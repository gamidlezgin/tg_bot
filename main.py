import asyncio
import os
import sqlite3
import re
import html
import csv
import io
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo
from collections import Counter
from openpyxl import Workbook

from aiogram import Bot, Dispatcher, F, types, BaseMiddleware
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    InlineKeyboardButton, InlineKeyboardMarkup, BotCommand,
    ReplyKeyboardMarkup, KeyboardButton, FSInputFile, BufferedInputFile
)
from aiogram.exceptions import TelegramBadRequest
from dotenv import load_dotenv

load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
if not BOT_TOKEN: raise RuntimeError("BOT_TOKEN не задан в .env")

DEVELOPER_ID = 6569042484
ADMIN_IDS = {int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip().isdigit()}
if DEVELOPER_ID not in ADMIN_IDS: ADMIN_IDS.add(DEVELOPER_ID)
PREMIUM_IDS = {int(x) for x in os.getenv("PREMIUM_IDS", "").split(",") if x.strip().isdigit()}
DB_PATH = "ufc.db"

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

TZ_OPTIONS = ["UTC", "Europe/Moscow", "Asia/Jerusalem", "Europe/London", "America/New_York", "Asia/Almaty", "Asia/Tbilisi"]

main_kb = ReplyKeyboardMarkup(keyboard=[
    [KeyboardButton(text="🥊 Турниры"), KeyboardButton(text="📊 Итоги")],
    [KeyboardButton(text="📅 Календарь"), KeyboardButton(text="📰 Новости")],
    [KeyboardButton(text="👤 Мой профиль"), KeyboardButton(text="💳 Купить Премиум")],
    [KeyboardButton(text="🔍 Поиск бойца"), KeyboardButton(text="📈 Топ прогнозов")],
    [KeyboardButton(text="⚙️ Настройки"), KeyboardButton(text="🎁 Промокод")],
    [KeyboardButton(text="🐛 Баг-репорт"), KeyboardButton(text="❓ Помощь")]
], resize_keyboard=True)

def now_utc() -> datetime: return datetime.now(timezone.utc).replace(tzinfo=None)

def safe_truncate_html(text: str, max_len: int) -> str:
    if len(text) <= max_len: return text
    cut = max_len - 80
    if cut < 1: return text[:max_len]
    text = text[:cut]
    text = re.sub(r'<[^>]*$', '', text)
    opened = re.findall(r'<(\w+)(?:\s[^>]*)?>', text)
    closed = re.findall(r'</(\w+)>', text)
    balance = Counter(opened) - Counter(closed)
    for tag, count in balance.items():
        if count > 0: text += f'</{tag}>' * count
    return text + "\n\n... (обрезано)"

def normalize_method(method: str) -> str:
    if not method: return ""
    m = method.upper()
    if "KO" in m or "TKO" in m or "НОКАУТ" in m: return "KO"
    if "SUB" in m or "БОЛЕВ" in m or "УДУШ" in m or "SAB" in m: return "SUB"
    return "DEC"

# ---------------------- БАЗА ----------------------
def db_init() -> None:
    con = sqlite3.connect(DB_PATH, timeout=10)
    con.executescript("""
    CREATE TABLE IF NOT EXISTS tournaments(
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, date_utc TEXT NOT NULL,
        location TEXT, event_type TEXT DEFAULT 'numbered', status TEXT DEFAULT 'announced', photo_id TEXT);
    CREATE TABLE IF NOT EXISTS fights(
        id INTEGER PRIMARY KEY AUTOINCREMENT, tournament_id INTEGER NOT NULL, weight_class TEXT,
        fighter_1 TEXT NOT NULL, fighter_2 TEXT NOT NULL, is_main_card INTEGER DEFAULT 1, is_title INTEGER DEFAULT 0,
        status TEXT DEFAULT 'scheduled', winner TEXT, method TEXT, round INTEGER, time TEXT,
        FOREIGN KEY(tournament_id) REFERENCES tournaments(id));
    CREATE TABLE IF NOT EXISTS users(
        telegram_id INTEGER PRIMARY KEY, timezone TEXT DEFAULT 'UTC', premium_until TEXT,
        first_name TEXT, username TEXT, first_seen TEXT, last_seen TEXT, total_messages INTEGER DEFAULT 0, aura_points INTEGER DEFAULT 0);
    CREATE TABLE IF NOT EXISTS admins(telegram_id INTEGER PRIMARY KEY);
    CREATE TABLE IF NOT EXISTS subscriptions(
        telegram_id INTEGER, fighter_name TEXT, PRIMARY KEY(telegram_id, fighter_name));
    CREATE TABLE IF NOT EXISTS news(
        id INTEGER PRIMARY KEY AUTOINCREMENT, text TEXT, photo_id TEXT, is_vip INTEGER DEFAULT 0, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS predictions(
        id INTEGER PRIMARY KEY AUTOINCREMENT, telegram_id INTEGER, fight_id INTEGER,
        predicted_winner TEXT, predicted_method TEXT, status TEXT DEFAULT 'pending', UNIQUE(telegram_id, fight_id));
    CREATE TABLE IF NOT EXISTS logs(
        id INTEGER PRIMARY KEY AUTOINCREMENT, admin_id INTEGER, action TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS bug_reports(
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, text TEXT, status TEXT DEFAULT 'open', created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS user_actions(
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, action_type TEXT, details TEXT, created_at TEXT);
    CREATE TABLE IF NOT EXISTS tour_reminders(
        telegram_id INTEGER, tournament_id INTEGER, reminder_type TEXT,
        PRIMARY KEY (telegram_id, tournament_id, reminder_type));
    CREATE TABLE IF NOT EXISTS promocodes(
        code TEXT PRIMARY KEY, days INTEGER, max_uses INTEGER, uses INTEGER DEFAULT 0, creator_id INTEGER);
    CREATE TABLE IF NOT EXISTS promo_activations(
        user_id INTEGER, code TEXT, PRIMARY KEY(user_id, code));
    """)
    migrations = [
        ("users", "first_name", "TEXT"), ("users", "username", "TEXT"), 
        ("users", "first_seen", "TEXT"), ("users", "last_seen", "TEXT"), 
        ("users", "total_messages", "INTEGER DEFAULT 0"), ("news", "is_vip", "INTEGER DEFAULT 0"),
        ("predictions", "predicted_method", "TEXT"),
        ("users", "aura_points", "INTEGER DEFAULT 0"),
        ("promocodes", "creator_id", "INTEGER")
    ]
    for table, col, typ in migrations:
        try: con.execute(f"ALTER TABLE {table} ADD COLUMN {col} {typ}")
        except sqlite3.OperationalError: pass
    con.commit(); con.close()

def log_action(admin_id: int, action: str):
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        con.execute("INSERT INTO logs(admin_id, action, created_at) VALUES(?,?,?)", (admin_id, action, now_utc().strftime("%Y-%m-%d %H:%M:%S")))
        con.commit()
    finally: con.close()

def log_user_activity(user_id: int, full_name: str, username: str, action_type: str, details: str):
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        now = now_utc().strftime("%Y-%m-%d %H:%M:%S")
        con.execute("""
            INSERT INTO users (telegram_id, first_name, username, first_seen, last_seen, total_messages)
            VALUES (?, ?, ?, ?, ?, 1)
            ON CONFLICT(telegram_id) DO UPDATE SET
                first_name = excluded.first_name, username = excluded.username,
                last_seen = excluded.last_seen, total_messages = total_messages + 1
        """, (user_id, full_name, username, now, now))
        con.execute("""
            INSERT INTO user_actions (user_id, action_type, details, created_at) VALUES (?, ?, ?, ?)
        """, (user_id, action_type, str(details)[:500], now))
        con.commit()
    except Exception as e: print(f"Ошибка логирования: {e}")
    finally: con.close()

class ActivityMiddleware(BaseMiddleware):
    async def __call__(self, handler, event, data):
        user = None; action_type = None; details = None
        if isinstance(event, types.Update):
            if event.message:
                user = event.message.from_user; action_type = "message"
                details = event.message.text or event.message.caption or f"[{event.message.content_type}]"
            elif event.callback_query:
                user = event.callback_query.from_user; action_type = "callback"; details = event.callback_query.data
        if user and not user.is_bot: log_user_activity(user.id, user.full_name, user.username or "", action_type, details or "")
        return await handler(event, data)

dp.update.outer_middleware(ActivityMiddleware())

def is_admin(uid: int) -> bool: return uid in ADMIN_IDS

def is_premium(uid: int) -> bool:
    if uid in PREMIUM_IDS or uid == DEVELOPER_ID: return True
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        row = con.execute("SELECT premium_until FROM users WHERE telegram_id=?", (uid,)).fetchone()
        if row and row[0]:
            try: return datetime.strptime(row[0], "%Y-%m-%d %H:%M:%S") > now_utc()
            except ValueError: return False
    finally: con.close()
    return False

def get_user_role_badge(uid: int) -> str:
    if uid == DEVELOPER_ID: return "👑 <b>Разработчик</b>"
    if is_admin(uid): return "🛠 <b>Админ</b>"
    if is_premium(uid): return "⭐ <b>Премиум</b>"
    return "👤 <b>Пользователь</b>"

def get_premium_status_text(uid: int) -> str:
    if uid in PREMIUM_IDS or uid == DEVELOPER_ID: return "✅ <b>Активен навсегда (VIP)</b>"
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        row = con.execute("SELECT premium_until FROM users WHERE telegram_id=?", (uid,)).fetchone()
        if row and row[0]:
            try:
                dt = datetime.strptime(row[0], "%Y-%m-%d %H:%M:%S")
                if dt > now_utc(): return f"✅ Активен до {dt.strftime('%d.%m.%Y')}"
            except ValueError: pass
    finally: con.close()
    return "❌ Не активен"

def get_user_tz(uid: int) -> ZoneInfo:
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        row = con.execute("SELECT timezone FROM users WHERE telegram_id=?", (uid,)).fetchone()
        name = row[0] if row else "UTC"
        try: return ZoneInfo(name)
        except Exception: return ZoneInfo("UTC")
    finally: con.close()

def fmt_utc_to_local(date_utc: str, tz: ZoneInfo) -> str:
    dt = datetime.strptime(date_utc, "%Y-%m-%d %H:%M").replace(tzinfo=ZoneInfo("UTC"))
    return dt.astimezone(tz).strftime("%Y-%m-%d %H:%M %Z")

async def broadcast_to_all(text: str = None, photo_id: str = None, exclude_id: int = None, only_premium: bool = False, exclude_premium: bool = False, copy_msg: types.Message = None):
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: users = con.execute("SELECT telegram_id FROM users").fetchall()
    finally: con.close()
    success = 0; failed = 0
    safe_text = text[:1024] if text else None
    for (uid,) in users:
        if uid == exclude_id: continue
        if only_premium and not is_premium(uid): continue
        if exclude_premium and is_premium(uid): continue
        try:
            if copy_msg:
                await bot.copy_message(chat_id=uid, from_chat_id=copy_msg.chat.id, message_id=copy_msg.message_id)
            elif photo_id: 
                await bot.send_photo(chat_id=uid, photo=photo_id, caption=safe_text, parse_mode="HTML")
            else: 
                await bot.send_message(chat_id=uid, text=text, parse_mode="HTML")
            success += 1; await asyncio.sleep(0.05)
        except Exception: failed += 1
    return success, failed

async def notify_subscribers(fighters: list, tournament_name: str):
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        for fighter in fighters:
            safe = fighter.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
            subs = con.execute("SELECT telegram_id FROM subscriptions WHERE fighter_name LIKE ? ESCAPE '\\'", (f"%{safe}%",)).fetchall()
            for (uid,) in subs:
                if is_premium(uid):
                    try: await bot.send_message(uid, f"⭐ <b>Внимание!</b>\n\nТвой боец <b>{html.escape(fighter)}</b> на <b>{html.escape(tournament_name)}</b>!", parse_mode="HTML")
                    except Exception: pass
    finally: con.close()

def get_fighter_stats(name: str) -> dict:
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        fights = con.execute("""
            SELECT fighter_1, fighter_2, winner, method, round, is_title
            FROM fights WHERE (fighter_1=? OR fighter_2=?) AND status='finished'
        """, (name, name)).fetchall()
    finally: con.close()
    total = len(fights); wins = 0; losses = 0; methods = Counter(); rounds = []; title_wins = 0
    for f1, f2, winner, method, rnd, is_title in fights:
        if not winner: continue
        w = winner.lower().strip()
        if w in (f1.lower().strip(), f2.lower().strip()):
            if w == name.lower().strip():
                wins += 1
                if method: methods[method.upper()] += 1
                if rnd: rounds.append(rnd)
                if is_title: title_wins += 1
            else: losses += 1
    avg_round = round(sum(rounds) / len(rounds), 1) if rounds else 0
    win_pct = round(wins / total * 100) if total else 0
    return {"total": total, "wins": wins, "losses": losses, "win_pct": win_pct, "methods": methods, "avg_round": avg_round, "title_wins": title_wins}

# ---------------------- ФОНОВАЯ ЗАДАЧА ----------------------
async def reminder_worker():
    while True:
        try:
            now = now_utc()
            con = sqlite3.connect(DB_PATH, timeout=10)
            try:
                tours = con.execute("SELECT id, name, date_utc FROM tournaments WHERE status='announced'").fetchall()
                for tid, name, date_utc in tours:
                    tour_time = datetime.strptime(date_utc, "%Y-%m-%d %H:%M")
                    diff = (tour_time - now).total_seconds() / 60
                    
                    if 1415 <= diff <= 1465:
                        users = con.execute("SELECT telegram_id FROM tour_reminders WHERE tournament_id=? AND reminder_type='24h'", (tid,)).fetchall()
                        for (uid,) in users:
                            try:
                                await bot.send_message(uid, f"⏳ <b>Напоминание!</b>\n\nТурнир <b>{html.escape(name)}</b> начнется уже завтра!\nНе пропусти.", parse_mode="HTML")
                                con.execute("DELETE FROM tour_reminders WHERE telegram_id=? AND tournament_id=? AND reminder_type='24h'", (uid, tid))
                                con.commit()
                            except Exception: pass
                            
                    if 55 <= diff <= 65:
                        users = con.execute("SELECT telegram_id FROM tour_reminders WHERE tournament_id=? AND reminder_type='1h'", (tid,)).fetchall()
                        for (uid,) in users:
                            try:
                                await bot.send_message(uid, f"🔔 <b>Напоминание!</b>\n\nТурнир <b>{html.escape(name)}</b> начнется примерно через час!\nНе пропусти главный кард.", parse_mode="HTML")
                                con.execute("DELETE FROM tour_reminders WHERE telegram_id=? AND tournament_id=? AND reminder_type='1h'", (uid, tid))
                                con.commit()
                            except Exception: pass
            finally: con.close()
        except Exception as e: print(f"Ошибка воркера напоминаний: {e}")
        await asyncio.sleep(60)

# ---------------------- FSM ----------------------
class AddEvent(StatesGroup):
    name = State(); date = State(); location = State(); event_type = State(); photo = State()
class EditEvent(StatesGroup):
    name = State(); date = State(); location = State(); event_type = State(); photo = State()
class EditFight(StatesGroup):
    weight = State(); f1 = State(); f2 = State(); main_card = State(); title = State()
class SetResult(StatesGroup):
    tournament = State(); fight = State(); winner = State(); method = State(); round = State(); time = State()
class PromoState(StatesGroup):
    waiting_for_code = State()

async def set_commands(bot: Bot):
    commands = [
        BotCommand(command="/start", description="Главное меню"),
        BotCommand(command="/fights", description="Турниры"),
        BotCommand(command="/results", description="Итоги"),
        BotCommand(command="/profile", description="Мой профиль"),
        BotCommand(command="/calendar", description="Календарь"),
        BotCommand(command="/news", description="Новости"),
        BotCommand(command="/search", description="Поиск бойца"),
        BotCommand(command="/top_predictors", description="Топ прогнозистов"),
        BotCommand(command="/premium", description="Купить Премиум"),
        BotCommand(command="/timezone", description="Часовой пояс"),
        BotCommand(command="/myid", description="Мой ID"),
        BotCommand(command="/bug", description="Сообщить о проблеме"),
        BotCommand(command="/help", description="Помощь"),
    ]
    await bot.set_my_commands(commands)

# ---------------------- МЕНЮ ----------------------
def get_main_inline_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🥊 Турниры", callback_data="menu_fights"),
         InlineKeyboardButton(text="📊 Итоги", callback_data="menu_results")],
        [InlineKeyboardButton(text="📅 Календарь", callback_data="menu_calendar"),
         InlineKeyboardButton(text="📰 Новости", callback_data="menu_news")],
        [InlineKeyboardButton(text="👤 Мой профиль", callback_data="menu_profile"),
         InlineKeyboardButton(text="💳 Премиум", callback_data="menu_premium")],
        [InlineKeyboardButton(text="🔍 Поиск бойца", callback_data="menu_search"),
         InlineKeyboardButton(text="📈 Топ прогнозов", callback_data="menu_top")],
        [InlineKeyboardButton(text="⚙️ Часовой пояс", callback_data="menu_timezone"),
         InlineKeyboardButton(text="❓ Помощь", callback_data="menu_help")]
    ])

@dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    kb = get_main_inline_menu()
    await message.answer(
        f"Привет, {message.from_user.full_name}!\n\nВыбери нужное действие на кнопках ниже (или используй клавиатуру внизу):", 
        reply_markup=kb
    )
    await message.answer("Меню активировано 👇", reply_markup=main_kb)

@dp.callback_query(F.data == "menu_back")
async def cb_menu_back(callback: types.CallbackQuery):
    await callback.answer()
    try: await callback.message.edit_text("Главное меню:", reply_markup=get_main_inline_menu())
    except TelegramBadRequest: pass

@dp.callback_query(F.data == "menu_fights")
async def cb_menu_fights(callback: types.CallbackQuery):
    await callback.answer()
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🥊 Fight Night", callback_data="next_type:fight_night:0")],
        [InlineKeyboardButton(text="🏆 Номерной", callback_data="next_type:numbered:0")],
        [InlineKeyboardButton(text="⬅️ Назад", callback_data="menu_back")]
    ])
    try: await callback.message.edit_text("Выбери тип турнира:", reply_markup=kb)
    except TelegramBadRequest: pass

@dp.callback_query(F.data == "menu_results")
async def cb_menu_results(callback: types.CallbackQuery):
    await callback.answer()
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        tours = con.execute("SELECT id, name, date_utc FROM tournaments WHERE status='finished' ORDER BY date_utc DESC LIMIT 3").fetchall()
        if not tours: return await callback.message.edit_text("Итогов пока нет.", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⬅️ Назад", callback_data="menu_back")]]))
        tz = get_user_tz(callback.from_user.id)
        parts = []
        for tid, name, date_utc in tours:
            block = f"✅ <b>{html.escape(name)}</b>\n📅 {fmt_utc_to_local(date_utc, tz)}\n"
            fights = con.execute("SELECT fighter_1, fighter_2, winner, method, round, time FROM fights WHERE tournament_id=? AND status='finished' ORDER BY is_main_card ASC, id", (tid,)).fetchall()
            for f1, f2, w, m, r, t in fights:
                if w: block += f"• {html.escape(f1)} vs {html.escape(f2)} → {html.escape(w)} ({html.escape(m or '?')}, R{r} {t})\n"
                else: block += f"• {html.escape(f1)} vs {html.escape(f2)} → нет\n"
            parts.append(block)
    finally: con.close()
    result = "\n".join(parts)
    if len(result) > 4000: result = safe_truncate_html(result, 4000)
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⬅️ Назад", callback_data="menu_back")]])
    try: await callback.message.edit_text(result, parse_mode="HTML", reply_markup=kb)
    except TelegramBadRequest: pass

@dp.callback_query(F.data == "menu_calendar")
async def cb_menu_calendar(callback: types.CallbackQuery):
    await callback.answer()
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: tours = con.execute("SELECT id, name, date_utc, location, event_type FROM tournaments WHERE status='announced' ORDER BY date_utc ASC LIMIT 20").fetchall()
    finally: con.close()
    if not tours: return await callback.message.edit_text("📅 Турниров пока нет.", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⬅️ Назад", callback_data="menu_back")]]))
    tz = get_user_tz(callback.from_user.id)
    lines = ["📅 <b>Календарь турниров:</b>\n"]
    for tid, name, date_utc, loc, etype in tours:
        local = fmt_utc_to_local(date_utc, tz)
        emoji = "🏆" if etype == "numbered" else "🥊"
        lines.append(f"{emoji} <b>{html.escape(name)}</b>\n📅 {local}\n📍 {html.escape(loc or '—')}\n")
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⬅️ Назад", callback_data="menu_back")]])
    try: await callback.message.edit_text(safe_truncate_html("\n".join(lines), 4000), parse_mode="HTML", reply_markup=kb)
    except TelegramBadRequest: pass

@dp.callback_query(F.data == "menu_news")
async def cb_menu_news(callback: types.CallbackQuery):
    await callback.answer()
    premium = is_premium(callback.from_user.id)
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: rows = con.execute("SELECT id, text, photo_id, is_vip, created_at FROM news ORDER BY id DESC LIMIT 10").fetchall()
    finally: con.close()
    if not rows: return await callback.message.edit_text("📰 Новостей пока нет.", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⬅️ Назад", callback_data="menu_back")]]))
    try: await callback.message.delete()
    except Exception: pass
    await callback.message.answer(f"📰 <b>Последние новости:</b>", parse_mode="HTML")
    for nid, text, photo_id, is_vip, dt in rows[:5]:
        if is_vip and not premium: continue
        prefix = "👑 <b>СЛИВКИ (VIP)</b>" if is_vip else "📰"
        cap = f"{prefix} <b>#{nid}</b> · <i>{dt}</i>\n\n{html.escape(text or '')}"
        try:
            if photo_id: await callback.message.answer_photo(photo=photo_id, caption=cap[:1024], parse_mode="HTML")
            else: await callback.message.answer(cap[:4000], parse_mode="HTML")
        except Exception: pass
    await callback.message.answer("Выбери действие:", reply_markup=get_main_inline_menu())

@dp.callback_query(F.data == "menu_profile")
async def cb_menu_profile(callback: types.CallbackQuery):
    await callback.answer()
    uid = callback.from_user.id
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        user = con.execute("SELECT first_name, username, first_seen, last_seen, total_messages, premium_until, timezone, aura_points FROM users WHERE telegram_id=?", (uid,)).fetchone()
        preds = con.execute("SELECT status, predicted_method FROM predictions WHERE telegram_id=?", (uid,)).fetchall()
        subs = con.execute("SELECT fighter_name FROM subscriptions WHERE telegram_id=?", (uid,)).fetchall()
    finally: con.close()
    if not user: return await callback.message.edit_text("Профиль не найден. Нажми /start")
    first_name, username, first_seen, last_seen, total_messages, premium_until, tz_name, aura = user
    status = get_premium_status_text(uid)
    role_badge = get_user_role_badge(uid)
    won = sum(1 for p in preds if p[0] == 'won')
    lost = sum(1 for p in preds if p[0] == 'lost')
    pending = sum(1 for p in preds if p[0] == 'pending')
    perfect = sum(1 for p in preds if p[0] == 'won' and p[1] is not None)
    total_preds = won + lost
    pct = round(won / total_preds * 100) if total_preds > 0 else 0
    subs_list = ", ".join([s[0] for s in subs[:5]])
    if len(subs) > 5: subs_list += f" и еще {len(subs)-5}"
    text = (
        f"👤 <b>Твой профиль</b>\n\n"
        f"🆔 ID: <code>{uid}</code>\n"
        f"📛 Имя: {html.escape(first_name or '—')}\n"
        f"🔗 Юзернейм: @{username or '—'}\n"
        f"🎖 Роль: {role_badge}\n"
        f"⭐ Статус: {status}\n"
        f"✨ Аура: <b>{aura}</b>\n\n"
        f"📅 В боте с: {first_seen[:10] if first_seen else '—'}\n"
        f"💬 Сообщений: {total_messages}\n"
        f"⚙️ Часовой пояс: {tz_name}\n\n"
        f"🎯 <b>Прогнозы:</b>\n"
        f"• Угадано: {won} из {total_preds} (Точность: {pct}%)\n"
        f"• Идеальных (с методом): {perfect}\n"
        f"• В ожидании: {pending}\n\n"
        f"🔔 <b>Подписки:</b>\n"
        f"• {html.escape(subs_list) if subs_list else 'Нет подписок'}\n"
    )
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⚙️ Настройки", callback_data="menu_timezone")],
        [InlineKeyboardButton(text="⬅️ Назад", callback_data="menu_back")]
    ])
    try: await callback.message.edit_text(text, parse_mode="HTML", reply_markup=kb)
    except TelegramBadRequest: pass

@dp.callback_query(F.data == "menu_premium")
async def cb_menu_premium(callback: types.CallbackQuery):
    await callback.answer()
    status = get_premium_status_text(callback.from_user.id)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔔 Мои подписки", callback_data="prem_subs")],
        [InlineKeyboardButton(text="📈 Мои прогнозы", callback_data="prem_my_stats")],
        [InlineKeyboardButton(text="📊 Скачать прогнозы (Excel)", callback_data="prem_export_preds")],
        [InlineKeyboardButton(text="🎁 Ввести промокод", callback_data="prem_enter_promo")],
        [InlineKeyboardButton(text="💬 Написать @sultanovg", url="https://t.me/sultanovg")],
        [InlineKeyboardButton(text="⬅️ Назад", callback_data="menu_back")]
    ])
    try:
        await callback.message.edit_text(
            f"⭐ <b>Премиум-подписка UFC Cards</b>\n\n"
            f"Ваш статус: {status}\n\n"
            f"💳 <b>Стоимость:</b> 15 ₽ / 30 дней\n"
            f"📩 <b>Для покупки:</b> напишите <a href='https://t.me/sultanovg'>@sultanovg</a>\n\n"
            f"🔥 <b>Эксклюзивные функции:</b>\n"
            f"┣ 🎯 <b>Прогнозы с методом</b> — угадывай не только победителя, но и как он победит (KO, Sub, Dec).\n"
            f"┣ 🔔 <b>Подписки на бойцов</b> — узнавай первым, когда твой любимчик выходит на ринг.\n"
            f"┣ 🆕 <b>Ранние анонсы</b> — уведомления о новых кардах до общей рассылки.\n"
            f"┣ ⚡ <b>Ранний доступ</b> — результаты боев приходят на 10 минут раньше остальных!\n"
            f"┣ 🔮 <b>Мнение сообщества</b> — смотри, на кого ставит большинство перед своим прогнозом.\n"
            f"┣ 📈 <b>Личная статистика</b> — детальный анализ твоих прогнозов.\n"
            f"┣ 📊 <b>Экспорт прогнозов</b> — скачай свои ставки в Excel.\n"
            f"┣ 👑 <b>Сливки (VIP-новости)</b> — доступ к самым горячим инсайдам и слухам.\n"
            f"┣ ⏰ <b>Напоминания</b> — бот пинганет тебя за час до начала турнира.\n"
            f"┗ ⭐ <b>VIP-значок</b> — выделяйся звездой в топе прогнозистов.\n\n"
            f"<i>Оформи подписку и получи максимальный доступ к миру MMA!</i>",
            reply_markup=kb, parse_mode="HTML"
        )
    except TelegramBadRequest: pass

@dp.callback_query(F.data == "menu_search")
async def cb_menu_search(callback: types.CallbackQuery):
    await callback.answer()
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⬅️ Назад", callback_data="menu_back")]])
    try: await callback.message.edit_text("🔍 Напиши в чат: `/search Имя`\n\nНапример: `/search Ган`", parse_mode="Markdown", reply_markup=kb)
    except TelegramBadRequest: pass

@dp.callback_query(F.data == "menu_top")
async def cb_menu_top(callback: types.CallbackQuery):
    await callback.answer()
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        rows = con.execute("""
            SELECT telegram_id,
                   SUM(CASE WHEN status='won' THEN 1 ELSE 0 END) as won,
                   SUM(CASE WHEN status IN ('won','lost') THEN 1 ELSE 0 END) as total
            FROM predictions GROUP BY telegram_id HAVING total > 0
            ORDER BY won DESC, total ASC LIMIT 10
        """).fetchall()
    finally: con.close()
    if not rows: return await callback.message.edit_text("🎯 Прогнозов пока нет.", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⬅️ Назад", callback_data="menu_back")]]))
    lines = ["🎯 <b>Топ прогнозистов:</b>\n"]
    for i, (uid, won, total) in enumerate(rows, 1):
        pct = round(won / total * 100)
        med = {1: "🥇", 2: "🥈", 3: "🥉"}.get(i, f"{i}.")
        if uid == DEVELOPER_ID: badge = " 👑"
        elif is_admin(uid): badge = " 🛠"
        elif is_premium(uid): badge = " ⭐"
        else: badge = ""
        lines.append(f"{med} ID <code>{uid}</code>{badge}: {won}/{total} ({pct}%)")
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⬅️ Назад", callback_data="menu_back")]])
    try: await callback.message.edit_text("\n".join(lines), parse_mode="HTML", reply_markup=kb)
    except TelegramBadRequest: pass

@dp.callback_query(F.data == "menu_timezone")
async def cb_menu_timezone(callback: types.CallbackQuery):
    await callback.answer()
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=tz, callback_data=f"tz:{tz}")] for tz in TZ_OPTIONS])
    kb.inline_keyboard.append([InlineKeyboardButton(text="⬅️ Назад", callback_data="menu_back")])
    try: await callback.message.edit_text("Выбери свой часовой пояс:", reply_markup=kb)
    except TelegramBadRequest: pass

@dp.callback_query(F.data == "menu_help")
async def cb_menu_help(callback: types.CallbackQuery):
    await callback.answer()
    user_help = (
        "👤 <b>КОМАНДЫ ПОЛЬЗОВАТЕЛЯ:</b>\n"
        "/start - Главное меню\n"
        "/profile - Мой профиль\n"
        "/fights - Турниры\n"
        "/results - Итоги\n"
        "/calendar - Календарь\n"
        "/news - Новости\n"
        "/search - Поиск\n"
        "/top_predictors - Топ\n"
        "/premium - Премиум\n"
        "/promo - Промокод\n"
        "/timezone - Часовой пояс\n"
        "/bug - Сообщить о проблеме\n"
    )
    admin_help = (
        "\n\n👑 <b>КОМАНДЫ АДМИНИСТРАТОРА:</b>\n"
        "/user - Все юзеры (список в чате + Excel)\n"
        "/premium_list - Список премиум-юзеров\n"
        "/add_event - Добавить новый турнир\n"
        "/edit_event - Редактировать турнир\n"
        "/delete_event - Удалить турнир\n"
        "/finish_event - Завершить турнир\n"
        "/card - Загрузить кард / список ID\n"
        "/edit_fight - Редактировать бой\n"
        "/delete_fight - Удалить бой\n"
        "/set_result - Внести результат боя\n"
        "/reset_result - Сбросить результат боя\n"
        "/edit_id - Изменить ID турнира\n"
        "/news_add - Опубликовать новость (или ответ на медиа)\n"
        "/vip_news - Опубликовать VIP-новость\n"
        "/news_delete - Удалить новость\n"
        "/broadcast - Сделать рассылку (или ответ на опрос)\n"
        "/send_id - Написать юзеру по ID\n"
        "/stats - Статистика бота\n"
        "/backup - Скачать базу данных\n"
        "/export_excel - Скачать базу в Excel\n"
        "/export_csv - Скачать бои в CSV\n"
        "/archive - Архивировать турниры\n"
        "/logs - Логи действий админов\n"
        "/bug_list - Открытые баги\n"
        "/add_admin - Добавить админа\n"
        "/give_premium - Выдать Премиум\n"
        "/give_aura - Начислить Ауру\n"
        "/remove_aura - Списать Ауру\n"
        "/create_promo - Создать промокод\n"
        "/promo_list - Список промокодов\n"
        "/delete_promo - Удалить промокод\n"
        "/cancel - Отменить ввод\n"
    )
    
    dev_help = ""
    if callback.from_user.id == DEVELOPER_ID:
        dev_help = (
            "\n\n🛠 <b>КОМАНДЫ РАЗРАБОТЧИКА:</b>\n"
            "/del_admin - Удалить админа\n"
        )
    
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⬅️ Назад", callback_data="menu_back")]])
    if callback.from_user.id == DEVELOPER_ID:
        text = user_help + admin_help + dev_help
    elif is_admin(callback.from_user.id):
        text = user_help + admin_help
    else:
        text = user_help
        
    try: await callback.message.edit_text(text, parse_mode="HTML", reply_markup=kb)
    except TelegramBadRequest: pass

# ---------------------- ОСТАЛЬНЫЕ БАЗОВЫЕ КОМАНДЫ ----------------------
@dp.message(Command("help"))
@dp.message(F.text == "❓ Помощь")
async def cmd_help(message: types.Message, state: FSMContext):
    await state.clear()
    user_help = (
        "👤 <b>КОМАНДЫ ПОЛЬЗОВАТЕЛЯ:</b>\n"
        "/start - Главное меню\n"
        "/profile - Мой профиль (статистика, аура, подписки)\n"
        "/fights - Список предстоящих турниров\n"
        "/results - Результаты прошедших турниров\n"
        "/calendar - Календарь турниров\n"
        "/news - Последние новости\n"
        "/search - Поиск бойца по имени\n"
        "/top_predictors - Топ-10 прогнозистов\n"
        "/premium - Меню Премиум-подписки\n"
        "/promo - Активировать промокод\n"
        "/timezone - Выбрать часовой пояс\n"
        "/myid - Узнать свой Telegram ID\n"
        "/bug - Сообщить о проблеме\n"
        "/help - Показать это сообщение"
    )
    admin_help = (
        "\n\n👑 <b>КОМАНДЫ АДМИНИСТРАТОРА:</b>\n"
        "/user - Все юзеры (список в чате + Excel)\n"
        "/premium_list - Список премиум-юзеров\n"
        "/add_event - Добавить новый турнир\n"
        "/edit_event - Редактировать турнир\n"
        "/delete_event - Удалить турнир\n"
        "/finish_event - Завершить турнир\n"
        "/card - Загрузить кард / список ID\n"
        "/edit_fight - Редактировать бой\n"
        "/delete_fight - Удалить бой\n"
        "/set_result - Внести результат боя\n"
        "/reset_result - Сбросить результат боя\n"
        "/edit_id - Изменить ID турнира\n"
        "/news_add - Опубликовать новость (или ответ на медиа)\n"
        "/vip_news - Опубликовать VIP-новость\n"
        "/news_delete - Удалить новость\n"
        "/broadcast - Сделать рассылку (или ответ на опрос)\n"
        "/send_id - Написать юзеру по ID\n"
        "/stats - Статистика бота\n"
        "/backup - Скачать базу данных\n"
        "/export_excel - Скачать базу в Excel\n"
        "/export_csv - Скачать бои в CSV\n"
        "/archive - Архивировать турниры\n"
        "/logs - Логи действий админов\n"
        "/bug_list - Открытые баги\n"
        "/add_admin - Добавить админа\n"
        "/give_premium - Выдать Премиум\n"
        "/give_aura - Начислить Ауру\n"
        "/remove_aura - Списать Ауру\n"
        "/create_promo - Создать промокод\n"
        "/promo_list - Список промокодов\n"
        "/delete_promo - Удалить промокод\n"
        "/cancel - Отменить ввод\n"
    )
    
    if message.from_user.id == DEVELOPER_ID:
        dev_help = "\n\n🛠 <b>КОМАНДЫ РАЗРАБОТЧИКА:</b>\n/del_admin - Удалить админа\n"
        await message.answer(user_help + admin_help + dev_help, reply_markup=main_kb, parse_mode="HTML")
    elif is_admin(message.from_user.id):
        await message.answer(user_help + admin_help, reply_markup=main_kb, parse_mode="HTML")
    else:
        await message.answer(user_help, reply_markup=main_kb, parse_mode="HTML")

@dp.message(Command("myid"))
async def cmd_myid(message: types.Message): await message.answer(f"Твой ID: {message.from_user.id}")

@dp.message(Command("cancel"))
async def cmd_cancel(message: types.Message, state: FSMContext):
    await state.clear(); await message.answer("Отменено.", reply_markup=main_kb)

@dp.message(Command("profile"))
@dp.message(F.text == "👤 Мой профиль")
async def cmd_profile(message: types.Message, state: FSMContext):
    await state.clear()
    uid = message.from_user.id
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        user = con.execute("SELECT first_name, username, first_seen, last_seen, total_messages, premium_until, timezone, aura_points FROM users WHERE telegram_id=?", (uid,)).fetchone()
        preds = con.execute("SELECT status, predicted_method FROM predictions WHERE telegram_id=?", (uid,)).fetchall()
        subs = con.execute("SELECT fighter_name FROM subscriptions WHERE telegram_id=?", (uid,)).fetchall()
    finally: con.close()
    if not user: return await message.answer("Профиль не найден. Нажми /start")
    first_name, username, first_seen, last_seen, total_messages, premium_until, tz_name, aura = user
    status = get_premium_status_text(uid)
    role_badge = get_user_role_badge(uid)
    won = sum(1 for p in preds if p[0] == 'won')
    lost = sum(1 for p in preds if p[0] == 'lost')
    pending = sum(1 for p in preds if p[0] == 'pending')
    perfect = sum(1 for p in preds if p[0] == 'won' and p[1] is not None)
    total_preds = won + lost
    pct = round(won / total_preds * 100) if total_preds > 0 else 0
    subs_list = ", ".join([s[0] for s in subs[:5]])
    if len(subs) > 5: subs_list += f" и еще {len(subs)-5}"
    text = (
        f"👤 <b>Твой профиль</b>\n\n"
        f"🆔 ID: <code>{uid}</code>\n"
        f"📛 Имя: {html.escape(first_name or '—')}\n"
        f"🔗 Юзернейм: @{username or '—'}\n"
        f"🎖 Роль: {role_badge}\n"
        f"⭐ Статус: {status}\n"
        f"✨ Аура: <b>{aura}</b>\n\n"
        f"📅 В боте с: {first_seen[:10] if first_seen else '—'}\n"
        f"💬 Сообщений: {total_messages}\n"
        f"⚙️ Часовой пояс: {tz_name}\n\n"
        f"🎯 <b>Прогнозы:</b>\n"
        f"• Угадано: {won} из {total_preds} (Точность: {pct}%)\n"
        f"• Идеальных (с методом): {perfect}\n"
        f"• В ожидании: {pending}\n\n"
        f"🔔 <b>Подписки:</b>\n"
        f"• {html.escape(subs_list) if subs_list else 'Нет подписок'}\n"
    )
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⚙️ Настройки", callback_data="open_settings")]])
    await message.answer(text, parse_mode="HTML", reply_markup=kb)

@dp.callback_query(F.data == "open_settings")
async def cb_open_settings(callback: types.CallbackQuery):
    await callback.answer()
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=tz, callback_data=f"tz:{tz}")] for tz in TZ_OPTIONS])
    await callback.message.answer("Выбери свой часовой пояс:", reply_markup=kb)

# ---------------------- ПРОМОКОДЫ И АУРА ----------------------
@dp.message(Command("create_promo"))
async def cmd_create_promo(message: types.Message):
    if not is_admin(message.from_user.id): return
    parts = message.text.split()
    if len(parts) < 4 or not parts[2].isdigit() or not parts[3].isdigit():
        return await message.answer("❌ Использование: `/create_promo КОД ДНИ ЛИМИТ`\n\n*(Если ответить на сообщение юзера, аура за активации пойдет ему)*", parse_mode="Markdown")
    
    code = parts[1].upper(); days = int(parts[2]); max_uses = int(parts[3])
    creator_id = None
    
    if message.reply_to_message:
        creator_id = message.reply_to_message.from_user.id
        
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        existing = con.execute("SELECT code FROM promocodes WHERE code=?", (code,)).fetchone()
        if existing: return await message.answer(f"❌ Промокод `{code}` уже существует!", parse_mode="Markdown")
        con.execute("INSERT INTO promocodes(code, days, max_uses, creator_id) VALUES(?,?,?,?)", (code, days, max_uses, creator_id))
        con.commit()
        log_action(message.from_user.id, f"Создал промокод {code} на {days} дней")
        
        text = f"✅ Промокод <b>{code}</b> создан!\n🎁 Даёт: {days} дней Премиума\n📊 Лимит: {max_uses}"
        if creator_id:
            text += f"\n\n✨ <b>Владелец:</b> <code>{creator_id}</code> (будет получать +200 Ауры за активацию)."
        
        await message.answer(text, parse_mode="HTML")
    finally: con.close()

@dp.message(Command("promo_list"))
async def cmd_promo_list(message: types.Message):
    if not is_admin(message.from_user.id): return
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: rows = con.execute("SELECT code, days, max_uses, uses, creator_id FROM promocodes ORDER BY uses DESC").fetchall()
    finally: con.close()
    if not rows: return await message.answer("Промокодов пока нет.")
    lines = ["🎟 <b>Список промокодов:</b>\n"]
    for code, days, max_uses, uses, creator_id in rows:
        status = "✅ Активен" if uses < max_uses else "❌ Исчерпан"
        creator_str = f" [Владелец: {creator_id}]" if creator_id else ""
        lines.append(f"<code>{code}</code> | {days} дн. | {uses}/{max_uses} | {status}{creator_str}")
    await message.answer("\n".join(lines), parse_mode="HTML")

@dp.message(Command("delete_promo"))
async def cmd_delete_promo(message: types.Message):
    if not is_admin(message.from_user.id): return
    parts = message.text.split()
    if len(parts) < 2: return await message.answer("❌ `/delete_promo КОД`", parse_mode="Markdown")
    code = parts[1].upper()
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        con.execute("DELETE FROM promocodes WHERE code=?", (code,))
        con.execute("DELETE FROM promo_activations WHERE code=?", (code,))
        con.commit()
        await message.answer(f"✅ Промокод <b>{code}</b> удален.", parse_mode="HTML")
    finally: con.close()

@dp.message(Command("give_aura"))
async def cmd_give_aura(message: types.Message):
    if not is_admin(message.from_user.id): return
    parts = message.text.split()
    if len(parts) < 3 or not parts[1].isdigit() or not parts[2].isdigit():
        return await message.answer("❌ Использование: `/give_aura ID СУММА`", parse_mode="Markdown")
    target_id, amount = int(parts[1]), int(parts[2])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        con.execute("UPDATE users SET aura_points = aura_points + ? WHERE telegram_id=?", (amount, target_id))
        con.commit()
        await message.answer(f"✅ Пользователю <code>{target_id}</code> начислено <b>{amount}</b> Ауры.", parse_mode="HTML")
        try: await bot.send_message(target_id, f"✨ Тебе начислено <b>{amount}</b> Ауры!", parse_mode="HTML")
        except Exception: pass
    finally: con.close()

@dp.message(Command("remove_aura"))
async def cmd_remove_aura(message: types.Message):
    if not is_admin(message.from_user.id): return
    parts = message.text.split()
    if len(parts) < 3 or not parts[1].isdigit() or not parts[2].isdigit():
        return await message.answer("❌ Использование: `/remove_aura ID СУММА`", parse_mode="Markdown")
    target_id, amount = int(parts[1]), int(parts[2])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        con.execute("UPDATE users SET aura_points = aura_points - ? WHERE telegram_id=?", (amount, target_id))
        con.commit()
        await message.answer(f"✅ У пользователя <code>{target_id}</code> списано <b>{amount}</b> Ауры.", parse_mode="HTML")
    finally: con.close()

@dp.message(Command("promo"))
@dp.message(F.text == "🎁 Промокод")
async def cmd_promo(message: types.Message, state: FSMContext):
    await state.clear()
    if message.text == "🎁 Промокод":
        await state.set_state(PromoState.waiting_for_code)
        return await message.answer("🎁 Введи свой промокод (например: <code>UFC2026</code>):", parse_mode="HTML")
    parts = message.text.split()
    if len(parts) < 2: return await message.answer("❌ `/promo КОД`", parse_mode="Markdown")
    await process_promo_code(message.from_user.id, parts[1].upper(), message)

@dp.callback_query(F.data == "prem_enter_promo")
async def cb_prem_enter_promo(callback: types.CallbackQuery, state: FSMContext):
    await callback.answer()
    await state.set_state(PromoState.waiting_for_code)
    await callback.message.answer("🎁 Введи свой промокод:", parse_mode="HTML")

@dp.message(PromoState.waiting_for_code)
async def handle_promo_input(message: types.Message, state: FSMContext):
    await state.clear()
    await process_promo_code(message.from_user.id, message.text.strip().upper(), message)

async def process_promo_code(uid: int, code: str, message: types.Message):
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        promo = con.execute("SELECT days, max_uses, uses, creator_id FROM promocodes WHERE code=?", (code,)).fetchone()
        if not promo: return await message.answer("❌ Такого промокода не существует.")
        days, max_uses, uses, creator_id = promo
        if uses >= max_uses: return await message.answer("❌ Лимит активаций исчерпан.")
        activated = con.execute("SELECT 1 FROM promo_activations WHERE user_id=? AND code=?", (uid, code)).fetchone()
        if activated: return await message.answer("❌ Ты уже активировал этот промокод.")
        
        row = con.execute("SELECT premium_until FROM users WHERE telegram_id=?", (uid,)).fetchone()
        current = now_utc()
        if row and row[0]:
            try:
                old = datetime.strptime(row[0], "%Y-%m-%d %H:%M:%S")
                if old > current: current = old
            except ValueError: pass
                
        new_date = current + timedelta(days=days)
        con.execute("UPDATE users SET premium_until=? WHERE telegram_id=?", (new_date.strftime("%Y-%m-%d %H:%M:%S"), uid))
        con.execute("UPDATE promocodes SET uses = uses + 1 WHERE code=?", (code,))
        con.execute("INSERT INTO promo_activations(user_id, code) VALUES(?,?)", (uid, code))
        
        if creator_id and creator_id != uid:
            con.execute("UPDATE users SET aura_points = aura_points + 200 WHERE telegram_id=?", (creator_id,))
            try: await bot.send_message(creator_id, f"✨ По твоему промокоду <code>{code}</code> зарегистрировались! Тебе начислено <b>+200 Ауры</b>.", parse_mode="HTML")
            except Exception: pass
            
        con.commit()
        await message.answer(f"🎉 Промокод активирован! Начислено {days} дней Премиума.", parse_mode="HTML")
        
        try:
            notify_text = (
                f"🎁 <b>Активация промокода!</b>\n\n"
                f"Код: <code>{code}</code>\n"
                f"👤 Юзер: {message.from_user.full_name} (@{message.from_user.username or '—'})\n"
                f"🆔 ID: <code>{uid}</code>"
            )
            await bot.send_message(DEVELOPER_ID, notify_text, parse_mode="HTML")
        except Exception: pass
        
    finally: con.close()

# ---------------------- БАГ-РЕПОРТЫ ----------------------
@dp.message(Command("bug"))
@dp.message(F.text == "🐛 Баг-репорт")
async def cmd_bug(message: types.Message):
    if message.text == "🐛 Баг-репорт":
        return await message.answer("❌ Напиши: `/bug Текст проблемы`", parse_mode="Markdown")
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2: return await message.answer("❌ `/bug Текст проблемы`", parse_mode="Markdown")
    bug_text = parts[1].strip()
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        cur = con.execute("INSERT INTO bug_reports(user_id, text, created_at) VALUES(?,?,?)", (message.from_user.id, bug_text, now_utc().strftime("%Y-%m-%d %H:%M:%S")))
        bid = cur.lastrowid; con.commit()
    finally: con.close()
    await message.answer(f"✅ Баг #{bid} отправлен!")
    for aid in ADMIN_IDS:
        try: await bot.send_message(aid, f"🐛 <b>Баг #{bid}</b>\nID: {message.from_user.id}\n{html.escape(bug_text)}", parse_mode="HTML")
        except Exception: pass

@dp.message(Command("bug_list"))
async def cmd_bug_list(message: types.Message):
    if not is_admin(message.from_user.id): return
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: bugs = con.execute("SELECT id, user_id, text FROM bug_reports WHERE status='open' ORDER BY id DESC LIMIT 20").fetchall()
    finally: con.close()
    if not bugs: return await message.answer("Открытых багов нет.")
    for bid, uid, text in bugs: await message.answer(f"🐛 <b>#{bid}</b> | {uid}\n{html.escape(text)}", parse_mode="HTML")

@dp.message(Command("logs"))
async def cmd_logs(message: types.Message):
    if not is_admin(message.from_user.id): return
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: rows = con.execute("SELECT admin_id, action, created_at FROM logs ORDER BY id DESC LIMIT 30").fetchall()
    finally: con.close()
    if not rows: return await message.answer("Логов нет.")
    lines = [f"<i>{dt}</i> — {aid}: {html.escape(act)}" for aid, act, dt in rows]
    await message.answer("📋 <b>Действия:</b>\n\n" + "\n".join(lines), parse_mode="HTML")

# ---------------------- ЮЗЕРЫ (АДМИН) ----------------------
@dp.message(Command("user"))
async def cmd_user(message: types.Message):
    if not is_admin(message.from_user.id): return
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        rows = con.execute("SELECT telegram_id, first_name, username, last_seen FROM users ORDER BY last_seen DESC").fetchall()
    finally: con.close()
    
    total = len(rows)
    if total == 0:
        return await message.answer("👥 В базе пока нет пользователей.")
        
    await message.answer(f"👥 <b>Всего пользователей:</b> {total}\nВывожу список в чат...", parse_mode="HTML")
    
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="📥 Скачать всех в Excel", callback_data="export_users_excel")]])
    await message.answer("Хочешь скачать полный список файлом?", reply_markup=kb)
    
    chunk = []
    current_length = 0
    
    for uid, fname, uname, last_seen in rows:
        u_str = f"@{uname}" if uname else "—"
        line = f"<code>{uid}</code> | {html.escape(fname or '—')} ({u_str})\n"
        
        if current_length + len(line) > 3800:
            await message.answer("".join(chunk), parse_mode="HTML")
            chunk = [line]
            current_length = len(line)
        else:
            chunk.append(line)
            current_length += len(line)
            
    if chunk:
        await message.answer("".join(chunk), parse_mode="HTML")

@dp.callback_query(F.data == "export_users_excel")
async def cb_export_users_excel(callback: types.CallbackQuery):
    await callback.answer()
    if not is_admin(callback.from_user.id): return
    
    await callback.message.answer("⏳ Собираю базу пользователей...")
    wb = Workbook()
    ws = wb.active
    ws.title = "Пользователи"
    ws.append(["ID", "Имя", "Юзернейм", "Первый визит", "Последний визит", "Сообщений", "Премиум до", "Аура"])
    
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: 
        rows = con.execute("SELECT telegram_id, first_name, username, first_seen, last_seen, total_messages, premium_until, aura_points FROM users ORDER BY last_seen DESC").fetchall()
    finally: con.close()
    
    for r in rows:
        ws.append(r)
        
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    await callback.message.answer_document(BufferedInputFile(out.read(), filename="all_users.xlsx"), caption=f"📊 Всего пользователей: {len(rows)}")

@dp.message(Command("user_list"))
async def cmd_user_list(message: types.Message):
    if not is_admin(message.from_user.id): return
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: rows = con.execute("SELECT telegram_id, first_name, username FROM users ORDER BY last_seen DESC LIMIT 50").fetchall()
    finally: con.close()
    if not rows: return await message.answer("Юзеров нет.")
    lines = ["👥 <b>Последние 50 юзеров:</b>\n"]
    for uid, fname, uname in rows:
        u_str = f"@{uname}" if uname else "—"
        lines.append(f"<code>{uid}</code> | {html.escape(fname or '—')} ({u_str})")
    await message.answer("\n".join(lines), parse_mode="HTML")

# ---------------------- ЭКСПОРТ ----------------------
@dp.message(Command("export_excel"))
async def cmd_export_excel(message: types.Message):
    if not is_admin(message.from_user.id): return
    await message.answer("⏳ Собираю Excel...")
    wb = Workbook()
    ws1 = wb.active; ws1.title = "Прогнозы"; ws1.append(["ID", "Юзер", "Бой", "Победитель", "Метод", "Статус"])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: preds = con.execute("SELECT id, telegram_id, fight_id, predicted_winner, predicted_method, status FROM predictions").fetchall()
    finally: con.close()
    for p in preds: ws1.append(p)
    out = io.BytesIO(); wb.save(out); out.seek(0)
    await message.answer_document(BufferedInputFile(out.read(), filename="predictions.xlsx"), caption="📊 База прогнозов")

@dp.callback_query(F.data == "prem_export_preds")
async def cb_prem_export_preds(callback: types.CallbackQuery):
    await callback.answer()
    if not is_premium(callback.from_user.id): return await callback.message.answer("❌ Только для Премиум.")
    await callback.message.answer("⏳ Собираю твои прогнозы...")
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        rows = con.execute("SELECT t.name, f.fighter_1, f.fighter_2, p.predicted_winner, p.predicted_method, p.status FROM predictions p JOIN fights f ON p.fight_id = f.id JOIN tournaments t ON f.tournament_id = t.id WHERE p.telegram_id = ? ORDER BY t.date_utc DESC", (callback.from_user.id,)).fetchall()
    finally: con.close()
    if not rows: return await callback.message.answer("У тебя пока нет прогнозов.")
    wb = Workbook(); ws = wb.active; ws.title = "Мои прогнозы"
    ws.append(["Турнир", "Боец 1", "Боец 2", "Мой прогноз", "Метод", "Статус"])
    for t_name, f1, f2, pred, method, status in rows:
        status_text = "⏳ Ожидание" if status == 'pending' else ("✅ Угадал" if status == 'won' else "❌ Не угадал")
        ws.append([t_name, f1, f2, pred, method or "—", status_text])
    out = io.BytesIO(); wb.save(out); out.seek(0)
    await callback.message.answer_document(BufferedInputFile(out.read(), filename="my_predictions.xlsx"), caption="📊 Твои прогнозы")

# ---------------------- ПОИСК И КАЛЕНДАРЬ ----------------------
@dp.message(Command("search"))
@dp.message(F.text == "🔍 Поиск бойца")
async def cmd_search(message: types.Message, state: FSMContext):
    await state.clear()
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2: return await message.answer("🔍 Напиши `/search Имя`", parse_mode="Markdown")
    query = parts[1].strip()
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: fights = con.execute("SELECT DISTINCT f.fighter_1, f.fighter_2 FROM fights f WHERE f.fighter_1 LIKE ? OR f.fighter_2 LIKE ?", (f"%{query}%", f"%{query}%")).fetchall()
    finally: con.close()
    names = set()
    for f1, f2 in fights:
        if query.lower() in f1.lower(): names.add(f1)
        if query.lower() in f2.lower(): names.add(f2)
    if not names: return await message.answer(f"❌ Никого не нашли.")
    buttons = [[InlineKeyboardButton(text=n, callback_data=f"fighter:{n}")] for n in sorted(names)]
    await message.answer("🔍 <b>Найдено:</b>", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")

@dp.callback_query(F.data.startswith("fighter:"))
async def cb_fighter_profile(callback: types.CallbackQuery):
    await callback.answer()
    name = callback.data.split(":", 1)[1]
    stats = get_fighter_stats(name)
    premium = is_premium(callback.from_user.id)
    text = f"🥊 <b>{html.escape(name)}</b>\n\n"
    if stats["total"] == 0: text += "📊 Боёв нет."
    else:
        text += f"📊 <b>Статистика:</b>\n• Боёв: {stats['total']}\n• Побед: {stats['wins']} ({stats['win_pct']}%)\n• Поражений: {stats['losses']}\n"
        if premium and stats["wins"] > 0:
            methods_str = ", ".join(f"{m} — {c}" for m, c in stats["methods"].most_common())
            text += f"\n⭐ <b>Расширенная:</b>\n• Методы: {methods_str or '—'}\n• Ср. раунд: {stats['avg_round']}\n• Титульных: {stats['title_wins']}\n"
        elif not premium: text += f"\n<i>💳 Расширенная статистика доступна с Премиумом</i>"
    await callback.message.answer(safe_truncate_html(text, 4000), parse_mode="HTML")

@dp.message(Command("calendar"))
@dp.message(F.text == "📅 Календарь")
async def cmd_calendar(message: types.Message, state: FSMContext):
    await state.clear()
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: tours = con.execute("SELECT id, name, date_utc, location, event_type FROM tournaments WHERE status='announced' ORDER BY date_utc ASC LIMIT 20").fetchall()
    finally: con.close()
    if not tours: return await message.answer("📅 Турниров пока нет.")
    tz = get_user_tz(message.from_user.id)
    lines = ["📅 <b>Календарь:</b>\n"]
    for tid, name, date_utc, loc, etype in tours:
        local = fmt_utc_to_local(date_utc, tz)
        emoji = "🏆" if etype == "numbered" else "🥊"
        lines.append(f"{emoji} <b>{html.escape(name)}</b>\n📅 {local}\n📍 {html.escape(loc or '—')}\n")
    await message.answer(safe_truncate_html("\n".join(lines), 4000), parse_mode="HTML")

@dp.message(Command("top_predictors"))
@dp.message(F.text == "📈 Топ прогнозов")
async def cmd_top_predictors(message: types.Message):
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        rows = con.execute("""
            SELECT telegram_id,
                   SUM(CASE WHEN status='won' THEN 1 ELSE 0 END) as won,
                   SUM(CASE WHEN status IN ('won','lost') THEN 1 ELSE 0 END) as total
            FROM predictions GROUP BY telegram_id HAVING total > 0
            ORDER BY won DESC, total ASC LIMIT 10
        """).fetchall()
    finally: con.close()
    if not rows: return await message.answer("🎯 Прогнозов пока нет.")
    lines = ["🎯 <b>Топ прогнозистов:</b>\n"]
    for i, (uid, won, total) in enumerate(rows, 1):
        pct = round(won / total * 100)
        med = {1: "🥇", 2: "🥈", 3: "🥉"}.get(i, f"{i}.")
        badge = " 👑" if uid == DEVELOPER_ID else (" 🛠" if is_admin(uid) else (" ⭐" if is_premium(uid) else ""))
        lines.append(f"{med} ID <code>{uid}</code>{badge}: {won}/{total} ({pct}%)")
    await message.answer("\n".join(lines), parse_mode="HTML")

# ---------------------- ПРЕМИУМ И АДМИНЫ ----------------------
@dp.message(Command("add_admin"))
async def cmd_add_admin(message: types.Message):
    if not is_admin(message.from_user.id): return
    target_id = message.reply_to_message.from_user.id if message.reply_to_message else None
    if not target_id:
        parts = message.text.split()
        if len(parts) > 1 and parts[1].isdigit(): target_id = int(parts[1])
    if not target_id: return await message.answer("❌ `/add_admin ID`", parse_mode="Markdown")
    if target_id in ADMIN_IDS: return await message.answer("ℹ️ Уже админ.")
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: con.execute("INSERT OR IGNORE INTO admins(telegram_id) VALUES(?)", (target_id,)); con.commit()
    finally: con.close()
    ADMIN_IDS.add(target_id)
    await message.answer(f"✅ {target_id} теперь админ!")

@dp.message(Command("del_admin"))
async def cmd_del_admin(message: types.Message):
    if message.from_user.id != DEVELOPER_ID: 
        return await message.answer("❌ Эта команда доступна только Разработчику.")
        
    target_id = message.reply_to_message.from_user.id if message.reply_to_message else None
    if not target_id:
        parts = message.text.split()
        if len(parts) > 1 and parts[1].isdigit(): target_id = int(parts[1])
    if not target_id: return await message.answer("❌ `/del_admin ID` (или ответом на сообщение)", parse_mode="Markdown")
    
    if target_id == DEVELOPER_ID:
        return await message.answer("❌ Разработчика нельзя удалить из админов!")
        
    if target_id not in ADMIN_IDS: return await message.answer("ℹ️ Этот пользователь не админ.")
    
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: 
        con.execute("DELETE FROM admins WHERE telegram_id=?", (target_id,))
        con.commit()
    finally: con.close()
    
    if target_id in ADMIN_IDS: ADMIN_IDS.remove(target_id)
    
    await message.answer(f"✅ {target_id} больше не админ.")

@dp.message(Command("give_premium"))
async def cmd_give_premium(message: types.Message):
    if not is_admin(message.from_user.id): return
    parts = message.text.split()
    if len(parts) < 3 or not parts[1].isdigit() or not parts[2].isdigit(): return await message.answer("❌ `/give_premium ID ДНИ`", parse_mode="Markdown")
    target_id, days = int(parts[1]), int(parts[2])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        row = con.execute("SELECT premium_until FROM users WHERE telegram_id=?", (target_id,)).fetchone()
        if not row: return await message.answer("❌ Нет такого юзера.")
        current = now_utc()
        if row[0]:
            try:
                old = datetime.strptime(row[0], "%Y-%m-%d %H:%M:%S")
                if old > current: current = old
            except ValueError: pass
        new_date = current + timedelta(days=days)
        con.execute("UPDATE users SET premium_until=? WHERE telegram_id=?", (new_date.strftime("%Y-%m-%d %H:%M:%S"), target_id))
        con.commit()
    finally: con.close()
    await message.answer(f"✅ Премиум до {new_date.strftime('%d.%m.%Y')}.")
    try: await bot.send_message(target_id, f"⭐ Тебе активирован Премиум на {days} дней!")
    except Exception: pass

@dp.message(Command("premium_list"))
async def cmd_premium_list(message: types.Message):
    if not is_admin(message.from_user.id): return
    lines = []
    
    if PREMIUM_IDS:
        lines.append("⭐ <b>VIP (навсегда):</b>")
        for uid in PREMIUM_IDS:
            con = sqlite3.connect(DB_PATH, timeout=10)
            try: row = con.execute("SELECT username, first_name FROM users WHERE telegram_id=?", (uid,)).fetchone()
            finally: con.close()
            if row:
                uname = f" @{row[0]}" if row[0] else ""
                fname = f" ({html.escape(row[1])})" if row[1] else ""
                lines.append(f"  • <code>{uid}</code>{uname}{fname}")
            else:
                lines.append(f"  • <code>{uid}</code> (не заходил)")
                
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: rows = con.execute("SELECT telegram_id, premium_until, username, first_name FROM users WHERE premium_until > ? ORDER BY premium_until DESC", (now_utc().strftime("%Y-%m-%d %H:%M:%S"),)).fetchall()
    finally: con.close()
    
    if rows:
        lines.append("\n⏱ <b>Временные:</b>")
        for uid, until, uname, fname in rows:
            try: 
                dt = datetime.strptime(until, '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')
                u_str = f" @{uname}" if uname else ""
                n_str = f" ({html.escape(fname)})" if fname else ""
                lines.append(f"  • <code>{uid}</code>{u_str}{n_str} — до {dt}")
            except ValueError: pass
            
    if not lines: return await message.answer("Премиум-юзеров нет.")
    await message.answer("\n".join(lines), parse_mode="HTML")

@dp.message(Command("premium"))
@dp.message(F.text == "💳 Купить Премиум")
async def cmd_premium(message: types.Message, state: FSMContext):
    await state.clear()
    status = get_premium_status_text(message.from_user.id)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔔 Мои подписки", callback_data="prem_subs")],
        [InlineKeyboardButton(text="📈 Мои прогнозы", callback_data="prem_my_stats")],
        [InlineKeyboardButton(text="📊 Скачать прогнозы (Excel)", callback_data="prem_export_preds")],
        [InlineKeyboardButton(text="🎁 Ввести промокод", callback_data="prem_enter_promo")],
        [InlineKeyboardButton(text="💬 Написать @sultanovg", url="https://t.me/sultanovg")]
    ])
    await message.answer(
        f"⭐ <b>Премиум-подписка UFC Cards</b>\n\nВаш статус: {status}\n\n"
        f"💳 <b>Стоимость:</b> 15 ₽ / 30 дней\n"
        f"📩 <b>Для покупки:</b> @sultanovg\n\n"
        f"🔥 <b>Функции:</b>\n"
        f"┣ 🎯 <b>Прогнозы с методом</b> (KO/Sub/Dec)\n"
        f"┣ 🔔 <b>Подписки на бойцов</b>\n"
        f"┣ ⚡ <b>Ранний доступ к результатам</b> (+10 мин)\n"
        f"┣ 🔮 <b>Мнение сообщества</b>\n"
        f"┣ 📊 <b>Экспорт в Excel</b>\n"
        f"┗ 👑 <b>VIP-новости и значок</b>", reply_markup=kb, parse_mode="HTML")

@dp.callback_query(F.data == "prem_my_stats")
async def cb_prem_my_stats(callback: types.CallbackQuery):
    await callback.answer()
    if not is_premium(callback.from_user.id): return await callback.message.answer("❌ Только для Премиум.")
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: rows = con.execute("SELECT predicted_winner, predicted_method, status FROM predictions WHERE telegram_id=?", (callback.from_user.id,)).fetchall()
    finally: con.close()
    if not rows: return await callback.message.answer("Ты еще не делал прогнозов.")
    total = len(rows); won = sum(1 for _, m, st in rows if st == "won"); lost = sum(1 for _, m, st in rows if st == "lost"); pending = sum(1 for _, m, st in rows if st == "pending")
    perfect = sum(1 for _, m, st in rows if st == "won" and m is not None)
    pct = round(won / (won + lost) * 100) if (won + lost) > 0 else 0
    fighters = Counter([r[0] for r in rows if r[0]])
    top_fighters = "\n".join(f"• {f} ({c} раз)" for f, c in fighters.most_common(3))
    text = (f"📈 <b>Статистика:</b>\n\n• Всего: {total}\n• Угадано: {won} ✅\n• Идеально (с методом): {perfect} 🎯\n• Не угадано: {lost} ❌\n• Ожидание: {pending} ⏳\n• Точность: <b>{pct}%</b>\n\n🥊 <b>Любимчики:</b>\n{top_fighters or '—'}")
    await callback.message.answer(text, parse_mode="HTML")

@dp.message(Command("subscribe"))
async def cmd_subscribe(message: types.Message):
    if not is_premium(message.from_user.id): return await message.answer("❌ Только для Премиум.")
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2: return await message.answer("❌ `/subscribe Имя`", parse_mode="Markdown")
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: con.execute("INSERT OR IGNORE INTO subscriptions(telegram_id, fighter_name) VALUES(?,?)", (message.from_user.id, parts[1].strip())); con.commit()
    finally: con.close()
    await message.answer(f"✅ Подписка на <b>{html.escape(parts[1].strip())}</b>!", parse_mode="HTML")

@dp.message(Command("unsubscribe"))
async def cmd_unsubscribe(message: types.Message):
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2: return await message.answer("❌ `/unsubscribe Имя`", parse_mode="Markdown")
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: con.execute("DELETE FROM subscriptions WHERE telegram_id=? AND fighter_name=?", (message.from_user.id, parts[1].strip())); con.commit()
    finally: con.close()
    await message.answer(f"🗑 Подписка удалена.")

@dp.callback_query(F.data == "prem_subs")
async def cb_prem_subs(callback: types.CallbackQuery):
    await callback.answer()
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: subs = con.execute("SELECT fighter_name FROM subscriptions WHERE telegram_id=?", (callback.from_user.id,)).fetchall()
    finally: con.close()
    if not subs: text = "Подписок нет."
    else: text = "🔔 <b>Твои подписки:</b>\n" + "\n".join([f"• {html.escape(s[0])}" for s in subs])
    await callback.message.answer(text, parse_mode="HTML")

def get_tournament_fighters(tid: int) -> list:
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: fights = con.execute("SELECT fighter_1, fighter_2 FROM fights WHERE tournament_id=?", (tid,)).fetchall()
    finally: con.close()
    return sorted(list(set([f.strip() for pair in fights for f in pair if f])))

@dp.callback_query(F.data.startswith("sub_list:"))
async def cb_sub_list(callback: types.CallbackQuery):
    if not is_premium(callback.from_user.id): return await callback.answer("❌ Только для Премиум!", show_alert=True)
    await callback.answer()
    tid = int(callback.data.split(":")[1])
    fighters = get_tournament_fighters(tid)
    if not fighters: return await callback.message.answer("Нет бойцов.")
    buttons = []
    for i in range(0, len(fighters), 2):
        row = [InlineKeyboardButton(text=fighters[i], callback_data=f"sub_do:{tid}:{i}")]
        if i + 1 < len(fighters): row.append(InlineKeyboardButton(text=fighters[i+1], callback_data=f"sub_do:{tid}:{i+1}"))
        buttons.append(row)
    buttons.append([InlineKeyboardButton(text="⬅️ Назад", callback_data=f"show_tour:{tid}:0")])
    await callback.message.answer("Выбери бойца:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))

@dp.callback_query(F.data.startswith("sub_do:"))
async def cb_sub_do(callback: types.CallbackQuery):
    if not is_premium(callback.from_user.id): return await callback.answer("❌ Только для Премиум!", show_alert=True)
    parts = callback.data.split(":")
    tid, idx = int(parts[1]), int(parts[2])
    fighters = get_tournament_fighters(tid)
    if idx >= len(fighters): return await callback.answer("Не найден.", show_alert=True)
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: con.execute("INSERT OR IGNORE INTO subscriptions(telegram_id, fighter_name) VALUES(?,?)", (callback.from_user.id, fighters[idx])); con.commit()
    finally: con.close()
    await callback.answer(f"✅ Подписан на {fighters[idx]}!", show_alert=True)

# ---------------------- НОВОСТИ (С МЕДИА/ОПРОСАМИ) ----------------------
@dp.message(Command("news_add"))
async def cmd_news_add(message: types.Message):
    if not is_admin(message.from_user.id): return
    
    if message.reply_to_message:
        con = sqlite3.connect(DB_PATH, timeout=10)
        try: cur = con.execute("INSERT INTO news(text, created_at) VALUES(?,?)", ("[Медиа/Опрос]", now_utc().strftime("%Y-%m-%d %H:%M:%S"))); nid = cur.lastrowid; con.commit()
        finally: con.close()
        await message.answer(f"✅ Новость #{nid}.\n⏳ Рассылка...")
        s, f = await broadcast_to_all(copy_msg=message.reply_to_message, exclude_id=message.from_user.id)
        await message.answer(f"✅ Доставлено: {s}\nБлок: {f}")
        return

    if message.photo: photo_id = message.photo[-1].file_id; text = (message.caption or "").replace("/news_add", "").strip()
    else: photo_id = None; text = (message.text or "").replace("/news_add", "").strip()
    if not text and not photo_id: return await message.answer("❌ `/news_add Текст` или ответь на сообщение.")
    
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: cur = con.execute("INSERT INTO news(text, photo_id, created_at) VALUES(?,?,?)", (text, photo_id, now_utc().strftime("%Y-%m-%d %H:%M:%S"))); nid = cur.lastrowid; con.commit()
    finally: con.close()
    await message.answer(f"✅ Новость #{nid}.\n⏳ Рассылка...")
    s, f = await broadcast_to_all(f"📰 <b>НОВОСТЬ #{nid}</b>\n\n{text}", photo_id, exclude_id=message.from_user.id)
    await message.answer(f"✅ Доставлено: {s}\nБлок: {f}")

@dp.message(Command("vip_news"))
async def cmd_vip_news(message: types.Message):
    if not is_admin(message.from_user.id): return
    
    if message.reply_to_message:
        con = sqlite3.connect(DB_PATH, timeout=10)
        try: cur = con.execute("INSERT INTO news(text, is_vip, created_at) VALUES(?,1,?)", ("[Медиа/Опрос]", now_utc().strftime("%Y-%m-%d %H:%M:%S"))); nid = cur.lastrowid; con.commit()
        finally: con.close()
        await message.answer(f"✅ VIP-новость #{nid}.\n⏳ Рассылка...")
        s, f = await broadcast_to_all(copy_msg=message.reply_to_message, only_premium=True, exclude_id=message.from_user.id)
        await message.answer(f"✅ Доставлено VIP: {s}\nБлок: {f}")
        return

    if message.photo: photo_id = message.photo[-1].file_id; text = (message.caption or "").replace("/vip_news", "").strip()
    else: photo_id = None; text = (message.text or "").replace("/vip_news", "").strip()
    if not text and not photo_id: return await message.answer("❌ `/vip_news Текст` или ответь на сообщение.")
    
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: cur = con.execute("INSERT INTO news(text, photo_id, is_vip, created_at) VALUES(?,?,1,?)", (text, photo_id, now_utc().strftime("%Y-%m-%d %H:%M:%S"))); nid = cur.lastrowid; con.commit()
    finally: con.close()
    await message.answer(f"✅ VIP-новость #{nid}.\n⏳ Рассылка...")
    s, f = await broadcast_to_all(f"👑 <b>СЛИВКИ #{nid}</b>\n\n{text}", photo_id, only_premium=True, exclude_id=message.from_user.id)
    await message.answer(f"✅ Доставлено VIP: {s}\nБлок: {f}")

@dp.message(Command("news_delete"))
async def cmd_news_delete(message: types.Message):
    if not is_admin(message.from_user.id): return
    parts = message.text.split()
    if len(parts) < 2: return await message.answer("❌ `/news_delete ID`")
    nid = int(parts[1])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: con.execute("DELETE FROM news WHERE id=?", (nid,)); con.commit()
    finally: con.close()
    await message.answer(f"✅ Новость #{nid} удалена.")

@dp.message(Command("news"))
@dp.message(F.text == "📰 Новости")
async def cmd_news_view(message: types.Message, state: FSMContext):
    await state.clear()
    premium = is_premium(message.from_user.id)
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: rows = con.execute("SELECT id, text, photo_id, is_vip, created_at FROM news ORDER BY id DESC LIMIT 10").fetchall()
    finally: con.close()
    if not rows: return await message.answer("📰 Новостей нет.")
    await message.answer(f"📰 <b>Последние новости:</b>", parse_mode="HTML")
    for nid, text, photo_id, is_vip, dt in rows[:5]:
        if is_vip and not premium: continue
        prefix = "👑 <b>СЛИВКИ</b>" if is_vip else "📰"
        cap = f"{prefix} <b>#{nid}</b> · <i>{dt}</i>\n\n{html.escape(text or '')}"
        try:
            if photo_id: await message.answer_photo(photo=photo_id, caption=cap[:1024], parse_mode="HTML")
            else: await message.answer(cap[:4000], parse_mode="HTML")
        except Exception: pass

# ---------------------- ПРОГНОЗЫ (С МЕТОДОМ) ----------------------
@dp.callback_query(F.data.startswith("pred_list:"))
async def cb_pred_list(callback: types.CallbackQuery):
    await callback.answer()
    tid = int(callback.data.split(":")[1])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: fights = con.execute("SELECT id, fighter_1, fighter_2 FROM fights WHERE tournament_id=? AND status='scheduled'", (tid,)).fetchall()
    finally: con.close()
    if not fights: return await callback.message.answer("Нет доступных боёв.")
    buttons = [[InlineKeyboardButton(text=f"{f1} vs {f2}", callback_data=f"pred_fight:{fid}")] for fid, f1, f2 in fights[:10]]
    buttons.append([InlineKeyboardButton(text="⬅️ Назад", callback_data=f"show_tour:{tid}:0")])
    await callback.message.answer("🎯 Выбери бой:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))

@dp.callback_query(F.data.startswith("pred_fight:"))
async def cb_pred_fight(callback: types.CallbackQuery):
    await callback.answer()
    fid = int(callback.data.split(":")[1])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        fight = con.execute("SELECT fighter_1, fighter_2, status, tournament_id FROM fights WHERE id=?", (fid,)).fetchone()
        existing = con.execute("SELECT predicted_winner, predicted_method FROM predictions WHERE telegram_id=? AND fight_id=?", (callback.from_user.id, fid)).fetchone()
    finally: con.close()
    if not fight: return await callback.message.answer("Бой не найден.")
    if fight[2] == "finished": return await callback.message.answer("Бой уже завершен.")
    f1, f2, tid = fight[0], fight[1], fight[3]
    header = f"🥊 <b>{html.escape(f1)} vs {html.escape(f2)}</b>\n\n"
    if existing:
        m_str = f" (Метод: {existing[1]})" if existing[1] else ""
        header += f"⚠️ Ты поставил на: <b>{html.escape(existing[0])}</b>{m_str}\nМожешь изменить прогноз.\n\n"
    else:
        header += "Кто победит?\n\n"
    kb_btns = []
    if is_premium(callback.from_user.id):
        kb_btns.append([InlineKeyboardButton(text="👀 Мнение сообщества", callback_data=f"pred_stats:{fid}")])
    kb_btns.append([InlineKeyboardButton(text=f"🥊 {f1}", callback_data=f"pred_win:{fid}:1")])
    kb_btns.append([InlineKeyboardButton(text=f"🥊 {f2}", callback_data=f"pred_win:{fid}:2")])
    if existing:
        kb_btns.append([InlineKeyboardButton(text="🗑 Удалить прогноз", callback_data=f"pred_del:{fid}")])
    kb_btns.append([InlineKeyboardButton(text="⬅️ К списку", callback_data=f"pred_list:{tid}")])
    kb = InlineKeyboardMarkup(inline_keyboard=kb_btns)
    try: await callback.message.edit_text(header, reply_markup=kb, parse_mode="HTML")
    except TelegramBadRequest: pass

@dp.callback_query(F.data.startswith("pred_win:"))
async def cb_pred_win(callback: types.CallbackQuery):
    parts = callback.data.split(":"); fid, who = int(parts[1]), int(parts[2])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        fight = con.execute("SELECT fighter_1, fighter_2, status, tournament_id FROM fights WHERE id=?", (fid,)).fetchone()
        if not fight: return await callback.answer("Не найден.", show_alert=True)
        if fight[2] == "finished": return await callback.answer("Бой уже завершен.", show_alert=True)
        f1, f2, tid = fight[0], fight[1], fight[3]
        winner = f1 if who == 1 else f2
        
        if is_premium(callback.from_user.id):
            kb_btns = [
                [InlineKeyboardButton(text="💥 KO / TKO", callback_data=f"pred_method:{fid}:{who}:KO")],
                [InlineKeyboardButton(text="🤼 Сабмишен", callback_data=f"pred_method:{fid}:{who}:SUB")],
                [InlineKeyboardButton(text="⚖️ Решение", callback_data=f"pred_method:{fid}:{who}:DEC")],
                [InlineKeyboardButton(text="⬅️ Назад", callback_data=f"pred_fight:{fid}")]
            ]
            await callback.message.edit_text(
                f"🥊 <b>{html.escape(f1)} vs {html.escape(f2)}</b>\n\n"
                f"Ты выбрал победу: <b>{html.escape(winner)}</b>\n\n"
                f"⭐ <b>Премиум:</b> Как он победит?",
                reply_markup=InlineKeyboardMarkup(inline_keyboard=kb_btns), parse_mode="HTML"
            )
            await callback.answer()
        else:
            con.execute(
                "INSERT INTO predictions(telegram_id, fight_id, predicted_winner) VALUES(?,?,?) "
                "ON CONFLICT(telegram_id, fight_id) DO UPDATE SET predicted_winner=excluded.predicted_winner",
                (callback.from_user.id, fid, winner)
            )
            con.commit()
            await callback.answer(f"✅ Поставил на: {winner}", show_alert=True)
            await cb_pred_fight(callback)
    finally: con.close()

@dp.callback_query(F.data.startswith("pred_method:"))
async def cb_pred_method(callback: types.CallbackQuery):
    parts = callback.data.split(":")
    fid, who, method = int(parts[1]), int(parts[2]), parts[3]
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        fight = con.execute("SELECT fighter_1, fighter_2 FROM fights WHERE id=?", (fid,)).fetchone()
        if not fight: return await callback.answer("Не найден.", show_alert=True)
        winner = fight[0] if who == 1 else fight[1]
        
        con.execute(
            "INSERT INTO predictions(telegram_id, fight_id, predicted_winner, predicted_method) VALUES(?,?,?,?) "
            "ON CONFLICT(telegram_id, fight_id) DO UPDATE SET predicted_winner=excluded.predicted_winner, predicted_method=excluded.predicted_method",
            (callback.from_user.id, fid, winner, method)
        )
        con.commit()
    finally: con.close()
    await callback.answer(f"✅ Поставил на {winner} ({method})", show_alert=True)
    await cb_pred_fight(callback)

@dp.callback_query(F.data.startswith("pred_stats:"))
async def cb_pred_stats(callback: types.CallbackQuery):
    if not is_premium(callback.from_user.id): return await callback.answer("❌ Только для Премиум!", show_alert=True)
    await callback.answer()
    fid = int(callback.data.split(":")[1])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: stats = con.execute("SELECT predicted_winner, COUNT(*) FROM predictions WHERE fight_id=? GROUP BY predicted_winner", (fid,)).fetchall()
    finally: con.close()
    total = sum(c for _, c in stats) or 1
    text = f"👀 <b>Мнение сообщества:</b>\n\n"
    for fighter, count in stats:
        pct = round(count / total * 100)
        bar = "🟩" * (pct // 10) + "⬜️" * (10 - pct // 10)
        text += f"{fighter}: {pct}% ({count} чел.)\n{bar}\n\n"
    if not stats: text += "Пока никто не ставил."
    await callback.message.answer(text, parse_mode="HTML")

@dp.callback_query(F.data.startswith("pred_del:"))
async def cb_pred_del(callback: types.CallbackQuery):
    fid = int(callback.data.split(":")[1])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: con.execute("DELETE FROM predictions WHERE telegram_id=? AND fight_id=?", (callback.from_user.id, fid)); con.commit()
    finally: con.close()
    await callback.answer("🗑 Прогноз удалён!", show_alert=True)
    await cb_pred_fight(callback)

# ---------------------- ЖИВОЙ ЧАТ + РАССЫЛКА ----------------------
@dp.message(F.from_user.id.in_(ADMIN_IDS) & F.reply_to_message)
async def admin_reply_message(message: types.Message):
    replied = message.reply_to_message
    if not replied.from_user or not replied.from_user.is_bot: return
    text = replied.text or replied.caption or ""
    match = re.search(r"🆔 ID: (\d+)", text)
    if match:
        target_id = int(match.group(1))
        try:
            if message.photo: await bot.send_photo(target_id, message.photo[-1].file_id, caption=f"👨‍💼 Ответ:\n\n{message.caption or ''}"[:1024])
            else: await bot.send_message(target_id, f"👨‍💼 Ответ:\n\n{message.text or '(пусто)'}")
            await message.answer("✅ Отправлено!")
        except Exception as e: await message.answer(f"❌ {e}")

@dp.message(Command("broadcast"))
async def cmd_broadcast(message: types.Message):
    if not is_admin(message.from_user.id): return
    
    if message.reply_to_message:
        target = message.reply_to_message
        await message.answer("⏳ Рассылка медиа/опроса...")
        s, f = await broadcast_to_all(copy_msg=target, exclude_id=message.from_user.id)
        await message.answer(f"✅ Доставлено: {s}\nБлок: {f}")
        return
        
    if message.photo: 
        photo_id = message.photo[-1].file_id
        text = (message.caption or "").replace("/broadcast", "").strip()
    else: 
        photo_id = None
        text = (message.text or "").replace("/broadcast", "").strip()
        
    if not text and not photo_id: return await message.answer("❌ `/broadcast Текст` или ответь на сообщение.")
    await message.answer("⏳ Рассылка...")
    s, f = await broadcast_to_all(text, photo_id, exclude_id=message.from_user.id)
    await message.answer(f"✅ Доставлено: {s}\nБлок: {f}")

@dp.message(Command("send_id"))
async def cmd_send_id(message: types.Message):
    if not is_admin(message.from_user.id): return
    parts = message.text.split(maxsplit=2)
    if len(parts) < 3 or not parts[1].isdigit(): return await message.answer("❌ `/send_id ID Текст`", parse_mode="Markdown")
    try: await bot.send_message(int(parts[1]), f"👨‍💼 {parts[2]}"); await message.answer("✅ Отправлено!")
    except Exception as e: await message.answer(f"❌ {e}")

# ---------------------- СТАТИСТИКА И БЭКАП ----------------------
@dp.message(Command("stats"))
async def cmd_stats(message: types.Message):
    if not is_admin(message.from_user.id): return
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        users = con.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        tours = con.execute("SELECT COUNT(*) FROM tournaments").fetchone()[0]
        fights = con.execute("SELECT COUNT(*) FROM fights").fetchone()[0]
        prem = con.execute("SELECT COUNT(*) FROM users WHERE premium_until > ?", (now_utc().strftime("%Y-%m-%d %H:%M:%S"),)).fetchone()[0]
    finally: con.close()
    await message.answer(f"📊 <b>Статистика</b>\n\n👥 Юзеров: {users}\n⭐ Премиум: {prem + len(PREMIUM_IDS)}\n🏆 Турниров: {tours}\n🥊 Боёв: {fights}", parse_mode="HTML")

@dp.message(Command("backup"))
async def cmd_backup(message: types.Message):
    if not is_admin(message.from_user.id): return
    if not os.path.exists(DB_PATH): return await message.answer("❌ Файл не найден.")
    try: await message.answer_document(FSInputFile(DB_PATH), caption="📦 Бэкап базы")
    except Exception as e: await message.answer(f"❌ {e}")

@dp.message(Command("archive"))
async def cmd_archive(message: types.Message):
    if not is_admin(message.from_user.id): return
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: con.execute("UPDATE tournaments SET status='archived' WHERE status='finished'"); con.commit()
    finally: con.close()
    await message.answer("✅ В архив.")

# ---------------------- РЕДАКТИРОВАНИЕ ----------------------
@dp.message(Command("edit_event"))
async def edit_event_start(message: types.Message, state: FSMContext):
    if not is_admin(message.from_user.id): return
    parts = message.text.split()
    if len(parts) < 2 or not parts[1].isdigit(): return await message.answer("❌ `/edit_event ID`", parse_mode="Markdown")
    tid = int(parts[1])
    await state.update_data(edit_tid=tid); await state.set_state(EditEvent.name)
    await message.answer("Название или `-`:")

@dp.message(EditEvent.name)
async def edit_event_name(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("Текст или `-`:")
    if message.text.strip() != "-": await state.update_data(edit_name=message.text.strip())
    await state.set_state(EditEvent.date); await message.answer("Дата (2026-11-15 23:00) или `-`:")

@dp.message(EditEvent.date)
async def edit_event_date(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("Текст или `-`:")
    if message.text.strip() != "-":
        try: datetime.strptime(message.text.strip(), "%Y-%m-%d %H:%M"); await state.update_data(edit_date=message.text.strip())
        except ValueError: return await message.answer("Формат или `-`:")
    await state.set_state(EditEvent.location); await message.answer("Место или `-`:")

@dp.message(EditEvent.location)
async def edit_event_location(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("Текст или `-`:")
    if message.text.strip() != "-": await state.update_data(edit_loc=message.text.strip())
    await state.set_state(EditEvent.event_type)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Fight Night", callback_data="edit_type:fight_night")],
        [InlineKeyboardButton(text="Номерной", callback_data="edit_type:numbered")],
        [InlineKeyboardButton(text="Оставить", callback_data="edit_type:keep")]])
    await message.answer("Тип:", reply_markup=kb)

@dp.callback_query(F.data.startswith("edit_type:"), EditEvent.event_type)
async def edit_event_type(callback: types.CallbackQuery, state: FSMContext):
    await state.update_data(edit_type_choice=callback.data.split(":", 1)[1])
    await state.set_state(EditEvent.photo)
    await callback.message.edit_text("Постер или `-`:"); await callback.answer()

@dp.message(EditEvent.photo)
async def edit_event_photo(message: types.Message, state: FSMContext):
    data = await state.get_data()
    if not message.photo and not (message.text and message.text.strip() == "-"): return await message.answer("Фото или `-`:")
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        if data.get("edit_type_choice") and data["edit_type_choice"] != "keep": con.execute("UPDATE tournaments SET event_type=? WHERE id=?", (data["edit_type_choice"], data["edit_tid"]))
        if "edit_name" in data: con.execute("UPDATE tournaments SET name=? WHERE id=?", (data["edit_name"], data["edit_tid"]))
        if "edit_date" in data: con.execute("UPDATE tournaments SET date_utc=? WHERE id=?", (data["edit_date"], data["edit_tid"]))
        if "edit_loc" in data: con.execute("UPDATE tournaments SET location=? WHERE id=?", (data["edit_loc"], data["edit_tid"]))
        if message.photo: con.execute("UPDATE tournaments SET photo_id=? WHERE id=?", (message.photo[-1].file_id, data["edit_tid"]))
        con.commit()
    finally: con.close()
    await state.clear(); await message.answer("✅ Обновлен!", reply_markup=main_kb)

@dp.message(Command("edit_fight"))
async def edit_fight_start(message: types.Message, state: FSMContext):
    if not is_admin(message.from_user.id): return
    parts = message.text.split()
    if len(parts) < 2 or not parts[1].isdigit(): return await message.answer("❌ `/edit_fight ID`", parse_mode="Markdown")
    fid = int(parts[1])
    await state.update_data(edit_fid=fid); await state.set_state(EditFight.weight)
    await message.answer("Вес или `-`:")

@dp.message(EditFight.weight)
async def edit_fight_weight(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("Текст или `-`:")
    if message.text.strip() != "-": await state.update_data(edit_weight=message.text.strip())
    await state.set_state(EditFight.f1); await message.answer("Боец 1 или `-`:")

@dp.message(EditFight.f1)
async def edit_fight_f1(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("Текст или `-`:")
    if message.text.strip() != "-": await state.update_data(edit_f1=message.text.strip())
    await state.set_state(EditFight.f2); await message.answer("Боец 2 или `-`:")

@dp.message(EditFight.f2)
async def edit_fight_f2(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("Текст или `-`:")
    if message.text.strip() != "-": await state.update_data(edit_f2=message.text.strip())
    await state.set_state(EditFight.main_card); await message.answer("Главный? (да/нет/`-`):")

@dp.message(EditFight.main_card)
async def edit_fight_main(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("да/нет/`-`:")
    ans = message.text.strip().lower()
    if ans in ("да", "yes"): await state.update_data(edit_main=1)
    elif ans in ("нет", "no"): await state.update_data(edit_main=0)
    await state.set_state(EditFight.title); await message.answer("Титульный? (да/нет/`-`):")

@dp.message(EditFight.title)
async def edit_fight_title(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("да/нет/`-`:")
    ans = message.text.strip().lower()
    if ans in ("да", "yes"): await state.update_data(edit_title=1)
    elif ans in ("нет", "no"): await state.update_data(edit_title=0)
    data = await state.get_data()
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        if "edit_weight" in data: con.execute("UPDATE fights SET weight_class=? WHERE id=?", (data["edit_weight"], data["edit_fid"]))
        if "edit_f1" in data: con.execute("UPDATE fights SET fighter_1=? WHERE id=?", (data["edit_f1"], data["edit_fid"]))
        if "edit_f2" in data: con.execute("UPDATE fights SET fighter_2=? WHERE id=?", (data["edit_f2"], data["edit_fid"]))
        if "edit_main" in data: con.execute("UPDATE fights SET is_main_card=? WHERE id=?", (data["edit_main"], data["edit_fid"]))
        if "edit_title" in data: con.execute("UPDATE fights SET is_title=? WHERE id=?", (data["edit_title"], data["edit_fid"]))
        con.commit()
    finally: con.close()
    await state.clear(); await message.answer("✅ Обновлен!", reply_markup=main_kb)

@dp.message(Command("delete_fight"))
async def cmd_delete_fight(message: types.Message):
    if not is_admin(message.from_user.id): return
    parts = message.text.split()
    if len(parts) < 2 or not parts[1].isdigit(): return await message.answer("❌ `/delete_fight ID`", parse_mode="Markdown")
    fid = int(parts[1])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        con.execute("DELETE FROM fights WHERE id=?", (fid,))
        con.execute("DELETE FROM predictions WHERE fight_id=?", (fid,))
        con.commit()
    finally: con.close()
    await message.answer(f"✅ Удален.", reply_markup=main_kb)

@dp.message(Command("timezone"))
@dp.message(F.text == "⚙️ Настройки")
async def cmd_timezone(message: types.Message, state: FSMContext):
    await state.clear()
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=tz, callback_data=f"tz:{tz}")] for tz in TZ_OPTIONS])
    await message.answer("Часовой пояс:", reply_markup=kb)

@dp.callback_query(F.data.startswith("tz:"))
async def cb_timezone(callback: types.CallbackQuery):
    await callback.answer()
    tz_name = callback.data.split(":", 1)[1]
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: con.execute("INSERT INTO users(telegram_id, timezone) VALUES(?,?) ON CONFLICT(telegram_id) DO UPDATE SET timezone=excluded.timezone", (callback.from_user.id, tz_name)); con.commit()
    finally: con.close()
    await callback.message.edit_text(f"Часовой пояс: {tz_name}")

# ---------------------- ТУРНИРЫ (+ НАПОМИНАНИЯ-ПЕРЕКЛЮЧАТЕЛИ) ----------------------
@dp.message(Command("fights"))
@dp.message(F.text == "🥊 Турниры")
async def cmd_fights(message: types.Message, state: FSMContext):
    await state.clear()
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🥊 Fight Night", callback_data="next_type:fight_night:0")],
        [InlineKeyboardButton(text="🏆 Номерной", callback_data="next_type:numbered:0")]])
    await message.answer("Тип:", reply_markup=kb)

@dp.callback_query(F.data == "back_to_menu")
async def cb_back_to_menu(callback: types.CallbackQuery):
    await callback.answer()
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🥊 Fight Night", callback_data="next_type:fight_night:0")],
        [InlineKeyboardButton(text="🏆 Номерной", callback_data="next_type:numbered:0")]])
    try: await callback.message.edit_text("Тип:", reply_markup=kb)
    except TelegramBadRequest:
        await callback.message.delete(); await callback.message.answer("Тип:", reply_markup=kb)

@dp.callback_query(F.data.startswith("next_type:"))
async def cb_next_type(callback: types.CallbackQuery):
    await callback.answer()
    parts = callback.data.split(":"); event_type = parts[1]; page = int(parts[2]) if len(parts) > 2 else 0
    limit = 5; offset = page * limit
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        total = con.execute("SELECT COUNT(*) FROM tournaments WHERE status='announced' AND event_type=?", (event_type,)).fetchone()[0]
        rows = con.execute("SELECT id, name FROM tournaments WHERE status='announced' AND event_type=? ORDER BY date_utc ASC LIMIT ? OFFSET ?", (event_type, limit, offset)).fetchall()
    finally: con.close()
    if total == 0:
        kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⬅️ Назад", callback_data="back_to_menu")]])
        try: await callback.message.edit_text("Нет турниров.", reply_markup=kb)
        except TelegramBadRequest: await callback.message.delete(); await callback.message.answer("Нет турниров.", reply_markup=kb)
        return
    buttons = [[InlineKeyboardButton(text=name, callback_data=f"show_tour:{tid}:{page}")] for tid, name in rows]
    nav = []
    if page > 0: nav.append(InlineKeyboardButton(text="⬅️", callback_data=f"next_type:{event_type}:{page-1}"))
    if (page + 1) * limit < total: nav.append(InlineKeyboardButton(text="➡️", callback_data=f"next_type:{event_type}:{page+1}"))
    if nav: buttons.append(nav)
    buttons.append([InlineKeyboardButton(text="⬅️ В меню", callback_data="back_to_menu")])
    kb = InlineKeyboardMarkup(inline_keyboard=buttons)
    try: await callback.message.edit_text(f"Турниры (Стр. {page+1}):", reply_markup=kb)
    except TelegramBadRequest: await callback.message.delete(); await callback.message.answer(f"Турниры (Стр. {page+1}):", reply_markup=kb)

@dp.callback_query(F.data.startswith("show_tour:"))
async def cb_show_tour(callback: types.CallbackQuery):
    try: await callback.answer()
    except Exception: pass
    
    parts = callback.data.split(":"); tid = int(parts[1]); page = int(parts[2]) if len(parts) > 2 else 0
    tz = get_user_tz(callback.from_user.id)
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        row = con.execute("SELECT name, date_utc, location, event_type, photo_id FROM tournaments WHERE id=?", (tid,)).fetchone()
        if not row:
            try: await callback.message.edit_text("Не найден.")
            except TelegramBadRequest: await callback.message.delete(); await callback.message.answer("Не найден.")
            return
        name, date_utc, location, event_type, photo_id = row
        fights = con.execute("SELECT fighter_1, fighter_2, weight_class, is_title, is_main_card FROM fights WHERE tournament_id=? ORDER BY is_main_card ASC, id ASC", (tid,)).fetchall()
        
        rems = con.execute("SELECT reminder_type FROM tour_reminders WHERE telegram_id=? AND tournament_id=?", (callback.from_user.id, tid)).fetchall()
        has_1h = any(r[0] == '1h' for r in rems)
        has_24h = any(r[0] == '24h' for r in rems)
    finally: con.close()
    
    local = datetime.strptime(date_utc, "%Y-%m-%d %H:%M").replace(tzinfo=ZoneInfo("UTC")).astimezone(tz)
    text = f"🏆 <b>{html.escape(name)}</b>\n📅 {local.strftime('%d %B %Y г.')}\n📍 {html.escape(location or '—')}\n\n"
    if fights:
        main = [f for f in fights if f[4] == 1]; pre = [f for f in fights if f[4] == 2]; early = [f for f in fights if f[4] == 3]
        if main:
            text += "🔥 <b>Главный кард:</b>\n"
            for f1, f2, wc, t, _ in main: text += f"<i>{html.escape(wc or 'Без веса')}</i>: <b>{html.escape(f1)}</b>{' (C)' if t else ''} vs <b>{html.escape(f2)}</b>{' (C)' if t else ''}\n\n"
        if pre:
            text += "🥊 <b>Прелимы:</b>\n"
            for f1, f2, wc, t, _ in pre: text += f"<i>{html.escape(wc or 'Без веса')}</i>: <b>{html.escape(f1)}</b>{' (C)' if t else ''} vs <b>{html.escape(f2)}</b>{' (C)' if t else ''}\n\n"
        if early:
            text += "🥊 <b>Ранние:</b>\n"
            for f1, f2, wc, t, _ in early: text += f"<i>{html.escape(wc or 'Без веса')}</i>: <b>{html.escape(f1)}</b>{' (C)' if t else ''} vs <b>{html.escape(f2)}</b>{' (C)' if t else ''}\n\n"
    else: text += "Кард не объявлен."
    
    if photo_id and len(text) > 1024: text = safe_truncate_html(text, 1000)
    elif not photo_id and len(text) > 4000: text = safe_truncate_html(text, 4000)
    
    kb_btns = []
    if fights:
        kb_btns.append([InlineKeyboardButton(text="🔔 Подписаться", callback_data=f"sub_list:{tid}"),
                        InlineKeyboardButton(text="🎯 Прогноз", callback_data=f"pred_list:{tid}")])
    
    btn_1h = "🔕 Отменить за час" if has_1h else "🔔 Напомнить за час"
    btn_24h = "🔕 Отменить за день" if has_24h else "🔔 Напомнить за день"

    kb_btns.append([InlineKeyboardButton(text=btn_1h, callback_data=f"remind_1h:{tid}:{page}"),
                    InlineKeyboardButton(text=btn_24h, callback_data=f"remind_24h:{tid}:{page}")])

    kb_btns.append([InlineKeyboardButton(text="⬅️ Назад", callback_data=f"next_type:{event_type}:{page}")])
    kb = InlineKeyboardMarkup(inline_keyboard=kb_btns)
    
    try: await callback.message.delete()
    except Exception: pass
    try:
        if photo_id: await callback.message.answer_photo(photo=photo_id, caption=text, parse_mode="HTML", reply_markup=kb)
        else: await callback.message.answer(text, parse_mode="HTML", reply_markup=kb)
    except TelegramBadRequest:
        if photo_id: await callback.message.answer_photo(photo=photo_id, caption=text[:1024], reply_markup=kb)
        else: await callback.message.answer(text[:4096], reply_markup=kb)

@dp.callback_query(F.data.startswith("remind_1h:"))
async def cb_remind_1h(callback: types.CallbackQuery):
    parts = callback.data.split(":")
    tid = int(parts[1])
    page = int(parts[2]) if len(parts) > 2 else 0
    
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        existing = con.execute("SELECT 1 FROM tour_reminders WHERE telegram_id=? AND tournament_id=? AND reminder_type='1h'", (callback.from_user.id, tid)).fetchone()
        if existing:
            con.execute("DELETE FROM tour_reminders WHERE telegram_id=? AND tournament_id=? AND reminder_type='1h'", (callback.from_user.id, tid))
            con.commit()
            await callback.answer("🔕 Напоминание за час отменено!", show_alert=True)
        else:
            con.execute("INSERT INTO tour_reminders(telegram_id, tournament_id, reminder_type) VALUES(?,?,'1h')", (callback.from_user.id, tid))
            con.commit()
            await callback.answer("✅ Напомню за час до турнира!", show_alert=True)
    finally: con.close()
    
    callback.data = f"show_tour:{tid}:{page}"
    await cb_show_tour(callback)

@dp.callback_query(F.data.startswith("remind_24h:"))
async def cb_remind_24h(callback: types.CallbackQuery):
    parts = callback.data.split(":")
    tid = int(parts[1])
    page = int(parts[2]) if len(parts) > 2 else 0
    
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        existing = con.execute("SELECT 1 FROM tour_reminders WHERE telegram_id=? AND tournament_id=? AND reminder_type='24h'", (callback.from_user.id, tid)).fetchone()
        if existing:
            con.execute("DELETE FROM tour_reminders WHERE telegram_id=? AND tournament_id=? AND reminder_type='24h'", (callback.from_user.id, tid))
            con.commit()
            await callback.answer("🔕 Напоминание за день отменено!", show_alert=True)
        else:
            con.execute("INSERT INTO tour_reminders(telegram_id, tournament_id, reminder_type) VALUES(?,?,'24h')", (callback.from_user.id, tid))
            con.commit()
            await callback.answer("✅ Напомню за день до турнира!", show_alert=True)
    finally: con.close()
    
    callback.data = f"show_tour:{tid}:{page}"
    await cb_show_tour(callback)

# ---------------------- ИТОГИ ----------------------
@dp.message(Command("results"))
@dp.message(F.text == "📊 Итоги")
async def cmd_results(message: types.Message, state: FSMContext):
    await state.clear()
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        tours = con.execute("SELECT id, name, date_utc FROM tournaments WHERE status='finished' ORDER BY date_utc DESC LIMIT 3").fetchall()
        if not tours: return await message.answer("Итогов нет.")
        tz = get_user_tz(message.from_user.id)
        parts = []
        for tid, name, date_utc in tours:
            block = f"✅ <b>{html.escape(name)}</b>\n📅 {fmt_utc_to_local(date_utc, tz)}\n"
            fights = con.execute("SELECT fighter_1, fighter_2, winner, method, round, time FROM fights WHERE tournament_id=? AND status='finished' ORDER BY is_main_card ASC, id", (tid,)).fetchall()
            for f1, f2, w, m, r, t in fights:
                if w: block += f"• {html.escape(f1)} vs {html.escape(f2)} → {html.escape(w)} ({html.escape(m or '?')}, R{r} {t})\n"
                else: block += f"• {html.escape(f1)} vs {html.escape(f2)} → нет\n"
            parts.append(block)
    finally: con.close()
    result = "\n".join(parts)
    if len(result) > 4000: result = safe_truncate_html(result, 4000)
    await message.answer(result, parse_mode="HTML")

# ---------------------- ДОБАВЛЕНИЕ ТУРНИРА ----------------------
@dp.message(Command("add_event"))
async def add_event_start(message: types.Message, state: FSMContext):
    if not is_admin(message.from_user.id): return
    await state.set_state(AddEvent.name); await message.answer("Название:")

@dp.message(AddEvent.name)
async def add_event_name(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("Текст:")
    await state.update_data(name=message.text.strip()); await state.set_state(AddEvent.date); await message.answer("Дата UTC (2026-11-15 23:00):")

@dp.message(AddEvent.date)
async def add_event_date(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("Текст:")
    try: datetime.strptime(message.text.strip(), "%Y-%m-%d %H:%M")
    except ValueError: return await message.answer("Формат:")
    await state.update_data(date=message.text.strip()); await state.set_state(AddEvent.location); await message.answer("Место:")

@dp.message(AddEvent.location)
async def add_event_location(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("Текст:")
    await state.update_data(location=message.text.strip()); await state.set_state(AddEvent.event_type)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Fight Night", callback_data="set_type:fight_night")],
        [InlineKeyboardButton(text="Номерной", callback_data="set_type:numbered")]])
    await message.answer("Тип:", reply_markup=kb)

@dp.callback_query(F.data.startswith("set_type:"), AddEvent.event_type)
async def add_event_type(callback: types.CallbackQuery, state: FSMContext):
    await state.update_data(event_type=callback.data.split(":", 1)[1])
    await state.set_state(AddEvent.photo)
    await callback.message.edit_text("Постер или `-`:"); await callback.answer()

@dp.message(AddEvent.photo)
async def add_event_photo(message: types.Message, state: FSMContext):
    if not message.photo and not (message.text and message.text.strip() == "-"): return await message.answer("Фото или `-`:")
    data = await state.get_data()
    photo_id = message.photo[-1].file_id if message.photo else None
    
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: 
        con.execute("INSERT INTO tournaments(name, date_utc, location, event_type, photo_id) VALUES(?,?,?,?,?)", (data["name"], data["date"], data["location"], data["event_type"], photo_id))
        con.commit()
    finally: con.close()
    
    log_action(message.from_user.id, f"Добавил турнир: {data['name']}")
    await state.clear(); 
    await message.answer(f"✅ Турнир добавлен!", reply_markup=main_kb)
    
    text = f"🆕 <b>Новый турнир!</b>\n\n🏆 <b>{html.escape(data['name'])}</b>\n📅 {data['date']} UTC\n📍 {html.escape(data['location'])}\n\nОткрой /fights → посмотри кард и сделай прогноз."
    await message.answer("⏳ Начинаю рассылку...")
    s, f = await broadcast_to_all(text, photo_id=photo_id, exclude_id=message.from_user.id)
    await message.answer(f"✅ Доставлено: {s}\nБлок: {f}")

# ---------------------- КАРД ----------------------
@dp.message(Command("card"))
async def cmd_card_fast(message: types.Message):
    if not is_admin(message.from_user.id): return
    if not message.text: return await message.answer("Текст.")
    text = message.text.replace("/card", "").strip()
    
    if not text:
        con = sqlite3.connect(DB_PATH, timeout=10)
        try: rows = con.execute("SELECT id, name FROM tournaments WHERE status='announced' ORDER BY id ASC").fetchall()
        finally: con.close()
        if not rows: return await message.answer("База пуста.")
        listing = "\n".join(f"ID <b>{tid}</b>: {html.escape(name)}" for tid, name in rows)
        return await message.answer(f"📋 <b>Турниры:</b>\n\n{listing}\n\n<code>/card ID\n[Текст]</code>", parse_mode="HTML")
    
    lines = text.split("\n")
    tid_str = lines[0].strip()
    if not tid_str.isdigit(): return await message.answer("Первая строка должна быть ID турнира (цифра).")
    tid = int(tid_str)
    
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        tour = con.execute("SELECT name, photo_id FROM tournaments WHERE id=?", (tid,)).fetchone()
        if not tour: return await message.answer(f"❌ Турнир с ID {tid} не найден.")
        
        cur_section = 1
        added_count = 0
        added_fighters = []
        
        for line in lines[1:]:
            line = line.strip()
            if not line: continue
            lower_line = line.lower()
            
            if "главный кард" in lower_line: 
                cur_section = 1
                continue
            elif "прелимы" in lower_line or "предварительный" in lower_line: 
                cur_section = 2
                continue
            elif "ранние" in lower_line: 
                cur_section = 3
                continue
            
            if "ufc" in lower_line and "vs" in lower_line: continue
            
            if ":" in line and " vs" in lower_line:
                try:
                    weight_raw = line.split(":")[0].strip()
                    fighters_raw = line.split(":", 1)[1].strip()
                    
                    is_title = 0
                    if "(титул" in weight_raw.lower() or "титульный" in weight_raw.lower():
                        is_title = 1
                    
                    weight = re.sub(r'\([^)]*\)', '', weight_raw).strip()
                    
                    sf = re.split(r'\s+vs\.?\s+', fighters_raw, maxsplit=1, flags=re.IGNORECASE)
                    if len(sf) < 2: continue
                    
                    f1 = sf[0].strip()
                    f2 = sf[1].strip()
                    
                    if not f1 or not f2 or f1.startswith("UFC"): continue
                    
                    added_fighters.extend([f1, f2])
                    con.execute("""
                        INSERT INTO fights(tournament_id, weight_class, fighter_1, fighter_2, is_main_card, is_title) 
                        VALUES(?,?,?,?,?,?)
                    """, (tid, weight, f1, f2, cur_section, is_title))
                    added_count += 1
                except Exception as e:
                    print(f"Ошибка парсинга строки '{line}': {e}")
                    continue
        con.commit()
    finally:
        con.close()
        
    log_action(message.from_user.id, f"Опубликовал кард турнира {tid}")
    await message.answer(f"✅ <b>Кард добавлен!</b>\n👊 Боёв: {added_count}", parse_mode="HTML")
    
    try:
        tour_name, tour_photo = tour[0], tour[1]
        s, f = await broadcast_to_all(
            f"🆕 <b>Новый кард!</b>\n\n🏆 {html.escape(tour_name)}\nОткрой /fights → посмотри бои и сделай прогноз.", 
            photo_id=tour_photo, only_premium=True, exclude_id=message.from_user.id
        )
        await message.answer(f"⭐ Премиумам отправлено: {s}")
    except Exception as e: print(f"Ошибка рассылки: {e}")
    
    if added_fighters: await notify_subscribers(added_fighters, tour_name)

# ---------------------- РЕЗУЛЬТАТ (+ ПРОВЕРКА МЕТОДА) ----------------------
@dp.message(Command("set_result"))
async def set_result_start(message: types.Message, state: FSMContext):
    if not is_admin(message.from_user.id): return
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: rows = con.execute("SELECT id, name FROM tournaments WHERE status='announced' ORDER BY date_utc DESC LIMIT 20").fetchall()
    finally: con.close()
    if not rows: return await message.answer("Нет турниров.")
    listing = "\n".join(f"{tid}. {html.escape(name)}" for tid, name in rows)
    await state.set_state(SetResult.tournament)
    await message.answer(f"ID турнира:\n{listing}", parse_mode="HTML")

@dp.message(SetResult.tournament)
async def set_result_tournament(message: types.Message, state: FSMContext):
    if not message.text or not message.text.strip().isdigit(): return await message.answer("ID.")
    tid = int(message.text.strip())
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: fights = con.execute("SELECT id, fighter_1, fighter_2 FROM fights WHERE tournament_id=? AND status='scheduled'", (tid,)).fetchall()
    finally: con.close()
    if not fights: return await message.answer("Нет боёв.")
    await state.update_data(tid=tid)
    listing = "\n".join(f"{fid}. {html.escape(f1)} vs {html.escape(f2)}" for fid, f1, f2 in fights)
    await state.set_state(SetResult.fight); await message.answer(f"ID боя:\n{listing}", parse_mode="HTML")

@dp.message(SetResult.fight)
async def set_result_fight(message: types.Message, state: FSMContext):
    if not message.text or not message.text.strip().isdigit(): return await message.answer("ID.")
    await state.update_data(fid=int(message.text.strip())); await state.set_state(SetResult.winner)
    await message.answer("Победитель (или draw/nc):")

@dp.message(SetResult.winner)
async def set_result_winner(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("Текст:")
    await state.update_data(winner=message.text.strip()); await state.set_state(SetResult.method); await message.answer("Метод (KO/Sub/Dec или текст):")

@dp.message(SetResult.method)
async def set_result_method(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("Текст:")
    await state.update_data(method=message.text.strip()); await state.set_state(SetResult.round); await message.answer("Раунд (1-5 или 0):")

@dp.message(SetResult.round)
async def set_result_round(message: types.Message, state: FSMContext):
    if not message.text or not message.text.strip().isdigit(): return await message.answer("Число:")
    await state.update_data(rnd=int(message.text.strip())); await state.set_state(SetResult.time); await message.answer("Время (5:00):")

async def delayed_broadcast(text: str):
    await asyncio.sleep(600)
    try:
        await broadcast_to_all(text, exclude_premium=True)
        print("Отложенная рассылка результатов завершена.")
    except Exception as e: print(f"Ошибка отложенной рассылки: {e}")

@dp.message(SetResult.time)
async def set_result_time(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("Текст (5:00):")
    data = await state.get_data()
    
    admin_method_std = normalize_method(data["method"])

    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        fi = con.execute("SELECT fighter_1, fighter_2, status FROM fights WHERE id=?", (data["fid"],)).fetchone()
        if not fi: await message.answer("❌ Бой не найден."); await state.clear(); return
        if fi[2] == "finished": await message.answer("⚠️ Уже внесен!"); await state.clear(); return
        con.execute("UPDATE fights SET status='finished', winner=?, method=?, round=?, time=? WHERE id=?", (data["winner"], data["method"], data["rnd"], message.text.strip(), data["fid"]))
        
        preds = con.execute("SELECT id, telegram_id, predicted_winner, predicted_method FROM predictions WHERE fight_id=? AND status='pending'", (data["fid"],)).fetchall()
        won, lost, perfect = 0, 0, 0
        
        for pid, uid, pred, pred_method in preds:
            if pred.lower().strip() == data["winner"].lower().strip():
                con.execute("UPDATE predictions SET status='won' WHERE id=?", (pid,))
                
                if pred_method == admin_method_std:
                    perfect += 1
                    try: await bot.send_message(uid, f"🎯 <b>ИДЕАЛЬНЫЙ ПРОГНОЗ!</b>\n\nТы угадал и победителя ({pred}), и метод ({data['method']})!", parse_mode="HTML")
                    except Exception: pass
                else:
                    try: await bot.send_message(uid, f"🎉 <b>Угадал!</b>\n{pred} победил.", parse_mode="HTML")
                    except Exception: pass
                won += 1
            else:
                con.execute("UPDATE predictions SET status='lost' WHERE id=?", (pid,))
                try: await bot.send_message(uid, f"❌ <b>Не угадал.</b>\nПобедил {data['winner']}.", parse_mode="HTML")
                except Exception: pass
                lost += 1
        con.commit()
    finally: con.close()
    log_action(message.from_user.id, f"Внес результат боя #{data['fid']}")
    await state.clear()
    await message.answer(f"✅ Результат сохранён.\n🎯 Угадали победителя: {won} (из них идеально: {perfect}), не угадали: {lost}")
    
    f1, f2 = fi[0], fi[1]
    if data["winner"].lower() in ("draw", "nc"):
        text = f"🥊 РЕЗУЛЬТАТ\n\n{f1} vs {f2}\nИтог: {data['winner'].upper()} ({data['method']})"
    else:
        text = f"🥊 РЕЗУЛЬТАТ\n\n{f1} vs {f2}\n\n🏆 {data['winner']} побеждает!\nМетод: {data['method']}\nРаунд: {data['rnd']}, Время: {message.text.strip()}"
    
    await broadcast_to_all(text, only_premium=True)
    await message.answer("⚡ Премиум-юзерам результат отправлен мгновенно!")
    asyncio.create_task(delayed_broadcast(text))
    await message.answer("⏳ Остальным пользователям результат уйдет через 10 минут (Ранний доступ).", reply_markup=main_kb)

@dp.message(Command("reset_result"))
async def cmd_reset_result(message: types.Message):
    if not is_admin(message.from_user.id): return
    parts = message.text.split()
    if len(parts) < 2 or not parts[1].isdigit(): return await message.answer("❌ `/reset_result ID_боя`", parse_mode="Markdown")
    fid = int(parts[1])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        fight = con.execute("SELECT fighter_1, fighter_2, status FROM fights WHERE id=?", (fid,)).fetchone()
        if not fight: return await message.answer(f"❌ Не найден.")
        if fight[2] != "finished": return await message.answer(f"⚠️ Бой не завершен.")
        con.execute("UPDATE fights SET status='scheduled', winner=NULL, method=NULL, round=NULL, time=NULL WHERE id=?", (fid,))
        con.execute("UPDATE predictions SET status='pending' WHERE fight_id=?", (fid,))
        con.commit()
        log_action(message.from_user.id, f"Сбросил результат боя #{fid}")
        await message.answer(f"✅ <b>Результат сброшен!</b>", parse_mode="HTML")
    finally: con.close()

@dp.message(Command("delete_event"))
async def cmd_delete_event(message: types.Message):
    if not is_admin(message.from_user.id): return
    parts = message.text.split()
    if len(parts) < 2:
        con = sqlite3.connect(DB_PATH, timeout=10)
        try: rows = con.execute("SELECT id, name FROM tournaments ORDER BY id DESC").fetchall()
        finally: con.close()
        if not rows: return await message.answer("Пусто.")
        listing = "\n".join(f"{tid}. {html.escape(name)}" for tid, name in rows)
        return await message.answer(f"`/delete_event ID`\n\n{listing}", parse_mode="HTML")
    if not parts[1].isdigit(): return await message.answer("ID.")
    tid = int(parts[1])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        fids = con.execute("SELECT id FROM fights WHERE tournament_id=?", (tid,)).fetchall()
        for (fid,) in fids: con.execute("DELETE FROM predictions WHERE fight_id=?", (fid,))
        con.execute("DELETE FROM fights WHERE tournament_id=?", (tid,)); con.execute("DELETE FROM tournaments WHERE id=?", (tid,)); con.commit()
    finally: con.close()
    log_action(message.from_user.id, f"Удалил турнир #{tid}")
    await message.answer(f"Удален.")

@dp.message(Command("finish_event"))
async def cmd_finish_event(message: types.Message):
    if not is_admin(message.from_user.id): return
    parts = message.text.split()
    if len(parts) < 2 or not parts[1].isdigit(): return await message.answer("`/finish_event ID`", parse_mode="Markdown")
    tid = int(parts[1])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try: con.execute("UPDATE tournaments SET status='finished' WHERE id=?", (tid,)); con.commit()
    finally: con.close()
    log_action(message.from_user.id, f"Завершил турнир #{tid}")
    await message.answer("Турнир завершён.")

@dp.message(Command("edit_id"))
async def cmd_edit_id(message: types.Message):
    if not is_admin(message.from_user.id): return
    parts = message.text.split()
    if len(parts) < 3 or not parts[1].isdigit() or not parts[2].isdigit(): return await message.answer("`/edit_id старый новый`", parse_mode="Markdown")
    old_id, new_id = int(parts[1]), int(parts[2])
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        if con.execute("SELECT name FROM tournaments WHERE id=?", (new_id,)).fetchone(): return await message.answer("Занято.")
        con.execute("UPDATE tournaments SET id=? WHERE id=?", (new_id, old_id))
        con.execute("UPDATE fights SET tournament_id=? WHERE tournament_id=?", (new_id, old_id))
        con.commit()
        await message.answer(f"✅ {old_id} → {new_id}")
    finally: con.close()

# ---------------------- ЖИВОЙ ЧАТ ----------------------
@dp.message((F.text | F.photo) & ~F.from_user.id.in_(ADMIN_IDS))
async def forward_user_message(message: types.Message, state: FSMContext):
    if message.text and message.text.startswith('/'): return
    if message.text in ["🥊 Турниры", "📊 Итоги", "📰 Новости", "📅 Календарь", "👤 Мой профиль", "💳 Купить Премиум", "🔍 Поиск бойца", "⚙️ Настройки", "❓ Помощь", "📈 Топ прогнозов", "🎁 Промокод", "🐛 Баг-репорт"]: return
    if await state.get_state() is not None: return
    info = f"📩 Сообщение\n👤 {message.from_user.full_name}\n🆔 ID: {message.from_user.id}"
    for aid in ADMIN_IDS:
        try:
            if message.photo: await bot.send_photo(aid, message.photo[-1].file_id, caption=f"{info}\n💬 {message.caption or '—'}"[:1024])
            else: await bot.send_message(aid, f"{info}\n💬 {message.text or '(пусто)'}")
        except Exception: pass
    await message.answer("✅ Отправлено.")

@dp.message()
async def fallback(message: types.Message):
    await message.answer("Смотри /help", reply_markup=main_kb)

# ---------------------- ЗАПУСК ----------------------
async def main():
    db_init()
    con = sqlite3.connect(DB_PATH, timeout=10)
    try:
        for (uid,) in con.execute("SELECT telegram_id FROM admins").fetchall(): ADMIN_IDS.add(uid)
    finally: con.close()
    await set_commands(bot)
    
    asyncio.create_task(reminder_worker())
    
    print(f"Бот запущен. Админов: {len(ADMIN_IDS)}, VIP: {len(PREMIUM_IDS)}")
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

#if __name__ == "__main__":
    #asyncio.run(main())