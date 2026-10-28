from daemon import daemon
from main import main

with daemon.DaemonContext(): 
    main()
