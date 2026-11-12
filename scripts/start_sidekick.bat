@echo off
setlocal
set PYW="C:\Users\Rohit Patil\AppData\Local\Programs\Python\Python312\pythonw.exe"
set DAEMON_DIR=C:\Users\Rohit Patil\sidekick-os\daemon

cd /d "%DAEMON_DIR%"
start "" %PYW% "%DAEMON_DIR%\server.py"
timeout /t 1 /nobreak >nul
start "" %PYW% "%DAEMON_DIR%\mascot.py"

