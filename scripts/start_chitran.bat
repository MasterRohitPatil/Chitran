@echo off
setlocal
set DAEMON_DIR=%~dp0..\daemon

if exist "%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe" (
    set PYW="%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe"
) else (
    set PYW=pythonw
)

cd /d "%DAEMON_DIR%"
start "" %PYW% "%DAEMON_DIR%\server.py"
timeout /t 1 /nobreak >nul
start "" %PYW% "%DAEMON_DIR%\mascot.py"

