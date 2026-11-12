@echo off
set CHROME_BIN=C:\Program Files\Google\Chrome\Application\chrome.exe

if exist "%LOCALAPPDATA%\Google\Chrome\ChitranProfile" (
    set PROFILE_DIR=%LOCALAPPDATA%\Google\Chrome\ChitranProfile
) else (
    set PROFILE_DIR=%LOCALAPPDATA%\Google\Chrome\SidekickProfile
)

if not exist "%CHROME_BIN%" (
    echo [Chitran] Error: Chrome not found at "%CHROME_BIN%"
    pause
    exit /b 1
)

echo [Chitran] Launching Chrome with CDP on port 9222...
echo Profile path: "%PROFILE_DIR%"

start "" "%CHROME_BIN%" ^
  --remote-debugging-port=9222 ^
  --user-data-dir="%PROFILE_DIR%" ^
  --disable-background-timer-throttling ^
  --disable-backgrounding-occluded-windows ^
  --disable-renderer-backgrounding ^
  --no-first-run ^
  --no-default-browser-check
