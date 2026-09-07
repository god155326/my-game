@echo off
REM stop-server.bat
REM Stops whatever is listening on port 8787 (the game's local server),
REM even though it's running hidden with no visible window.

setlocal enabledelayedexpansion
set FOUND=0

for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8787" ^| findstr "LISTENING"') do (
    echo Stopping server process %%a ...
    taskkill /F /PID %%a >nul 2>&1
    set FOUND=1
)

if "!FOUND!"=="1" (
    echo Server stopped.
) else (
    echo No server was running on port 8787.
)
pause
