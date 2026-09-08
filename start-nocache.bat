@echo off
setlocal
cd /d "%~dp0"

set PORT=8787

echo ============================================
echo   Starting no-cache local server for the game (port %PORT%)
echo   This version forces the browser to NEVER use cached
echo   xlsx / image files, so every refresh always gets the
echo   latest edits. Close this window to stop the server.
echo ============================================
echo.

where python >nul 2>nul
if %ERRORLEVEL%==0 (
    start "" http://localhost:%PORT%/index.html
    python "%~dp0nocache_server.py"
    goto :eof
)

where py >nul 2>nul
if %ERRORLEVEL%==0 (
    start "" http://localhost:%PORT%/index.html
    py "%~dp0nocache_server.py"
    goto :eof
)

echo Python not found. Cannot start the no-cache server.
echo Please install Python first, then run this file again:
echo   https://www.python.org/downloads/
echo   (Check "Add python.exe to PATH" during install)
echo.
echo If you have Node.js instead, use start.bat, which will
echo automatically fall back to "npx http-server -c-1"
echo (same no-cache effect as this script).
echo.
pause
