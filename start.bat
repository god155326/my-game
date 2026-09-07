@echo off
setlocal
cd /d "%~dp0"

set PORT=8787

echo ============================================
echo   啟動「不朽之旅」本機伺服器 (port %PORT%)
echo   關閉這個視窗即可停止伺服器
echo ============================================
echo.

where python >nul 2>nul
if %ERRORLEVEL%==0 (
    start "" http://localhost:%PORT%/index.html
    python -m http.server %PORT%
    goto :eof
)

where py >nul 2>nul
if %ERRORLEVEL%==0 (
    start "" http://localhost:%PORT%/index.html
    py -m http.server %PORT%
    goto :eof
)

where npx >nul 2>nul
if %ERRORLEVEL%==0 (
    start "" http://localhost:%PORT%/index.html
    npx --yes http-server -p %PORT% -c-1
    goto :eof
)

echo 找不到 Python 或 Node.js，無法啟動本機伺服器。
echo 請先安裝其中一個，再重新執行這個檔案：
echo   Python: https://www.python.org/downloads/  (安裝時記得勾選 "Add python.exe to PATH")
echo   Node.js: https://nodejs.org/
echo.
pause
