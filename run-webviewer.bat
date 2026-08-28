@echo off
setlocal
cd /d "%~dp0"

rem Uncomment/edit any of these if you need non-default settings.
rem set DOCVIEW_PORT=8000
rem set DOCVIEW_DATA_DIR=C:\path\to\persistent\data
rem set DOCVIEW_SECRET=a-fixed-production-secret

if not exist node_modules (
    echo [webviewer] node_modules not found, running npm install first...
    call npm install
    if errorlevel 1 (
        echo [webviewer] npm install failed
        pause
        exit /b 1
    )
)

echo [webviewer] Starting dev server (default address: http://127.0.0.1:8000) ...
echo [webviewer] Closing this window stops the server.
call npm run dev

echo.
echo [webviewer] Process exited. Press any key to close this window.
pause
