@echo off
setlocal
cd /d "%~dp0"

rem Production startup script.
rem - Binds to 0.0.0.0 so it is reachable from outside this machine.
rem - DOCVIEW_SECRET is pinned so restarts do not invalidate already-issued tokens.
rem - DOCVIEW_API_KEYS / DOCVIEW_ALLOWED_ORIGINS must be set for production; see README.
rem NOTE: keep this file ASCII-only and CRLF line endings. Non-ASCII comments plus
rem LF-only endings have been observed to corrupt cmd.exe parsing on some consoles
rem (wrong active code page), silently swallowing the "set" lines below.

set DOCVIEW_HOST=0.0.0.0
set DOCVIEW_PORT=8000
set DOCVIEW_SECRET=ROTATED_OUT_SECRET_REDACTED_FROM_HISTORY

rem Auth for the upload / token-issuance APIs. List of "name:key" pairs, comma-separated.
rem Replace with the real production key, and set the exact same key on every project
rem that calls into this webviewer (e.g. maintenance-management's WEBVIEWER_API_KEY).
rem set DOCVIEW_API_KEYS=maintenance-management:CHANGE_ME_TO_A_REAL_SECRET_KEY

rem Origins allowed to embed this viewer in an iframe. Comma-separated if more than one.
set DOCVIEW_ALLOWED_ORIGINS=http://localhost:5173,http://vdi.openminds.co.kr:30443

if not exist node_modules (
    echo [webviewer] node_modules not found, running npm install first...
    call npm install
    if errorlevel 1 (
        echo [webviewer] npm install failed
        pause
        exit /b 1
    )
)

echo [webviewer] Starting production server on %DOCVIEW_HOST%:%DOCVIEW_PORT% ...
echo [webviewer] Closing this window stops the server.
call npm start

echo.
echo [webviewer] Process exited. Press any key to close this window.
pause
