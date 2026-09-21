@echo off
setlocal
cd /d "%~dp0"

rem Production startup script.
rem - Binds to 0.0.0.0 so it is reachable from outside this machine.
rem - DOCVIEW_SECRET is pinned (via start.local.bat, see below) so restarts do not
rem   invalidate already-issued tokens.
rem - DOCVIEW_API_KEYS / DOCVIEW_ALLOWED_ORIGINS must be set for production; see README.
rem NOTE: keep this file ASCII-only and CRLF line endings. Non-ASCII comments plus
rem LF-only endings have been observed to corrupt cmd.exe parsing on some consoles
rem (wrong active code page), silently swallowing the "set" lines below.

set DOCVIEW_HOST=0.0.0.0
set DOCVIEW_PORT=8000

rem Secrets (DOCVIEW_SECRET, optionally DOCVIEW_API_KEYS) live in start.local.bat, which is
rem gitignored so real secrets never land in git. Copy start.local.bat.example to that name
rem and fill in real values. Without it, server.ts falls back to a random per-boot secret and
rem warns loudly at startup (every restart invalidates all issued tokens) rather than failing
rem outright, so a fresh checkout still runs for local development.
if exist "%~dp0start.local.bat" call "%~dp0start.local.bat"

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
