@echo off
setlocal
cd /d "%~dp0"

rem 운영 배포용 실행 스크립트 — 외부(브라우저)에서 접속 가능하도록 0.0.0.0으로 바인딩하고,
rem 재시작해도 기존에 발급된 토큰이 무효화되지 않도록 DOCVIEW_SECRET을 고정값으로 둔다.
set DOCVIEW_HOST=0.0.0.0
set DOCVIEW_PORT=8000
set DOCVIEW_SECRET=ROTATED_OUT_SECRET_REDACTED_FROM_HISTORY

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
