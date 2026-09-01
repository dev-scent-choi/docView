@echo off
setlocal
cd /d "%~dp0"

rem 운영 배포용 실행 스크립트 — 외부(브라우저)에서 접속 가능하도록 0.0.0.0으로 바인딩하고,
rem 재시작해도 기존에 발급된 토큰이 무효화되지 않도록 DOCVIEW_SECRET을 고정값으로 둔다.
set DOCVIEW_HOST=0.0.0.0
set DOCVIEW_PORT=8000
set DOCVIEW_SECRET=ROTATED_OUT_SECRET_REDACTED_FROM_HISTORY

rem 업로드/토큰 발급 API 인증 — "이름:키" 쌍을 콤마로 나열. 실제 운영 값으로 바꾸고, 각
rem 연동 프로젝트(예: maintenance-management)의 설정에도 똑같은 키를 넣어야 한다.
set DOCVIEW_API_KEYS=maintenance-management:CHANGE_ME_TO_A_REAL_SECRET_KEY

rem 이 화면을 iframe으로 띄울 수 있는 출처. 실제로 임베드할 사이트 주소로 바꾼다(콤마로 여러 개 가능).
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
