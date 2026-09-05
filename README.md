# DocView (webviewer)

권한 제어형 문서 웹뷰어. 업로드된 문서(pdf/이미지/xlsx/csv/docx/pptx 등)를 서버에서 정규화해
열람시키고, docx/xlsx는 Yjs 기반 실시간 동시편집(Office 365 스타일)을 지원한다. 원본 파일은
열람/편집 토큰이 있어야만 접근 가능하며, 브라우저로 직접 전달되지 않는다(다운로드 권한이 별도로
있는 경우만 예외).

다른 프로젝트(예: `maintenance-management`)에서 이 서비스의 HTTP API만 호출해 문서를 업로드하고
편집 화면을 iframe으로 임베드하는 방식으로 연동할 수 있다 — 이 프로젝트는 항상 독립적으로
빌드/구동되며, 다른 프로젝트와 소스를 합치지 않는다.

## 요구 사항

- Node.js 20 이상
- (선택) LibreOffice — docx/pptx 등을 PDF로 변환하려면 필요. docx는 없어도 자동으로 서식 없는
  텍스트 미리보기 폴백으로 동작하지만, pptx는 이 폴백이 없어 LibreOffice 없이는 업로드가 거부된다
  (xlsx/csv는 LibreOffice 없이도 항상 정상 동작 — 셀 데이터만 직접 추출).

## 설치

```bash
npm install
```

## 개발 실행

```bash
npm run dev
```

`editor.bundle.js`/`grid-editor.bundle.js`를 esbuild로 먼저 빌드한 뒤, `tsx watch`로 서버를 띄운다
(소스 변경 시 자동 재시작). 기본 주소: `http://127.0.0.1:8000`.

## 운영 실행

```bash
npm run dev    # 위와 동일하지만 파일 변경 감시 없이 1회 실행하려면:
npm run build:editor && npm start
```

`start`는 번들 빌드 후 `tsx`로 서버를 1회 실행한다(별도의 트랜스파일 산출물을 디스크에 남기지 않음
— TypeScript를 tsx가 즉시 실행). 프로세스 관리자(pm2, systemd, Windows 서비스 등)로 감싸서 운영하면
된다.

## 환경 변수

| 변수 | 기본값 | 설명 |
|---|---|---|
| `DOCVIEW_PORT` | `8000` | 서버 포트 |
| `DOCVIEW_HOST` | `127.0.0.1` | 바인딩 호스트 |
| `DOCVIEW_DATA_DIR` | `<cwd>/data` | 원본/변환본/메타/감사로그 저장 위치 |
| `DOCVIEW_SECRET` | (미설정 시 매 재시작마다 임시 생성) | 열람/편집 토큰 서명 키 — **운영 환경에서는 반드시 고정값으로 설정**. 미설정 시 서버 재시작마다 기존에 발급된 토큰이 전부 무효화된다. |
| `DOCVIEW_TOKEN_TTL` | `3600` (초) | 발급 토큰 유효기간 |
| `DOCVIEW_STORAGE` | `disk` | `disk` \| `memory` — `memory`는 재시작 시 전부 소멸(디스크 비잔류) |
| `DOCVIEW_MEMORY_BUDGET_MB` | `1024` | `memory` 모드일 때 원본 보관 메모리 상한 |
| `DOCVIEW_CACHE_BUDGET_MB` | `512` | 렌더링/파싱 캐시(그리드·PDF 텍스트·페이지 이미지) 메모리 상한 |
| `DOCVIEW_SOFFICE` | (자동 탐지) | LibreOffice(`soffice`) 실행 파일 경로를 직접 지정할 때 |
| `DOCVIEW_API_KEYS` | (미설정 시 인증 없음) | 업로드·토큰 발급 API를 호출할 수 있는 프로젝트별 키. `이름:키` 쌍을 콤마로 나열 (예: `maintenance-management:abc123,other-project:def456`). **운영 환경에서는 반드시 설정하세요** — 미설정 시 이 두 API는 누구나 호출할 수 있습니다. |
| `DOCVIEW_ALLOWED_ORIGINS` | (미설정 시 제한 없음) | 이 화면(뷰어/편집기)을 iframe으로 띄울 수 있는 출처 목록, 콤마로 나열 (예: `https://maintenance.example.com`). **운영 환경에서는 반드시 설정하세요** — 미설정 시 아무 사이트에서나 이 화면을 임베드할 수 있습니다. |

예시(Windows, Git Bash):

```bash
export DOCVIEW_PORT=8000
export DOCVIEW_DATA_DIR="/c/path/to/persistent/data"
export DOCVIEW_SECRET="운영용-고정-시크릿-여기에"
npm start
```

## 검증

```bash
npm run typecheck        # tsc --noEmit
npm test                 # vitest run
node scripts/browser-check.mjs            # Playwright 헤드리스 브라우저 회귀(업로드→뷰어→편집기)
node scripts/browser-check-presence.mjs   # 2인 동시편집 프레즌스 회귀
node scripts/formula-check.mjs            # xlsx 수식 엔진(VLOOKUP/MATCH/SUMIFS 등) 회귀
```

뒤의 세 스크립트는 실행 전에 서버가 떠 있어야 하며, 서버가 사용한 것과 같은
`DOCVIEW_PORT`/`DOCVIEW_DATA_DIR`/`DOCVIEW_SECRET` 값을 환경 변수로 맞춰줘야 한다(JWT 서명 키가
다르면 토큰 검증이 실패한다).

## 다른 프로젝트에서 연동하기 (예: maintenance-management)

이 서비스는 4개의 API로 연동한다. 이 중 업로드·토큰 발급 API는 `DOCVIEW_API_KEYS`에 등록된
프로젝트별 키가 있어야 호출 가능하다 — 연동하는 쪽(반드시 서버 사이드, 브라우저 아님)에서
`Authorization: Bearer <발급받은 키>` 헤더를 붙여서 호출한다(`DOCVIEW_API_KEYS` 미설정 시에는
개발 편의상 이 검사를 건너뛴다):

1. `POST /api/documents` (multipart, 필드명 `file`, **API 키 필요**) — 문서 업로드, `{data.id}` 반환.
2. `GET /api/documents/:id` — 변환 상태 폴링, `status: 'ready'`가 될 때까지.
3. `POST /api/documents/:id/tokens` `{viewer, canEdit, canDownload, canPrint, canCopy}`
   (**API 키 필요**) — 토큰 발급, `{data.token, data.editorUrl 또는 data.viewerUrl}` 반환.
   `editorUrl`/`viewerUrl`은 이
   서비스의 상대경로이므로, 호출하는 쪽에서 이 서비스의 접속 주소를 앞에 붙여 `<iframe>`으로
   그대로 띄우면 된다(브라우저가 이 서비스에 직접 WebSocket 연결을 맺어야 하므로 iframe은 반드시
   이 서비스 자체의 URL로 로드해야 한다).
4. `GET /api/view/download?token=...` — `canDownload` 권한이 있는 토큰으로 최신 저장 상태의 원본
   파일 바이트를 내려받는다(동시편집 중인 문서는 유휴 저장 디바운스 후 반영됨).

`maintenance-management`의 "업체별 접속현황" 탭이 이 4개 API로 실제 연동되어 있다 — 자세한 흐름은
`../maintenance-management/WEBVIEWER_INTEGRATION.md` 참고.
