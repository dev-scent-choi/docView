# TODO / 백로그

당장 진행하지 않지만 나중에 다시 다룰 항목들. 각 항목은 실제 작업을 시작할 때 별도 브레인스토밍/설계
과정을 거친다(아래 내용은 요구사항 메모이지 설계가 아님).

## 1. 문서 버전 관리 (DB 기반 이력 · 복원 · Diff)

**요구사항 (사용자 원문 기준):**
- 웹뷰어에서 동시편집 기능 제공 (이미 구현됨 — docx/xlsx/pptx)
- 편집 완료 시 변경된 내용을 DB에 저장
- 문서 버전을 DB로 관리, 버전별 수정 이력(수정자, 수정일시, 변경 사유) 조회 가능
- 이전 버전 문서 복원 가능
- 버전 간 문서 비교(Diff)로 변경 내용 확인 가능

**결정된 것:**
- DB는 **PostgreSQL** 사용 (SQLite 아님) — 실제 서버 필요, 연결 정보/드라이버/마이그레이션 도구 준비 필요

**시작 전 확인할 것:**
- 이미 쓰고 있는 Postgres 인스턴스가 있는지, 아니면 이 프로젝트용으로 새로 준비할지
- docx/xlsx/pptx 세 문서 타입의 데이터 모델이 완전히 달라서(ProseMirror / GridDocument / 슬라이드 포인터 배열) Diff·버전 저장 방식이 타입별로 다르게 설계될 가능성 높음
- 현재 프로젝트엔 로그인/인증 시스템이 없음(자유 텍스트 `viewer` 문자열만 존재) — "수정자" 필드의 신뢰 수준이 이 한계를 따라감
- `src/collab/room.ts`의 `flushToStore`가 모든 동시편집 저장이 지나가는 유일한 지점 — 버전 스냅샷 훅이 여기 필요

**독립성:** 아래 2번(메모리 이슈)과는 서로 관계 없음 — 어느 쪽을 먼저 해도 무방.

---

## 2. 대용량 xlsx 동시편집 시 서버 메모리(Room) 구조 개선

2026-09-09 세션에서 진단, 그리고 아래 1번(baseline 전체 복제본 제거)·2번(동시 룸 개수 상한)·
4번(DELETE 시 Room 미정리)·3번(장기 지연 materialization, 2026-09-10 완료)까지 전부 완료.
지금도 임시방편(`.npmrc`의 힙 한도 16GB)이 남아 있음 — 이건 별개로 겪은 ExcelJS `writeBuffer()`
자체의 대용량 한계(아래 "저장 실패" 항목 참고)에 대한 안전망이라 이번 메모리 개선과 무관하게 유지.

**진단된 문제(원래):**
- 동시편집 룸(Room)이 대용량 문서 하나당 canonical 그리드 + baseline 그리드(전체 복제본) +
  원본 파일 바이트, 세 벌을 메모리에 동시 보관 — 29만 행급 파일 하나 여는 데 실측 약 8GB
- 동시에 열린 Room 개수에 상한이 없고, 유휴/메모리 압박 기반 정리도 없음(닫힐 때만 즉시 해제됨)
- `DELETE /api/documents/:id`가 메모리상 살아있는 Room은 안 건드림(부수적 발견 — 아래 4번에서 해결)

**1. baseline 그리드 전체 복제본 제거 — 완료(2026-09-09):**
- `Room.baselineGrid`(canonicalGrid 전체 깊은 복사본)를 없애고 대신 세 가지 가벼운 기록으로
  대체: `baselineSheetExtents`(시트별 rowCount/colCount만), `dirtyCells`(실제로 편집된 칸의
  좌표 집합, gridCellKey 형식), `structuralOps`(행/열 삽입·삭제 연산의 순서 기록).
  `exportXlsxFromGrid`(`src/xlsx/collab.ts`)는 이제 셀 단위 값 비교 대신 원본을 다시 로드한
  워크시트에 `structuralOps`를 같은 순서로 재생(spliceRows/spliceColumns)한 뒤, `dirtyCells`에
  담긴 좌표만 canonicalGrid의 현재 값으로 덮어쓴다.
- **실측 메모리 절감**(실제 29만 행 파일 기준): canonical 그리드만 약 535MB, 기존 방식(canonical +
  baseline 전체 복제본)은 약 1,353MB — 이 두 자료구조만 놓고 보면 약 60% 감소.
- **알려진 트레이드오프**: 한 번 편집된 칸은 이후 원래 값으로 되돌려도 "편집됨" 표시가 남는다 —
  원래 수식이었던 칸을 편집했다가 정확히 그 수식의 계산값과 같은 문자열로 되돌리면, 예전에는
  원본 수식이 보존됐지만 지금은 정적 값으로 저장된다. 실사용에서 마주치기 어려운 좁은 차이라
  받아들임.
- 회귀 테스트: `tests/xlsx-collab.test.ts`(새 시그니처로 전면 수정), `tests/xlsxExportYielding.test.ts`
  (더티-셀 쓰기 루프의 양보 지점 검증으로 재작성), 기존 `tests/roomStructureSync.test.ts`/
  `tests/roomRowColShift.test.ts`(mid-sheet 삽입·삭제, 여러 연산 연쇄, 1600행 시트, 새 시트 추가
  등 34개 테스트)가 전부 그대로 통과 — 이 테스트들이 재설계의 실질적인 검증 역할을 했다.

**2. 동시 대용량 xlsx 룸 개수 상한 — 완료(2026-09-09):**
- `room.ts`의 `checkRoomCapacity(docId, ext)` — 원본 파일이 5MB 이상(`LARGE_XLSX_BYTE_THRESHOLD`)인
  xlsx 문서에 한해, 동시에 열려 있는 그런 문서가 3개(`MAX_CONCURRENT_LARGE_XLSX_ROOMS`)를 넘으면
  **새로 여는** 시도만 거부한다(이미 열려 있는 문서의 추가 접속자는 항상 허용). 파일 크기는
  `store.originalSize`(새로 추가, 파싱 없이 바로 조회)로 판단 — 룸을 열어(파싱해) 봐야 크기를
  아는 모순을 피하기 위함.
- `routes/edit.ts`가 `joinEditSession` 호출 전에 이 검사를 하고, 거부되면 WS 종료 코드 4429 +
  승인된 문구("사용자가 많습니다. 잠시 후 다시 시도해주세요.")로 소켓을 닫는다.
- 클라이언트(`grid-editor.ts`)는 `provider.on('connection-close', ...)`으로 4400번대 종료 코드를
  받으면 y-websocket의 기본 무한 재연결을 멈추고(`provider.shouldConnect = false`) 그 이유를
  화면에 그대로 보여준다 — 안 그러면 "연결 끊김 — 재연결 중…"만 계속 보이고 실제 이유(서버 혼잡)를
  사용자가 알 수 없었다.
- 회귀 테스트: `tests/roomCapacity.test.ts`(신규, 5개) — 소형 파일 무제한, 비-xlsx 무제한, 4번째
  대용량 룸 거부, 같은 docId 중복 체크 시 슬롯 안 먹음, 룸이 완전히 닫히면 슬롯이 반환되는지까지 확인.

**4. `DELETE /api/documents/:id`의 Room 미정리 — 완료(2026-09-09):**
- `room.ts`에 `forceCloseRoom(docId, reason?)` 추가 — 해당 문서의 룸이 열려 있으면 모든 연결을
  WS 4410("문서가 삭제되었습니다.")로 닫고 closeConn의 정상 "마지막 참여자 퇴장" 경로를 그대로
  태워 룸·대용량 상한 슬롯을 즉시 정리한다.
- `documents.ts`의 DELETE 핸들러가 `store.deleteDocument` **다음에** 이걸 부른다 — 순서가
  중요하다: 먼저 지워야 `flushToStore`의 `meta` null 체크가 자동저장으로 파일을 되살리는 걸
  막아준다(forceCloseRoom 주석 참고).
- 회귀 테스트: `tests/roomForceClose.test.ts`(신규, 5개) — 4410 코드로 실제 닫히는지, 룸이
  메모리에서 사라지는지, 대용량 룸 상한 슬롯이 즉시 반환되는지, 룸이 없을 때 안전한 no-op인지,
  삭제 후 자동저장이 파일을 되살리지 않는지까지 확인.

**저장 실패(`Invalid array length`) — 완료(2026-09-09), 스트리밍 폴백으로 해결:**
- 실제 버그 하나 먼저 발견·수정함: `exportXlsxFromGrid`(`src/xlsx/collab.ts`)의 행/열 삭제
  감지 로직이 방금 재로드한 워크북의 raw `sheet.rowCount`/`columnCount`(서식만 있어도
  세는 ExcelJS 고유 카운트)를 기준으로 비교하고 있었음 — 편집이 전혀 없어도 "행이 삭제됐다"고
  오판해 `spliceRows`/`spliceColumns`가 잘못 호출될 수 있었음(이미 위 1번 재설계에 흡수됨).
- **하지만 이 수정만으로는 이 특정 파일(292,244행 × 36열, 서식 거의 전 셀)의 저장 문제가
  끝나지 않았다.** `workbook.xlsx.writeBuffer()` 자체가 `RangeError: Invalid array length`로
  실패 — ExcelJS 4.4.0(최신) 자체가 이 정도 규모의 워크북을 통째로 메모리에서 재직렬화하는 걸
  감당하지 못하는 라이브러리 레벨 한계로 확인(수정 없이 load+writeBuffer만 해도 동일 실패).
- **해결: `src/xlsx/exportStreaming.ts`에 `exportXlsxFromGridStreaming` 신규 추가** —
  `ExcelJS.stream.xlsx.WorkbookWriter`(행 단위 스트리밍 직렬화, 로컬 임시 파일에 쓴 뒤 읽어
  들여 반환)로 원본을 행 단위로 복사하면서 `dirtyCells`에 담긴 편집된 칸만 canonicalGrid 값으로
  덮어쓴다. **기존 `exportXlsxFromGrid`는 전혀 건드리지 않았다** — `room.ts`에 새로 추가한
  `exportXlsxWithFallback`이 정상 경로를 먼저 시도하고, 그게 실제로 실패할 때만(사전 임계값이
  아니라 실패 후 재시도) 이 스트리밍 함수로 넘어간다. 정상 크기 파일은 전혀 영향받지 않는다.
- **알려진 스코프 제한(의도적, 문서화됨)**: 이 폴백은 행/열 삽입·삭제(`structuralOps`)를
  지원하지 않는다(있으면 조용히 잘못 쓰는 대신 즉시 예외) — 이 정도로 거대한 파일에서 그런
  구조 변경까지 겪는 일은 드물다고 보고 스코프를 좁혔다. 시트 보호 재적용도 하지 않는다.
- **실측 검증(실제 문제의 원인이었던 그 파일로)**: 파싱 40초, 스트리밍 내보내기(편집 2건 포함)
  66초, 출력 64.5MB, 재오픈 검증 통과, 메모리 5.8GB(16GB 힙 한도 내). 최초에 이 파일이
  "connection refused"로 아예 안 열리던 문제(이제 열람은 이미 해결됨)에 이어, **저장까지
  end-to-end로 성공**하는 것을 확인.
- 회귀 테스트: `tests/xlsxExportStreaming.test.ts`(신규, 8개 — 구조연산 거부, 무편집 라운드트립,
  더티 셀만 덮어쓰기, 서식 적용, 협업 중 추가된 새 시트, 직접 rowCount 축소, 시트 순서, 시트
  제거), `tests/roomExportFallback.test.ts`(신규, 2개 — 정상 경로 실패 시 실제로 폴백이 성공하는지,
  폴백마저 실패하면 에러가 조용히 삼켜지지 않고 그대로 드러나는지).

**3. 진짜 지연(lazy) materialization — 완료(2026-09-09~10):**
- `Room.canonicalGrid: GridDocument`(시트 전체를 실제 GridCell 배열로 담아 seedRoom 시점에 파일
  전체를 파싱해 메모리에 올리던 구조)를 `Room.lazyGrid: LazyGridStore`로 완전히 교체 —
  편집된 칸만 `dirtyOverlay`(Map)에 두고, 나머지는 온디스크 행 저장소(JSONL, 룸당 임시 파일)에서
  그때그때 읽어 작은 LRU 캐시(`readCache`)에 담는다. 설계 문서:
  `docs/superpowers/specs/2026-09-09-xlsx-lazy-canonical-grid-design.md`, 계획 문서:
  `docs/superpowers/plans/2026-09-09-xlsx-lazy-canonical-grid.md`(13개 태스크, 전부 완료).
- 수집(`ingestXlsxLazily`)은 ExcelJS의 스트리밍 리더로 한 번에 한 행만 메모리에 두고 온디스크
  저장소를 만든다. 읽기(`getCell`/`getRow`)는 dirtyOverlay → readCache → 디스크 순으로 확인한다.
  구조연산(행/열 삽입삭제)은 실제 셀 배열이 아니라 정수 매핑 배열(`logicalToOriginalRow/Col`)만
  스플라이스하므로 30만 행 시트에서도 항상 즉시·동기로 끝난다(예전엔 청크+양보가 필요한 무거운
  배열 패딩이었다).
- **실측 결과**: 이 항목의 발단이었던 그 실제 파일(292,244행×36열, 57MB)을 열었을 때 canonicalGrid
  파싱만으로 약 535MB(baseline 제거 후 기준)였던 것이, LazyGridStore 재설계 이후 강제 GC 기준
  steady-state heap 32~38MB로 줄었다(10배 이상, 파일 크기에 비례하지 않음 — 설계 목표 달성).
- 회귀 테스트: 기존 구조연산 회귀 스위트(`tests/roomStructureSync.test.ts`/
  `tests/roomRowColShift.test.ts`, 34개)가 재설계 후에도 전부 그대로 통과. 전체 스위트
  41 files / 298 tests PASS, `npx tsc --noEmit` 클린.
- **알려진 캐비앗(운영 환경에서 재확인 필요)**: 이 세션의 개발 머신에서, 그 실제 대용량 파일을
  room.ts의 실제 시딩 경로(`joinEditSession`)로 열람할 때 방금 정상적으로 쓰이고 닫힌 LazyGridStore
  행 저장소 파일의 `fs.open`이 간헐적으로 `ENOENT`를 던지는 현상을 반복 관측했다 — 실패 지점이
  실행마다 달랐고, 같은 파일을 더 단순한 직접 호출로 전량 읽을 때는 재현되지 않아 이 코드의
  결정적 버그라기보다 이 환경(Windows, 짧은 경로 표기 `OPENMI~1`이 나타나는 프로필)의
  OS/파일시스템 계층 불안정성으로 보인다(같은 임시 폴더에 대한 기본 `fs.watch` 호출조차 별개로
  libuv 내부 단언 실패로 죽는 것을 확인함). `fs.rm`에 이미 쓰던 것과 같은 재시도 관용구
  (`openRowStoreWithRetry`, 5회·100ms 간격)를 행 저장소 읽기에도 추가했지만 이 환경에서는
  완전히 해소되지 않았다 — 실제 배포 환경(짧은 경로 표기가 없는 Linux 서버 등)에서는 이 문제가
  없을 가능성이 높지만 확인하지 못했다. 대용량 파일 열람 시 간헐적 오류가 보고되면
  `src/xlsx/lazyGridStore.ts`의 `openRowStoreWithRetry` 주석부터 볼 것.

**참고:** 자세한 배경은 이 세션의 대화 기록 참조. 다시 시작할 때 코드 상태를 먼저 재확인할 것
(시간이 지나 내용이 달라졌을 수 있음).

---

## 3. pptx 텍스트박스 콘텐츠 동시편집 — 완료(2026-09-11), 발견된 버그 1건도 Task 13에서 수정 완료

계획: `docs/superpowers/plans/2026-09-10-pptx-content-editing.md` (13개 태스크, 전부 완료).
설계: `docs/superpowers/specs/2026-09-10-pptx-content-editing-design.md`.
SDD 기록: `.superpowers/sdd/2026-09-10-pptx-content-editing/`(태스크별 브리프·리포트·리뷰 diff,
`progress.md`에 전체 태스크 이력).

**구현된 것 (Task 1~11, 이전 세션에서 완료·리뷰됨):**
- 서버: `ppt/slides/*.xml`의 텍스트박스(`<p:sp>`+`<p:txBody>`)를 지연 로딩(슬라이드를 처음 열 때만
  파싱해 Yjs에 시딩)하고, 저장 시 편집된 도형만 원본 XML에 병합(`src/pptx/collab.ts`
  `exportPptxFromSlides`, `src/pptx/slideContent.ts`). 슬라이드 복제 시 원본이 아니라 "지금까지
  세션에서 편집된 내용"을 깊은 복사. baseline과 비교해 실제로 안 건드린 도형은 병합 자체를
  건너뛰어 서식이 안 깎이게 함(Task 6에서 발견·수정된 Critical 버그 2건 포함).
- 클라이언트: 왼쪽 슬라이드 레일 + 오른쪽 콘텐츠 캔버스 레이아웃(`editor-slides.ts`), 텍스트박스별
  실시간 ProseMirror(`pptxTextSchema`) + y-prosemirror 바인딩, 서식 툴바(굵게/기울임/밑줄/색),
  텍스트박스 추가/삭제 UI, 읽기 전용 모드(아직 아무도 안 연 슬라이드는 안내 문구로 대체).

**Task 12(이번 작업) — 통합 테스트 + 실제 파일 end-to-end 검증:**
- `tests/roomPptxContent.test.ts`에 새 통합 테스트 추가: 두 개의 `FakeSocket` + mirror `Y.Doc`이
  같은 룸에 join하고, 같은 텍스트박스(`Y.XmlFragment`)에 대해 **서로의 편집을 아직 보지 못한 채**
  각자 로컬 트랜잭션으로 서로 다른 위치에 서로 다른 텍스트를 삽입 → 캡처한 update를 room.ts의
  진짜 WebSocket 메시지 경로(`handleMessage`)로 순서대로 흘려보내 → 최종 canonical 문서·두 클라이언트
  mirror·실제 export된 pptx 바이트 세 곳 모두에서 두 편집이 전부 살아남았는지 확인(진짜 CRDT 병합
  증명, "둘 다 연결이 죽지 않았다" 수준이 아님). 전체 스위트: `npx vitest run` 44 files / 330 tests
  전부 PASS, `npx tsc --noEmit` 클린.
- **실제 PowerPoint 파일로 수동 E2E 검증**(python-pptx로 제작한 6슬라이드 `.pptx`: 굵게/일반/
  빨간색-기울임이 한 문단에 섞인 텍스트박스, 3x3 표가 있는 슬라이드, 실제 PNG 이미지가 있는 슬라이드,
  절대 건드리지 않을 대조군 슬라이드 2장, 복제용 소스 슬라이드 1장). 실행 중이던 dev 서버(포트 8000)에
  Playwright(headless Chromium)로 실제 업로드 → 문서 목록의 "동시편집 열기" 클릭 → 왼쪽 레일에서
  빈 슬라이드 추가/슬라이드 복제/슬라이드 삭제/드래그로 순서 바꾸기 → 오른쪽 캔버스에서 기존
  텍스트박스 끝에 타이핑 후 굵게 적용(기존 굵게/기울임/빨간색 런은 그대로 유지됨을 확인) → 다른
  텍스트박스 삭제 → 새 텍스트박스 추가 후 타이핑 → Ctrl+S로 저장(2회, 두 번째 저장 전에 추가
  편집도 함) → `/api/view/download`로 실제 다운로드 → 이 프로젝트의 실제 `fflate` unzip으로 재파싱해
  원본과 비교. 확인된 것:
  - 한 번도 안 연 대조군 슬라이드 2장, 표 슬라이드, 이미지 슬라이드가 **바이트 단위로 원본과 동일**
    (슬라이드 XML도, 이미지 미디어 파일도).
  - 표(`<a:tbl>`, 셀 텍스트)·이미지(`<p:pic>`, `ppt/media/image1.png`)가 손상 없이 그대로 보존.
  - 편집한 텍스트박스: 안 건드린 굵게/일반/빨간색-기울임 런이 그대로 남아 있고, 새로 타이핑한
    내용도 정확한 서식(bold `b="1"`)으로 저장됨.
  - 새로 추가한 텍스트박스가 **두 번째 저장에서도 중복되지 않고 그 자리에서 갱신됨**(Task 6에서
    고친 "저장마다 텍스트박스 중복" 버그가 실제 브라우저 E2E 흐름에서도 재발하지 않음을 재확인).
  - 슬라이드 복제본이 소스와 독립적인 콘텐츠를 가짐, 슬라이드 삭제·순서 바꾸기 모두 저장된 파일의
    `ppt/presentation.xml`/`<p:sldId>` 순서에 정확히 반영됨.
  - 어떤 슬라이드에서도 도형 id 중복 없음.

**알려진 문제 1건 — 해결됨(Task 13, 2026-09-11):**
- **텍스트박스 삭제(×버튼)가 저장된 파일에는 반영되지 않는다.** Task 12의 수동 검증에서 텍스트박스
  하나를 삭제 → 두 번 저장 → 다운로드한 실제 파일을 열어 보니, 삭제한 텍스트박스의 `<p:sp>` 블록이
  **원본 그대로 남아 있었다**(라이브 세션/화면에서는 정상적으로 사라져 보임 — 다른 접속자에게도
  실시간으로 사라진 것처럼 보임 — 그런데 저장된 파일에는 한 번도 지워지지 않음).
  - 근본 원인: `src/pptx/collab.ts`의 export 병합 루프(`exportPptxFromSlides`)가
    `for (const [shapeId, shapeMap] of perSlide.entries())`로 **지금 `slideContents`에 남아 있는
    도형만** 순회했다. 삭제된 도형은 애초에 이 순회에 아예 안 걸리므로 "이 도형은 없어졌으니
    XML에서 제거하라"는 코드 경로 자체가 존재하지 않았다(`src/pptx/slideContent.ts`에
    `appendNewShape`/`spliceShapeTxBody`만 있고 도형을 XML에서 제거하는 함수가 없었음).
  - 이건 특정 태스크의 구현 실수가 아니라 **설계 문서 자체의 공백**이었다 — 설계 문서
    (`docs/superpowers/specs/2026-09-10-pptx-content-editing-design.md`) "Export" 절은
    "기존 도형이면 교체", "새 도형이면 추가" 두 경우만 명시하고, "이전엔 있었는데 지금은 없는
    도형(=삭제됨)을 원본 XML에서 제거"하는 세 번째 경우를 아예 다루지 않았다. 이전 태스크(Task 11)의
    수동 확인도 "라이브 화면/DOM에서 사라지는지"만 봤고, 저장 → 다운로드 → 재파싱까지 가는 전체
    왕복은 검증하지 않았다(자동화 테스트인 `tests/roomPptxContent.test.ts`에도 삭제→export 시나리오를
    확인하는 테스트가 없었음).
  - **수정 내용(Task 13)**: `src/pptx/slideContent.ts`에 `removeShape(slideXml, shapeId)` 추가 —
    `hasShape`/`spliceShapeTxBody`와 정확히 같은 `<p:sp>` 블록 탐지 정규식을 써서(id 매칭을 숫자로
    제한하지 않으므로 `new-<uuid>` 형태의 id도 매치) 해당 블록 전체를 제거한다. 다른 도형은 전혀
    건드리지 않고, 못 찾으면 throw한다(이 파일의 기존 "절대 조용히 실패하지 않는다" 원칙과 동일).
    `exportPptxFromSlides`의 기존 도형별 update/append 루프(`for (const [shapeId, shapeMap] of
    perSlide.entries())`) 직후, 슬라이드 XML을 zip에 확정하기 전에 제거 단계를 추가: 이미 있는
    `baselineShapes`(Task 6이 "수정 안 된 도형은 건드리지 않기" 최적화용으로 만든, `sourceSlidePart`별
    원본 도형 `Map<string, ParsedShape>`)를 순회하며, `perSlide.has(shapeId)`가 false인(=baseline에는
    있었는데 지금 라이브 상태에는 없는, 즉 사용자가 삭제한) shapeId에 대해서만 `removeShape`를 호출한다.
    `baselineShapes`가 null인 경우(`entry.kind === 'blank'`, 애초에 baseline이 없는 슬라이드)는
    "삭제"라는 개념이 성립하지 않으므로 자연히 건너뛴다.
  - **회귀 위험 검증**: `removeShape` 호출 대상은 항상 "baseline에는 있고 perSlide에는 없는" shapeId로만
    한정되므로, 위 update/append 루프가 건드리는 도형(=perSlide에 있는 도형)과는 절대 겹치지 않는다 —
    Task 6의 두 Critical 버그(저장마다 새 텍스트박스 중복 추가, 안 건드린 도형의 서식 유실)를 만드는
    `hasShape`/baseline-비교 로직과 이번 제거 로직은 서로 다른 shapeId 집합에서만 동작해 간섭하지
    않는다. `removeShape`가 호출된 shapeId는 원본 XML에서 파싱해 baseline에 들어간 것이므로(=
    export 시작 시점의 슬라이드 XML에 실제로 그 `<p:sp>` 블록이 존재), 그리고 위 루프가 `<p:sp>`
    블록을 지우는 일은 전혀 없으므로, `removeShape`의 "못 찾으면 throw"는 정상 사용에서는 결코
    발생하지 않는 안전장치일 뿐임을 확인했다.
  - **테스트**: `tests/pptxSlideContent.test.ts`에 `removeShape` 단위 테스트 3개(다른 도형 안 건드림,
    `new-` 접두어 id도 동작, 없는 id는 throw). `tests/roomPptxContent.test.ts`에 서버 통합 테스트 4개
    (삭제 후 export하면 실제로 사라짐+다른 도형은 그대로, 두 번째 저장에서도 안 되살아남, 한 번도
    export 안 된 `new-` 도형을 삭제하면 안전한 no-op, 슬라이드의 모든 도형을 지워도 에러 없이 빈
    `<p:spTree>`로 export됨). 전체 스위트(`npx vitest run`, 44 files / 336 tests) + `npx tsc --noEmit`
    모두 클린 — Task 6의 기존 회귀 테스트(Critical #1/#2, Important #6)도 전부 그대로 통과해 이번
    변경이 그 수정들을 깨뜨리지 않았음을 확인.
  - **실제 pptx로 재확인**: Task 12가 쓴 것과 같은 실제 6슬라이드 `.pptx` 픽스처로, 실제 production
    함수(`seedSlidesFromPptx`/`seedSlideContentFromXml`/`exportPptxFromSlides`/`parseSlideShapes`)를
    직접 호출해 텍스트박스 하나(굵게/일반/빨간색-기울임 런이 섞인 슬라이드의 두 도형 중 하나)를
    삭제 → 저장 → 재파싱 → 삭제한 도형이 실제로 없어졌고 남은 도형은 서식까지 바이트 단위로 그대로
    보존됨을 확인. 두 번째 저장(직전 저장 결과를 다음 원본으로 삼음, `flushToStore`와 동일한 패턴)
    에서도 되살아나지 않음을 확인. 한 번도 안 연 다른 5개 슬라이드는 두 번의 저장 모두에서 원본과
    바이트 단위로 동일했다.

**그 외 참고(이번 세션 범위 밖, 새 버그 아님):** 슬라이드 레일 카드 썸네일이 `kind: 'duplicate'`
슬라이드에 대해 깨진 이미지 요청(`/api/view/pages/null`)을 보내는 것을 수동 E2E 중 관찰함 — 이
콘텐츠 편집 기능이 생기기 전부터(`811b376` 커밋, 최초 슬라이드 관리 UI) 있던 코드라 이번 플랜과
무관한 기존 이슈. 텍스트 콘텐츠나 저장 데이터와는 무관한 순수 썸네일 표시 문제.

이 세션의 "원본 유실이나 버전 관리만 잘되면 됨"이라는 최우선 제약(레이아웃/마스터 상속 미해결,
표/차트/이미지 편집 제외 등은 설계 문서에 이미 문서화된 의도된 스코프 제한)은 지켜졌다 — 위
검증에서 확인한 대로 안 연 슬라이드·표·이미지는 전부 바이트 단위로 보존됐다. "삭제" 기능 자체가
저장 시 무효화되던 위 문제도 Task 13에서 고쳤다 — 이 계획(13개 태스크)은 이제 전부 완료다.
