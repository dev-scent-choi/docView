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
