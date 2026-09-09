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

2026-09-09 세션에서 진단, 그리고 아래 1번(baseline 전체 복제본 제거)·2번(동시 룸 개수 상한)
완료. 3번(장기 지연 materialization)은 아직 미착수. 지금도 임시방편(`.npmrc`의 힙 한도 16GB)이
남아 있음 — 이건 별개로 겪은 ExcelJS `writeBuffer()` 자체의 대용량 한계(아래 "저장 실패" 항목
참고)에 대한 안전망이라 이번 메모리 개선과 무관하게 유지.

**진단된 문제(원래):**
- 동시편집 룸(Room)이 대용량 문서 하나당 canonical 그리드 + baseline 그리드(전체 복제본) +
  원본 파일 바이트, 세 벌을 메모리에 동시 보관 — 29만 행급 파일 하나 여는 데 실측 약 8GB
- 동시에 열린 Room 개수에 상한이 없고, 유휴/메모리 압박 기반 정리도 없음(닫힐 때만 즉시 해제됨)
- `DELETE /api/documents/:id`가 메모리상 살아있는 Room은 안 건드림(부수적 발견, **아직 미해결** —
  아래 "남은 항목" 참고)

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

**남은 항목:**
2. (장기) 대용량 시트를 통째로 메모리에 올리지 않고 필요한 부분만 그때그때 읽는 진짜 지연
   materialization — 이번 세션에 끝낸 xlsx 윈도잉 프로젝트급의 별도 큰 작업.

**저장 실패(`Invalid array length`) — 원인 규명 완료, 부분 수정:**
- 실제 버그 하나 발견·수정함: `exportXlsxFromGrid`(`src/xlsx/collab.ts`)의 행/열 삭제
  감지 로직이 방금 재로드한 워크북의 raw `sheet.rowCount`/`columnCount`(서식만 있어도
  세는 ExcelJS 고유 카운트)를 기준으로 비교하고 있었음 — 원래는 `baselineGrid`의
  `actualRowCount` 기반 카운트와 비교했어야 함. 이 버그 때문에 편집이 전혀 없어도
  "행이 삭제됐다"고 오판해 `spliceRows`/`spliceColumns`가 잘못 호출될 수 있었음.
  회귀 테스트 추가 완료(`tests/xlsx-collab.test.ts`), 전체 테스트 통과.
- **하지만 이 수정만으로는 이 특정 파일(292,244행 × 36열, 서식 거의 전 셀)의 저장 문제가
  끝나지 않음.** 위 수정을 적용한 후에도 `workbook.xlsx.writeBuffer()` 자체가
  `RangeError: Invalid array length`로 실패함 — 독립적으로 확인(이 파일을 로드 후
  아무 수정 없이 바로 `writeBuffer()`만 호출해도 동일하게 실패, 힙 제한 없는 환경에서는
  OOM). 즉 ExcelJS 4.4.0(최신) 자체가 이 정도 규모의 워크북을 통째로 메모리에서
  재직렬화하는 걸 감당하지 못하는 라이브러리 레벨 한계.
- **우회 방법 확인 완료(스파이크):** `ExcelJS.stream.xlsx.WorkbookWriter`(스트리밍 writer)로
  시트/행 단위로 순회하며 새 워크북에 다시 써내려가는 방식은 이 파일에서 성공함
  (64.5MB 출력, 행/시트/서식 모두 보존, 재오픈 검증 완료). 메모리도 안정적 — 일반
  방식은 16GB 힙에서도 죽었는데 스트리밍 방식은 파일 로드(4.5GB) 위에 추가로 약
  1.2GB만 사용.
  - 단, API가 완전히 다름: in-place 편집(`spliceRows` 등) 없이 원본을 행 단위로
    복사하면서 편집된 셀만 canonicalGrid 값으로 덮어쓰는 방식으로 export 로직을
    다시 짜야 함. 삭제된 행은 복사 단계에서 건너뛰는 식으로 재구현.
  - 출력도 버퍼 반환이 아니라 파일 스트림에 직접 쓰는 방식이라 `room.ts`의
    `flushToStore` 인터페이스도 손볼 가능성 있음.
  - **범위가 커서 별도 브레인스토밍/설계+계획을 거쳐야 하는 작업으로 판단, 지금은
    보류.** 다음에 시작할 때 코드 상태(특히 `exportXlsxFromGrid`)를 먼저 재확인할 것.

**참고:** 자세한 배경은 이 세션의 대화 기록 참조. 다시 시작할 때 코드 상태를 먼저 재확인할 것
(시간이 지나 내용이 달라졌을 수 있음).
