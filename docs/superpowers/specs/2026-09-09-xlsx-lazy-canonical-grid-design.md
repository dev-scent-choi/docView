# xlsx 협업 편집 — 지연(Lazy) Canonical Grid 설계

## 목표

`Room.canonicalGrid`(`src/collab/room.ts`)가 xlsx 파일 전체를 `parseXlsxToGrid`로 한 번에 파싱해
메모리에 통째로 올려두는 현재 방식을, **사용자가 실제로 편집한 칸만 메모리에 두고 나머지는 디스크에서
그때그때 읽어오는 방식**으로 바꾼다. 목표는 특정 파일 크기나 동시 접속자 수 숫자가 아니라 —
"파일이 아무리 커도 메모리 사용량이 파일 크기가 아니라 편집량에 비례한다"는 **원칙적으로 올바른
구조**를 만드는 것이다.

## 배경

2026-09-09 세션에서 대용량 xlsx(292,244행 × 36열, ~57MB) 협업 편집 중 겪은 문제들은 이미 별도로
해결됐다(`docs/TODO.md` 참고 — baseline 그리드 복제본 제거, 동시 대용량 룸 상한, 문서 삭제 시 룸
정리, ExcelJS `writeBuffer()` 대용량 저장 실패의 스트리밍 폴백). 이 문서가 다루는 것은 그 이후에도
남아 있는 근본 원인: **룸을 여는 순간 파일 전체가 여전히 메모리에 파싱된다**는 것이다(baseline
제거 후에도 이 파일 하나에 canonical 그리드만으로 실측 약 535MB).

## 범위

- **포함**: xlsx 협업 편집 룸(`src/collab/room.ts`)의 canonical 상태 표현 전체를 새 구조로 교체.
  **크기와 무관하게 모든 xlsx 문서**에 통일 적용한다(작은 파일 전용 경로를 따로 두지 않는다 —
  작은 파일은 디스크 접근 횟수가 적어 체감 차이가 없고, 코드 경로를 하나로 유지하는 게 낫다).
- **제외(이번 프로젝트 범위 밖)**:
  - 읽기전용 뷰어가 쓰는 `grid.json`/`gridindex.json` 생성 경로(`src/spreadsheet.ts`의
    `serializeGrid`/`computeLineOffsets`, `src/routes/viewer.ts`의 `loadGridUncached`) — 이건
    "첫 화면 속도"라는 다른 목적의 요청 단위 캐시라 이번 문제(협업 편집 세션의 상주 메모리)와
    결이 다르다. 형식(JSONL + 바이트 오프셋 인덱스)은 참고하되 코드는 공유하지 않는다.
  - 문서 버전 관리(DB) 기능([[project-webviewer-version-history]])과는 무관.
  - docx/pptx 협업 편집 — canonical 상태 자체가 xlsx와 근본적으로 다른 자료구조라 대상 아님.

## 현재 구조 요약

`Room.canonicalGrid: GridDocument | null` — `seedRoom`(`room.ts`)이 룸 생성 시
`parseXlsxToGrid(originalBytes)`(`src/spreadsheet.ts`)로 파일 전체를 한 번에 파싱해 채운다.
`GridDocument.sheets[i].rows`는 모든 행 × 열의 `GridCell | 0`을 담은 배열의 배열이다. 이후 모든
읽기(수식 계산, Yjs 페이지 시딩)와 쓰기(`applyCellsChangeToCanonical` 등)가 이 배열을 직접
인덱싱한다. 행/열 삽입·삭제는 `gridShift.ts`가 이 배열 자체를 `splice`한다.

## 새 구조: `LazyGridStore`

룸마다 하나씩 갖는, `canonicalGrid: GridDocument`를 대체하는 새 상태:

```ts
interface LazyGridStore {
  readonly docId: string;
  readonly rowStorePath: string;          // 이 룸 전용 온디스크 원본 행 저장소(JSONL)
  readonly rowIndex: ReadonlyArray<ReadonlyArray<{ byteOffset: number; byteLength: number } | undefined>>;
  // ^ sheetIdx -> 원본 행 번호(ingestion 시점 기준, 0-based) -> 그 행의 디스크 위치. 불변(ingestion 이후 절대 안 바뀜).
  readonly sheetMeta: LazyGridSheetMeta[];  // name, rowCount, colCount, freezeRows/Cols, colWidths, merges — GridSheet에서 rows만 뺀 모양
  readonly styles: GridStyle[];             // 공용 스타일 팔레트(작음, 상주)
  readonly logicalToOriginalRow: number[][];// sheetIdx -> 논리 행(현재 번호) -> 원본 행 번호(-1=원본에 없던 새 행)
  readonly logicalToOriginalCol: number[][];// 위와 동일(열 기준)
  readonly dirtyOverlay: Map<string, GridCell | 0>; // gridCellKey -> 실제로 편집된 칸의 현재 값(옛 dirtyCells를 값까지 담도록 확장)
  readonly readCache: LRUCache<string, ParsedRow>;  // "sheetIdx:원본행번호" -> 파싱된 행(캐시 단위는 칸이 아니라 행)
}

// ParsedRow = Array<GridCell | 0> — GridSheet.rows의 한 행과 완전히 같은 모양(JSONL 한 줄을 그대로 파싱한 것).
```

`readCache`의 상한은 기본 5,000행(시트당이 아니라 스토어 전체 — 페이지 몇 개 + 흩어진 수식 참조를
넉넉히 덮는 크기, 튜닝 가능한 상수로 둔다). `removedSheetIndices`/`sheetDisplayOrder`(Room의 기존
필드)는 이 설계와 무관하게 지금 그대로 유지된다 — 삭제된 시트·탭 순서는 여전히 Room 레벨에서
별도로 추적하고, export가 지금처럼 독립적으로 참조한다.

**왜 캐시 단위가 칸이 아니라 행인가**: 디스크 I/O도, 온디스크 저장 형식(JSONL 한 줄 = 한 행)도
행 단위다. 한 번 읽은 행의 모든 칸을 한 번에 캐시해 두면, 한 페이지(500행 × 36열)를 채울 때
디스크 접근이 각 칸이 아니라 각 행당 한 번으로 줄어든다.

**왜 `logicalToOriginalRow/Col`가 필요한가**: 행/열 삽입·삭제가 일어나면 "지금 3번째 행"이
"원래는 2번째 행이었다" 같은 어긋남이 생긴다. 이 배열이 그 대응을 유지한다 — 매 읽기마다
`structuralOps` 로그 전체를 재생하는 대신, 삽입·삭제가 일어난 그 순간에 이 작은 정수 배열만
한 번 `splice`해 두면 이후 읽기는 배열 인덱싱 한 번으로 끝난다(오늘 `insertCanonicalRows` 등이
행 데이터 배열에 하던 것과 똑같은 연산을, 셀 값이 아니라 정수 하나에 대해서 하는 것뿐이라
비용은 무시할 수준이다 — 오늘 실측된 "0.002~0.97ms/30만 행" 그대로 적용된다).

## 수집(ingestion) 파이프라인

`seedRoom`이 룸을 생성할 때, `parseXlsxToGrid` 대신 새 함수(가칭 `ingestXlsxLazily`, 위치 제안:
`src/xlsx/lazyGridStore.ts`)를 부른다:

1. `ExcelJS.stream.xlsx.WorkbookReader`(이미 `exportStreaming.ts`에서 쓴 `WorkbookWriter`의
   짝, forward-only)로 원본 바이트를 앞에서부터 훑는다.
2. 시트 메타(이름, 스타일 팔레트, 병합, 틀고정, 열 너비)는 오늘의 `parseXlsxToGrid`와 같은
   규칙으로 그대로 뽑아 `sheetMeta`/`styles`에 채운다(작으므로 그대로 메모리 보관).
3. 각 행이 스트리밍으로 도착할 때마다, 그 행을 `parseXlsxToGrid`가 만드는 것과 같은 모양
   (`Array<GridCell | 0>`)으로 JSON 직렬화해 `rowStorePath`(문서별 임시 파일, 위치는
   `os.tmpdir()`, 이름에 `docId` + `randomUUID()` 포함)에 이어붙이고, 그 시작 바이트 오프셋과
   길이를 `rowIndex`에 기록한 뒤 **그 행 자체는 버린다**(변수가 스코프를 벗어나 GC됨). 이 시점에
   메모리에 남는 건 "지금 처리 중인 딱 한 행"뿐이다.
4. `logicalToOriginalRow/Col`을 항등 배열(`[0, 1, 2, ...]`)로 초기화한다.

**실패 처리**: 오늘의 `seedRoom`과 동일하게 이 전체를 하나의 try/catch로 감싸 실패 시 로그만
남기고(룸은 canonical 없이 남는다 — 오늘의 `!room.canonicalGrid` 가드가 이미 이 상태를 안전하게
처리한다).

## 읽기 경로

`getCell(store, sheetIdx, logicalRow, logicalCol): Promise<GridCell | 0>`:

1. `key = gridCellKey(sheetIdx, logicalRow, logicalCol)`. `dirtyOverlay.get(key)` — 있으면
   반환(편집된 칸의 현재 값이 유일한 진실이므로 더 볼 것도 없다).
2. `originalRow = logicalToOriginalRow[sheetIdx][logicalRow]`,
   `originalCol = logicalToOriginalCol[sheetIdx][logicalCol]`. 둘 중 하나라도 `-1`(원본에 없던
   자리, 즉 협업 중 삽입됐지만 아직 아무도 안 건드린 칸)이면 `0`(빈 칸) 반환 — 디스크를 볼 필요가
   없다.
3. `readCache.get("${sheetIdx}:${originalRow}")` — 있으면 그 파싱된 행에서 `originalCol` 인덱싱해
   반환.
4. 없으면 `rowIndex[sheetIdx][originalRow]`로 `rowStorePath`의 해당 바이트 범위를 읽어
   (`fs.read`에 `position` 지정, 파일 핸들은 스토어당 한 번만 열어 재사용) `JSON.parse`, 캐시에
   넣고 `originalCol` 인덱싱해 반환. `rowIndex`에 그 자리가 없으면(원본에 그 행 자체가 없었음)
   `0` 반환.

`getRow`(페이지 전체를 채울 때 쓰는 배치 버전)도 같은 규칙으로 동작하되, 한 페이지(500행)를
한 번에 처리하며 청크마다 이벤트 루프에 양보한다(기존 `PAGE_SIZE_ROWS`/`yieldToEventLoop` 패턴
그대로).

**쓰기는 훨씬 단순해진다**: 편집은 항상 `dirtyOverlay.set(key, value)` 하나로 끝난다. 오늘
`ensureCanonicalCellSlot`이 하던 "이 좌표를 담을 수 있게 배열을 늘린다"는 로직 전체가
필요 없어진다 — `Map`은 미리 크기를 맞출 필요가 없다.

**시트 끝을 직접 늘릴 때(구조 연산 없이 rowCount/colCount만 커지는 경로,
`applyStructureChangeToCanonical`)는 예외적으로 배열을 건드려야 한다**: 이 경로는 삽입 연산이
아니라(그래서 `structuralOps`에 기록되지 않는다) `sheetMeta`의 rowCount/colCount만 올리는
것인데, `logicalToOriginalRow/Col`도 새 rowCount/colCount까지 `-1`로 채워 늘려야 한다 — 안 그러면
그 새 영역의 논리 좌표가 배열 밖이라 `getCell`이 범위를 벗어난 접근을 하게 된다. 늘리는 것만
하면 되고(줄어들 때는 `sheetMeta`의 rowCount/colCount만 낮추고 배열은 그대로 둔다 — 뒤쪽은
`getCell` 호출부가 `rowCount`/`colCount` 경계 체크로 이미 걸러낸다, 오늘의 canonical 그리드가
축소 시 배열을 자르지 않는 것과 같은 원칙).

## 수식 계산 — 2단계 읽기

`computeGridCellDisplay`(`src/xlsx/formulas.ts`)는 오늘 동기 함수로, 수식이 참조하는 임의의 칸을
그 자리에서 바로 인덱싱한다. `getCell`이 비동기가 되므로 그대로는 못 쓴다. 대신:

1. **1단계(prefetch)**: 수식 AST를 먼저 훑어 이 수식이 참조하는 모든 (시트, 행, 열) 좌표를
   모은다(이미 같은 파싱이 계산 자체에도 필요하므로 새 작업이 아니라 순서만 바꾸는 것). 이
   좌표들을 시트·원본행별로 묶어 `getRow`를 배치로 호출해 `readCache`를 미리 데운다.
2. **2단계(evaluate)**: 캐시가 데워진 상태에서, 오늘과 똑같은 동기 평가 로직을 그대로 돌린다
   (캐시 히트만 있으므로 실제로는 동기 함수를 그대로 재사용할 수 있다 — 캐시 조회 자체는
   동기다).

`recomputeFormulasForPage`(`room.ts`)는 이 두 단계를 감싸는 얇은 async 래퍼가 된다. 페이지 단위
청크+양보 구조는 그대로 유지.

## 행/열 삽입·삭제

`applyGridStructureOp`(`room.ts`)가 오늘 `insertCanonicalRows`/`deleteCanonicalRows`/
`insertCanonicalCols`/`deleteCanonicalCols`(`gridShift.ts`)로 행 데이터 배열 자체를 스플라이스
하던 것을, **`logicalToOriginalRow`/`logicalToOriginalCol` 배열만 스플라이스**하는 것으로
바꾼다:

- `insertRows(at, count)`: `logicalToOriginalRow[sheetIdx]`의 `at` 위치에 `-1`을 `count`개
  끼워 넣는다.
- `deleteRows(at, count)`: 그 위치의 `count`개 항목을 제거한다(가리키던 원본 행은 그냥
  참조되지 않게 될 뿐 — 온디스크 저장소에서 굳이 지우지 않는다. 룸이 닫히면 임시 파일째
  버려지므로 낭비가 크지 않다).
- 열도 동일하게 `logicalToOriginalCol`에 적용.

이 연산 자체가 O(1)에 가까워진다(정수 배열 스플라이스, 오늘 이미 실측된 저비용 그대로).
**`dirtyOverlay`의 키 이동은 오늘 만든 `shiftDirtyCellKeys`를 그대로 재사용**한다(`Set<string>`
대신 `Map<string, GridCell|0>`을 순회하도록 아주 작게 확장하면 된다 — 로직은 동일).

## 저장(export)과의 연결

`exportXlsxFromGrid`/`exportXlsxFromGridStreaming`은 오늘도 이미 "원본을 다시 로드 → 구조 연산
재생 → 편집된 칸만 덮어쓰기" 구조라 `LazyGridStore`와 자연스럽게 들어맞는다. 변경점은 인자를
`canonicalGrid: GridDocument` 대신 `store: LazyGridStore`로 바꾸고, 값 diff 루프가 `dirtyOverlay`
를 직접 순회하도록 바꾸는 것뿐이다(오히려 코드가 줄어든다 — 더 이상 "이 칸이 편집됐는지" 판단할
필요 없이 `dirtyOverlay`에 있으면 편집된 것이다).

## 룸 생명주기와 정리

- 룸이 완전히 닫힐 때(`closeConn`의 마지막 참여자 퇴장 경로, `forceCloseRoom`) `rowStorePath`
  임시 파일을 지운다(`fs.rm`, 기존 EBUSY 재시도 패턴 재사용).
- **서버 비정상 종료 대비**: 서버 시작 시(`ensureDataDirs` 근처) `os.tmpdir()`에서 이 프로젝트가
  만드는 패턴(`docview-lazygrid-*`)의 남은 임시 파일을 정리하는 스윕을 추가한다 — 정상 종료라면
  이미 다 지워져 있으므로 이 스윕은 사실상 비정상 종료를 복구하는 안전망이다.
- 디스크 공간: 이 임시 파일은 파일 크기에 비례해 **로컬 디스크**를 쓴다(전에는 메모리만 썼다) —
  운영상 새로 지켜봐야 할 자원이다.

## 에러 처리

- 온디스크 저장소 읽기 실패(파일이 세션 도중 지워지는 등)는 `getCell`의 Promise 거부로
  전파된다 — `recomputeFormulasForPage`의 기존 페이지별 catch-and-log, `flushToStore`의 기존
  저장 실패 처리(및 스트리밍 폴백)가 이미 이 클래스의 실패를 다룰 수 있는 구조라 새 메커니즘이
  필요 없다.

## 적용 순서(로드맵 — 구현 계획은 별도 문서)

1. `LazyGridStore`(수집 + 읽기 API)를 독립 모듈로 만들고 자체 테스트로 검증. 아직 `room.ts`에
   연결하지 않는다.
2. 룸의 읽기 경로(Yjs 페이지 시딩, 수식 계산)를 이 저장소를 쓰도록 전환.
3. `applyGridStructureOp`를 배열 스플라이스 대신 `logicalToOriginalRow/Col` 스플라이스로 전환.
4. `exportXlsxFromGrid`/`exportXlsxFromGridStreaming`이 `LazyGridStore`에서 직접 읽도록 전환.
5. `Room.canonicalGrid: GridDocument` 필드와 그 자리에서 쓰이던 `parseXlsxToGrid` 전체 파싱
   호출을 제거.

## 테스트 전략

- `LazyGridStore` 자체: 작은 픽스처 + 실제 대용량 파일(292,244행)로 수집 → 읽기 → 캐시 밀려남
  → 재읽기(같은 값 재확인)까지 검증.
- **오늘 세션에 이미 확립된 구조연산 회귀 스위트**(`tests/roomStructureSync.test.ts`,
  `tests/roomRowColShift.test.ts` — 중간 삽입·삭제, 여러 연산 연쇄, 1600행 시트, 구독 안 된
  페이지, 새 시트 추가 등)가 이 재설계 위에서도 전부 그대로 통과해야 한다 — 이게 사실상의
  회귀 방지 기준선이다.
- 수식 2단계 읽기: 현재 구독 페이지 밖의 셀을 참조하는 수식이 정확히 계산되는지.
- 캐시 상한이 실제로 오래된 행부터 밀어내는지, 밀려난 뒤 재조회해도 같은 값이 나오는지(안전성).
- 룸 종료 시 임시 파일 삭제, 서버 재시작 시 orphan 스윕 동작.

## 알려진 트레이드오프 / 리스크

- 대부분의 canonical 접근 지점(페이지 시딩, 수식 계산)이 동기에서 비동기로 바뀐다 — 호출부
  전체를 async로 전파해야 하는, 코드베이스 전반에 걸친 변경이다. 단계별 적용(위 로드맵)으로
  리스크를 분산한다.
- 디스크 공간이라는 새 자원 제약이 생긴다(메모리 대신).
- 룸이 사는 동안 로컬 디스크 I/O에 의존하게 된다 — 디스크가 매우 느리거나(네트워크 드라이브 등)
  가득 찬 환경에서는 오히려 오늘보다 나쁠 수 있다. 배포 환경이 로컬 SSD를 전제한다는 가정을
  문서화해 둔다.
