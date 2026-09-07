# 엑셀 동시편집기 진짜 윈도우 방식 지연 로딩 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 대형 xlsx를 동시편집기로 열어도 서버 메모리(Yjs)와 클라이언트 DOM 둘 다 "지금 보이는
부분"만큼만 유지되도록 만든다 — 파일 크기와 무관하게 빠르게 열리고 가볍게 유지.

**Architecture:** 서버가 항상 전체 그리드(canonical)를 메모리에 들고, 수식 계산도 서버가
전담한다. Yjs는 지금 구독 중인 500행 페이지만 담는 캐시로 축소되고, 편집은 발생 즉시
canonical에 반영되어 저장/다운로드가 항상 최신을 보장한다. 클라이언트는 가상 스크롤로 화면에
보이는 행만 렌더링하고, 스크롤에 따라 필요한 페이지를 서버에 구독 요청한다.

**Tech Stack:** TypeScript(Fastify + ws), Yjs(y-protocols, lib0), ExcelJS, fast-formula-parser,
Vitest.

**Spec:** [docs/superpowers/specs/2026-09-07-xlsx-editor-windowed-loading-design.md](../specs/2026-09-07-xlsx-editor-windowed-loading-design.md)

## Global Constraints

- canonical 그리드를 훑거나 채우는 모든 반복문(수식 재계산, 페이지 채움/비움)은
  `yieldToEventLoop`/`YIELD_EVERY_N_ROWS` 양보 지점 패턴을 반드시 재사용한다 — 대형 xlsx가
  서버 전체를 멈추는 버그(`6355a17`로 수정)가 새 반복문에서 재발하면 안 된다.
- Yjs `GRID_CELLS_MAP`(`cells`)에는 항상 원본 텍스트(수식이면 수식 자체)만 쓴다. 계산된
  표시값은 반드시 새 `GRID_COMPUTED_MAP`(`'computed'`)에만 쓴다.
- 셀 편집이 Yjs에 반영되면 같은 이벤트 처리 안에서 canonical 반영까지 끝낸다 — 배치·지연 반영
  금지(편집 직후 다운로드가 항상 최신이어야 한다는 요구사항).
- `GRID_STRUCTURE_MAP`/`GRID_MERGES_MAP`은 절대 윈도잉하지 않는다 — 항상 모든 시트의 전체
  메타데이터가 Yjs에 있다.
- 페이지 크기는 정확히 500행으로 고정한다(0-based 페이지 인덱스: 페이지 N = 행 `[N*500,
  (N+1)*500)`). 바로보기 뷰어의 `/api/view/grid/rows`(유연한 from/count 쿼리)와는 별개의,
  고정 페이지 단위 체계다 — 혼동하지 않는다.
- 페이지 구독이 합집합에서 빠져도 10초간 아무도 재구독하지 않을 때만 실제로 Yjs에서 지운다
  (스크롤 경계 스래싱 방지).
- room.ts의 xlsx 동시편집 테스트는 반드시 이 세션에서 이미 확립된 패턴을 그대로 재사용한다:
  `FakeSocket`(EventEmitter 기반, `mirror: Y.Doc`으로 서버가 보낸 걸 디코드) +
  `joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false)` 직접 호출 +
  `vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] })`
  (setImmediate는 반드시 real로 남겨야 한다 — ExcelJS의 zip/stream 내부 동작이 진짜
  setImmediate 턴에 의존하기 때문에, 전부 fake로 두면 `tests/roomGridIndex.test.ts`에서 실제로
  겪었듯 무한 대기한다) + `flushMicrotasks`(수백~수천 회 `await Promise.resolve()`와
  `await new Promise(r => setImmediate(r))`를 함께 반복) 헬퍼.
- 찾기/바꾸기, 채우기 핸들 드래그, 붙여넣기가 현재 로드된(구독된) 페이지 범위를 벗어나는 경우는
  이번 계획의 범위 밖이다 — 이 계획에서 다루지 않되, 찾기/바꾸기는 "지금 로드된 범위에서만
  검색합니다" 경고를 UI에 추가해 조용히 틀린 결과를 주지 않게 한다(Task 9).

---

### Task 1: 공유 상수 정리 — YIELD_EVERY_N_ROWS 중복 제거, GRID_COMPUTED_MAP 추가

**Files:**
- Modify: `src/spreadsheet.ts:14` (`YIELD_EVERY_N_ROWS`를 export)
- Modify: `src/xlsx/collab.ts:23` (로컬 중복 상수 제거, `spreadsheet.ts`에서 import)
- Modify: `src/collab/gridShared.ts` (`GRID_COMPUTED_MAP` 추가)
- Test: `tests/spreadsheet-derived.test.ts`(기존 파일에 케이스 추가), `tests/gridShared.test.ts`(신규)

**Interfaces:**
- Produces: `export const YIELD_EVERY_N_ROWS = 2000;`(`src/spreadsheet.ts`에서), 이후 모든
  파일이 이 하나만 참조. `export const GRID_COMPUTED_MAP = 'computed';`(`src/collab/gridShared.ts`).

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/gridShared.test.ts`(신규 파일):
```ts
import { describe, expect, test } from 'vitest';
import { GRID_CELLS_MAP, GRID_COMPUTED_MAP, GRID_STYLES_MAP } from '../src/collab/gridShared.js';

describe('gridShared: GRID_COMPUTED_MAP', () => {
  test('is a distinct map name from cells/styles', () => {
    expect(GRID_COMPUTED_MAP).toBe('computed');
    expect(GRID_COMPUTED_MAP).not.toBe(GRID_CELLS_MAP);
    expect(GRID_COMPUTED_MAP).not.toBe(GRID_STYLES_MAP);
  });
});
```

`tests/spreadsheet-derived.test.ts`에 추가(기존 `describe` 블록 밖, 파일 끝에 새 블록):
```ts
import { YIELD_EVERY_N_ROWS } from '../src/spreadsheet.js';

describe('spreadsheet.ts: YIELD_EVERY_N_ROWS is exported', () => {
  test('is a positive integer', () => {
    expect(Number.isInteger(YIELD_EVERY_N_ROWS)).toBe(true);
    expect(YIELD_EVERY_N_ROWS).toBeGreaterThan(0);
  });
});
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/gridShared.test.ts tests/spreadsheet-derived.test.ts`
Expected: FAIL — `GRID_COMPUTED_MAP`/`YIELD_EVERY_N_ROWS`가 아직 export되지 않아 import 오류.

- [ ] **Step 3: 소스 수정**

`src/spreadsheet.ts:14`의
```ts
const YIELD_EVERY_N_ROWS = 2000;
```
을
```ts
export const YIELD_EVERY_N_ROWS = 2000;
```
로 바꾼다(값은 그대로, `export`만 추가).

`src/collab/gridShared.ts`의 `export const GRID_CELLS_MAP = 'cells';` 바로 뒤에 추가:
```ts
/** 수식 셀의 계산된 표시값만 담는 Y.Map — "시트:행:열" 키(gridCellKey와 동일 형식). 수식이
 * 아닌 셀은 이 맵에 항목이 없다. cells는 항상 원본 텍스트(수식이면 수식 자체)만 담고, 여기에
 * 계산된 값을 절대 섞지 않는다 — 섞으면 저장 시 수식이 정적 값으로 깨진다. */
export const GRID_COMPUTED_MAP = 'computed';
```

`src/xlsx/collab.ts:23`의
```ts
const YIELD_EVERY_N_ROWS = 2000;
```
줄을 삭제하고, 파일 상단 import에 `yieldToEventLoop`와 나란히 `YIELD_EVERY_N_ROWS`를 추가:
```ts
import { parseXlsxToGrid, saveGridDerived, yieldToEventLoop, YIELD_EVERY_N_ROWS } from '../spreadsheet.js';
```
(정확한 기존 import 목록은 파일을 직접 열어 확인하고 병합할 것 — `parseXlsxToGrid`,
`saveGridDerived`, `yieldToEventLoop`가 이미 import되어 있으므로 그 목록에 `YIELD_EVERY_N_ROWS`만
추가한다.)

- [ ] **Step 4: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/gridShared.test.ts tests/spreadsheet-derived.test.ts tests/xlsx-collab.test.ts tests/xlsxYielding.test.ts`
Expected: PASS 전부 — 특히 `xlsx-collab.test.ts`/`xlsxYielding.test.ts`가 여전히 통과해야
`xlsx/collab.ts`의 상수 교체가 동작을 안 바꿨음을 증명한다(값이 똑같이 2000이므로 당연히
통과해야 하지만, import 경로 실수를 여기서 잡는다).

- [ ] **Step 5: 타입 체크**

Run: `npx tsc --noEmit`
Expected: PASS

- [ ] **Step 6: 커밋**

```bash
git add src/spreadsheet.ts src/xlsx/collab.ts src/collab/gridShared.ts tests/gridShared.test.ts tests/spreadsheet-derived.test.ts
git commit -m "refactor: YIELD_EVERY_N_ROWS 중복 제거, GRID_COMPUTED_MAP 추가"
```

---

### Task 2: 서버 — Room에 canonical 그리드 유지 + 편집 즉시 반영 (과도기: Yjs 전체 시딩은 그대로 유지)

이 태스크는 아직 아무 사용자 동작도 바꾸지 않는다 — Yjs는 여전히 전체를 시딩한다. 다만 이제
Room이 canonical 그리드도 "같이" 들고, 모든 xlsx 편집이 그 canonical에도 즉시 반영되기
시작한다. 이렇게 과도기를 두는 이유: 이 시점엔 canonical과 Yjs 내용이 항상 동일해야 하므로,
"canonical이 정확히 최신 상태를 반영하는가"를 기존 테스트 데이터로 안전하게 검증할 수 있다
(윈도잉을 동시에 켜면 canonical/Yjs가 서로 다른 부분집합이 되어 버그를 구분하기 어려워진다).

**Files:**
- Modify: `src/collab/room.ts:70-97`(`Room` 인터페이스), `:119-136`(`seedRoom`),
  `:331-380`(`getOrCreateRoom`, 특히 `ydoc.on('update', ...)` 핸들러)
- Test: `tests/roomCanonicalGrid.test.ts`(신규)

**Interfaces:**
- Consumes: `parseXlsxToGrid`(`src/spreadsheet.ts`, 기존), `GridDocument`/`GridSheet`/`GridCell`
  타입(`src/spreadsheet.ts`, 기존), `GRID_CELLS_MAP`/`gridCellKey`/`parseGridCellKey`
  (`src/collab/gridShared.ts`, 기존).
- Produces: `Room.canonicalGrid: GridDocument | null`(xlsx가 아닌 룸은 항상 `null`) —
  Task 3(저장/다운로드), Task 6(윈도잉 컷오버)가 이 필드를 읽는다. 내부 헬퍼
  `applyCellEditToCanonical(canonicalGrid: GridDocument, sheetIdx: number, row: number, col: number, rawValue: string): void`
  — Task 6에서도 재사용 가능하도록 이름을 고정해 둔다(exported하지 않아도 되지만, 같은
  이름으로 정의해 다음 태스크 작성자가 grep으로 찾을 수 있게 한다).

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/roomCanonicalGrid.test.ts`(신규 파일 — Global Constraints의 FakeSocket 패턴을 그대로
따른다):

```ts
import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeUpdate } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

const MESSAGE_SYNC = 0;

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  readonly mirror = new Y.Doc();

  send(data: Uint8Array, cb?: (err?: Error) => void): void {
    const decoder = decoding.createDecoder(data);
    const messageType = decoding.readVarUint(decoder);
    if (messageType === MESSAGE_SYNC) {
      const syncType = decoding.readVarUint(decoder);
      if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) {
        readUpdate(decoder, this.mirror, this);
      }
    }
    cb?.();
  }
  close(): void {}
  ping(): void {}
}

async function makeSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 'Name';
  sheet.getCell('A2').value = 'Alice';
  sheet.getCell('B2').value = 10;
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
}

function sendCellEdit(socket: FakeSocket, key: string, value: string): void {
  const scratch = new Y.Doc();
  scratch.getMap<string>('cells').set(key, value);
  const update = Y.encodeStateAsUpdate(scratch);
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeUpdate(encoder, update);
  socket.emit('message', encoding.toUint8Array(encoder));
}

describe('room.ts: xlsx room keeps a canonical grid in sync with edits', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-canonical-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.useRealTimers();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  test('canonical grid matches the original file immediately after seeding, before any edit', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomCanonicalGridForTest } = await import('../src/collab/room.js');

    const docId = randomUUID();
    const original = await makeSampleXlsx();
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId,
      name: 'x.xlsx',
      ext: 'xlsx',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
      progress: 100,
      viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);

    const canonical = getRoomCanonicalGridForTest(docId);
    expect(canonical).not.toBeNull();
    expect(canonical!.sheets[0].rows[0][0]).toEqual({ v: 'Name' });
    expect(canonical!.sheets[0].rows[1][0]).toEqual({ v: 'Alice' });
  });

  test('an edit is reflected in the canonical grid immediately, without waiting for autosave', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomCanonicalGridForTest } = await import('../src/collab/room.js');
    const { gridCellKey } = await import('../src/collab/gridShared.js');

    const docId = randomUUID();
    const original = await makeSampleXlsx();
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId,
      name: 'x.xlsx',
      ext: 'xlsx',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
      progress: 100,
      viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);

    // A2("Alice")를 "Alice Kim"으로 편집 — 자동저장(8초)을 전혀 기다리지 않는다.
    sendCellEdit(socket, gridCellKey(0, 1, 0), 'Alice Kim');

    const canonical = getRoomCanonicalGridForTest(docId);
    expect(canonical!.sheets[0].rows[1][0]).toEqual({ v: 'Alice Kim' });
  });
});
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/roomCanonicalGrid.test.ts`
Expected: FAIL — `getRoomCanonicalGridForTest`가 아직 export되지 않음.

- [ ] **Step 3: `src/collab/room.ts` 수정**

`Room` 인터페이스(`:70-97`)에 필드 추가(마지막 `suppressAutosaveReschedule: boolean;` 다음 줄):
```ts
  /** xlsx 룸에서만 쓰인다 — 항상 전체 시트 데이터를 담는 "진짜" 데이터. Yjs cells 맵은 이제
   * (Task 6 이후) 지금 구독 중인 페이지만 담는 캐시가 되므로, 저장·다운로드·수식 계산은 항상
   * 이 필드를 기준으로 한다. docx/pptx 룸에서는 항상 null. */
  canonicalGrid: GridDocument | null;
```
`import`에 `type { GridDocument }`를 `../spreadsheet.js`에서 추가(이미 `parseXlsxToGrid`를
import하고 있으므로 같은 줄에 병합).

`seedRoom`(`:119-136`)의 xlsx 분기를 수정 — 파싱한 `grid`를 룸에 보관해야 하므로, `seedRoom`이
`room: Room` 전체를 받도록 시그니처를 바꾼다(기존 `docId, kind, ydoc` 세 인자 대신):
```ts
async function seedRoom(room: Room): Promise<void> {
  const { docId, kind, ydoc } = room;
  try {
    const meta = await store.readMeta(docId);
    if (!meta) return;
    const original = await store.readOriginal(docId, meta.ext);
    if (kind === 'xlsx') {
      const grid = await parseXlsxToGrid(original);
      room.canonicalGrid = grid;
      await seedGridFromXlsx(ydoc, grid);
    } else if (kind === 'pptx') {
      seedSlidesFromPptx(original, ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY));
    } else {
      const { doc } = importDocx(original);
      prosemirrorToYXmlFragment(doc, ydoc.getXmlFragment(XML_FRAGMENT_NAME));
    }
  } catch (error) {
    console.error(`동시편집 세션 초기화 실패 (문서 ${docId}):`, error);
  }
}
```
이 시그니처 변경 때문에 `getOrCreateRoom`(`:331-380`) 안의 `ready: seedRoom(docId, kind, ydoc)`
호출부도 고쳐야 한다 — 다만 `room` 객체는 그 줄이 실행되는 시점엔 아직 완성 전(생성자 리터럴
안)이므로, `seedRoom` 호출을 `Room` 객체 리터럴 밖으로 옮겨야 한다. `getOrCreateRoom`을 아래처럼
바꾼다(기존 로직 순서는 최대한 유지, `room` 선언 후 `seedRoom(room)` 호출로 `ready`를 별도
할당):

```ts
function getOrCreateRoom(docId: string, ext: string): Room {
  const existing = rooms.get(docId);
  if (existing) return existing;

  const kind = kindOfExt(ext);
  const ydoc = new Y.Doc({ gc: true });
  const awareness = new Awareness(ydoc);
  awareness.setLocalState(null);
  const conns = new Map<WebSocket, Set<number>>();

  const room: Room = {
    docId,
    kind,
    ydoc,
    awareness,
    conns,
    readOnlyConns: new Set<WebSocket>(),
    autosaveTimer: null,
    ready: Promise.resolve(), // 아래에서 즉시 덮어씀 — seedRoom(room)이 room 자신을 참조해야 해서
    suppressAutosaveReschedule: false,
    canonicalGrid: null,
  };
  room.ready = seedRoom(room);
  rooms.set(docId, room);

  ydoc.on('update', (update: Uint8Array) => {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeUpdate(encoder, update);
    broadcast(room, encoding.toUint8Array(encoder));
    if (room.kind === 'xlsx' && room.canonicalGrid) {
      applyYjsUpdateToCanonical(room.canonicalGrid, ydoc, update);
    }
    if (!room.suppressAutosaveReschedule) scheduleAutosave(room);
  });

  // ... 이하 awareness.on(...) 부분은 기존과 동일, 손대지 않는다
```

**주의**: `ready: Promise.resolve()`로 임시 채운 뒤 바로 `room.ready = seedRoom(room)`로
덮어쓰는 순서가 중요하다 — `Room`이 `readonly ready: Promise<void>`로 선언되어 있으면(기존
인터페이스 확인 필요) 이 재할당이 컴파일 에러가 난다. `Room` 인터페이스의 `ready` 필드 선언에서
`readonly`를 제거해야 한다(`:84` 근처, `readonly ready: Promise<void>` → `ready: Promise<void>`)
— 이 변경도 이 스텝에 포함시킨다.

이제 `applyYjsUpdateToCanonical` 헬퍼를 추가한다 — `seedRoom` 함수 앞(파일에서 `seedRoom`보다
먼저 정의되도록):
```ts
/**
 * Yjs 'update' 이벤트 하나가 xlsx 룸의 GRID_CELLS_MAP에 담은 셀 편집을 canonical 그리드에도
 * 즉시 반영한다. Y.Doc.on('update')는 원시 바이너리 업데이트만 주므로, 이 업데이트를 별도
 * "관찰용" Y.Doc에 적용해 어떤 키가 바뀌었는지 관찰한 뒤 그 값을 canonical에 옮겨 쓴다 — 이미
 * ydoc 자체에 적용된 값을 다시 읽으면 되므로 실제로는 관찰용 Y.Doc 없이 ydoc.getMap(...)에서
 * 직접 최신 값을 읽으면 된다(update 바이트 자체를 파싱할 필요 없음): update가 어떤 키를
 * 건드렸는지는 몰라도, ydoc의 GRID_CELLS_MAP은 이미 최신 상태이므로 전체를 다시 훑어
 * canonical과 다른 값만 반영하면 정확하다 — 다만 매 update마다 전체를 훑으면 페이지 하나(500행)
 * 안에서도 비효율적이므로, Task 6(윈도잉 컷오버) 전까지는 Yjs가 전체를 담고 있어 이 함수가
 * "전체를 훑는" 접근이어도 안전하다(과도기 단계). Task 6에서 "지금 구독 중인 페이지만 훑는"
 * 최적화된 버전으로 교체될 예정이다.
 */
function applyYjsUpdateToCanonical(canonicalGrid: GridDocument, ydoc: Y.Doc, _update: Uint8Array): void {
  const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
  for (const [key, value] of cells) {
    const parsed = parseGridCellKey(key);
    if (!parsed) continue;
    const { sheetIdx, row, col } = parsed;
    const sheet = canonicalGrid.sheets[sheetIdx];
    if (!sheet || row >= sheet.rows.length || col >= sheet.colCount) continue;
    const currentCell = sheet.rows[row][col];
    const currentText = currentCell === 0 ? '' : currentCell.v;
    if (currentText === value) continue;
    (sheet.rows[row] as Array<{ v: string; s?: number } | 0>)[col] =
      value === '' ? 0 : { v: value };
  }
}
```
`GRID_CELLS_MAP`/`parseGridCellKey`를 `../collab/gridShared.js` import 목록에 추가(이미
`DOC_META_MAP` 등을 이 파일에서 import하고 있으므로 같은 줄에 병합).

**테스트 전용 export** — 파일 끝(`activeRoomCount` export 근처)에 추가:
```ts
/** 테스트 전용 — 룸의 canonical 그리드를 직접 들여다본다. 프로덕션 코드에서는 쓰지 않는다. */
export function getRoomCanonicalGridForTest(docId: string): GridDocument | null {
  return rooms.get(docId)?.canonicalGrid ?? null;
}
```

- [ ] **Step 4: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/roomCanonicalGrid.test.ts`
Expected: PASS (2 tests)

- [ ] **Step 5: 회귀 확인**

Run: `npx vitest run tests/roomAutosaveLoop.test.ts tests/roomGridIndex.test.ts tests/xlsx-collab.test.ts && npx tsc --noEmit`
Expected: 모두 PASS — `seedRoom`/`getOrCreateRoom` 시그니처 변경이 기존 docx/pptx 흐름을
깨지 않았는지 확인.

- [ ] **Step 6: 커밋**

```bash
git add src/collab/room.ts tests/roomCanonicalGrid.test.ts
git commit -m "feat: xlsx 룸이 canonical 그리드를 유지하고 편집을 즉시 반영하도록 함(과도기)"
```

---

### Task 3: 서버 — 저장/다운로드를 canonical 기준으로 전환

이 시점엔 Task 2 덕에 canonical이 항상 Yjs와 동일한 내용을 담고 있으므로(아직 윈도잉 전), 이
전환은 **동작을 바꾸지 않으면서** 내보내기 로직의 데이터 소스만 Yjs에서 canonical로 옮기는
안전한 리팩터링이다. 기존 `tests/xlsx-collab.test.ts`가 여전히 통과해야 회귀가 없음을 증명한다.

**Files:**
- Modify: `src/xlsx/collab.ts`(`exportXlsxFromGrid` 시그니처와 본문)
- Modify: `src/collab/room.ts`(`flushToStore`의 xlsx 분기, `getCurrentBytes`의 xlsx 분기)
- Modify: `tests/xlsx-collab.test.ts`(시그니처 변경에 맞춰 호출부 업데이트)
- Test: 위 수정된 `tests/xlsx-collab.test.ts` 자체가 검증

**Interfaces:**
- Consumes: Task 2의 `Room.canonicalGrid: GridDocument | null`.
- Produces: `exportXlsxFromGrid(originalXlsxBytes: Uint8Array, canonicalGrid: GridDocument, baselineGrid: GridDocument): Promise<Uint8Array>`
  — 시그니처가 `(bytes, ydoc)`에서 `(bytes, canonicalGrid, baselineGrid)`로 바뀐다.
  `baselineGrid`는 "원본 그대로"를 뜻하며, Task 2에서 룸이 `parseXlsxToGrid`로 처음 만든 grid를
  깊은 복사해 별도로 보관해야 한다(canonical은 편집되며 바뀌지만 baseline은 안 바뀜 — 지금
  Yjs의 `baseline`/`stylesBaseline` 맵과 같은 역할). Task 2의 결과물에 `Room.baselineGrid:
  GridDocument | null` 필드를 이 태스크에서 추가한다(Task 2가 놓친 부분이 아니라, "베이스라인
  분리 보관"이 이 태스크의 스코프이기 때문에 여기서 추가).

- [ ] **Step 1: 실패하는 테스트 작성 (기존 테스트를 새 시그니처에 맞게 고쳐 쓴다)**

`tests/xlsx-collab.test.ts`는 지금 `seedGridFromXlsx(ydoc, grid)` 후 `ydoc`의 Y.Map을 직접
조작해 편집을 흉내내고 `exportXlsxFromGrid(original, ydoc)`을 호출한다. 새 시그니처에 맞춰
각 테스트를 아래 패턴으로 고친다 — 예를 들어 "seeds a Y.Doc from xlsx, edits a cell, and
flushes back..." 테스트를:

```ts
test('seeds a canonical grid from xlsx, edits a cell, and flushes back to a valid xlsx while preserving formulas', async () => {
  const original = await makeSampleXlsx();
  const baselineGrid = await parseXlsxToGrid(original);
  const canonicalGrid = structuredClone(baselineGrid) as GridDocument;

  // "다른 브라우저 탭"이 편집한 것처럼 canonical을 직접 바꾼다(room.ts의
  // applyYjsUpdateToCanonical이 하는 것과 동일한 형태의 직접 대입).
  (canonicalGrid.sheets[0].rows[1] as Array<GridCell | 0>)[0] = { v: 'Alice Kim' };

  const exported = await exportXlsxFromGrid(original, canonicalGrid, baselineGrid);
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(exported as unknown as ArrayBuffer);
  const sheet = reopened.getWorksheet('Sheet1')!;

  expect(sheet.getCell('A2').value).toBe('Alice Kim');
  expect(sheet.getCell('B2').value).toBe(10); // 손대지 않은 숫자 셀
  const formulaCell = sheet.getCell('B4').value as { formula?: string } | number | null;
  expect(typeof formulaCell === 'object' && formulaCell !== null ? formulaCell.formula : undefined).toBe('SUM(B2:B3)');
});
```

**나머지 5개 테스트도 같은 패턴으로 고친다**: `seedGridFromXlsx(ydoc, grid)` +
`ydoc.getMap(...).set(...)` 조합을 전부 `parseXlsxToGrid` → `structuredClone`으로 만든
`canonicalGrid` 직접 수정으로 바꾸고, `exportXlsxFromGrid(original, ydoc)` 호출을
`exportXlsxFromGrid(original, canonicalGrid, baselineGrid)`로 바꾼다. "두 독립적인 편집이 함께
살아남는다(CRDT 병합)" 테스트는 canonical 모델에서는 의미가 달라진다 — CRDT 병합은 이제 Yjs
레벨(Task 6 이전까지는 여전히 전체 시딩)에서 일어나고 canonical은 그 병합 결과가 반영된
"단일 진실"이므로, 이 테스트는 "두 클라이언트의 Y.Doc을 병합한 뒤 그 병합 결과를 canonical에
반영하고 export하면 둘 다 반영된다"로 재구성한다(정확한 코드는 구현자가 기존 테스트 의도를
보존하며 작성 — CRDT 병합 자체는 Yjs가 하므로, `Y.applyUpdate`로 두 Y.Doc을 병합한 뒤 그
결과로 `canonicalGrid`를 만들면 된다).

`GridCell`/`GridDocument` 타입을 이 테스트 파일 상단 import에 추가:
```ts
import type { GridCell, GridDocument } from '../src/spreadsheet.js';
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/xlsx-collab.test.ts`
Expected: FAIL — `exportXlsxFromGrid`가 아직 옛 시그니처(2개 인자)라 타입 오류.

- [ ] **Step 3: `src/xlsx/collab.ts` 수정**

`exportXlsxFromGrid`의 시그니처와, Yjs 맵을 읽던 부분을 `canonicalGrid`/`baselineGrid` 순회로
바꾼다:
```ts
export async function exportXlsxFromGrid(
  originalXlsxBytes: Uint8Array,
  canonicalGrid: GridDocument,
  baselineGrid: GridDocument,
): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(originalXlsxBytes as unknown as ArrayBuffer);
  const visibleSheets = workbook.worksheets.filter((sheet) => !sheet.state || sheet.state === 'visible');
  const originalSheetCount = visibleSheets.length;

  const wasProtected = visibleSheets.map((sheet) =>
    Boolean((sheet as unknown as { sheetProtection?: unknown }).sheetProtection),
  );

  // 시트 구조(추가/삭제/이름변경) 처리는 canonicalGrid.sheets vs baselineGrid.sheets 비교로
  // 기존 structure map 순회 로직을 그대로 이식한다(정확한 로직은 현재 파일의 구조 처리 부분을
  // 그대로 옮기되, `structure.get(key)` 형태의 Yjs 조회를 `canonicalGrid.sheets[i]` 배열 인덱싱
  // 으로 바꾼다 — 이 부분은 코드량이 많으므로 원본 함수 본문을 참조해 옮길 것, 새로 설계하지 않는다).
  // ...(기존 구조 처리 로직 이식)...

  for (let sheetIdx = 0; sheetIdx < canonicalGrid.sheets.length; sheetIdx += 1) {
    const currentSheet = canonicalGrid.sheets[sheetIdx];
    const baselineSheet = baselineGrid.sheets[sheetIdx];
    const worksheet = /* 이 시트에 대응하는 ExcelJS worksheet, 기존 로직에서 찾던 방식 재사용 */;
    for (let r = 0; r < currentSheet.rows.length; r += 1) {
      const currentRow = currentSheet.rows[r];
      const baselineRow = baselineSheet?.rows[r];
      for (let c = 0; c < currentRow.length; c += 1) {
        const currentCell = currentRow[c];
        const baselineCell = baselineRow?.[c];
        const currentText = currentCell === 0 ? '' : currentCell.v;
        const baselineText = baselineCell === 0 || baselineCell === undefined ? '' : baselineCell.v;
        if (currentText === baselineText) continue; // 손대지 않은 셀 — 원본 그대로 둔다(수식 보존)
        // 값이 실제로 바뀐 셀만 기존 coerceCellValue 로직으로 써넣는다(기존 로직 재사용).
      }
    }
  }

  // 서식(styles/stylesBaseline) 비교도 같은 패턴으로 canonicalGrid의 각 셀 .s(스타일 참조
  // 인덱스)를 baselineGrid와 비교해 바뀐 것만 적용한다.

  return workbook.xlsx.writeBuffer() as unknown as Promise<Uint8Array>;
}
```

**중요**: 이 함수는 코드량이 많아(기존 파일에서 원본 대비 200줄 이상) 여기 완전한 코드를 다
쓰지 않는다 — 구현자는 반드시 **원본 `exportXlsxFromGrid`(git으로 현재 커밋 상태를 확인)를
열어**, Yjs 맵 조회(`structure.get(key)`, `cells.get(key)`, `baseline.get(key)`,
`styles.get(key)`, `stylesBaseline.get(key)`) 하나하나를 위 패턴(canonicalGrid/baselineGrid
배열 인덱싱)으로 1:1 대응시켜 옮긴다 — 로직 자체(구조 변경 감지, 값 diff, 서식 diff, 행/열
축소 처리)는 그대로 보존하고 **데이터 소스만** 바꾼다. 옮긴 뒤 Step 4의 테스트가 전부
통과해야 정확히 옮겨졌다는 뜻이다.

`src/collab/room.ts`의 `flushToStore` xlsx 분기(`:151-163` 근처)를 수정:
```ts
  if (kind === 'xlsx') {
    const original = await store.readOriginal(docId, meta.ext);
    if (!room.canonicalGrid || !room.baselineGrid) {
      console.error(`동시편집 저장 실패 — canonical 그리드 없음 (문서 ${docId})`);
    } else {
      const newXlsxBytes = await exportXlsxFromGrid(original, room.canonicalGrid, room.baselineGrid);
      await store.saveOriginal(docId, meta.ext, Buffer.from(newXlsxBytes));
      try {
        const grid = await reparseGridForViewer(newXlsxBytes);
        await saveGridDerived(store, docId, grid);
        await store.writeMeta({ ...meta, pageCount: grid.sheets.length, status: 'ready', viewKind: 'grid', progress: 100 });
      } catch (error) {
        console.error(`동시편집 저장 후 그리드 재생성 실패 (문서 ${docId}):`, error);
      }
    }
  } else if (kind === 'pptx') {
```

`getCurrentBytes`의 xlsx 분기(`:304-306` 근처)도 동일하게 수정:
```ts
  if (room.kind === 'xlsx') {
    if (!room.canonicalGrid || !room.baselineGrid) return null;
    return exportXlsxFromGrid(original, room.canonicalGrid, room.baselineGrid);
  }
```

`Room` 인터페이스에 `baselineGrid: GridDocument | null;` 필드를 `canonicalGrid` 옆에 추가하고,
`seedRoom`의 xlsx 분기에서 `room.canonicalGrid = grid;` 다음 줄에
`room.baselineGrid = structuredClone(grid);`를 추가한다(Node 17+ 내장 `structuredClone` 사용
— 이 프로젝트가 이미 Node 최신 버전을 쓰고 있는지 `package.json`의 `engines` 필드로 확인하고,
없으면 `JSON.parse(JSON.stringify(grid))`로 대체).

- [ ] **Step 4: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/xlsx-collab.test.ts`
Expected: PASS (전체 6개 테스트, 새 시그니처로).

- [ ] **Step 5: 회귀 확인**

Run: `npx vitest run tests/roomCanonicalGrid.test.ts tests/roomAutosaveLoop.test.ts tests/roomGridIndex.test.ts && npx tsc --noEmit`
Expected: 모두 PASS.

- [ ] **Step 6: 커밋**

```bash
git add src/xlsx/collab.ts src/collab/room.ts tests/xlsx-collab.test.ts
git commit -m "refactor: xlsx 저장/다운로드가 Yjs 대신 canonical 그리드에서 직렬화하도록 전환"
```

---

### Task 4: 서버 — 수식 계산 엔진을 서버로 이식

**Files:**
- Create: `src/xlsx/formulas.ts` (수식 파서 설정 + 계산 함수, `src/client/formulaOverrides.ts`의
  로직을 거의 그대로 이식)
- Modify: `src/collab/room.ts` (계산 결과를 `GRID_COMPUTED_MAP`에 반영하는 함수 추가 — 트리거
  연결은 Task 5에서)
- Delete 예정 표시만: `src/client/grid-editor.ts`의 수식 계산 부분은 Task 7에서 제거(이
  태스크에서는 아직 손대지 않는다 — 서버 쪽이 먼저 준비되어야 클라이언트를 안전하게 뗄 수 있다).
- Test: `tests/xlsxFormulas.test.ts`(신규)

**Interfaces:**
- Consumes: `src/client/formulaOverrides.ts`의 `formulaFunctions`/`formulaFunctionsNeedContext`
  (이식 대상 — 새 파일이 기존 파일을 import해도 되고, 완전히 옮겨도 된다. 클라이언트 번들에서
  빼야 하므로(Task 7) **완전히 옮기고 원본은 Task 7에서 삭제**하는 쪽을 권장 — 이 태스크에서는
  일단 새 서버 파일에 복사해 넣고, 기존 클라이언트 파일은 그대로 둔다(둘 다 존재해도 당장
  문제없음, 정리는 Task 7).
- Produces: `computeGridCellDisplay(grid: GridDocument, sheetIdx: number, row: number, col: number): unknown`
  — 셀 하나의 계산된 표시값(수식이 아니면 원본 그대로, 수식이면 계산 결과 또는 에러 이름)을
  돌려준다. `formatComputedValueForStorage(value: unknown): string` — 계산 결과를
  `GRID_COMPUTED_MAP`에 저장할 문자열로 변환(클라이언트의 기존 `formatComputedValue`+
  `formatNumberForDisplay` 로직 중 "값 자체"를 문자열화하는 부분만 — 서식 적용은 클라이언트가
  스타일 정보로 계속 처리하므로 서버는 순수 계산값만 문자열화한다).

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/xlsxFormulas.test.ts`(신규 파일):
```ts
import { describe, expect, test } from 'vitest';
import { computeGridCellDisplay, formatComputedValueForStorage } from '../src/xlsx/formulas.js';
import type { GridDocument } from '../src/spreadsheet.js';

function makeGrid(rows: Array<Array<{ v: string } | 0>>): GridDocument {
  return {
    styles: [],
    sheets: [
      {
        name: 'Sheet1',
        rowCount: rows.length,
        colCount: Math.max(...rows.map((r) => r.length)),
        colWidths: [],
        merges: [],
        freezeRows: 0,
        freezeCols: 0,
        rows,
      },
    ],
  };
}

describe('xlsx/formulas.ts: computeGridCellDisplay', () => {
  test('a plain value cell returns its own text (numbers coerced to number)', () => {
    const grid = makeGrid([[{ v: '10' }, { v: 'hello' }]]);
    expect(computeGridCellDisplay(grid, 0, 0, 0)).toBe(10);
    expect(computeGridCellDisplay(grid, 0, 0, 1)).toBe('hello');
  });

  test('SUM formula referencing cells on the same visible page computes correctly', () => {
    const grid = makeGrid([
      [{ v: '10' }],
      [{ v: '20' }],
      [{ v: '=SUM(A1:A2)' }],
    ]);
    expect(computeGridCellDisplay(grid, 0, 2, 0)).toBe(30);
  });

  test('a formula referencing a row far outside any 500-row page still computes correctly', () => {
    // 화면 밖(다른 페이지) 참조가 서버 canonical 기준으로는 항상 정확해야 한다는 게 이 설계의
    // 핵심 — 5000번째 행(10번째 페이지)을 0번째 행에서 참조.
    const rows: Array<Array<{ v: string } | 0>> = Array.from({ length: 5001 }, () => [0]);
    rows[4999] = [{ v: '42' }]; // 0-based 행 4999 = 5000번째 행
    rows[0] = [{ v: '=A5000' }];
    const grid = makeGrid(rows);
    expect(computeGridCellDisplay(grid, 0, 0, 0)).toBe(42);
  });

  test('circular reference returns the REF error name, not an infinite loop', () => {
    const grid = makeGrid([[{ v: '=A1' }]]);
    expect(computeGridCellDisplay(grid, 0, 0, 0)).toBe('#REF!');
  });

  test('MAX (a fast-formula-parser stub overridden in formulaOverrides) works correctly', () => {
    const grid = makeGrid([[{ v: '3' }, { v: '7' }, { v: '=MAX(A1:B1)' }]]);
    expect(computeGridCellDisplay(grid, 0, 0, 2)).toBe(7);
  });
});

describe('xlsx/formulas.ts: formatComputedValueForStorage', () => {
  test('formats a number, trimming floating point noise', () => {
    expect(formatComputedValueForStorage(0.1 + 0.2)).toBe('0.3');
  });
  test('formats a boolean as TRUE/FALSE', () => {
    expect(formatComputedValueForStorage(true)).toBe('TRUE');
  });
  test('formats an error name as-is', () => {
    expect(formatComputedValueForStorage('#DIV/0!')).toBe('#DIV/0!');
  });
});
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/xlsxFormulas.test.ts`
Expected: FAIL — `src/xlsx/formulas.ts`가 아직 없음.

- [ ] **Step 3: `src/xlsx/formulas.ts` 작성**

```ts
import FormulaParser, { FormulaError, type CellRef, type RangeRef } from 'fast-formula-parser';
import { formulaFunctions, formulaFunctionsNeedContext } from '../client/formulaOverrides.js';
import type { GridDocument } from '../spreadsheet.js';

/** 셀 하나의 원본 텍스트를 canonical 그리드에서 읽는다 — computeGridCellDisplay의 재귀
 * 참조(onCell/onRange)와 최초 호출이 공유하는 저수준 접근자. */
function rawCellText(grid: GridDocument, sheetIdx: number, row: number, col: number): string {
  const sheet = grid.sheets[sheetIdx];
  if (!sheet) return '';
  const cell = sheet.rows[row]?.[col];
  return cell === 0 || cell === undefined ? '' : cell.v;
}

/**
 * 셀 하나의 "표시용" 값을 계산한다 — 클라이언트의 옛 computeCellDisplay(grid-editor.ts)와
 * 동일한 로직이지만, Yjs cells 맵 대신 canonical GridDocument를 직접 읽으므로 화면 밖(로드
 * 안 된 페이지) 참조도 항상 정확하다. 순환 참조 가드도 호출마다 새로 만든다(재계산이 매번
 * 독립적이어야 하므로, 클라이언트처럼 클로저에 고정된 evalGuard를 쓰지 않는다).
 */
export function computeGridCellDisplay(
  grid: GridDocument,
  sheetIdx: number,
  row: number,
  col: number,
  evalGuard: Set<string> = new Set(),
): unknown {
  const key = `${sheetIdx}:${row}:${col}`;
  const raw = rawCellText(grid, sheetIdx, row, col);
  if (!raw.startsWith('=')) {
    if (raw.trim() === '') return '';
    const num = Number(raw);
    return Number.isFinite(num) ? num : raw;
  }
  if (evalGuard.has(key)) return FormulaError.REF.name;
  evalGuard.add(key);
  try {
    const parser = new FormulaParser({
      functions: formulaFunctions,
      functionsNeedContext: formulaFunctionsNeedContext,
      onCell: (ref: CellRef) => computeGridCellDisplay(grid, sheetIdx, ref.row - 1, ref.col - 1, evalGuard),
      onRange: (ref: RangeRef) => {
        const rows: unknown[][] = [];
        for (let r = ref.from.row; r <= ref.to.row; r += 1) {
          const rowVals: unknown[] = [];
          for (let c = ref.from.col; c <= ref.to.col; c += 1) {
            rowVals.push(computeGridCellDisplay(grid, sheetIdx, r - 1, c - 1, evalGuard));
          }
          rows.push(rowVals);
        }
        return rows;
      },
    });
    return parser.parse(raw.slice(1), { row: row + 1, col: col + 1, sheet: 'Sheet1' });
  } catch (error) {
    return error instanceof FormulaError ? error.name : '#ERROR!';
  } finally {
    evalGuard.delete(key);
  }
}

/** computeGridCellDisplay의 결과를 GRID_COMPUTED_MAP에 저장할 문자열로 바꾼다 — 서식(쉼표·
 * 통화 등)은 클라이언트가 스타일 정보로 따로 입히므로 여기서는 순수 값만 문자열화한다. */
export function formatComputedValueForStorage(value: unknown): string {
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value !== 'number') return value === null || value === undefined ? '' : String(value);
  return String(Math.round(value * 1e10) / 1e10);
}
```

**주의**: `FormulaParser`를 매 셀 계산마다 새로 만드는 건 비효율적이다(재귀 호출마다도 새로
만들어진다, 위 코드가 그렇다) — 이건 의도적 단순화다(YAGNI, 스펙의 "의존성 그래프 없이 단순함
유지" 원칙과 일관). Task 5에서 실제 재계산 트리거를 연결할 때, "페이지 하나를 재계산"하는
바깥 루프에서 `FormulaParser` 인스턴스 하나를 만들어 그 안의 여러 셀 계산이 공유하도록 리팩터링
할 여지를 남겨둔다(지금 이 태스크에서는 정확성만 확보, 최적화는 Task 5로 미룸 — 다만 Task 5는
그 리팩터링을 반드시 하지는 않아도 된다, 실측 후 필요하면 한다).

- [ ] **Step 4: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/xlsxFormulas.test.ts`
Expected: PASS (7 tests)

- [ ] **Step 5: 타입 체크**

Run: `npx tsc --noEmit`
Expected: PASS

- [ ] **Step 6: 커밋**

```bash
git add src/xlsx/formulas.ts tests/xlsxFormulas.test.ts
git commit -m "feat: 수식 계산 엔진을 canonical 그리드 기준으로 서버에 이식"
```

---

### Task 5: 서버 — 수식 재계산 트리거 (편집 시 + 페이지 로드 시) + 양보 지점

**Files:**
- Modify: `src/collab/room.ts` (Task 2의 `applyYjsUpdateToCanonical` 호출부 근처에 재계산 훅
  추가, `Room`에 재계산 대상 페이지 추적용 필드는 아직 필요 없음 — Task 5 시점엔 아직 Yjs가
  전체를 담고 있으므로 "지금 구독 중인 페이지"라는 개념이 없다. 이 태스크는 **"편집이 있으면
  그 시트 전체의 수식 셀을 재계산해 GRID_COMPUTED_MAP에 반영"**까지만 한다 — "구독 중인
  페이지만"으로 범위를 좁히는 건 Task 6 이후. 과도기 단계이므로 지금은 시트 전체가 사실상
  "항상 로드된 페이지"와 같다.)
- Test: `tests/xlsxFormulaRecompute.test.ts`(신규)

**Interfaces:**
- Consumes: Task 2의 `Room.canonicalGrid`, Task 4의 `computeGridCellDisplay`/
  `formatComputedValueForStorage`, `yieldToEventLoop`/`YIELD_EVERY_N_ROWS`(Task 1에서 export).
- Produces: `recomputeFormulasForSheet(room: Room, sheetIdx: number): Promise<void>` — 이
  함수를 Task 6에서 "구독 중인 페이지만" 버전으로 확장한다(함수 이름과 시그니처를 그대로
  유지하되 내부에서 행 범위를 제한하도록 Task 6이 수정할 예정 — 지금은 항상 시트 전체를
  훑는다).

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/xlsxFormulaRecompute.test.ts`(신규 — FakeSocket 패턴, Task 2 테스트와 유사한 골격):
```ts
import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeUpdate } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

const MESSAGE_SYNC = 0;

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  readonly mirror = new Y.Doc();

  send(data: Uint8Array, cb?: (err?: Error) => void): void {
    const decoder = decoding.createDecoder(data);
    const messageType = decoding.readVarUint(decoder);
    if (messageType === MESSAGE_SYNC) {
      const syncType = decoding.readVarUint(decoder);
      if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) {
        readUpdate(decoder, this.mirror, this);
      }
    }
    cb?.();
  }
  close(): void {}
  ping(): void {}
}

async function makeSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 10;
  sheet.getCell('A2').value = 20;
  sheet.getCell('A3').value = { formula: 'SUM(A1:A2)' };
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
}

function sendCellEdit(socket: FakeSocket, key: string, value: string): void {
  const scratch = new Y.Doc();
  scratch.getMap<string>('cells').set(key, value);
  const update = Y.encodeStateAsUpdate(scratch);
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeUpdate(encoder, update);
  socket.emit('message', encoding.toUint8Array(encoder));
}

describe('room.ts: formula cells recompute into GRID_COMPUTED_MAP on edit', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-formula-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.useRealTimers();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  test('editing a cell that a formula depends on updates the formula cell\'s computed value', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession } = await import('../src/collab/room.js');
    const { gridCellKey, GRID_COMPUTED_MAP } = await import('../src/collab/gridShared.js');

    const docId = randomUUID();
    const original = await makeSampleXlsx();
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);

    // A3는 =SUM(A1:A2) — 처음엔 30이어야 한다(시딩 직후 최초 재계산).
    let computed = socket.mirror.getMap<string>(GRID_COMPUTED_MAP);
    // 시딩 시점에도 최초 1회 재계산이 돌아야 하므로, sync 응답을 받기 위해 클라이언트가 자기
    // 상태를 먼저 보내야 한다 — 여기서는 서버가 보낸 초기 syncStep1에 대해 별도 응답 없이도
    // 서버가 broadcast하는 update만으로 mirror가 채워지는지 확인한다(room.ts가 시딩 완료 후
    // 자체 update를 브로드캐스트하지 않는다면, 이 값은 편집 이후에야 처음 나타난다 — 그 경우
    // 아래 편집 이후의 값만 검증하면 충분하다. 구현 결과에 따라 이 초기 검증 줄은 생략 가능).

    sendCellEdit(socket, gridCellKey(0, 0, 0), '100'); // A1을 100으로 편집 → SUM은 120이 되어야 함

    computed = socket.mirror.getMap<string>(GRID_COMPUTED_MAP);
    expect(computed.get(gridCellKey(0, 2, 0))).toBe('120');
  });
});
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/xlsxFormulaRecompute.test.ts`
Expected: FAIL — `computed.get(...)`가 `undefined`(아직 재계산 로직이 없음).

- [ ] **Step 3: `src/collab/room.ts` 수정**

`recomputeFormulasForSheet` 함수를 추가(Task 2의 `applyYjsUpdateToCanonical` 근처):
```ts
/**
 * 시트 하나의 모든 수식 셀을 canonical 그리드 기준으로 다시 계산해 GRID_COMPUTED_MAP에
 * 반영한다. 넓은 범위를 참조하는 수식이 많으면 이 루프가 길어질 수 있으므로
 * YIELD_EVERY_N_ROWS마다 이벤트 루프에 양보한다 — 대형 xlsx가 서버를 멈추던 버그(6355a17)가
 * 이 새 반복문에서 재발하지 않도록 하는 필수 조치(계획 Global Constraints 참고).
 */
async function recomputeFormulasForSheet(room: Room, sheetIdx: number): Promise<void> {
  if (!room.canonicalGrid) return;
  const sheet = room.canonicalGrid.sheets[sheetIdx];
  if (!sheet) return;
  const computed = room.ydoc.getMap<string>(GRID_COMPUTED_MAP);
  for (let start = 0; start < sheet.rows.length; start += YIELD_EVERY_N_ROWS) {
    const end = Math.min(start + YIELD_EVERY_N_ROWS, sheet.rows.length);
    room.ydoc.transact(() => {
      for (let r = start; r < end; r += 1) {
        const row = sheet.rows[r];
        for (let c = 0; c < row.length; c += 1) {
          const cell = row[c];
          const key = gridCellKey(sheetIdx, r, c);
          if (cell === 0 || !cell.v.startsWith('=')) {
            if (computed.has(key)) computed.delete(key);
            continue;
          }
          const value = computeGridCellDisplay(room.canonicalGrid!, sheetIdx, r, c);
          computed.set(key, formatComputedValueForStorage(value));
        }
      }
    });
    await yieldToEventLoop();
  }
}
```
Import 추가: `import { computeGridCellDisplay, formatComputedValueForStorage } from '../xlsx/formulas.js';`,
`import { GRID_COMPUTED_MAP } from './gridShared.js';`(기존 import 목록에 병합),
`import { YIELD_EVERY_N_ROWS, yieldToEventLoop } from '../spreadsheet.js';`(기존
`parseXlsxToGrid` import와 병합).

`seedRoom`의 xlsx 분기 끝(`room.baselineGrid = structuredClone(grid);` 다음)에 최초 재계산
호출을 추가:
```ts
      room.baselineGrid = structuredClone(grid);
      for (let sheetIdx = 0; sheetIdx < grid.sheets.length; sheetIdx += 1) {
        await recomputeFormulasForSheet(room, sheetIdx);
      }
```

`ydoc.on('update', ...)` 핸들러(Task 2에서 추가한 `applyYjsUpdateToCanonical` 호출부) 바로
다음 줄에 재계산 트리거를 추가 — **주의: 이 재계산 자체가 Yjs에 또 update를 쓰므로
(computed.set), suppressAutosaveReschedule로 감싸지 않으면 무한 루프가 재발한다(이 세션
앞부분에서 겪은 버그와 같은 클래스)**:
```ts
    if (room.kind === 'xlsx' && room.canonicalGrid) {
      applyYjsUpdateToCanonical(room.canonicalGrid, ydoc, update);
      const sheetIdx = 0; // Task 6 이전엔 항상 활성 시트 0으로 단순화 — 다중 시트 재계산 확장은 Task 6
      room.suppressAutosaveReschedule = true;
      void recomputeFormulasForSheet(room, sheetIdx).finally(() => {
        room.suppressAutosaveReschedule = false;
      });
    }
```

**중요**: `recomputeFormulasForSheet`가 `await yieldToEventLoop()`를 포함한 **비동기**
함수이므로, `suppressAutosaveReschedule`를 껐다 켜는 타이밍이 이 재계산이 끝난 뒤여야 한다
(`void ....finally(...)`로 처리 — 위 코드가 그렇게 되어 있다). 다만 이 재계산이 진행되는
동안(양보 지점 사이) **진짜 사용자 편집이 도착하면** 그 편집의 update 이벤트도
`suppressAutosaveReschedule`가 `true`인 상태에서 처리되어 자동저장이 예약되지 않는 문제가
생길 수 있다 — 이건 실제 위험이므로, 구현자는 이 정확한 레이스 컨디션을 테스트로 확인해야
한다(아래 Step 3.5).

- [ ] **Step 3.5: 레이스 컨디션 테스트 추가 (재계산 도중 도착하는 진짜 편집이 자동저장을 놓치지 않는지)**

`tests/xlsxFormulaRecompute.test.ts`에 테스트 추가:
```ts
  test('a real edit arriving while a formula recompute is still yielding still schedules autosave', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

    const { store } = await import('../src/storage.js');
    const { joinEditSession } = await import('../src/collab/room.js');
    const { gridCellKey } = await import('../src/collab/gridShared.js');

    const docId = randomUUID();
    // 재계산이 최소 한 번은 양보하도록(YIELD_EVERY_N_ROWS=2000을 넘도록) 큰 시트를 만든다.
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    for (let r = 1; r <= 3000; r += 1) sheet.getCell(r, 1).value = r;
    const original = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);

    const saveSpy = vi.spyOn(store, 'saveOriginal');
    sendCellEdit(socket, gridCellKey(0, 0, 0), '999');

    // 재계산이 setImmediate로 양보하는 동안 실제 이벤트 루프가 여러 턴 돌아야 하므로, 마이크로
    // 태스크뿐 아니라 setImmediate 턴도 충분히 흘려보낸다.
    for (let i = 0; i < 100; i += 1) await new Promise((resolve) => setImmediate(resolve));

    await vi.advanceTimersByTimeAsync(8_100); // 8초 자동저장 대기
    for (let i = 0; i < 100; i += 1) await new Promise((resolve) => setImmediate(resolve));

    expect(saveSpy).toHaveBeenCalled();
  });
```

이 테스트가 실패하면(자동저장이 안 불림), `room.suppressAutosaveReschedule`를 편집이 도착한
"그 순간"만 좁게 감싸는 대신 재계산 전체 기간 동안 켜 둔 것이 원인이다 — 수정 방향:
`applyYjsUpdateToCanonical` 호출(동기)만 `suppressAutosaveReschedule`로 감싸고,
`recomputeFormulasForSheet`(비동기, 양보 지점 있음)는 **감싸지 않는다** — 재계산 자체가
Yjs에 쓰는 매 `ydoc.transact(...)` 호출마다 스스로 `scheduleAutosave`를 다시 트리거하는 건
사실 문제가 아니다(그 트리거들은 8초 디바운스로 자연스럽게 하나로 합쳐진다) — 진짜 막아야 할
무한 루프는 "재계산 결과 자체가 편집처럼 보여 또 재계산을 유발"하는 경우인데,
`recomputeFormulasForSheet`는 `computed` 맵만 쓰고 `cells` 맵은 안 건드리므로애초에 재귀
트리거가 안 걸린다 — 즉 **이 함수는 suppressAutosaveReschedule로 감쌀 필요가 자체가 없을 수
있다**. 구현자는 이 가능성을 먼저 검토하고(감싸지 않는 더 단순한 버전으로 테스트가 통과하는지
먼저 시도), 정말 필요한 경우에만 좁게 감싼다 — "일단 안전하게 넓게 감싸고 본다"는 이
세션에서 이미 겪은 실수(자동저장 무한 루프)의 반대 방향 실수(자동저장 누락)이므로 신중하게
판단할 것.

- [ ] **Step 4: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/xlsxFormulaRecompute.test.ts`
Expected: PASS (2 tests)

- [ ] **Step 5: 회귀 확인**

Run: `npx vitest run tests/roomCanonicalGrid.test.ts tests/roomAutosaveLoop.test.ts tests/roomGridIndex.test.ts tests/xlsx-collab.test.ts && npx tsc --noEmit`
Expected: 모두 PASS.

- [ ] **Step 6: 커밋**

```bash
git add src/collab/room.ts tests/xlsxFormulaRecompute.test.ts
git commit -m "feat: 셀 편집 시 수식 셀을 canonical 기준으로 재계산해 computed 맵에 반영"
```

---

### Task 6: 서버 — 페이지 구독 프로토콜 (아직 실제 윈도잉은 켜지 않음, 구독 추적만)

**Files:**
- Modify: `src/collab/room.ts` (`MESSAGE_PAGE_SUBSCRIBE` 상수, `Room`에 구독 추적 필드,
  `handleMessage`에 새 case, `closeConn`에서 구독 정리)
- Test: `tests/roomPageSubscription.test.ts`(신규)

**Interfaces:**
- Consumes: 없음(새 프로토콜).
- Produces: `MESSAGE_PAGE_SUBSCRIBE = 2`(export 불필요, room.ts 내부 상수).
  `Room.pageSubscriptions: Map<WebSocket, { sheet: number; pages: readonly number[] }>`.
  `getSubscribedPageUnion(room: Room, sheet: number): Set<number>` — 이 함수를 Task 8이
  재사용해 실제 Yjs 채움/비움을 구현한다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/roomPageSubscription.test.ts`(신규 — FakeSocket 패턴):
```ts
import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as encoding from 'lib0/encoding';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

const MESSAGE_SYNC = 0;
const MESSAGE_PAGE_SUBSCRIBE = 2;

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  send(_data: Uint8Array, cb?: (err?: Error) => void): void { cb?.(); }
  close(): void {}
  ping(): void {}
}

function sendSubscribe(socket: FakeSocket, sheet: number, pages: readonly number[]): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_PAGE_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify({ sheet, pages }));
  socket.emit('message', encoding.toUint8Array(encoder));
}

async function makeSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Sheet1').getCell('A1').value = 'x';
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

describe('room.ts: page subscription tracking', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-subscribe-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  test('a single client subscribing to pages [0, 2] is reflected in the union', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getSubscribedPageUnionForTest } = await import('../src/collab/room.js');

    const docId = randomUUID();
    const original = await makeSampleXlsx();
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
    sendSubscribe(socket, 0, [0, 2]);

    const union = getSubscribedPageUnionForTest(docId, 0);
    expect([...union].sort()).toEqual([0, 2]);
  });

  test('two clients subscribing to overlapping pages produce the correct union', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getSubscribedPageUnionForTest } = await import('../src/collab/room.js');

    const docId = randomUUID();
    const original = await makeSampleXlsx();
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid',
    });

    const socketA = new FakeSocket();
    const socketB = new FakeSocket();
    await joinEditSession(socketA as never, {} as never, docId, 'xlsx', [], false);
    await joinEditSession(socketB as never, {} as never, docId, 'xlsx', [], false);
    sendSubscribe(socketA, 0, [0, 1]);
    sendSubscribe(socketB, 0, [1, 2]);

    const union = getSubscribedPageUnionForTest(docId, 0);
    expect([...union].sort()).toEqual([0, 1, 2]);
  });

  test('a client re-subscribing to a new sheet/page list replaces its previous subscription entirely', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getSubscribedPageUnionForTest } = await import('../src/collab/room.js');

    const docId = randomUUID();
    const original = await makeSampleXlsx();
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
    sendSubscribe(socket, 0, [0, 1]);
    sendSubscribe(socket, 0, [5]); // 스크롤해서 완전히 다른 페이지로 이동

    const union = getSubscribedPageUnionForTest(docId, 0);
    expect([...union]).toEqual([5]);
  });

  test('disconnecting a client removes its subscription from the union', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getSubscribedPageUnionForTest } = await import('../src/collab/room.js');

    const docId = randomUUID();
    const original = await makeSampleXlsx();
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
    sendSubscribe(socket, 0, [0]);
    expect([...getSubscribedPageUnionForTest(docId, 0)]).toEqual([0]);

    socket.emit('close');
    expect([...getSubscribedPageUnionForTest(docId, 0)]).toEqual([]);
  });
});
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/roomPageSubscription.test.ts`
Expected: FAIL — `getSubscribedPageUnionForTest`가 없음.

- [ ] **Step 3: `src/collab/room.ts` 수정**

`MESSAGE_AWARENESS = 1;` 다음 줄에 추가:
```ts
const MESSAGE_PAGE_SUBSCRIBE = 2;
```

`Room` 인터페이스에 필드 추가(`baselineGrid` 다음):
```ts
  /** xlsx 룸에서만 쓰인다 — 연결별로 "지금 이 시트의 이 페이지들을 보고 있다"를 추적한다.
   * 연결이 끊기면(closeConn) 이 맵에서 즉시 제거한다. */
  readonly pageSubscriptions: Map<WebSocket, { sheet: number; pages: readonly number[] }>;
```

`getOrCreateRoom`의 `Room` 객체 리터럴에 `pageSubscriptions: new Map(),` 추가.

`handleMessage`의 `switch`에 새 case 추가(`case MESSAGE_AWARENESS:` 다음):
```ts
      case MESSAGE_PAGE_SUBSCRIBE: {
        try {
          const payload = decoding.readVarString(decoder);
          const parsed = JSON.parse(payload) as { sheet: number; pages: readonly number[] };
          if (
            typeof parsed.sheet === 'number' &&
            Array.isArray(parsed.pages) &&
            parsed.pages.every((p) => typeof p === 'number')
          ) {
            room.pageSubscriptions.set(conn, { sheet: parsed.sheet, pages: parsed.pages });
          }
        } catch (error) {
          console.error('페이지 구독 메시지 처리 오류:', error);
        }
        break;
      }
```

`closeConn`에서 `room.readOnlyConns.delete(conn);` 다음 줄에 추가:
```ts
    room.pageSubscriptions.delete(conn);
```

`getSubscribedPageUnion` 함수 추가(파일 하단, `activeRoomCount` 근처):
```ts
function getSubscribedPageUnion(room: Room, sheet: number): Set<number> {
  const union = new Set<number>();
  for (const subscription of room.pageSubscriptions.values()) {
    if (subscription.sheet !== sheet) continue;
    for (const page of subscription.pages) union.add(page);
  }
  return union;
}

/** 테스트 전용 */
export function getSubscribedPageUnionForTest(docId: string, sheet: number): Set<number> {
  const room = rooms.get(docId);
  if (!room) return new Set();
  return getSubscribedPageUnion(room, sheet);
}
```

- [ ] **Step 4: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/roomPageSubscription.test.ts`
Expected: PASS (4 tests)

- [ ] **Step 5: 회귀 확인 + 타입 체크**

Run: `npx vitest run tests/roomCanonicalGrid.test.ts tests/roomAutosaveLoop.test.ts tests/roomGridIndex.test.ts tests/xlsx-collab.test.ts tests/xlsxFormulaRecompute.test.ts && npx tsc --noEmit`
Expected: 모두 PASS.

- [ ] **Step 6: 커밋**

```bash
git add src/collab/room.ts tests/roomPageSubscription.test.ts
git commit -m "feat: 페이지 구독 메시지 프로토콜 추가(추적만, 아직 실제 윈도잉 없음)"
```

---

### Task 7: 서버 — seedRoom을 윈도잉으로 전환 (컷오버, 가장 위험한 태스크)

**이 태스크가 실제로 동작을 바꾼다** — Task 1~6은 전부 "동작은 그대로, 내부 배관만 준비"였다.
이 태스크부터 Yjs는 더 이상 전체 셀을 담지 않는다.

**Files:**
- Modify: `src/collab/room.ts` (`seedRoom`의 xlsx 분기 — 전체 `seedGridFromXlsx` 호출 제거,
  페이지 기반 채움으로 교체. `getSubscribedPageUnion` 변경 시 Yjs 채움/비움을 트리거하는 로직
  추가. `recomputeFormulasForSheet`를 "구독 중인 페이지만" 도는 버전으로 좁힘)
- Modify: `src/xlsx/collab.ts` (페이지 하나 분량만 시딩하는 새 함수 추가 — 기존
  `seedGridFromXlsx`의 내부 루프 로직을 재사용)
- Test: `tests/roomWindowedSeeding.test.ts`(신규), 기존 `tests/roomPageSubscription.test.ts`
  확장

**Interfaces:**
- Consumes: Task 6의 `getSubscribedPageUnion`, Task 2의 `canonicalGrid`, Task 5의
  `recomputeFormulasForSheet`.
- Produces: `seedGridPageFromXlsx(ydoc: Y.Doc, canonicalGrid: GridDocument, sheetIdx: number, page: number): Promise<void>`
  (`src/xlsx/collab.ts`, export) — 페이지 하나(500행)만 Yjs에 채운다.
  `unseedGridPageFromXlsx(ydoc: Y.Doc, sheetIdx: number, page: number): void` — 페이지 하나를
  Yjs에서 지운다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/roomWindowedSeeding.test.ts`(신규):
```ts
import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeSyncStep1 } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

const MESSAGE_SYNC = 0;
const MESSAGE_PAGE_SUBSCRIBE = 2;

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  readonly mirror = new Y.Doc();

  send(data: Uint8Array, cb?: (err?: Error) => void): void {
    const decoder = decoding.createDecoder(data);
    const messageType = decoding.readVarUint(decoder);
    if (messageType === MESSAGE_SYNC) {
      const syncType = decoding.readVarUint(decoder);
      if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) {
        readUpdate(decoder, this.mirror, this);
      }
    }
    cb?.();
  }
  close(): void {}
  ping(): void {}

  requestSync(): void {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeSyncStep1(encoder, this.mirror);
    this.emit('message', encoding.toUint8Array(encoder));
  }
}

function sendSubscribe(socket: FakeSocket, sheet: number, pages: readonly number[]): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_PAGE_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify({ sheet, pages }));
  socket.emit('message', encoding.toUint8Array(encoder));
}

async function makeBigXlsx(rows: number): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  for (let r = 1; r <= rows; r += 1) sheet.getCell(r, 1).value = `row-${r}`;
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

describe('room.ts: xlsx seeding is windowed by subscribed page', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-windowed-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.useRealTimers();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  test('opening a large sheet does NOT seed all rows into Yjs — only the subscribed page', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession } = await import('../src/collab/room.js');
    const { GRID_CELLS_MAP, gridCellKey } = await import('../src/collab/gridShared.js');

    const docId = randomUUID();
    const original = await makeBigXlsx(3000); // 6개 페이지(0~5, 각 500행)
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
    sendSubscribe(socket, 0, [0]); // 첫 페이지(0~499행)만 구독
    socket.requestSync();

    const cells = socket.mirror.getMap<string>(GRID_CELLS_MAP);
    expect(cells.get(gridCellKey(0, 0, 0))).toBe('row-1'); // 구독한 페이지 안 — 있어야 함
    expect(cells.get(gridCellKey(0, 2500, 0))).toBeUndefined(); // 구독 안 한 페이지(5번) — 없어야 함
  });

  test('subscribing to a new page fills it in; the union still includes previously subscribed pages', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession } = await import('../src/collab/room.js');
    const { GRID_CELLS_MAP, gridCellKey } = await import('../src/collab/gridShared.js');

    const docId = randomUUID();
    const original = await makeBigXlsx(3000);
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
    sendSubscribe(socket, 0, [0, 3]); // 0페이지(0~499행)와 3페이지(1500~1999행) 둘 다
    socket.requestSync();

    const cells = socket.mirror.getMap<string>(GRID_CELLS_MAP);
    expect(cells.get(gridCellKey(0, 0, 0))).toBe('row-1');
    expect(cells.get(gridCellKey(0, 1500, 0))).toBe('row-1501');
    expect(cells.get(gridCellKey(0, 2500, 0))).toBeUndefined();
  });

  test('a page removed from the union stays in Yjs for the grace period, then is removed', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

    const { store } = await import('../src/storage.js');
    const { joinEditSession } = await import('../src/collab/room.js');
    const { GRID_CELLS_MAP, gridCellKey } = await import('../src/collab/gridShared.js');

    const docId = randomUUID();
    const original = await makeBigXlsx(1000);
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
    sendSubscribe(socket, 0, [0]);
    socket.requestSync();
    for (let i = 0; i < 50; i += 1) await new Promise((resolve) => setImmediate(resolve));

    let cells = socket.mirror.getMap<string>(GRID_CELLS_MAP);
    expect(cells.get(gridCellKey(0, 0, 0))).toBe('row-1');

    sendSubscribe(socket, 0, [1]); // 0페이지 구독 해제, 1페이지로 이동
    for (let i = 0; i < 50; i += 1) await new Promise((resolve) => setImmediate(resolve));

    // 유예 시간(10초) 이전 — 0페이지가 아직 Yjs에 남아있어야 한다(canonical에는 항상 있으므로
    // mirror 쪽에서 "지워짐"을 직접 관찰하긴 어렵다 — 대신 서버 내부 상태를 테스트 전용 함수로
    // 확인한다. 구현자는 Task 6의 getSubscribedPageUnionForTest와 유사하게, "지금 Yjs에 실제로
    // 채워져 있는 페이지 집합"을 들여다보는 테스트 전용 export를 이 태스크에서 추가한다 —
    // 예: getSeededPagesForTest(docId, sheet): Set<number>).
    const { getSeededPagesForTest } = await import('../src/collab/room.js');
    expect(getSeededPagesForTest(docId, 0).has(0)).toBe(true);

    await vi.advanceTimersByTimeAsync(10_100); // 유예 시간 경과
    for (let i = 0; i < 50; i += 1) await new Promise((resolve) => setImmediate(resolve));

    expect(getSeededPagesForTest(docId, 0).has(0)).toBe(false);
    expect(getSeededPagesForTest(docId, 0).has(1)).toBe(true);
  });
});
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/roomWindowedSeeding.test.ts`
Expected: FAIL — 지금은 `seedRoom`이 여전히 전체를 시딩하므로 "구독 안 한 페이지가 없어야
한다"는 assertion들이 실패한다.

- [ ] **Step 3: `src/xlsx/collab.ts`에 페이지 단위 시딩/해제 함수 추가**

기존 `seedGridFromXlsx`(Task 1 이후 상태) 함수 바로 뒤에 추가:
```ts
const PAGE_SIZE_ROWS = 500;

/** 시트 하나의 페이지 하나(500행)만 Yjs에 채운다 — seedGridFromXlsx의 내부 페이지 루프
 * 본문과 동일한 로직을 한 페이지 분량으로 좁힌 버전. */
export async function seedGridPageFromXlsx(
  ydoc: Y.Doc,
  canonicalGrid: GridDocument,
  sheetIdx: number,
  page: number,
): Promise<void> {
  const sheet = canonicalGrid.sheets[sheetIdx];
  if (!sheet) return;
  const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
  const baseline = ydoc.getMap<string>(GRID_BASELINE_MAP);
  const styles = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_MAP);
  const stylesBaseline = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_BASELINE_MAP);
  const covered = mergeCoveredCells(sheet.merges);
  const start = page * PAGE_SIZE_ROWS;
  const end = Math.min(start + PAGE_SIZE_ROWS, sheet.rows.length);

  ydoc.transact(() => {
    for (let r = start; r < end; r += 1) {
      const row = sheet.rows[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c += 1) {
        const cell = row[c];
        if (cell === 0) continue;
        if (covered.has(`${r}:${c}`)) continue;
        const key = gridCellKey(sheetIdx, r, c);
        const style = cell.s !== undefined ? canonicalGrid.styles[cell.s] : undefined;
        if (cell.v !== '') {
          const value = seedableCellValue(cell.v, style);
          cells.set(key, value);
          baseline.set(key, value);
        }
        if (style) {
          const attrs = toCellStyleAttrs(style);
          if (Object.keys(attrs).length > 0) {
            styles.set(key, attrs);
            stylesBaseline.set(key, attrs);
          }
        }
      }
    }
  });
}

/** 페이지 하나를 Yjs에서 지운다(canonical에는 영향 없음 — 그쪽이 항상 진실이다). */
export function unseedGridPageFromXlsx(ydoc: Y.Doc, sheetIdx: number, page: number, rowCount: number): void {
  const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
  const baseline = ydoc.getMap<string>(GRID_BASELINE_MAP);
  const styles = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_MAP);
  const stylesBaseline = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_BASELINE_MAP);
  const computed = ydoc.getMap<string>(GRID_COMPUTED_MAP);
  const start = page * PAGE_SIZE_ROWS;
  const end = Math.min(start + PAGE_SIZE_ROWS, rowCount);
  ydoc.transact(() => {
    for (const map of [cells, baseline, styles, stylesBaseline, computed]) {
      for (const key of [...map.keys()]) {
        const parsed = parseGridCellKey(key);
        if (parsed && parsed.sheetIdx === sheetIdx && parsed.row >= start && parsed.row < end) {
          map.delete(key);
        }
      }
    }
  });
}
```
`PAGE_SIZE_ROWS`를 export해서 room.ts와 클라이언트(Task 9)가 같은 값을 참조하게 한다 —
`export const PAGE_SIZE_ROWS = 500;`로 바꾸고 `gridShared.ts`로 옮기는 게 더 적절하다(서버·
클라이언트 양쪽에서 import 가능해야 하므로 `xlsx/collab.ts`는 서버 전용 모듈일 수 있다 — 실제로
`GRID_CELLS_MAP` 등도 `gridShared.ts`에 있는 이유가 이것이다). **구현자는 `PAGE_SIZE_ROWS`를
`src/collab/gridShared.ts`에 정의하고 이 파일에서 import하도록** 위 코드를 조정할 것.

`GRID_COMPUTED_MAP`/`parseGridCellKey`를 이 파일의 `gridShared.js` import 목록에 추가.

- [ ] **Step 4: `src/collab/room.ts`의 `seedRoom` 전환**

`seedRoom`의 xlsx 분기를 수정 — 전체 시딩(`await seedGridFromXlsx(ydoc, grid);`) 대신, 초기
구독 없음 상태로 시작하고(아무 페이지도 안 채움), 첫 연결이 구독 메시지를 보내면 그때 채운다:
```ts
    if (kind === 'xlsx') {
      const grid = await parseXlsxToGrid(original);
      room.canonicalGrid = grid;
      room.baselineGrid = structuredClone(grid);
      for (let sheetIdx = 0; sheetIdx < grid.sheets.length; sheetIdx += 1) {
        await recomputeFormulasForSheet(room, sheetIdx);
      }
      // 이제 seedGridFromXlsx(전체 시딩)를 호출하지 않는다 — 구조/병합 메타데이터
      // (GRID_STRUCTURE_MAP/GRID_MERGES_MAP)만 채우고, 셀 데이터는 페이지 구독이 들어올 때
      // 채운다(핵심 제약: 구조 메타데이터는 절대 윈도잉하지 않는다).
      await seedGridStructureOnly(ydoc, grid);
    } else if (kind === 'pptx') {
```
`src/xlsx/collab.ts`에 `seedGridStructureOnly` 함수를 추가(기존 `seedGridFromXlsx`에서
`structure`/`merges` 맵 채우는 부분만 뽑아낸 것):
```ts
/** 셀 데이터 없이 구조 메타데이터(시트 이름·행렬 수·틀고정·병합)만 채운다 — 윈도잉 시작 시
 * 룸을 열자마자 호출한다. 클라이언트가 아직 로드 안 된 페이지의 스크롤 높이·병합 경계를 미리
 * 알아야 하므로 이건 절대 윈도잉하지 않는다(계획 Global Constraints). */
export async function seedGridStructureOnly(ydoc: Y.Doc, grid: GridDocument): Promise<void> {
  const structure = ydoc.getMap<SheetStructure>(GRID_STRUCTURE_MAP);
  const merges = ydoc.getMap<readonly GridMergeShape[]>(GRID_MERGES_MAP);
  ydoc.transact(() => {
    grid.sheets.forEach((sheet, sheetIdx) => {
      structure.set(String(sheetIdx), {
        name: sheet.name, rowCount: sheet.rowCount, colCount: sheet.colCount,
        freezeRows: sheet.freezeRows, freezeCols: sheet.freezeCols, colWidths: sheet.colWidths,
      });
      if (sheet.merges.length > 0) merges.set(String(sheetIdx), sheet.merges);
    });
  });
}
```

이제 `handleMessage`의 `MESSAGE_PAGE_SUBSCRIBE` case(Task 6)를 확장해, 구독이 바뀌면 실제로
Yjs를 채우고/지연 제거를 예약하도록 한다:
```ts
      case MESSAGE_PAGE_SUBSCRIBE: {
        try {
          const payload = decoding.readVarString(decoder);
          const parsed = JSON.parse(payload) as { sheet: number; pages: readonly number[] };
          if (
            typeof parsed.sheet === 'number' &&
            Array.isArray(parsed.pages) &&
            parsed.pages.every((p) => typeof p === 'number')
          ) {
            const before = getSubscribedPageUnion(room, parsed.sheet);
            room.pageSubscriptions.set(conn, { sheet: parsed.sheet, pages: parsed.pages });
            const after = getSubscribedPageUnion(room, parsed.sheet);
            void reconcilePageSubscriptions(room, parsed.sheet, before, after);
          }
        } catch (error) {
          console.error('페이지 구독 메시지 처리 오류:', error);
        }
        break;
      }
```

`reconcilePageSubscriptions`와 지연 제거 예약을 위한 상태를 추가:
```ts
const PAGE_UNSEED_GRACE_MS = 10_000;

/** 룸별·시트별로 "지연 제거 예약된 페이지 → 타이머" — 유예 시간 안에 재구독되면 취소한다. */
const pendingPageUnseeds = new Map<string, NodeJS.Timeout>();

function pageUnseedKey(docId: string, sheet: number, page: number): string {
  return `${docId}:${sheet}:${page}`;
}

async function reconcilePageSubscriptions(
  room: Room,
  sheet: number,
  before: Set<number>,
  after: Set<number>,
): Promise<void> {
  if (!room.canonicalGrid) return;
  // 새로 추가된 페이지 — 즉시 채우고, 대기 중이던 지연 제거가 있었다면 취소한다.
  for (const page of after) {
    const key = pageUnseedKey(room.docId, sheet, page);
    const pending = pendingPageUnseeds.get(key);
    if (pending) {
      clearTimeout(pending);
      pendingPageUnseeds.delete(key);
    }
    if (!before.has(page)) {
      await seedGridPageFromXlsx(room.ydoc, room.canonicalGrid, sheet, page);
      await recomputeFormulasForPage(room, sheet, page);
    }
  }
  // 합집합에서 빠진 페이지 — 유예 시간 뒤 조건부 제거를 예약한다.
  for (const page of before) {
    if (after.has(page)) continue;
    const key = pageUnseedKey(room.docId, sheet, page);
    const timer = setTimeout(() => {
      pendingPageUnseeds.delete(key);
      const stillUnsubscribed = !rooms.has(room.docId) || !getSubscribedPageUnion(room, sheet).has(page);
      if (stillUnsubscribed && room.canonicalGrid) {
        unseedGridPageFromXlsx(room.ydoc, sheet, page, room.canonicalGrid.sheets[sheet]?.rowCount ?? 0);
      }
    }, PAGE_UNSEED_GRACE_MS);
    pendingPageUnseeds.set(key, timer);
  }
}

/** 테스트 전용 */
export function getSeededPagesForTest(docId: string, sheet: number): Set<number> {
  const room = rooms.get(docId);
  if (!room) return new Set();
  const cells = room.ydoc.getMap<string>(GRID_CELLS_MAP);
  const pages = new Set<number>();
  for (const key of cells.keys()) {
    const parsed = parseGridCellKey(key);
    if (parsed && parsed.sheetIdx === sheet) pages.add(Math.floor(parsed.row / PAGE_SIZE_ROWS));
  }
  return pages;
}
```

`recomputeFormulasForPage`(Task 5의 `recomputeFormulasForSheet`를 페이지 범위로 좁힌 버전)를
추가 — `recomputeFormulasForSheet`의 본문을 거의 그대로 쓰되 `start`/`end`를
`sheet.rows.length` 전체가 아니라 `page * PAGE_SIZE_ROWS` ~ `min((page+1)*PAGE_SIZE_ROWS,
rowCount)`로 제한한다(정확한 코드는 `recomputeFormulasForSheet`를 복사해 범위만 좁혀
작성 — 두 함수 다 남겨두거나, `recomputeFormulasForSheet`를 `recomputeFormulasForPage`를
모든 페이지에 대해 호출하는 래퍼로 재작성해도 된다 — 구현자 판단에 맡긴다).

**편집 시 재계산 범위도 좁힌다** — Task 5에서 `sheetIdx`를 항상 `0`으로 하드코딩해 시트
전체를 재계산했는데, 이제는 "그 편집이 속한 페이지 + 그 페이지를 구독 중인 모든 클라이언트가
보는 페이지들"만 재계산해야 한다(정확히는: 수식이 다른 페이지 값을 참조할 수 있으므로 "편집된
셀이 속한 페이지"가 아니라 **"지금 구독 중인 모든 페이지"**를 재계산 대상으로 삼는 게 맞다 —
Task 5의 원래 의도와 같다, 다만 "시트 전체"가 아니라 "구독된 페이지들의 합집합"으로 범위를
좁힌다). 구현자는 `ydoc.on('update', ...)` 핸들러의 재계산 호출을
`recomputeFormulasForSheet(room, sheetIdx)` 대신
`for (const page of getSubscribedPageUnion(room, sheetIdx)) await recomputeFormulasForPage(room, sheetIdx, page);`
로 바꾼다.

`closeConn`에서 연결이 끊길 때도 구독 변화를 반영해야 한다 — `room.pageSubscriptions.delete(conn);`
다음 줄에:
```ts
    if (room.canonicalGrid) {
      // 이 연결이 마지막으로 구독하던 정보는 이미 지워졌으므로, 남은 구독자들의 합집합과
      // 비교할 "이전" 값을 conn 삭제 전에 미리 계산해 뒀어야 한다 — 위 delete보다 먼저
      // before를 구하고, delete 이후 after를 구해 reconcile한다(정확한 순서는 구현자가
      // room.pageSubscriptions.delete(conn) 호출 위치를 조정해 처리).
    }
```
**이 부분은 정확한 순서가 미묘하다** — 구현자는 `closeConn` 안에서 `conn`이 구독하던
`(sheet, pages)`를 **삭제하기 전에** 읽어 두고, 삭제 후 `getSubscribedPageUnion`으로 새 합집합을
구해 `reconcilePageSubscriptions`를 호출해야 한다. 여러 시트를 동시에 구독하는 경우는 없다는
전제(스펙의 "시트 전환" 절)이므로 `conn`이 구독하던 시트 하나만 reconcile하면 된다.

- [ ] **Step 5: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/roomWindowedSeeding.test.ts`
Expected: PASS (3 tests)

- [ ] **Step 6: 회귀 확인 — 이 태스크가 가장 위험하므로 전체 스위트를 돈다**

Run: `npx vitest run && npx tsc --noEmit`
Expected: 모두 PASS. 특히 `tests/xlsx-collab.test.ts`(Task 3에서 canonical 기준으로 이미
바뀌어 있으므로 윈도잉 자체와는 무관하게 계속 통과해야 함), `tests/roomPageSubscription.test.ts`,
`tests/xlsxFormulaRecompute.test.ts`가 전부 통과하는지 확인.

- [ ] **Step 7: 커밋**

```bash
git add src/collab/room.ts src/xlsx/collab.ts src/collab/gridShared.ts tests/roomWindowedSeeding.test.ts
git commit -m "feat: xlsx Yjs 시딩을 페이지 구독 기반 윈도잉으로 전환(컷오버)"
```

---

### Task 8: 클라이언트 — 수식 계산 제거, computed 맵 읽기로 전환

**Files:**
- Modify: `src/client/grid-editor.ts:1-20`(import 정리), `:341-468`(수식 계산 섹션 제거/축소)
- Modify: `package.json`(그리드 편집기 번들에서 `fast-formula-parser` 의존 제거는 이 태스크
  범위 밖 — import만 제거하면 esbuild 트리셰이킹이 자동으로 번들에서 빠진다, `package.json`
  자체는 서버 쪽(`src/xlsx/formulas.ts`)이 여전히 이 패키지를 쓰므로 손대지 않는다)

**Interfaces:**
- Consumes: Task 1의 `GRID_COMPUTED_MAP`.
- Produces: 없음(클라이언트 내부 변경).

- [ ] **Step 1: `grid-editor.ts` 수정**

파일 상단 import에서 제거:
```ts
import FormulaParser, { FormulaError, type CellRef, type RangeRef } from 'fast-formula-parser';
```
```ts
import { formulaFunctions, formulaFunctionsNeedContext } from './formulaOverrides.js';
```
`GRID_COMPUTED_MAP`을 `../collab/gridShared.js` import 목록에 추가.

`main()` 안에서 `const cells = ydoc.getMap<string>(GRID_CELLS_MAP);` 다음 줄에 추가:
```ts
  const computed = ydoc.getMap<string>(GRID_COMPUTED_MAP);
```

`:341-468`(수식 계산 섹션 — `formulaParser`, `computeCellDisplay`, `evalGuard`)을 통째로
삭제하고, 그 자리를 대체하는 훨씬 단순한 함수로 바꾼다:
```ts
  /** 셀의 표시용 값을 얻는다 — 수식 셀이면 서버가 계산해 둔 값(computed 맵)을, 아니면 원본
   * 텍스트(cells 맵)를 그대로 쓴다. 클라이언트는 더 이상 수식을 직접 계산하지 않는다(서버가
   * canonical 전체 데이터 기준으로 계산해야 화면 밖 참조도 정확하기 때문 — 설계 문서의 핵심
   * 결정 참고). */
  function computeCellDisplay(sheetIdx: number, row: number, col: number): unknown {
    const key = gridCellKey(sheetIdx, row, col);
    const computedValue = computed.get(key);
    if (computedValue !== undefined) return computedValue;
    const raw = cells.get(key) ?? '';
    if (raw.startsWith('=')) return ''; // 서버 계산이 아직 도착 전(그 페이지가 막 로드된 직후) — 잠시 빈 값
    if (raw.trim() === '') return '';
    const num = Number(raw);
    return Number.isFinite(num) ? num : raw;
  }
```
`formatComputedValue`(기존 `:387-392`)는 그대로 둔다(표시 형식 처리는 여전히 클라이언트가
한다 — 서버는 순수 계산값만 준다). `renderCellDisplay`/`applySpillDisplay`(기존
`:415-468`)에서 `computeCellDisplay` 호출부는 시그니처가 그대로(`sheetIdx, row, col`)이므로
수정 불필요 — 다만 내부적으로 `evalGuard`/`formulaParser`를 참조하던 다른 코드가 있는지
grep으로 확인하고(`formulaParser\.` 검색), 있으면 함께 제거한다.

- [ ] **Step 2: 빌드 확인**

Run: `npm run build:editor`
Expected: 성공. `grid-editor.bundle.js` 용량이 이전보다 줄어드는지 확인(참고용, 실패 조건
아님 — `fast-formula-parser` 제거로 번들이 작아지는 게 기대되는 부수 효과).

- [ ] **Step 3: 타입 체크**

Run: `npx tsc --noEmit`
Expected: PASS — `formulaOverrides.ts`를 더 이상 아무도 import 안 하면 미사용 파일 경고가 날
수 있으나(빌드는 안 깨짐), 이 파일은 삭제하지 않는다(Task 4에서 서버 쪽
`src/xlsx/formulas.ts`가 여전히 `../client/formulaOverrides.js`를 import하고 있으므로 —
파일 위치가 `client` 디렉터리에 있어도 순수 로직이라 서버에서 import해도 문제없다, 다만
이름이 어색하다는 지적이 최종 리뷰에서 나올 수 있음을 감안해 구현자는 원한다면
`src/xlsx/formulaOverrides.ts`로 옮기고 Task 4의 import 경로도 함께 고쳐도 된다 — 선택 사항).

- [ ] **Step 4: 커밋**

```bash
git add src/client/grid-editor.ts
git commit -m "refactor: 그리드 편집기 클라이언트에서 수식 계산 제거, computed 맵 읽기로 전환"
```

---

### Task 9: 클라이언트 — 가상 스크롤 테이블 렌더링 + 페이지 구독 전송 + 찾기/바꾸기 범위 경고

이 태스크는 코드량이 많고 UI 동작 검증이 필요해(이 프로젝트 관례상 DOM 테스트 하네스가
없음), 완료 후 반드시 `npm run dev`로 서버를 띄우고 실제 대형 xlsx로 브라우저 확인을 병행한다
(Step 마지막에 명시).

**Files:**
- Modify: `src/client/grid-editor.ts` (`renderCountsFor`/`renderTable` 함수를 스크롤 기반
  가상 렌더링으로 교체, 스크롤 리스너 추가, 페이지 구독 전송 로직 추가, 찾기/바꾸기 경고 문구
  추가)

**Interfaces:**
- Consumes: `PAGE_SIZE_ROWS`(Task 7에서 `gridShared.ts`로 옮긴 값), `provider`(기존
  `WebsocketProvider` 인스턴스, `main()` 안에 이미 존재).

- [ ] **Step 1: 가상 스크롤 렌더링으로 교체**

기존 `renderCountsFor`(`:1067-1081`)와 `renderTable`(`:1106`부터, 전체 시트 행만큼 `<tr>`을
만들던 로직)을, `public/js/viewer.js`의 `renderGridViewport`(뷰어 쪽 참조 구현, `:1014-1082`)
패턴을 따라 재작성한다 — 정확히 이식하지 말고(뷰어는 서버에서 행 데이터를 fetch하지만, 이
편집기는 이미 로컬 Yjs `cells` 맵에서 직접 읽으므로 fetch 단계가 없다) 다음 구조로 새로
작성한다:

```ts
  const ROW_HEIGHT_PX = 27; // 기존 상수 재사용(이미 파일에 정의돼 있음)
  const OVERSCAN_ROWS = 50; // 화면 위아래로 여유분(스크롤 시 빈 화면 깜빡임 방지)

  /** 전체 행 수만큼의 스크롤 높이를 유지하되, 실제 <tr>은 보이는 범위 + 여유분만 만든다. */
  function renderTableVirtualized(sheet: SheetStructure): void {
    const totalRows = sheet.rowCount + GRID_ROW_PADDING; // 기존 패딩 로직 유지(빈 행도 편집 가능해야 하므로)
    const { cols: renderColCount } = renderCountsFor(sheet); // 열은 기존처럼 화면 너비 기준(윈도잉 대상 아님 — 열은 최대 120개라 문제 없음)

    tableEl.style.height = `${totalRows * ROW_HEIGHT_PX}px`; // 스크롤 가능한 전체 높이 확보(뷰어의 translateY 트릭과 동일 원리)

    const scrollTop = scrollHostEl.scrollTop;
    const viewportHeight = scrollHostEl.clientHeight;
    const firstVisible = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT_PX) - OVERSCAN_ROWS);
    const lastVisible = Math.min(
      totalRows - 1,
      Math.ceil((scrollTop + viewportHeight) / ROW_HEIGHT_PX) + OVERSCAN_ROWS,
    );

    // ... 이하 기존 renderTable의 <thead>/<tbody> 생성 로직을 그대로 재사용하되,
    // `for (let r = 0; r < renderRowCount; r += 1)` 루프를
    // `for (let r = firstVisible; r <= lastVisible; r += 1)` 로 바꾸고,
    // 각 행의 <tr>에 `tr.style.position = 'absolute'; tr.style.top = `${r * ROW_HEIGHT_PX}px`;`
    // 를 추가해(기존 sticky용 position:relative 설정과 함께 — freeze 행은 여전히 sticky 유지,
    // 나머지 행만 absolute) 스크롤 위치와 무관하게 올바른 자리에 배치한다.
    // (기존 renderTable의 셀 생성/이벤트 바인딩 코드는 전부 그대로 유지 — 루프 범위와 위치
    // 지정 방식만 바뀐다. 구현자는 기존 renderTable 본문을 이 함수 안으로 옮기고 위 두 가지만
    // 수정할 것.)
  }
```

`renderTable`을 호출하던 모든 곳(초기 `provider.on('sync', ...)`, 시트 전환, 구조 변경 감지
등)을 `renderTableVirtualized`로 바꾼다(정확한 호출부는 `renderTable(` grep으로 전부 찾아
치환).

**스크롤 리스너 추가** — 기존 `scrollHostEl.addEventListener('scroll', updateHScrollbar);`
(`:2354`) 옆에 추가:
```ts
  let scrollRenderScheduled = false;
  scrollHostEl.addEventListener('scroll', () => {
    if (scrollRenderScheduled) return;
    scrollRenderScheduled = true;
    requestAnimationFrame(() => {
      scrollRenderScheduled = false;
      renderTableVirtualized(currentSheetStructure); // currentSheetStructure는 activateSheet 등에서 이미 추적 중인 값 — 정확한 변수명은 기존 코드 확인
    });
    scheduleSubscriptionUpdate();
  });
```

- [ ] **Step 2: 페이지 구독 전송**

```ts
  const MESSAGE_PAGE_SUBSCRIBE = 2;
  let subscriptionDebounceTimer: ReturnType<typeof setTimeout> | null = null;

  function sendPageSubscription(sheetIdx: number, pages: readonly number[]): void {
    const encoder = lib0encoding.createEncoder(); // 파일 상단에 이미 lib0/encoding을 어떤 이름으로 import했는지 확인 후 맞출 것(y-websocket 내부에서만 쓰였다면 새로 import 필요: `import * as encoding from 'lib0/encoding';`)
    encoding.writeVarUint(encoder, MESSAGE_PAGE_SUBSCRIBE);
    encoding.writeVarString(encoder, JSON.stringify({ sheet: sheetIdx, pages }));
    // WebsocketProvider는 원시 소켓을 직접 노출하지 않을 수 있다 — provider.ws가 실제 WebSocket
    // 인스턴스인지(y-websocket 라이브러리 구현 확인 필요) 확인하고, 있으면
    // provider.ws?.send(encoding.toUint8Array(encoder)); 로 보낸다. 없으면 provider의 내부
    // 소켓 접근 방법을 y-websocket 타입 선언(node_modules/y-websocket)에서 확인해 맞는
    // 방법으로 전송한다 — 이 부분은 구현자가 실제 라이브러리 API를 확인 후 정확한 코드를
    // 작성해야 한다(플레이스홀더 아님, 실제 조사 필요).
  }

  function scheduleSubscriptionUpdate(): void {
    if (subscriptionDebounceTimer) clearTimeout(subscriptionDebounceTimer);
    subscriptionDebounceTimer = setTimeout(() => {
      subscriptionDebounceTimer = null;
      const scrollTop = scrollHostEl.scrollTop;
      const viewportHeight = scrollHostEl.clientHeight;
      const firstRow = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT_PX) - OVERSCAN_ROWS);
      const lastRow = Math.ceil((scrollTop + viewportHeight) / ROW_HEIGHT_PX) + OVERSCAN_ROWS;
      const firstPage = Math.floor(firstRow / PAGE_SIZE_ROWS);
      const lastPage = Math.floor(lastRow / PAGE_SIZE_ROWS);
      const pages: number[] = [];
      for (let p = firstPage; p <= lastPage; p += 1) pages.push(p);
      sendPageSubscription(activeSheetIndex, pages); // activeSheetIndex는 기존 코드에서 활성 시트를 추적하는 변수 — 정확한 이름은 기존 코드 확인
    }, 200);
  }
```
시트 전환(기존 `activateSheet` 함수) 끝에도 `scheduleSubscriptionUpdate()`를 즉시 호출하도록
추가(디바운스 없이 — 시트를 바꾼 직후엔 바로 새 시트의 페이지를 알려야 한다. 정확히는:
시트 전환은 디바운스를 건너뛰고 즉시 전송해도 되지만, 최소한 스펙과 일관되게 동작하도록
`scheduleSubscriptionUpdate()`를 호출하고 200ms 기다리는 것으로 충분하다 — 구현자 판단).

초기 로드 시(provider `'sync'` 이벤트 핸들러 안, `renderTableVirtualized` 최초 호출 직후)에도
`scheduleSubscriptionUpdate()`를 호출해 첫 화면의 페이지를 구독한다.

- [ ] **Step 3: 찾기/바꾸기 범위 경고**

찾기/바꾸기 UI 관련 코드(`:2581` 근처, 정확한 UI 요소는 기존 코드에서 검색/치환 패널을 여는
함수를 찾아 확인)에 안내 문구를 추가한다 — 검색 패널이 열릴 때 한 번 표시되는 작은 안내
텍스트: `"현재 로드된 범위에서만 검색합니다."` (정확한 DOM 삽입 위치는 기존 검색 패널 HTML
구조를 확인해 구현자가 결정 — `public/editor-grid.html`의 검색 UI 마크업 확인 필요).

- [ ] **Step 4: 빌드 + 타입 체크**

Run: `npm run build:editor && npx tsc --noEmit`
Expected: 둘 다 성공.

- [ ] **Step 5: 실제 브라우저 확인 (필수 — 이 프로젝트 관례)**

`npm run dev`로 서버를 띄우고, 3000행 이상짜리 실제(또는 합성) xlsx를 업로드해 동시편집기로
연다. 확인 항목:
- 여는 즉시(전체 로딩 없이) 빠르게 화면이 뜨는지
- 스크롤하면 데이터가 자연스럽게 나타나는지(빈 화면이 오래 보이지 않는지)
- 수식 셀이 정확한 값을 보여주는지(특히 화면 밖 참조가 있는 수식)
- 셀을 편집하고 즉시 다운로드했을 때 그 편집이 반영돼 있는지
- 두 브라우저 탭을 열어 서로 다른 위치를 보면서 동시에 편집해도 각자 정상 동작하는지

- [ ] **Step 6: 커밋**

```bash
git add src/client/grid-editor.ts public/editor-grid.html
git commit -m "feat: 그리드 편집기 클라이언트에 가상 스크롤 + 페이지 구독 전송 추가"
```

---

## 최종 확인

모든 태스크가 끝나면:

```bash
npx tsc --noEmit && npx vitest run && npm run build:editor
```

모두 초록불이어야 한다. 이 계획은 이번 세션에서 가장 리스크가 큰 변경이므로, 최종
전체 브랜치 리뷰(subagent-driven-development의 final review)에서 특히 다음을 중점적으로
확인할 것: (1) canonical과 Yjs 사이의 데이터 유실 가능성이 정말 없는지(편집 즉시 반영이
모든 경로에서 실제로 지켜지는지), (2) 페이지 재구독/해제 경계에서의 레이스 컨디션,
(3) 수식 재계산 루프가 정말 이벤트 루프를 막지 않는지, (4) 대형 파일에 대한 실제 브라우저
확인이 Task 9에서 실제로 이뤄졌는지.
