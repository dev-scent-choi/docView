import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import {
  messageYjsSyncStep2,
  messageYjsUpdate,
  readUpdate,
  writeSyncStep1,
  writeUpdate,
} from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

const MESSAGE_SYNC = 0;
const MESSAGE_PAGE_SUBSCRIBE = 2;
const MESSAGE_GRID_STRUCTURE_OP = 4;

/**
 * 이 파일이 붙잡는 회귀(최종 전체 브랜치 리뷰의 C1/C2/I2):
 * - C1: 값은 그대로 두고 **서식만** 바꾼 편집이 canonical에 반영되지 않아 저장된 파일에서 사라진다.
 * - C2: 원본 시트 범위 **밖**의 편집과 모든 구조 변경(시트 이름/추가/삭제/순서, 행·열 삽입)이
 *   canonical에 반영되지 않아 저장된 파일에서 사라진다.
 * - I2: MESSAGE_PAGE_SUBSCRIBE가 개수/정수 검증 없이 아무 페이지 목록이나 받아들인다.
 *
 * 전부 "Yjs에는 보이고 협업 상대에게도 전파되지만, 다운로드한 파일에는 없다"는 모양의 조용한
 * 데이터 손실이라, 검증은 반드시 **실제 내보낸 바이트를 다시 열어보는** 것으로 한다.
 */

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

function sendSubscribe(socket: FakeSocket, payload: unknown): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_PAGE_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify(payload));
  socket.emit('message', encoding.toUint8Array(encoder));
}

/**
 * 행/열 삽입·삭제는 이제 클라이언트가 Yjs 키를 직접 옮기지 않고 이 메시지만 보낸다 — 서버가
 * canonical 그리드에서 옮긴다(room.ts의 applyGridStructureOp). 재설계의 이유와 이 경로의 본 시험대는
 * tests/roomRowColShift.test.ts 참고(1000행이 넘는 시트, 구독되지 않은 페이지까지 포함).
 */
function sendStructureOp(socket: FakeSocket, payload: unknown): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_GRID_STRUCTURE_OP);
  encoding.writeVarString(encoder, JSON.stringify(payload));
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** 행/열 연산은 canonical 갱신 체인에서 양보하며 돌므로 체인이 실제로 비는 것을 기다린다. */
async function settleStructureOp(docId: string): Promise<void> {
  const { awaitCanonicalGrowthForTest } = await import('../src/collab/room.js');
  await awaitCanonicalGrowthForTest(docId);
  await flushEventLoop();
  await awaitCanonicalGrowthForTest(docId);
}

/** 페이지 채우기/재계산은 청크마다 실제 setImmediate로 양보하므로 그 턴을 흘려보낸다. */
async function flushEventLoop(turns = 60): Promise<void> {
  for (let i = 0; i < turns; i += 1) await new Promise((resolve) => setImmediate(resolve));
}

/**
 * 실제 클라이언트가 접속 직후 하는 것과 같은 동기화 — 이걸 먼저 해야 mirror가 룸과 같은 항목
 * 식별자를 공유하게 되어, mirror에서 만든 편집/삭제가 룸의 실제 항목에 정확히 병합된다
 * (roomCanonicalGrid.test.ts의 같은 헬퍼와 동일한 이유).
 */
function syncMirrorWithRoom(socket: FakeSocket): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeSyncStep1(encoder, socket.mirror);
  socket.emit('message', encoding.toUint8Array(encoder));
}

/**
 * mirror(= 이 클라이언트의 Y.Doc)에서 **한 트랜잭션 안에** 편집을 수행하고, 그 결과 생긴 update를
 * 서버로 보낸다. 클라이언트 코드(grid-editor.ts)가 실제로 하는 것과 같은 모양이다 — 특히 행/열
 * 삽입은 "셀 키 이동 + structure의 행/열 수 변경"을 반드시 한 트랜잭션 안에서 하므로, 그 순서·
 * 원자성까지 그대로 재현해야 서버 관찰자들의 처리 순서를 진짜로 검증할 수 있다.
 */
function applyOnMirror(socket: FakeSocket, fn: (doc: Y.Doc) => void): void {
  let captured: Uint8Array | null = null;
  const onUpdate = (update: Uint8Array) => {
    captured = captured === null ? update : Y.mergeUpdates([captured, update]);
  };
  socket.mirror.on('update', onUpdate);
  socket.mirror.transact(() => fn(socket.mirror));
  socket.mirror.off('update', onUpdate);
  if (captured === null) return;
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeUpdate(encoder, captured);
  socket.emit('message', encoding.toUint8Array(encoder));
}

async function makeSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 'Name';
  sheet.getCell('B1').value = 'Score';
  sheet.getCell('A2').value = 'Alice';
  sheet.getCell('B2').value = 10;
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

async function makeTwoSheetXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const first = workbook.addWorksheet('First');
  first.getCell('A1').value = 'first-a1';
  const second = workbook.addWorksheet('Second');
  second.getCell('A1').value = 'second-a1';
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

/**
 * 데이터가 **삽입 지점 위아래 양쪽에** 있고, 중간에 빈 칸(구멍)이 섞인 시트. 가장자리 삽입만으로는
 * 절대 드러나지 않는 "행 밀기"의 실패 모드를 겨냥한 픽스처다.
 *
 *      A          B
 *   1  Header     H2
 *   2  alpha      1
 *   3  (빈 칸)     2     ← A열에만 구멍
 *   4  omega      3
 *
 * 시딩은 빈 칸을 GRID_CELLS_MAP에 **아예 넣지 않으므로**(seedGridPageFromXlsx: `cell.v !== ''`),
 * 이 구멍은 Yjs에서 "키 없음"으로 나타난다 — 밀기 로직이 "원본이 비어 있으면 목적지도 비워야 한다"를
 * 빠뜨리면 그 자리에 직전 값이 그대로 남아 데이터가 복제된다.
 */
async function makeSparseMiddleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 'Header';
  sheet.getCell('B1').value = 'H2';
  sheet.getCell('A2').value = 'alpha';
  sheet.getCell('B2').value = 1;
  /* A3은 일부러 비워 둔다 */ sheet.getCell('B3').value = 2;
  sheet.getCell('A4').value = 'omega';
  sheet.getCell('B4').value = 3;
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

async function seedDoc(docId: string, original: Buffer): Promise<void> {
  const { store } = await import('../src/storage.js');
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
}

/** 지금 룸이 내보낼 바이트를 실제로 만들어 다시 ExcelJS로 연다 — "다운로드한 파일"과 동일한 경로. */
async function downloadWorkbook(docId: string): Promise<ExcelJS.Workbook> {
  const { getCurrentBytes } = await import('../src/collab/room.js');
  const bytes = await getCurrentBytes(docId);
  expect(bytes).not.toBeNull();
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(bytes! as unknown as ArrayBuffer);
  return reopened;
}

describe('room.ts: style-only and structural edits reach the canonical grid (and the saved file)', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-structure-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.useRealTimers();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  async function openRoom(
    original: Buffer,
    pages: readonly number[] = [0],
  ): Promise<{ docId: string; socket: FakeSocket }> {
    const { joinEditSession } = await import('../src/collab/room.js');
    const docId = randomUUID();
    await seedDoc(docId, original);
    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
    sendSubscribe(socket, { sheet: 0, pages });
    await flushEventLoop();
    syncMirrorWithRoom(socket);
    return { docId, socket };
  }

  // ---------------------------------------------------------------- C1
  test('a style-only edit (no text change) survives a round-trip through the saved file', async () => {
    const { GRID_STYLES_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
    const { docId, socket } = await openRoom(await makeSampleXlsx());

    // 사용자가 A1에 굵게 + 배경색만 적용한다 — 셀 텍스트는 한 글자도 바뀌지 않는다.
    applyOnMirror(socket, (doc) => {
      doc.getMap(GRID_STYLES_MAP).set(gridCellKey(0, 0, 0), { b: 1, bg: '#FFFF00' });
    });
    await flushEventLoop();

    const workbook = await downloadWorkbook(docId);
    const cell = workbook.getWorksheet('Sheet1')!.getCell('A1');
    expect(cell.value).toBe('Name'); // 값은 그대로
    expect(cell.font?.bold).toBe(true);
    expect((cell.fill as ExcelJS.FillPattern | undefined)?.fgColor?.argb).toBe('FFFFFF00');
  });

  test('removing a style (Y.Map key delete) also reaches the saved file', async () => {
    const workbookIn = new ExcelJS.Workbook();
    const sheetIn = workbookIn.addWorksheet('Sheet1');
    sheetIn.getCell('A1').value = 'Bold';
    sheetIn.getCell('A1').font = { bold: true };
    const original = Buffer.from((await workbookIn.xlsx.writeBuffer()) as ArrayBuffer);

    const { GRID_STYLES_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
    const { docId, socket } = await openRoom(original);

    applyOnMirror(socket, (doc) => {
      doc.getMap(GRID_STYLES_MAP).delete(gridCellKey(0, 0, 0));
    });
    await flushEventLoop();

    const workbook = await downloadWorkbook(docId);
    const cell = workbook.getWorksheet('Sheet1')!.getCell('A1');
    expect(cell.value).toBe('Bold');
    expect(cell.font?.bold).toBeFalsy();
  });

  test('a style-only edit does not clobber the cell text in the canonical grid', async () => {
    const { getRoomLazyGridForTest } = await import('../src/collab/room.js');
    const { getCell } = await import('../src/xlsx/lazyGridStore.js');
    const { GRID_STYLES_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
    const { docId, socket } = await openRoom(await makeSampleXlsx());

    applyOnMirror(socket, (doc) => {
      doc.getMap(GRID_STYLES_MAP).set(gridCellKey(0, 1, 0), { i: 1 });
    });
    await flushEventLoop();

    const lazyGrid = getRoomLazyGridForTest(docId)!;
    const cell = await getCell(lazyGrid, 0, 1, 0);
    expect(cell).not.toBe(0);
    expect((cell as { v: string }).v).toBe('Alice');
    expect((cell as { s?: number }).s).toBeDefined();
  });

  // ---------------------------------------------------------------- C2
  test('a cell written at exactly col === colCount (past the original extent) reaches the saved file', async () => {
    const { GRID_CELLS_MAP, GRID_STRUCTURE_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
    const { docId, socket } = await openRoom(await makeSampleXlsx()); // 2열(A,B)짜리 시트

    // 클라이언트(commitCell)가 원본 범위 밖 칸에 입력할 때 하는 것과 정확히 같은 한 트랜잭션:
    // 셀 값을 쓰고, 같은 트랜잭션에서 structure의 colCount를 넓힌다.
    applyOnMirror(socket, (doc) => {
      const structure = doc.getMap<{ name: string; rowCount: number; colCount: number }>(GRID_STRUCTURE_MAP);
      const current = structure.get('0')!;
      doc.getMap<string>(GRID_CELLS_MAP).set(gridCellKey(0, 0, 2), 'PastEdge');
      structure.set('0', { ...current, colCount: Math.max(current.colCount, 3) });
    });
    await flushEventLoop();

    const workbook = await downloadWorkbook(docId);
    expect(workbook.getWorksheet('Sheet1')!.getCell('C1').value).toBe('PastEdge');
  });

  test('a cell written past the last row reaches the saved file', async () => {
    const { GRID_CELLS_MAP, GRID_STRUCTURE_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
    const { docId, socket } = await openRoom(await makeSampleXlsx()); // 2행짜리 시트

    applyOnMirror(socket, (doc) => {
      const structure = doc.getMap<{ name: string; rowCount: number; colCount: number }>(GRID_STRUCTURE_MAP);
      const current = structure.get('0')!;
      doc.getMap<string>(GRID_CELLS_MAP).set(gridCellKey(0, 4, 0), 'Row5');
      structure.set('0', { ...current, rowCount: Math.max(current.rowCount, 5) });
    });
    await flushEventLoop();

    const workbook = await downloadWorkbook(docId);
    expect(workbook.getWorksheet('Sheet1')!.getCell('A5').value).toBe('Row5');
  });

  test('a sheet rename survives the round-trip', async () => {
    const { GRID_STRUCTURE_MAP } = await import('../src/collab/gridShared.js');
    const { docId, socket } = await openRoom(await makeSampleXlsx());

    applyOnMirror(socket, (doc) => {
      const structure = doc.getMap<{ name: string }>(GRID_STRUCTURE_MAP);
      structure.set('0', { ...structure.get('0')!, name: '이름바꾼시트' });
    });
    await flushEventLoop();

    const workbook = await downloadWorkbook(docId);
    expect(workbook.worksheets.map((s) => s.name)).toContain('이름바꾼시트');
    expect(workbook.getWorksheet('이름바꾼시트')!.getCell('A1').value).toBe('Name');
  });

  test('adding a sheet and typing into it survives the round-trip', async () => {
    const { GRID_CELLS_MAP, GRID_STRUCTURE_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
    const { docId, socket } = await openRoom(await makeSampleXlsx());

    // addSheet(grid-editor.ts)가 하는 것과 같은 모양 — 새 키에 최소 구조를 쓴다.
    applyOnMirror(socket, (doc) => {
      doc
        .getMap<{ name: string; rowCount: number; colCount: number; order?: number }>(GRID_STRUCTURE_MAP)
        .set('1', { name: 'Sheet2', rowCount: 1, colCount: 1, order: 1 });
    });
    await flushEventLoop();
    // 새 시트에 입력 — 클라이언트는 시트를 전환하고 그 시트 페이지를 구독한 뒤 입력한다.
    sendSubscribe(socket, { sheet: 1, pages: [0] });
    await flushEventLoop();
    applyOnMirror(socket, (doc) => {
      const structure = doc.getMap<{ name: string; rowCount: number; colCount: number }>(GRID_STRUCTURE_MAP);
      const current = structure.get('1')!;
      doc.getMap<string>(GRID_CELLS_MAP).set(gridCellKey(1, 0, 0), 'brand-new');
      structure.set('1', { ...current, rowCount: Math.max(current.rowCount, 1), colCount: Math.max(current.colCount, 1) });
    });
    await flushEventLoop();

    const workbook = await downloadWorkbook(docId);
    const added = workbook.getWorksheet('Sheet2');
    expect(added).toBeDefined();
    expect(added!.getCell('A1').value).toBe('brand-new');
    // 원래 시트도 그대로 남아 있어야 한다.
    expect(workbook.getWorksheet('Sheet1')!.getCell('A1').value).toBe('Name');
  });

  test('deleting a sheet removes it from the saved file (and keeps the other one intact)', async () => {
    const { GRID_STRUCTURE_MAP } = await import('../src/collab/gridShared.js');
    const { docId, socket } = await openRoom(await makeTwoSheetXlsx());

    applyOnMirror(socket, (doc) => {
      doc.getMap(GRID_STRUCTURE_MAP).delete('0'); // 첫 시트를 지운다(꼬리가 아닌 중간 삭제)
    });
    await flushEventLoop();

    const workbook = await downloadWorkbook(docId);
    expect(workbook.worksheets.map((s) => s.name)).toEqual(['Second']);
    expect(workbook.getWorksheet('Second')!.getCell('A1').value).toBe('second-a1');
  });

  test('reordering sheets (order field) is reflected in the saved file', async () => {
    const { GRID_STRUCTURE_MAP } = await import('../src/collab/gridShared.js');
    const { docId, socket } = await openRoom(await makeTwoSheetXlsx());

    applyOnMirror(socket, (doc) => {
      const structure = doc.getMap<{ name: string; order?: number }>(GRID_STRUCTURE_MAP);
      structure.set('0', { ...structure.get('0')!, order: 1 });
      structure.set('1', { ...structure.get('1')!, order: 0 });
    });
    await flushEventLoop();

    const workbook = await downloadWorkbook(docId);
    expect(workbook.worksheets.map((s) => s.name)).toEqual(['Second', 'First']);
  });

  test('inserting a row shifts the data down in the saved file', async () => {
    const { docId, socket } = await openRoom(await makeSampleXlsx());

    sendStructureOp(socket, { sheet: 0, kind: 'insertRows', at: 0, count: 1 });
    await settleStructureOp(docId);

    const workbook = await downloadWorkbook(docId);
    const sheet = workbook.getWorksheet('Sheet1')!;
    expect(sheet.getCell('A1').value).toBeFalsy(); // 새로 삽입된 빈 행
    expect(sheet.getCell('A2').value).toBe('Name');
    expect(sheet.getCell('A3').value).toBe('Alice');
    expect(sheet.getCell('B3').value).toBe(10);
  });

  test('inserting a column shifts the data right in the saved file', async () => {
    const { docId, socket } = await openRoom(await makeSampleXlsx());

    sendStructureOp(socket, { sheet: 0, kind: 'insertCols', at: 0, count: 1 });
    await settleStructureOp(docId);

    const workbook = await downloadWorkbook(docId);
    const sheet = workbook.getWorksheet('Sheet1')!;
    expect(sheet.getCell('A1').value).toBeFalsy();
    expect(sheet.getCell('B1').value).toBe('Name');
    expect(sheet.getCell('C1').value).toBe('Score');
    expect(sheet.getCell('B2').value).toBe('Alice');
  });

  test('inserting a row in the MIDDLE of a sheet with data on both sides (and a gap) shifts exactly once', async () => {
    const { docId, socket } = await openRoom(await makeSparseMiddleXlsx()); // 4행 × 2열, A3이 빈 칸

    // insertRowAt(2) — 위(1,2행)에도 아래(3,4행)에도 데이터가 있는 진짜 "가운데" 삽입.
    // 구멍(A3)이 섞인 픽스처를 쓰는 이유: 성긴 셀 맵에서 빈 칸이 목적지에 반영되지 않으면 그
    // 자리에 직전 값이 남아 데이터가 복제된다 — 서버 재시딩 경로에서도 같은 실패가 가능하다.
    sendStructureOp(socket, { sheet: 0, kind: 'insertRows', at: 2, count: 1 });
    await settleStructureOp(docId);

    const sheet = (await downloadWorkbook(docId)).getWorksheet('Sheet1')!;
    expect(sheet.getCell('A1').value).toBe('Header'); // 삽입 지점 위 — 그대로
    expect(sheet.getCell('A2').value).toBe('alpha');
    expect(sheet.getCell('B2').value).toBe(1);
    expect(sheet.getCell('A3').value).toBeFalsy(); // 새로 삽입된 빈 행
    expect(sheet.getCell('B3').value).toBeFalsy();
    // 원본의 "구멍 있던 행"이 한 칸 내려온 자리 — A는 여전히 비어 있어야 한다.
    // 밀기가 빈 원본을 목적지에 반영하지 않으면 여기에 'omega'가 복제되어 남는다.
    expect(sheet.getCell('A4').value).toBeFalsy();
    expect(sheet.getCell('B4').value).toBe(2);
    expect(sheet.getCell('A5').value).toBe('omega'); // 아래쪽 데이터 — 정확히 한 칸만 내려온다
    expect(sheet.getCell('B5').value).toBe(3);
    expect(sheet.getCell('A6').value).toBeFalsy(); // 두 칸 밀리지 않았는지
  });

  test('inserting a column in the MIDDLE of a sheet with data on both sides (and a gap) shifts exactly once', async () => {
    // 열 방향으로 같은 모양을 만든다: 1행은 A,B,(빈 C),D / 2행은 전부 채움.
    const workbookIn = new ExcelJS.Workbook();
    const sheetIn = workbookIn.addWorksheet('Sheet1');
    sheetIn.getCell('A1').value = 'a1';
    sheetIn.getCell('B1').value = 'b1';
    /* C1은 비워 둔다 */ sheetIn.getCell('D1').value = 'd1';
    sheetIn.getCell('A2').value = 'a2';
    sheetIn.getCell('B2').value = 'b2';
    sheetIn.getCell('C2').value = 'c2';
    sheetIn.getCell('D2').value = 'd2';
    const original = Buffer.from((await workbookIn.xlsx.writeBuffer()) as ArrayBuffer);

    const { docId, socket } = await openRoom(original);

    // insertColumnAt(2) — B(왼쪽)와 D(오른쪽) 사이의 진짜 "가운데" 삽입.
    sendStructureOp(socket, { sheet: 0, kind: 'insertCols', at: 2, count: 1 });
    await settleStructureOp(docId);

    const sheet = (await downloadWorkbook(docId)).getWorksheet('Sheet1')!;
    expect(sheet.getCell('A1').value).toBe('a1'); // 삽입 지점 왼쪽 — 그대로
    expect(sheet.getCell('B1').value).toBe('b1');
    expect(sheet.getCell('C1').value).toBeFalsy(); // 새로 삽입된 빈 열
    expect(sheet.getCell('D1').value).toBeFalsy(); // 원본 C1(빈 칸)이 내려온 자리 — 'd1'이 복제되면 안 된다
    expect(sheet.getCell('E1').value).toBe('d1'); // 오른쪽 데이터 — 정확히 한 칸만 이동
    expect(sheet.getCell('A2').value).toBe('a2');
    expect(sheet.getCell('C2').value).toBeFalsy();
    expect(sheet.getCell('D2').value).toBe('c2');
    expect(sheet.getCell('E2').value).toBe('d2');
  });

  /**
   * 예전 이 자리에는 클라이언트 밀기 루프의 **하드코딩된 상한 1000** 주변을 겨냥한 테스트 두
   * 개가 있었다(경계 행 1000의 데이터를 파괴하는 회귀, 그리고 그 열 짝). 그 루프 자체가 사라졌고
   * — 애초에 그 상한이 이번 재설계의 원인이다 — 같은 실패 모드는 이제 서버 경로에서, 훨씬 강한
   * 형태로 검증한다: tests/roomRowColShift.test.ts의 1600행 시트 테스트는 행 999·1000·1500을
   * 한꺼번에 확인하고, 그 행들이 **Yjs에 아예 올라와 있지 않은 상태**(페이지 0만 구독)에서도
   * 정확히 밀리는지까지 본다.
   */
  test('deleting a row (rowCount shrink) removes the tail row from the saved file', async () => {
    const { docId, socket } = await openRoom(await makeSampleXlsx());

    sendStructureOp(socket, { sheet: 0, kind: 'deleteRows', at: 0, count: 1 });
    await settleStructureOp(docId);

    const workbook = await downloadWorkbook(docId);
    const sheet = workbook.getWorksheet('Sheet1')!;
    expect(sheet.getCell('A1').value).toBe('Alice');
    expect(sheet.getCell('B1').value).toBe(10);
    expect(sheet.getCell('A2').value).toBeFalsy();
  });

  // ---------------------------------------------------------------- I2
  test('MESSAGE_PAGE_SUBSCRIBE rejects an oversized pages array and non-integer/negative page numbers', async () => {
    const { getSubscribedPageUnionForTest } = await import('../src/collab/room.js');
    const { joinEditSession } = await import('../src/collab/room.js');

    const docId = randomUUID();
    await seedDoc(docId, await makeSampleXlsx());
    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);

    sendSubscribe(socket, { sheet: 0, pages: Array.from({ length: 5000 }, (_, i) => i) });
    await flushEventLoop(10);
    expect(getSubscribedPageUnionForTest(docId, 0).size).toBe(0);

    sendSubscribe(socket, { sheet: 0, pages: [0.5] });
    await flushEventLoop(10);
    expect(getSubscribedPageUnionForTest(docId, 0).size).toBe(0);

    sendSubscribe(socket, { sheet: 0, pages: [-1] });
    await flushEventLoop(10);
    expect(getSubscribedPageUnionForTest(docId, 0).size).toBe(0);

    sendSubscribe(socket, { sheet: 0, pages: [10_000_000] });
    await flushEventLoop(10);
    expect(getSubscribedPageUnionForTest(docId, 0).size).toBe(0);

    sendSubscribe(socket, { sheet: -1, pages: [0] });
    await flushEventLoop(10);
    expect(getSubscribedPageUnionForTest(docId, -1).size).toBe(0);

    // 정상 구독은 그대로 받아들여야 한다 — 검증이 필요 이상으로 넓어지지 않았는지 함께 확인.
    sendSubscribe(socket, { sheet: 0, pages: [0, 1] });
    await flushEventLoop(10);
    expect([...getSubscribedPageUnionForTest(docId, 0)].sort()).toEqual([0, 1]);
  });
}, 120_000);
