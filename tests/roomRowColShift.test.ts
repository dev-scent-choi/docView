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

/**
 * 이 파일이 붙잡는 회귀 — **이 계획 전체보다 오래된, 진짜 데이터 손상 결함**:
 *
 * 예전에는 행/열 삽입·삭제를 클라이언트(grid-editor.ts의 shiftRowsAndMerges/shiftColsAndMerges)가
 * Yjs의 cells/styles 키를 직접 한 칸씩 옮겨서 처리했다. 두 가지가 겹쳐 조용한 데이터 손상이 났다:
 *
 * 1. 그 루프의 상한이 시트 실제 크기가 아니라 **하드코딩된 1000**이었다(1000행 넘는 시트를
 *    지원하기 전 시대의 코드). 1001행이 넘는 시트에서 행을 하나 삽입하면 행 1000의 내용이
 *    덮어써지고, **행 1001 이후는 아예 밀리지 않아** 시트 나머지 전체가 위쪽과 한 칸씩 어긋났다.
 * 2. 그 접근 자체가 "시트의 모든 셀이 Yjs에 있다"를 전제하는데, 이 프로젝트가 Yjs를 "지금 구독
 *    중인 500행 페이지"로 좁힌 뒤로 그 전제가 깨졌다 — 창 밖의 행은 애초에 옮길 수가 없다.
 *
 * 그래서 이제 클라이언트는 "이 자리에 행을 넣어 달라"는 메시지(MESSAGE_GRID_STRUCTURE_OP)만 보내고,
 * 실제 이동은 서버가 **항상 시트 전체를 담고 있는** canonical 그리드에서 한다.
 *
 * 검증은 반드시 **실제로 내보낸 바이트를 다시 열어보는** 것으로 한다 — 이 버그 클래스는 전부
 * "화면에는 맞게 보이는데 다운로드한 파일에는 없다/어긋나 있다" 모양이기 때문이다.
 */

const MESSAGE_SYNC = 0;
const MESSAGE_PAGE_SUBSCRIBE = 2;
const MESSAGE_GRID_STRUCTURE_OP = 4;

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  readonly mirror = new Y.Doc();

  send(data: Uint8Array, cb?: (err?: Error) => void): void {
    const decoder = decoding.createDecoder(data);
    if (decoding.readVarUint(decoder) === MESSAGE_SYNC) {
      const syncType = decoding.readVarUint(decoder);
      if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) readUpdate(decoder, this.mirror, this);
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

/** 클라이언트의 sendGridStructureOp와 완전히 같은 인코딩(varUint 메시지 타입 + varString JSON). */
function sendStructureOp(socket: FakeSocket, payload: unknown): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_GRID_STRUCTURE_OP);
  encoding.writeVarString(encoder, JSON.stringify(payload));
  socket.emit('message', encoding.toUint8Array(encoder));
}

async function flushEventLoop(turns = 80): Promise<void> {
  for (let i = 0; i < turns; i += 1) await new Promise((resolve) => setImmediate(resolve));
}

function syncMirrorWithRoom(socket: FakeSocket): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeSyncStep1(encoder, socket.mirror);
  socket.emit('message', encoding.toUint8Array(encoder));
}

/**
 * 행/열 연산은 canonical 갱신 체인에서 돌고(enqueueCanonicalGrowth) 그 안에서 페이지 재시딩·수식
 * 재계산이 다시 청크마다 양보하므로, 고정된 틱 수만 흘려보내는 것으로는 완료를 보장할 수 없다 —
 * 체인이 실제로 비는 것을 기다린 뒤 남은 후속 작업까지 한 번 더 흘려보낸다.
 */
async function settleStructureOp(docId: string): Promise<void> {
  const { awaitCanonicalGrowthForTest } = await import('../src/collab/room.js');
  await awaitCanonicalGrowthForTest(docId);
  await flushEventLoop();
  await awaitCanonicalGrowthForTest(docId);
}

/**
 * 1600행짜리 시트 — 이 파일의 주력 픽스처.
 *
 *   B열: 1~1600행 전부 'b'  ← parseXlsxToGrid의 rowCount는 `actualRowCount`(= **값이 있는 행의
 *        개수**)라, 성긴 시트를 쓰면 시트가 몇 줄짜리로 잘려 경계 행 자체가 사라진다.
 *   A열 표식(괄호는 0-based):
 *     A1    (0)    'top'      ← 삽입 지점 위 — 절대 움직이면 안 된다
 *     A501  (500)  'mid500'   ← 삽입 지점 아래, 옛 루프 상한 안쪽
 *     A1000 (999)  'row999'   ← 옛 루프의 마지막 원본 행
 *     A1001 (1000) 'row1000'  ← **옛 코드가 덮어써서 날려버린 행**
 *     A1501 (1500) 'row1500'  ← **옛 코드가 아예 밀지 않아 어긋난 행**
 */
async function makeTallXlsx(rows = 1600): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  for (let excelRow = 1; excelRow <= rows; excelRow += 1) sheet.getCell(excelRow, 2).value = 'b';
  sheet.getCell('A1').value = 'top';
  sheet.getCell('A501').value = 'mid500';
  sheet.getCell('A1000').value = 'row999';
  sheet.getCell('A1001').value = 'row1000';
  sheet.getCell('A1501').value = 'row1500';
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

async function downloadWorkbook(docId: string): Promise<ExcelJS.Workbook> {
  const { getCurrentBytes } = await import('../src/collab/room.js');
  const bytes = await getCurrentBytes(docId);
  expect(bytes).not.toBeNull();
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(bytes! as unknown as ArrayBuffer);
  return reopened;
}

/** 한 열에서 그 표식이 몇 번 나타나는지 — "복제되지 않았다"를 직접 확인하기 위한 헬퍼. */
function countInColumn(sheet: ExcelJS.Worksheet, col: number, marker: string): number {
  let found = 0;
  for (let r = 1; r <= sheet.rowCount; r += 1) {
    if (sheet.getCell(r, col).value === marker) found += 1;
  }
  return found;
}

describe('room.ts: server-side row/column insert and delete (canonical shift)', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-rowcol-shift-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
    // setImmediate는 반드시 real로 남긴다 — ExcelJS의 zip/stream 내부가 진짜 setImmediate 턴에
    // 의존해서, 전부 fake로 두면 내보내기가 영원히 끝나지 않는다(계획의 Global Constraints).
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
    readOnly = false,
  ): Promise<{ docId: string; socket: FakeSocket }> {
    const { joinEditSession } = await import('../src/collab/room.js');
    const docId = randomUUID();
    await seedDoc(docId, original);
    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], readOnly);
    sendSubscribe(socket, { sheet: 0, pages });
    await flushEventLoop();
    syncMirrorWithRoom(socket);
    return { docId, socket };
  }

  // ------------------------------------------------------------------ 원래의 실패 시나리오
  test('inserting a row on a 1600-row sheet shifts EVERY row below it — including rows 1000+ that are not loaded into Yjs at all', async () => {
    // 페이지 0(행 0~499)만 구독한다 — 행 999/1000/1500은 Yjs에 **한 글자도 올라와 있지 않다.**
    // 옛 클라이언트 밀기 방식으로는 애초에 옮길 수 없었던 바로 그 영역이다.
    const { docId, socket } = await openRoom(await makeTallXlsx(), [0]);

    sendStructureOp(socket, { sheet: 0, kind: 'insertRows', at: 100, count: 1 });
    await settleStructureOp(docId);

    const sheet = (await downloadWorkbook(docId)).getWorksheet('Sheet1')!;
    expect(sheet.getCell('A1').value).toBe('top'); // 삽입 지점 위 — 그대로
    expect(sheet.getCell('A101').value).toBeFalsy(); // 새로 끼워진 빈 행(0-based 100)
    expect(sheet.getCell('B101').value).toBeFalsy();
    expect(sheet.getCell('A502').value).toBe('mid500'); // 0-based 500 → 501
    expect(sheet.getCell('A1001').value).toBe('row999'); // 0-based 999 → 1000
    // 핵심 (1): 옛 코드가 **덮어써서 날려버린** 행.
    expect(sheet.getCell('A1002').value).toBe('row1000'); // 0-based 1000 → 1001
    // 핵심 (2): 옛 코드가 **아예 밀지 않아** 위쪽과 한 칸 어긋나 있던 행.
    expect(sheet.getCell('A1502').value).toBe('row1500'); // 0-based 1500 → 1501
    // 복제도 유실도 없다 — 각 표식은 파일 전체에서 정확히 한 번.
    for (const marker of ['top', 'mid500', 'row999', 'row1000', 'row1500']) {
      expect(countInColumn(sheet, 1, marker)).toBe(1);
    }
    // 옛 자리는 비었다(두 번 밀리거나 옛 값이 남지 않았는지).
    expect(sheet.getCell('A1001').value).not.toBe('row1000');
    expect(sheet.getCell('A1501').value).not.toBe('row1500');
    // 시트 높이도 정확히 한 행만 늘었다.
    const lazyGrid = (await import('../src/collab/room.js')).getRoomLazyGridForTest(docId)!;
    expect(lazyGrid.sheetMeta[0].rowCount).toBe(1601);
  }, 120_000);

  test('deleting a row on a 1600-row sheet pulls up every row below it, including unloaded rows 1000+', async () => {
    const { docId, socket } = await openRoom(await makeTallXlsx(), [0]);

    sendStructureOp(socket, { sheet: 0, kind: 'deleteRows', at: 100, count: 1 });
    await settleStructureOp(docId);

    const sheet = (await downloadWorkbook(docId)).getWorksheet('Sheet1')!;
    expect(sheet.getCell('A1').value).toBe('top');
    expect(sheet.getCell('A500').value).toBe('mid500'); // 0-based 500 → 499
    expect(sheet.getCell('A999').value).toBe('row999'); // 0-based 999 → 998
    expect(sheet.getCell('A1000').value).toBe('row1000'); // 0-based 1000 → 999
    expect(sheet.getCell('A1500').value).toBe('row1500'); // 0-based 1500 → 1499
    for (const marker of ['top', 'mid500', 'row999', 'row1000', 'row1500']) {
      expect(countInColumn(sheet, 1, marker)).toBe(1);
    }
    const lazyGrid = (await import('../src/collab/room.js')).getRoomLazyGridForTest(docId)!;
    expect(lazyGrid.sheetMeta[0].rowCount).toBe(1599);
    // 꼬리 행이 실제로 사라졌는지 — 예전엔 rowCount만 줄고 파일 꼬리가 그대로 남았다.
    expect(sheet.getCell(1600, 2).value).toBeFalsy();
  }, 120_000);

  test('deleting a row that itself holds data past row 1000 removes exactly that row', async () => {
    const { docId, socket } = await openRoom(await makeTallXlsx(), [0]);

    // 0-based 1000('row1000')을 직접 지운다 — 구독되지 않은 페이지 2의 행이다.
    sendStructureOp(socket, { sheet: 0, kind: 'deleteRows', at: 1000, count: 1 });
    await settleStructureOp(docId);

    const sheet = (await downloadWorkbook(docId)).getWorksheet('Sheet1')!;
    expect(countInColumn(sheet, 1, 'row1000')).toBe(0); // 지워진 행
    expect(sheet.getCell('A1000').value).toBe('row999'); // 그 위는 그대로
    expect(sheet.getCell('A1500').value).toBe('row1500'); // 아래는 한 칸 당겨졌다
    expect(countInColumn(sheet, 1, 'row1500')).toBe(1);
  }, 120_000);

  // ------------------------------------------------------------------ 열 방향
  /**
   * 열 방향은 **1000열 시나리오가 이제 도달 불가능**하다: 직전 픽스 라운드에서
   * MAX_CANONICAL_COLS를 파서의 실제 상한(MAX_COLS=120)으로 낮췄기 때문이다. 그래서 같은 실패
   * 모드(밀기 루프가 시트 실제 폭이 아닌 하드코딩 상한을 쓰고, 창 밖 데이터는 옮기지 못함)를
   * **도달 가능한 규모**에서 검증한다 — 열 데이터는 어차피 모든 페이지에 걸쳐 있으므로
   * "구독되지 않은 페이지의 열 데이터가 밀리는가"가 여기서도 핵심이다.
   */
  test('inserting a column shifts cells in rows that are not loaded into Yjs', async () => {
    const { docId, socket } = await openRoom(await makeTallXlsx(), [0]);

    sendStructureOp(socket, { sheet: 0, kind: 'insertCols', at: 0, count: 1 });
    await settleStructureOp(docId);

    const sheet = (await downloadWorkbook(docId)).getWorksheet('Sheet1')!;
    expect(sheet.getCell('A1').value).toBeFalsy(); // 새로 끼워진 빈 열
    expect(sheet.getCell('B1').value).toBe('top'); // A → B
    expect(sheet.getCell('C1').value).toBe('b'); // B → C
    // 구독되지 않은 페이지의 행들도 똑같이 밀렸다(열 삽입은 **행 번호를 바꾸지 않는다**).
    expect(sheet.getCell('B1000').value).toBe('row999'); // 0-based (999, 0) → (999, 1)
    expect(sheet.getCell('B1001').value).toBe('row1000'); // 0-based (1000, 0) → (1000, 1)
    expect(sheet.getCell('B1002').value).toBeFalsy(); // 원본 A1002는 빈 칸이었다(복제 금지)
    expect(sheet.getCell('B1501').value).toBe('row1500');
    expect(sheet.getCell('C1501').value).toBe('b');
    expect(countInColumn(sheet, 1, 'row1500')).toBe(0); // 옛 A열에는 아무것도 남지 않았다
    expect(countInColumn(sheet, 2, 'row1500')).toBe(1);
    const lazyGrid = (await import('../src/collab/room.js')).getRoomLazyGridForTest(docId)!;
    expect(lazyGrid.sheetMeta[0].colCount).toBe(3);
  }, 120_000);

  test('deleting a column removes it in every row, loaded or not', async () => {
    const { docId, socket } = await openRoom(await makeTallXlsx(), [0]);

    sendStructureOp(socket, { sheet: 0, kind: 'deleteCols', at: 0, count: 1 });
    await settleStructureOp(docId);

    const sheet = (await downloadWorkbook(docId)).getWorksheet('Sheet1')!;
    expect(sheet.getCell('A1').value).toBe('b'); // B열이 A로 당겨졌다
    expect(countInColumn(sheet, 1, 'top')).toBe(0); // 지워진 A열
    expect(countInColumn(sheet, 1, 'row1500')).toBe(0);
    expect(sheet.getCell('A1501').value).toBe('b'); // 구독 밖의 행도 같이 당겨졌다
    const lazyGrid = (await import('../src/collab/room.js')).getRoomLazyGridForTest(docId)!;
    expect(lazyGrid.sheetMeta[0].colCount).toBe(1);
  }, 120_000);

  // ------------------------------------------------------------------ 살아 있는 협업 화면(Yjs)
  test("a subscribed client's live Yjs state reflects the shift (not just canonical/export)", async () => {
    const { GRID_CELLS_MAP, GRID_STRUCTURE_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
    // 페이지 0·1을 구독한다 — 삽입 지점(행 10)과 그 아래가 전부 창 안에 있다.
    const { docId, socket } = await openRoom(await makeTallXlsx(), [0, 1]);
    const mirrorCells = socket.mirror.getMap<string>(GRID_CELLS_MAP);
    expect(mirrorCells.get(gridCellKey(0, 0, 0))).toBe('top');
    expect(mirrorCells.get(gridCellKey(0, 500, 0))).toBe('mid500');

    sendStructureOp(socket, { sheet: 0, kind: 'insertRows', at: 10, count: 1 });
    await settleStructureOp(docId);

    // 클라이언트의 Yjs 미러가 밀린 상태를 그대로 반영한다.
    expect(mirrorCells.get(gridCellKey(0, 0, 0))).toBe('top');
    expect(mirrorCells.get(gridCellKey(0, 10, 0))).toBeUndefined(); // 새로 끼워진 빈 행
    expect(mirrorCells.get(gridCellKey(0, 501, 0))).toBe('mid500');
    expect(mirrorCells.get(gridCellKey(0, 500, 0))).toBeUndefined(); // 옛 자리는 비었다
    // 구조 메타데이터(가상 스크롤 높이의 근거)도 함께 도착한다.
    const structure = socket.mirror.getMap<{ rowCount: number }>(GRID_STRUCTURE_MAP);
    expect(structure.get('0')!.rowCount).toBe(1601);
  }, 120_000);

  test('a row insert moves cell styles with their cells in the live Yjs styles map', async () => {
    const workbookIn = new ExcelJS.Workbook();
    const sheetIn = workbookIn.addWorksheet('Sheet1');
    sheetIn.getCell('A1').value = 'plain';
    sheetIn.getCell('A2').value = 'bold-me';
    sheetIn.getCell('A2').font = { bold: true };
    sheetIn.getCell('A3').value = 'after';
    const original = Buffer.from((await workbookIn.xlsx.writeBuffer()) as ArrayBuffer);

    const { GRID_STYLES_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
    const { docId, socket } = await openRoom(original, [0]);
    const mirrorStyles = socket.mirror.getMap<{ b?: 1 }>(GRID_STYLES_MAP);
    expect(mirrorStyles.get(gridCellKey(0, 1, 0))?.b).toBe(1);

    sendStructureOp(socket, { sheet: 0, kind: 'insertRows', at: 0, count: 1 });
    await settleStructureOp(docId);

    // 서식은 셀과 함께 한 칸 내려간다 — 옛 자리에 남아 "굵게가 엉뚱한 행에 걸리는" 일이 없어야 한다.
    expect(mirrorStyles.get(gridCellKey(0, 2, 0))?.b).toBe(1);
    expect(mirrorStyles.get(gridCellKey(0, 1, 0))?.b).toBeUndefined();
    // 저장된 파일에서도 같아야 한다(canonical의 GridCell.s가 셀과 함께 움직였는지).
    const sheet = (await downloadWorkbook(docId)).getWorksheet('Sheet1')!;
    expect(sheet.getCell('A3').value).toBe('bold-me');
    expect(sheet.getCell('A3').font?.bold).toBe(true);
    expect(sheet.getCell('A2').font?.bold).toBeFalsy();
  }, 120_000);

  // ------------------------------------------------------------------ 수식
  test('formula cells referencing shifted rows show updated computed values after the shift', async () => {
    const workbookIn = new ExcelJS.Workbook();
    const sheetIn = workbookIn.addWorksheet('Sheet1');
    sheetIn.getCell('A1').value = 10;
    sheetIn.getCell('A2').value = 20;
    sheetIn.getCell('C1').value = { formula: 'SUM(A1:A5)' };
    const original = Buffer.from((await workbookIn.xlsx.writeBuffer()) as ArrayBuffer);

    const { GRID_COMPUTED_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
    const { docId, socket } = await openRoom(original, [0]);
    const computed = socket.mirror.getMap<string>(GRID_COMPUTED_MAP);
    expect(computed.get(gridCellKey(0, 0, 2))).toBe('30');

    // A1 위에 행을 하나 끼운다 — 수식은 C2로 내려가고, 참조 범위 A1:A5 안의 값 배치가 바뀐다.
    sendStructureOp(socket, { sheet: 0, kind: 'insertRows', at: 0, count: 1 });
    await settleStructureOp(docId);

    // 수식 자체(원본 텍스트)도 한 칸 내려간다.
    const { GRID_CELLS_MAP } = await import('../src/collab/gridShared.js');
    expect(socket.mirror.getMap<string>(GRID_CELLS_MAP).get(gridCellKey(0, 1, 2))).toBe('=SUM(A1:A5)');
    // 새 자리에서 다시 계산된 값이 computed 맵으로 도착한다(A2=10, A3=20 → 여전히 30).
    expect(computed.get(gridCellKey(0, 1, 2))).toBe('30');
    expect(computed.get(gridCellKey(0, 0, 2))).toBeUndefined(); // 옛 자리의 계산값은 남지 않는다

    // 이제 참조 범위 밖으로 값을 밀어낸다 — 행 3개를 더 끼우면 20이 A6으로 나가 합계가 10이 된다.
    sendStructureOp(socket, { sheet: 0, kind: 'insertRows', at: 2, count: 3 });
    await settleStructureOp(docId);
    expect(computed.get(gridCellKey(0, 1, 2))).toBe('10');
  }, 120_000);

  // ------------------------------------------------------------------ 병합
  test('merge ranges are shifted with the data (canonical and the live merges map)', async () => {
    const workbookIn = new ExcelJS.Workbook();
    const sheetIn = workbookIn.addWorksheet('Sheet1');
    sheetIn.getCell('A1').value = 'title';
    sheetIn.mergeCells('A1:B1'); // 1-based r=1,c=1,rs=1,cs=2
    sheetIn.getCell('A3').value = 'body';
    const original = Buffer.from((await workbookIn.xlsx.writeBuffer()) as ArrayBuffer);

    const { GRID_MERGES_MAP } = await import('../src/collab/gridShared.js');
    const { docId, socket } = await openRoom(original, [0]);
    expect(socket.mirror.getMap<readonly { r: number }[]>(GRID_MERGES_MAP).get('0')).toEqual([
      { r: 1, c: 1, rs: 1, cs: 2 },
    ]);

    sendStructureOp(socket, { sheet: 0, kind: 'insertRows', at: 0, count: 1 });
    await settleStructureOp(docId);

    const shifted = [{ r: 2, c: 1, rs: 1, cs: 2 }];
    const lazyGrid = (await import('../src/collab/room.js')).getRoomLazyGridForTest(docId)!;
    expect(lazyGrid.sheetMeta[0].merges).toEqual(shifted);
    expect(socket.mirror.getMap<readonly { r: number }[]>(GRID_MERGES_MAP).get('0')).toEqual(shifted);
  }, 120_000);

  // ------------------------------------------------------------------ 잘못된/조작된 입력
  test('malformed or out-of-bounds structure ops are rejected without crashing or corrupting state', async () => {
    const { getRoomLazyGridForTest } = await import('../src/collab/room.js');
    const { docId, socket } = await openRoom(await makeTallXlsx(600), [0]);
    const before = getRoomLazyGridForTest(docId)!;
    const rowCountBefore = before.sheetMeta[0].rowCount;
    const colCountBefore = before.sheetMeta[0].colCount;

    const bad: unknown[] = [
      null,
      'not-an-object',
      42,
      {},
      { sheet: 0, kind: 'insertRows' }, // at/count 없음
      { sheet: 0, kind: 'nope', at: 0, count: 1 }, // 없는 kind
      { sheet: 0, kind: 'insertRows', at: -1, count: 1 },
      { sheet: 0, kind: 'insertRows', at: 0.5, count: 1 },
      { sheet: 0, kind: 'insertRows', at: 0, count: 0 },
      { sheet: 0, kind: 'insertRows', at: 0, count: -5 },
      { sheet: 0, kind: 'insertRows', at: 0, count: 1.5 },
      { sheet: 0, kind: 'insertRows', at: 0, count: 1e9 }, // 상한 초과(메모리 폭탄 차단)
      { sheet: 0, kind: 'insertRows', at: 0, count: Number.NaN },
      { sheet: 0, kind: 'insertRows', at: Number.POSITIVE_INFINITY, count: 1 },
      { sheet: -1, kind: 'insertRows', at: 0, count: 1 },
      { sheet: 9999, kind: 'insertRows', at: 0, count: 1 }, // 시트 상한 초과
      { sheet: 1, kind: 'insertRows', at: 0, count: 1 }, // 존재하지 않는 시트
      { sheet: 0, kind: 'insertRows', at: 10_000, count: 1 }, // 시트 끝보다 뒤
      { sheet: 0, kind: 'deleteRows', at: 10_000, count: 1 },
      // 시트 끝을 넘는 삭제 범위. makeTallXlsx(600)이라도 A1000/A1001/A1501에 값을 직접 써서 실제
      // 행 수는 1501까지 간다(뒤쪽 여유를 크게 둬 그 실제 크기에 안전하게 못 미친다).
      { sheet: 0, kind: 'deleteRows', at: 1499, count: 100 },
      { sheet: 0, kind: 'deleteCols', at: 0, count: 2 }, // 마지막 남은 열까지 지우려는 시도
    ];
    for (const payload of bad) {
      sendStructureOp(socket, payload);
    }
    // 형식이 깨진 JSON 자체도 던지지 않아야 한다.
    const rawEncoder = encoding.createEncoder();
    encoding.writeVarUint(rawEncoder, MESSAGE_GRID_STRUCTURE_OP);
    encoding.writeVarString(rawEncoder, '{not json');
    socket.emit('message', encoding.toUint8Array(rawEncoder));
    await settleStructureOp(docId);

    const after = getRoomLazyGridForTest(docId)!;
    expect(after.sheetMeta[0].rowCount).toBe(rowCountBefore);
    expect(after.sheetMeta[0].colCount).toBe(colCountBefore);
    expect(after.sheetMeta.length).toBe(1); // 없는 시트 번호가 자리를 만들어내지 않았다
    // 룸이 여전히 정상 동작한다 — 거부 뒤에도 올바른 연산은 받아들여야 한다.
    sendStructureOp(socket, { sheet: 0, kind: 'insertRows', at: 0, count: 1 });
    await settleStructureOp(docId);
    expect(getRoomLazyGridForTest(docId)!.sheetMeta[0].rowCount).toBe(rowCountBefore + 1);
  }, 120_000);

  test('a read-only connection cannot insert or delete rows', async () => {
    const { getRoomLazyGridForTest } = await import('../src/collab/room.js');
    const { docId, socket } = await openRoom(await makeTallXlsx(600), [0], true);
    const rowCountBefore = getRoomLazyGridForTest(docId)!.sheetMeta[0].rowCount;

    sendStructureOp(socket, { sheet: 0, kind: 'insertRows', at: 0, count: 1 });
    sendStructureOp(socket, { sheet: 0, kind: 'deleteRows', at: 0, count: 1 });
    await settleStructureOp(docId);

    expect(getRoomLazyGridForTest(docId)!.sheetMeta[0].rowCount).toBe(rowCountBefore);
  }, 120_000);

  /**
   * 행/열 삽입이 **유휴 자동저장으로 디스크까지** 도달하는지 — 위 테스트들은 전부
   * getCurrentBytes(즉시 다운로드) 경로를 보므로 이건 별개의 확인이다.
   *
   * 여기가 어긋날 수 있는 지점: 이 연산이 Yjs에 쓰는 트랜잭션에는 GRID_STRUCTURE_SHIFT_ORIGIN이
   * 붙는데, 창 시딩(GRID_WINDOW_SEED_ORIGIN)이나 계산값 쓰기처럼 **자동저장 면제**로 분류해 버리면
   * 편집이 화면·다운로드에는 보이지만 자동저장으로는 절대 파일에 반영되지 않는다(이 프로젝트가
   * 이미 여러 번 겪은 조용한 손실 클래스). 셀 재시딩은 면제 origin을 쓰므로, 면제되지 않는
   * 구조 메타데이터 쓰기가 맨 마지막에 오는 순서에 실제로 의존한다.
   */
  test('a row insert reaches the disk through idle autosave (not just the immediate-download path)', async () => {
    const { store } = await import('../src/storage.js');
    const { docId, socket } = await openRoom(await makeTallXlsx(600), [0]);
    const saveSpy = vi.spyOn(store, 'saveOriginal');

    sendStructureOp(socket, { sheet: 0, kind: 'insertRows', at: 0, count: 1 });
    await settleStructureOp(docId);

    await vi.advanceTimersByTimeAsync(8_100); // AUTOSAVE_IDLE_MS(8초) + 여유
    await flushEventLoop();
    expect(saveSpy).toHaveBeenCalled();

    // 디스크에 실제로 밀린 상태가 저장됐는지 — 저장된 바이트를 다시 열어 확인한다.
    const savedBytes = await store.readOriginal(docId, 'xlsx');
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(savedBytes as unknown as ArrayBuffer);
    const sheet = reopened.getWorksheet('Sheet1')!;
    expect(sheet.getCell('A1').value).toBeFalsy(); // 새로 삽입된 빈 행
    expect(sheet.getCell('A2').value).toBe('top'); // 0-based 0 → 1
    expect(sheet.getCell('A502').value).toBe('mid500'); // 0-based 500 → 501
    saveSpy.mockRestore();
  }, 120_000);

  test('two structure ops in a row compose correctly (the chain serializes them)', async () => {
    const { docId, socket } = await openRoom(await makeTallXlsx(), [0]);

    // 연달아 보낸다 — 체인이 순서대로 처리해야 두 삽입이 정확히 두 칸으로 누적된다.
    sendStructureOp(socket, { sheet: 0, kind: 'insertRows', at: 10, count: 1 });
    sendStructureOp(socket, { sheet: 0, kind: 'insertCols', at: 0, count: 1 });
    await settleStructureOp(docId);

    const sheet = (await downloadWorkbook(docId)).getWorksheet('Sheet1')!;
    expect(sheet.getCell('B1').value).toBe('top'); // 열 하나 밀림
    expect(sheet.getCell('B1002').value).toBe('row1000'); // 행·열 각각 하나씩 밀림
    expect(sheet.getCell('B1502').value).toBe('row1500');
    expect(countInColumn(sheet, 2, 'row1500')).toBe(1);
  }, 120_000);
});

// --------------------------------------------------------------------- 순수 연산 단위 테스트
describe('gridShift.ts: pure canonical shift primitives', () => {
  test('merge ranges shift, grow, shrink and vanish as the data moves', async () => {
    const { shiftMerges } = await import('../src/xlsx/gridShift.js');
    // 삽입 지점(0-based 4) 아래의 병합은 통째로 밀린다.
    expect(shiftMerges([{ r: 10, c: 1, rs: 2, cs: 2 }], 'row', 4, 1, true)).toEqual([
      { r: 11, c: 1, rs: 2, cs: 2 },
    ]);
    // 병합 **안쪽**에 끼워지면 병합이 넓어진다(엑셀과 같은 동작).
    expect(shiftMerges([{ r: 3, c: 1, rs: 4, cs: 1 }], 'row', 4, 1, true)).toEqual([
      { r: 3, c: 1, rs: 5, cs: 1 },
    ]);
    // 삽입 지점 위의 병합은 그대로.
    expect(shiftMerges([{ r: 1, c: 1, rs: 2, cs: 2 }], 'row', 4, 1, true)).toEqual([
      { r: 1, c: 1, rs: 2, cs: 2 },
    ]);
    // 삭제가 병합의 일부만 먹으면 그만큼 좁아진다.
    expect(shiftMerges([{ r: 3, c: 1, rs: 4, cs: 2 }], 'row', 3, 2, false)).toEqual([
      { r: 3, c: 1, rs: 2, cs: 2 },
    ]);
    // 병합보다 앞에서 지워지면 시작점이 당겨진다.
    expect(shiftMerges([{ r: 10, c: 1, rs: 2, cs: 2 }], 'row', 0, 3, false)).toEqual([
      { r: 7, c: 1, rs: 2, cs: 2 },
    ]);
    // 병합 전체가 지워지면 목록에서 사라진다.
    expect(shiftMerges([{ r: 3, c: 1, rs: 2, cs: 2 }], 'row', 2, 4, false)).toEqual([]);
    // 1×1로 쪼그라든 것은 더 이상 병합이 아니다.
    expect(shiftMerges([{ r: 3, c: 1, rs: 2, cs: 1 }], 'row', 3, 1, false)).toEqual([]);
    // 열 축도 같은 규칙.
    expect(shiftMerges([{ r: 1, c: 5, rs: 1, cs: 3 }], 'col', 2, 2, true)).toEqual([
      { r: 1, c: 7, rs: 1, cs: 3 },
    ]);
    expect(shiftMerges([{ r: 1, c: 2, rs: 1, cs: 4 }], 'col', 2, 1, false)).toEqual([
      { r: 1, c: 2, rs: 1, cs: 3 },
    ]);
  });

});
