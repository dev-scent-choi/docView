import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeSyncStep1, writeUpdate } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

/**
 * 이 파일이 붙잡는 회귀(픽스 웨이브 재리뷰의 New Important #2):
 *
 * canonical 그리드는 **신뢰할 수 없는 클라이언트 입력**을 따라 스스로 넓어진다 — 범위 밖 셀 키
 * ("0:299999:1023") 하나, 또는 GRID_STRUCTURE_MAP에 쓴 거대한 rowCount/colCount 하나면 된다.
 * 예전(canonical이 시트 전체를 실제 GridCell 배열로 담던 시절)에는 그 성장 루프에 양보 지점이
 * 하나도 없었고 상한(MAX_CANONICAL_COLS=1024)도 이 코드베이스가 실제로 파싱하는 열 수
 * (MAX_COLS=120)보다 8.5배 넉넉해서, 메시지 한 통으로 약 2.4GB를 할당하며 이벤트 루프를 1초 넘게
 * 통째로 멈출 수 있었다.
 *
 * **LazyGridStore 재설계(2026-09-09) 이후**: `growSheet`는 더 이상 실제 `GridCell` 행을 채우지
 * 않는다 — `logicalToOriginalRow/Col`(정수 하나씩만 담는 매핑 배열)에 -1을 채울 뿐이다. 상한까지
 * 늘려도 정수 300,000개 push는 그 자체로 이미 충분히 싸서, 예전처럼 청크+양보를 할 이유가
 * 사라졌다(청크/양보 제거 자체가 이 재설계의 목적 중 하나였다) — 그래서 아래 첫 테스트는 더 이상
 * "양보하는지"가 아니라 "커진 만큼 매핑 배열도 정확히 자랐는지"만 검증한다. 상한 클램프(열은
 * MAX_COLS, 행은 MAX_ROWS) 자체는 여전히 유효한 회귀이므로 나머지 테스트는 그대로 유지한다.
 */

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
    if (decoding.readVarUint(decoder) === MESSAGE_SYNC) {
      const syncType = decoding.readVarUint(decoder);
      if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) readUpdate(decoder, this.mirror, this);
    }
    cb?.();
  }
  close(): void {}
  ping(): void {}
}

async function flushEventLoop(turns = 60): Promise<void> {
  for (let i = 0; i < turns; i += 1) await new Promise((resolve) => setImmediate(resolve));
}

function sendSubscribe(socket: FakeSocket, payload: unknown): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_PAGE_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify(payload));
  socket.emit('message', encoding.toUint8Array(encoder));
}

function syncMirrorWithRoom(socket: FakeSocket): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeSyncStep1(encoder, socket.mirror);
  socket.emit('message', encoding.toUint8Array(encoder));
}

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
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

describe('room.ts: canonical grid growth is bounded and yields to the event loop', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-growth-'));
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

  async function openRoom(): Promise<{ docId: string; socket: FakeSocket }> {
    const { joinEditSession } = await import('../src/collab/room.js');
    const { store } = await import('../src/storage.js');
    const docId = randomUUID();
    await store.saveOriginal(docId, 'xlsx', await makeSampleXlsx());
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
    sendSubscribe(socket, { sheet: 0, pages: [0] });
    await flushEventLoop();
    syncMirrorWithRoom(socket);
    return { docId, socket };
  }

  test(
    'a huge rowCount growth request completes immediately (growSheet only pads an integer mapping array)',
    async () => {
      const { getRoomLazyGridForTest, awaitCanonicalGrowthForTest } = await import('../src/collab/room.js');
      const { GRID_STRUCTURE_MAP } = await import('../src/collab/gridShared.js');
      const { YIELD_EVERY_N_ROWS } = await import('../src/spreadsheet.js');
      const { docId, socket } = await openRoom();

      const GROWN_ROWS = YIELD_EVERY_N_ROWS * 8; // 예전 청크 경계를 여러 번 넘기던 규모 그대로 유지
      applyOnMirror(socket, (doc) => {
        const structure = doc.getMap<Record<string, unknown>>(GRID_STRUCTURE_MAP);
        const current = structure.get('0')!;
        structure.set('0', { ...current, rowCount: GROWN_ROWS });
      });
      await awaitCanonicalGrowthForTest(docId);

      const lazyGrid = getRoomLazyGridForTest(docId)!;
      expect(lazyGrid.sheetMeta[0].rowCount).toBe(GROWN_ROWS);
      // 매핑 배열도 정확히 그만큼 자랐는지(패딩이 짧게 끝나 뒤쪽 인덱스가 undefined로 남지 않는지).
      expect(lazyGrid.logicalToOriginalRow[0].length).toBe(GROWN_ROWS);
    },
    120_000,
  );

  test('a colCount beyond the real xlsx column ceiling is clamped, not allocated', async () => {
    const { getRoomLazyGridForTest, awaitCanonicalGrowthForTest } = await import('../src/collab/room.js');
    const { GRID_STRUCTURE_MAP } = await import('../src/collab/gridShared.js');
    const { MAX_COLS } = await import('../src/spreadsheet.js');
    const { docId, socket } = await openRoom();

    // 예전 상한(1024)에서는 그대로 통과하던 값 — 새 상한에서는 반드시 잘려야 한다.
    applyOnMirror(socket, (doc) => {
      const structure = doc.getMap<Record<string, unknown>>(GRID_STRUCTURE_MAP);
      const current = structure.get('0')!;
      structure.set('0', { ...current, colCount: 1024 });
    });
    await awaitCanonicalGrowthForTest(docId);

    const lazyGrid = getRoomLazyGridForTest(docId)!;
    expect(lazyGrid.sheetMeta[0].colCount).toBe(MAX_COLS);
    expect(lazyGrid.logicalToOriginalCol[0].length).toBeLessThanOrEqual(MAX_COLS);
  });

  test('a cell key past the real xlsx column ceiling is rejected outright', async () => {
    const { getRoomLazyGridForTest, awaitCanonicalGrowthForTest } = await import('../src/collab/room.js');
    const { GRID_CELLS_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
    const { MAX_COLS } = await import('../src/spreadsheet.js');
    const { getCell } = await import('../src/xlsx/lazyGridStore.js');
    const { docId, socket } = await openRoom();

    applyOnMirror(socket, (doc) => {
      // 열 500: 예전 상한(1024) 아래라 통과하던 좌표. 이제는 버려져야 한다.
      doc.getMap<string>(GRID_CELLS_MAP).set(gridCellKey(0, 0, 500), 'evil');
      // 바로 옆 유효 좌표는 그대로 반영되어야 한다(상한이 멀쩡한 편집까지 막지는 않는다).
      doc.getMap<string>(GRID_CELLS_MAP).set(gridCellKey(0, 0, 2), 'ok');
    });
    await awaitCanonicalGrowthForTest(docId);
    await flushEventLoop();

    const lazyGrid = getRoomLazyGridForTest(docId)!;
    expect(lazyGrid.sheetMeta[0].colCount).toBeLessThanOrEqual(MAX_COLS);
    // 상한 밖으로 거부된 좌표는 애초에 dirtyOverlay에 쓰인 적이 없다 — getCell은 "값 없음"을 0으로
    // 나타낸다(원본에도 있을 수 없는 칸이므로 로직상 undefined가 아니라 0).
    expect(await getCell(lazyGrid, 0, 0, 500)).toBe(0);
    expect(await getCell(lazyGrid, 0, 0, 2)).toEqual({ v: 'ok' });
  });

  test('a rowCount far beyond the parser ceiling is clamped to MAX_ROWS', async () => {
    const { getRoomLazyGridForTest, awaitCanonicalGrowthForTest } = await import('../src/collab/room.js');
    const { GRID_STRUCTURE_MAP } = await import('../src/collab/gridShared.js');
    const { MAX_ROWS } = await import('../src/spreadsheet.js');
    const { docId, socket } = await openRoom();

    applyOnMirror(socket, (doc) => {
      const structure = doc.getMap<Record<string, unknown>>(GRID_STRUCTURE_MAP);
      const current = structure.get('0')!;
      structure.set('0', { ...current, rowCount: 50_000_000 });
    });
    await awaitCanonicalGrowthForTest(docId);

    expect(getRoomLazyGridForTest(docId)!.sheetMeta[0].rowCount).toBe(MAX_ROWS);
  }, 120_000);
});
