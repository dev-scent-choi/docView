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
 * 그런데 그 성장 루프에는 양보 지점이 하나도 없었고 상한(MAX_CANONICAL_COLS=1024)도 이 코드베이스가
 * 실제로 파싱하는 열 수(MAX_COLS=120)보다 8.5배 넉넉해서, 메시지 한 통으로 약 2.4GB를 할당하며
 * 이벤트 루프를 1초 넘게 통째로 멈출 수 있었다 — 이 프로젝트의 윈도잉 작업이 없애려는 바로 그
 * 버그 클래스(6355a17, Task 4의 onRange 상한과 같은 계열)다.
 *
 * **측정 방법**: xlsxExportYielding.test.ts와 같은 방식으로, 절대 틱 수 같은 불안정한 임계값을 쓰지
 * 않는다. canonical 시트의 rows 배열을 Proxy로 감싸 `push` 한 번마다 "그 시점까지 이벤트 루프가 몇 번
 * 돌았는지"를 기록한다 — 양보가 전혀 없으면 모든 push가 **같은 틱 값**을 가진다(그 사이 루프가 한
 * 번도 돌지 못했다는 직접 증거).
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

function startTicker(): { count: () => number; stop: () => void } {
  let ticks = 0;
  let running = true;
  function schedule(): void {
    setImmediate(() => {
      if (!running) return;
      ticks += 1;
      schedule();
    });
  }
  schedule();
  return {
    count: () => ticks,
    stop: () => {
      running = false;
    },
  };
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
    'a huge rowCount growth request is chunked with real event-loop yields',
    async () => {
      const { getRoomCanonicalGridForTest, awaitCanonicalGrowthForTest } = await import('../src/collab/room.js');
      const { GRID_STRUCTURE_MAP } = await import('../src/collab/gridShared.js');
      const { YIELD_EVERY_N_ROWS } = await import('../src/spreadsheet.js');
      const { docId, socket } = await openRoom();

      // 성장 루프가 행을 하나 밀어 넣을 때마다 "지금까지 이벤트 루프가 몇 번 돌았는지"를 기록한다.
      const canonical = getRoomCanonicalGridForTest(docId)!;
      const sheet = canonical.sheets[0] as unknown as { rows: unknown[] };
      const realRows = sheet.rows;
      const ticker = startTicker();
      const ticksAtPush: number[] = [];
      sheet.rows = new Proxy(realRows, {
        get(target, prop, receiver) {
          if (prop === 'push') {
            return (...args: unknown[]): number => {
              ticksAtPush.push(ticker.count());
              return Array.prototype.push.apply(target, args);
            };
          }
          return Reflect.get(target, prop, receiver) as unknown;
        },
      });

      const GROWN_ROWS = YIELD_EVERY_N_ROWS * 8; // 청크 경계를 여러 번 넘긴다
      applyOnMirror(socket, (doc) => {
        const structure = doc.getMap<Record<string, unknown>>(GRID_STRUCTURE_MAP);
        const current = structure.get('0')!;
        structure.set('0', { ...current, rowCount: GROWN_ROWS });
      });
      await awaitCanonicalGrowthForTest(docId);
      ticker.stop();

      // 성장 자체는 실제로 끝났는가.
      expect(realRows.length).toBe(GROWN_ROWS);
      expect(ticksAtPush.length).toBeGreaterThanOrEqual(GROWN_ROWS - 1);

      // 양보가 없던 예전 코드에서는 이 값이 1이다 — 전체 할당이 한 호흡의 동기 블록이라
      // 그동안 이벤트 루프가 단 한 번도 돌지 못하기 때문이다.
      expect(new Set(ticksAtPush).size).toBeGreaterThanOrEqual(5);

      // "한 번도 안 쉬고 연달아 할당한 최대 행 수"도 직접 확인한다.
      let longestUnyieldedRun = 1;
      let currentRun = 1;
      for (let i = 1; i < ticksAtPush.length; i += 1) {
        currentRun = ticksAtPush[i] === ticksAtPush[i - 1] ? currentRun + 1 : 1;
        longestUnyieldedRun = Math.max(longestUnyieldedRun, currentRun);
      }
      expect(longestUnyieldedRun).toBeLessThanOrEqual(YIELD_EVERY_N_ROWS + 100);
    },
    120_000,
  );

  test('a colCount beyond the real xlsx column ceiling is clamped, not allocated', async () => {
    const { getRoomCanonicalGridForTest, awaitCanonicalGrowthForTest } = await import('../src/collab/room.js');
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

    const sheet = getRoomCanonicalGridForTest(docId)!.sheets[0];
    expect(sheet.colCount).toBe(MAX_COLS);
    for (const row of sheet.rows) expect(row.length).toBeLessThanOrEqual(MAX_COLS);
  });

  test('a cell key past the real xlsx column ceiling is rejected outright', async () => {
    const { getRoomCanonicalGridForTest, awaitCanonicalGrowthForTest } = await import('../src/collab/room.js');
    const { GRID_CELLS_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
    const { MAX_COLS } = await import('../src/spreadsheet.js');
    const { docId, socket } = await openRoom();

    applyOnMirror(socket, (doc) => {
      // 열 500: 예전 상한(1024) 아래라 통과하던 좌표. 이제는 버려져야 한다.
      doc.getMap<string>(GRID_CELLS_MAP).set(gridCellKey(0, 0, 500), 'evil');
      // 바로 옆 유효 좌표는 그대로 반영되어야 한다(상한이 멀쩡한 편집까지 막지는 않는다).
      doc.getMap<string>(GRID_CELLS_MAP).set(gridCellKey(0, 0, 2), 'ok');
    });
    await awaitCanonicalGrowthForTest(docId);
    await flushEventLoop();

    const sheet = getRoomCanonicalGridForTest(docId)!.sheets[0];
    expect(sheet.colCount).toBeLessThanOrEqual(MAX_COLS);
    expect(sheet.rows[0][500]).toBeUndefined();
    expect(sheet.rows[0][2]).toEqual({ v: 'ok' });
  });

  test('a rowCount far beyond the parser ceiling is clamped to MAX_ROWS', async () => {
    const { getRoomCanonicalGridForTest, awaitCanonicalGrowthForTest } = await import('../src/collab/room.js');
    const { GRID_STRUCTURE_MAP } = await import('../src/collab/gridShared.js');
    const { MAX_ROWS } = await import('../src/spreadsheet.js');
    const { docId, socket } = await openRoom();

    applyOnMirror(socket, (doc) => {
      const structure = doc.getMap<Record<string, unknown>>(GRID_STRUCTURE_MAP);
      const current = structure.get('0')!;
      structure.set('0', { ...current, rowCount: 50_000_000 });
    });
    await awaitCanonicalGrowthForTest(docId);

    expect(getRoomCanonicalGridForTest(docId)!.sheets[0].rowCount).toBe(MAX_ROWS);
  }, 120_000);
});
