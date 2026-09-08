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

/**
 * 페이지 채우기/비우기는 청크마다 `yieldToEventLoop()`(= 실제 setImmediate)로 양보하므로,
 * 구독 메시지를 보낸 뒤 그 결과를 관찰하려면 실제 이벤트 루프 턴을 여러 번 흘려보내야 한다.
 * 가짜 타이머를 쓰는 테스트에서도 setImmediate는 일부러 가짜로 만들지 않으므로(그렇게 하면
 * ExcelJS의 zip 스트림 파싱이 영원히 끝나지 않는다 — roomGridIndex.test.ts 주석 참고) 이
 * 헬퍼는 두 경우 모두에서 동작한다.
 */
async function flushEventLoop(turns = 50): Promise<void> {
  for (let i = 0; i < turns; i += 1) await new Promise((resolve) => setImmediate(resolve));
}

async function makeBigXlsx(rows: number): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  for (let r = 1; r <= rows; r += 1) sheet.getCell(r, 1).value = `row-${r}`;
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

/**
 * 1000행(= 페이지 0과 1) 시트 — A1:A3은 숫자, B601(0-based로 행 600·열 1, 즉 **페이지 1**)에
 * `SUM(A1:A3)` 수식이 있다. 수식은 페이지 1에 있고 참조 대상은 페이지 0에 있으므로, "창 밖의
 * 셀을 참조하는 수식도 canonical 기준으로 정확히 계산된다"를 검증할 수 있다.
 */
async function makeCrossPageFormulaXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  for (let r = 1; r <= 1000; r += 1) sheet.getCell(r, 1).value = r <= 3 ? r : `row-${r}`;
  // 계산 결과(result)를 일부러 넣지 않는다 — 넣으면 parseXlsxToGrid가 그 캐시된 값을 셀 값으로
  // 삼아 "수식 셀"이 아니게 되고, 재계산 경로 자체를 검증할 수 없다.
  sheet.getCell(601, 2).value = { formula: 'SUM(A1:A3)' } as ExcelJS.CellFormulaValue;
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

  test(
    'opening a large sheet does NOT seed all rows into Yjs — only the subscribed page',
    async () => {
      const { joinEditSession } = await import('../src/collab/room.js');
      const { GRID_CELLS_MAP, GRID_STRUCTURE_MAP, gridCellKey } = await import('../src/collab/gridShared.js');

      const docId = randomUUID();
      await seedDoc(docId, await makeBigXlsx(3000)); // 6개 페이지(0~5, 각 500행)

      const socket = new FakeSocket();
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);

      // 아직 아무 구독도 없는 시점 — 셀은 하나도 없어야 하지만 구조 메타데이터는 이미 전량
      // 있어야 한다(GRID_STRUCTURE_MAP/GRID_MERGES_MAP은 절대 윈도잉하지 않는다).
      socket.requestSync();
      await flushEventLoop();
      expect(socket.mirror.getMap<string>(GRID_CELLS_MAP).size).toBe(0);
      expect(socket.mirror.getMap<{ rowCount: number }>(GRID_STRUCTURE_MAP).get('0')?.rowCount).toBe(3000);

      sendSubscribe(socket, 0, [0]); // 첫 페이지(0~499행)만 구독
      await flushEventLoop();
      socket.requestSync();

      const cells = socket.mirror.getMap<string>(GRID_CELLS_MAP);
      expect(cells.get(gridCellKey(0, 0, 0))).toBe('row-1'); // 구독한 페이지 안 — 있어야 함
      expect(cells.get(gridCellKey(0, 499, 0))).toBe('row-500'); // 페이지 경계 마지막 행
      expect(cells.get(gridCellKey(0, 500, 0))).toBeUndefined(); // 바로 다음 페이지(1번) — 없어야 함
      expect(cells.get(gridCellKey(0, 2500, 0))).toBeUndefined(); // 구독 안 한 페이지(5번) — 없어야 함
      expect(cells.size).toBe(500); // 3000행 전체가 아니라 딱 한 페이지 분량
    },
    30_000,
  );

  test(
    'subscribing to a new page fills it in; the union still includes previously subscribed pages',
    async () => {
      const { joinEditSession } = await import('../src/collab/room.js');
      const { GRID_CELLS_MAP, gridCellKey } = await import('../src/collab/gridShared.js');

      const docId = randomUUID();
      await seedDoc(docId, await makeBigXlsx(3000));

      const socket = new FakeSocket();
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
      sendSubscribe(socket, 0, [0, 3]); // 0페이지(0~499행)와 3페이지(1500~1999행) 둘 다
      await flushEventLoop();
      socket.requestSync();

      const cells = socket.mirror.getMap<string>(GRID_CELLS_MAP);
      expect(cells.get(gridCellKey(0, 0, 0))).toBe('row-1');
      expect(cells.get(gridCellKey(0, 1500, 0))).toBe('row-1501');
      expect(cells.get(gridCellKey(0, 2500, 0))).toBeUndefined();
      expect(cells.size).toBe(1000); // 딱 두 페이지 분량
    },
    30_000,
  );

  test(
    'a page removed from the union stays in Yjs for the grace period, then is removed',
    async () => {
      vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

      const { joinEditSession, getSeededPagesForTest } = await import('../src/collab/room.js');
      const { GRID_CELLS_MAP, gridCellKey } = await import('../src/collab/gridShared.js');

      const docId = randomUUID();
      await seedDoc(docId, await makeBigXlsx(1000)); // 페이지 0과 1

      const socket = new FakeSocket();
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
      sendSubscribe(socket, 0, [0]);
      await flushEventLoop();
      socket.requestSync();

      expect(socket.mirror.getMap<string>(GRID_CELLS_MAP).get(gridCellKey(0, 0, 0))).toBe('row-1');
      expect(getSeededPagesForTest(docId, 0).has(0)).toBe(true);

      sendSubscribe(socket, 0, [1]); // 0페이지 구독 해제, 1페이지로 이동
      await flushEventLoop();

      // 유예 시간(10초) 이전 — 0페이지는 아직 Yjs에 남아 있어야 한다. canonical에는 항상 전체가
      // 있으므로 클라이언트 mirror만으로는 "지워짐"을 관찰하기 어렵다 — 서버 내부 상태를 보는
      // 테스트 전용 창구로 확인한다.
      expect(getSeededPagesForTest(docId, 0).has(0)).toBe(true);
      expect(getSeededPagesForTest(docId, 0).has(1)).toBe(true);

      await vi.advanceTimersByTimeAsync(10_100); // 유예 시간 경과
      await flushEventLoop();

      expect(getSeededPagesForTest(docId, 0).has(0)).toBe(false);
      expect(getSeededPagesForTest(docId, 0).has(1)).toBe(true);
    },
    30_000,
  );

  test(
    're-subscribing within the grace period cancels the pending removal',
    async () => {
      vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

      const { joinEditSession, getSeededPagesForTest } = await import('../src/collab/room.js');

      const docId = randomUUID();
      await seedDoc(docId, await makeBigXlsx(1000));

      const socket = new FakeSocket();
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
      sendSubscribe(socket, 0, [0]);
      await flushEventLoop();
      expect(getSeededPagesForTest(docId, 0).has(0)).toBe(true);

      sendSubscribe(socket, 0, [1]); // 0페이지 제거 예약
      await flushEventLoop();
      await vi.advanceTimersByTimeAsync(5_000); // 유예 시간의 절반만 흐름
      sendSubscribe(socket, 0, [0, 1]); // 스크롤이 되돌아와 0페이지 재구독 → 예약 취소
      await flushEventLoop();

      // 원래 예약 시점으로부터 유예 시간이 한참 지난 뒤에도 0페이지는 그대로 있어야 한다.
      await vi.advanceTimersByTimeAsync(20_000);
      await flushEventLoop();

      expect(getSeededPagesForTest(docId, 0).has(0)).toBe(true);
      expect(getSeededPagesForTest(docId, 0).has(1)).toBe(true);
    },
    30_000,
  );

  test(
    'unseeding a page does not wipe those rows from the canonical grid (no silent data loss)',
    async () => {
      vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

      const { joinEditSession, getRoomCanonicalGridForTest, getSeededPagesForTest } = await import(
        '../src/collab/room.js'
      );

      const docId = randomUUID();
      await seedDoc(docId, await makeBigXlsx(1000));

      const socket = new FakeSocket();
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
      sendSubscribe(socket, 0, [0]);
      await flushEventLoop();

      sendSubscribe(socket, 0, [1]); // 0페이지가 창에서 빠진다
      await flushEventLoop();
      await vi.advanceTimersByTimeAsync(10_100);
      await flushEventLoop();

      // Yjs에서는 사라졌지만…
      expect(getSeededPagesForTest(docId, 0).has(0)).toBe(false);
      // …canonical 그리드(저장·다운로드의 유일한 기준)에는 그대로 남아 있어야 한다. 여기서
      // 0으로 비워지면 "스크롤만 했는데 500행이 파일에서 사라지는" 조용한 데이터 손실이 된다
      // (unseed의 키 삭제를 cells.observe가 "사용자가 지웠다"로 오해하는 경로).
      const canonical = getRoomCanonicalGridForTest(docId);
      expect(canonical).not.toBeNull();
      expect(canonical!.sheets[0].rows[0][0]).toEqual({ v: 'row-1' });
      expect(canonical!.sheets[0].rows[499][0]).toEqual({ v: 'row-500' });
    },
    30_000,
  );

  test(
    'a disconnecting client releases the pages only it was viewing, after the grace period',
    async () => {
      vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

      const { joinEditSession, getSeededPagesForTest } = await import('../src/collab/room.js');

      const docId = randomUUID();
      await seedDoc(docId, await makeBigXlsx(1000)); // 페이지 0과 1

      const socketA = new FakeSocket();
      const socketB = new FakeSocket();
      await joinEditSession(socketA as never, {} as never, docId, 'xlsx', [], false);
      await joinEditSession(socketB as never, {} as never, docId, 'xlsx', [], false);
      sendSubscribe(socketA, 0, [0]);
      sendSubscribe(socketB, 0, [1]);
      await flushEventLoop();
      expect([...getSeededPagesForTest(docId, 0)].sort()).toEqual([0, 1]);

      socketA.emit('close');
      await flushEventLoop();
      await vi.advanceTimersByTimeAsync(10_100);
      await flushEventLoop();

      // A만 보던 0페이지는 비워지고, B가 여전히 보는 1페이지는 남아야 한다. closeConn이 구독을
      // 지우기 "전"에 이전 합집합을 확보하지 않으면 0페이지가 차이에 나타나지 않아 영원히 남는다.
      expect([...getSeededPagesForTest(docId, 0)]).toEqual([1]);
    },
    30_000,
  );

  test(
    'a disconnecting client does NOT release a page another client is still viewing',
    async () => {
      vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

      const { joinEditSession, getSeededPagesForTest } = await import('../src/collab/room.js');

      const docId = randomUUID();
      await seedDoc(docId, await makeBigXlsx(1000));

      const socketA = new FakeSocket();
      const socketB = new FakeSocket();
      await joinEditSession(socketA as never, {} as never, docId, 'xlsx', [], false);
      await joinEditSession(socketB as never, {} as never, docId, 'xlsx', [], false);
      sendSubscribe(socketA, 0, [0]);
      sendSubscribe(socketB, 0, [0]); // 둘이 같은 페이지를 본다
      await flushEventLoop();
      expect([...getSeededPagesForTest(docId, 0)]).toEqual([0]);

      socketA.emit('close');
      await flushEventLoop();
      await vi.advanceTimersByTimeAsync(10_100);
      await flushEventLoop();

      // 합집합이 A 퇴장 전후로 똑같이 {0}이므로 비울 페이지가 없다 — 여기서 비워지면 B의 화면이
      // 통째로 빈 칸이 된다.
      expect([...getSeededPagesForTest(docId, 0)]).toEqual([0]);
    },
    30_000,
  );

  test(
    'a newly subscribed page arrives with its computed formula values, even for formulas referencing rows outside the window',
    async () => {
      const { joinEditSession } = await import('../src/collab/room.js');
      const { GRID_CELLS_MAP, GRID_COMPUTED_MAP, gridCellKey } = await import('../src/collab/gridShared.js');

      const docId = randomUUID();
      await seedDoc(docId, await makeCrossPageFormulaXlsx());

      const socket = new FakeSocket();
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
      sendSubscribe(socket, 0, [1]); // 수식이 있는 페이지 1만 구독 — 참조 대상(A1:A3)은 페이지 0
      await flushEventLoop();
      socket.requestSync();

      const cells = socket.mirror.getMap<string>(GRID_CELLS_MAP);
      const computed = socket.mirror.getMap<string>(GRID_COMPUTED_MAP);
      // 수식 원문은 구독한 페이지 안이라 있어야 하고, 참조 대상 셀은 창 밖이라 없어야 한다.
      expect(cells.get(gridCellKey(0, 600, 1))).toBe('=SUM(A1:A3)');
      expect(cells.get(gridCellKey(0, 0, 0))).toBeUndefined();
      // 계산은 항상 canonical 전체를 보므로, 창 밖의 값을 참조해도 정확히 계산된다 — 그리고 그
      // 계산값은 페이지가 처음 구독되는 그 시점에 셀 데이터와 함께 도착해야 한다(시딩 시점에
      // 전 시트를 미리 계산해 두는 낭비 없이).
      expect(computed.get(gridCellKey(0, 600, 1))).toBe('6');
      // 구독하지 않은 페이지 0에는 계산값도 쌓이지 않아야 한다.
      expect(computed.get(gridCellKey(0, 0, 0))).toBeUndefined();
    },
    30_000,
  );
});
