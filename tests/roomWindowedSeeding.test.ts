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
  for (let i = 0; i < turns; i += 1) {
    await new Promise((resolve) => setImmediate(resolve));
    await drainPendingDiskIo();
    await drainRoomBackgroundWork();
  }
}

/**
 * 서버가 흘려보낸(fire-and-forget) 백그라운드 작업 — 페이지 창 반영, 수식 재계산, 셀/서식 편집의
 * canonical 되반영 — 이 실제로 끝날 때까지 기다린다. 턴 수 같은 **대리 지표가 아니라 그 작업의
 * 완료 자체**를 기다리는 것이 핵심이다(room.ts의 Room.backgroundWork 주석 참고).
 *
 * 룸 모듈을 아직(또는 전혀) 불러오지 않은 테스트에서도 안전하도록 조용히 넘어간다.
 */
async function drainRoomBackgroundWork(): Promise<void> {
  try {
    const { drainRoomBackgroundWorkForTest } = await import('../src/collab/room.js');
    await drainRoomBackgroundWorkForTest();
  } catch {
    /* 이 테스트는 룸을 쓰지 않는다 */
  }
}


/**
 * setImmediate 턴만 흘려보내는 것으로는 **진짜 디스크 I/O를 기다릴 수 없다** — 이 한 줄이 이
 * 파일(과 같은 헬퍼를 쓰는 다른 룸 테스트들)의 타이밍 모델에서 유일하게 바뀐 부분이다.
 *
 * 왜 필요해졌나: 예전에는 룸을 열 때 `seedGridStylesFromXlsx`가 워크북 전체를 훑으면서
 * `LazyGridStore.readCache`를 통째로 데워 놓았고, `joinEditSession`이 그 완료를 기다렸다. 그래서
 * 테스트가 페이지를 구독하는 시점에는 그 페이지의 모든 행이 **이미 메모리에** 있었고,
 * `seedGridPageFromXlsx`는 디스크를 단 한 번도 건드리지 않았다(마이크로태스크만으로 끝났다) —
 * 그 조건에서만 "setImmediate 턴 N번"이 충분한 대기였다. 서식 시딩이 콘텐츠와 함께 페이지 단위로
 * 윈도잉되면서 그 선점적 전량 데우기가 사라졌고, 이제 페이지를 처음 구독하면 실제로 디스크를
 * 읽는다.
 *
 * 그런데 setImmediate 턴은 **벽시계 시간을 전혀 사주지 않는다**: 실측하면 이 루프는 500턴을 돌아도
 * 약 3ms밖에 걸리지 않는 반면(턴마다 poll 타임아웃 0으로 그냥 스핀한다), 한 페이지(500행)를 읽는 데는
 * 실제로 약 10~25ms가 걸린다. 즉 턴 수를 아무리 늘려도 I/O를 기다리는 대기가 되지 못한다 —
 * 턴 수를 부풀리는 것은 해결책이 아니라 더 조용한 플래키의 원인이다(그래서 턴 수는 예전 그대로 둔다).
 *
 * `fs.stat`은 행 읽기와 **같은 libuv 스레드풀**(기본 4스레드)의 FIFO 작업 큐를 쓴다. 그래서 지금
 * 큐에 쌓여 있는 읽기들보다 뒤에 들어가고, 이 stat이 끝났다는 것은 그 앞의 읽기들이 이미 큐에서
 * 빠져나갔다는 뜻이다 — 임의의 sleep이 아니라 "지금까지 제출된 디스크 작업을 흘려보낸다"는
 * 의미가 분명한 대기다. 타이머(setTimeout)를 쓰지 않으므로 가짜 타이머를 켠 테스트에서도 안전하고,
 * setImmediate를 real로 남겨 둬야 하는 ExcelJS 제약과도 무관하다.
 */
async function drainPendingDiskIo(): Promise<void> {
  await fs.stat(os.tmpdir()).catch(() => undefined);
}

/**
 * 클라이언트가 보낸 Yjs 업데이트로 서버 쪽 맵 하나에 값을 쓴다 — 실제 편집 경로(소켓 메시지)를
 * 그대로 타므로 서버의 관찰자·origin 판정이 프로덕션과 동일하게 동작한다.
 *
 * clientId를 표현 가능한 최댓값 근처로 고정하는 이유와, 한 테스트에서 두 번 이상 부를 때 서로 다른
 * clientId를 줘야 하는 이유는 xlsxFormulaRecompute.test.ts의 sendCellEdit 주석 참고.
 */
function sendMapEdit(
  socket: FakeSocket,
  mapName: string,
  key: string,
  value: unknown,
  clientId = 0xffffffff,
): void {
  const scratch = new Y.Doc();
  scratch.clientID = clientId;
  scratch.getMap(mapName).set(key, value);
  const update = Y.encodeStateAsUpdate(scratch);
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeUpdate(encoder, update);
  socket.emit('message', encoding.toUint8Array(encoder));
}

/**
 * 이벤트 루프에 실제로 제어권이 돌아오는 횟수를 센다(tests/xlsxYielding.test.ts와 같은 패턴).
 * 측정 대상이 한 번도 안 쉬는 동기 블록이라면 이 재귀 티커는 그 블록이 끝날 때까지 단 한 번도
 * 돌지 못한다 — 그래서 "양보가 실제로 일어났는지"를 값의 정확성과 무관하게 직접 관찰할 수 있다.
 */
function startTicker(): { count: () => number; stop: () => number } {
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
      return ticks;
    },
  };
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

/**
 * 3000행(= 페이지 0~5) 시트 — A열은 숫자, **각 페이지마다 하나씩** B열에 수식이 있다(1, 501,
 * 1001, 1501, 2001, 2501행). 페이지마다 수식이 최소 하나 있어야 그 페이지의 재계산이 실제로
 * computed 맵에 쓰고(= Yjs 트랜잭션 하나 + 전파 하나를 만들고) 그래서 "페이지 사이에 양보가
 * 있었는지"를 트랜잭션 단위로 관찰할 수 있다(수식이 없는 페이지의 재계산은 아무것도 쓰지 않아
 * 빈 트랜잭션이 되고 'update'가 발생하지 않는다).
 */
async function makeFormulaOnEveryPageXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  for (let r = 1; r <= 3000; r += 1) sheet.getCell(r, 1).value = r;
  // 계산 결과(result)는 일부러 넣지 않는다 — 넣으면 parseXlsxToGrid가 그 캐시값을 셀 값으로 삼아
  // "수식 셀"이 아니게 되고 재계산 경로 자체를 검증할 수 없다.
  for (const row of [1, 501, 1001, 1501, 2001, 2501]) {
    sheet.getCell(row, 2).value = { formula: 'SUM(A1:A100)' } as ExcelJS.CellFormulaValue;
  }
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

/**
 * **한 페이지(500행) 안에** 수식이 고르게 흩어진 시트 — B열의 1, 126, 251, 376행(1-based)에 수식이
 * 있다. 페이지 하나를 재계산하는 루프가 페이지 안에서도 실제로 양보하는지(I3)를 관찰하기 위한
 * 픽스처다: 페이지 하나를 통째로 한 청크로 도는 구현에서는 이 네 수식이 전부 같은 트랜잭션·같은
 * 이벤트 루프 턴에서 계산되고, 페이지 내부에도 양보 지점이 있는 구현에서는 청크마다 따로 계산된다.
 */
async function makeFormulasWithinOnePageXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  for (let r = 1; r <= 500; r += 1) sheet.getCell(r, 1).value = r;
  // 계산 결과(result)는 일부러 넣지 않는다 — 넣으면 parseXlsxToGrid가 그 캐시값을 셀 값으로 삼아
  // "수식 셀"이 아니게 되고 재계산 경로 자체를 검증할 수 없다.
  for (const row of [1, 126, 251, 376]) {
    sheet.getCell(row, 2).value = { formula: 'SUM(A1:A100)' } as ExcelJS.CellFormulaValue;
  }
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

/** 1000행(= 페이지 0과 1) 시트 — A1에 원본 파일에서부터 굵게 서식이 걸려 있다. A2에는 아무 서식도 없다. */
async function makeStyledXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  for (let r = 1; r <= 1000; r += 1) sheet.getCell(r, 1).value = `row-${r}`;
  sheet.getCell(1, 1).font = { bold: true };
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

/**
 * 1000행(= 페이지 0과 1) 시트 — **A1에 원본 파일에서부터 서식이 여러 겹 걸려 있다**(굵게 +
 * 기울임 + 노란 배경). 서식 편집이 도착했을 때 "원래 있던 나머지 서식"이 살아남는지를 보려면
 * 편집으로 들어오는 속성(굵게) 말고도 잃을 것이 있어야 한다.
 */
async function makeRichlyStyledXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  for (let r = 1; r <= 1000; r += 1) sheet.getCell(r, 1).value = r === 1 ? 'title' : `row-${r}`;
  sheet.getCell(1, 1).font = { bold: true, italic: true };
  sheet.getCell(1, 1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFF00' },
  } as ExcelJS.FillPattern;
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
      await flushEventLoop();

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

      const { joinEditSession, getRoomLazyGridForTest, getSeededPagesForTest } = await import(
        '../src/collab/room.js'
      );
      const { getCell } = await import('../src/xlsx/lazyGridStore.js');

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
      const lazyGrid = getRoomLazyGridForTest(docId);
      expect(lazyGrid).not.toBeNull();
      expect(await getCell(lazyGrid!, 0, 0, 0)).toEqual({ v: 'row-1' });
      expect(await getCell(lazyGrid!, 0, 499, 0)).toEqual({ v: 'row-500' });
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

  test(
    'one edit while subscribed to many pages yields to the event loop between pages ' +
      '(no unyielded whole-window recompute block)',
    async () => {
      const { joinEditSession } = await import('../src/collab/room.js');
      const { GRID_CELLS_MAP, GRID_COMPUTED_MAP, gridCellKey } = await import('../src/collab/gridShared.js');

      const docId = randomUUID();
      await seedDoc(docId, await makeFormulaOnEveryPageXlsx());

      const socket = new FakeSocket();
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
      // 6개 페이지 전부 구독 — 편집 하나가 6페이지 분량의 재계산을 유발하는 상황을 만든다.
      sendSubscribe(socket, 0, [0, 1, 2, 3, 4, 5]);
      await flushEventLoop(300);

      // 페이지 재계산이 만드는 서버 트랜잭션 하나하나가 전파되어 mirror의 computed 맵을 바꾼다 —
      // 그때마다 "지금까지 이벤트 루프가 몇 번 돌았는지"를 기록한다. 페이지들이 양보 없이 연달아
      // 돌면 여러 기록이 **똑같은 틱 값**을 갖는다(그 사이 이벤트 루프가 한 번도 못 돌았다는 뜻).
      const ticker = startTicker();
      const ticksAtComputedWrite: number[] = [];
      socket.mirror.getMap<string>(GRID_COMPUTED_MAP).observe(() => {
        ticksAtComputedWrite.push(ticker.count());
      });

      // A1 편집 — 구독 중인 6페이지 전부의 재계산을 유발한다.
      sendMapEdit(socket, GRID_CELLS_MAP, gridCellKey(0, 0, 0), '999');
      await flushEventLoop(300);
      ticker.stop();

      // 페이지마다 수식이 하나씩 있으므로, 페이지들이 서로 다른 이벤트 루프 턴에서 돌았다면
      // 재계산 트랜잭션도 페이지 수만큼 나온다.
      //
      // 픽스 이전 코드(페이지마다 `void`)에서는 이 값이 **1**이었다(실측) — 여섯 페이지의 재계산이
      // 전부 편집 트랜잭션의 cleanup 안에서 연달아 동기 실행되는 바람에 Yjs가 그것을 트랜잭션
      // 하나로 합쳐 버리기 때문이다. 즉 이 숫자 자체가 "여섯 페이지가 한 호흡에 붙어 돌았다"는
      // 직접적인 증거다.
      expect(ticksAtComputedWrite.length).toBeGreaterThanOrEqual(6);

      // 같은 틱 값으로 연달아 기록된 최대 길이 = "양보 한 번 없이 연달아 실행된 페이지 재계산 수".
      // 페이지마다 void로 흘려보내던 예전 코드에서는 6개 페이지의 동기 블록이 전부 같은 틱에
      // 몰려(= 최대 길이 6) 이 값이 페이지 수만큼 커진다. 한 체인으로 await하면 페이지 사이마다
      // 이벤트 루프가 실제로 한 턴 돌아 이 값이 1로 떨어진다.
      let longestUnyieldedRun = 1;
      let currentRun = 1;
      for (let i = 1; i < ticksAtComputedWrite.length; i += 1) {
        currentRun = ticksAtComputedWrite[i] === ticksAtComputedWrite[i - 1] ? currentRun + 1 : 1;
        longestUnyieldedRun = Math.max(longestUnyieldedRun, currentRun);
      }
      expect(longestUnyieldedRun).toBeLessThanOrEqual(2);

      // 양보가 실제로 여러 번 일어났다는 직접 증거 — 페이지 수만큼(경계 5개 이상) 틱이 늘어났다.
      expect(ticksAtComputedWrite[ticksAtComputedWrite.length - 1] - ticksAtComputedWrite[0]).toBeGreaterThanOrEqual(5);
    },
    60_000,
  );

  test(
    'loading a page by scrolling (no user edit) does not trigger a save, but a real edit still does',
    async () => {
      vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

      const { store } = await import('../src/storage.js');
      const { joinEditSession } = await import('../src/collab/room.js');
      const { GRID_CELLS_MAP, gridCellKey } = await import('../src/collab/gridShared.js');

      const docId = randomUUID();
      await seedDoc(docId, await makeCrossPageFormulaXlsx()); // 페이지 1에 수식이 있는 1000행 시트

      const socket = new FakeSocket();
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
      await flushEventLoop();

      const saveSpy = vi.spyOn(store, 'saveOriginal');

      // 스크롤로 페이지 1을 열었을 뿐 — 사용자 편집은 하나도 없다. 이 구독이 페이지 시딩과 그
      // 페이지의 수식 재계산(computed 맵 쓰기)을 유발한다.
      sendSubscribe(socket, 0, [1]);
      await flushEventLoop(200);
      await vi.advanceTimersByTimeAsync(8_100); // 유휴 자동저장 디바운스를 한참 넘긴다
      // 저장 파이프라인(ExcelJS 로드 → export → saveOriginal)은 실제 이벤트 루프 턴을 꽤 많이
      // 먹는다 — 여기서 턴을 아끼면 "아직 도달하지 못한 저장"을 "저장이 예약되지 않았다"로
      // 오판해 이 테스트가 무의미하게 통과한다(실측으로 확인: 50턴에서는 픽스 없이도 통과했다).
      await flushEventLoop(400);

      // computed 맵은 저장 대상이 아니다(exportXlsxFromGrid는 canonicalGrid/baselineGrid만 읽는다)
      // — 그 맵에 쓰였다는 이유로 전체 xlsx export + 재파싱 + 뷰어 캐시 재생성이 돌면 안 된다.
      expect(saveSpy).not.toHaveBeenCalled();

      // 반대로 진짜 편집은 여전히 저장을 예약해야 한다 — 억제가 필요 이상으로 넓어지지 않았는지
      // 함께 검증한다(계산값 쓰기만 면제되고 셀 쓰기는 그대로).
      sendMapEdit(socket, GRID_CELLS_MAP, gridCellKey(0, 600, 0), 'edited');
      await flushEventLoop(200);
      await vi.advanceTimersByTimeAsync(8_100);
      await flushEventLoop(400);

      expect(saveSpy).toHaveBeenCalled();
    },
    60_000,
  );

  test(
    'styles are now windowed with content: not seeded before subscribe, unseeded with the page, ' +
      'and correctly restored (both original and user-edited styles) on reseed',
    async () => {
      vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

      const { joinEditSession, getSeededPagesForTest } = await import('../src/collab/room.js');
      const { GRID_STYLES_MAP, gridCellKey } = await import('../src/collab/gridShared.js');

      const docId = randomUUID();
      await seedDoc(docId, await makeStyledXlsx()); // A1은 원본부터 굵게, 1000행(페이지 0·1)

      const socket = new FakeSocket();
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);

      // 아직 아무 페이지도 구독하지 않은 시점 — 서식도 콘텐츠처럼 아직 전혀 채워지지 않아야 한다
      // (룸을 열 때 워크북 전체를 훑어 서식을 미리 채우던 옛 seedGridStylesFromXlsx가 사라졌다는
      // 직접적인 증거 — 이게 이 세션이 고친 룸 열기 지연의 핵심이다).
      socket.requestSync();
      await flushEventLoop();
      const styles = socket.mirror.getMap<{ b?: 1; i?: 1 }>(GRID_STYLES_MAP);
      expect(styles.size).toBe(0);

      sendSubscribe(socket, 0, [0]); // 서식이 있는 셀이 속한 페이지를 구독
      await flushEventLoop();
      socket.requestSync();

      // 구독한 페이지의 서식은 콘텐츠와 같은 시점에 함께 도착한다.
      expect(styles.get(gridCellKey(0, 0, 0))?.b).toBe(1);

      // 사용자가 A2에 기울임을 적용한다 — 클라이언트가 GRID_STYLES_MAP에 직접 쓰는 실제 경로.
      // room.ts의 styles.observe → applyStylesChangeToCanonical이 이 편집을 lazyGrid.dirtyOverlay에
      // 반영해 두므로, 아래에서 이 페이지가 비워졌다 다시 채워져도 이 편집이 사라지면 안 된다.
      sendMapEdit(socket, GRID_STYLES_MAP, gridCellKey(0, 1, 0), { i: 1 });
      await flushEventLoop();
      expect(styles.get(gridCellKey(0, 1, 0))?.i).toBe(1);

      // 페이지 0에서 스크롤해 나간 뒤 유예 시간을 넘긴다 → 페이지 0이 실제로 비워진다.
      sendSubscribe(socket, 0, [1]);
      await flushEventLoop();
      await vi.advanceTimersByTimeAsync(10_100);
      await flushEventLoop();
      expect(getSeededPagesForTest(docId, 0).has(0)).toBe(false); // 셀 데이터는 정말 비워졌다
      // 서식도 콘텐츠와 함께 정말로 비워졌다 — 예전처럼 "서식은 절대 안 지운다"가 아니라는 것을
      // 직접 확인한다(윈도잉이 실제로 일어났다는 증거, 회귀하면 이 값이 여전히 남아 있을 것이다).
      expect(styles.get(gridCellKey(0, 0, 0))).toBeUndefined();
      expect(styles.get(gridCellKey(0, 1, 0))).toBeUndefined();

      // 다시 스크롤해 돌아온다 → 페이지 0이 다시 채워진다.
      sendSubscribe(socket, 0, [0, 1]);
      await flushEventLoop();
      socket.requestSync();
      expect(getSeededPagesForTest(docId, 0).has(0)).toBe(true);

      // 원본 파일의 서식도, 사용자가 적용한 서식도 정확히 복원돼야 한다 — 콘텐츠와 똑같이
      // dirtyOverlay를 거쳐 가는 resolveGridCell 덕분에, 일시적으로 지워졌던 서식이 원본으로
      // 되돌아가지 않고 편집된 그대로 되살아난다(데이터 손실 없음).
      expect(styles.get(gridCellKey(0, 0, 0))?.b).toBe(1);
      expect(styles.get(gridCellKey(0, 1, 0))?.i).toBe(1);
    },
    60_000,
  );

  test(
    'applying a style to a cell whose page has NOT been seeded yet preserves the cell\'s original formatting ' +
      '(the client\'s read-modify-write starts from an empty view — the server must not take that as "no style")',
    async () => {
      vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

      const { joinEditSession, getRoomLazyGridForTest, getCurrentBytes } = await import('../src/collab/room.js');
      const { GRID_STYLES_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
      const { getCell } = await import('../src/xlsx/lazyGridStore.js');

      const docId = randomUUID();
      await seedDoc(docId, await makeRichlyStyledXlsx());

      const socket = new FakeSocket();
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
      await flushEventLoop();

      // **일부러 아무 페이지도 구독하지 않는다** — 서식이 콘텐츠와 함께 윈도잉된 뒤로 생긴
      // 실제 경쟁 구간을 그대로 재현한다: 브라우저에는 그 칸이 이미 그려져 있는데 서버가 아직
      // 그 페이지를 Yjs에 채우지 않은 순간. 그때 클라이언트의 읽고-고쳐-쓰기
      // (`const current = styles.get(key) ?? {}` → 속성 하나 얹기 → `styles.set`)는 **빈 객체**를
      // 바탕으로 굴러가므로, 서버에 도착하는 것은 "굵게" 하나뿐이다.
      expect(socket.mirror.getMap(GRID_STYLES_MAP).size).toBe(0);
      sendMapEdit(socket, GRID_STYLES_MAP, gridCellKey(0, 0, 0), { b: 1 });
      await flushEventLoop();

      // 서버는 그 편집을 **canonical에 이미 있는 서식 위에** 얹어야 한다 — 통째로 갈아치우면
      // 원본의 배경색·기울임·표시 형식이 조용히 사라지고, 그 결과가 dirtyOverlay를 타고
      // **저장되는 파일까지** 간다(이 프로젝트가 절대 용납하지 않는 원본 유실).
      const lazyGrid = getRoomLazyGridForTest(docId)!;
      const cell = await getCell(lazyGrid, 0, 0, 0);
      expect(cell).not.toBe(0);
      const styleIndex = (cell as { s?: number }).s;
      expect(styleIndex).toBeDefined();
      const style = lazyGrid.styles[styleIndex!];
      expect(style.b).toBe(1); // 방금 적용한 굵게
      expect(style.i).toBe(1); // 원본 기울임 — 예전 구현에서는 여기서 사라졌다
      expect(style.bg?.toUpperCase()).toContain('FFFF00'); // 원본 배경색

      // 저장 파일까지 확인한다 — 이 결함은 "화면은 멀쩡한데 파일에서 서식이 사라지는" 모양이다.
      const bytes = await getCurrentBytes(docId);
      expect(bytes).not.toBeNull();
      const reopened = new ExcelJS.Workbook();
      await reopened.xlsx.load(bytes! as unknown as ArrayBuffer);
      const saved = reopened.getWorksheet('Sheet1')!.getCell('A1');
      expect(saved.value).toBe('title');
      expect(saved.font?.bold).toBe(true);
      expect(saved.font?.italic).toBe(true);
      expect((saved.fill as ExcelJS.FillPattern | undefined)?.fgColor?.argb?.toUpperCase()).toContain('FFFF00');
    },
    60_000,
  );

  test(
    'removing one attribute (e.g. un-bolding) on an already-seeded cell still removes exactly that attribute ' +
      '(the canonical merge must not resurrect what the user just took off)',
    async () => {
      vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

      const { joinEditSession, getRoomLazyGridForTest, getCurrentBytes } = await import('../src/collab/room.js');
      const { GRID_STYLES_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
      const { getCell } = await import('../src/xlsx/lazyGridStore.js');

      const docId = randomUUID();
      await seedDoc(docId, await makeRichlyStyledXlsx()); // A1 = 굵게 + 기울임 + 노란 배경

      const socket = new FakeSocket();
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
      sendSubscribe(socket, 0, [0]); // 이번에는 페이지가 **이미 시딩된** 통상 경로
      await flushEventLoop();
      socket.requestSync();

      const styles = socket.mirror.getMap<{ b?: 1; i?: 1; bg?: string }>(GRID_STYLES_MAP);
      const seen = styles.get(gridCellKey(0, 0, 0))!;
      expect(seen.b).toBe(1);

      // 툴바의 굵게 버튼을 한 번 더 누른 것 = 클라이언트가 자기 시야에서 b만 뺀 객체를 쓴다.
      const { b: _removed, ...withoutBold } = seen;
      sendMapEdit(socket, GRID_STYLES_MAP, gridCellKey(0, 0, 0), withoutBold);
      await flushEventLoop();

      const lazyGrid = getRoomLazyGridForTest(docId)!;
      const cell = await getCell(lazyGrid, 0, 0, 0);
      const style = lazyGrid.styles[(cell as { s?: number }).s!];
      expect(style.b).toBeFalsy(); // 정말로 해제됐다(병합이 되살리지 않았다)
      expect(style.i).toBe(1); // 건드리지 않은 나머지는 그대로
      expect(style.bg?.toUpperCase()).toContain('FFFF00');

      const bytes = await getCurrentBytes(docId);
      const reopened = new ExcelJS.Workbook();
      await reopened.xlsx.load(bytes! as unknown as ArrayBuffer);
      const saved = reopened.getWorksheet('Sheet1')!.getCell('A1');
      expect(saved.font?.bold).toBeFalsy();
      expect(saved.font?.italic).toBe(true);
    },
    60_000,
  );

  test(
    'recomputing a SINGLE page yields to the event loop partway through it ' +
      '(PAGE_SIZE_ROWS < YIELD_EVERY_N_ROWS must not mean "one page = one unyielded chunk")',
    async () => {
      const { joinEditSession } = await import('../src/collab/room.js');
      const { GRID_CELLS_MAP, GRID_COMPUTED_MAP, gridCellKey } = await import('../src/collab/gridShared.js');

      const docId = randomUUID();
      await seedDoc(docId, await makeFormulasWithinOnePageXlsx());

      const socket = new FakeSocket();
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
      sendSubscribe(socket, 0, [0]); // 페이지 **하나만** 구독 — 페이지 사이 양보(Task 7)는 여기서 일어날 수 없다
      await flushEventLoop(200);

      const ticker = startTicker();
      const ticksAtComputedWrite: number[] = [];
      socket.mirror.getMap<string>(GRID_COMPUTED_MAP).observe(() => {
        ticksAtComputedWrite.push(ticker.count());
      });

      // A1 편집 — 이 한 페이지 안의 수식 네 개가 전부 다시 계산된다.
      sendMapEdit(socket, GRID_CELLS_MAP, gridCellKey(0, 0, 0), '999');
      await flushEventLoop(200);
      ticker.stop();

      // 페이지 하나를 통째로 한 청크로 도는 예전 구현에서는 네 수식이 트랜잭션 하나에 몰려
      // 이 값이 1이었다 — 즉 500행 전체가 양보 한 번 없는 동기 블록이었다는 직접 증거.
      expect(ticksAtComputedWrite.length).toBeGreaterThanOrEqual(4);

      // 그 트랜잭션들이 서로 다른 이벤트 루프 턴에서 일어났는지(= 페이지 중간에 실제로 쉬었는지).
      expect(new Set(ticksAtComputedWrite).size).toBeGreaterThanOrEqual(4);
    },
    60_000,
  );
});

/**
 * Task 8 — seedGridPageFromXlsx가 GridDocument 대신 LazyGridStore를 읽는 새 시그니처로 동작하는지
 * 직접(room.ts를 거치지 않고) 검증한다. room.ts는 아직 canonicalGrid(GridDocument)를 임시
 * 어댑터로 감싸 이 함수를 호출하므로(Task 14 이전), 위 describe 블록의 room 통합 테스트들은 이
 * 태스크 이후에도 정확히 같은 결과를 내야 한다(회귀 없음) — 이 블록은 그 어댑터를 거치지 않고
 * 실제 LazyGridStore(ingestXlsxLazily)를 바로 물려 "디스크에서 처음 읽는 캐시 미스 경로"까지
 * 실제로 타는지 확인한다.
 */
describe('xlsx/collab.ts: seedGridPageFromXlsx reads through LazyGridStore', () => {
  test('subscribing a page fills exactly that page\'s cells, including a row not yet warmed in the read cache (forces a disk read)', async () => {
    const { ingestXlsxLazily, closeLazyGridStore } = await import('../src/xlsx/lazyGridStore.js');
    const { seedGridPageFromXlsx } = await import('../src/xlsx/collab.js');
    const { GRID_CELLS_MAP, gridCellKey } = await import('../src/collab/gridShared.js');

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    for (let r = 1; r <= 600; r += 1) sheet.getCell(r, 1).value = `row-${r}`; // 페이지 0(0~499)과 페이지 1의 일부
    const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);

    const store = await ingestXlsxLazily(randomUUID(), bytes);
    // ingestion만으로는 읽기 캐시가 데워지지 않는다 — 아래 seedGridPageFromXlsx 호출이 실제로
    // 디스크에서 처음 읽어야 한다(캐시 히트가 아니라 미스 경로를 강제로 태운다).
    expect(store.readCache.size).toBe(0);

    try {
      const ydoc = new Y.Doc();
      await seedGridPageFromXlsx(ydoc, store, 0, 0); // 페이지 0(0~499행)만 구독

      const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
      expect(cells.get(gridCellKey(0, 0, 0))).toBe('row-1');
      expect(cells.get(gridCellKey(0, 499, 0))).toBe('row-500'); // 페이지 경계 마지막 행 — 디스크에서 새로 읽힘
      expect(cells.get(gridCellKey(0, 500, 0))).toBeUndefined(); // 다음 페이지(1번) — 구독하지 않았으므로 없어야 함
      expect(cells.size).toBe(500);
      // 방금 읽은 페이지 0의 행들은 이제 캐시에 남아 있어야 한다(다음 구독에서 재사용).
      expect(store.readCache.size).toBeGreaterThan(0);
    } finally {
      await closeLazyGridStore(store);
    }
  }, 30_000);

  test('seedGridPageFromXlsx fills that page\'s styles together with its content, and none outside the page', async () => {
    const { ingestXlsxLazily, closeLazyGridStore } = await import('../src/xlsx/lazyGridStore.js');
    const { seedGridPageFromXlsx } = await import('../src/xlsx/collab.js');
    const { GRID_STYLES_MAP, gridCellKey } = await import('../src/collab/gridShared.js');

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    for (let r = 1; r <= 600; r += 1) sheet.getCell(r, 1).value = `row-${r}`;
    sheet.getCell(1, 1).font = { bold: true }; // 페이지 0 안(0-based 행 0)
    // 굵게 서식을 페이지 0 밖(600행, 0-based 599 = 페이지 1)에도 둔다 — 구독하지 않은 페이지의
    // 서식은 채워지지 않아야(예전의 전량 선점 시딩이 사라졌어야) 한다는 것을 증명하기 위함.
    sheet.getCell(600, 1).font = { bold: true };
    const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);

    const store = await ingestXlsxLazily(randomUUID(), bytes);

    try {
      const ydoc = new Y.Doc();
      await seedGridPageFromXlsx(ydoc, store, 0, 0); // 페이지 0(0~499행)만 구독

      const styles = ydoc.getMap<{ b?: 1 }>(GRID_STYLES_MAP);
      // 구독한 페이지 안의 서식은 콘텐츠와 정확히 같은 호출로, 추가 디스크 스캔 없이 채워진다
      // (이게 이번 성능 수정의 핵심 — 예전에는 이 서식이 별도의 seedGridStylesFromXlsx가 워크북
      // 전체를 다시 훑어야만 나왔다).
      expect(styles.get(gridCellKey(0, 0, 0))?.b).toBe(1);
      // 구독하지 않은 페이지(1번)의 서식은 아직 없어야 한다 — 전량 선점 시딩이 사라졌다는 직접
      // 증거.
      expect(styles.get(gridCellKey(0, 599, 0))).toBeUndefined();
      // 서식 없는 칸은 여전히 안 심긴다(희소 유지) — 예전 동작과 동일.
      expect(styles.get(gridCellKey(0, 1, 0))).toBeUndefined();

      // 페이지 1을 나중에 구독하면 그 페이지의 서식도 같은 경로로, 그제서야 채워진다.
      await seedGridPageFromXlsx(ydoc, store, 0, 1);
      expect(styles.get(gridCellKey(0, 599, 0))?.b).toBe(1);
    } finally {
      await closeLazyGridStore(store);
    }
  }, 30_000);
});
