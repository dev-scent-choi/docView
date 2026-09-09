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
import { computeGridCellDisplayAsync } from '../src/xlsx/formulas.js';
import { ingestXlsxLazily, setDirtyCell } from '../src/xlsx/lazyGridStore.js';

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

/**
 * Task 7(윈도잉 컷오버) 이후 수식 재계산은 **구독 중인 페이지**에 대해서만 일어난다 — 룸을 열기만
 * 하고 아무 페이지도 구독하지 않으면 Yjs에는 셀도 계산값도 들어오지 않는다(그게 윈도잉의 요점).
 * 그래서 이 파일의 모든 테스트는 실제 클라이언트와 마찬가지로 "지금 보고 있는 페이지"를 먼저
 * 알려준다. 여기 쓰이는 시트는 전부 500행 미만이라 페이지 0 하나가 시트 전체를 덮는다.
 */
function sendSubscribe(socket: FakeSocket, sheet: number, pages: readonly number[]): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_PAGE_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify({ sheet, pages }));
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** 페이지 채우기·재계산은 청크마다 실제 setImmediate로 양보하므로 그 턴을 흘려보내야 한다. */
async function flushEventLoop(turns = 50): Promise<void> {
  for (let i = 0; i < turns; i += 1) await new Promise((resolve) => setImmediate(resolve));
}

function sendCellEdit(socket: FakeSocket, key: string, value: string, clientId = 0xffffffff): void {
  // scratch는 room의 Y.Doc과 인과관계를 공유하지 않는 완전히 새 Y.Doc이므로, 이미 값이 있는
  // 같은 키에 값을 쓰면 Y.Map의 동시편집 충돌 해소 규칙(더 큰 clientID가 이김)에 따라 결과가
  // 무작위로 갈린다 — clientID를 표현 가능한 최댓값 근처로 고정해 "이 편집이 항상 이긴다"를
  // 결정적으로 만든다(roomAutosaveLoop.test.ts, roomCanonicalGrid.test.ts와 동일한 패턴).
  // 한 테스트 안에서 sendCellEdit를 두 번 이상 부를 때는 반드시 서로 다른 clientId를 줘야 한다
  // — 같은 clientId를 가진 서로 다른 scratch Y.Doc 두 개가 각자 clock 0부터 새로 시작해 같은
  // 대상 ydoc에 병합되면, 대상의 상태 벡터가 이미 그 clientId의 clock을 앞서 있어 두 번째
  // scratch의 아이템이 "이미 통합된 것"으로 오인되어 조용히 버려진다(실측으로 확인됨).
  const scratch = new Y.Doc();
  scratch.clientID = clientId;
  scratch.getMap<string>('cells').set(key, value);
  const update = Y.encodeStateAsUpdate(scratch);
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeUpdate(encoder, update);
  socket.emit('message', encoding.toUint8Array(encoder));
}

describe('room.ts: formula cells recompute into GRID_COMPUTED_MAP', () => {
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

  // 이 파일 첫 테스트는 fast-formula-parser 등 무거운 의존성 그래프를 콜드 스타트로 처음
  // 임포트하는 비용을 떠안아 기본 5초 타임아웃에 걸릴 수 있다(격리 실행 시엔 1초 미만) —
  // 넉넉히 20초로 늘려둔다.
  test('formula cells are computed as soon as their page is subscribed, before any edit', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession } = await import('../src/collab/room.js');
    const { gridCellKey, GRID_COMPUTED_MAP } = await import('../src/collab/gridShared.js');

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
    sendSubscribe(socket, 0, [0]);
    await flushEventLoop();

    // A3 = SUM(A1:A2) = 30. 편집이 하나도 없었으므로 이 값은 페이지가 처음 구독될 때 셀 데이터와
    // 함께 채워지는 최초 재계산에서만 나올 수 있다.
    const computed = socket.mirror.getMap<string>(GRID_COMPUTED_MAP);
    expect(computed.get(gridCellKey(0, 2, 0))).toBe('30');
  }, 20_000);

  test("editing a cell that a formula depends on updates the formula cell's computed value", async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession } = await import('../src/collab/room.js');
    const { gridCellKey, GRID_COMPUTED_MAP } = await import('../src/collab/gridShared.js');

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
    sendSubscribe(socket, 0, [0]);
    await flushEventLoop();

    // A1을 100으로 편집 → SUM(A1:A2)은 120이 되어야 한다.
    sendCellEdit(socket, gridCellKey(0, 0, 0), '100');
    await flushEventLoop();

    const computed = socket.mirror.getMap<string>(GRID_COMPUTED_MAP);
    expect(computed.get(gridCellKey(0, 2, 0))).toBe('120');
  }, 20_000);

  test('overwriting a formula cell with a plain value deletes its stale computed entry', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession } = await import('../src/collab/room.js');
    const { gridCellKey, GRID_COMPUTED_MAP } = await import('../src/collab/gridShared.js');

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
    sendSubscribe(socket, 0, [0]);
    await flushEventLoop();

    let computed = socket.mirror.getMap<string>(GRID_COMPUTED_MAP);
    expect(computed.get(gridCellKey(0, 2, 0))).toBe('30');

    // A3(=SUM(A1:A2))을 더 이상 수식이 아닌 고정 텍스트로 덮어쓴다.
    sendCellEdit(socket, gridCellKey(0, 2, 0), 'not a formula anymore');
    await flushEventLoop();

    computed = socket.mirror.getMap<string>(GRID_COMPUTED_MAP);
    expect(computed.has(gridCellKey(0, 2, 0))).toBe(false);
  }, 20_000);

  test(
    "a real edit arriving while a previous edit's recompute is still yielding still schedules " +
      'autosave, converges to the correct computed value, and does not cause a perpetual resave loop',
    async () => {
      vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

      const { store } = await import('../src/storage.js');
      const { joinEditSession } = await import('../src/collab/room.js');
      const { gridCellKey, GRID_COMPUTED_MAP } = await import('../src/collab/gridShared.js');

      const docId = randomUUID();
      // 재계산이 여러 번 실제 setImmediate로 양보하도록 여러 페이지(3000행 = 500행짜리 6페이지)
      // 분량의 시트를 만든다. B1은 A1:A3000 전체를 더하는 수식 셀 — 재계산 도중 겹쳐 들어오는 두
      // 번째 편집이 최종 결과에 정확히 반영되는지, 그리고 그 수식이 (창 안이긴 하지만) 자기
      // 페이지 밖의 2500행을 참조해도 canonical 기준으로 정확히 계산되는지 함께 검증한다.
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Sheet1');
      for (let r = 1; r <= 3000; r += 1) sheet.getCell(r, 1).value = r;
      sheet.getCell('B1').value = { formula: 'SUM(A1:A3000)' };
      const original = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
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
      // 3000행 = 페이지 0~5. 여섯 페이지를 모두 구독해 시트 전체를 창 안에 둔다 — 그래야 편집
      // 하나가 유발하는 재계산이 여러 페이지에 걸쳐 여러 번 양보하고, 그 양보 틈에 다음 편집이
      // 끼어드는 상황(이 테스트의 핵심)이 실제로 재현된다.
      sendSubscribe(socket, 0, [0, 1, 2, 3, 4, 5]);
      await flushEventLoop(200);

      const saveSpy = vi.spyOn(store, 'saveOriginal');

      // 편집 1: A1을 1 → 999로. 이 편집의 cells.observe 콜백이 구독된 페이지마다
      // recomputeFormulasForPage를 시작시키고, 그 함수들은 실제 setImmediate로 양보하므로 아직
      // 그 양보를 기다리는 동안(같은 동기 실행 구간 안에서) 편집 2를 곧바로 보낸다 — "재계산이
      // 아직 진행 중일 때 도착하는 진짜 편집"을 재현한다.
      sendCellEdit(socket, gridCellKey(0, 0, 0), '999', 0xffffffff);
      sendCellEdit(socket, gridCellKey(0, 1, 0), '999', 0xfffffffe);

      // 두 재계산이 각자의 양보 지점을 모두 통과해 끝날 때까지 실제 setImmediate 턴을 충분히
      // 흘려보낸다(가짜 타이머는 setImmediate에 영향을 주지 않는다).
      for (let i = 0; i < 200; i += 1) await new Promise((resolve) => setImmediate(resolve));

      // 8초 유휴 자동저장 대기 — 재계산 도중 도착한 편집 2의 update 이벤트가
      // suppressAutosaveReschedule=true인 상태에서 처리되었다면(브리프 원안의 "재계산 전체를
      // 감싼다" 실수) 이 시점에 자동저장이 걸리지 않았을 것이다.
      await vi.advanceTimersByTimeAsync(8_100);
      for (let i = 0; i < 200; i += 1) await new Promise((resolve) => setImmediate(resolve));

      expect(saveSpy).toHaveBeenCalled();
      const callsAfterFirstSave = saveSpy.mock.calls.length;

      // 최종 계산값: A1=999, A2=999로 바뀐 뒤의 SUM(A1:A3000). 원래 합(1+...+3000=4,501,500)에서
      // 1과 2를 999, 999로 바꾼 차이만큼 더한 값이어야 한다 — 겹쳐 실행된 두 재계산이 서로를
      // 밀어내지 않고 최종적으로 두 편집을 모두 반영해 수렴했는지 검증한다.
      const computed = socket.mirror.getMap<string>(GRID_COMPUTED_MAP);
      expect(computed.get(gridCellKey(0, 0, 1))).toBe('4503495');

      // 재계산 자신의 computed 맵 쓰기(매 청크마다 ydoc 'update'를 발생시킨다)가 스스로 또
      // 자동저장을 예약해 실제 편집 없이도 영원히 재저장이 반복되는 루프를 만들지 않아야 한다
      // (이 세션에서 겪은 자동저장 무한 루프와 같은 클래스의 버그, roomAutosaveLoop.test.ts 참고).
      await vi.advanceTimersByTimeAsync(30_000);
      for (let i = 0; i < 50; i += 1) await new Promise((resolve) => setImmediate(resolve));
      expect(saveSpy.mock.calls.length).toBe(callsAfterFirstSave);
    },
    20_000,
  );
});

describe('xlsx/formulas.ts: computeGridCellDisplayAsync (LazyGridStore, cache-miss retry)', () => {
  test('a formula referencing a row far outside any cached page is still evaluated correctly (forces at least one cache-miss retry)', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.getCell('A1').value = { formula: 'A1000' }; // 1000번째 행을 참조
    sheet.getCell('A1000').value = 42;
    const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);

    const store = await ingestXlsxLazily(randomUUID(), bytes);
    // A1의 원본 텍스트는 파싱 시 '=A1000'으로 저장돼 있어야 함(이미 검증된 기존 규칙) — 캐시는
    // 비어 있는 상태로 시작(강제로 미스가 나게 함).
    expect(store.readCache.size).toBe(0);

    const result = await computeGridCellDisplayAsync(store, 0, 0, 0);
    expect(result).toBe(42);

    await fs.rm(store.rowStorePath, { force: true });
  });

  test('IF() with a value-dependent branch resolves only the branch actually taken, without prefetching the other branch', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.getCell('A1').value = 1; // 조건이 참이 되게
    sheet.getCell('B1').value = { formula: 'IF(A1>0, C1, D1)' };
    sheet.getCell('C1').value = 'yes';
    sheet.getCell('D1').value = 'no';
    const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);

    const store = await ingestXlsxLazily(randomUUID(), bytes);
    const result = await computeGridCellDisplayAsync(store, 0, 0, 1); // B1
    expect(result).toBe('yes');

    await fs.rm(store.rowStorePath, { force: true });
  });
});
