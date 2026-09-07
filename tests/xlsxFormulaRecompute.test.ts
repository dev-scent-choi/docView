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
  test('formula cells are computed once right after seeding, before any edit', async () => {
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

    // A3 = SUM(A1:A2) = 30. 편집이 하나도 없었으므로 이 값은 시딩 직후의 최초 재계산에서만
    // 나올 수 있다.
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

    // A1을 100으로 편집 → SUM(A1:A2)은 120이 되어야 한다.
    sendCellEdit(socket, gridCellKey(0, 0, 0), '100');

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

    let computed = socket.mirror.getMap<string>(GRID_COMPUTED_MAP);
    expect(computed.get(gridCellKey(0, 2, 0))).toBe('30');

    // A3(=SUM(A1:A2))을 더 이상 수식이 아닌 고정 텍스트로 덮어쓴다.
    sendCellEdit(socket, gridCellKey(0, 2, 0), 'not a formula anymore');

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
      // 재계산이 최소 한 번은 실제 setImmediate로 양보하도록(YIELD_EVERY_N_ROWS=2000을 넘도록)
      // 큰 시트를 만든다. B1은 A1:A3000 전체를 더하는 수식 셀 — 재계산 도중 겹쳐 들어오는 두
      // 번째 편집이 최종 결과에 정확히 반영되는지 함께 검증한다.
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

      const saveSpy = vi.spyOn(store, 'saveOriginal');

      // 편집 1: A1을 1 → 999로. 이 편집의 cells.observe 콜백이 recomputeFormulasForSheet를
      // 시작시키고, 그 함수는 실제 setImmediate로 양보하므로(3000행짜리 시트) 아직 그 양보를
      // 기다리는 동안(같은 동기 실행 구간 안에서) 편집 2를 곧바로 보낸다 — "재계산이 아직
      // 진행 중일 때 도착하는 진짜 편집"을 재현한다.
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
