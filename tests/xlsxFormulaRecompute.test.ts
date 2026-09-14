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
    'a formula referencing a row outside the subscribed page still recomputes correctly after an edit',
    async () => {
      const { store } = await import('../src/storage.js');
      const { joinEditSession } = await import('../src/collab/room.js');
      const { gridCellKey, GRID_COMPUTED_MAP } = await import('../src/collab/gridShared.js');

      const docId = randomUUID();
      // A1000(원본 행 999)은 페이지 1(500~999행)에 속해 페이지 0만 구독하면 Yjs 창 밖에 남는다.
      // B1의 수식은 그래도 canonical 전체를 기준으로 정확히 계산돼야 한다
      // (recomputeFormulasForPage가 computeGridCellDisplayAsync/LazyGridStore 어댑터를 거쳐도
      // 구독 여부와 무관하게 canonicalGrid 전체를 읽는다는 기존 동작이 유지되는지 검증).
      // 모든 행(1~1000)을 밀도 있게 채운다 — parseXlsxToGrid가 rowCount를
      // sheet.actualRowCount(= 실제로 값이 있는 행의 "개수")로 정하므로, 듬성듬성(1행과 1000행만)
      // 채우면 actualRowCount가 2가 되어 1000번째 행이 애초에 grid에 들어오지 않는다(위 4번째
      // 테스트가 3000행을 모두 채우는 것과 같은 이유).
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Sheet1');
      for (let r = 1; r <= 1000; r += 1) sheet.getCell(r, 1).value = r;
      sheet.getCell('B1').value = { formula: 'A1+A1000' };
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
      // 페이지 0만 구독한다 — A1000(페이지 1)은 창 밖에 남는다.
      sendSubscribe(socket, 0, [0]);
      await flushEventLoop();

      let computed = socket.mirror.getMap<string>(GRID_COMPUTED_MAP);
      // A1(=1) + A1000(=1000) = 1001
      expect(computed.get(gridCellKey(0, 0, 1))).toBe('1001');

      // A1을 1 → 100으로 편집(페이지 0 안, 구독 중). B1은 여전히 창 밖의 A1000을 정확히 참조해
      // 재계산돼야 한다.
      sendCellEdit(socket, gridCellKey(0, 0, 0), '100');
      await flushEventLoop();

      computed = socket.mirror.getMap<string>(GRID_COMPUTED_MAP);
      // 100 + 1000 = 1100
      expect(computed.get(gridCellKey(0, 0, 1))).toBe('1100');
    },
    20_000,
  );

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

  // 실제 버그 리포트의 시나리오 그대로: 페이지 하나(500행)에 체인형(cascading) 수식 컬럼 하나
  // (러닝 합계, B{r} = B{r-1} + A{r}) — 브리프에 따르면 memo 없이는 90초 안에도 끝나지 않던 경우.
  // room.ts의 실제 프로덕션 경로(joinEditSession → 페이지 구독 → recomputeFormulasForPage)를 그대로
  // 거치므로, recomputeFormulasForPage가 실제로 페이지 하나 전체에 걸쳐 memo를 공유하도록 배선돼
  // 있는지(단위 테스트가 아니라 배선 자체)를 검증하는 종단 간(end-to-end) 회귀 테스트다.
  test(
    'end-to-end: subscribing to a page with 500 rows of a chained running-total formula computes ' +
      'all values quickly and correctly (matches the real page-load bug report)',
    async () => {
      const { store } = await import('../src/storage.js');
      const { joinEditSession } = await import('../src/collab/room.js');
      const { gridCellKey, GRID_COMPUTED_MAP } = await import('../src/collab/gridShared.js');

      const ROWS = 500;
      const docId = randomUUID();
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Sheet1');
      for (let r = 1; r <= ROWS; r += 1) {
        sheet.getCell(r, 1).value = r; // A{r} = r
        sheet.getCell(r, 2).value = r === 1 ? { formula: 'A1' } : { formula: `B${r - 1}+A${r}` }; // B{r}
      }
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

      const start = Date.now();
      sendSubscribe(socket, 0, [0]); // 500행 = 페이지 0 하나 전체
      await flushEventLoop(300);
      const elapsedMs = Date.now() - start;

      const computed = socket.mirror.getMap<string>(GRID_COMPUTED_MAP);
      // 정확성: B{r} = 1+2+...+r = r(r+1)/2.
      expect(computed.get(gridCellKey(0, 0, 1))).toBe('1'); // B1
      expect(computed.get(gridCellKey(0, 9, 1))).toBe('55'); // B10
      expect(computed.get(gridCellKey(0, 99, 1))).toBe('5050'); // B100
      expect(computed.get(gridCellKey(0, ROWS - 1, 1))).toBe(String((ROWS * (ROWS + 1)) / 2)); // B500 = 125250

      // 속도: 브리프의 실측(같은 시나리오가 90초 넘게도 끝나지 않았음)과 대비되는 넉넉한 상한.
      // 전체 테스트 스위트를 병렬로 돌릴 때의 CPU 경합까지 고려해 여유를 넉넉히 둔다(격리 실행
      // 시엔 1.5초 안팎) — 그래도 O(N²) 상태로 되돌아가면 절대 넘길 수 없는 상한이라는 게 핵심.
      expect(elapsedMs).toBeLessThan(15_000);
    },
    30_000,
  );

  // da53baf(FormulaParser 풀링)가 실제로 고치려던 원래 버그 리포트의 패턴: 서로 참조하지 않는
  // *독립된* 수식 셀들 — 위 체인형(cascading) 테스트와 달리 memo는 여기서 전혀 도움이 안 된다
  // (각 칸이 정확히 한 번씩만 계산되므로). memo 없이 풀링도 없던 시절엔 칸마다 `new
  // FormulaParser(...)`를 새로 만드는 비용(실측 ~1.9ms/칸, formulaParserPool.ts 문서 주석 참고)을
  // 그대로 물어 페이지 하나(500행 × 2개 수식 컬럼 = 1000칸)를 재계산하는 데만 수 초가 걸렸다.
  // room.ts의 실제 프로덕션 경로(joinEditSession → 페이지 구독 → recomputeFormulasForPage)를 그대로
  // 거치므로, recomputeFormulasForPage가 실제로 풀링된 파서를 쓰도록 배선돼 있는지(단위 테스트가
  // 아니라 배선 자체)를 검증하는 종단 간(end-to-end) 회귀 테스트다.
  test(
    'end-to-end: subscribing to a page with 500 rows of two independent (non-chained) formula ' +
      'columns computes all values quickly and correctly (matches the original page-load bug report)',
    async () => {
      const { store } = await import('../src/storage.js');
      const { joinEditSession } = await import('../src/collab/room.js');
      const { gridCellKey, GRID_COMPUTED_MAP } = await import('../src/collab/gridShared.js');

      const ROWS = 500;
      const docId = randomUUID();
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Sheet1');
      for (let r = 1; r <= ROWS; r += 1) {
        sheet.getCell(r, 1).value = r; // A{r} = r
        sheet.getCell(r, 2).value = r * 2; // B{r} = 2r
        // C{r}과 D{r}은 오직 A{r}/B{r}(수식이 아닌 입력 칸)만 참조한다 — 서로도, 다른 수식 칸도
        // 참조하지 않는 독립된 수식이다(체인형과 대비되는 이 테스트의 핵심 성질).
        sheet.getCell(r, 3).value = { formula: `A${r}*B${r}` }; // C{r} = A{r}*B{r}
        sheet.getCell(r, 4).value = { formula: `A${r}+B${r}` }; // D{r} = A{r}+B{r}
      }
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

      const start = Date.now();
      sendSubscribe(socket, 0, [0]); // 500행 = 페이지 0 하나 전체
      await flushEventLoop(300);
      const elapsedMs = Date.now() - start;

      const computed = socket.mirror.getMap<string>(GRID_COMPUTED_MAP);
      // 정확성: C{r} = A{r}*B{r} = r*2r = 2r², D{r} = A{r}+B{r} = 3r.
      expect(computed.get(gridCellKey(0, 9, 2))).toBe(String(10 * 20)); // C10 = A10*B10 = 200
      expect(computed.get(gridCellKey(0, 9, 3))).toBe(String(10 + 20)); // D10 = A10+B10 = 30
      expect(computed.get(gridCellKey(0, ROWS - 1, 2))).toBe(String(ROWS * (ROWS * 2))); // C500 = A500*B500
      expect(computed.get(gridCellKey(0, ROWS - 1, 3))).toBe(String(ROWS + ROWS * 2)); // D500 = A500+B500

      // 속도: 리뷰의 실측(이 정확한 시나리오가 풀링된 파서로 ~41ms)과 대비되는 넉넉한 상한 — 풀링
      // 없이 칸마다 새 FormulaParser를 만들던 시절의 수 초짜리 지연으로 되돌아가면 절대 넘길 수
      // 없을 만큼 두세 자릿수 낮게 잡되, CI/느린 머신의 변동성에는 여유를 둔다.
      expect(elapsedMs).toBeLessThan(2_000);
    },
    30_000,
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

  // memo write-log(LayeredMemo) 회귀 테스트: 미스가 난 "더러운" 시도 도중 체인 여러 칸(B1→C1→D1)이
  // 전부 blank(빈 칸) 취급된 A1000을 바탕으로 잘못된 중간값을 계산해 각자의 writeLog에 적어 넣는다
  // — 이 시도는 결국 버려져야 하고(공유 memo에 절대 섞이면 안 됨), 캐시가 채워진 뒤의 "깨끗한" 마지막
  // 시도에서 다시 계산된 올바른 값만 공유 memo에 들어가야 한다. 또한 이 호출과 무관한, 공유 memo에
  // 이미 들어있던 다른 항목은 시도가 몇 번 버려지든 그대로 보존돼야 한다(LayeredMemo.set은 writeLog
  // 에만 쓰고, 승격은 writeLog의 새 항목만 옮기지 base를 건드리지 않으므로).
  test(
    'a chain of formula cells (B1→C1→D1) that only resolves after a deeper cache-miss retry ' +
      'promotes only the correctly-recomputed values into the shared memo, discarding whatever the ' +
      'dirty (blank-substituted) attempt computed for those same keys along the way',
    async () => {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Sheet1');
      sheet.getCell('B1').value = { formula: 'A1000+1' };
      sheet.getCell('C1').value = { formula: 'B1+1' };
      sheet.getCell('D1').value = { formula: 'C1+1' };
      sheet.getCell('A1000').value = 42;
      const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);

      const store = await ingestXlsxLazily(randomUUID(), bytes);
      expect(store.readCache.size).toBe(0); // 캐시가 비어있는 상태로 시작 — row 0도, row 999도 미스가 난다.

      // 이 호출과 무관한, 공유 memo에 이미 들어있던 항목 — 시도가 버려지고 재시도되는 과정에서
      // 사라지거나 바뀌면 안 된다.
      const memo = new Map<string, unknown>([['0:5:5', 'unrelated-preexisting-value']]);

      const result = await computeGridCellDisplayAsync(store, 0, 0, 3, new Set(), memo); // D1
      expect(result).toBe(45); // D1 = C1+1 = (B1+1)+1 = (A1000+1+1)+1 = 42+1+1+1

      // 공유 memo에는 각 칸의 "정확한" 최종 값만 들어가 있어야 한다 — 미스가 난 더러운 시도가
      // A1000을 빈 칸(0)으로 취급해 계산했을 B1=1/C1=2/D1=3 같은 값이 남아있으면 안 된다.
      expect(memo.get('0:0:1')).toBe(43); // B1
      expect(memo.get('0:0:2')).toBe(44); // C1
      expect(memo.get('0:0:3')).toBe(45); // D1

      // 무관한 기존 항목은 그대로 보존된다.
      expect(memo.get('0:5:5')).toBe('unrelated-preexisting-value');

      await fs.rm(store.rowStorePath, { force: true });
    },
  );

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

  // 이 태스크의 핵심 회귀 테스트: 체인형(cascading) 수식 — 러닝 합계처럼 B{r}이 B{r-1}을 참조하는
  // 패턴 — 를 memo 없이 재계산하면 셀 N개에 대해 총 계산 횟수가 O(N²)로 폭증한다(N번째 셀을 계산할
  // 때 1..N-1번째를 전부 재귀로 다시 계산하고, N-1번째 셀이 "그 자신도 재계산 대상 수식 셀이라"
  // 별도로 다시 계산될 때 또 1..N-2번째를 전부 다시 계산하는 식). 500개 셀이면 250,000번 가까운
  // FormulaParser 생성/파싱이 필요해, 실측(브리프 참고)으로는 90초 안에 끝나지 않았다.
  // recomputeFormulasForPage가 페이지 하나(최대 PAGE_SIZE_ROWS=500행)의 모든 수식 셀에 걸쳐 공유하는
  // memo(Map<cellKey, result>)를 만들어 computeGridCellDisplayAsync에 넘기면, 이미 이 패스에서
  // 계산된 셀은 다시 계산하지 않으므로 총 계산 횟수가 O(N)으로 줄어든다 — 이 테스트는 그 memo를
  // recomputeFormulasForPage와 똑같은 방식(모든 호출이 같은 Map 인스턴스를 공유)으로 직접 사용해
  // "빠르다"(관대하지만 명확한 시간 상한)와 "정확하다"(몇몇 행의 러닝 합계 값)를 함께 검증한다.
  test(
    '500 rows of a chained/cascading formula (running total, B{r} = B{r-1} + A{r}) recompute quickly ' +
      'and correctly when computeGridCellDisplayAsync calls share one memo Map, matching recomputeFormulasForPage',
    async () => {
      const ROWS = 500;
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Sheet1');
      for (let r = 1; r <= ROWS; r += 1) {
        sheet.getCell(r, 1).value = r; // A{r} = r
        sheet.getCell(r, 2).value = r === 1 ? { formula: 'A1' } : { formula: `B${r - 1}+A${r}` }; // B{r}
      }
      const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
      const store = await ingestXlsxLazily(randomUUID(), bytes);

      // recomputeFormulasForPage가 한 페이지의 모든 수식 셀에 걸쳐 공유하는 것과 똑같은 memo — 오름차순
      // 행 순서로 호출해 "앞 셀을 먼저 계산해 memo에 남기고, 뒤 셀은 그걸 재사용"하는 실제 패턴을 그대로 재현한다.
      const memo = new Map<string, unknown>();
      const start = Date.now();
      const results: unknown[] = [];
      for (let r = 0; r < ROWS; r += 1) {
        results.push(await computeGridCellDisplayAsync(store, 0, r, 1, new Set(), memo));
      }
      const elapsedMs = Date.now() - start;

      // 정확성: 러닝 합계 B{r} = 1+2+...+r = r(r+1)/2.
      expect(results[0]).toBe(1); // B1
      expect(results[1]).toBe(3); // B2
      expect(results[9]).toBe(55); // B10
      expect(results[99]).toBe(5050); // B100
      expect(results[ROWS - 1]).toBe((ROWS * (ROWS + 1)) / 2); // B500 = 125250

      // 속도: memo가 없으면(O(N²)) 90초 넘게도 끝나지 않던 것이(브리프 실측) memo 공유로
      // O(N)이 되어 훨씬 빨리 끝나야 한다 — 전체 스위트를 병렬로 돌릴 때의 CPU 경합까지 고려해
      // 여유를 넉넉히 둔다(격리 실행 시엔 1.5초 안팎). 타이트한 타이밍 어서션으로 인한 플레이크를
      // 피하면서도, O(N²)로 되돌아가면 절대 넘길 수 없는 상한이라는 게 핵심.
      expect(elapsedMs).toBeLessThan(15_000);

      await fs.rm(store.rowStorePath, { force: true });
    },
    30_000,
  );
});
