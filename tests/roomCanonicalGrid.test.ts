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

/** A2("Alice")에 굵게(bold) 스타일이 걸린 샘플 — canonical에 편집 후에도 스타일(.s)이
 * 유지되는지 검증하기 위한 픽스처. */
async function makeStyledSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 'Name';
  const aliceCell = sheet.getCell('A2');
  aliceCell.value = 'Alice';
  aliceCell.font = { bold: true };
  sheet.getCell('B2').value = 10;
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
}

/** A2에 천단위 구분 쉼표 서식이 걸린 숫자 셀 — parseXlsxToGrid는 이걸 "1,000"으로 표시하지만
 * Yjs cells 맵엔 편집 가능하도록 "1000"으로 시딩된다(seedableCellValue). canonical이 시딩
 * 직후 이 쉼표-스트립 값에 오염되지 않는지 검증하기 위한 픽스처. */
async function makeNumberFormattedXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 'Total';
  const amountCell = sheet.getCell('A2');
  amountCell.value = 1000;
  amountCell.numFmt = '#,##0';
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
}

/**
 * Task 7(윈도잉 컷오버) 이후 Yjs cells 맵은 "구독 중인 페이지"만 담는다 — 룸을 열기만 하고
 * 아무 페이지도 구독하지 않으면 셀이 하나도 채워지지 않으므로, 실제 클라이언트가 셀을 보고
 * 편집·삭제할 수 있는 상태를 재현하려면 먼저 그 페이지를 구독해야 한다. 이 파일의 픽스처는
 * 전부 몇 행뿐이라 페이지 0 하나가 시트 전체를 덮는다.
 */
function sendSubscribe(socket: FakeSocket, sheet: number, pages: readonly number[]): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_PAGE_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify({ sheet, pages }));
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** 페이지 채우기는 청크마다 실제 setImmediate로 양보하므로 그 턴을 흘려보내야 한다. */
async function flushEventLoop(turns = 50): Promise<void> {
  for (let i = 0; i < turns; i += 1) await new Promise((resolve) => setImmediate(resolve));
}

function sendCellEdit(socket: FakeSocket, key: string, value: string): void {
  // scratch는 room의 Y.Doc과 아무 인과관계도 공유하지 않는 완전히 새 Y.Doc이라, 이미 값이 있는
  // 같은 키에 값을 쓰면 Yjs Y.Map의 동시편집(concurrent) 충돌 해소 규칙(더 큰 clientID를 가진
  // 쪽이 이긴다 — 적용 순서와 무관)에 따라 어느 쪽이 "최종값"이 될지가 두 Y.Doc의 무작위
  // clientID 대소 비교로 결정돼 버린다(실측: 약 40~50% 확률로 새 값이 조용히 무시됨). 실제
  // y-websocket 클라이언트는 접속 시 항상 먼저 서버 상태를 동기화한 뒤에 편집하므로 이런 충돌이
  // 없다 — 이 테스트 헬퍼도 같은 보장을 흉내 내기 위해, scratch의 clientID를 표현 가능한 최댓값으로
  // 고정해 "이 편집이 항상 이긴다"를 결정적으로 만든다(가능한 32비트 clientID 공간에서 room의
  // ydoc이 우연히 똑같은 최댓값을 뽑을 확률은 무시할 수 있는 수준).
  const scratch = new Y.Doc();
  scratch.clientID = 0xffffffff;
  scratch.getMap<string>('cells').set(key, value);
  const update = Y.encodeStateAsUpdate(scratch);
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeUpdate(encoder, update);
  socket.emit('message', encoding.toUint8Array(encoder));
}

/**
 * 클라이언트가 접속 직후 자기 상태 벡터(비어 있음)를 서버로 보내는 것을 흉내낸다 — 서버의
 * handleMessage가 messageYjsSyncStep1 분기에서 room.ydoc과의 diff(=현재 룸의 전체 내용)를
 * syncStep2로 되돌려 보내주고, 그 응답이 FakeSocket.send를 거쳐 socket.mirror에 적용된다.
 * 이렇게 해야 socket.mirror가 room.ydoc과 같은 아이템 식별자(client/clock)를 공유하게 되어,
 * 그 뒤 mirror에서 만든 delete가 room의 실제 항목을 지우는 delete로 정확히 병합된다(실제
 * y-websocket 클라이언트가 접속 시 거치는 것과 같은 절차).
 */
function syncMirrorWithRoom(socket: FakeSocket): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeSyncStep1(encoder, socket.mirror);
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** 실제 클라이언트(grid-editor.ts의 commitCell)가 셀을 지울 때 하는 것과 같은 동작 —
 * `cells.set(key, '')`가 아니라 `cells.delete(key)` — 를 mirror Y.Doc에서 수행하고, 그 결과
 * 생긴 update를 서버로 전송한다. syncMirrorWithRoom을 먼저 호출해 mirror가 room과 같은 항목을
 * 알고 있어야만 delete가 실제로 뭔가를 지우는 update를 만들어낸다(안 그러면 mirror 로컬에
 * 해당 키의 항목이 없어 delete가 아무 update도 생성하지 않는다). */
function deleteCellViaClient(socket: FakeSocket, key: string): void {
  let captured: Uint8Array | null = null;
  const onUpdate = (update: Uint8Array) => {
    captured = update;
  };
  socket.mirror.on('update', onUpdate);
  socket.mirror.getMap<string>('cells').delete(key);
  socket.mirror.off('update', onUpdate);
  if (!captured) {
    throw new Error('delete produced no update — was syncMirrorWithRoom called first?');
  }
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeUpdate(encoder, captured);
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
    const { joinEditSession, getRoomLazyGridForTest } = await import('../src/collab/room.js');
    const { getCell } = await import('../src/xlsx/lazyGridStore.js');

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

    const lazyGrid = getRoomLazyGridForTest(docId);
    expect(lazyGrid).not.toBeNull();
    expect(await getCell(lazyGrid!, 0, 0, 0)).toEqual({ v: 'Name' });
    expect(await getCell(lazyGrid!, 0, 1, 0)).toEqual({ v: 'Alice' });
  });

  test('an edit is reflected in the canonical grid immediately, without waiting for autosave', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomLazyGridForTest } = await import('../src/collab/room.js');
    const { gridCellKey } = await import('../src/collab/gridShared.js');
    const { getCell } = await import('../src/xlsx/lazyGridStore.js');

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
    // applyCellsChangeToCanonical은 async(서식 보존을 위해 getCell로 현재 값을 먼저 읽는다) —
    // cells.observe 콜백은 이를 void로 흘려보내므로, dirtyOverlay에 실제로 쓰이기까지 최소
    // 한 틱 이상 걸린다(과거의 완전 동기 GridDocument 대입과 다른 점).
    await flushEventLoop();

    const lazyGrid = getRoomLazyGridForTest(docId);
    expect(await getCell(lazyGrid!, 0, 1, 0)).toEqual({ v: 'Alice Kim' });
  });

  test('deleting a cell (Y.Map key delete, not set to "") clears it to 0 in the canonical grid', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomLazyGridForTest } = await import('../src/collab/room.js');
    const { gridCellKey } = await import('../src/collab/gridShared.js');
    const { getCell } = await import('../src/xlsx/lazyGridStore.js');

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

    // 편집 전 확인: A2는 "Alice"를 담고 있어야 한다.
    const before = getRoomLazyGridForTest(docId);
    expect(await getCell(before!, 0, 1, 0)).toEqual({ v: 'Alice' });

    // 실제 클라이언트처럼 먼저 보고 있는 페이지를 구독한다 — 그래야 그 페이지의 셀이 Yjs에
    // 채워지고, 아래 syncMirrorWithRoom이 mirror에 그 셀을 가져올 수 있다(윈도잉 이후 구독
    // 없이는 지울 셀 자체가 Yjs에 없다).
    sendSubscribe(socket, 0, [0]);
    await flushEventLoop();

    // 실제 클라이언트(grid-editor.ts의 commitCell)가 셀을 지울 때 하는 것과 동일하게
    // cells.delete(key)를 수행한다 — cells.set(key, '')가 아니다.
    syncMirrorWithRoom(socket);
    deleteCellViaClient(socket, gridCellKey(0, 1, 0));
    await flushEventLoop();

    const lazyGrid = getRoomLazyGridForTest(docId);
    expect(await getCell(lazyGrid!, 0, 1, 0)).toBe(0);
  });

  test('editing a cell that already has a style preserves its style (.s) in the canonical grid', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomLazyGridForTest } = await import('../src/collab/room.js');
    const { gridCellKey } = await import('../src/collab/gridShared.js');
    const { getCell } = await import('../src/xlsx/lazyGridStore.js');

    const docId = randomUUID();
    const original = await makeStyledSampleXlsx();
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

    const before = getRoomLazyGridForTest(docId);
    const beforeCell = await getCell(before!, 0, 1, 0);
    expect(beforeCell).not.toBe(0);
    const styleIndex = (beforeCell as { v: string; s?: number }).s;
    expect(styleIndex).toBeDefined();

    // 텍스트만 바꾸는 편집 — 스타일은 건드리지 않는다.
    sendCellEdit(socket, gridCellKey(0, 1, 0), 'Alice Kim');
    await flushEventLoop();

    const lazyGrid = getRoomLazyGridForTest(docId);
    expect(await getCell(lazyGrid!, 0, 1, 0)).toEqual({ v: 'Alice Kim', s: styleIndex });
  });

  test('seeding does not corrupt the canonical grid with the stripped (editable) cell text', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomLazyGridForTest } = await import('../src/collab/room.js');
    const { getCell } = await import('../src/xlsx/lazyGridStore.js');

    const docId = randomUUID();
    const original = await makeNumberFormattedXlsx();
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

    // 윈도잉 이후 시딩은 룸 생성 시점이 아니라 페이지가 구독될 때 일어난다 — 즉 cells.observe가
    // 이미 걸린 뒤에 계속 일어난다. 그래서 구독을 실제로 보내 그 시딩을 유발해야 이 테스트가
    // 검증하려는 오염 경로가 재현된다(구독을 보내지 않으면 아무것도 시딩되지 않아 통과가 무의미).
    sendSubscribe(socket, 0, [0]);
    await flushEventLoop();

    // 편집 메시지를 하나도 보내지 않았으므로 — canonical은 여전히 parseXlsxToGrid가 만든 올바른
    // 표시 텍스트("1,000")를 담고 있어야 한다. seedGridPageFromXlsx가 Yjs cells 맵에 심는 편집용
    // 스트립 값("1000")으로 오염되면 안 된다(그 시딩 트랜잭션은 GRID_WINDOW_SEED_ORIGIN 표식을
    // 달고 있어 cells.observe가 canonical 되반영을 건너뛴다).
    const lazyGrid = getRoomLazyGridForTest(docId);
    const cell = await getCell(lazyGrid!, 0, 1, 0);
    expect(cell).not.toBe(0);
    expect((cell as { v: string }).v).toBe('1,000');
  });
});
