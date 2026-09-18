import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import ExcelJS from 'exceljs';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeSyncStep1, writeUpdate } from 'y-protocols/sync';
import { describe, expect, test } from 'vitest';

/**
 * "누가/언제/어디를/무엇으로" 바꿨는지 조회하는 getCellChangeLog(room.ts) 기능의 테스트 —
 * tests/roomCanonicalGrid.test.ts와 같은 FakeSocket/sendCellEdit 패턴을 그대로 쓴다.
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

/** roomCanonicalGrid.test.ts의 sendCellEdit과 같은 clientID 고정 트릭을 쓰지만, 한 소켓이
 * **여러 번** 편집을 보내는 테스트가 있어(같은 칸을 두 번 고치는 경우 등) scratch Y.Doc을 호출마다
 * 새로 만들지 않고 소켓별로 하나씩만 만들어 재사용한다 — 매번 새 Y.Doc을 쓰면 clock이 매번 0으로
 * 리셋되는데 clientID는 그대로 고정이라, 두 번째 호출이 Yjs 입장에서는 "이미 적용된 것과 똑같은
 * (clientID, clock) 항목"이 되어 조용히 무시된다(관찰자가 아예 안 불림 — 실제로 이 버그로 처음엔
 * 테스트가 실패했다). 소켓마다 하나의 scratch를 계속 쓰면 실제 브라우저 클라이언트의 Y.Doc처럼
 * clock이 매 편집마다 자연스럽게 올라간다. */
const scratchDocs = new WeakMap<FakeSocket, Y.Doc>();
let nextScratchClientId = 0xffffffff;
function scratchDocFor(socket: FakeSocket): Y.Doc {
  let scratch = scratchDocs.get(socket);
  if (!scratch) {
    scratch = new Y.Doc();
    // 소켓마다 다른(하지만 여전히 room의 무작위 32비트 clientID보다 확실히 큰) 값을 준다 — 두
    // 소켓이 동시에 서로 다른 칸을 고치는 테스트에서, 둘 다 같은 clientID를 쓰면 clock이 각자
    // 0부터 다시 시작해 서로 다른 항목인데도 Yjs가 같은 (clientID, clock) 좌표로 오인할 수 있다.
    scratch.clientID = nextScratchClientId;
    nextScratchClientId -= 1;
    scratchDocs.set(socket, scratch);
  }
  return scratch;
}

function sendCellEdit(socket: FakeSocket, key: string, value: string): void {
  const scratch = scratchDocFor(socket);
  scratch.getMap<string>('cells').set(key, value);
  const update = Y.encodeStateAsUpdate(scratch);
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeUpdate(encoder, update);
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** roomCanonicalGrid.test.ts의 동명 헬퍼와 동일 — Task 7(윈도잉 컷오버) 이후 Yjs cells 맵은
 * "구독 중인 페이지"만 담으므로, 원본에 이미 있던 값(예: "Alice")을 지우는 걸 재현하려면 먼저 그
 * 페이지를 구독해 room이 그 값을 cells 맵에 시딩하게 해야 한다 — 안 그러면 mirror도 room도
 * 그 키를 아예 모르는 상태라 delete가 아무것도 지우지 못한다. */
function sendSubscribe(socket: FakeSocket, sheet: number, pages: readonly number[]): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_PAGE_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify({ sheet, pages }));
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** roomCanonicalGrid.test.ts의 동명 헬퍼와 동일 — 접속 직후 클라이언트가 자기 상태 벡터(비어
 * 있음)를 보내는 것을 흉내낸다. socket.mirror가 room.ydoc과 같은 아이템 식별자를 공유해야만
 * 그 뒤 mirror에서 만든 delete가 room의 실제 항목을 지우는 delete로 정확히 병합된다. */
function syncMirrorWithRoom(socket: FakeSocket): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeSyncStep1(encoder, socket.mirror);
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** 서로 다른 소켓이 **같은 칸**을 순서대로 고치는 시나리오에서 sendCellEdit(독립 scratch,
 * clientID로 승부)를 쓰면 안 된다 — 두 스크래치가 서로의 존재를 몰라 Yjs가 "동시 편집"으로
 * 보고 clientID가 더 큰 쪽을 무조건 이기게 해서, 실제로 나중에 온 편집이 조용히 버려질 수
 * 있다(Y.Map의 동시-편집 충돌 해소 규칙 — 적용 순서가 아니라 clientID 크기로 승자를 정한다).
 * `syncMirrorWithRoom`을 먼저 불러 mirror가 room의 "지금까지" 상태(이전 편집 포함)를 안 뒤에
 * 편집해야, 이 편집이 그 이전 편집의 인과적 후속으로 인식되어 항상(clientID와 무관하게) 이긴다.
 * deleteCellViaClient와 같은 이유로 전체 상태가 아니라 델타만 캡처해 보낸다. */
function editCellViaClient(socket: FakeSocket, key: string, value: string): void {
  let captured: Uint8Array | null = null;
  const onUpdate = (update: Uint8Array) => {
    captured = update;
  };
  socket.mirror.on('update', onUpdate);
  socket.mirror.getMap<string>('cells').set(key, value);
  socket.mirror.off('update', onUpdate);
  if (!captured) {
    throw new Error('edit produced no update');
  }
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeUpdate(encoder, captured);
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** roomCanonicalGrid.test.ts의 동명 헬퍼와 동일 — 실제 클라이언트(grid-editor.ts의 commitCell)가
 * 칸을 지울 때 하는 것과 같은 동작(cells.delete(key), cells.set(key,'')가 아니다)을 mirror에서
 * 수행하고, 그 결과로 생긴 update만(전체 상태가 아니라 델타만) 캡처해 서버로 보낸다. */
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

async function drainPendingDiskIo(): Promise<void> {
  await fs.stat(os.tmpdir()).catch(() => undefined);
}

async function flushEventLoop(turns = 50): Promise<void> {
  const { drainRoomBackgroundWorkForTest } = await import('../src/collab/room.js');
  for (let i = 0; i < turns; i += 1) {
    await new Promise((resolve) => setImmediate(resolve));
    await drainPendingDiskIo();
    await drainRoomBackgroundWorkForTest();
  }
}

describe('collab/room.ts: getCellChangeLog', () => {
  test('returns null when no room is open for the document', async () => {
    const { getCellChangeLog } = await import('../src/collab/room.js');
    expect(getCellChangeLog(randomUUID())).toBeNull();
  });

  test('returns an empty array for an open room with no edits yet', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getCellChangeLog } = await import('../src/collab/room.js');

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
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false, 'viewer-a');
    expect(getCellChangeLog(docId)).toEqual([]);
  });

  test('records before/after/viewer/timestamp for a single edit, with 1-based row/col', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getCellChangeLog } = await import('../src/collab/room.js');
    const { gridCellKey } = await import('../src/collab/gridShared.js');

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
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false, 'viewer-a');

    const before = Date.now();
    // (row=1, col=0) internal = A2 = "Alice"
    sendCellEdit(socket, gridCellKey(0, 1, 0), 'Bob');
    await flushEventLoop();

    const log = getCellChangeLog(docId);
    expect(log).toHaveLength(1);
    expect(log![0]).toMatchObject({ sheet: 0, row: 2, col: 1, before: 'Alice', after: 'Bob', viewer: 'viewer-a' });
    expect(new Date(log![0].at).getTime()).toBeGreaterThanOrEqual(before);
  });

  test('editing the same cell twice keeps the ORIGINAL before-value but updates after/viewer/at to the latest', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getCellChangeLog } = await import('../src/collab/room.js');
    const { gridCellKey } = await import('../src/collab/gridShared.js');

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

    const socketA = new FakeSocket();
    await joinEditSession(socketA as never, {} as never, docId, 'xlsx', [], false, 'viewer-a');
    syncMirrorWithRoom(socketA);
    editCellViaClient(socketA, gridCellKey(0, 1, 0), 'Bob');
    await flushEventLoop();

    const socketB = new FakeSocket();
    await joinEditSession(socketB as never, {} as never, docId, 'xlsx', [], false, 'viewer-b');
    // A의 편집(Bob) 이후에 동기화해야 B의 편집이 그 편집의 인과적 후속이 된다(editCellViaClient
    // 문서 주석 참고) — 안 그러면 "동시 편집"으로 보여 clientID 크기로 승자가 정해진다.
    syncMirrorWithRoom(socketB);
    editCellViaClient(socketB, gridCellKey(0, 1, 0), 'Carol');
    await flushEventLoop();

    const log = getCellChangeLog(docId);
    expect(log).toHaveLength(1);
    // before는 최초 원본("Alice")을 계속 유지 — "Bob"(중간값)으로 갱신되지 않는다.
    expect(log![0]).toMatchObject({ before: 'Alice', after: 'Carol', viewer: 'viewer-b' });
  });

  test('two concurrent viewers editing different cells are each attributed correctly (no cross-contamination)', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getCellChangeLog } = await import('../src/collab/room.js');
    const { gridCellKey } = await import('../src/collab/gridShared.js');

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

    const socketA = new FakeSocket();
    const socketB = new FakeSocket();
    await joinEditSession(socketA as never, {} as never, docId, 'xlsx', [], false, 'viewer-a');
    await joinEditSession(socketB as never, {} as never, docId, 'xlsx', [], false, 'viewer-b');

    sendCellEdit(socketA, gridCellKey(0, 1, 0), 'A_EDIT'); // A2
    sendCellEdit(socketB, gridCellKey(0, 1, 1), '999'); // B2
    await flushEventLoop();

    const log = getCellChangeLog(docId);
    expect(log).toHaveLength(2);
    const byViewer = Object.fromEntries(log!.map((e) => [e.viewer, e]));
    expect(byViewer['viewer-a']).toMatchObject({ row: 2, col: 1, before: 'Alice', after: 'A_EDIT' });
    expect(byViewer['viewer-b']).toMatchObject({ row: 2, col: 2, before: '10', after: '999' });
  });

  test('a delete records the pre-delete value as before and an empty string as after', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getCellChangeLog } = await import('../src/collab/room.js');
    const { gridCellKey } = await import('../src/collab/gridShared.js');

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
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false, 'viewer-a');
    sendSubscribe(socket, 0, [0]); // room이 이 페이지의 원본 값을 cells 맵에 시딩하게 한다
    await flushEventLoop();

    syncMirrorWithRoom(socket);
    await flushEventLoop(5);
    deleteCellViaClient(socket, gridCellKey(0, 1, 0));
    await flushEventLoop();

    const log = getCellChangeLog(docId);
    expect(log).toHaveLength(1);
    expect(log![0]).toMatchObject({ row: 2, col: 1, before: 'Alice', after: '', viewer: 'viewer-a' });
  });

  test('returns null for a docx room (change log is xlsx-only)', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getCellChangeLog } = await import('../src/collab/room.js');
    const { makeSampleDocx } = await import('./helpers.js');

    const docId = randomUUID();
    const bytes = makeSampleDocx(['hello']);
    await store.saveOriginal(docId, 'docx', bytes);
    await store.writeMeta({
      id: docId,
      name: 'x.docx',
      ext: 'docx',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
      progress: 100,
      viewKind: 'pages',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'docx', [], false, 'viewer-a');
    expect(getCellChangeLog(docId)).toBeNull();
  });
});
