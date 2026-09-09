import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

/**
 * `DELETE /api/documents/:id`가 메모리상 살아있는 협업 Room을 정리하지 않던 문제(room.ts의
 * forceCloseRoom) 검증.
 *
 * 이 문제가 없으면: 문서를 삭제해도 이미 열려 있던 Room(canonicalGrid 등 대용량 메모리)은
 * 마지막 참여자가 스스로 나갈 때까지 계속 살아 있었고, 대용량 룸 상한(checkRoomCapacity) 슬롯도
 * 계속 차지했다.
 */

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  closedWith: { code?: number; reason?: string } | null = null;
  send(_data: Uint8Array, cb?: (err?: Error) => void): void {
    cb?.();
  }
  // 실제 ws의 WebSocket#close()는 이미 CLOSING/CLOSED 상태에서 다시 부르면 아무 일도 하지
  // 않는다(코드/이유를 덮어쓰지 않는다) — forceCloseRoom이 명시적 코드로 닫은 뒤 closeConn이
  // 인자 없이 또 부르는 경우가 실제로 있으므로, 이 더블도 첫 close만 기록해야 실제 동작과 맞다.
  close(code?: number, reason?: string): void {
    if (this.closedWith) return;
    this.closedWith = { code, reason };
  }
  ping(): void {}
}

async function makeSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Sheet1').getCell('A1').value = 'x';
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

describe('room.ts: forceCloseRoom cleans up an open room after document deletion', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-forceclose-'));
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

  test('closes every live connection with a 4410 code and the given reason', async () => {
    const { joinEditSession, forceCloseRoom } = await import('../src/collab/room.js');

    const docId = randomUUID();
    await seedDoc(docId, await makeSampleXlsx());
    const socketA = new FakeSocket();
    const socketB = new FakeSocket();
    await joinEditSession(socketA as never, {} as never, docId, 'xlsx', [], false);
    await joinEditSession(socketB as never, {} as never, docId, 'xlsx', [], false);

    forceCloseRoom(docId, '문서가 삭제되었습니다.');

    expect(socketA.closedWith).toEqual({ code: 4410, reason: '문서가 삭제되었습니다.' });
    expect(socketB.closedWith).toEqual({ code: 4410, reason: '문서가 삭제되었습니다.' });
  });

  test('removes the room from memory so a later checkRoomCapacity/canonical lookup sees it as gone', async () => {
    const { joinEditSession, forceCloseRoom, getRoomCanonicalGridForTest } = await import('../src/collab/room.js');

    const docId = randomUUID();
    await seedDoc(docId, await makeSampleXlsx());
    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
    await new Promise((resolve) => setImmediate(resolve)); // room.ready 시딩 완료를 기다린다

    expect(getRoomCanonicalGridForTest(docId)).not.toBeNull();

    forceCloseRoom(docId);

    expect(getRoomCanonicalGridForTest(docId)).toBeNull();
  });

  test('frees a large-room capacity slot immediately, without waiting for a natural disconnect', async () => {
    const { joinEditSession, forceCloseRoom, checkRoomCapacity } = await import('../src/collab/room.js');
    const { store } = await import('../src/storage.js');

    const large = Buffer.alloc(6 * 1024 * 1024);
    const docIds = await Promise.all(
      Array.from({ length: 3 }, async () => {
        const docId = randomUUID();
        await store.saveOriginal(docId, 'xlsx', large);
        await store.writeMeta({
          id: docId,
          name: 'big.xlsx',
          ext: 'xlsx',
          status: 'ready',
          pageCount: 1,
          createdAt: new Date().toISOString(),
          progress: 100,
          viewKind: 'grid',
        });
        return docId;
      }),
    );

    const sockets: FakeSocket[] = [];
    for (const docId of docIds) {
      expect((await checkRoomCapacity(docId, 'xlsx')).ok).toBe(true);
      const socket = new FakeSocket();
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
      sockets.push(socket);
    }

    const fourthDocId = randomUUID();
    await store.saveOriginal(fourthDocId, 'xlsx', large);
    await store.writeMeta({
      id: fourthDocId,
      name: 'big4.xlsx',
      ext: 'xlsx',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
      progress: 100,
      viewKind: 'grid',
    });
    expect((await checkRoomCapacity(fourthDocId, 'xlsx')).ok).toBe(false);

    // 문서 삭제 시나리오: 스토리지에서 지운 뒤 룸을 강제로 닫는다(documents.ts DELETE 핸들러와 같은 순서).
    await store.deleteDocument(docIds[0]);
    forceCloseRoom(docIds[0]);

    expect((await checkRoomCapacity(fourthDocId, 'xlsx')).ok).toBe(true);
  });

  test('is a safe no-op when no room is open for the given docId', async () => {
    const { forceCloseRoom } = await import('../src/collab/room.js');
    expect(() => forceCloseRoom(randomUUID())).not.toThrow();
  });

  test('does not resurrect the deleted file via autosave when the room is force-closed after deletion', async () => {
    const { joinEditSession, forceCloseRoom } = await import('../src/collab/room.js');
    const { store } = await import('../src/storage.js');

    const docId = randomUUID();
    await seedDoc(docId, await makeSampleXlsx());
    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
    await new Promise((resolve) => setImmediate(resolve));

    // documents.ts DELETE 핸들러와 같은 순서: 스토리지 삭제가 먼저, 룸 강제 종료가 나중.
    await store.deleteDocument(docId);
    forceCloseRoom(docId);
    // 종료 경로가 예약하는 queueFlush(flushToStore)까지 흘려보낸다 — meta가 없으니 아무 것도
    // 쓰지 않아야 한다(flushToStore의 이른 반환).
    await new Promise((resolve) => setImmediate(resolve));
    await new Promise((resolve) => setImmediate(resolve));

    await expect(store.readMeta(docId)).resolves.toBeNull();
  });
});
