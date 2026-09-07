import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as encoding from 'lib0/encoding';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

const MESSAGE_SYNC = 0;
const MESSAGE_PAGE_SUBSCRIBE = 2;

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  send(_data: Uint8Array, cb?: (err?: Error) => void): void { cb?.(); }
  close(): void {}
  ping(): void {}
}

function sendSubscribe(socket: FakeSocket, sheet: number, pages: readonly number[]): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_PAGE_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify({ sheet, pages }));
  socket.emit('message', encoding.toUint8Array(encoder));
}

async function makeSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Sheet1').getCell('A1').value = 'x';
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

describe('room.ts: page subscription tracking', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-subscribe-'));
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

  test('a single client subscribing to pages [0, 2] is reflected in the union', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getSubscribedPageUnionForTest } = await import('../src/collab/room.js');

    const docId = randomUUID();
    const original = await makeSampleXlsx();
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
    sendSubscribe(socket, 0, [0, 2]);

    const union = getSubscribedPageUnionForTest(docId, 0);
    expect([...union].sort()).toEqual([0, 2]);
  });

  test('two clients subscribing to overlapping pages produce the correct union', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getSubscribedPageUnionForTest } = await import('../src/collab/room.js');

    const docId = randomUUID();
    const original = await makeSampleXlsx();
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid',
    });

    const socketA = new FakeSocket();
    const socketB = new FakeSocket();
    await joinEditSession(socketA as never, {} as never, docId, 'xlsx', [], false);
    await joinEditSession(socketB as never, {} as never, docId, 'xlsx', [], false);
    sendSubscribe(socketA, 0, [0, 1]);
    sendSubscribe(socketB, 0, [1, 2]);

    const union = getSubscribedPageUnionForTest(docId, 0);
    expect([...union].sort()).toEqual([0, 1, 2]);
  });

  test('a client re-subscribing to a new sheet/page list replaces its previous subscription entirely', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getSubscribedPageUnionForTest } = await import('../src/collab/room.js');

    const docId = randomUUID();
    const original = await makeSampleXlsx();
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
    sendSubscribe(socket, 0, [0, 1]);
    sendSubscribe(socket, 0, [5]); // 스크롤해서 완전히 다른 페이지로 이동

    const union = getSubscribedPageUnionForTest(docId, 0);
    expect([...union]).toEqual([5]);
  });

  test('disconnecting a client removes its subscription from the union', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getSubscribedPageUnionForTest } = await import('../src/collab/room.js');

    const docId = randomUUID();
    const original = await makeSampleXlsx();
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
    sendSubscribe(socket, 0, [0]);
    expect([...getSubscribedPageUnionForTest(docId, 0)]).toEqual([0]);

    socket.emit('close');
    expect([...getSubscribedPageUnionForTest(docId, 0)]).toEqual([]);
  });
});
