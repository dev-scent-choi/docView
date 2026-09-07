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
  sheet.getCell('A1').value = 'Name';
  sheet.getCell('A2').value = 'Alice';
  sheet.getCell('B2').value = 10;
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
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
    const { joinEditSession, getRoomCanonicalGridForTest } = await import('../src/collab/room.js');

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

    const canonical = getRoomCanonicalGridForTest(docId);
    expect(canonical).not.toBeNull();
    expect(canonical!.sheets[0].rows[0][0]).toEqual({ v: 'Name' });
    expect(canonical!.sheets[0].rows[1][0]).toEqual({ v: 'Alice' });
  });

  test('an edit is reflected in the canonical grid immediately, without waiting for autosave', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomCanonicalGridForTest } = await import('../src/collab/room.js');
    const { gridCellKey } = await import('../src/collab/gridShared.js');

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

    const canonical = getRoomCanonicalGridForTest(docId);
    expect(canonical!.sheets[0].rows[1][0]).toEqual({ v: 'Alice Kim' });
  });
});
