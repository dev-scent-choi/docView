import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeUpdate } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import ExcelJS from 'exceljs';

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
      // messageYjsSyncStep1은 이 테스트에서 응답할 필요가 없다(셀 편집만 서버로 보내면 된다).
    }
    cb?.();
  }
  close(): void {}
  ping(): void {}
}

async function flushMicrotasks(times = 2_000): Promise<void> {
  for (let i = 0; i < times; i += 1) {
    await Promise.resolve();
    // setTimeout/setInterval만 가짜로 만들었으므로(위 useFakeTimers 참고) setImmediate는 실제
    // 매크로태스크다 — ExcelJS의 zip 스트림 처리(exportXlsxFromGrid/reparseGridForViewer 내부)가
    // 그걸 타므로, 마이크로태스크만 비워서는 그 체인이 끝까지 진행되지 않는다. 실측 결과 이
    // load→write 체인이 완주하는 데 수백~천 회 단위의 실제 setImmediate 턴이 필요했다(50회는
    // 부족해 항상 타임아웃 없이 조용히 hasDerived만 false로 남았다) — 반복 자체는 가벼워서
    // (총 몇십 ms) 횟수를 넉넉히 늘리는 것이 안전하다.
    await new Promise((resolve) => setImmediate(resolve));
  }
}

async function makeSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 'Name';
  sheet.getCell('A2').value = 'Alice';
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
}

describe('room.ts: xlsx collab save keeps gridindex in sync with the grid it just wrote', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-gridindex-'));
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

  test('after an xlsx cell edit autosaves, gridindex offsets match the freshly saved grid bytes', async () => {
    // 일반 vi.useFakeTimers()는 기본으로 setImmediate/queueMicrotask까지 가짜로 만든다.
    // ExcelJS의 xlsx 파싱(zip 스트림 처리)이 내부적으로 그 두 API에 의존해, 가짜 시계를 한 번도
    // 감기 전에는(=joinEditSession이 room 시딩을 위해 parseXlsxToGrid를 기다리는 이 시점) 영원히
    // 끝나지 않는다 — 실제로 이 테스트를 처음 작성했을 때 5초 타임아웃으로 그대로 걸렸다.
    // 자동저장 디바운스만 제어하면 되므로 setTimeout/setInterval 계열만 가짜로 만든다.
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });

    const { store } = await import('../src/storage.js');
    const { joinEditSession } = await import('../src/collab/room.js');
    const { GRID_CELLS_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
    const { computeLineOffsets } = await import('../src/spreadsheet.js');

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

    // 실제 원격 편집처럼, 별도 Y.Doc에서 셀 하나를 바꾼 업데이트를 서버로 보낸다.
    const scratch = new Y.Doc();
    scratch.getMap<string>(GRID_CELLS_MAP).set(gridCellKey(0, 1, 0), 'Alice Kim');
    const update = Y.encodeStateAsUpdate(scratch);
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeUpdate(encoder, update);
    socket.emit('message', encoding.toUint8Array(encoder));

    await vi.advanceTimersByTimeAsync(8_100);
    await flushMicrotasks();

    expect(await store.hasDerived(docId, 'gridindex')).toBe(true);
    const savedGrid = await store.readDerived(docId, 'grid');
    const { lineOffsets } = JSON.parse((await store.readDerived(docId, 'gridindex')).toString('utf8'));
    expect(lineOffsets).toEqual(computeLineOffsets(savedGrid));
  });
});
