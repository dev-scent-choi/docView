import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

/**
 * room.ts의 exportXlsxWithFallback이 실제로 정상 경로 실패 시 스트리밍 폴백으로 넘어가는지
 * 검증한다 — exportStreaming.ts 자체의 정확성은 tests/xlsxExportStreaming.test.ts가 다루고,
 * 여기서는 오직 "room.ts가 실패를 감지해 실제로 폴백을 호출하고, 그 결과가 저장에 정상적으로
 * 쓰이는지, 그리고 폴백마저 실패하면 그 실패가 조용히 삼켜지지 않고 그대로 드러나는지"만
 * 확인한다.
 */

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  send(_data: Uint8Array, cb?: (err?: Error) => void): void {
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

describe('room.ts: exportXlsxWithFallback falls back to the streaming exporter when the normal path throws', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-fallback-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.doUnmock('../src/xlsx/collab.js');
    vi.doUnmock('../src/xlsx/exportStreaming.js');
    vi.unstubAllEnvs();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  test('a save still succeeds via the streaming fallback when exportXlsxFromGrid throws', async () => {
    vi.doMock('../src/xlsx/collab.js', async () => {
      const actual = await vi.importActual<typeof import('../src/xlsx/collab.js')>('../src/xlsx/collab.js');
      return {
        ...actual,
        exportXlsxFromGrid: vi.fn().mockRejectedValue(new Error('simulated: RangeError: Invalid array length')),
      };
    });

    const { store } = await import('../src/storage.js');
    const { joinEditSession, flushNow } = await import('../src/collab/room.js');

    const docId = randomUUID();
    const original = await makeSampleXlsx();
    await seedDoc(docId, original);

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);

    await expect(flushNow(docId)).resolves.toBe(true);

    const saved = await store.readOriginal(docId, 'xlsx');
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(saved as unknown as ArrayBuffer);
    const sheet = reopened.getWorksheet('Sheet1')!;
    expect(sheet.getCell('A1').value).toBe('Name');
    expect(sheet.getCell('A2').value).toBe('Alice');
  });

  test('when both the normal path and the streaming fallback fail, the error propagates (not silently swallowed)', async () => {
    vi.doMock('../src/xlsx/collab.js', async () => {
      const actual = await vi.importActual<typeof import('../src/xlsx/collab.js')>('../src/xlsx/collab.js');
      return { ...actual, exportXlsxFromGrid: vi.fn().mockRejectedValue(new Error('normal path failed')) };
    });
    vi.doMock('../src/xlsx/exportStreaming.js', () => ({
      exportXlsxFromGridStreaming: vi.fn().mockRejectedValue(new Error('streaming fallback also failed')),
    }));

    const docId = randomUUID();
    await seedDoc(docId, await makeSampleXlsx());
    const { joinEditSession, flushNow } = await import('../src/collab/room.js');

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);

    await expect(flushNow(docId)).rejects.toThrow('streaming fallback also failed');
  });
});
