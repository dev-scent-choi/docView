import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, test, vi } from 'vitest';
import { makeSamplePdf } from './helpers.js';
import type { Converter } from '../src/converters/types.js';

async function loadDocumentsModuleWithTempDir() {
  const dir = mkdtempSync(path.join(os.tmpdir(), 'docview-documents-test-'));
  vi.resetModules();
  process.env.DOCVIEW_DATA_DIR = dir;
  process.env.DOCVIEW_STORAGE = 'disk';
  const storageModule = await import('../src/storage.js');
  await storageModule.ensureDataDirs();
  const documentsModule = await import('../src/routes/documents.js');
  return { dir, storageModule, documentsModule };
}

describe('documents.ts: convertToPdf persists pagesizes', () => {
  let tempDirs: string[] = [];

  afterEach(() => {
    delete process.env.DOCVIEW_DATA_DIR;
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
    for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true });
    tempDirs = [];
  });

  test('after convertToPdf, pagesizes derived data matches the converted PDF page sizes', async () => {
    const { dir, storageModule, documentsModule } = await loadDocumentsModuleWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();
    await store.saveOriginal(id, 'docx', Buffer.from('fake docx bytes', 'utf8'));
    await store.writeMeta({
      id,
      name: 'a.docx',
      ext: 'docx',
      status: 'processing',
      pageCount: 0,
      createdAt: new Date().toISOString(),
    });
    const meta = (await store.readMeta(id))!;

    const samplePdf = await makeSamplePdf(2);
    const fakeConverter: Converter = {
      name: 'fake',
      extensions: ['docx'],
      convert: async (_source, targetPdfPath) => {
        writeFileSync(targetPdfPath, samplePdf);
      },
    };

    const { pageCount } = await documentsModule.convertToPdf(meta, fakeConverter);
    expect(pageCount).toBe(2);
    expect(await store.hasDerived(id, 'pagesizes')).toBe(true);
    const sizes = JSON.parse((await store.readDerived(id, 'pagesizes')).toString('utf8'));
    expect(sizes).toEqual([
      { width: 595, height: 842 },
      { width: 595, height: 842 },
    ]);
  });
});
