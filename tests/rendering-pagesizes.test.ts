import { describe, expect, test, vi } from 'vitest';
import { getPageDimensions, getPdfPageInfo, renderPage, type PageSize } from '../src/rendering.js';
import { makeSamplePdf } from './helpers.js';

function nullPersistence() {
  return {
    load: vi.fn(async () => null as readonly PageSize[] | null),
    save: vi.fn(async () => undefined),
  };
}

describe('rendering: getPdfPageInfo', () => {
  test('returns page count and per-page sizes from a single parse', async () => {
    const pdf = await makeSamplePdf(3);
    const info = await getPdfPageInfo(pdf);
    expect(info.pageCount).toBe(3);
    expect(info.sizes).toEqual([
      { width: 595, height: 842 },
      { width: 595, height: 842 },
      { width: 595, height: 842 },
    ]);
  });
});

describe('rendering: getPageDimensions with persisted page sizes', () => {
  test('uses persisted sizes without ever calling loadBytes when they exist', async () => {
    const persisted = {
      load: vi.fn(async () => [{ width: 100, height: 200 }] as readonly PageSize[]),
      save: vi.fn(async () => undefined),
    };
    const loadBytes = vi.fn(async () => {
      throw new Error('loadBytes should not be called when persisted sizes exist');
    });
    const dims = await getPageDimensions(`doc-${Math.random()}`, loadBytes, persisted);
    expect(dims).toEqual([{ w: 100, h: 200 }]);
    expect(loadBytes).not.toHaveBeenCalled();
    expect(persisted.save).not.toHaveBeenCalled();
  });

  test('falls back to parsing bytes and self-heals by calling persisted.save when nothing is persisted yet', async () => {
    const pdf = await makeSamplePdf(2);
    const persisted = nullPersistence();
    const loadBytes = vi.fn(async () => pdf);
    const dims = await getPageDimensions(`doc-${Math.random()}`, loadBytes, persisted);
    expect(dims).toEqual([
      { w: 595, h: 842 },
      { w: 595, h: 842 },
    ]);
    expect(loadBytes).toHaveBeenCalledTimes(1);
    expect(persisted.save).toHaveBeenCalledTimes(1);
    expect(persisted.save).toHaveBeenCalledWith([
      { width: 595, height: 842 },
      { width: 595, height: 842 },
    ]);
  });
});

describe('rendering: renderPage with persisted page sizes', () => {
  test('renders using persisted sizes, calling loadBytes only once (for rasterization, not the size lookup)', async () => {
    const pdf = await makeSamplePdf(1);
    const persisted = {
      load: vi.fn(async () => [{ width: 595, height: 842 }] as readonly PageSize[]),
      save: vi.fn(async () => undefined),
    };
    // loadBytes는 renderPage 내부의 실제 래스터화(openDocument)에는 여전히 필요하므로 진짜
    // PDF를 준다 — 대신 호출 횟수를 세어, persisted가 있으면 "크기 조회" 몫으로는 단 한 번도
    // 더 불리지 않는다는 것(래스터화용 1회만 호출됨)을 확인한다.
    const loadBytes = vi.fn(async () => pdf);
    const png = await renderPage(`doc-${Math.random()}`, 1, 1, loadBytes, persisted);
    expect(png.length).toBeGreaterThan(0);
    expect(persisted.load).toHaveBeenCalledTimes(1);
    expect(loadBytes).toHaveBeenCalledTimes(1);
  });
});
