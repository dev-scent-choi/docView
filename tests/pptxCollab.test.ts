import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { PPTX_SLIDES_ARRAY, type SlideEntry } from '../src/collab/pptxShared.js';
import { exportPptxFromSlides, seedSlidesFromPptx } from '../src/pptx/collab.js';
import { makeSamplePptx } from './helpers.js';

describe('seedSlidesFromPptx', () => {
  test('populates one original entry per slide, in presentation order', () => {
    const bytes = makeSamplePptx(['First', 'Second', 'Third']);
    const ydoc = new Y.Doc();
    const slidesArray = ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);

    seedSlidesFromPptx(bytes, slidesArray);

    expect(slidesArray.length).toBe(3);
    const entries = slidesArray.toArray();
    expect(entries.map((e) => e.kind)).toEqual(['original', 'original', 'original']);
    expect(entries.map((e) => e.sourceSlidePart)).toEqual(['slide1.xml', 'slide2.xml', 'slide3.xml']);
    expect(entries.map((e) => e.sourcePageNumber)).toEqual([1, 2, 3]);
    expect(entries.every((e) => typeof e.id === 'string' && e.id.length > 0)).toBe(true);
  });
});

function slidesArrayFrom(ydoc: Y.Doc): Y.Array<SlideEntry> {
  return ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);
}

function slideOrderOf(pptxBytes: Uint8Array): string[] {
  const entries = unzipSync(pptxBytes);
  const presentationXml = strFromU8(entries['ppt/presentation.xml']);
  const presentationRelsXml = strFromU8(entries['ppt/_rels/presentation.xml.rels']);
  const relMap = new Map<string, string>();
  for (const relTag of presentationRelsXml.match(/<Relationship\b[^>]*\/>/g) ?? []) {
    const idMatch = /Id="(rId\d+)"/.exec(relTag);
    const targetMatch = /Target="(?:[^"]*\/)?(slide\d+\.xml)"/.exec(relTag);
    if (idMatch && targetMatch) relMap.set(idMatch[1], targetMatch[1]);
  }
  const rIds = [...presentationXml.matchAll(/<p:sldId\b[^>]*r:id="(rId\d+)"/g)].map((m) => m[1]);
  return rIds.map((rId) => relMap.get(rId)!);
}

describe('exportPptxFromSlides', () => {
  test('reordering the array reorders the exported presentation', () => {
    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    seedSlidesFromPptx(makeSamplePptx(['A', 'B', 'C']), slidesArray);
    const [first] = slidesArray.slice(0, 1);
    ydoc.transact(() => {
      slidesArray.delete(0, 1);
      slidesArray.insert(2, [first]);
    });

    const original = makeSamplePptx(['A', 'B', 'C']);
    const out = exportPptxFromSlides(original, slidesArray);
    const entries = unzipSync(out);
    expect(strFromU8(entries['ppt/slides/slide1.xml'])).toContain('A'); // 원본 파트 그대로
    expect(slideOrderOf(out)).toEqual(['slide2.xml', 'slide3.xml', 'slide1.xml']);
  });

  test('deleting a slide removes its part and shrinks the presentation', () => {
    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    seedSlidesFromPptx(makeSamplePptx(['A', 'B', 'C']), slidesArray);
    slidesArray.delete(1, 1); // "B" 삭제

    const original = makeSamplePptx(['A', 'B', 'C']);
    const out = exportPptxFromSlides(original, slidesArray);
    const entries = unzipSync(out);
    expect(slideOrderOf(out)).toEqual(['slide1.xml', 'slide3.xml']);
    expect(entries['ppt/slides/slide2.xml']).toBeUndefined();
    expect(strFromU8(entries['[Content_Types].xml'])).not.toContain('/ppt/slides/slide2.xml');
    expect(strFromU8(entries['ppt/_rels/presentation.xml.rels'])).not.toContain('slide2.xml');
  });

  test('duplicating a slide copies its content byte-for-byte under a new part name', () => {
    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    seedSlidesFromPptx(makeSamplePptx(['A', 'B']), slidesArray);
    const source = slidesArray.get(0);
    slidesArray.insert(1, [{ id: 'dup-1', kind: 'duplicate', sourceSlidePart: source.sourceSlidePart, sourcePageNumber: null }]);

    const original = makeSamplePptx(['A', 'B']);
    const out = exportPptxFromSlides(original, slidesArray);
    const entries = unzipSync(out);
    expect(slideOrderOf(out)).toEqual(['slide1.xml', 'slide3.xml', 'slide2.xml']);
    expect(strFromU8(entries['ppt/slides/slide3.xml'])).toContain('A'); // "A" 슬라이드가 그대로 복제됨
    expect(entries['ppt/slides/_rels/slide3.xml.rels']).toBeDefined();
  });

  test('adding a blank slide inserts a minimal valid slide referencing the first layout', () => {
    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    seedSlidesFromPptx(makeSamplePptx(['A', 'B']), slidesArray);
    slidesArray.insert(2, [{ id: 'blank-1', kind: 'blank', sourceSlidePart: null, sourcePageNumber: null }]);

    const original = makeSamplePptx(['A', 'B']);
    const out = exportPptxFromSlides(original, slidesArray);
    const entries = unzipSync(out);
    expect(slideOrderOf(out)).toEqual(['slide1.xml', 'slide2.xml', 'slide3.xml']);
    expect(entries['ppt/slides/slide3.xml']).toBeDefined();
    expect(entries['ppt/slides/_rels/slide3.xml.rels']).toBeDefined();
    expect(strFromU8(entries['ppt/slides/_rels/slide3.xml.rels'])).toContain('slideLayout1.xml');
    expect(strFromU8(entries['[Content_Types].xml'])).toContain('/ppt/slides/slide3.xml');
  });
});
