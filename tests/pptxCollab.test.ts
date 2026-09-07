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

/** 주어진 슬라이드 파트(예: 'slide2.xml')가 presentation.xml.rels/presentation.xml에서 쓰는
 * (rId, sldId 숫자) 쌍을 찾는다. 살아남은 'original' 슬라이드가 내보내기 전후로 이 값을
 * 그대로 유지하는지 검증하는 데 쓴다(p14:sectionLst/custShowLst 등 외부 참조가 깨지지 않도록). */
function relIdAndSldIdFor(pptxBytes: Uint8Array, part: string): { relId: string; sldId: number } {
  const entries = unzipSync(pptxBytes);
  const presentationXml = strFromU8(entries['ppt/presentation.xml']);
  const presentationRelsXml = strFromU8(entries['ppt/_rels/presentation.xml.rels']);
  const relTag = (presentationRelsXml.match(/<Relationship\b[^>]*\/>/g) ?? []).find((tag) =>
    new RegExp(`Target="(?:[^"]*/)?${part.replace('.', '\\.')}"`).test(tag),
  );
  const relId = relTag ? /Id="(rId\d+)"/.exec(relTag)?.[1] : undefined;
  if (!relId) throw new Error(`관계를 찾을 수 없습니다: ${part}`);
  const sldIdTag = (presentationXml.match(/<p:sldId\b[^>]*\/>/g) ?? []).find((tag) =>
    tag.includes(`r:id="${relId}"`),
  );
  const sldIdMatch = sldIdTag ? /\bid="(\d+)"/.exec(sldIdTag) : null;
  if (!sldIdMatch) throw new Error(`sldId를 찾을 수 없습니다: ${part}`);
  return { relId, sldId: Number(sldIdMatch[1]) };
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
    const before = relIdAndSldIdFor(original, 'slide2.xml');
    const { bytes: out } = exportPptxFromSlides(original, slidesArray);
    const entries = unzipSync(out);
    expect(strFromU8(entries['ppt/slides/slide1.xml'])).toContain('A'); // 원본 파트 그대로
    expect(slideOrderOf(out)).toEqual(['slide2.xml', 'slide3.xml', 'slide1.xml']);
    // 살아남은 원본 슬라이드는 순서만 바뀌고 rId/sldId는 그대로 유지되어야 한다(섹션/사용자 지정
    // 쇼처럼 이 값들을 참조하는 PowerPoint 기능이 재정렬만으로 깨지지 않도록).
    expect(relIdAndSldIdFor(out, 'slide2.xml')).toEqual(before);
  });

  test('deleting a slide removes its part and shrinks the presentation', () => {
    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    seedSlidesFromPptx(makeSamplePptx(['A', 'B', 'C']), slidesArray);
    slidesArray.delete(1, 1); // "B" 삭제

    const original = makeSamplePptx(['A', 'B', 'C']);
    const { bytes: out } = exportPptxFromSlides(original, slidesArray);
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
    const { bytes: out } = exportPptxFromSlides(original, slidesArray);
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
    const { bytes: out } = exportPptxFromSlides(original, slidesArray);
    const entries = unzipSync(out);
    expect(slideOrderOf(out)).toEqual(['slide1.xml', 'slide2.xml', 'slide3.xml']);
    expect(entries['ppt/slides/slide3.xml']).toBeDefined();
    expect(entries['ppt/slides/_rels/slide3.xml.rels']).toBeDefined();
    expect(strFromU8(entries['ppt/slides/_rels/slide3.xml.rels'])).toContain('slideLayout1.xml');
    expect(strFromU8(entries['[Content_Types].xml'])).toContain('/ppt/slides/slide3.xml');
  });

  // --- Critical 1 회귀: 첫 저장의 duplicate가 그 이후 사라진 원본을 가리키는 stale 상태가 되어도
  // 다음 저장이 크래시하지 않아야 한다. ---
  test('a duplicate entry whose sourceSlidePart no longer exists in the zip is silently dropped, not thrown', () => {
    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    // 원본에는 slide1.xml만 있다 — 배열에는 이미 사라진 slide9.xml을 가리키는 stale duplicate가
    // 섞여 있는 상태를 흉내낸다(예: 직전 저장에서 원본이 삭제된 뒤에도 남아있던 항목).
    slidesArray.insert(0, [
      { id: 'orig-1', kind: 'original', sourceSlidePart: 'slide1.xml', sourcePageNumber: 1 },
      { id: 'dup-stale', kind: 'duplicate', sourceSlidePart: 'slide9.xml', sourcePageNumber: null },
    ]);

    const original = makeSamplePptx(['A']);
    expect(() => exportPptxFromSlides(original, slidesArray)).not.toThrow();
    const { bytes: out, finalParts } = exportPptxFromSlides(original, slidesArray);
    expect(finalParts).toEqual(['slide1.xml']);
    expect(slideOrderOf(out)).toEqual(['slide1.xml']);
  });

  test('a second export using the array-rewrite output (finalParts as fresh originals) succeeds and is idempotent', () => {
    // flush #1: 2장짜리 덱에서 슬라이드 1을 복제 — duplicate가 slide1.xml을 가리킨다.
    const ydoc1 = new Y.Doc();
    const slidesArray1 = slidesArrayFrom(ydoc1);
    seedSlidesFromPptx(makeSamplePptx(['A', 'B']), slidesArray1);
    const source = slidesArray1.get(0);
    slidesArray1.insert(1, [
      { id: 'dup-1', kind: 'duplicate', sourceSlidePart: source.sourceSlidePart, sourcePageNumber: null },
    ]);
    const original1 = makeSamplePptx(['A', 'B']);
    const { bytes: afterFlush1, finalParts } = exportPptxFromSlides(original1, slidesArray1);
    expect(finalParts).toEqual(['slide1.xml', 'slide3.xml', 'slide2.xml']);

    // room.ts의 flushToStore가 flush #1 직후에 하는 것과 동일하게, 배열을 finalParts 순서의
    // 'original' 항목들로 재구성한다.
    const ydoc2 = new Y.Doc();
    const slidesArray2 = slidesArrayFrom(ydoc2);
    slidesArray2.insert(
      0,
      finalParts.map((part, index) => ({
        id: `rewritten-${index}`,
        kind: 'original' as const,
        sourceSlidePart: part,
        sourcePageNumber: index + 1,
      })),
    );

    // flush #2: 새로 재구성된 배열로, 이제 디스크에 있는 afterFlush1(원본이 flush #1의 결과물)을
    // 대상으로 다시 export한다 — 크래시하지 않고, 슬라이드 개수도 그대로 유지되어야 한다(멱등).
    expect(() => exportPptxFromSlides(afterFlush1, slidesArray2)).not.toThrow();
    const { bytes: afterFlush2 } = exportPptxFromSlides(afterFlush1, slidesArray2);
    expect(slideOrderOf(afterFlush2)).toEqual(['slide1.xml', 'slide3.xml', 'slide2.xml']);
  });

  // --- Critical 2 회귀: 남는 슬라이드가 하나도 없으면 저장 자체를 거부해야 한다. ---
  test('an entry list that filters down to zero entries throws a clear error', () => {
    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    slidesArray.insert(0, [
      { id: 'bad-1', kind: 'original', sourceSlidePart: 'not-a-slide-part', sourcePageNumber: 1 },
      { id: 'bad-2', kind: 'duplicate', sourceSlidePart: 'slide99.xml', sourcePageNumber: null },
    ]);
    const original = makeSamplePptx(['A', 'B']);
    expect(() => exportPptxFromSlides(original, slidesArray)).toThrow(
      '슬라이드가 하나도 남지 않아 저장할 수 없습니다.',
    );
  });

  // --- Critical 3 회귀: sourceSlidePart에 담긴 인젝션 시도는 절대 저장된 XML에 반영되지 않아야
  // 한다 — 'slideN.xml' 정확한 형식이 아닌 값은 통째로 걸러진다. ---
  test('an invalid sourceSlidePart used for XML injection is filtered out, never reaching the saved XML', () => {
    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    slidesArray.insert(0, [
      {
        id: 'evil-1',
        kind: 'original',
        sourceSlidePart: 'slide1.xml" TargetMode="External" Target="http://evil.example/x',
        sourcePageNumber: 1,
      },
      { id: 'safe-1', kind: 'original', sourceSlidePart: 'slide1.xml', sourcePageNumber: 1 },
    ]);
    const original = makeSamplePptx(['A']);
    const { bytes: out } = exportPptxFromSlides(original, slidesArray);
    const entries = unzipSync(out);
    expect(strFromU8(entries['ppt/_rels/presentation.xml.rels'])).not.toContain('evil.example');
    expect(strFromU8(entries['[Content_Types].xml'])).not.toContain('evil.example');
  });
});
