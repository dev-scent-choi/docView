import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { strFromU8, unzipSync, zipSync } from 'fflate';
import { PPTX_SLIDES_ARRAY, type SlideEntry } from '../src/collab/pptxShared.js';
import { exportPptxFromSlides, seedSlidesFromPptx } from '../src/pptx/collab.js';
import { makeSamplePng, makeSamplePptx } from './helpers.js';

/** zip 로컬 파일 헤더를 순서대로 파싱해 엔트리별 압축 방식(0=저장/무압축, 8=deflate)을 읽는다 —
 * fflate의 zipSync는 엔트리마다 어떤 압축 레벨이 실제로 적용됐는지 공개 API로 조회할 방법을 주지
 * 않으므로, 산출된 zip 바이트 자체(ZIP 스펙의 로컬 파일 헤더: signature@0, method@8(2B),
 * compressedSize@18(4B), nameLen@26(2B), extraLen@28(2B), name@30)를 직접 읽어 검증한다.
 * fflate가 데이터 기술자(data descriptor) 없이 실제 크기를 로컬 헤더에 바로 쓰기 때문에(스트리밍이
 * 아닌 zipSync 경로) compressedSize만으로 다음 엔트리 오프셋을 정확히 계산할 수 있다. */
function readZipCompressionMethods(zipBytes: Uint8Array): Map<string, number> {
  const view = new DataView(zipBytes.buffer, zipBytes.byteOffset, zipBytes.byteLength);
  const methods = new Map<string, number>();
  let offset = 0;
  while (offset + 30 <= zipBytes.length && view.getUint32(offset, true) === 0x04034b50) {
    const method = view.getUint16(offset + 8, true);
    const compressedSize = view.getUint32(offset + 18, true);
    const nameLen = view.getUint16(offset + 26, true);
    const extraLen = view.getUint16(offset + 28, true);
    const name = strFromU8(zipBytes.subarray(offset + 30, offset + 30 + nameLen));
    methods.set(name, method);
    offset += 30 + nameLen + extraLen + compressedSize;
  }
  return methods;
}

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

  test('survivingEntryIds matches finalParts 1:1, in the same order, using pre-save ids', () => {
    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    seedSlidesFromPptx(makeSamplePptx(['First', 'Second']), slidesArray);
    const idsBeforeExport = slidesArray.toArray().map((e) => e.id);

    const original = makeSamplePptx(['First', 'Second']);
    const result = exportPptxFromSlides(original, slidesArray);

    expect(result.survivingEntryIds).toEqual(idsBeforeExport);
    expect(result.survivingEntryIds).toHaveLength(result.finalParts.length);
  });

  test('survivingEntryIds drops ids of entries filterValidEntries rejects (broken sourceSlidePart)', () => {
    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    seedSlidesFromPptx(makeSamplePptx(['First', 'Second']), slidesArray);
    const survivingIdBefore = slidesArray.get(0).id;
    // 두 번째 항목을 존재하지 않는 슬라이드 파트를 가리키도록 깨뜨린다.
    const broken = { ...slidesArray.get(1), sourceSlidePart: 'slide999.xml' };
    ydoc.transact(() => {
      slidesArray.delete(1, 1);
      slidesArray.insert(1, [broken]);
    });

    const original = makeSamplePptx(['First', 'Second']);
    const result = exportPptxFromSlides(original, slidesArray);

    expect(result.survivingEntryIds).toEqual([survivingIdBefore]); // 깨진 두 번째 항목은 빠짐
  });
});

// --- 성능 최적화: ppt/media/ 아래 이미지처럼 이미 압축된 바이너리는 zipSync 단계에서 재압축을
// 건너뛴다(level 0/저장) — 이 아카이브의 최종 zipSync 호출은 편집 여부와 무관하게 매번 전체
// 엔트리에 적용되는 유일한 비용이라, 자동저장/수동저장/다운로드 전부가 이 압축 레벨 선택에
// 영향을 받는다. 아래 테스트는 (1) 그 선택이 실제로 산출된 zip 바이트에 반영되는지(로컬 파일
// 헤더의 압축 방식 필드로 직접 검증), (2) 그 선택이 압축 해제 결과(실제 콘텐츠 바이트)를 단
// 1바이트도 바꾸지 않는지를 함께 확인한다 — 순수 CPU 비용 최적화이지 콘텐츠 변경이 아니라는 것.
describe('exportPptxFromSlides — media entries skip recompression', () => {
  test('ppt/media/ 아래 이미지는 level 0(저장)으로, 슬라이드 XML은 그대로 deflate로 저장된다', () => {
    const png = makeSamplePng();
    const baseEntries = unzipSync(makeSamplePptx(['A']));
    const original = Buffer.from(zipSync({ ...baseEntries, 'ppt/media/image1.png': png }));

    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    seedSlidesFromPptx(original, slidesArray);

    const { bytes: out } = exportPptxFromSlides(original, slidesArray);
    const methods = readZipCompressionMethods(out);

    expect(methods.get('ppt/media/image1.png')).toBe(0); // 저장만, deflate 안 함
    expect(methods.get('ppt/slides/slide1.xml')).toBe(8); // 기존과 동일하게 deflate
    expect(methods.get('ppt/presentation.xml')).toBe(8);
  });

  test('미디어 엔트리는 압축 레벨과 무관하게 압축 해제 결과가 원본과 byte-for-byte 동일하다', () => {
    const png = makeSamplePng();
    const baseEntries = unzipSync(makeSamplePptx(['A']));
    const original = Buffer.from(zipSync({ ...baseEntries, 'ppt/media/image1.png': png }));

    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    seedSlidesFromPptx(original, slidesArray);

    const { bytes: out } = exportPptxFromSlides(original, slidesArray);
    const outEntries = unzipSync(out);

    expect(outEntries['ppt/media/image1.png']).toEqual(new Uint8Array(png));
  });

  test('ppt/media/ 밖에 있는 같은 확장자 파일은 압축 대상에서 제외되지 않는다(경로 프리픽스 검사)', () => {
    const png = makeSamplePng();
    const baseEntries = unzipSync(makeSamplePptx(['A']));
    // 실제로 PowerPoint가 만들어내지는 않지만, isPrecompressedMediaPath가 확장자만이 아니라
    // "ppt/media/ 아래"라는 경로 조건까지 함께 봐야 한다는 것을 확인해 둔다.
    const original = Buffer.from(zipSync({ ...baseEntries, 'ppt/notmedia/image1.png': png }));

    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    seedSlidesFromPptx(original, slidesArray);

    const { bytes: out } = exportPptxFromSlides(original, slidesArray);
    const methods = readZipCompressionMethods(out);

    expect(methods.get('ppt/notmedia/image1.png')).toBe(8);
  });
});
