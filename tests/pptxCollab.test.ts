import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { PPTX_SLIDES_ARRAY, type SlideEntry } from '../src/collab/pptxShared.js';
import { seedSlidesFromPptx } from '../src/pptx/collab.js';
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
