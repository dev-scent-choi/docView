import { describe, expect, test } from 'vitest';
import { EDITABLE_EXTS, editorPathFor } from '../src/routes/documents.js';

describe('pptx editing routing', () => {
  test('pptx is editable', () => {
    expect(EDITABLE_EXTS.has('pptx')).toBe(true);
  });
  test('pptx routes to editor-slides.html', () => {
    expect(editorPathFor('pptx')).toBe('/editor-slides.html');
  });
});
