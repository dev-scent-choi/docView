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

describe('hwp/hwpx editing routing', () => {
  test('hwp and hwpx are editable', () => {
    expect(EDITABLE_EXTS.has('hwp')).toBe(true);
    expect(EDITABLE_EXTS.has('hwpx')).toBe(true);
  });
  test('hwp and hwpx route to editor-hwp.html (native @rhwp/editor, not the docx editor)', () => {
    expect(editorPathFor('hwp')).toBe('/editor-hwp.html');
    expect(editorPathFor('hwpx')).toBe('/editor-hwp.html');
  });
});
