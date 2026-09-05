import { describe, expect, test } from 'vitest';
import { officeConverter } from '../src/converters/libreOffice.js';

describe('officeConverter', () => {
  test('supports pptx alongside the other LibreOffice-backed office formats', () => {
    expect(officeConverter.extensions).toContain('pptx');
    expect(officeConverter.extensions).toContain('ppt');
    expect(officeConverter.extensions).toContain('docx');
  });
});
