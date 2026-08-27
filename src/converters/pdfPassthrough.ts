import fs from 'node:fs/promises';
import type { Converter } from './types.js';

export const pdfPassthroughConverter: Converter = {
  name: 'pdf-passthrough',
  extensions: ['pdf'],
  async convert(sourcePath, targetPdfPath) {
    await fs.copyFile(sourcePath, targetPdfPath);
  },
};
