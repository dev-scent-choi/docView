import fs from 'node:fs/promises';
import path from 'node:path';
import { PDFDocument } from 'pdf-lib';
import type { Converter } from './types.js';

export const imageConverter: Converter = {
  name: 'image',
  extensions: ['png', 'jpg', 'jpeg'],
  async convert(sourcePath, targetPdfPath) {
    const bytes = await fs.readFile(sourcePath);
    const ext = path.extname(sourcePath).slice(1).toLowerCase();
    const doc = await PDFDocument.create();
    const image = ext === 'png' ? await doc.embedPng(bytes) : await doc.embedJpg(bytes);
    const page = doc.addPage([image.width, image.height]);
    page.drawImage(image, { x: 0, y: 0, width: image.width, height: image.height });
    await fs.writeFile(targetPdfPath, await doc.save());
  },
};
