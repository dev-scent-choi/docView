import fs from 'node:fs/promises';
import { PDFDocument } from 'pdf-lib';
import sharp from 'sharp';
import type { Converter } from './types.js';

/**
 * pdf-lib이 직접 임베드하지 못하는 래스터 포맷(TIFF/BMP/WebP/GIF)을
 * sharp로 PNG 페이지로 풀어 PDF로 조립한다. 다페이지 TIFF(스캔 도면/팩스)도
 * 페이지별로 나눠 담는다.
 */
export const rasterConverter: Converter = {
  name: 'raster-sharp',
  extensions: ['tif', 'tiff', 'bmp', 'webp', 'gif'],
  async convert(sourcePath, targetPdfPath) {
    const bytes = await fs.readFile(sourcePath);
    const metadata = await sharp(bytes).metadata();
    const pageCount = metadata.pages && metadata.pages > 0 ? metadata.pages : 1;

    const doc = await PDFDocument.create();
    for (let index = 0; index < pageCount; index += 1) {
      const png = await sharp(bytes, { page: index }).png().toBuffer();
      const image = await doc.embedPng(png);
      const page = doc.addPage([image.width, image.height]);
      page.drawImage(image, { x: 0, y: 0, width: image.width, height: image.height });
    }
    await fs.writeFile(targetPdfPath, await doc.save());
  },
};
