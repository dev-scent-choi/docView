import fs from 'node:fs/promises';
import { PDFDocument } from 'pdf-lib';
import { pdf } from 'pdf-to-img';
import sharp from 'sharp';

const TRIM_RENDER_SCALE = 1;
/** 잘라낸 뒤 내용 주변에 남길 여백 비율 */
const PADDING_RATIO = 0.015;
/** 내용이 페이지의 이 비율 이상을 차지하면 크롭하지 않는다 */
const MAX_CONTENT_AREA_RATIO = 0.8;
const TRIM_THRESHOLD = 10;

/**
 * CAD 변환 PDF는 도면이 거대한 빈 페이지 한가운데 작게 놓이는 경우가 많다.
 * 페이지를 렌더링해 내용 영역(bounding box)을 찾고 CropBox를 좁혀
 * 뷰어에서 도면이 화면을 가득 채우도록 만든다.
 */
export async function trimPdfMargins(pdfPath: string): Promise<void> {
  const bytes = await fs.readFile(pdfPath);
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
  const rendered = await pdf(pdfPath, { scale: TRIM_RENDER_SCALE });
  let changed = false;

  for (let index = 0; index < doc.getPageCount(); index += 1) {
    const png = await rendered.getPage(index + 1);
    const full = await sharp(png).metadata();
    if (!full.width || !full.height) continue;

    const { info } = await sharp(png)
      .flatten({ background: '#ffffff' })
      .trim({ threshold: TRIM_THRESHOLD })
      .toBuffer({ resolveWithObject: true });

    const contentLeft = -(info.trimOffsetLeft ?? 0);
    const contentTop = -(info.trimOffsetTop ?? 0);
    const contentAreaRatio = (info.width * info.height) / (full.width * full.height);
    if (contentAreaRatio >= MAX_CONTENT_AREA_RATIO || info.width <= 0 || info.height <= 0) {
      continue;
    }

    const page = doc.getPage(index);
    const { width: pageWidth, height: pageHeight } = page.getSize();
    const scaleX = pageWidth / full.width;
    const scaleY = pageHeight / full.height;
    const padding = Math.max(pageWidth, pageHeight) * PADDING_RATIO;

    const x = Math.max(0, contentLeft * scaleX - padding);
    const yFromTop = Math.max(0, contentTop * scaleY - padding);
    const width = Math.min(pageWidth - x, info.width * scaleX + padding * 2);
    const height = Math.min(pageHeight - yFromTop, info.height * scaleY + padding * 2);
    const y = pageHeight - yFromTop - height;

    page.setCropBox(x, y, width, height);
    page.setMediaBox(x, y, width, height);
    changed = true;
  }

  if (changed) {
    await fs.writeFile(pdfPath, await doc.save());
  }
}
