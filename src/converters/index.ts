import { dwg2dxfPath, dwgConverter, odaConverterPath } from './dwgConverter.js';
import { docxTextPreviewConverter } from './docxTextPreview.js';
import { imageConverter } from './imageConverter.js';
import {
  dxfConverter,
  gridPrintConverter,
  officeConverter,
  sofficePath,
  spreadsheetConverter,
} from './libreOffice.js';
import { pdfPassthroughConverter } from './pdfPassthrough.js';
import { rasterConverter } from './rasterConverter.js';
import type { Converter } from './types.js';

const converters: readonly Converter[] = [
  pdfPassthroughConverter,
  imageConverter,
  rasterConverter,
  // xlsx/csv 그리드 열람은 LibreOffice 없이 동작하므로 항상 등록한다.
  // (출력용 PDF 지연 생성만 LibreOffice가 필요하며, 그 시점에 없으면 별도로 실패한다)
  gridPrintConverter,
  // LibreOffice가 있으면 정확한 렌더링(officeConverter)을 쓰고, 없으면 docx만은 서식 없는
  // 텍스트 미리보기로라도 업로드·동시편집이 가능하게 한다(동시편집 대상이 docx뿐이라 이 정도로 충분).
  ...(sofficePath ? [officeConverter, spreadsheetConverter, dxfConverter] : [docxTextPreviewConverter]),
  // dwgConverter 자체는 ODA File Converter/LibreDWG(dwg2dxf) 둘 중 하나만 있어도 동작하는데,
  // 등록 여부를 dwg2dxfPath만으로 판단하면 ODA만 설치된 환경에서는 실제로는 변환 가능한데도
  // .dwg 업로드가 415로 거부됐다 — 둘 중 하나라도 있으면 등록한다.
  ...(sofficePath && (dwg2dxfPath || odaConverterPath) ? [dwgConverter] : []),
];

const byExtension = new Map<string, Converter>(
  converters.flatMap((converter) => converter.extensions.map((ext) => [ext, converter] as const)),
);

export function findConverter(ext: string): Converter | null {
  return byExtension.get(ext.toLowerCase()) ?? null;
}

export const supportedExtensions: readonly string[] = [...byExtension.keys()];
