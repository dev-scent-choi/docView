import { randomBytes } from 'node:crypto';
import { strToU8, zipSync } from 'fflate';
import { PDFDocument, StandardFonts } from 'pdf-lib';

/** fastify.inject용 multipart 본문을 직접 조립한다. */
export function multipartPayload(
  filename: string,
  contentType: string,
  data: Buffer,
): { payload: Buffer; headers: Record<string, string> } {
  const boundary = `----docviewtest${randomBytes(8).toString('hex')}`;
  const head = Buffer.from(
    `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="file"; filename="${filename}"\r\n` +
      `Content-Type: ${contentType}\r\n\r\n`,
  );
  const tail = Buffer.from(`\r\n--${boundary}--\r\n`);
  return {
    payload: Buffer.concat([head, data, tail]),
    headers: { 'content-type': `multipart/form-data; boundary=${boundary}` },
  };
}

export async function makeSamplePdf(pages = 2): Promise<Buffer> {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (let i = 0; i < pages; i += 1) {
    const page = doc.addPage([595, 842]);
    page.drawText(`DocView Search Target page ${i + 1}`, { x: 50, y: 760, size: 18, font });
  }
  return Buffer.from(await doc.save());
}

/** 1x1 픽셀 유효 PNG */
export function makeSamplePng(): Buffer {
  return Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  );
}

const CONTENT_TYPES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ' +
  'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/styles.xml" ' +
  'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
  '</Types>';

const ROOT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ' +
  'Target="word/document.xml"/>' +
  '</Relationships>';

const DOCUMENT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" ' +
  'Target="styles.xml"/>' +
  '</Relationships>';

/** Heading1~3 스타일 정의를 담은 최소 styles.xml — export한 문서를 Word가 열었을 때 제목 서식이 실제로 적용되도록 한다. */
const STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  [1, 2, 3]
    .map(
      (level) =>
        `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/>` +
        `<w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="${32 - level * 4}"/></w:rPr></w:style>`,
    )
    .join('') +
  '</w:styles>';

function documentXml(paragraphs: readonly string[]): string {
  const body = paragraphs
    .map((text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`)
    .join('');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}<w:sectPr/></w:body></w:document>`
  );
}

/**
 * 최소한의 유효한 docx를 직접 zip으로 조립한다(Word/LibreOffice 없이).
 * word/document.xml, styles.xml, 필수 관계·콘텐츠 타입 파트만 포함한다.
 */
export function makeSampleDocx(paragraphs: readonly string[] = ['Hello DocView']): Buffer {
  const entries = {
    '[Content_Types].xml': strToU8(CONTENT_TYPES_XML),
    '_rels/.rels': strToU8(ROOT_RELS),
    'word/document.xml': strToU8(documentXml(paragraphs)),
    'word/_rels/document.xml.rels': strToU8(DOCUMENT_RELS),
    'word/styles.xml': strToU8(STYLES_XML),
  };
  return Buffer.from(zipSync(entries));
}
