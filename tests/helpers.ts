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

const PPTX_CONTENT_TYPES_HEAD =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/ppt/presentation.xml" ' +
  'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
  '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ' +
  'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
  '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ' +
  'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>';

const PPTX_ROOT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ' +
  'Target="ppt/presentation.xml"/>' +
  '</Relationships>';

const PPTX_SLIDE_MASTER_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/></p:spTree></p:cSld>' +
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" ' +
  'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" ' +
  'folHlink="folHlink"/>' +
  '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
  '</p:sldMaster>';

const PPTX_SLIDE_MASTER_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" ' +
  'Target="../slideLayouts/slideLayout1.xml"/>' +
  '</Relationships>';

const PPTX_SLIDE_LAYOUT_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>';

const PPTX_SLIDE_LAYOUT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" ' +
  'Target="../slideMasters/slideMaster1.xml"/>' +
  '</Relationships>';

function pptxSlideXml(text: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
    '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr/>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    `<p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>` +
    '</p:spTree></p:cSld></p:sld>'
  );
}

const PPTX_SLIDE_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" ' +
  'Target="../slideLayouts/slideLayout1.xml"/>' +
  '</Relationships>';

/**
 * 최소한의 유효한 pptx를 직접 zip으로 조립한다(PowerPoint/LibreOffice 없이).
 * slideTexts.length장의 슬라이드를 만들고, 전부 같은 slideMaster/slideLayout 하나를 공유한다.
 */
export function makeSamplePptx(slideTexts: readonly string[] = ['Slide 1']): Buffer {
  const slideOverrides = slideTexts
    .map(
      (_, i) =>
        `<Override PartName="/ppt/slides/slide${i + 1}.xml" ` +
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>',
    )
    .join('');
  const contentTypesXml = `${PPTX_CONTENT_TYPES_HEAD}${slideOverrides}</Types>`;

  const sldIdLst = slideTexts
    .map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`)
    .join('');
  const presentationXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
    `<p:sldIdLst>${sldIdLst}</p:sldIdLst>` +
    '<p:sldSz cx="9144000" cy="6858000"/>' +
    '</p:presentation>';

  const presentationRelsEntries = [
    '<Relationship Id="rId1" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" ' +
      'Target="slideMasters/slideMaster1.xml"/>',
    ...slideTexts.map(
      (_, i) =>
        `<Relationship Id="rId${i + 2}" ` +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" ' +
        `Target="slides/slide${i + 1}.xml"/>`,
    ),
  ].join('');
  const presentationRelsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `${presentationRelsEntries}</Relationships>`;

  const entries: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(contentTypesXml),
    '_rels/.rels': strToU8(PPTX_ROOT_RELS),
    'ppt/presentation.xml': strToU8(presentationXml),
    'ppt/_rels/presentation.xml.rels': strToU8(presentationRelsXml),
    'ppt/slideMasters/slideMaster1.xml': strToU8(PPTX_SLIDE_MASTER_XML),
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': strToU8(PPTX_SLIDE_MASTER_RELS),
    'ppt/slideLayouts/slideLayout1.xml': strToU8(PPTX_SLIDE_LAYOUT_XML),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': strToU8(PPTX_SLIDE_LAYOUT_RELS),
  };
  slideTexts.forEach((text, i) => {
    entries[`ppt/slides/slide${i + 1}.xml`] = strToU8(pptxSlideXml(text));
    entries[`ppt/slides/_rels/slide${i + 1}.xml.rels`] = strToU8(PPTX_SLIDE_RELS);
  });

  return Buffer.from(zipSync(entries));
}
