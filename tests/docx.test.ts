import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import { exportDocx, renderDocumentXml } from '../src/docx/export.js';
import { getDocxDefaultFontFamily, getDocxDefaultFontSizePt, importDocx } from '../src/docx/import.js';
import { docxSchema } from '../src/docx/schema.js';
import { makeSampleDocx } from './helpers.js';

describe('docx import', () => {
  test('extracts paragraph text in order', () => {
    const bytes = makeSampleDocx(['First paragraph', 'Second paragraph']);
    const { doc, hasUnsupportedContent } = importDocx(bytes);
    expect(hasUnsupportedContent).toBe(false);
    expect(doc.childCount).toBe(2);
    expect(doc.child(0).type.name).toBe('paragraph');
    expect(doc.child(0).textContent).toBe('First paragraph');
    expect(doc.child(1).textContent).toBe('Second paragraph');
  });

  test('flags documents containing drawings as having unsupported content', () => {
    const withDrawing = makeSampleDocx(['plain text']);
    const entries = unzipSync(withDrawing);
    const xml = strFromU8(entries['word/document.xml']).replace(
      '<w:sectPr/>',
      '<w:p><w:r><w:drawing/></w:r></w:p><w:sectPr/>',
    );
    entries['word/document.xml'] = strToU8(xml);
    const rezipped = Buffer.from(zipSync(entries));
    const { hasUnsupportedContent } = importDocx(rezipped);
    expect(hasUnsupportedContent).toBe(true);
  });
});

describe('docx import: table merges', () => {
  function docxWithTableXml(tableXml: string): Buffer {
    const base = makeSampleDocx([]);
    const entries = unzipSync(base);
    const xml = strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', `${tableXml}<w:sectPr/>`);
    entries['word/document.xml'] = strToU8(xml);
    return Buffer.from(zipSync(entries));
  }

  test('reads w:gridSpan as colspan on the merged cell', () => {
    const tableXml =
      '<w:tbl><w:tr>' +
      '<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Merged</w:t></w:r></w:p></w:tc>' +
      '</w:tr><w:tr>' +
      '<w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>';
    const { doc } = importDocx(docxWithTableXml(tableXml));
    const table = doc.child(0);
    expect(table.type.name).toBe('table');
    const row1 = table.child(0);
    expect(row1.childCount).toBe(1); // 병합된 칸은 하나의 셀로만 존재해야 한다
    expect(row1.child(0).attrs.colspan).toBe(2);
    expect(row1.child(0).textContent).toBe('Merged');
    const row2 = table.child(1);
    expect(row2.childCount).toBe(2);
  });

  test('reads w:vMerge restart/continue as rowspan and drops the continuation cell', () => {
    const tableXml =
      '<w:tbl><w:tr>' +
      '<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Tall</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>' +
      '</w:tr><w:tr>' +
      '<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>' +
      '<w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>';
    const { doc } = importDocx(docxWithTableXml(tableXml));
    const table = doc.child(0);
    const row1 = table.child(0);
    const row2 = table.child(1);
    expect(row1.childCount).toBe(2);
    expect(row1.child(0).attrs.rowspan).toBe(2);
    expect(row1.child(0).textContent).toBe('Tall');
    // 세로 병합의 "이어지는 칸"은 별개 셀로 남지 않아야 한다 — row2엔 B2 하나뿐
    expect(row2.childCount).toBe(1);
    expect(row2.child(0).textContent).toBe('B2');
  });
});

describe('docx import: table cell width and shading', () => {
  function docxWithTableXml(tableXml: string): Buffer {
    const base = makeSampleDocx([]);
    const entries = unzipSync(base);
    const xml = strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', `${tableXml}<w:sectPr/>`);
    entries['word/document.xml'] = strToU8(xml);
    return Buffer.from(zipSync(entries));
  }

  test('reads column width from w:tblGrid as colwidth (dxa -> px, 1440dxa = 96px)', () => {
    const tableXml =
      '<w:tbl><w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="2880"/></w:tblGrid><w:tr>' +
      '<w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>';
    const { doc } = importDocx(docxWithTableXml(tableXml));
    const row = doc.child(0).child(0);
    expect(row.child(0).attrs.colwidth).toEqual([96]);
    expect(row.child(1).attrs.colwidth).toEqual([192]);
  });

  test('gives a colspan-2 cell one width per spanned column, taken from w:tblGrid', () => {
    const tableXml =
      '<w:tbl><w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="2880"/></w:tblGrid><w:tr>' +
      '<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Merged</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>';
    const { doc } = importDocx(docxWithTableXml(tableXml));
    const cell = doc.child(0).child(0).child(0);
    expect(cell.attrs.colwidth).toEqual([96, 192]);
  });

  test('reads w:shd fill as the cell background color', () => {
    const tableXml =
      '<w:tbl><w:tr>' +
      '<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FFD43B"/></w:tcPr><w:p><w:r><w:t>Shaded</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>Plain</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>';
    const { doc } = importDocx(docxWithTableXml(tableXml));
    const row = doc.child(0).child(0);
    expect(row.child(0).attrs.background).toBe('#ffd43b');
    expect(row.child(1).attrs.background).toBeNull();
  });

  test('treats w:shd fill="auto" (no real fill) as no background', () => {
    const tableXml =
      '<w:tbl><w:tr>' +
      '<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:tcPr><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>';
    const { doc } = importDocx(docxWithTableXml(tableXml));
    expect(doc.child(0).child(0).child(0).attrs.background).toBeNull();
  });
});

describe('docx import: table borders', () => {
  function docxWithTableXml(tableXml: string): Buffer {
    const base = makeSampleDocx([]);
    const entries = unzipSync(base);
    const xml = strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', `${tableXml}<w:sectPr/>`);
    entries['word/document.xml'] = strToU8(xml);
    return Buffer.from(zipSync(entries));
  }

  const TABLE_BORDERS =
    '<w:tblBorders>' +
    '<w:top w:val="single" w:sz="4" w:color="000000"/>' +
    '<w:left w:val="single" w:sz="4" w:color="000000"/>' +
    '<w:bottom w:val="single" w:sz="4" w:color="000000"/>' +
    '<w:right w:val="single" w:sz="4" w:color="000000"/>' +
    '<w:insideH w:val="dashed" w:sz="8" w:color="FF0000"/>' +
    '<w:insideV w:val="dashed" w:sz="8" w:color="FF0000"/>' +
    '</w:tblBorders>';

  test('uses the outer border on edge cells and the inside border between cells (2x2)', () => {
    const tableXml =
      `<w:tbl><w:tblPr>${TABLE_BORDERS}</w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>` +
      '<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>';
    const { doc } = importDocx(docxWithTableXml(tableXml));
    const table = doc.child(0);
    const a1 = table.child(0).child(0).attrs.borders;
    // A1은 왼쪽·위가 바깥 변(검정 실선), 오른쪽·아래는 안쪽 grid선(빨간 점선)이어야 한다
    expect(a1.top).toEqual({ style: 'solid', widthPx: 1, color: '#000000' });
    expect(a1.left).toEqual({ style: 'solid', widthPx: 1, color: '#000000' });
    expect(a1.right).toEqual({ style: 'dashed', widthPx: 1.5, color: '#ff0000' });
    expect(a1.bottom).toEqual({ style: 'dashed', widthPx: 1.5, color: '#ff0000' });

    const b2 = table.child(1).child(1).attrs.borders;
    // B2(오른쪽 아래 끝 칸)는 오른쪽·아래가 바깥 변이어야 한다
    expect(b2.right).toEqual({ style: 'solid', widthPx: 1, color: '#000000' });
    expect(b2.bottom).toEqual({ style: 'solid', widthPx: 1, color: '#000000' });
  });

  test('a cell-level w:tcBorders overrides the table default for that cell only', () => {
    const tableXml =
      `<w:tbl><w:tblPr>${TABLE_BORDERS}</w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>` +
      '<w:tr>' +
      '<w:tc><w:tcPr><w:tcBorders><w:top w:val="single" w:sz="24" w:color="00FF00"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>';
    const { doc } = importDocx(docxWithTableXml(tableXml));
    const table = doc.child(0);
    expect(table.child(0).child(0).attrs.borders.top).toEqual({ style: 'solid', widthPx: 4, color: '#00ff00' });
    // 재정의 없는 B1은 여전히 표 기본값(검정 실선)을 쓴다
    expect(table.child(0).child(1).attrs.borders.top).toEqual({ style: 'solid', widthPx: 1, color: '#000000' });
  });

  test('w:val="nil" means no border on that side', () => {
    const tableXml =
      '<w:tbl><w:tblPr><w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/>' +
      '<w:bottom w:val="nil"/><w:right w:val="nil"/></w:tblBorders></w:tblPr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const { doc } = importDocx(docxWithTableXml(tableXml));
    const borders = doc.child(0).child(0).child(0).attrs.borders;
    expect(borders).toEqual({ top: null, right: null, bottom: null, left: null });
  });
});

describe('docx import: manual line breaks (w:br)', () => {
  function docxWithBodyXml(bodyXml: string): Buffer {
    const base = makeSampleDocx([]);
    const entries = unzipSync(base);
    const xml = strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', `${bodyXml}<w:sectPr/>`);
    entries['word/document.xml'] = strToU8(xml);
    return Buffer.from(zipSync(entries));
  }

  test('keeps text on either side of a w:br as separate lines instead of joining them', () => {
    const bodyXml =
      '<w:p><w:r><w:t xml:space="preserve">first line</w:t></w:r>' +
      '<w:r><w:br/></w:r>' +
      '<w:r><w:t xml:space="preserve">second line</w:t></w:r></w:p>';
    const { doc } = importDocx(docxWithBodyXml(bodyXml));
    const paragraph = doc.child(0);
    expect(paragraph.childCount).toBe(3);
    expect(paragraph.child(0).text).toBe('first line');
    expect(paragraph.child(1).type.name).toBe('hard_break');
    expect(paragraph.child(2).text).toBe('second line');
    // textContent는 hard_break 같은 텍스트 아닌 리프 노드를 그냥 건너뛰므로(구분자 없음)
    // 여기선 leafText로 줄바꿈을 대신 넣어주는 textBetween으로 실제 줄 구분을 확인한다
    expect(paragraph.textBetween(0, paragraph.content.size, undefined, '\n')).toBe(
      'first line\nsecond line',
    );
  });
});

describe('docx import: page setup, table alignment, row height, cell vertical align', () => {
  function docxWithSectPr(sectPrXml: string): Buffer {
    const base = makeSampleDocx([]);
    const entries = unzipSync(base);
    const xml = strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', sectPrXml);
    entries['word/document.xml'] = strToU8(xml);
    return Buffer.from(zipSync(entries));
  }

  test('reads page size and margins from w:sectPr (A4 portrait, 1-inch margins)', () => {
    // A4: 11906 x 16838 dxa. 1인치 여백 = 1440dxa.
    const sectPr =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
    const { pageSetup } = importDocx(docxWithSectPr(sectPr));
    expect(pageSetup).not.toBeNull();
    expect(pageSetup!.widthPx).toBe(794); // 11906/15
    expect(pageSetup!.heightPx).toBe(1123); // 16838/15
    expect(pageSetup!.marginTopPx).toBe(96); // 1440/15
    expect(pageSetup!.marginLeftPx).toBe(96);
    expect(pageSetup!.orientation).toBe('portrait');
  });

  test('reads landscape orientation', () => {
    const base = makeSampleDocx([]);
    const entries = unzipSync(base);
    const xml = strFromU8(entries['word/document.xml']).replace(
      '<w:sectPr/>',
      '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>',
    );
    entries['word/document.xml'] = strToU8(xml);
    const { pageSetup } = importDocx(Buffer.from(zipSync(entries)));
    expect(pageSetup!.orientation).toBe('landscape');
    expect(pageSetup!.widthPx).toBe(1123);
  });

  test('returns null pageSetup when w:sectPr has no w:pgSz', () => {
    const { pageSetup } = importDocx(makeSampleDocx(['x']));
    expect(pageSetup).toBeNull();
  });

  test('reads w:tblPr/w:jc as the table-level align attribute (distinct from cell paragraph align)', () => {
    const tableXml =
      '<w:tbl><w:tblPr><w:jc w:val="center"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const base = makeSampleDocx([]);
    const entries = unzipSync(base);
    const xml = strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', `${tableXml}<w:sectPr/>`);
    entries['word/document.xml'] = strToU8(xml);
    const { doc } = importDocx(Buffer.from(zipSync(entries)));
    const table = doc.child(0);
    expect(table.attrs.align).toBe('center'); // 표 자체의 위치
    const cellParagraph = table.child(0).child(0).firstChild!;
    expect(cellParagraph.attrs.align).toBe('right'); // 셀 안 문단의 정렬 — 별개
  });

  test('reads w:trHeight (atLeast) as row height in px, and treats hRule="auto" as unset', () => {
    const tableXml =
      '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:trHeight w:val="720" w:hRule="atLeast"/></w:trPr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:trPr><w:trHeight w:val="720" w:hRule="auto"/></w:trPr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>';
    const base = makeSampleDocx([]);
    const entries = unzipSync(base);
    const xml = strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', `${tableXml}<w:sectPr/>`);
    entries['word/document.xml'] = strToU8(xml);
    const { doc } = importDocx(Buffer.from(zipSync(entries)));
    const table = doc.child(0);
    expect(table.child(0).attrs.height).toEqual({ rule: 'atLeast', px: 48 }); // 720/15
    expect(table.child(1).attrs.height).toBeNull();
  });

  test('reads w:tcPr/w:vAlign as the cell verticalAlign attribute', () => {
    const tableXml =
      '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr>' +
      '<w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>';
    const base = makeSampleDocx([]);
    const entries = unzipSync(base);
    const xml = strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', `${tableXml}<w:sectPr/>`);
    entries['word/document.xml'] = strToU8(xml);
    const { doc } = importDocx(Buffer.from(zipSync(entries)));
    expect(doc.child(0).child(0).child(0).attrs.verticalAlign).toBe('middle');
  });

  const A4_SECTPR =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

  function docxWithTableAndSectPr(tableXml: string, sectPrXml: string = A4_SECTPR): Buffer {
    const base = makeSampleDocx([]);
    const entries = unzipSync(base);
    const xml = strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', `${tableXml}${sectPrXml}`);
    entries['word/document.xml'] = strToU8(xml);
    return Buffer.from(zipSync(entries));
  }

  test('scales column widths down to fit the page content width when w:tblGrid overflows it', () => {
    // A4, 1인치 여백 -> 내용 폭 602px(794-96-96). 표는 그 폭보다 훨씬 넓은 800px(6000dxa*2)로 지정됨.
    const tableXml =
      '<w:tbl><w:tblGrid><w:gridCol w:w="6000"/><w:gridCol w:w="6000"/></w:tblGrid><w:tr>' +
      '<w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>';
    const { doc } = importDocx(docxWithTableAndSectPr(tableXml));
    const row = doc.child(0).child(0);
    // 800px -> 602px로 축소되므로 비율(1:1)을 유지한 채 각각 301px이 되어야 한다.
    expect(row.child(0).attrs.colwidth).toEqual([301]);
    expect(row.child(1).attrs.colwidth).toEqual([301]);
  });

  test('leaves column widths untouched when w:tblGrid already fits within the page content width', () => {
    const tableXml =
      '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>' +
      '<w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>';
    const { doc } = importDocx(docxWithTableAndSectPr(tableXml));
    expect(doc.child(0).child(0).child(0).attrs.colwidth).toEqual([200]); // 3000/15, 좁은 표는 그대로
  });

  test('leaves column widths untouched when there is no page setup to compare against', () => {
    const tableXml =
      '<w:tbl><w:tblGrid><w:gridCol w:w="6000"/><w:gridCol w:w="6000"/></w:tblGrid><w:tr>' +
      '<w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>';
    const { doc } = importDocx(docxWithTableAndSectPr(tableXml, '<w:sectPr/>'));
    const row = doc.child(0).child(0);
    expect(row.child(0).attrs.colwidth).toEqual([400]);
    expect(row.child(1).attrs.colwidth).toEqual([400]);
  });
});

describe('docx import: default font size (w:docDefaults, for documents with no direct/style font size)', () => {
  function docxWithStylesXml(stylesXml: string): Buffer {
    const base = makeSampleDocx(['x']);
    const entries = unzipSync(base);
    entries['word/styles.xml'] = strToU8(stylesXml);
    return Buffer.from(zipSync(entries));
  }

  test('reads w:docDefaults/w:rPrDefault/w:rPr/w:sz as the document default font size (half-points -> pt)', () => {
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '</w:styles>';
    expect(getDocxDefaultFontSizePt(docxWithStylesXml(stylesXml))).toBe(11);
  });

  test('falls back to 10pt when styles.xml has no w:docDefaults sz', () => {
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>';
    expect(getDocxDefaultFontSizePt(docxWithStylesXml(stylesXml))).toBe(10);
  });

  test('prefers the default paragraph style (w:type="paragraph" w:default="1") own w:sz over w:docDefaults', () => {
    // 실제로 관찰된 패턴: 모든 문단이 암묵적으로 이 "기본" 스타일을 쓰고, 그 스타일 자체가
    // docDefaults와 다른 크기를 직접 지정한다 — 이 경우 실제로 적용되는 크기는 스타일 쪽이다.
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="a0"><w:name w:val="Normal"/>' +
      '<w:rPr><w:sz w:val="20"/></w:rPr></w:style>' +
      '</w:styles>';
    expect(getDocxDefaultFontSizePt(docxWithStylesXml(stylesXml))).toBe(10);
  });
});

describe('docx import: default font family (w:docDefaults/w:rFonts, direct or theme-resolved)', () => {
  const THEME1_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>' +
    '<a:fontScheme name="Office">' +
    '<a:majorFont><a:latin typeface="Calibri Light"/><a:font script="Hang" typeface="맑은 고딕 제목"/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:font script="Hang" typeface="맑은 고딕"/></a:minorFont>' +
    '</a:fontScheme></a:themeElements></a:theme>';

  function docxWithStylesAndTheme(rFontsAttrs: string): Buffer {
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts ${rFontsAttrs}/></w:rPr></w:rPrDefault></w:docDefaults>` +
      '</w:styles>';
    const base = makeSampleDocx(['x']);
    const entries = unzipSync(base);
    entries['word/styles.xml'] = strToU8(stylesXml);
    entries['word/theme/theme1.xml'] = strToU8(THEME1_XML);
    return Buffer.from(zipSync(entries));
  }

  test('resolves w:eastAsiaTheme="minorEastAsia" to the theme minorFont Korean (Hang) typeface', () => {
    const docx = docxWithStylesAndTheme('w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"');
    expect(getDocxDefaultFontFamily(docx)).toBe('"맑은 고딕", "Calibri"');
  });

  test('resolves w:eastAsiaTheme="majorEastAsia" to the theme majorFont Korean (Hang) typeface', () => {
    const docx = docxWithStylesAndTheme('w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia"');
    expect(getDocxDefaultFontFamily(docx)).toBe('"맑은 고딕 제목", "Calibri Light"');
  });

  test('prefers a literal w:eastAsia typeface over a theme reference when both are absent/present', () => {
    const docx = docxWithStylesAndTheme('w:ascii="Arial" w:eastAsia="바탕체"');
    expect(getDocxDefaultFontFamily(docx)).toBe('"바탕체", "Arial"');
  });

  test('returns null when styles.xml has no w:docDefaults rFonts at all', () => {
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>';
    const base = makeSampleDocx(['x']);
    const entries = unzipSync(base);
    entries['word/styles.xml'] = strToU8(stylesXml);
    expect(getDocxDefaultFontFamily(Buffer.from(zipSync(entries)))).toBeNull();
  });

  test('prefers the default paragraph style\'s own literal w:rFonts over w:docDefaults (theme or literal)', () => {
    // 실제로 관찰된 패턴: docDefaults는 테마(예: 맑은 고딕)를 가리키지만, 모든 문단이 암묵적으로
    // 쓰는 "기본" 스타일 자체가 리터럴로 다른 서체를 직접 지정한다 — 실제로 적용되는 건 후자다.
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr>' +
      '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/>' +
      '</w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="a0"><w:name w:val="Normal"/>' +
      '<w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="굴림"/></w:rPr></w:style>' +
      '</w:styles>';
    const base = makeSampleDocx(['x']);
    const entries = unzipSync(base);
    entries['word/styles.xml'] = strToU8(stylesXml);
    expect(getDocxDefaultFontFamily(Buffer.from(zipSync(entries)))).toBe('"굴림", "Arial"');
  });
});

describe('docx import: default paragraph alignment inherited from the default paragraph style', () => {
  function docxWithDefaultStyleAndBody(defaultStylePPr: string, bodyXml: string): Buffer {
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      `<w:style w:type="paragraph" w:default="1" w:styleId="a0"><w:name w:val="Normal"/><w:pPr>${defaultStylePPr}</w:pPr></w:style>` +
      '</w:styles>';
    const base = makeSampleDocx([]);
    const entries = unzipSync(base);
    entries['word/styles.xml'] = strToU8(stylesXml);
    const xml = strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', `${bodyXml}<w:sectPr/>`);
    entries['word/document.xml'] = strToU8(xml);
    return Buffer.from(zipSync(entries));
  }

  test('applies the default style\'s w:jc="both" (justify) to a plain paragraph with no direct jc and no pStyle', () => {
    const bodyXml = '<w:p><w:r><w:t xml:space="preserve">plain text</w:t></w:r></w:p>';
    const { doc } = importDocx(docxWithDefaultStyleAndBody('<w:jc w:val="both"/>', bodyXml));
    expect(doc.child(0).attrs.align).toBe('justify');
  });

  test('a paragraph\'s own direct w:jc still wins over the default style', () => {
    const bodyXml =
      '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p>';
    const { doc } = importDocx(docxWithDefaultStyleAndBody('<w:jc w:val="both"/>', bodyXml));
    expect(doc.child(0).attrs.align).toBe('right');
  });

  test('does not apply the default style\'s jc to a paragraph that references some other w:pStyle', () => {
    const bodyXml =
      '<w:p><w:pPr><w:pStyle w:val="SomeOtherStyle"/></w:pPr><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p>';
    const { doc } = importDocx(docxWithDefaultStyleAndBody('<w:jc w:val="both"/>', bodyXml));
    expect(doc.child(0).attrs.align).toBeNull();
  });

  test('applies the default style\'s alignment inside table cells too (e.g. a shaded title cell with no direct jc)', () => {
    const tableXml =
      '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr>' +
      '<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="CCCCCC"/></w:tcPr>' +
      '<w:p><w:r><w:t xml:space="preserve">제목</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const { doc } = importDocx(docxWithDefaultStyleAndBody('<w:jc w:val="both"/>', tableXml));
    const cellParagraph = doc.child(0).child(0).child(0).firstChild!;
    expect(cellParagraph.attrs.align).toBe('justify');
  });
});

describe('docx import: table cell margins (w:tblCellMar/w:tcMar, cascaded like real Word)', () => {
  // 우리 CSS는 모든 셀에 고정된 0.4rem 0.6rem(≈6.4px/9.6px) padding을 쓰는데, 실제 문서가
  // 이보다 좁은 여백(예: 108dxa=7.2px)을 지정하면 그 차이만큼 셀 안 텍스트가 덜 들어가
  // 원본에선 한 줄인 게 편집기에선 두 줄로 줄바꿈되는 원인이 된다.
  function docxWithTablesXml(stylesTblPrXml: string, tableXml: string): Buffer {
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      (stylesTblPrXml
        ? `<w:style w:type="table" w:default="1" w:styleId="a2"><w:name w:val="Normal Table"/><w:tblPr>${stylesTblPrXml}</w:tblPr></w:style>`
        : '') +
      '</w:styles>';
    const base = makeSampleDocx([]);
    const entries = unzipSync(base);
    entries['word/styles.xml'] = strToU8(stylesXml);
    const xml = strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', `${tableXml}<w:sectPr/>`);
    entries['word/document.xml'] = strToU8(xml);
    return Buffer.from(zipSync(entries));
  }

  const SIMPLE_TABLE =
    '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr>' +
    '<w:tc><w:p><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';

  test('falls back to Word\'s standard cell margin (108dxa=7.2px left/right, 0 top/bottom) when nothing specifies one', () => {
    const { doc } = importDocx(docxWithTablesXml('', SIMPLE_TABLE));
    const cell = doc.child(0).child(0).child(0);
    expect(cell.attrs.cellPadding).toEqual({ topPx: 0, rightPx: 7.2, bottomPx: 0, leftPx: 7.2 });
  });

  test('uses the default table style\'s own w:tblCellMar when the table itself does not specify one', () => {
    const stylesTblPr = '<w:tblCellMar><w:top w:w="20" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>' +
      '<w:bottom w:w="20" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>';
    const { doc } = importDocx(docxWithTablesXml(stylesTblPr, SIMPLE_TABLE));
    const cell = doc.child(0).child(0).child(0);
    expect(cell.attrs.cellPadding).toEqual({ topPx: 20 / 15, rightPx: 7.2, bottomPx: 20 / 15, leftPx: 7.2 });
  });

  test('the table\'s own w:tblCellMar overrides the default table style', () => {
    const tableXml =
      '<w:tbl><w:tblPr><w:tblCellMar><w:left w:w="200" w:type="dxa"/><w:right w:w="200" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr>' +
      '<w:tc><w:p><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const { doc } = importDocx(docxWithTablesXml('', tableXml));
    const cell = doc.child(0).child(0).child(0);
    expect(cell.attrs.cellPadding).toEqual({ topPx: 0, rightPx: 200 / 15, bottomPx: 0, leftPx: 200 / 15 });
  });

  test('a cell-level w:tcMar overrides the table default for that cell only', () => {
    const tableXml =
      '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr>' +
      '<w:tc><w:tcPr><w:tcMar><w:left w:w="300" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:t xml:space="preserve">a</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t xml:space="preserve">b</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>';
    const { doc } = importDocx(docxWithTablesXml('', tableXml));
    const row = doc.child(0).child(0);
    expect(row.child(0).attrs.cellPadding).toEqual({ topPx: 0, rightPx: 7.2, bottomPx: 0, leftPx: 300 / 15 });
    expect(row.child(1).attrs.cellPadding).toEqual({ topPx: 0, rightPx: 7.2, bottomPx: 0, leftPx: 7.2 });
  });
});

describe('docx import: paragraphs referencing a named custom style (w:pStyle) pick up that style\'s own jc/bold', () => {
  // 실제로 발견된 패턴: 문단 자신에는 직접 서식이 전혀 없고(런에 w:b 없음, w:pPr에 w:jc 없음)
  // w:pStyle로 사용자 정의 스타일("테이블 제목" 등)만 참조하는데, 그 스타일 자체가
  // jc=center + w:b(굵게)를 지정한다 — 이건 "기본(default) 스타일" 하나가 아니라 이름이 있는
  // 임의의 스타일이라, 지금까지의 defaultParagraphStyleInfoOf(기본 스타일 전용)로는 못 잡는다.
  function docxWithNamedStyle(styleDefXml: string, bodyXml: string): Buffer {
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      styleDefXml +
      '</w:styles>';
    const base = makeSampleDocx([]);
    const entries = unzipSync(base);
    entries['word/styles.xml'] = strToU8(stylesXml);
    const xml = strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', `${bodyXml}<w:sectPr/>`);
    entries['word/document.xml'] = strToU8(xml);
    return Buffer.from(zipSync(entries));
  }

  const TABLE_TITLE_STYLE =
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="ab"><w:name w:val="테이블 제목"/>' +
    '<w:basedOn w:val="a0"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>';

  test('applies the referenced style\'s jc=center and w:b(bold) when the paragraph has no direct formatting of its own', () => {
    const bodyXml =
      '<w:p><w:pPr><w:pStyle w:val="ab"/></w:pPr><w:r><w:t xml:space="preserve">인수인계(확인)서</w:t></w:r></w:p>';
    const { doc } = importDocx(docxWithNamedStyle(TABLE_TITLE_STYLE, bodyXml));
    const paragraph = doc.child(0);
    expect(paragraph.attrs.align).toBe('center');
    const textNode = paragraph.firstChild!;
    expect(textNode.marks.map((m) => m.type.name)).toContain('strong');
  });

  test('applies inside table cells too (the real-world case: a table-title cell referencing this style)', () => {
    const tableXml =
      '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr>' +
      '<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="CCCCCC"/></w:tcPr>' +
      '<w:p><w:pPr><w:pStyle w:val="ab"/></w:pPr><w:r><w:t xml:space="preserve">인수인계(확인)서</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>';
    const { doc } = importDocx(docxWithNamedStyle(TABLE_TITLE_STYLE, tableXml));
    const cellParagraph = doc.child(0).child(0).child(0).firstChild!;
    expect(cellParagraph.attrs.align).toBe('center');
    expect(cellParagraph.firstChild!.marks.map((m) => m.type.name)).toContain('strong');
  });

  test('a paragraph\'s own direct w:jc still wins over the referenced style\'s jc', () => {
    const bodyXml =
      '<w:p><w:pPr><w:pStyle w:val="ab"/><w:jc w:val="right"/></w:pPr><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p>';
    const { doc } = importDocx(docxWithNamedStyle(TABLE_TITLE_STYLE, bodyXml));
    expect(doc.child(0).attrs.align).toBe('right');
  });

  test('a run\'s own explicit w:b still applies even if the style did not specify bold (no double-negative surprises)', () => {
    const styleWithoutBold =
      '<w:style w:type="paragraph" w:customStyle="1" w:styleId="a"><w:name w:val="테이블 내용구분"/>' +
      '<w:basedOn w:val="a0"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>';
    const bodyXml =
      '<w:p><w:pPr><w:pStyle w:val="a"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">x</w:t></w:r></w:p>';
    const { doc } = importDocx(docxWithNamedStyle(styleWithoutBold, bodyXml));
    const paragraph = doc.child(0);
    expect(paragraph.attrs.align).toBeNull();
    expect(paragraph.firstChild!.marks.map((m) => m.type.name)).toContain('strong');
  });

  test('a style with no jc/bold of its own (e.g. just numbering) does not force either on the paragraph', () => {
    const styleWithoutBold =
      '<w:style w:type="paragraph" w:customStyle="1" w:styleId="a"><w:name w:val="테이블 내용구분"/>' +
      '<w:basedOn w:val="a0"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>';
    const bodyXml = '<w:p><w:pPr><w:pStyle w:val="a"/></w:pPr><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p>';
    const { doc } = importDocx(docxWithNamedStyle(styleWithoutBold, bodyXml));
    const paragraph = doc.child(0);
    expect(paragraph.attrs.align).toBeNull();
    expect(paragraph.firstChild!.marks.map((m) => m.type.name)).not.toContain('strong');
  });
});

describe('docx export', () => {
  test('round-trips formatted paragraphs, headings and tables through a full docx cycle', () => {
    const original = makeSampleDocx(['placeholder']);

    const heading = docxSchema.nodes.heading.create({ level: 2 }, docxSchema.text('Report Title'));
    const boldItalic = docxSchema.text('bold and italic', [
      docxSchema.marks.strong.create(),
      docxSchema.marks.em.create(),
    ]);
    const paragraph = docxSchema.nodes.paragraph.create(null, boldItalic);
    const cell = (text: string) =>
      docxSchema.nodes.table_cell.create(null, docxSchema.nodes.paragraph.create(null, docxSchema.text(text)));
    const row = docxSchema.nodes.table_row.create(null, [cell('A1'), cell('B1')]);
    const table = docxSchema.nodes.table.create(null, [row]);

    const doc = docxSchema.nodes.doc.create(null, [heading, paragraph, table]);

    const xml = renderDocumentXml(doc);
    expect(xml).toContain('Heading2');
    expect(xml).toContain('<w:b/>');
    expect(xml).toContain('<w:i/>');
    expect(xml).toContain('<w:tbl>');

    const exported = exportDocx(original, doc);
    // 재압축 결과가 여전히 유효한 zip이며 다른 파트(styles.xml)는 그대로 보존됐는지 확인한다
    const entries = unzipSync(exported);
    expect(entries['word/styles.xml']).toBeDefined();
    expect(strFromU8(entries['word/document.xml'])).toBe(xml);

    const reimported = importDocx(exported);
    expect(reimported.doc.child(0).type.name).toBe('heading');
    expect(reimported.doc.child(0).textContent).toBe('Report Title');
    expect(reimported.doc.child(1).textContent).toBe('bold and italic');
    const marks = reimported.doc.child(1).firstChild?.marks.map((m) => m.type.name) ?? [];
    expect(marks).toContain('strong');
    expect(marks).toContain('em');
    expect(reimported.doc.child(2).type.name).toBe('table');
  });

  test('round-trips a horizontally and vertically merged table through a full docx cycle', () => {
    const original = makeSampleDocx(['placeholder']);
    const cell = (text: string, attrs: { colspan?: number; rowspan?: number } = {}) =>
      docxSchema.nodes.table_cell.create(
        { colspan: attrs.colspan ?? 1, rowspan: attrs.rowspan ?? 1 },
        docxSchema.nodes.paragraph.create(null, docxSchema.text(text)),
      );
    // 1행: 가로로 2칸 병합된 셀 하나
    const row1 = docxSchema.nodes.table_row.create(null, [cell('Merged', { colspan: 2 })]);
    // 2행~3행: 1열이 세로로 2칸 병합, 2열은 각 행마다 별개 셀
    const row2 = docxSchema.nodes.table_row.create(null, [cell('Tall', { rowspan: 2 }), cell('B2')]);
    const row3 = docxSchema.nodes.table_row.create(null, [cell('B3')]);
    const table = docxSchema.nodes.table.create(null, [row1, row2, row3]);
    const doc = docxSchema.nodes.doc.create(null, [table]);

    const xml = renderDocumentXml(doc);
    expect(xml).toContain('<w:gridSpan w:val="2"/>');
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
    expect(xml).toContain('<w:vMerge/>');

    const exported = exportDocx(original, doc);
    const reimported = importDocx(exported);
    const reimportedTable = reimported.doc.child(0);
    const reimportedRow1 = reimportedTable.child(0);
    const reimportedRow2 = reimportedTable.child(1);
    const reimportedRow3 = reimportedTable.child(2);

    expect(reimportedRow1.childCount).toBe(1);
    expect(reimportedRow1.child(0).attrs.colspan).toBe(2);
    expect(reimportedRow1.child(0).textContent).toBe('Merged');

    expect(reimportedRow2.childCount).toBe(2);
    expect(reimportedRow2.child(0).attrs.rowspan).toBe(2);
    expect(reimportedRow2.child(0).textContent).toBe('Tall');
    expect(reimportedRow2.child(1).textContent).toBe('B2');

    // 세로 병합이 이어지는 3행에는 원래 셀(B3)만 남고, 병합된 1열 자리는 별개 셀로 다시 생기지 않는다
    expect(reimportedRow3.childCount).toBe(1);
    expect(reimportedRow3.child(0).textContent).toBe('B3');
  });

  test('round-trips cell column width and background shading through a full docx cycle', () => {
    const original = makeSampleDocx(['placeholder']);
    const cell = (text: string, attrs: { colwidth?: number[]; background?: string | null } = {}) =>
      docxSchema.nodes.table_cell.create(
        { colspan: 1, rowspan: 1, colwidth: attrs.colwidth ?? null, background: attrs.background ?? null },
        docxSchema.nodes.paragraph.create(null, docxSchema.text(text)),
      );
    const row = docxSchema.nodes.table_row.create(null, [
      cell('Narrow', { colwidth: [96] }),
      cell('Shaded', { background: '#ffd43b' }),
    ]);
    const table = docxSchema.nodes.table.create(null, [row]);
    const doc = docxSchema.nodes.doc.create(null, [table]);

    const xml = renderDocumentXml(doc);
    expect(xml).toContain('<w:gridCol w:w="1440"/>'); // 96px -> 1440dxa
    expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FFD43B"/>');

    const exported = exportDocx(original, doc);
    const reimported = importDocx(exported);
    const reimportedRow = reimported.doc.child(0).child(0);
    expect(reimportedRow.child(0).attrs.colwidth).toEqual([96]);
    expect(reimportedRow.child(1).attrs.background).toBe('#ffd43b');
  });

  test('round-trips a manual line break (hard_break) inside a paragraph', () => {
    const original = makeSampleDocx(['placeholder']);
    const paragraph = docxSchema.nodes.paragraph.create(null, [
      docxSchema.text('first line'),
      docxSchema.nodes.hard_break.create(),
      docxSchema.text('second line'),
    ]);
    const doc = docxSchema.nodes.doc.create(null, [paragraph]);

    const xml = renderDocumentXml(doc);
    expect(xml).toContain('<w:br/>');

    const exported = exportDocx(original, doc);
    const reimported = importDocx(exported);
    const reimportedParagraph = reimported.doc.child(0);
    expect(reimportedParagraph.childCount).toBe(3);
    expect(reimportedParagraph.child(1).type.name).toBe('hard_break');
    expect(
      reimportedParagraph.textBetween(0, reimportedParagraph.content.size, undefined, '\n'),
    ).toBe('first line\nsecond line');
  });

  test('round-trips per-cell borders, including a side with no border (nil)', () => {
    const original = makeSampleDocx(['placeholder']);
    const cell = docxSchema.nodes.table_cell.create(
      {
        colspan: 1,
        rowspan: 1,
        borders: {
          top: { style: 'solid', widthPx: 1, color: '#000000' },
          right: { style: 'dashed', widthPx: 2, color: '#ff0000' },
          bottom: null,
          left: { style: 'double', widthPx: 1.5, color: '#0000ff' },
        },
      },
      docxSchema.nodes.paragraph.create(null, docxSchema.text('X')),
    );
    const table = docxSchema.nodes.table.create(null, [
      docxSchema.nodes.table_row.create(null, [cell]),
    ]);
    const doc = docxSchema.nodes.doc.create(null, [table]);

    const xml = renderDocumentXml(doc);
    expect(xml).toContain('<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>');
    expect(xml).toContain('<w:right w:val="dashed" w:sz="12" w:space="0" w:color="FF0000"/>');
    expect(xml).toContain('<w:bottom w:val="nil"/>');
    expect(xml).toContain('<w:left w:val="double" w:sz="9" w:space="0" w:color="0000FF"/>');

    const reimported = importDocx(exportDocx(original, doc));
    const reimportedBorders = reimported.doc.child(0).child(0).child(0).attrs.borders;
    expect(reimportedBorders.top).toEqual({ style: 'solid', widthPx: 1, color: '#000000' });
    expect(reimportedBorders.right).toEqual({ style: 'dashed', widthPx: 2, color: '#ff0000' });
    expect(reimportedBorders.bottom).toBeNull();
    expect(reimportedBorders.left).toEqual({ style: 'double', widthPx: 1.5, color: '#0000ff' });
  });

  test('round-trips a composite table: merges + shading + per-column widths + borders together', () => {
    const original = makeSampleDocx(['placeholder']);
    const redBorder = { style: 'solid', widthPx: 1, color: '#ff0000' };
    const borders = (over: Partial<Record<'top' | 'right' | 'bottom' | 'left', unknown>> = {}) => ({
      top: redBorder,
      right: redBorder,
      bottom: redBorder,
      left: redBorder,
      ...over,
    });
    const cell = (
      text: string,
      attrs: { colspan?: number; rowspan?: number; colwidth?: number[]; background?: string | null } = {},
    ) =>
      docxSchema.nodes.table_cell.create(
        {
          colspan: attrs.colspan ?? 1,
          rowspan: attrs.rowspan ?? 1,
          colwidth: attrs.colwidth ?? null,
          background: attrs.background ?? null,
          borders: borders(),
        },
        docxSchema.nodes.paragraph.create(null, text === '' ? null : docxSchema.text(text)),
      );
    // 1행: 가로 병합 + 음영, 2열 폭 지정
    const row1 = docxSchema.nodes.table_row.create(null, [
      cell('제목(병합+음영)', { colspan: 2, colwidth: [96, 192], background: '#ffd43b' }),
    ]);
    // 2~3행: 1열이 세로 병합, 2열은 빈 셀 포함
    const row2 = docxSchema.nodes.table_row.create(null, [
      cell('세로병합', { rowspan: 2, colwidth: [96] }),
      cell('', { colwidth: [192] }), // 빈 셀
    ]);
    const row3 = docxSchema.nodes.table_row.create(null, [cell('B3', { colwidth: [192] })]);
    const table = docxSchema.nodes.table.create(null, [row1, row2, row3]);
    const doc = docxSchema.nodes.doc.create(null, [table]);

    const exported = exportDocx(original, doc);
    const reimported = importDocx(exported);
    const t = reimported.doc.child(0);

    const r1c0 = t.child(0).child(0);
    expect(r1c0.attrs.colspan).toBe(2);
    expect(r1c0.attrs.background).toBe('#ffd43b');
    expect(r1c0.attrs.colwidth).toEqual([96, 192]);
    expect(r1c0.attrs.borders.top).toEqual(redBorder);

    const r2c0 = t.child(1).child(0);
    expect(r2c0.attrs.rowspan).toBe(2);
    expect(r2c0.textContent).toBe('세로병합');
    expect(t.child(1).child(1).textContent).toBe(''); // 빈 셀 그대로 보존

    expect(t.child(2).childCount).toBe(1); // 세로 병합 이어짐 칸은 별개 셀로 안 생긴다
    expect(t.child(2).child(0).textContent).toBe('B3');
  });

  test('preserves the original w:sectPr (page size/margins) across an edit that never touches it', () => {
    // makeSampleDocx는 빈 <w:sectPr/>만 만들어주므로, 실제 페이지 크기가 있는 원본을 직접 조립한다.
    const entries = unzipSync(makeSampleDocx(['placeholder']));
    const realSectPr =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>';
    entries['word/document.xml'] = strToU8(
      strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', realSectPr),
    );
    const original = Buffer.from(zipSync(entries));

    const doc = docxSchema.nodes.doc.create(null, [
      docxSchema.nodes.paragraph.create(null, docxSchema.text('edited text')),
    ]);
    const exported = exportDocx(original, doc);
    const exportedXml = strFromU8(unzipSync(exported)['word/document.xml']);
    expect(exportedXml).toContain(realSectPr);

    const reimported = importDocx(exported);
    expect(reimported.pageSetup).toEqual({
      widthPx: 794,
      heightPx: 1123,
      marginTopPx: 96,
      marginBottomPx: 96,
      marginLeftPx: 120,
      marginRightPx: 120,
      orientation: 'portrait',
    });
  });

  test('round-trips table-level alignment (w:tblPr/w:jc) and row height (w:trHeight)', () => {
    const original = makeSampleDocx(['placeholder']);
    const cell = docxSchema.nodes.table_cell.create(
      { colwidth: [100] },
      docxSchema.nodes.paragraph.create(null, docxSchema.text('x')),
    );
    const row = docxSchema.nodes.table_row.create({ height: { rule: 'exact', px: 40 } }, [cell]);
    const table = docxSchema.nodes.table.create({ align: 'center' }, [row]);
    const doc = docxSchema.nodes.doc.create(null, [table]);

    const xml = renderDocumentXml(doc);
    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml).toContain(`<w:trHeight w:val="${40 * 15}" w:hRule="exact"/>`);

    const reimported = importDocx(exportDocx(original, doc));
    const reimportedTable = reimported.doc.child(0);
    expect(reimportedTable.attrs.align).toBe('center');
    expect(reimportedTable.child(0).attrs.height).toEqual({ rule: 'exact', px: 40 });
  });

  test('round-trips cell vertical alignment (w:tcPr/w:vAlign)', () => {
    const original = makeSampleDocx(['placeholder']);
    const cell = docxSchema.nodes.table_cell.create(
      { colwidth: [100], verticalAlign: 'middle' },
      docxSchema.nodes.paragraph.create(null, docxSchema.text('x')),
    );
    const row = docxSchema.nodes.table_row.create(null, [cell]);
    const table = docxSchema.nodes.table.create(null, [row]);
    const doc = docxSchema.nodes.doc.create(null, [table]);

    const xml = renderDocumentXml(doc);
    expect(xml).toContain('<w:vAlign w:val="center"/>');

    const reimported = importDocx(exportDocx(original, doc));
    const reimportedCell = reimported.doc.child(0).child(0).child(0);
    expect(reimportedCell.attrs.verticalAlign).toBe('middle');
  });

  test('round-trips cell margins (w:tcPr/w:tcMar) read from the original OOXML', () => {
    const original = makeSampleDocx(['placeholder']);
    const cell = docxSchema.nodes.table_cell.create(
      { colwidth: [100], cellPadding: { topPx: 1, rightPx: 12, bottomPx: 1, leftPx: 12 } },
      docxSchema.nodes.paragraph.create(null, docxSchema.text('x')),
    );
    const row = docxSchema.nodes.table_row.create(null, [cell]);
    const table = docxSchema.nodes.table.create(null, [row]);
    const doc = docxSchema.nodes.doc.create(null, [table]);

    const xml = renderDocumentXml(doc);
    expect(xml).toContain('<w:tcMar>');
    expect(xml).toContain(`<w:right w:w="${12 * 15}" w:type="dxa"/>`);

    const reimported = importDocx(exportDocx(original, doc));
    const reimportedCell = reimported.doc.child(0).child(0).child(0);
    expect(reimportedCell.attrs.cellPadding).toEqual({ topPx: 1, rightPx: 12, bottomPx: 1, leftPx: 12 });
  });

  test('round-trips font color, font size and paragraph alignment', () => {
    const original = makeSampleDocx();
    const colored = docxSchema.text('red bold text', [
      docxSchema.marks.textColor.create({ color: '#ff0000' }),
      docxSchema.marks.fontSize.create({ size: 18 }),
    ]);
    const paragraph = docxSchema.nodes.paragraph.create({ align: 'center' }, colored);
    const doc = docxSchema.nodes.doc.create(null, [paragraph]);

    const xml = renderDocumentXml(doc);
    expect(xml).toContain('<w:color w:val="FF0000"/>');
    expect(xml).toContain('<w:sz w:val="36"/>'); // 18pt -> 36 half-points
    expect(xml).toContain('<w:jc w:val="center"/>');

    const exported = exportDocx(original, doc);
    const reimported = importDocx(exported);
    const reimportedParagraph = reimported.doc.child(0);
    expect(reimportedParagraph.attrs.align).toBe('center');
    const textNode = reimportedParagraph.firstChild!;
    const markNames = textNode.marks.map((m) => m.type.name);
    expect(markNames).toContain('textColor');
    expect(markNames).toContain('fontSize');
    expect(textNode.marks.find((m) => m.type.name === 'textColor')?.attrs.color).toBe('#ff0000');
    expect(textNode.marks.find((m) => m.type.name === 'fontSize')?.attrs.size).toBe(18);
  });

  test('flattens bullet lists into prefixed paragraphs (v1 known limitation)', () => {
    const original = makeSampleDocx();
    const item = (text: string) =>
      docxSchema.nodes.list_item.create(null, docxSchema.nodes.paragraph.create(null, docxSchema.text(text)));
    const list = docxSchema.nodes.bullet_list.create(null, [item('first'), item('second')]);
    const doc = docxSchema.nodes.doc.create(null, [list]);

    const xml = renderDocumentXml(doc);
    expect(xml).toContain('• ');
    expect(xml).not.toContain('<w:numPr>');

    const exported = exportDocx(original, doc);
    const reimported = importDocx(exported);
    // 목록 서식은 유실되지만 텍스트(접두사 포함)는 보존된다
    expect(reimported.doc.textBetween(0, reimported.doc.content.size, '\n')).toContain('• first');
  });
});
