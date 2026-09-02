import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import { exportDocx, renderDocumentXml } from '../src/docx/export.js';
import { importDocx } from '../src/docx/import.js';
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
