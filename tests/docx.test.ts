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
