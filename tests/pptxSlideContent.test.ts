import { describe, expect, test } from 'vitest';
import { parseSlideShapes } from '../src/pptx/slideContent.js';

const XML_HEAD =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/>';
const XML_TAIL = '</p:spTree></p:cSld></p:sld>';

describe('parseSlideShapes', () => {
  test('parses a single plain-text shape with explicit position', () => {
    const xml =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="838200" y="365125"/><a:ext cx="7772400" cy="1470025"/></a:xfrm></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>' +
      XML_TAIL;

    const shapes = parseSlideShapes(xml);

    expect(shapes).toHaveLength(1);
    expect(shapes[0].shapeId).toBe('2');
    expect(shapes[0].x).toBe(838200);
    expect(shapes[0].y).toBe(365125);
    expect(shapes[0].w).toBe(7772400);
    expect(shapes[0].h).toBe(1470025);
    expect(shapes[0].doc.textContent).toBe('Hello');
  });

  test('position is null when <a:xfrm> is absent (layout-inherited)', () => {
    const xml =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>' +
      XML_TAIL;

    const shapes = parseSlideShapes(xml);

    expect(shapes[0].x).toBeNull();
    expect(shapes[0].y).toBeNull();
  });

  test('preserves mixed formatting within one paragraph (bold run next to plain run)', () => {
    const xml =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/><a:p>' +
      '<a:r><a:rPr b="1"/><a:t>Bold</a:t></a:r>' +
      '<a:r><a:t> plain</a:t></a:r>' +
      '</a:p></p:txBody></p:sp>' +
      XML_TAIL;

    const shapes = parseSlideShapes(xml);
    const para = shapes[0].doc.child(0);
    expect(para.child(0).text).toBe('Bold');
    expect(para.child(0).marks.map((m) => m.type.name)).toEqual(['strong']);
    expect(para.child(1).text).toBe(' plain');
    expect(para.child(1).marks).toHaveLength(0);
  });

  test('reads color and font size from <a:rPr>', () => {
    const xml =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r>' +
      '<a:rPr sz="2400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr>' +
      '<a:t>Red big</a:t></a:r></a:p></p:txBody></p:sp>' +
      XML_TAIL;

    const shapes = parseSlideShapes(xml);
    const run = shapes[0].doc.child(0).child(0);
    const color = run.marks.find((m) => m.type.name === 'textColor');
    const size = run.marks.find((m) => m.type.name === 'fontSize');
    expect(color?.attrs.color).toBe('#FF0000');
    expect(size?.attrs.size).toBe(24); // 2400 hundredths-of-a-point -> 24pt
  });

  test('multiple paragraphs become separate paragraph nodes', () => {
    const xml =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="5" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/>' +
      '<a:p><a:r><a:t>First</a:t></a:r></a:p>' +
      '<a:p><a:r><a:t>Second</a:t></a:r></a:p>' +
      '</p:txBody></p:sp>' +
      XML_TAIL;

    const shapes = parseSlideShapes(xml);
    expect(shapes[0].doc.childCount).toBe(2);
    expect(shapes[0].doc.child(0).textContent).toBe('First');
    expect(shapes[0].doc.child(1).textContent).toBe('Second');
  });

  test('shapes without <p:txBody> (images, plain shapes) are skipped', () => {
    const xml =
      XML_HEAD +
      '<p:pic><p:nvPicPr><p:cNvPr id="6" name="Picture"/></p:nvPicPr></p:pic>' +
      XML_TAIL;

    expect(parseSlideShapes(xml)).toHaveLength(0);
  });

  test('an empty paragraph (no runs) still produces a paragraph node with empty text', () => {
    const xml =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="7" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>' +
      XML_TAIL;

    const shapes = parseSlideShapes(xml);
    expect(shapes[0].doc.childCount).toBe(1);
    expect(shapes[0].doc.child(0).textContent).toBe('');
  });
});
