import { describe, expect, test } from 'vitest';
import { appendNewShape, parseSlideShapes, serializeTxBodyXml, spliceShapeTxBody } from '../src/pptx/slideContent.js';
import { pptxTextSchema as schema } from '../src/pptx/textSchema.js';

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

describe('serializeTxBodyXml + spliceShapeTxBody (export)', () => {
  test('round-trips plain text unchanged', () => {
    const original =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>' +
      XML_TAIL;

    const [shape] = parseSlideShapes(original);
    const txBodyXml = serializeTxBodyXml(shape.doc);
    const spliced = spliceShapeTxBody(original, '2', txBodyXml);

    const [reparsed] = parseSlideShapes(spliced);
    expect(reparsed.doc.textContent).toBe('Hello');
  });

  test('round-trips mixed bold/plain runs and does not touch the other shape', () => {
    const original =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="A"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:rPr b="1"/><a:t>Bold</a:t></a:r><a:r><a:t> plain</a:t></a:r></a:p></p:txBody></p:sp>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="B"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Untouched</a:t></a:r></a:p></p:txBody></p:sp>' +
      XML_TAIL;

    const shapes = parseSlideShapes(original);
    const editedDoc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('Edited ', [schema.marks.strong.create()]),
        schema.text('text'),
      ]),
    ]);
    const spliced = spliceShapeTxBody(original, '2', serializeTxBodyXml(editedDoc));

    const reparsed = parseSlideShapes(spliced);
    const shapeA = reparsed.find((s) => s.shapeId === '2')!;
    const shapeB = reparsed.find((s) => s.shapeId === '3')!;
    expect(shapeA.doc.child(0).child(0).text).toBe('Edited ');
    expect(shapeA.doc.child(0).child(0).marks.map((m) => m.type.name)).toEqual(['strong']);
    expect(shapeB.doc.textContent).toBe('Untouched'); // 다른 도형은 안 건드림
  });

  test('color and font size survive the round trip', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('Red big', [
          schema.marks.textColor.create({ color: '#00FF00' }),
          schema.marks.fontSize.create({ size: 18 }),
        ]),
      ]),
    ]);
    const original =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="9" name="X"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>placeholder</a:t></a:r></a:p></p:txBody></p:sp>' +
      XML_TAIL;

    const spliced = spliceShapeTxBody(original, '9', serializeTxBodyXml(doc));
    const [reparsed] = parseSlideShapes(spliced);
    const run = reparsed.doc.child(0).child(0);
    expect(run.marks.find((m) => m.type.name === 'textColor')?.attrs.color).toBe('#00FF00');
    expect(run.marks.find((m) => m.type.name === 'fontSize')?.attrs.size).toBe(18);
  });

  test('appendNewShape inserts a brand-new text box with its own position', () => {
    const original = XML_HEAD + XML_TAIL; // 텍스트박스가 하나도 없는 슬라이드
    const doc = schema.node('doc', null, [schema.node('paragraph', null, [schema.text('New box')])]);

    const withShape = appendNewShape(
      original,
      { shapeId: 'new-abc123', x: 1000000, y: 2000000, w: 3000000, h: 500000 },
      serializeTxBodyXml(doc),
    );

    const shapes = parseSlideShapes(withShape);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].shapeId).toBe('new-abc123');
    expect(shapes[0].x).toBe(1000000);
    expect(shapes[0].doc.textContent).toBe('New box');
  });

  test('throws instead of silently no-op when the matched <p:sp> block has no <p:txBody>', () => {
    // id는 일치하지만 <p:txBody>가 구조적으로 없는 <p:sp> 블록(그림 등)을 손으로 만든다 —
    // parseSlideShapes는 이런 도형을 애초에 ParsedShape로 만들지 않지만, spliceShapeTxBody는
    // 원시 XML 문자열을 그대로 받으므로 이런 입력도 걸러내지 못하면 안 된다.
    const original =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="42" name="NoBody"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr/></p:sp>' +
      XML_TAIL;

    expect(() => spliceShapeTxBody(original, '42', '<p:txBody><a:bodyPr/><a:p/></p:txBody>')).toThrow();
  });
});
