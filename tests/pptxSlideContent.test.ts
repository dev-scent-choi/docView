import { describe, expect, test } from 'vitest';
import {
  appendNewShape,
  hasShape,
  parseSlideShapes,
  removeShape,
  serializeTxBodyXml,
  spliceShapeTxBody,
} from '../src/pptx/slideContent.js';
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

  test('parses a shape whose opening tag carries an attribute (e.g. useBgFill, as PowerPoint emits for background-fill shapes)', () => {
    // fast-xml-parser 자체는 항상 실제 파서이므로 이 케이스에서 문제가 된 적이 없다 — 이 테스트는
    // 회귀 방지용 대조군: 아래 hasShape/spliceShapeTxBody/removeShape의 정규식-기반 탐지가
    // parseSlideShapes와 항상 같은 도형 집합을 봐야 한다는 전제를 명시적으로 남겨 둔다.
    const xml =
      XML_HEAD +
      '<p:sp useBgFill="1"><p:nvSpPr><p:cNvPr id="8" name="BgShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Attributed</a:t></a:r></a:p></p:txBody></p:sp>' +
      XML_TAIL;

    const shapes = parseSlideShapes(xml);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].shapeId).toBe('8');
    expect(shapes[0].doc.textContent).toBe('Attributed');
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

  // Task 13 fix 회귀 테스트: hasShape/spliceShapeTxBody가 공유하던 <p:sp> 블록 탐지 정규식이
  // 리터럴 <p:sp>만 매치하고, 여는 태그에 속성이 붙은 형태(PowerPoint가 배경 채우기 도형에 실제로
  // 흔히 붙이는 useBgFill="1" 등)는 매치하지 못했다. hasShape가 그런 도형을 "없다"고 잘못 판단하면
  // src/pptx/collab.ts의 export 병합 루프가 spliceShapeTxBody(교체) 대신 appendNewShape(추가)를
  // 타 버려, 같은 shapeId 아래 옛 텍스트와 새 텍스트가 동시에 남는 조용한 중복이 생겼다(Task 6이
  // 이미 한 번 고친 것과 같은 종류의 버그가 속성 있는 도형에서만 재발한 것).
  test('hasShape finds a shape whose opening tag carries an attribute (regression: previously reported false, causing appendNewShape to duplicate it instead of spliceShapeTxBody updating in place)', () => {
    const original =
      XML_HEAD +
      '<p:sp useBgFill="1"><p:nvSpPr><p:cNvPr id="8" name="BgShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Original</a:t></a:r></a:p></p:txBody></p:sp>' +
      XML_TAIL;

    expect(hasShape(original, '8')).toBe(true);
    expect(hasShape(original, '999')).toBe(false);
  });

  test('spliceShapeTxBody updates an attributed <p:sp useBgFill="1"> shape in place — exactly one shape with that id, holding the new text, not a duplicate', () => {
    const original =
      XML_HEAD +
      '<p:sp useBgFill="1"><p:nvSpPr><p:cNvPr id="8" name="BgShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Old Text</a:t></a:r></a:p></p:txBody></p:sp>' +
      '<p:sp flipH="1"><p:nvSpPr><p:cNvPr id="9" name="Sibling"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Sibling Untouched</a:t></a:r></a:p></p:txBody></p:sp>' +
      XML_TAIL;

    const newDoc = schema.node('doc', null, [schema.node('paragraph', null, [schema.text('New Text')])]);
    const result = spliceShapeTxBody(original, '8', serializeTxBodyXml(newDoc));

    const shapes = parseSlideShapes(result);
    const edited = shapes.filter((s) => s.shapeId === '8');
    expect(edited).toHaveLength(1); // 정확히 하나 — appendNewShape로 중복 추가되지 않았다
    expect(edited[0].doc.textContent).toBe('New Text');
    expect(result).not.toContain('Old Text');

    // 다른 속성 있는(flipH="1") 도형은 전혀 건드리지 않는다 — 그 쪽도 <p:sp\b 넓힌 정규식이 잘못
    // 건드리지 않는지 함께 확인한다.
    const sibling = shapes.find((s) => s.shapeId === '9')!;
    expect(sibling.doc.textContent).toBe('Sibling Untouched');
  });
});

describe('removeShape', () => {
  test('removes the matching <p:sp> block entirely, leaving other shapes untouched', () => {
    const original =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="A"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Keep me</a:t></a:r></a:p></p:txBody></p:sp>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="B"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Delete me</a:t></a:r></a:p></p:txBody></p:sp>' +
      XML_TAIL;

    const result = removeShape(original, '3');

    const shapes = parseSlideShapes(result);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].shapeId).toBe('2');
    expect(shapes[0].doc.textContent).toBe('Keep me');
    expect(result).not.toContain('Delete me');
  });

  test('works on a new--prefixed shapeId (a client-added text box that was later saved then deleted)', () => {
    const original = XML_HEAD + XML_TAIL;
    const doc = schema.node('doc', null, [schema.node('paragraph', null, [schema.text('New')])]);
    const withShape = appendNewShape(
      original,
      { shapeId: 'new-abc', x: 0, y: 0, w: 1000000, h: 500000 },
      serializeTxBodyXml(doc),
    );

    const result = removeShape(withShape, 'new-abc');

    expect(parseSlideShapes(result)).toHaveLength(0);
  });

  test('throws when the shapeId does not exist in the slide XML', () => {
    const original = XML_HEAD + XML_TAIL;
    expect(() => removeShape(original, '999')).toThrow();
  });

  // Task 13 fix 회귀 테스트(Critical, 코드 리뷰에서 발견): removeShape가 쓰던 <p:sp> 블록 탐지
  // 정규식이 리터럴 <p:sp>만 매치하고, 여는 태그에 속성이 붙은 <p:sp useBgFill="1">(PowerPoint가
  // 배경 채우기 도형에 실제로 흔히 붙이는 속성) 같은 형태는 매치하지 못했다. parseSlideShapes(실제
  // XML 파서)는 이런 도형을 정확히 baseline에 포함시키므로, exportPptxFromSlides의 삭제 병합 단계가
  // 그 shapeId로 removeShape를 호출하면 블록이 실제로 존재하는데도 못 찾아 throw했다 — 이 throw는
  // 저장/다운로드 경로 안에서 일어나므로 그 문서의 이후 모든 저장을 영구히 막았다(사용자 쪽에서
  // 되돌릴 방법이 없는 orphan baseline id). 지금은 <p:sp\b로 넓혀 이 케이스가 정상적으로 제거된다.
  test('removes an attributed <p:sp useBgFill="1"> shape without throwing, leaving the sibling shape untouched (Task 13 fix — Critical)', () => {
    const original =
      XML_HEAD +
      '<p:sp useBgFill="1"><p:nvSpPr><p:cNvPr id="8" name="BgShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Delete Me</a:t></a:r></a:p></p:txBody></p:sp>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="9" name="Sibling"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Keep Me</a:t></a:r></a:p></p:txBody></p:sp>' +
      XML_TAIL;

    const result = removeShape(original, '8');

    const shapes = parseSlideShapes(result);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].shapeId).toBe('9');
    expect(shapes[0].doc.textContent).toBe('Keep Me');
    expect(result).not.toContain('Delete Me');
    expect(result).not.toContain('useBgFill');
  });
});

describe('<p:sp\\b block regex: boundary correctness against <p:spPr>/<p:spTree>', () => {
  // Task 13 fix가 <p:sp>를 <p:sp\b로 넓히면서 가장 걱정했던 회귀: <p:spPr>(도형의 위치/스타일
  // 하위 요소)와 <p:spTree>(슬라이드 최상위 도형 컨테이너, 모든 픽스처에 이미 등장한다)를 <p:sp>
  // 블록의 시작으로 잘못 인식하면 안 된다 — 'sp' 바로 다음이 'P'/'T'처럼 단어 문자로 이어지는
  // 위치에는 \b가 성립하지 않으므로 올바르게 거부되어야 한다. 도형 하나에 <p:spPr>가 중첩돼 있고
  // (실제 pptx의 전형적인 모양), 전체가 <p:spTree> 안에 있는(모든 XML_HEAD/XML_TAIL 픽스처가 이미
  // 그렇다) 상태에서 hasShape/spliceShapeTxBody/removeShape가 정확히 도형 하나만 인식하는지 본다.
  const oneShapeWithSpPrAndSpTree =
    XML_HEAD +
    '<p:sp useBgFill="1"><p:nvSpPr><p:cNvPr id="8" name="BgShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Only Shape</a:t></a:r></a:p></p:txBody></p:sp>' +
    XML_TAIL;

  test('hasShape finds exactly the real shape and does not treat the nested <p:spPr> or the wrapping <p:spTree> as a shape', () => {
    expect(hasShape(oneShapeWithSpPrAndSpTree, '8')).toBe(true);
    expect(parseSlideShapes(oneShapeWithSpPrAndSpTree)).toHaveLength(1);
  });

  test('spliceShapeTxBody replaces only the real shape\'s <p:txBody>, leaving <p:spPr>/<p:spTree> structure intact', () => {
    const newDoc = schema.node('doc', null, [schema.node('paragraph', null, [schema.text('Updated')])]);
    const result = spliceShapeTxBody(oneShapeWithSpPrAndSpTree, '8', serializeTxBodyXml(newDoc));

    expect(result).toContain('<a:prstGeom prst="rect">'); // spPr 하위 구조가 그대로 남아 있다
    expect(result).toContain('</p:spTree>'); // 최상위 컨테이너도 손상되지 않았다
    const [shape] = parseSlideShapes(result);
    expect(shape.doc.textContent).toBe('Updated');
    expect(shape.w).toBe(100); // spPr의 위치/크기 정보도 그대로 보존됐다
  });

  test('removeShape removes exactly one <p:sp> block and leaves an empty but structurally valid <p:spTree>', () => {
    const result = removeShape(oneShapeWithSpPrAndSpTree, '8');

    expect(parseSlideShapes(result)).toHaveLength(0);
    expect(result).toContain('<p:spTree>'); // <p:spTree> 자체는 절대 지워지지 않는다
    expect(result).toContain('</p:spTree>');
    expect(result).not.toContain('<p:spPr>'); // 도형과 함께 그 하위 spPr도 통째로 사라졌다
  });
});
