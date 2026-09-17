import { describe, expect, test } from 'vitest';
import { parseSlideVisuals } from '../src/pptx/slideContent.js';

const XML_HEAD =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/>';
const XML_TAIL = '</p:spTree></p:cSld></p:sld>';

function zip(entries: Record<string, string | Uint8Array>): Record<string, Uint8Array> {
  const out: Record<string, Uint8Array> = {};
  for (const [key, value] of Object.entries(entries)) {
    out[key] = typeof value === 'string' ? new TextEncoder().encode(value) : value;
  }
  return out;
}

describe('parseSlideVisuals', () => {
  test('extracts an image with its position and a resolved data URI', () => {
    const slideXml =
      XML_HEAD +
      '<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId2"/></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="3000" cy="4000"/></a:xfrm></p:spPr></p:pic>' +
      XML_TAIL;
    const relsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId2" Type=".../image" Target="../media/image1.png"/>' +
      '</Relationships>';
    const pngBytes = new Uint8Array([137, 80, 78, 71]); // 진짜 PNG일 필요는 없다 — base64 인코딩 경로만 검증

    const entries = zip({
      'ppt/slides/slide1.xml': slideXml,
      'ppt/slides/_rels/slide1.xml.rels': relsXml,
      'ppt/media/image1.png': pngBytes,
    });

    const visuals = parseSlideVisuals(entries, 'slide1.xml');

    expect(visuals).toHaveLength(1);
    expect(visuals[0]).toMatchObject({ kind: 'image', x: 1000, y: 2000, w: 3000, h: 4000 });
    expect(visuals[0].dataUri).toBe(`data:image/png;base64,${Buffer.from(pngBytes).toString('base64')}`);
  });

  test('skips an image whose media format the browser cannot render (emf)', () => {
    const slideXml =
      XML_HEAD +
      '<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId2"/></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></p:spPr></p:pic>' +
      XML_TAIL;
    const relsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId2" Type=".../image" Target="../media/image1.emf"/>' +
      '</Relationships>';
    const entries = zip({
      'ppt/slides/slide1.xml': slideXml,
      'ppt/slides/_rels/slide1.xml.rels': relsXml,
      'ppt/media/image1.emf': new Uint8Array([1, 2, 3]),
    });

    expect(parseSlideVisuals(entries, 'slide1.xml')).toHaveLength(0);
  });

  test('extracts a textless shape as a filled rectangle', () => {
    const slideXml =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="5" name="Rectangle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr></p:sp>' +
      XML_TAIL;
    const entries = zip({ 'ppt/slides/slide1.xml': slideXml });

    const visuals = parseSlideVisuals(entries, 'slide1.xml');

    expect(visuals).toEqual([{ kind: 'shape', x: 100, y: 200, w: 300, h: 400, fill: '#4472C4' }]);
  });

  test('a shape with a <p:txBody> is left for parseSlideShapes — not duplicated here', () => {
    const slideXml =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="6" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>' +
      XML_TAIL;
    const entries = zip({ 'ppt/slides/slide1.xml': slideXml });

    expect(parseSlideVisuals(entries, 'slide1.xml')).toHaveLength(0);
  });

  test('extracts a table with real cell text content, row by row', () => {
    const slideXml =
      XML_HEAD +
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Table"/></p:nvGraphicFramePr>' +
      '<p:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></p:xfrm>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
      '<a:tbl><a:tr h="1"><a:tc><a:txBody><a:p><a:r><a:t>A1</a:t></a:r></a:p></a:txBody></a:tc>' +
      '<a:tc><a:txBody><a:p><a:r><a:t>B1</a:t></a:r></a:p></a:txBody></a:tc></a:tr>' +
      '<a:tr h="1"><a:tc><a:txBody><a:p><a:r><a:t>A2</a:t></a:r></a:p></a:txBody></a:tc>' +
      '<a:tc><a:txBody><a:p><a:r><a:t>B2</a:t></a:r></a:p></a:txBody></a:tc></a:tr>' +
      '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>' +
      XML_TAIL;
    const entries = zip({ 'ppt/slides/slide1.xml': slideXml });

    const visuals = parseSlideVisuals(entries, 'slide1.xml');

    expect(visuals).toEqual([
      {
        kind: 'table',
        x: 10,
        y: 20,
        w: 30,
        h: 40,
        rows: [
          ['A1', 'B1'],
          ['A2', 'B2'],
        ],
      },
    ]);
  });

  test('a chart graphicFrame becomes a labeled placeholder at the correct position', () => {
    const slideXml =
      XML_HEAD +
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8" name="Chart"/></p:nvGraphicFramePr>' +
      '<p:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></p:xfrm>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      '<c:chart r:id="rId3"/></a:graphicData></a:graphic></p:graphicFrame>' +
      XML_TAIL;
    const entries = zip({ 'ppt/slides/slide1.xml': slideXml });

    expect(parseSlideVisuals(entries, 'slide1.xml')).toEqual([
      { kind: 'placeholder', x: 1, y: 2, w: 3, h: 4, label: '차트' },
    ]);
  });

  test('a diagram (SmartArt) graphicFrame becomes a labeled placeholder', () => {
    const slideXml =
      XML_HEAD +
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="SmartArt"/></p:nvGraphicFramePr>' +
      '<p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">' +
      '<dgm:relIds/></a:graphicData></a:graphic></p:graphicFrame>' +
      XML_TAIL;
    const entries = zip({ 'ppt/slides/slide1.xml': slideXml });

    const visuals = parseSlideVisuals(entries, 'slide1.xml');
    expect(visuals[0]).toMatchObject({ kind: 'placeholder', label: '스마트아트' });
  });

  test('returns an empty array when the slide part is missing from the zip', () => {
    expect(parseSlideVisuals(zip({}), 'slide1.xml')).toEqual([]);
  });
});
