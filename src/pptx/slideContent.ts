import { XMLParser } from 'fast-xml-parser';
import type { Node as PMNode } from 'prosemirror-model';
import { pptxTextSchema as schema } from './textSchema.js';

export interface ParsedShape {
  readonly shapeId: string;
  readonly x: number | null;
  readonly y: number | null;
  readonly w: number | null;
  readonly h: number | null;
  readonly doc: PMNode;
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  // 기본값(true)이면 <a:t> 텍스트의 앞뒤 공백을 잘라낸다 — pptx 런(run) 경계의 공백(예: "Bold"
  // 다음 런의 " plain")이 서식 구분을 위한 의미 있는 문자라서, 잘리면 텍스트가 눈에 띄게
  // 달라진다(docx/import.ts가 같은 이유로 trimValues: false를 쓰는 것과 동일한 근거).
  trimValues: false,
  // <a:r>가 하나뿐이면 배열이 아니라 객체로 파싱되는 fast-xml-parser의 기본 동작을 막는다 —
  // 문단 하나에 런이 하나뿐인 경우가 흔해서, 이 옵션 없이는 매번 "배열인지 객체인지" 분기해야 한다.
  isArray: (name) => ['a:p', 'a:r', 'p:sp'].includes(name),
});

type XmlNode = Record<string, unknown>;

function asArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined) return [];
  return Array.isArray(value) ? value : [value];
}

/** <a:rPr>에서 마크(굵게/기울임/밑줄/색/크기)를 뽑아 그 런의 텍스트에 적용한다. */
function marksFromRPr(rPr: XmlNode | undefined): ReturnType<typeof schema.marks.strong.create>[] {
  if (!rPr) return [];
  const marks = [];
  if (rPr['@_b'] === '1') marks.push(schema.marks.strong.create());
  if (rPr['@_i'] === '1') marks.push(schema.marks.em.create());
  if (rPr['@_u'] && rPr['@_u'] !== 'none') marks.push(schema.marks.underline.create());
  const solidFill = rPr['a:solidFill'] as XmlNode | undefined;
  const srgbClr = solidFill?.['a:srgbClr'] as XmlNode | undefined;
  const colorVal = srgbClr?.['@_val'] as string | undefined;
  if (colorVal) marks.push(schema.marks.textColor.create({ color: `#${colorVal}` }));
  const sz = rPr['@_sz'] as string | undefined;
  if (sz) marks.push(schema.marks.fontSize.create({ size: Number(sz) / 100 }));
  return marks;
}

function paragraphFromXml(pNode: XmlNode): PMNode {
  const runs = asArray<XmlNode>(pNode['a:r'] as XmlNode | XmlNode[] | undefined);
  const inline = runs
    .map((run) => {
      const text = (run['a:t'] as string | number | undefined) ?? '';
      const textStr = String(text);
      if (textStr === '') return null;
      return schema.text(textStr, marksFromRPr(run['a:rPr'] as XmlNode | undefined));
    })
    .filter((n): n is NonNullable<typeof n> => n !== null);
  return schema.node('paragraph', null, inline);
}

function emuAttr(xfrmChild: XmlNode | undefined, attr: '@_x' | '@_y' | '@_cx' | '@_cy'): number | null {
  const value = xfrmChild?.[attr];
  return typeof value === 'number' ? value : value !== undefined ? Number(value) : null;
}

/** pptx 슬라이드 XML(p:sld) 안의 각 텍스트 도형(p:sp)을 위치 정보 + ProseMirror 문서로 파싱한다.
 * <p:txBody>가 없는 도형(그림 등)은 편집 대상이 아니므로 결과에서 제외한다. */
export function parseSlideShapes(slideXml: string): ParsedShape[] {
  const parsed = parser.parse(slideXml) as XmlNode;
  const sld = parsed['p:sld'] as XmlNode;
  const spTree = (sld['p:cSld'] as XmlNode)['p:spTree'] as XmlNode;
  const shapeNodes = asArray<XmlNode>(spTree['p:sp'] as XmlNode | XmlNode[] | undefined);

  const shapes: ParsedShape[] = [];
  for (const sp of shapeNodes) {
    const txBody = sp['p:txBody'] as XmlNode | undefined;
    if (!txBody) continue; // <p:txBody> 없는 도형(그림 등)은 편집 대상이 아니다 — 건너뛴다.

    const cNvPr = (sp['p:nvSpPr'] as XmlNode)['p:cNvPr'] as XmlNode;
    const shapeId = String(cNvPr['@_id']);

    const xfrm = (sp['p:spPr'] as XmlNode | undefined)?.['a:xfrm'] as XmlNode | undefined;
    const off = xfrm?.['a:off'] as XmlNode | undefined;
    const ext = xfrm?.['a:ext'] as XmlNode | undefined;

    const paragraphs = asArray<XmlNode>(txBody['a:p'] as XmlNode | XmlNode[] | undefined).map(paragraphFromXml);
    const doc = schema.node(
      'doc',
      null,
      paragraphs.length > 0 ? paragraphs : [schema.node('paragraph', null, [])],
    );

    shapes.push({
      shapeId,
      x: emuAttr(off, '@_x'),
      y: emuAttr(off, '@_y'),
      w: emuAttr(ext, '@_cx'),
      h: emuAttr(ext, '@_cy'),
      doc,
    });
  }
  return shapes;
}

function escapeXmlText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** 텍스트 노드의 마크를 <a:rPr>로 되돌린다. schema.marks 순서와 무관하게 항상 같은 속성
 * 순서로 내보내 스냅샷/디버깅이 안정적이게 한다. */
function rPrXmlFor(node: PMNode): string {
  const b = node.marks.some((m) => m.type.name === 'strong') ? ' b="1"' : '';
  const i = node.marks.some((m) => m.type.name === 'em') ? ' i="1"' : '';
  const u = node.marks.some((m) => m.type.name === 'underline') ? ' u="sng"' : '';
  const sizeMark = node.marks.find((m) => m.type.name === 'fontSize');
  const sz = sizeMark ? ` sz="${Math.round((sizeMark.attrs.size as number) * 100)}"` : '';
  const colorMark = node.marks.find((m) => m.type.name === 'textColor');
  const attrs = `${b}${i}${u}${sz}`;
  if (!colorMark) {
    return attrs ? `<a:rPr${attrs}/>` : '';
  }
  const hex = (colorMark.attrs.color as string).replace('#', '').toUpperCase();
  return `<a:rPr${attrs}><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></a:rPr>`;
}

export function serializeTxBodyXml(doc: PMNode): string {
  const paragraphsXml: string[] = [];
  doc.forEach((paragraph) => {
    const runsXml: string[] = [];
    paragraph.forEach((textNode) => {
      if (!textNode.isText || !textNode.text) return;
      const rPr = rPrXmlFor(textNode);
      runsXml.push(`<a:r>${rPr}<a:t>${escapeXmlText(textNode.text)}</a:t></a:r>`);
    });
    paragraphsXml.push(`<a:p>${runsXml.join('')}</a:p>`);
  });
  return `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphsXml.join('')}</p:txBody>`;
}

/** 이미 있는 도형(shapeId로 찾음)의 <p:txBody>만 통째로 교체한다 — 그 도형의 <p:spPr>(위치·
 * 스타일)이나 다른 도형은 전혀 건드리지 않는다. 정규식으로 그 <p:sp>...</p:sp> 블록을 찾고,
 * 그 안의 <p:txBody>...</p:txBody>만 문자열 치환한다(전체를 다시 파싱해 직렬화하면 원본에
 * 있던, 우리가 모델링하지 않는 속성(원본에만 있는 <p:spPr> 하위 요소 등)이 유실될 위험이 있어
 * 문자열 서저리를 유지한다 — src/pptx/collab.ts의 기존 방식과 같은 원칙). */
export function spliceShapeTxBody(slideXml: string, shapeId: string, txBodyXml: string): string {
  const spBlockRe = /<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g;
  let replaced = false;
  const result = slideXml.replace(spBlockRe, (spBlock) => {
    const idMatch = /<p:cNvPr\b[^>]*\bid="(\d+)"/.exec(spBlock);
    if (!idMatch || idMatch[1] !== shapeId) return spBlock;
    replaced = true;
    return spBlock.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, txBodyXml);
  });
  if (!replaced) {
    throw new Error(`spliceShapeTxBody: shapeId ${shapeId}를 슬라이드 XML에서 찾지 못했습니다.`);
  }
  return result;
}

/** 원본에 없던 도형(새로 추가한 텍스트박스)을 <p:spTree> 끝에 끼워 넣는다. shapeId는
 * new-<uuid> 형태를 그대로 <p:cNvPr name="..">에는 안 쓰고 id 자리(문자열)에 그대로 쓴다 —
 * PowerPoint의 <p:cNvPr id>는 원래 숫자지만, 우리가 다시 읽을 때(parseSlideShapes)도
 * 문자열로만 비교하므로 실제로는 아무 문제가 없다(PowerPoint 자체가 다시 열 때 재검증하지
 * 않는 필드). */
export function appendNewShape(
  slideXml: string,
  meta: { shapeId: string; x: number; y: number; w: number; h: number },
  txBodyXml: string,
): string {
  const spXml =
    `<p:sp><p:nvSpPr><p:cNvPr id="${meta.shapeId}" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${meta.x}" y="${meta.y}"/><a:ext cx="${meta.w}" cy="${meta.h}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    txBodyXml +
    '</p:sp>';
  const closeIdx = slideXml.lastIndexOf('</p:spTree>');
  if (closeIdx === -1) {
    throw new Error('appendNewShape: 슬라이드 XML에서 </p:spTree>를 찾지 못했습니다.');
  }
  return slideXml.slice(0, closeIdx) + spXml + slideXml.slice(closeIdx);
}
