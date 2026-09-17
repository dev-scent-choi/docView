import { XMLParser } from 'fast-xml-parser';
import type { Node as PMNode } from 'prosemirror-model';
import { pptxTextSchema as schema } from './textSchema.js';
import type { SlideVisual } from '../collab/pptxShared.js';

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
  // 기본값(true)이면 태그 텍스트가 숫자처럼 보일 때(예: "007", "1.0", "0.10", "+7", "1e3",
  // "0x1A") fast-xml-parser가 조용히 JS 숫자로 바꾸고, 다시 문자열로 꺼낼 때 그 숫자의 "canonical"
  // 형태로 되돌린다("007" -> "7", "1.0" -> "1" 등) — 이 파서가 읽는 태그 텍스트 중 의미 있는 건
  // <a:t>(런 텍스트)뿐이다(위치/크기 같은 속성값은 attributeNamePrefix로 따로 읽혀
  // parseAttributeValue 기본값 false의 영향만 받으므로 이 옵션과 무관하다). 이 문제는 그 도형을
  // 한 번도 편집 안 하면 무해하다(baseline 비교 양쪽이 똑같이 망가지므로 "편집 안 된 도형은
  // 건드리지 않는다" 로직이 여전히 원본을 그대로 둔다) — 하지만 그 도형 안 어느 런이든 한 번
  // 편집되는 순간 <p:txBody> 전체가 이미 망가진 파싱 결과로부터 재조립되어, 숫자처럼 생긴 그
  // 텍스트가 저장 파일에 영구히 다른 값으로 남는다(FINAL REVIEW C3). parseTagValue만 꺼서 태그
  // 텍스트는 항상 원문 문자열 그대로 읽는다.
  parseTagValue: false,
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
  // CT_TextBody(OOXML 스키마)는 <a:p>가 최소 1개 있어야 한다(minOccurs="1") — 문단이 하나도 없는
  // <p:txBody>는 PowerPoint가 열 수 없는 손상된 파일이다(FINAL REVIEW C2). pptxTextSchema는
  // doc: 'paragraph+'라 정상적으로 만들어진 문서는 항상 문단이 1개 이상이지만, "텍스트박스 추가"
  // 직후 아직 아무도 타이핑하지 않은 빈 Y.XmlFragment를 y-prosemirror가 ProseMirror 문서로
  // 되돌릴 때는 문단이 0개인 doc이 나올 수 있다(마운트된 에디터 자신의 초기화가 최소 문단을 채워
  // 넣기 전에 자동저장이 먼저 끼어드는 경우 등) — 그 경우에도 항상 빈 문단 하나는 반드시 내보낸다.
  // 빈 문단의 출력 형태는 위 루프가 런이 0개인 문단에 이미 쓰는 것과 똑같이 `<a:p></a:p>`로
  // 맞춘다(자기 닫힘 `<a:p/>` 대신 — 이 함수의 기존 출력 스타일과 일관되게).
  if (paragraphsXml.length === 0) {
    paragraphsXml.push('<a:p></a:p>');
  }
  return `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphsXml.join('')}</p:txBody>`;
}

/** 이미 있는 도형(shapeId로 찾음)의 <p:txBody>만 통째로 교체한다 — 그 도형의 <p:spPr>(위치·
 * 스타일)이나 다른 도형은 전혀 건드리지 않는다. 정규식으로 그 <p:sp>...</p:sp> 블록을 찾고,
 * 그 안의 <p:txBody>...</p:txBody>만 문자열 치환한다(전체를 다시 파싱해 직렬화하면 원본에
 * 있던, 우리가 모델링하지 않는 속성(원본에만 있는 <p:spPr> 하위 요소 등)이 유실될 위험이 있어
 * 문자열 서저리를 유지한다 — src/pptx/collab.ts의 기존 방식과 같은 원칙).
 *
 * id 매칭은 숫자로 제한하지 않는다(`id="([^"]+)"` — 따옴표로 감싸인 값이면 무엇이든 매치) —
 * 원본 슬라이드의 도형은 항상 숫자 id를 쓰지만, 클라이언트가 새로 추가한 텍스트박스는
 * `new-<uuid>` 형태의 문자열 id를 쓴다(appendNewShape 참고). 이 함수가 숫자만 매치했다면
 * 두 번째 저장부터 그 도형을 "찾지 못했다"고 판단해 매번 새로 append하게 되어 저장할 때마다
 * 같은 텍스트박스가 중복 추가되는 버그가 있었다 — 이 완화는 순수하게 넓히는 방향이라(숫자
 * id는 전부 `[^"]+`에도 그대로 매치되므로) 기존에 매치되던 것은 전부 그대로 매치된다.
 *
 * 블록 탐지 정규식(`<p:sp\b...`)도 마찬가지로 순수 확장이다 — 원래 리터럴 `<p:sp>`만 매치했는데,
 * PowerPoint가 실제로 자주 내보내는 `<p:sp useBgFill="1">`처럼 여는 태그에 속성이 붙은 형태는
 * 매치하지 못했다(이 함수는 "못 찾음"으로 throw하지는 않지만, 아래 hasShape가 같은 이유로 "없다"고
 * 잘못 답하면 이 함수가 아니라 appendNewShape가 불려 같은 도형이 중복 추가됐다 — Task 13 fix로
 * `\b`(단어 경계)로 넓혀 속성 유무와 무관하게 매치하면서도 `<p:spPr>`/`<p:spTree>`는 계속
 * 거른다: `sp` 다음이 `Pr`/`Tree`처럼 단어 문자로 이어지면 그 사이엔 경계가 없기 때문).
 *
 * `<p:txBody>...</p:txBody>` 탐지 정규식도 (리뷰에서 지적된) 같은 계열의 문제를 안고 있었다 —
 * 리터럴 `<p:txBody>`만 매치해, 여는 태그에 속성이 붙은 변형(OOXML 스키마 자체는 CT_TextBody에
 * 속성을 정의하지 않으므로 `<p:sp>`보다 훨씬 드문 트리거지만, 어떤 재직렬화 도구가 네임스페이스
 * 재선언 등을 붙이면 여전히 가능하다)을 만나면 이 함수가 "<p:txBody> 없음"으로 잘못 throw했다.
 * 아래에서 `<p:txBody\b`로 넓혀 같은 방식으로 고쳤다. */
export function spliceShapeTxBody(slideXml: string, shapeId: string, txBodyXml: string): string {
  const spBlockRe = /<p:sp\b(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g;
  let replaced = false;
  const result = slideXml.replace(spBlockRe, (spBlock) => {
    const idMatch = /<p:cNvPr\b[^>]*\bid="([^"]+)"/.exec(spBlock);
    if (!idMatch || idMatch[1] !== shapeId) return spBlock;
    // id는 맞아도 이 <p:sp> 블록에 <p:txBody>가 구조적으로 없을 수 있다(그림 등) — 그 경우
    // 아래 정규식 치환은 아무것도 바꾸지 못한 채 spBlock을 그대로 돌려주는데, 그걸 놓치고
    // replaced만 true로 찍으면 "교체 성공"으로 잘못 보고되며 수정이 조용히 사라진다.
    // 원본 유실 금지가 최우선 기준이므로, 실제로 바꿀 대상이 없으면 여기서 즉시 throw한다.
    //
    // 이 <p:txBody> 탐지 정규식도 위 <p:sp\b 블록 정규식과 정확히 같은 이유로 `\b`(단어 경계)를
    // 쓴다 — 리터럴 <p:txBody>만 매치하면, 어떤 재직렬화 도구가 <p:txBody>에 네임스페이스
    // 재선언 같은 속성을 붙인(OOXML 스키마 자체는 CT_TextBody에 속성을 정의하지 않지만, 실제로
    // 그런 속성 붙은 변형을 만들어내는 도구가 있을 수 있다) 슬라이드를 만나면, parseSlideShapes
    // (실제 XML 파서)는 그 도형의 텍스트 본문을 정상적으로 인식하는데 이 함수만 "<p:txBody> 없음"
    // 으로 잘못 throw한다 — 이 throw는 저장/다운로드 경로 안에서 일어나므로 그 문서의 이후 모든
    // 저장을 영구히 막는다(<p:sp\b 정규식을 리터럴 <p:sp>로 뒀을 때 일어났던 것과 완전히 같은
    // 종류의 결과). `\b`는 <p:txBody>(바로 뒤가 `>`)도, <p:txBody xmlns:a="...">(바로 뒤가 공백)도
    // 매치하면서, <p:txBody> 접두어를 공유하되 실제로는 다른 태그(가정상의 <p:txBodyFoo> 등, 바로
    // 뒤가 단어 문자로 이어짐)는 여전히 거른다.
    if (!/<p:txBody\b[\s\S]*?<\/p:txBody>/.test(spBlock)) {
      throw new Error(`spliceShapeTxBody: shapeId ${shapeId}의 <p:sp> 블록에 <p:txBody>가 없습니다.`);
    }
    replaced = true;
    return spBlock.replace(/<p:txBody\b[\s\S]*?<\/p:txBody>/, txBodyXml);
  });
  if (!replaced) {
    throw new Error(`spliceShapeTxBody: shapeId ${shapeId}를 슬라이드 XML에서 찾지 못했습니다.`);
  }
  return result;
}

/** shapeId를 가진 <p:sp> 블록이 슬라이드 XML 안에 이미 존재하는지만 확인한다(내용은 건드리지
 * 않는다) — spliceShapeTxBody와 정확히 같은 블록 탐색/속성 매칭 정규식을 쓴다(둘이 어긋나면
 * "있다"고 했는데 spliceShapeTxBody는 못 찾는 모순이 생길 수 있으므로 반드시 같아야 한다).
 *
 * src/pptx/collab.ts의 export 병합 루프가 도형별로 이 함수를 먼저 호출해 "이미 있는 도형(교체 —
 * spliceShapeTxBody)"과 "처음 쓰는 도형(추가 — appendNewShape)"을 구분한다. 예전에는 shapeId가
 * `new-`로 시작하는지만 보고 무조건 appendNewShape를 불렀는데, 그러면 이미 한 번 저장되어 슬라이드
 * XML에 실제로 존재하는 `new-` 도형을 두 번째 저장에서 또 추가해 버려 저장할 때마다 같은
 * 텍스트박스가 중복되는 버그가 있었다 — 이 함수로 먼저 존재 여부를 확인하면 그 문제가 사라진다.
 *
 * **Task 13 fix로 고친, 그와 같은 계열의 별개 버그**: 이 함수가 쓰는 블록 탐지 정규식이 최근까지
 * (Task 3/6이 만든 그대로) 리터럴 `<p:sp>`만 매치했다 — 여는 태그에 속성이 붙은
 * `<p:sp useBgFill="1">` 같은, PowerPoint가 실제로 흔히 만들어내는 형태는 매치하지 못했다. 그
 * 결과 이 함수가 실제로 존재하는 속성 있는 도형을 "없다"고 잘못 답해, 위 병합 루프가 교체
 * (spliceShapeTxBody) 대신 추가(appendNewShape) 분기를 타 버려 그 도형이 같은 shapeId로 중복
 * 저장됐다(구 텍스트와 새 텍스트가 둘 다 파일에 남는 조용한 데이터 손상). `\b`(단어 경계)로
 * 정규식을 넓혀 지금은 속성 유무와 무관하게 매치한다. */
export function hasShape(slideXml: string, shapeId: string): boolean {
  const spBlockRe = /<p:sp\b(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g;
  for (const spBlock of slideXml.match(spBlockRe) ?? []) {
    const idMatch = /<p:cNvPr\b[^>]*\bid="([^"]+)"/.exec(spBlock);
    if (idMatch && idMatch[1] === shapeId) return true;
  }
  return false;
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

/** shapeId에 해당하는 <p:sp> 블록 전체를 슬라이드 XML에서 제거한다(사용자가 텍스트박스를
 * 삭제했을 때 씀 — src/pptx/collab.ts의 exportPptxFromSlides가 "baseline에는 있었는데 지금
 * perSlide에는 없는" 도형에 대해 호출한다). hasShape/spliceShapeTxBody와 정확히 같은 블록 탐색/
 * 속성 매칭 정규식을 쓴다(id 매칭을 숫자로 제한하지 않으므로 `new-<uuid>` 형태의 id도 그대로
 * 매치된다) — 셋이 어긋나면 "있다"고 판단한 도형을 못 찾거나, 지워야 할 도형을 못 찾는 모순이
 * 생길 수 있으므로 반드시 같아야 한다. 다른 도형의 <p:sp> 블록은 전혀 건드리지 않는다.
 *
 * 못 찾으면 throw한다(이 파일의 다른 함수들과 동일한 "절대 조용히 실패하지 않는다" 원칙) — 이
 * 호출부(exportPptxFromSlides)는 항상 baseline에 실제로 존재하는 shapeId로만 이 함수를 부르므로,
 * 정상적인 사용에서는 이 throw가 절대 발생하지 않아야 한다(발생한다면 그 자체가 병합 로직의
 * 불변식이 깨졌다는 신호이므로, 조용히 넘어가지 않고 크게 실패하는 편이 "원본 유실 금지"
 * 원칙에 맞다).
 *
 * **주의(Task 13 fix 이전에는 이 불변식이 실제로 깨져 있었다)**: baseline은 parseSlideShapes(실제
 * XML 파서)로 만들어지므로 여는 태그에 속성이 붙은 `<p:sp useBgFill="1">` 같은 도형도 정확히
 * 인식해 baseline에 포함시킨다. 반면 이 함수가 쓰는 블록 탐지 정규식은 최근까지 리터럴 `<p:sp>`만
 * 매치했다 — 그래서 baseline에는 있고(파서가 봤으니) perSlide에는 없는(사용자가 지웠으니) 속성
 * 있는 도형을 지우려 하면, 그 블록이 XML 안에 실제로 있는데도 정규식이 찾지 못해 이 throw가
 * 정상적인 삭제 조작만으로 발생했다 — 그리고 이 throw는 exportPptxFromSlides(저장/다운로드 경로)
 * 안에서 일어나므로, 그 문서의 이후 모든 저장·다운로드를 영구히 막았다(orphan된 baseline id는
 * 절대 사라지지 않으므로 사용자 쪽에서 되돌릴 방법도 없었다). 지금은 세 함수의 블록 정규식을
 * 전부 `<p:sp\b`로 넓혀(속성 유무와 무관하게 매치하되 `<p:spPr>`/`<p:spTree>`는 여전히 제외)
 * baseline과 이 함수가 항상 같은 도형 집합을 인식한다.
 *
 * **위 "절대 발생하지 않아야 한다"는 주장의 정확한 범위**: 이건 `<p:sp\b` 블록 탐지 정규식에
 * 한정된 얘기다 — 세 함수가 함께 쓰는 또 다른 정규식, `<p:cNvPr>`에서 id 속성값을 뽑는
 * `/<p:cNvPr\b[^>]*\bid="([^"]+)"/`는 여전히 실제 XML 파서와 완전히 동일하지 않다. PowerPoint와
 * LibreOffice가 실제로 내보내는, 큰따옴표를 쓰고 `id`가 다른 속성보다 먼저 오는 직렬화 형태에서는
 * 항상 일치하지만, 합법적이지만 어느 실제 생산자도 만들어내지 않는 변형 — 작은따옴표
 * (`id='7'`), `=` 앞뒤 공백(`id = "7"`), 또는 `id`보다 앞선 속성의 값 안에 이스케이프 안 된 `>`가
 * 들어있는 경우(예: `name="a > b" id="7"`) — 에서는 이 id 정규식이 이론적으로 여전히 실제 파서와
 * 다르게 판단할 수 있다. 그러므로 "이 throw가 절대 발생하지 않는다"는 건 실제 생산자가 만들어내는
 * 직렬화 형태에 한해 참이지, 위와 같은 이론상 가능한 모든 합법적 XML 변형까지 포괄하는 절대적
 * 보장은 아니다. */
export function removeShape(slideXml: string, shapeId: string): string {
  const spBlockRe = /<p:sp\b(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g;
  let removed = false;
  const result = slideXml.replace(spBlockRe, (spBlock) => {
    const idMatch = /<p:cNvPr\b[^>]*\bid="([^"]+)"/.exec(spBlock);
    if (!idMatch || idMatch[1] !== shapeId) return spBlock;
    removed = true;
    return '';
  });
  if (!removed) {
    throw new Error(`removeShape: shapeId ${shapeId}를 슬라이드 XML에서 찾지 못했습니다.`);
  }
  return result;
}

/* ==================== 표시 전용 요소(이미지·단순 도형·표/차트/스마트아트) ====================
 * 아래 전부는 parseSlideShapes(편집 대상 텍스트박스)와 완전히 별개 경로다 — 결과가
 * SlideVisual[]로 slideVisuals(협업 편집 상태와 무관한 표시 전용 맵)에만 쓰이고, 절대
 * export되지 않는다(그 요소들이 있던 자리는 원본 XML 그대로 복사되므로 여기서 뭘 하든
 * 저장되는 파일에는 영향이 없다). */

/** .rels 파일에서 "rId → Target" 매핑만 뽑는다(src/pptx/collab.ts의 slideRelsMap과 같은 패턴,
 * 이 파일은 그쪽을 import할 이유가 없을 만큼 단순해 그대로 복제했다). */
function relsMap(relsXml: string): Map<string, string> {
  const map = new Map<string, string>();
  const relRe = /<Relationship\b[^>]*\/>/g;
  for (const relTag of relsXml.match(relRe) ?? []) {
    const idMatch = /Id="(rId\d+)"/.exec(relTag);
    const targetMatch = /Target="([^"]+)"/.exec(relTag);
    if (idMatch && targetMatch) map.set(idMatch[1], targetMatch[1]);
  }
  return map;
}

/** .rels의 Target(예: "../media/image1.png")을 슬라이드 파트 기준 zip 내부 절대 경로로 푼다. */
function resolveRelativePath(basePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const stack = basePart.split('/').slice(0, -1);
  for (const part of target.split('/')) {
    if (part === '..') stack.pop();
    else if (part !== '.') stack.push(part);
  }
  return stack.join('/');
}

const IMAGE_MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  // emf/wmf/tiff 등은 의도적으로 뺐다 — 브라우저가 <img>로 직접 그릴 수 없는 포맷이라
  // data URI를 만들어봐야 화면에 아무것도 안 나온다(깨진 이미지 아이콘보다는 그냥 생략이 낫다).
};

function fillColorFromSpPr(spPr: XmlNode | undefined): string | undefined {
  const solidFill = spPr?.['a:solidFill'] as XmlNode | undefined;
  const srgbClr = solidFill?.['a:srgbClr'] as XmlNode | undefined;
  const val = srgbClr?.['@_val'] as string | undefined;
  // a:schemeClr(테마 참조 색, 예: "accent1")는 테마 XML까지 읽어야 풀리므로 이번 범위에서는
  // 생략한다 — 채우기 없이 위치만 정확한 도형으로 표시된다(완전히 안 보이는 것보다는 낫다).
  return val ? `#${val}` : undefined;
}

function graphicFrameXfrm(gf: XmlNode) {
  const xfrm = gf['p:xfrm'] as XmlNode | undefined; // 다른 도형과 달리 p:spPr 없이 바로 p:xfrm
  const off = xfrm?.['a:off'] as XmlNode | undefined;
  const ext = xfrm?.['a:ext'] as XmlNode | undefined;
  return { x: emuAttr(off, '@_x'), y: emuAttr(off, '@_y'), w: emuAttr(ext, '@_cx'), h: emuAttr(ext, '@_cy') };
}

/** <a:tc>(표 셀) 하나의 텍스트만 뽑는다 — 서식은 버리고 문단은 줄바꿈으로 잇는다(표는 편집
 * 대상이 아니라 화면 표시용이라 이 정도 단순화로 충분하다). */
function tableCellText(tc: XmlNode): string {
  const txBody = tc['a:txBody'] as XmlNode | undefined;
  if (!txBody) return '';
  return asArray<XmlNode>(txBody['a:p'] as XmlNode | XmlNode[] | undefined)
    .map((p) =>
      asArray<XmlNode>(p['a:r'] as XmlNode | XmlNode[] | undefined)
        .map((r) => String((r['a:t'] as string | number | undefined) ?? ''))
        .join(''),
    )
    .join('\n');
}

const TABLE_GRAPHIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';
const CHART_GRAPHIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

/** 슬라이드 XML 안의 이미지(p:pic)·글자 없는 단순 도형(p:sp)·표/차트/스마트아트
 * (p:graphicFrame)를 화면 표시용으로 뽑는다. 표는 셀 텍스트까지 그리고, 차트·스마트아트는
 * (렌더링 엔진 없이는 실제로 그릴 수 없어) 위치만 맞는 자리표시자로 남긴다. */
export function parseSlideVisuals(zipEntries: Record<string, Uint8Array>, sourceSlidePart: string): SlideVisual[] {
  const slideXmlBytes = zipEntries[`ppt/slides/${sourceSlidePart}`];
  if (!slideXmlBytes) return [];
  const slideXml = Buffer.from(slideXmlBytes).toString('utf-8');
  const relsBytes = zipEntries[`ppt/slides/_rels/${sourceSlidePart}.rels`];
  const rels = relsBytes ? relsMap(Buffer.from(relsBytes).toString('utf-8')) : new Map<string, string>();

  function resolveImageDataUri(rId: string | undefined): string | undefined {
    if (!rId) return undefined;
    const target = rels.get(rId);
    if (!target) return undefined;
    const normalized = resolveRelativePath(`ppt/slides/${sourceSlidePart}`, target);
    const mediaBytes = zipEntries[normalized];
    if (!mediaBytes) return undefined;
    const ext = normalized.split('.').pop()?.toLowerCase() ?? '';
    const mime = IMAGE_MIME_BY_EXT[ext];
    if (!mime) return undefined;
    return `data:${mime};base64,${Buffer.from(mediaBytes).toString('base64')}`;
  }

  const parsed = parser.parse(slideXml) as XmlNode;
  const sld = parsed['p:sld'] as XmlNode;
  const spTree = (sld['p:cSld'] as XmlNode)['p:spTree'] as XmlNode;
  const visuals: SlideVisual[] = [];

  for (const pic of asArray<XmlNode>(spTree['p:pic'] as XmlNode | XmlNode[] | undefined)) {
    const spPr = pic['p:spPr'] as XmlNode | undefined;
    const xfrm = spPr?.['a:xfrm'] as XmlNode | undefined;
    const off = xfrm?.['a:off'] as XmlNode | undefined;
    const ext = xfrm?.['a:ext'] as XmlNode | undefined;
    const blip = (pic['p:blipFill'] as XmlNode | undefined)?.['a:blip'] as XmlNode | undefined;
    const dataUri = resolveImageDataUri(blip?.['@_r:embed'] as string | undefined);
    if (!dataUri) continue; // 관계 유실 또는 브라우저가 못 그리는 포맷 — 자리표시자 없이 생략
    visuals.push({
      kind: 'image',
      x: emuAttr(off, '@_x'),
      y: emuAttr(off, '@_y'),
      w: emuAttr(ext, '@_cx'),
      h: emuAttr(ext, '@_cy'),
      dataUri,
    });
  }

  for (const sp of asArray<XmlNode>(spTree['p:sp'] as XmlNode | XmlNode[] | undefined)) {
    if (sp['p:txBody']) continue; // 텍스트박스는 parseSlideShapes가 이미 편집 대상으로 다룬다
    const spPr = sp['p:spPr'] as XmlNode | undefined;
    const xfrm = spPr?.['a:xfrm'] as XmlNode | undefined;
    const off = xfrm?.['a:off'] as XmlNode | undefined;
    const ext = xfrm?.['a:ext'] as XmlNode | undefined;
    visuals.push({
      kind: 'shape',
      x: emuAttr(off, '@_x'),
      y: emuAttr(off, '@_y'),
      w: emuAttr(ext, '@_cx'),
      h: emuAttr(ext, '@_cy'),
      fill: fillColorFromSpPr(spPr),
    });
  }

  for (const gf of asArray<XmlNode>(spTree['p:graphicFrame'] as XmlNode | XmlNode[] | undefined)) {
    const { x, y, w, h } = graphicFrameXfrm(gf);
    const graphicData = (gf['a:graphic'] as XmlNode | undefined)?.['a:graphicData'] as XmlNode | undefined;
    const uri = graphicData?.['@_uri'] as string | undefined;
    if (uri === TABLE_GRAPHIC_URI) {
      const tbl = graphicData?.['a:tbl'] as XmlNode | undefined;
      const rows = asArray<XmlNode>(tbl?.['a:tr'] as XmlNode | XmlNode[] | undefined).map((tr) =>
        asArray<XmlNode>(tr['a:tc'] as XmlNode | XmlNode[] | undefined).map(tableCellText),
      );
      visuals.push({ kind: 'table', x, y, w, h, rows });
    } else {
      const label = uri === CHART_GRAPHIC_URI ? '차트' : uri?.includes('/diagram') ? '스마트아트' : '콘텐츠';
      visuals.push({ kind: 'placeholder', x, y, w, h, label });
    }
  }

  return visuals;
}
