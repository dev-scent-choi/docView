import { randomUUID } from 'node:crypto';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import * as Y from 'yjs';
import { prosemirrorToYXmlFragment, yXmlFragmentToProseMirrorRootNode } from 'y-prosemirror';
import { appendNewShape, parseSlideShapes, serializeTxBodyXml, spliceShapeTxBody } from './slideContent.js';
import { pptxTextSchema } from './textSchema.js';
import type { SlideEntry, SlideShapeMeta } from '../collab/pptxShared.js';

const SLIDE_TARGET = /Target="(?:[^"]*\/)?(slide\d+\.xml)"/;

/** presentation.xml.rels에서 "rId → slideN.xml" 매핑만 뽑는다(슬라이드 관계만, 마스터 등은 무시). */
function slideRelsMap(presentationRelsXml: string): Map<string, string> {
  const map = new Map<string, string>();
  const relRe = /<Relationship\b[^>]*\/>/g;
  for (const relTag of presentationRelsXml.match(relRe) ?? []) {
    const idMatch = /Id="(rId\d+)"/.exec(relTag);
    const targetMatch = SLIDE_TARGET.exec(relTag);
    if (idMatch && targetMatch) map.set(idMatch[1], targetMatch[1]);
  }
  return map;
}

/** presentation.xml의 sldIdLst에서 r:id를 등장 순서대로 뽑는다. */
function sldIdListOrder(presentationXml: string): string[] {
  const sldIdLstMatch = /<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/.exec(presentationXml);
  if (!sldIdLstMatch) return [];
  const rIds: string[] = [];
  const sldIdRe = /<p:sldId\b[^>]*r:id="(rId\d+)"[^>]*\/>/g;
  let match: RegExpExecArray | null;
  while ((match = sldIdRe.exec(sldIdLstMatch[1])) !== null) rIds.push(match[1]);
  return rIds;
}

/** 원본 pptx를 열어 슬라이드 순서 배열을 초기 상태로 채운다 (room.ts의 seedRoom에서 호출). */
export function seedSlidesFromPptx(pptxBytes: Uint8Array, slidesArray: Y.Array<SlideEntry>): void {
  const zipEntries = unzipSync(pptxBytes);
  const presentationXml = strFromU8(zipEntries['ppt/presentation.xml']);
  const presentationRelsXml = strFromU8(zipEntries['ppt/_rels/presentation.xml.rels']);

  const relsMap = slideRelsMap(presentationRelsXml);
  const order = sldIdListOrder(presentationXml);

  const seeded: SlideEntry[] = order
    .map((rId) => relsMap.get(rId))
    .filter((part): part is string => Boolean(part))
    .map((sourceSlidePart, index) => ({
      id: randomUUID(),
      kind: 'original' as const,
      sourceSlidePart,
      sourcePageNumber: index + 1,
    }));

  slidesArray.insert(0, seeded);
}

const BLANK_SLIDE_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/></p:spTree></p:cSld>' +
  '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
  '</p:sld>';

const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const SLIDE_RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';

function blankSlideRelsXml(layoutTarget: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" ' +
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" ' +
    `Target="${layoutTarget}"/>` +
    '</Relationships>'
  );
}

/** zip 엔트리 키(ppt/slides/slideN.xml) 중 가장 큰 N. 슬라이드가 하나도 없으면 0. */
function maxSlideNumber(zipEntries: Record<string, Uint8Array>): number {
  let max = 0;
  for (const key of Object.keys(zipEntries)) {
    const match = /^ppt\/slides\/slide(\d+)\.xml$/.exec(key);
    if (match) max = Math.max(max, Number(match[1]));
  }
  return max;
}

/** sldIdLst 안 p:sldId의 id 속성 중 가장 큰 값. OOXML 규격상 새 id는 이보다 커야 하고 256 이상이어야
 * 한다 — 원본이 이미 256 이상이므로 +1만 하면 항상 유효하다. */
function maxSldId(presentationXml: string): number {
  const ids = [...presentationXml.matchAll(/<p:sldId\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1]));
  return ids.length > 0 ? Math.max(...ids) : 255;
}

/** presentation.xml.rels 안 Id="rIdN" 중 가장 큰 N. */
function maxRelId(relsXml: string): number {
  const ids = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
  return ids.length > 0 ? Math.max(...ids) : 0;
}

/** 원본 sldIdLst의 첫 슬라이드가 참조하는 슬라이드 레이아웃 파트(예: 'slideLayouts/slideLayout1.xml',
 * presentation.xml.rels의 Target 값 그대로) — 빈 슬라이드가 항상 이 레이아웃을 쓴다. */
function firstSlideLayoutTarget(
  zipEntries: Record<string, Uint8Array>,
  firstSlidePart: string,
): string {
  const slideRelsXml = strFromU8(zipEntries[`ppt/slides/_rels/${firstSlidePart}.rels`]);
  const match = /Target="([^"]*slideLayout\d+\.xml)"/.exec(slideRelsXml);
  if (!match) throw new Error(`슬라이드 레이아웃 참조를 찾을 수 없습니다: ${firstSlidePart}`);
  // 슬라이드 자신의 rels는 '../slideLayouts/slideLayoutN.xml'처럼 상대경로 — 새 슬라이드도 같은
  // 깊이(ppt/slides/)에 있으므로 그대로 재사용 가능하다.
  return match[1];
}

/** 슬라이드 파트 이름의 유일하게 허용되는 형태 — 정확히 'slideN.xml'(N은 숫자)만 통과시킨다.
 * 클라이언트(또는 그 사이 어떤 버그)가 만든 SlideEntry.sourceSlidePart를 그대로 XML 문자열
 * 접합(Relationship Target=, Override PartName= 등)에 쓰기 전에 반드시 이 검증을 거쳐야 한다 —
 * 그렇지 않으면 `slide1.xml" TargetMode="External" Target="http://evil` 같은 값이 저장되는
 * pptx의 실제 관계/콘텐츠 타입 XML에 임의 속성을 주입할 수 있다(패키지 인젝션). */
const SLIDE_PART_RE = /^slide\d+\.xml$/;

/**
 * 저장 가능한 항목만 남긴다: 'blank'는 항상 통과, 'original'/'duplicate'는 sourceSlidePart가
 * 정확히 'slideN.xml' 형태이고 실제로 zip 안에 그 파트가 존재할 때만 통과시킨다. 그 외
 * (kind가 알 수 없거나, 형식이 이상하거나, 참조하는 원본 파트가 이미 사라진 경우 — 예: 직전
 * 저장에서 원본이 삭제됐는데 그 이전에 만든 duplicate 항목이 여전히 그 원본을 가리키는 stale
 * 상태)는 조용히 걸러낸다(throw하지 않는다) — 그런 상태는 "더 이상 유효하지 않은 편집 의도"로
 * 보고 안전하게 무시하는 게, 전체 저장을 실패시켜 다음 자동저장마다 계속 크래시하는 것보다 낫다.
 */
function filterValidEntries(
  rawEntryList: readonly SlideEntry[],
  zipEntries: Record<string, Uint8Array>,
): SlideEntry[] {
  return rawEntryList.filter((entry) => {
    if (entry.kind === 'blank') return true;
    if (entry.kind !== 'original' && entry.kind !== 'duplicate') return false;
    return (
      typeof entry.sourceSlidePart === 'string' &&
      SLIDE_PART_RE.test(entry.sourceSlidePart) &&
      zipEntries[`ppt/slides/${entry.sourceSlidePart}`] !== undefined
    );
  });
}

/**
 * slideContents(PPTX_SLIDE_CONTENTS_MAP)의 도형 하나 — 'meta'(SlideShapeMeta, 도형 위치/크기)와
 * 'content'(Y.XmlFragment, ProseMirror 문서)를 키로 갖는다. Yjs Y.Map은 값 타입을 하나만 선언할
 * 수 있어(pptxShared.ts의 SlideShapeMeta 주석과 같은 이유) 값 타입을 느슨하게 unknown으로 둔다.
 */
export type ShapeContentMap = Y.Map<unknown>;
/** 슬라이드 하나 안의 모든 도형: shapeId -> ShapeContentMap. */
export type PerSlideContentsMap = Y.Map<ShapeContentMap>;
/** PPTX_SLIDE_CONTENTS_MAP의 최상위 타입: slideEntryId -> PerSlideContentsMap. */
export type SlideContentsMap = Y.Map<PerSlideContentsMap>;

/**
 * 슬라이드 하나(entry)의 텍스트박스를 원본 XML에서 파싱해 slideContents에 채운다(지연 시딩,
 * room.ts의 seedSlideContentIfNeeded가 호출). 이미 채워져 있으면 아무 것도 하지 않는다 — 호출부가
 * 먼저 has() 체크를 하지만, 두 클라이언트가 같은 슬라이드를 동시에 처음 여는 경쟁을 대비해 여기서도
 * 한 번 더 확인한다(이 함수는 통째로 동기이므로 이 두 번째 체크와 실제 slideContents.set(...) 사이에
 * 다른 코드가 끼어들 수 없다 — JS 단일 스레드).
 *
 * **attach 순서**: 아래는 perSlide/shapeMap/fragment를 먼저 부모(이미 Y.Doc에 연결된
 * slideContents)에 붙인 뒤에 내용을 채운다 — room.ts의 기존 docx 시딩
 * (`prosemirrorToYXmlFragment(doc, ydoc.getXmlFragment(XML_FRAGMENT_NAME))`, XML_FRAGMENT_NAME이라는
 * "이미 attach된" 최상위 fragment를 먼저 얻고 그 다음에 채우는 패턴)과 똑같은 순서다. 반대 순서
 * (자식을 전부 채운 detached 서브트리를 만든 뒤 마지막에 한 번에 attach)도 기능적으로는 똑같이
 * 동작하지만(prosemirrorToYXmlFragment는 `type.doc`가 없으면 `{ transact: (fn) => fn(undefined) }`
 * 대체 트랜잭션으로 처리한다), 이 룸의 `ydoc.on('update', ...)`가 매 트랜잭션마다 즉시 그 결과를
 * 인코딩해 브로드캐스트하는 실제 서버 환경에서 그 인코딩 단계가 방금 한 번에 attach된 깊은
 * 서브트리를 읽다가 Yjs의 "Invalid access: Add Yjs type to a document before reading data"
 * (premature access) 경고를 콘솔에 남기는 것을 구현 중 실제 room.ts 테스트로 확인했다(고립된
 * yjs 스크립트만으로는 재현되지 않았다 — 브로드캐스트 인코딩까지 거쳐야 드러난다). 치명적이지
 * 않지만(값은 항상 정확했다) 이미 검증된 docx의 "먼저 attach, 그 다음 채우기" 순서를 그대로
 * 따르면 이 경고 자체가 나지 않는다. 전체를 doc.transact(...)로 한 번 더 감싸는 이유는 도형이
 * 여러 개여도 피어에게는 여전히 하나의 원자적 업데이트로만 보이게 하기 위함이다(각 .set()이
 * 개별 트랜잭션이 되어 도형별로 쪼개진 업데이트가 나가는 것을 막는다).
 */
export function seedSlideContentFromXml(
  slideContents: SlideContentsMap,
  slideEntryId: string,
  pptxBytes: Uint8Array,
  sourceSlidePart: string,
): void {
  if (slideContents.has(slideEntryId)) return;
  const zipEntries = unzipSync(pptxBytes);
  const slideXml = strFromU8(zipEntries[`ppt/slides/${sourceSlidePart}`]);
  const shapes = parseSlideShapes(slideXml);

  // slideContents는 항상 ydoc.getMap(PPTX_SLIDE_CONTENTS_MAP)로 얻은, 이미 Y.Doc에 연결된 최상위
  // 공유 타입이다(AbstractType은 attach되는 즉시 .doc가 채워진다) — 그래서 .doc!는 항상 안전하다.
  slideContents.doc!.transact(() => {
    const perSlide: PerSlideContentsMap = new Y.Map();
    slideContents.set(slideEntryId, perSlide); // 먼저 attach
    for (const shape of shapes) {
      const shapeMap: ShapeContentMap = new Y.Map();
      perSlide.set(shape.shapeId, shapeMap); // 먼저 attach
      const meta: SlideShapeMeta = { shapeId: shape.shapeId, x: shape.x, y: shape.y, w: shape.w, h: shape.h };
      shapeMap.set('meta', meta);
      const fragment = new Y.XmlFragment();
      shapeMap.set('content', fragment); // 먼저 attach
      prosemirrorToYXmlFragment(shape.doc, fragment); // 그 다음 채우기 — docx 시딩과 같은 순서
    }
  });
}

/**
 * 복제된 슬라이드의 콘텐츠를 소스 슬라이드의 "지금까지의 편집 상태"에서 깊은 복사한다(원본 XML을
 * 다시 파싱하지 않는다 — 소스가 이미 편집되어 있었다면 그 편집을 그대로 이어받아야 한다는 뜻).
 * 소스가 아직 시딩되지 않아 slideContents에 없으면 아무 것도 하지 않는다(호출부인
 * room.ts의 seedSlideContentIfNeeded가 이 함수를 부르기 전에 항상 소스를 먼저 seed해 둔다).
 *
 * Y.Map/Y.XmlFragment 자체에는 깊은 복사 메서드가 없으므로, 도형별로 새 Y.Map/Y.XmlFragment를
 * 만들고 그 안의 자식(Y.XmlElement/Y.XmlText, 둘 다 Yjs 공식 clone() 지원)만 하나씩 옮겨 붙인다.
 * attach 순서는 seedSlideContentFromXml과 같은 이유로 "먼저 부모에 붙이고 그 다음 채우기"를
 * 따른다(전체를 doc.transact(...)로 감싸 피어에게는 하나의 원자적 업데이트로만 보이게 한다).
 */
export function duplicateSlideContent(
  slideContents: SlideContentsMap,
  sourceEntryId: string,
  newEntryId: string,
): void {
  const source = slideContents.get(sourceEntryId);
  if (!source) return;
  slideContents.doc!.transact(() => {
    const perSlide: PerSlideContentsMap = new Y.Map();
    slideContents.set(newEntryId, perSlide); // 먼저 attach
    for (const [shapeId, shapeMap] of source.entries()) {
      const newShapeMap: ShapeContentMap = new Y.Map();
      perSlide.set(shapeId, newShapeMap); // 먼저 attach
      newShapeMap.set('meta', shapeMap.get('meta'));
      const newFragment = new Y.XmlFragment();
      newShapeMap.set('content', newFragment); // 먼저 attach
      const sourceFragment = shapeMap.get('content') as Y.XmlFragment;
      for (const child of sourceFragment.toArray()) {
        if (!(child instanceof Y.XmlElement) && !(child instanceof Y.XmlText)) {
          // pptxTextSchema는 XmlHook을 만들 일이 없다(리프 노드가 없다) — 혹시라도 나타나면 조용히
          // 건너뛰지 않고 바로 던진다: 원본 유실 금지 원칙상 복제본에서 자식 하나가 조용히 빠지는
          // 것보다는 복제 자체가 눈에 띄게 실패하는 편이 낫다.
          throw new Error(`duplicateSlideContent: 지원하지 않는 Y.Xml 자식 타입입니다(도형 ${shapeId}).`);
        }
        newFragment.push([child.clone()]); // 그 다음 채우기 — 위와 같은 순서
      }
    }
  });
}

/** exportPptxFromSlides의 결과 — 저장할 바이트뿐 아니라, 그 바이트가 실제로 담고 있는 최종
 * 슬라이드 파트 순서도 함께 돌려준다. room.ts의 flushToStore는 이 finalParts로 Yjs 배열 자체를
 * "방금 디스크에 쓴 내용"과 다시 일치시켜, 두 번째 저장이 stale한 duplicate/sourcePageNumber
 * 때문에 깨지지 않게 한다. */
export interface ExportedPptx {
  readonly bytes: Uint8Array;
  readonly finalParts: readonly string[];
  /** finalParts[i]를 만들어낸 SlideEntry의 id(저장 전/후 동일 — flushToStore는 이 id를 그대로
   * 재사용해 새 SlideEntry를 만든다). flushToStore는 이 목록을 "이번 저장에서 살아남은 id
   * 전체"로도 써서, slideContents 중 여기 없는 키(삭제되거나 filterValidEntries가 걸러낸
   * 슬라이드)를 쓰레기로 간주해 정리한다. finalParts와 항상 같은 길이, 같은 순서로 1:1 대응한다. */
  readonly survivingEntryIds: readonly string[];
}

/**
 * 현재 슬라이드 배열 상태 + 원본 pptx bytes로 새 pptx를 조립한다
 * (room.ts의 flushToStore/getCurrentBytes에서 호출).
 */
export function exportPptxFromSlides(
  pptxBytes: Uint8Array,
  slidesArray: Y.Array<SlideEntry>,
  slideContents?: SlideContentsMap,
): ExportedPptx {
  const zipEntries = { ...unzipSync(pptxBytes) };
  const rawEntryList = slidesArray.toArray();
  const entryList = filterValidEntries(rawEntryList, zipEntries);
  if (entryList.length === 0) {
    throw new Error('슬라이드가 하나도 남지 않아 저장할 수 없습니다.');
  }
  const survivingEntryIds = entryList.map((entry) => entry.id);

  const originalRelsXml = strFromU8(zipEntries['ppt/_rels/presentation.xml.rels']);
  const originalPresentationXml = strFromU8(zipEntries['ppt/presentation.xml']);

  // 원본 슬라이드 파트 → 원본 rId, 원본 rId → 원본 sldId 매핑. 살아남는 'original' 슬라이드는 이
  // 값들을 그대로 재사용해야 한다 — 그래야 p14:sectionLst(슬라이드 정렬 "구역")나
  // p:custShowLst(사용자 지정 쇼)처럼 rId/sldId를 참조하는 기능이 내보내기마다 깨지지 않는다.
  const originalPartToRelId = new Map<string, string>(
    [...slideRelsMap(originalRelsXml)].map(([rId, part]) => [part, rId]),
  );
  const originalSldIdByRelId = new Map<string, number>();
  for (const tag of originalPresentationXml.match(/<p:sldId\b[^>]*\/>/g) ?? []) {
    const idMatch = /\bid="(\d+)"/.exec(tag);
    const relIdMatch = /r:id="(rId\d+)"/.exec(tag);
    if (idMatch && relIdMatch) originalSldIdByRelId.set(relIdMatch[1], Number(idMatch[1]));
  }

  let nextSlideNumber = maxSlideNumber(zipEntries) + 1;
  let nextRelId = maxRelId(originalRelsXml) + 1;
  let nextSldId = maxSldId(originalPresentationXml) + 1;

  const firstOriginalPart = entryList.find((e) => e.sourceSlidePart)?.sourceSlidePart;
  const layoutTarget = firstOriginalPart ? firstSlideLayoutTarget(zipEntries, firstOriginalPart) : null;

  // 최종 순서대로 (파트 이름, presentation.xml.rels용 rId) 목록을 만들면서, 필요한 zip 엔트리를
  // 채워 넣는다. 'original'은 기존 파트를 그대로 두고(이름도 안 바꿈), 'duplicate'/'blank'만
  // 새 번호를 할당해 zipEntries에 새로 추가한다.
  const finalParts: string[] = [];

  for (const entry of entryList) {
    if (entry.kind === 'original' && entry.sourceSlidePart) {
      finalParts.push(entry.sourceSlidePart);
      continue;
    }
    if (entry.kind === 'duplicate' && entry.sourceSlidePart) {
      const newPart = `slide${nextSlideNumber}.xml`;
      nextSlideNumber += 1;
      zipEntries[`ppt/slides/${newPart}`] = zipEntries[`ppt/slides/${entry.sourceSlidePart}`];
      zipEntries[`ppt/slides/_rels/${newPart}.rels`] =
        zipEntries[`ppt/slides/_rels/${entry.sourceSlidePart}.rels`];
      finalParts.push(newPart);
      continue;
    }
    // kind === 'blank'
    if (!layoutTarget) throw new Error('빈 슬라이드를 추가하려면 원본에 슬라이드가 최소 1장 있어야 합니다.');
    const newPart = `slide${nextSlideNumber}.xml`;
    nextSlideNumber += 1;
    zipEntries[`ppt/slides/${newPart}`] = strToU8(BLANK_SLIDE_XML);
    zipEntries[`ppt/slides/_rels/${newPart}.rels`] = strToU8(blankSlideRelsXml(layoutTarget));
    finalParts.push(newPart);
  }

  // 원본에는 있었지만 최종 목록에 없는(=삭제된) 슬라이드 파트를 zip에서 제거한다. finalParts에는
  // 이미 새로 만든 duplicate/blank 파트도 포함되어 있으므로, 여기 없다는 건 원래 있던 슬라이드가
  // 배열에서 완전히 빠졌다(삭제됐다)는 뜻뿐이다.
  for (const key of Object.keys(zipEntries)) {
    const match = /^ppt\/slides\/(slide\d+\.xml)$/.exec(key);
    if (match && !finalParts.includes(match[1])) {
      delete zipEntries[key];
      delete zipEntries[`ppt/slides/_rels/${match[1]}.rels`];
    }
  }

  // presentation.xml.rels용 rId 배정: 원본에 이미 있던 파트(=살아남은 'original' 슬라이드)는 원래
  // rId를 그대로 재사용하고, 새로 생긴 파트(duplicate/blank)만 새 rId를 할당한다.
  const partToRelId = new Map<string, string>();
  for (const part of finalParts) {
    if (partToRelId.has(part)) continue;
    const originalRelId = originalPartToRelId.get(part);
    if (originalRelId) {
      partToRelId.set(part, originalRelId);
    } else {
      partToRelId.set(part, `rId${nextRelId}`);
      nextRelId += 1;
    }
  }

  // presentation.xml.rels: 기존 슬라이드 관계 태그만 "타겟 제거"하고(Target이 slideN.xml인 것만
  // 정규식으로 지움 — 다른 관계 태그는 self-closing이 아니든, Target에 '>'가 들어있든 원본 문자열
  // 그대로 절대 건드리지 않는다), 최종 슬라이드 목록에 대한 관계 태그를 그 뒤에 새로 덧붙인다.
  // Content_Types의 Override 처리와 같은 패턴 — 파일 전체를 매치된 태그들로부터 재조립하지 않는다.
  const slideRelRemovalRe = /<Relationship\b[^>]*Target="(?:[^"]*\/)?slide\d+\.xml"[^>]*\/>/g;
  const slideRelTags = [...partToRelId.entries()].map(
    ([part, rId]) =>
      `<Relationship Id="${rId}" Type="${SLIDE_RELATIONSHIP_TYPE}" Target="slides/${part}"/>`,
  );
  const newRelsXml = originalRelsXml
    .replace(slideRelRemovalRe, '')
    .replace('</Relationships>', `${slideRelTags.join('')}</Relationships>`);
  zipEntries['ppt/_rels/presentation.xml.rels'] = strToU8(newRelsXml);

  // presentation.xml: sldIdLst를 최종 순서대로 재작성한다. 살아남은 'original' 슬라이드는 원래
  // sldId 숫자를 재사용하고(그 rId가 원본에도 있었다면), 새로 생긴 파트만 새 sldId를 할당한다.
  const sldIdTags = finalParts.map((part) => {
    const rId = partToRelId.get(part)!;
    let sldId = originalSldIdByRelId.get(rId);
    if (sldId === undefined) {
      sldId = nextSldId;
      nextSldId += 1;
    }
    return `<p:sldId id="${sldId}" r:id="${rId}"/>`;
  });
  const presentationXml = originalPresentationXml.replace(
    /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
    `<p:sldIdLst>${sldIdTags.join('')}</p:sldIdLst>`,
  );
  zipEntries['ppt/presentation.xml'] = strToU8(presentationXml);

  // [Content_Types].xml: 삭제된 슬라이드의 Override는 지우고, 새로 추가된(duplicate/blank) 슬라이드의
  // Override는 더한다. original 슬라이드는 이미 있던 Override를 그대로 둔다.
  let contentTypesXml = strFromU8(zipEntries['[Content_Types].xml']);
  const overrideRe = /<Override PartName="\/ppt\/slides\/(slide\d+\.xml)"[^>]*\/>/g;
  contentTypesXml = contentTypesXml.replace(overrideRe, (whole, part: string) =>
    finalParts.includes(part) ? whole : '',
  );
  const missingOverrides = finalParts
    .filter((part) => !contentTypesXml.includes(`/ppt/slides/${part}"`))
    .map((part) => `<Override PartName="/ppt/slides/${part}" ContentType="${SLIDE_CONTENT_TYPE}"/>`);
  contentTypesXml = contentTypesXml.replace('</Types>', `${missingOverrides.join('')}</Types>`);
  zipEntries['[Content_Types].xml'] = strToU8(contentTypesXml);

  // 슬라이드 "내용"(텍스트박스) 편집을 각 파트의 XML에 병합한다. entryList[i]/finalParts[i]는
  // 위에서 같은 for-of 한 번으로 1:1 대응하게 만들어졌다(같은 길이, 같은 순서) — survivingEntryIds와
  // 같은 불변식. 한 번도 안 연(=slideContents에 없는) 슬라이드는 지연 시딩이 아직 안 됐다는
  // 뜻이므로 원본 파트를 전혀 건드리지 않고 그대로 둔다(대부분의 슬라이드가 여기 해당한다).
  // slideContents 인자 자체가 없으면(콘텐츠 편집을 다루지 않는 기존 호출부/테스트) 아래 루프는
  // 통째로 건너뛰어 기존 동작과 100% 동일하다.
  if (slideContents) {
    for (let i = 0; i < entryList.length; i += 1) {
      const perSlide = slideContents.get(entryList[i].id);
      // room.ts의 flushToStore는 slideContents 값이 구조적으로 Y.Map이 아니면(예: 악의적이거나
      // 버그 있는 클라이언트가 원시 값을 써 넣은 경우) throw하지 않고 조용히 정리(prune)한다 —
      // 여기서도 같은 원칙을 따른다: 애초에 우리 시딩 코드(seedSlideContentFromXml/
      // duplicateSlideContent)가 만든 적 없는 구조는 "유효한 콘텐츠가 아예 없었다"로 취급해
      // 건너뛴다(= 병합할 것이 없으니 원본 파트를 그대로 둔다). 아래에서 던지는 throw는 반대로
      // "구조는 멀쩡한데 실제 병합(spliceShapeTxBody/appendNewShape/직렬화)이 실패한" 경우만을
      // 위한 것이다 — 그 경우는 진짜 사용자 편집이 유실될 수 있는 상황이므로 반드시 크게 실패해야
      // 한다("원본 유실 금지"). 이 둘을 구분하지 않으면 어느 한쪽이 깨진다: 여기서 전부 throw하면
      // Task 5가 이미 보장한 "손상된 슬라이드 하나가 전체 자동저장을 영원히 막지 않는다"가
      // 깨지고, 반대로 전부 조용히 건너뛰면 진짜 편집 유실이 조용히 파일에 반영된다.
      if (!(perSlide instanceof Y.Map)) continue; // 한 번도 안 열었거나(undefined) 구조가 잘못됨 — 원본 그대로 둔다
      const partPath = `ppt/slides/${finalParts[i]}`;
      let slideXml = strFromU8(zipEntries[partPath]);
      for (const [shapeId, shapeMap] of perSlide.entries()) {
        if (!(shapeMap instanceof Y.Map)) continue; // 위와 같은 이유 — 이 도형 항목만 구조가 잘못됨
        const fragment = shapeMap.get('content');
        if (!(fragment instanceof Y.XmlFragment)) continue; // 위와 같은 이유 — 진짜 콘텐츠가 아니었음
        const doc = yXmlFragmentToProseMirrorRootNode(fragment, pptxTextSchema);
        const txBodyXml = serializeTxBodyXml(doc);
        // 'new-' 접두어는 클라이언트가 새로 추가한(원본에 없던) 텍스트박스의 shapeId 관례다 —
        // spliceShapeTxBody는 `id="\d+"`만 매치하므로 이런 id를 넘기면 찾지 못해 throw한다(그게
        // 맞는 동작이다 — 아래 분기가 절대 그 경로를 타지 않게 해야 한다). 그 외의 모든 shapeId는
        // 원본에서 파싱된(parseSlideShapes) 실제 숫자 id이므로 spliceShapeTxBody로 기존
        // <p:txBody>를 교체한다.
        if (shapeId.startsWith('new-')) {
          const meta = shapeMap.get('meta') as SlideShapeMeta;
          slideXml = appendNewShape(
            slideXml,
            { shapeId, x: meta.x ?? 0, y: meta.y ?? 0, w: meta.w ?? 1000000, h: meta.h ?? 500000 },
            txBodyXml,
          );
        } else {
          slideXml = spliceShapeTxBody(slideXml, shapeId, txBodyXml);
        }
      }
      zipEntries[partPath] = strToU8(slideXml);
    }
  }

  return { bytes: zipSync(zipEntries), finalParts, survivingEntryIds };
}
