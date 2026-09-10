import { randomUUID } from 'node:crypto';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import type * as Y from 'yjs';
import type { SlideEntry } from '../collab/pptxShared.js';

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
export function exportPptxFromSlides(pptxBytes: Uint8Array, slidesArray: Y.Array<SlideEntry>): ExportedPptx {
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

  return { bytes: zipSync(zipEntries), finalParts, survivingEntryIds };
}
