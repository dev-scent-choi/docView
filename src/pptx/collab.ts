import { randomUUID } from 'node:crypto';
import { strFromU8, unzipSync } from 'fflate';
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
