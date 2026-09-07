# pptx 슬라이드 구성 편집 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** pptx 파일의 슬라이드를 복사·삭제·순서변경하고 빈 슬라이드를 추가할 수 있는 Yjs 실시간
동시편집 기능을 새 전용 편집기(`editor-slides.html`)로 제공한다.

**Architecture:** 슬라이드 내용(도형·텍스트·이미지)은 절대 파싱/재생성하지 않는다 — pptx zip
안에서 슬라이드 순서·존재 여부만 나타내는 `Y.Array<SlideEntry>` 하나를 Yjs로 동기화하고, 저장
시점에만 원본 zip의 슬라이드 파트(`ppt/slides/slideN.xml` 등)를 그대로 복제/삭제/재배치해서
새 pptx를 조립한다. docx의 `literalizeDocxThemeFonts`와 동일한 "zip을 열어 필요한 XML
텍스트만 정규식으로 고치고 다시 잠그기" 패턴.

**Tech Stack:** fflate(zip), yjs, TypeScript, Fastify, esbuild(클라이언트 번들), vitest.

**Spec:** `docs/superpowers/specs/2026-09-07-pptx-slide-editor-design.md`

## Global Constraints

- 슬라이드 내부 콘텐츠(도형·텍스트·이미지·애니메이션)는 절대 파싱·재생성하지 않는다 — 항상
  원본 zip 파트를 바이트 그대로 복제/이동한다.
- 지원 작업은 정확히 4가지: 복사, 삭제, 순서 변경, 빈 슬라이드 추가. 그 외(레이아웃/테마 변경,
  발표자 노트 편집, 다른 파일에서 슬라이드 가져오기)는 이번 범위가 아니다.
- docx/xlsx와 동일하게 Yjs 실시간 동시편집(`src/collab/room.ts`)을 통해 저장된다 — REST 기반
  단발성 저장 API를 새로 만들지 않는다.
- 새 외부 npm 의존성을 추가하지 않는다 — fflate(이미 의존성에 있음)만으로 구현한다.
- 빈 슬라이드가 참조할 슬라이드 레이아웃은 항상 **원본 `sldIdLst`의 첫 번째 슬라이드가 쓰는
  레이아웃**으로 고정한다.

---

### Task 1: pptx 협업 데이터 모델 + 테스트용 샘플 pptx 픽스처

**Files:**
- Create: `src/collab/pptxShared.ts`
- Modify: `tests/helpers.ts`
- Test: `tests/pptxShared.test.ts`

**Interfaces:**
- Produces: `SlideEntryKind`, `SlideEntry` 인터페이스, `PPTX_SLIDES_ARRAY` 상수 (이후 모든
  Task가 이 타입/상수를 그대로 가져다 쓴다). `makeSamplePptx(slideTexts: readonly string[]): Buffer`
  (Task 2, 3, 4의 모든 테스트가 이 픽스처를 쓴다).

- [ ] **Step 1: `src/collab/pptxShared.ts` 작성**

```ts
// src/collab/pptxShared.ts
/**
 * pptx 동시편집이 쓰는 Y.Doc 최상위 공유 타입 이름 — 서버(room.ts, src/pptx/collab.ts)와
 * 클라이언트 번들(src/client/editor-slides.ts) 양쪽에서 그대로 가져다 쓴다. 슬라이드 "내용"은
 * 절대 Yjs에 담지 않는다 — 이 배열은 오직 "어떤 슬라이드가 몇 번째냐"만 나타낸다.
 */
export type SlideEntryKind = 'original' | 'duplicate' | 'blank';

export interface SlideEntry {
  readonly id: string; // 이 배열 항목의 고유 UUID (Yjs 식별용, 슬라이드 파일 번호와 무관)
  readonly kind: SlideEntryKind;
  /** kind가 'original' | 'duplicate'일 때만 존재 — 원본 zip의 실제 슬라이드 파트 이름
   *  (예: 'slide3.xml', 확장자 포함, 'ppt/slides/' 접두어 없음). 'duplicate'도 원본과 같은
   *  sourceSlidePart를 가리킨다(내보낼 때 그 파트를 복제해서 새 번호를 붙인다). */
  readonly sourceSlidePart: string | null;
  /** kind가 'original' | 'duplicate'일 때만 존재 — 업로드 시점에 만들어 둔 "바로보기"용 PDF에서
   *  이 슬라이드가 몇 번째 페이지였는지(1-based). 썸네일을 그릴 때 `/api/view/pages/:page`에
   *  그대로 쓴다 — 편집 중에는 이 PDF를 다시 만들지 않으므로 항상 유효한 페이지 번호다. */
  readonly sourcePageNumber: number | null;
}

export const PPTX_SLIDES_ARRAY = 'slides';
```

- [ ] **Step 2: `tests/helpers.ts`에 `makeSamplePptx` 추가**

기존 `makeSampleDocx`와 동일한 스타일(fflate `zipSync`로 최소 유효 구조를 직접 조립)로,
슬라이드 N장짜리 pptx를 만든다. 각 슬라이드는 서로 다른 텍스트를 담아 테스트에서
"이 결과물이 어느 원본 슬라이드였는지"를 구분할 수 있게 한다. 파일 맨 끝에 추가:

```ts
const PPTX_CONTENT_TYPES_HEAD =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/ppt/presentation.xml" ' +
  'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
  '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ' +
  'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
  '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ' +
  'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>';

const PPTX_ROOT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ' +
  'Target="ppt/presentation.xml"/>' +
  '</Relationships>';

const PPTX_SLIDE_MASTER_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/></p:spTree></p:cSld>' +
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" ' +
  'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" ' +
  'folHlink="folHlink"/>' +
  '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
  '</p:sldMaster>';

const PPTX_SLIDE_MASTER_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" ' +
  'Target="../slideLayouts/slideLayout1.xml"/>' +
  '</Relationships>';

const PPTX_SLIDE_LAYOUT_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>';

const PPTX_SLIDE_LAYOUT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" ' +
  'Target="../slideMasters/slideMaster1.xml"/>' +
  '</Relationships>';

function pptxSlideXml(text: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
    '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr/>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    `<p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>` +
    '</p:spTree></p:cSld></p:sld>'
  );
}

const PPTX_SLIDE_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" ' +
  'Target="../slideLayouts/slideLayout1.xml"/>' +
  '</Relationships>';

/**
 * 최소한의 유효한 pptx를 직접 zip으로 조립한다(PowerPoint/LibreOffice 없이).
 * slideTexts.length장의 슬라이드를 만들고, 전부 같은 slideMaster/slideLayout 하나를 공유한다.
 */
export function makeSamplePptx(slideTexts: readonly string[] = ['Slide 1']): Buffer {
  const slideOverrides = slideTexts
    .map(
      (_, i) =>
        `<Override PartName="/ppt/slides/slide${i + 1}.xml" ` +
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>',
    )
    .join('');
  const contentTypesXml = `${PPTX_CONTENT_TYPES_HEAD}${slideOverrides}</Types>`;

  const sldIdLst = slideTexts
    .map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`)
    .join('');
  const presentationXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
    '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
    `<p:sldIdLst>${sldIdLst}</p:sldIdLst>` +
    '<p:sldSz cx="9144000" cy="6858000"/>' +
    '</p:presentation>';

  const presentationRelsEntries = [
    '<Relationship Id="rId1" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" ' +
      'Target="slideMasters/slideMaster1.xml"/>',
    ...slideTexts.map(
      (_, i) =>
        `<Relationship Id="rId${i + 2}" ` +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" ' +
        `Target="slides/slide${i + 1}.xml"/>`,
    ),
  ].join('');
  const presentationRelsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `${presentationRelsEntries}</Relationships>`;

  const entries: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(contentTypesXml),
    '_rels/.rels': strToU8(PPTX_ROOT_RELS),
    'ppt/presentation.xml': strToU8(presentationXml),
    'ppt/_rels/presentation.xml.rels': strToU8(presentationRelsXml),
    'ppt/slideMasters/slideMaster1.xml': strToU8(PPTX_SLIDE_MASTER_XML),
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': strToU8(PPTX_SLIDE_MASTER_RELS),
    'ppt/slideLayouts/slideLayout1.xml': strToU8(PPTX_SLIDE_LAYOUT_XML),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': strToU8(PPTX_SLIDE_LAYOUT_RELS),
  };
  slideTexts.forEach((text, i) => {
    entries[`ppt/slides/slide${i + 1}.xml`] = strToU8(pptxSlideXml(text));
    entries[`ppt/slides/_rels/slide${i + 1}.xml.rels`] = strToU8(PPTX_SLIDE_RELS);
  });

  return Buffer.from(zipSync(entries));
}
```

- [ ] **Step 3: 실패하는 테스트 작성 — `tests/pptxShared.test.ts`**

```ts
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import { makeSamplePptx } from './helpers.js';

describe('makeSamplePptx', () => {
  test('creates a valid zip with the requested number of slides', () => {
    const bytes = makeSamplePptx(['First', 'Second', 'Third']);
    const entries = unzipSync(bytes);
    expect(entries['ppt/slides/slide1.xml']).toBeDefined();
    expect(entries['ppt/slides/slide2.xml']).toBeDefined();
    expect(entries['ppt/slides/slide3.xml']).toBeDefined();
    expect(strFromU8(entries['ppt/slides/slide1.xml'])).toContain('First');
    expect(strFromU8(entries['ppt/slides/slide2.xml'])).toContain('Second');
  });

  test('presentation.xml sldIdLst lists slides in order', () => {
    const bytes = makeSamplePptx(['A', 'B']);
    const entries = unzipSync(bytes);
    const presentationXml = strFromU8(entries['ppt/presentation.xml']);
    const rIds = [...presentationXml.matchAll(/r:id="(rId\d+)"/g)].map((m) => m[1]);
    expect(rIds).toEqual(['rId2', 'rId3']); // rId1은 slideMaster
  });
});
```

- [ ] **Step 4: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/pptxShared.test.ts`
Expected: FAIL — `makeSamplePptx` is not exported / module not found (아직 helpers.ts에 추가 전)

- [ ] **Step 5: Step 1-2의 코드를 실제로 파일에 반영**

`src/collab/pptxShared.ts`를 새로 만들고, `tests/helpers.ts` 맨 위 import에
`strToU8, zipSync`는 이미 있으니 그대로 두고 파일 끝에 Step 2 코드를 추가한다.

- [ ] **Step 6: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/pptxShared.test.ts`
Expected: PASS (2 tests)

- [ ] **Step 7: 커밋**

```bash
git add src/collab/pptxShared.ts tests/helpers.ts tests/pptxShared.test.ts
git commit -m "feat: pptx 슬라이드 편집 데이터 모델 + 테스트용 샘플 pptx 픽스처"
```

---

### Task 2: pptx → SlideEntry[] 시딩 (`seedSlidesFromPptx`)

**Files:**
- Create: `src/pptx/collab.ts`
- Test: `tests/pptxCollab.test.ts`

**Interfaces:**
- Consumes: `SlideEntry`, `SlideEntryKind`, `PPTX_SLIDES_ARRAY` (Task 1), `makeSamplePptx` (Task 1).
- Produces: `seedSlidesFromPptx(pptxBytes: Uint8Array, slidesArray: Y.Array<SlideEntry>): void`
  (Task 4의 `room.ts` seedRoom이 그대로 호출).

원본 pptx를 열어 `ppt/presentation.xml`의 `sldIdLst` 순서대로 `ppt/_rels/presentation.xml.rels`를
통해 실제 슬라이드 파트 이름을 찾고, 그 순서 그대로 `SlideEntry[]`(전부 `kind:'original'`,
`sourcePageNumber`는 1부터 순서대로)를 만들어 배열에 채운다.

- [ ] **Step 1: 실패하는 테스트 작성 — `tests/pptxCollab.test.ts`**

```ts
import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { PPTX_SLIDES_ARRAY, type SlideEntry } from '../src/collab/pptxShared.js';
import { seedSlidesFromPptx } from '../src/pptx/collab.js';
import { makeSamplePptx } from './helpers.js';

describe('seedSlidesFromPptx', () => {
  test('populates one original entry per slide, in presentation order', () => {
    const bytes = makeSamplePptx(['First', 'Second', 'Third']);
    const ydoc = new Y.Doc();
    const slidesArray = ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);

    seedSlidesFromPptx(bytes, slidesArray);

    expect(slidesArray.length).toBe(3);
    const entries = slidesArray.toArray();
    expect(entries.map((e) => e.kind)).toEqual(['original', 'original', 'original']);
    expect(entries.map((e) => e.sourceSlidePart)).toEqual(['slide1.xml', 'slide2.xml', 'slide3.xml']);
    expect(entries.map((e) => e.sourcePageNumber)).toEqual([1, 2, 3]);
    expect(entries.every((e) => typeof e.id === 'string' && e.id.length > 0)).toBe(true);
  });
});
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/pptxCollab.test.ts`
Expected: FAIL — `src/pptx/collab.js`를 찾을 수 없음

- [ ] **Step 3: `src/pptx/collab.ts` 구현**

```ts
// src/pptx/collab.ts
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
```

- [ ] **Step 4: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/pptxCollab.test.ts`
Expected: PASS (1 test)

- [ ] **Step 5: 커밋**

```bash
git add src/pptx/collab.ts tests/pptxCollab.test.ts
git commit -m "feat: pptx 원본에서 슬라이드 순서 배열을 시딩하는 seedSlidesFromPptx"
```

---

### Task 3: SlideEntry[] → pptx 내보내기 (`exportPptxFromSlides`)

**Files:**
- Modify: `src/pptx/collab.ts`
- Test: `tests/pptxCollab.test.ts`

**Interfaces:**
- Consumes: Task 2의 `seedSlidesFromPptx`, `SlideEntry` (Task 1).
- Produces: `exportPptxFromSlides(pptxBytes: Uint8Array, slidesArray: Y.Array<SlideEntry>): Uint8Array`
  (Task 4의 `room.ts` flushToStore/getCurrentBytes가 그대로 호출).

가장 복잡한 부분 — 현재 배열 상태(순서·복사·삭제·빈 슬라이드)를 원본 zip에 반영해 새 pptx를
만든다. 슬라이드 내용 파트(`ppt/slides/slideN.xml`, 그 rels)는 **절대 파싱하지 않고 항상
바이트 그대로 복제**한다.

- [ ] **Step 1: 실패하는 테스트 4개 작성 — `tests/pptxCollab.test.ts`에 추가**

```ts
import { strFromU8, unzipSync } from 'fflate';
import { exportPptxFromSlides, seedSlidesFromPptx } from '../src/pptx/collab.js';

function slidesArrayFrom(ydoc: Y.Doc): Y.Array<SlideEntry> {
  return ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);
}

function slideOrderOf(pptxBytes: Uint8Array): string[] {
  const entries = unzipSync(pptxBytes);
  const presentationXml = strFromU8(entries['ppt/presentation.xml']);
  const presentationRelsXml = strFromU8(entries['ppt/_rels/presentation.xml.rels']);
  const relMap = new Map<string, string>();
  for (const relTag of presentationRelsXml.match(/<Relationship\b[^>]*\/>/g) ?? []) {
    const idMatch = /Id="(rId\d+)"/.exec(relTag);
    const targetMatch = /Target="(?:[^"]*\/)?(slide\d+\.xml)"/.exec(relTag);
    if (idMatch && targetMatch) relMap.set(idMatch[1], targetMatch[1]);
  }
  const rIds = [...presentationXml.matchAll(/<p:sldId\b[^>]*r:id="(rId\d+)"/g)].map((m) => m[1]);
  return rIds.map((rId) => relMap.get(rId)!);
}

describe('exportPptxFromSlides', () => {
  test('reordering the array reorders the exported presentation', () => {
    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    seedSlidesFromPptx(makeSamplePptx(['A', 'B', 'C']), slidesArray);
    const [first] = slidesArray.slice(0, 1);
    ydoc.transact(() => {
      slidesArray.delete(0, 1);
      slidesArray.insert(2, [first]);
    });

    const original = makeSamplePptx(['A', 'B', 'C']);
    const out = exportPptxFromSlides(original, slidesArray);
    const entries = unzipSync(out);
    expect(strFromU8(entries['ppt/slides/slide1.xml'])).toContain('A'); // 원본 파트 그대로
    expect(slideOrderOf(out)).toEqual(['slide2.xml', 'slide3.xml', 'slide1.xml']);
  });

  test('deleting a slide removes its part and shrinks the presentation', () => {
    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    seedSlidesFromPptx(makeSamplePptx(['A', 'B', 'C']), slidesArray);
    slidesArray.delete(1, 1); // "B" 삭제

    const original = makeSamplePptx(['A', 'B', 'C']);
    const out = exportPptxFromSlides(original, slidesArray);
    const entries = unzipSync(out);
    expect(slideOrderOf(out)).toEqual(['slide1.xml', 'slide3.xml']);
    expect(entries['ppt/slides/slide2.xml']).toBeUndefined();
    expect(strFromU8(entries['[Content_Types].xml'])).not.toContain('/ppt/slides/slide2.xml');
    expect(strFromU8(entries['ppt/_rels/presentation.xml.rels'])).not.toContain('slide2.xml');
  });

  test('duplicating a slide copies its content byte-for-byte under a new part name', () => {
    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    seedSlidesFromPptx(makeSamplePptx(['A', 'B']), slidesArray);
    const source = slidesArray.get(0);
    slidesArray.insert(1, [{ id: 'dup-1', kind: 'duplicate', sourceSlidePart: source.sourceSlidePart, sourcePageNumber: null }]);

    const original = makeSamplePptx(['A', 'B']);
    const out = exportPptxFromSlides(original, slidesArray);
    const entries = unzipSync(out);
    expect(slideOrderOf(out)).toEqual(['slide1.xml', 'slide3.xml', 'slide2.xml']);
    expect(strFromU8(entries['ppt/slides/slide3.xml'])).toContain('A'); // "A" 슬라이드가 그대로 복제됨
    expect(entries['ppt/slides/_rels/slide3.xml.rels']).toBeDefined();
  });

  test('adding a blank slide inserts a minimal valid slide referencing the first layout', () => {
    const ydoc = new Y.Doc();
    const slidesArray = slidesArrayFrom(ydoc);
    seedSlidesFromPptx(makeSamplePptx(['A', 'B']), slidesArray);
    slidesArray.insert(2, [{ id: 'blank-1', kind: 'blank', sourceSlidePart: null, sourcePageNumber: null }]);

    const original = makeSamplePptx(['A', 'B']);
    const out = exportPptxFromSlides(original, slidesArray);
    const entries = unzipSync(out);
    expect(slideOrderOf(out)).toEqual(['slide1.xml', 'slide2.xml', 'slide3.xml']);
    expect(entries['ppt/slides/slide3.xml']).toBeDefined();
    expect(entries['ppt/slides/_rels/slide3.xml.rels']).toBeDefined();
    expect(strFromU8(entries['ppt/slides/_rels/slide3.xml.rels'])).toContain('slideLayout1.xml');
    expect(strFromU8(entries['[Content_Types].xml'])).toContain('/ppt/slides/slide3.xml');
  });
});
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/pptxCollab.test.ts`
Expected: FAIL — `exportPptxFromSlides is not defined`

- [ ] **Step 3: `src/pptx/collab.ts`에 `exportPptxFromSlides` 구현 추가**

```ts
import { strToU8, unzipSync, zipSync } from 'fflate';
// (unzipSync, strFromU8은 이미 import되어 있음 — zipSync, strToU8 추가)

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

/**
 * 현재 슬라이드 배열 상태 + 원본 pptx bytes로 새 pptx를 조립한다
 * (room.ts의 flushToStore/getCurrentBytes에서 호출).
 */
export function exportPptxFromSlides(pptxBytes: Uint8Array, slidesArray: Y.Array<SlideEntry>): Uint8Array {
  const zipEntries = { ...unzipSync(pptxBytes) };
  const entryList = slidesArray.toArray();

  let nextSlideNumber = maxSlideNumber(zipEntries) + 1;
  let nextRelId = maxRelId(strFromU8(zipEntries['ppt/_rels/presentation.xml.rels'])) + 1;
  let nextSldId = maxSldId(strFromU8(zipEntries['ppt/presentation.xml'])) + 1;

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

  // presentation.xml.rels: 최종 파트 목록 전체에 대해 rId를 다시 배정한다(살아남은 기존 rId도
  // 재사용하지 않고 전부 새로 씀 — 중복/충돌 걱정 없이 항상 일관된 결과를 만들기 위해서다).
  const partToRelId = new Map<string, string>();
  for (const part of finalParts) {
    if (!partToRelId.has(part)) {
      partToRelId.set(part, `rId${nextRelId}`);
      nextRelId += 1;
    }
  }
  const otherRelsXml = strFromU8(zipEntries['ppt/_rels/presentation.xml.rels']);
  const nonSlideRels = (otherRelsXml.match(/<Relationship\b[^>]*\/>/g) ?? []).filter(
    (tag) => !/Target="(?:[^"]*\/)?slide\d+\.xml"/.test(tag),
  );
  const slideRelTags = [...partToRelId.entries()].map(
    ([part, rId]) =>
      `<Relationship Id="${rId}" Type="${SLIDE_RELATIONSHIP_TYPE}" Target="slides/${part}"/>`,
  );
  zipEntries['ppt/_rels/presentation.xml.rels'] =
    strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `${nonSlideRels.join('')}${slideRelTags.join('')}</Relationships>`);

  // presentation.xml: sldIdLst를 최종 순서대로 재작성한다.
  const sldIdTags = finalParts.map((part, index) => {
    const rId = partToRelId.get(part)!;
    return `<p:sldId id="${nextSldId + index}" r:id="${rId}"/>`;
  });
  const presentationXml = strFromU8(zipEntries['ppt/presentation.xml']).replace(
    /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
    `<p:sldIdLst>${sldIdTags.join('')}</p:sldIdLst>`,
  );
  zipEntries['ppt/presentation.xml'] = strToU8(presentationXml);

  // [Content_Types].xml: 삭제된 슬라이드의 Override는 지우고, 새로 추가된(duplicate/blank) 슬라이드의
  // Override는 더한다. original 슬라이드는 이미 있던 Override를 그대로 둔다.
  let contentTypesXml = strFromU8(zipEntries['[Content_Types].xml']);
  const overrideRe = /<Override PartName="\/ppt\/slides\/(slide\d+\.xml)"[^/]*\/>/g;
  contentTypesXml = contentTypesXml.replace(overrideRe, (whole, part: string) =>
    finalParts.includes(part) ? whole : '',
  );
  const missingOverrides = finalParts
    .filter((part) => !contentTypesXml.includes(`/ppt/slides/${part}"`))
    .map((part) => `<Override PartName="/ppt/slides/${part}" ContentType="${SLIDE_CONTENT_TYPE}"/>`);
  contentTypesXml = contentTypesXml.replace('</Types>', `${missingOverrides.join('')}</Types>`);
  zipEntries['[Content_Types].xml'] = strToU8(contentTypesXml);

  return zipSync(zipEntries);
}
```

**중요 — Step 3 코드 작성 시 구현자에게:** `maxSlideNumber`/`maxSldId`/`maxRelId`/
`firstSlideLayoutTarget`/`blankSlideRelsXml`/`BLANK_SLIDE_XML`/`SLIDE_CONTENT_TYPE`/
`SLIDE_RELATIONSHIP_TYPE`는 전부 `src/pptx/collab.ts` 안의 새 private 함수/상수이고,
`seedSlidesFromPptx`가 이미 정의한 `SLIDE_TARGET`은 재사용하지 않는다(정규식 캡처 그룹
용도가 달라 별도로 둔다). 기존 `unzipSync`, `strFromU8` import에 `zipSync`, `strToU8`을
추가한다.

- [ ] **Step 4: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/pptxCollab.test.ts`
Expected: PASS (5 tests: Task 2의 1개 + 이번 4개)

- [ ] **Step 5: 커밋**

```bash
git add src/pptx/collab.ts tests/pptxCollab.test.ts
git commit -m "feat: 슬라이드 배열 상태로 pptx를 재조립하는 exportPptxFromSlides"
```

---

### Task 4: 서버 통합 — room.ts / edit.ts / documents.ts

**Files:**
- Modify: `src/collab/room.ts:58,60-62,102-117,123-167,193-206`
- Modify: `src/routes/edit.ts:70`
- Modify: `src/routes/documents.ts:43-50`
- Test: `tests/routes-documents.test.ts` (신규 — `EDITABLE_EXTS`/`editorPathFor` 단위 테스트,
  아래 Step 9 참고. room.ts 자체의 pptx 분기는 Task 6의 end-to-end 검증에서 확인한다)

**Interfaces:**
- Consumes: `seedSlidesFromPptx`, `exportPptxFromSlides` (Task 2, 3), `PPTX_SLIDES_ARRAY`,
  `SlideEntry` (Task 1).
- Produces: docx/xlsx와 동일하게 pptx도 `EDITABLE_EXTS`에 포함되고 `/editor-slides.html`로
  라우팅됨. WebSocket `/ws/edit/:docId`가 pptx 문서를 받아들임.

- [ ] **Step 1: `src/routes/documents.ts` 수정**

```ts
// 43-50번째 줄 근처, 기존:
const GRID_VIEW_EXTS = new Set(['xlsx', 'csv']);
const EDITABLE_EXTS = new Set(['docx', 'xlsx']);
function editorPathFor(ext: string): string {
  return ext === 'xlsx' ? '/editor-grid.html' : '/editor.html';
}

// 아래로 교체 — Step 9에서 tests/routes-documents.test.ts가 직접 import해 검증할 수 있도록
// 둘 다 export를 붙인다(지금까지는 파일 스코프 전용이었다):
const GRID_VIEW_EXTS = new Set(['xlsx', 'csv']);
/** 실시간 동시편집(에디터)을 지원하는 확장자 — docx는 문서 편집기, xlsx는 그리드 편집기,
 * pptx는 슬라이드 구성 편집기로 연결된다 */
export const EDITABLE_EXTS = new Set(['docx', 'xlsx', 'pptx']);
export function editorPathFor(ext: string): string {
  if (ext === 'xlsx') return '/editor-grid.html';
  if (ext === 'pptx') return '/editor-slides.html';
  return '/editor.html';
}
```

- [ ] **Step 2: `src/routes/edit.ts` 수정**

```ts
// 70번째 줄 근처, 기존:
if (meta.ext !== 'docx' && meta.ext !== 'xlsx') {

// 아래로 교체:
if (meta.ext !== 'docx' && meta.ext !== 'xlsx' && meta.ext !== 'pptx') {
```
(에러 메시지 `'이 문서 형식은 동시편집을 지원하지 않습니다.'`와 주석은 그대로 둔다.)

- [ ] **Step 3: `src/collab/room.ts` 수정 — import + DocKind + kindOfExt**

```ts
// import 블록에 추가 (32번째 줄 parseXlsxToGrid import 아래):
import { exportPptxFromSlides, seedSlidesFromPptx } from '../pptx/collab.js';
import { PPTX_SLIDES_ARRAY, type SlideEntry } from './pptxShared.js';

// 58번째 줄:
type DocKind = 'docx' | 'xlsx' | 'pptx';

// 60-62번째 줄:
function kindOfExt(ext: string): DocKind {
  if (ext === 'xlsx') return 'xlsx';
  if (ext === 'pptx') return 'pptx';
  return 'docx';
}
```

- [ ] **Step 4: `seedRoom`에 pptx 분기 추가 (102-117번째 줄)**

```ts
async function seedRoom(docId: string, kind: DocKind, ydoc: Y.Doc): Promise<void> {
  try {
    const meta = await store.readMeta(docId);
    if (!meta) return;
    const original = await store.readOriginal(docId, meta.ext);
    if (kind === 'xlsx') {
      const grid = await parseXlsxToGrid(original);
      seedGridFromXlsx(ydoc, grid);
    } else if (kind === 'pptx') {
      seedSlidesFromPptx(original, ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY));
    } else {
      const { doc } = importDocx(original);
      prosemirrorToYXmlFragment(doc, ydoc.getXmlFragment(XML_FRAGMENT_NAME));
    }
  } catch (error) {
    console.error(`동시편집 세션 초기화 실패 (문서 ${docId}):`, error);
  }
}
```

- [ ] **Step 5: `flushToStore`에 pptx 분기 추가 (123-167번째 줄)**

xlsx 분기와 docx 분기 사이에 pptx 분기를 끼워 넣는다(if/else if/else 3갈래로):

```ts
async function flushToStore(docId: string, kind: DocKind, ydoc: Y.Doc): Promise<void> {
  const meta = await store.readMeta(docId);
  if (!meta) return;

  if (kind === 'xlsx') {
    // ...기존 xlsx 분기 그대로...
  } else if (kind === 'pptx') {
    const original = await store.readOriginal(docId, meta.ext);
    const newPptxBytes = exportPptxFromSlides(original, ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY));
    await store.saveOriginal(docId, meta.ext, Buffer.from(newPptxBytes));

    const converter = findConverter(meta.ext);
    if (converter) {
      try {
        const { pageCount } = await convertToPdf(meta, converter);
        await store.writeMeta({ ...meta, pageCount, status: 'ready', viewKind: 'pages', progress: 100 });
      } catch (error) {
        console.error(`동시편집 저장 후 PDF 재변환 실패 (문서 ${docId}):`, error);
      }
    }
  } else {
    // ...기존 docx 분기 그대로...
  }
  invalidateRenderCache(docId);
  invalidateViewCaches(docId);
  await recordAudit({ type: 'edit_saved', docId });
  ydoc.getMap<number>(DOC_META_MAP).set(DOC_META_SAVED_AT_KEY, Date.now());
}
```

- [ ] **Step 6: `getCurrentBytes`에 pptx 분기 추가 (193-206번째 줄)**

```ts
export async function getCurrentBytes(docId: string): Promise<Uint8Array | null> {
  const room = rooms.get(docId);
  if (!room) return null;
  await room.ready;
  const meta = await store.readMeta(docId);
  if (!meta) return null;
  const original = await store.readOriginal(docId, meta.ext);
  if (room.kind === 'xlsx') {
    return exportXlsxFromGrid(original, room.ydoc);
  }
  if (room.kind === 'pptx') {
    return exportPptxFromSlides(original, room.ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY));
  }
  const fragment = room.ydoc.getXmlFragment(XML_FRAGMENT_NAME);
  const pmDoc = yXmlFragmentToProseMirrorRootNode(fragment, docxSchema);
  return exportDocx(original, pmDoc);
}
```

**참고 — 이 Task의 테스트 범위:** `tests/collab.test.ts`의 기존 테스트(docx ↔ Yjs ↔
ProseMirror 브리지)는 `room.ts`의 WebSocket/룸 생명주기를 거치지 않고 변환 함수만 직접
호출하는 방식이다 — `Room`/`seedRoom`/`flushToStore`는 모듈 내부에 갇혀 있어(export 없음)
그 방식대로 pptx용 테스트를 추가하려면 room.ts의 공개 API를 바꿔야 한다(이번 범위 아님).
`seedSlidesFromPptx`/`exportPptxFromSlides` 자체의 정확성은 이미 Task 2-3에서 충분히
테스트했으므로, 이 Task에서는 **아래 Step 9(라우팅 단위 테스트)**로 wiring 코드가 올바른
분기를 타는지만 확인하고, `room.ts`를 실제 WebSocket으로 거치는 전체 흐름은 Task 6의
end-to-end 검증에서 확인한다.

- [ ] **Step 9: `EDITABLE_EXTS`/`editorPathFor` 단위 테스트**

Step 1에서 `export`를 붙였으므로, 새 테스트 파일 `tests/routes-documents.test.ts`에서 직접
import해 확인한다:

```ts
import { describe, expect, test } from 'vitest';
import { EDITABLE_EXTS, editorPathFor } from '../src/routes/documents.js';

describe('pptx editing routing', () => {
  test('pptx is editable', () => {
    expect(EDITABLE_EXTS.has('pptx')).toBe(true);
  });
  test('pptx routes to editor-slides.html', () => {
    expect(editorPathFor('pptx')).toBe('/editor-slides.html');
  });
});
```

- [ ] **Step 10: 전체 테스트 + 타입체크**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 전부 PASS

- [ ] **Step 11: 커밋**

```bash
git add src/collab/room.ts src/routes/edit.ts src/routes/documents.ts tests/routes-documents.test.ts
git commit -m "feat: pptx를 동시편집 룸/라우팅에 연결"
```

---

### Task 5: 클라이언트 편집기 — `editor-slides.html` + `editor-slides.ts`

**Files:**
- Create: `public/editor-slides.html`
- Create: `src/client/editor-slides.ts`
- Modify: `scripts/build-editor.mjs`

**Interfaces:**
- Consumes: `PPTX_SLIDES_ARRAY`, `SlideEntry` (Task 1). 서버 `/api/view/info`, `/ws/edit/:docId`
  (기존, 변경 없음 — Task 4가 이미 pptx를 통과시키게 해 둠). `/api/view/pages/:page` (기존 PDF
  페이지 렌더링 API, 그대로 재사용).
- Produces: 없음(최종 사용자 화면).

`public/editor-grid.html` + `src/client/grid-editor.ts`의 뼈대(토큰 파싱, `/api/view/info`
호출, WebsocketProvider 연결, presence 렌더링)를 그대로 재사용하되, 본문만 그리드 대신
슬라이드 썸네일 카드 그리드로 바꾼다.

- [ ] **Step 1: `public/editor-slides.html` 작성**

`public/editor-grid.html`의 `<head>`와 상단 툴바(1-42번째 줄, `toolbar-top-row` 전체)를
그대로 복사하되 `<title>`을 `DocView — 슬라이드 편집`으로, `body class`를
`editor-body slides-editor-page`로 바꾼다. `toolbar-format`(리본) 자리에는 버튼 3개만 둔다:

```html
    <div class="toolbar-group toolbar-format">
      <button id="btn-add-blank" class="btn-pill" title="빈 슬라이드 추가">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14M5 12h14"/></svg>
        빈 슬라이드
      </button>
    </div>
  </header>

  <main id="slides-grid" class="slides-editor-grid" aria-label="슬라이드 목록"></main>

  <template id="slide-card-template">
    <div class="slide-card" draggable="true">
      <div class="slide-card-thumb">
        <img alt="" />
        <span class="slide-card-number"></span>
      </div>
      <div class="slide-card-actions">
        <button class="icon-btn slide-card-duplicate" title="복사">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="8" y="8" width="12" height="13" rx="1.3"/><path d="M15 8V5.5A1.5 1.5 0 0 0 13.5 4H5.5A1.5 1.5 0 0 0 4 5.5v11A1.5 1.5 0 0 0 5.5 18H8"/></svg>
        </button>
        <button class="icon-btn slide-card-delete" title="삭제">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h16"/><path d="M10 11v6M14 11v6"/><path d="M6 7l1 13a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-13"/><path d="M9 7V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v3"/></svg>
        </button>
      </div>
    </div>
  </template>

  <div id="error-screen" class="error-screen" hidden>
    <div class="error-card">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M12 8v4"/><path d="M12 16h.01"/></svg>
      <p id="error-message"></p>
      <a href="/" class="btn-pill btn-pill-primary">문서 목록으로</a>
    </div>
  </div>

  <script type="module" src="/js/editor-slides.bundle.js"></script>
</body>
</html>
```

빈 슬라이드 카드는 `<img>` 대신 고정 플레이스홀더를 쓴다 — `public/img/blank-slide.svg`를
새로 만든다(회색 배경에 "+"만 있는 단순 SVG, 4:3 비율 `viewBox="0 0 400 300"`):

```svg
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300">
  <rect width="400" height="300" fill="#f3f4f7"/>
  <rect x="0.5" y="0.5" width="399" height="299" fill="none" stroke="#d7dae0"/>
  <path d="M170 150h60M200 120v60" stroke="#b6bcc7" stroke-width="6" stroke-linecap="round"/>
</svg>
```

- [ ] **Step 2: `src/client/editor-slides.ts` 작성**

`grid-editor.ts`의 `main()`(토큰 파싱 → `/api/view/info` → `WebsocketProvider` 연결 →
presence)과 완전히 같은 앞부분을 쓰고, 본문 렌더링만 슬라이드 그리드로 바꾼다:

```ts
import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import { PPTX_SLIDES_ARRAY, type SlideEntry } from '../collab/pptxShared.js';
import { pickPresenceColorForName, renderPresenceAvatars } from './presence.js';

interface Permissions {
  readonly canPrint: boolean;
  readonly canDownload: boolean;
  readonly canCopy: boolean;
  readonly canEdit: boolean;
}

interface ViewInfo {
  readonly docId: string;
  readonly name: string;
  readonly viewer: string;
  readonly permissions: Permissions;
}

const token = new URLSearchParams(location.search).get('token');
const titleEl = document.getElementById('doc-title')!;
const connStatusEl = document.getElementById('conn-status')!;
const gridEl = document.getElementById('slides-grid') as HTMLElement;
const cardTemplate = document.getElementById('slide-card-template') as HTMLTemplateElement;
const presenceListEl = document.getElementById('presence-list') as HTMLElement;

function showFatalError(message: string): void {
  document.getElementById('error-message')!.textContent = message;
  document.getElementById('error-screen')!.hidden = false;
}

async function fetchViewInfo(): Promise<ViewInfo> {
  const res = await fetch(`/api/view/info?token=${encodeURIComponent(token ?? '')}`);
  const json = (await res.json()) as { success: boolean; data?: ViewInfo; error?: string };
  if (!res.ok || !json.success || !json.data) throw new Error(json.error ?? '문서 정보를 불러오지 못했습니다.');
  return json.data;
}

let canEdit = false;
let slidesArray: Y.Array<SlideEntry>;
let dragFromIndex: number | null = null;

function cardFor(entry: SlideEntry, index: number): HTMLElement {
  const fragment = cardTemplate.content.cloneNode(true) as DocumentFragment;
  const card = fragment.querySelector('.slide-card') as HTMLElement;
  card.dataset.index = String(index);
  card.dataset.id = entry.id;
  const img = card.querySelector('img') as HTMLImageElement;
  img.src =
    entry.kind === 'blank'
      ? '/img/blank-slide.svg'
      : `/api/view/pages/${entry.sourcePageNumber}?token=${encodeURIComponent(token ?? '')}`;
  (card.querySelector('.slide-card-number') as HTMLElement).textContent = String(index + 1);
  if (!canEdit) {
    card.querySelector('.slide-card-actions')?.remove();
    card.draggable = false;
  }
  return card;
}

function renderSlides(): void {
  const entries = slidesArray.toArray();
  gridEl.replaceChildren(...entries.map((entry, index) => cardFor(entry, index)));
}

function setupDragAndDrop(): void {
  gridEl.addEventListener('dragstart', (event) => {
    const card = (event.target as HTMLElement).closest<HTMLElement>('.slide-card');
    if (!card) return;
    dragFromIndex = Number(card.dataset.index);
    event.dataTransfer?.setData('text/plain', card.dataset.id ?? '');
  });
  gridEl.addEventListener('dragover', (event) => {
    if (dragFromIndex !== null) event.preventDefault();
  });
  gridEl.addEventListener('drop', (event) => {
    event.preventDefault();
    const target = (event.target as HTMLElement).closest<HTMLElement>('.slide-card');
    if (!target || dragFromIndex === null) return;
    const toIndex = Number(target.dataset.index);
    if (toIndex === dragFromIndex) return;
    const ydoc = slidesArray.doc!;
    ydoc.transact(() => {
      const [moved] = slidesArray.slice(dragFromIndex!, dragFromIndex! + 1);
      slidesArray.delete(dragFromIndex!, 1);
      // 삭제 후 배열이 한 칸 줄어들므로, 드롭 대상이 삭제 지점보다 뒤에 있었다면 그 대상의
      // 실제(삭제 후) 위치는 원래 toIndex보다 1 작다 — 항상 "대상 카드 바로 앞"에 꽂히도록
      // 두 경우 모두 대상의 삭제-후 위치에 삽입한다.
      const insertAt = toIndex > dragFromIndex! ? toIndex - 1 : toIndex;
      slidesArray.insert(insertAt, [moved]);
    });
    dragFromIndex = null;
  });
}

function setupCardActions(): void {
  gridEl.addEventListener('click', (event) => {
    if (!canEdit) return;
    const target = event.target as HTMLElement;
    const card = target.closest<HTMLElement>('.slide-card');
    if (!card) return;
    const index = Number(card.dataset.index);
    const entry = slidesArray.get(index);
    if (target.closest('.slide-card-duplicate')) {
      const ydoc = slidesArray.doc!;
      ydoc.transact(() => {
        slidesArray.insert(index + 1, [
          {
            id: crypto.randomUUID(),
            kind: 'duplicate',
            sourceSlidePart: entry.sourceSlidePart,
            sourcePageNumber: null,
          },
        ]);
      });
    } else if (target.closest('.slide-card-delete')) {
      slidesArray.delete(index, 1);
    }
  });
  document.getElementById('btn-add-blank')?.addEventListener('click', () => {
    if (!canEdit) return;
    slidesArray.insert(slidesArray.length, [
      { id: crypto.randomUUID(), kind: 'blank', sourceSlidePart: null, sourcePageNumber: null },
    ]);
  });
}

async function main(): Promise<void> {
  if (!token) {
    showFatalError('열람 토큰이 없습니다.');
    return;
  }
  let info: ViewInfo;
  try {
    info = await fetchViewInfo();
  } catch (error) {
    showFatalError(error instanceof Error ? error.message : '문서 정보를 불러오지 못했습니다.');
    return;
  }
  if (!info.permissions.canEdit) {
    showFatalError('이 링크에는 동시편집 권한이 없습니다. 뷰어에서 열람만 가능합니다.');
    return;
  }

  titleEl.textContent = info.name;
  document.title = `${info.name} — DocView 편집`;
  canEdit = info.permissions.canEdit;

  const ydoc = new Y.Doc();
  slidesArray = ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);
  const wsProtocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
  const provider = new WebsocketProvider(`${wsProtocol}//${location.host}/ws/edit`, info.docId, ydoc, {
    params: { token },
  });

  const myColor = pickPresenceColorForName(info.viewer);
  provider.awareness.setLocalStateField('user', { name: info.viewer, color: myColor });

  provider.on('status', ({ status }: { status: string }) => {
    connStatusEl.dataset.status = status;
    connStatusEl.textContent = status === 'connected' ? '연결됨' : '연결 중…';
  });
  provider.awareness.on('change', () => renderPresenceAvatars(presenceListEl, provider.awareness));

  slidesArray.observe(renderSlides);
  renderSlides();
  setupDragAndDrop();
  setupCardActions();
}

void main();
```

- [ ] **Step 3: `scripts/build-editor.mjs`에 세 번째 엔트리포인트 추가**

```js
await build({
  ...shared,
  entryPoints: ['src/client/editor-slides.ts'],
  outfile: 'public/js/editor-slides.bundle.js',
});
```

- [ ] **Step 4: 빌드 확인**

Run: `node scripts/build-editor.mjs`
Expected: 에러 없이 `public/js/editor-slides.bundle.js` 생성

- [ ] **Step 5: 타입체크**

Run: `npx tsc --noEmit`
Expected: 에러 없음

- [ ] **Step 6: 커밋**

```bash
git add public/editor-slides.html public/img/blank-slide.svg src/client/editor-slides.ts scripts/build-editor.mjs
git commit -m "feat: pptx 슬라이드 구성 편집 클라이언트(editor-slides.html)"
```

---

### Task 6: 실제 문서로 end-to-end 검증

**Files:**
- 코드 변경 없음 — 수동/Playwright 검증만.

**Interfaces:**
- Consumes: Task 1-5의 전체 결과물.

- [ ] **Step 1: 개발 서버 기동**

Run: `npm run dev`

- [ ] **Step 2: 실제 pptx 업로드 → 토큰 발급 → editor-slides.html 열기**

`POST /api/documents`로 실제 pptx 업로드, `POST /api/documents/:id/tokens`로
`canEdit:true` 토큰 발급, `editorUrl`(`/editor-slides.html?token=...`)을 Playwright로 열어
슬라이드 썸네일이 원본 슬라이드 수만큼 나오는지 스크린샷으로 확인.

- [ ] **Step 3: 복사/삭제/빈 슬라이드 추가/드래그 순서변경을 실제로 조작**

각 조작 후 화면이 즉시 반영되는지 확인. 두 번째 브라우저 컨텍스트로 같은 문서를 열어 한쪽의
조작이 다른 쪽에도 실시간 반영되는지 확인(Yjs 동시편집 확인).

- [ ] **Step 4: 저장 후 결과 확인**

두 세션을 모두 닫아(또는 8초 대기해) `flushToStore`가 실행되게 한 뒤, `viewer.html`로 같은
문서를 열어 "바로보기"가 편집한 순서/장수를 정확히 반영하는지 확인. 다운로드한 pptx를 실제
PowerPoint/LibreOffice Impress로 직접 열어 깨짐 없이 표시되는지 확인(가능하면).

- [ ] **Step 5: 전체 테스트 스위트 + 타입체크 최종 확인**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 전부 PASS, Task 1-4에서 추가한 테스트 전부 포함

- [ ] **Step 6: 임시 테스트 문서/스크립트 정리**

업로드 테스트에 쓴 문서는 `DELETE /api/documents/:id`로 정리하고, 진단용으로 만든 임시
스크립트가 있다면 삭제한다.
