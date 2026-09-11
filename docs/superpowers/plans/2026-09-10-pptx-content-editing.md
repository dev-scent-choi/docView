# pptx 텍스트박스 콘텐츠 동시편집 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** pptx 편집기의 오른쪽 콘텐츠 영역에서 기존 텍스트박스 글자를 실시간 동시편집할 수
있게 하고, 왼쪽은 지금의 슬라이드 관리(추가/복제/삭제/순서바꾸기)를 세로 레일로 재배치한다.

**Architecture:** 슬라이드 "내용"(텍스트박스별 ProseMirror 문서)을 새 Yjs 공유 타입
`slideContents`에 담아, 이 저장소가 이미 docx에서 쓰고 있는 `y-prosemirror`
(`ySyncPlugin`/`yCursorPlugin`/`yUndoPlugin`)로 글자 단위 CRDT 동기화한다. 슬라이드는
클릭해서 열 때만 원본 XML을 파싱해 지연 시딩한다(xlsx 지연 로딩과 같은 원칙). 저장은
편집된 텍스트박스만 원본 슬라이드 XML의 `<p:txBody>`를 교체하는 방식으로 원본을 보존한다.

**Tech Stack:** TypeScript, Yjs + y-websocket + y-prosemirror, prosemirror-model/state/view,
fast-xml-parser(이미 docx import가 씀), fflate(zip), Fastify, vitest.

**Spec:** `docs/superpowers/specs/2026-09-10-pptx-content-editing-design.md`

## Global Constraints

- **원본 유실 금지가 최우선 기준**(사용자 명시): 파싱·재조립 과정에서 어떤 이유로든 실패하면
  "조용히 잘못된 파일을 만드는" 대신 반드시 예외를 던지거나, 최소한 해당 슬라이드/도형은
  원본 그대로 남긴다. 절대 원본 zip 엔트리를 먼저 지우고 나중에 실패하는 순서로 만들지 않는다.
- **버전 관리는 이번 범위 밖**(사용자 명시, DB 기반으로 추후 별도 프로젝트) — 이 작업은
  "지금 이 순간의 문서"만 다룬다. 히스토리·되돌리기(문서 버전 단위)는 만들지 않는다
  (텍스트박스 안 타이핑 되돌리기(`yUndoPlugin`)는 이미 범위 안 — 이건 다른 것).
- 도형 이동·크기조절·회전, 새 도형·이미지 추가, 표·차트·스마트아트·애니메이션 편집은 이번
  범위에 없다 — 그런 UI 자체를 만들지 않는다(스펙의 "제외" 절 그대로).
- 슬라이드 레이아웃/마스터 상속 완전 해석은 범위 밖 — 위치·서식이 슬라이드 XML에 직접
  없으면 기본값(중앙 배치)을 쓴다(스펙의 "알려진 한계" 절 그대로).
- Node 서버 코드(`src/`)는 ESM(`.js` 확장자로 import), 클라이언트 번들은
  `scripts/build-editor.mjs`(esbuild)가 `src/client/*.ts`를 `public/js/*.bundle.js`로 묶는다
  — 클라이언트 파일을 고치면 `node scripts/build-editor.mjs`를 다시 돌려야 브라우저에 반영된다.

---

### Task 1: 데이터 모델 타입 + 텍스트 스키마 + WS 메시지 상수

**Files:**
- Modify: `src/collab/pptxShared.ts`
- Create: `src/pptx/textSchema.ts`
- Test: `tests/pptxTextSchema.test.ts`

**Interfaces:**
- Produces: `PPTX_SLIDE_CONTENTS_MAP: string`, `PPTX_SLIDE_SHAPE_META_MAP: string`,
  `MESSAGE_SLIDE_CONTENT_SUBSCRIBE: number`, `interface SlideShapeMeta { shapeId: string; x: number | null; y: number | null; w: number | null; h: number | null }`,
  `pptxTextSchema: Schema`(prosemirror-model)를 이후 모든 태스크가 그대로 가져다 쓴다.

- [ ] **Step 1: `pptxShared.ts`에 상수·타입 추가**

`src/collab/pptxShared.ts` 맨 아래(`export const PPTX_SLIDES_ARRAY = 'slides';` 다음)에
추가:

```ts
/**
 * slideContents.get(slideEntryId) -> Y.Map(shapeId -> Y.Map({ meta: SlideShapeMeta, content: Y.XmlFragment })).
 * 슬라이드 "내용"은 이제 여기 담긴다 — 위 PPTX_SLIDES_ARRAY 자체는 여전히 순서/포인터만
 * 나타낸다(이 파일 위쪽 주석은 "슬라이드 배열 자체"에는 내용을 안 담는다는 뜻으로 남겨 둔다).
 */
export const PPTX_SLIDE_CONTENTS_MAP = 'slideContents';

/** slideContents 안, 텍스트박스 하나의 위치 메타데이터. content(Y.XmlFragment)는 같은
 * 텍스트박스의 Y.Map 안에 'content' 키로 별도로 들어간다(이 인터페이스에는 없음 — Yjs
 * 공유 타입은 plain interface로 표현할 수 없어서 분리한다). */
export interface SlideShapeMeta {
  readonly shapeId: string;
  /** EMU(1pt = 12700 EMU) 단위, 원본 슬라이드 XML의 <a:off>/<a:ext>와 같은 단위.
   * null = 원본 XML에 명시적 위치가 없어(레이아웃 상속 등) 기본값을 썼다는 뜻. */
  readonly x: number | null;
  readonly y: number | null;
  readonly w: number | null;
  readonly h: number | null;
}

/** 새로 추가한(원본에 없던) 텍스트박스의 기본 위치/크기 — 슬라이드 중앙에 가로 60%, 세로 20%.
 * 표준 16:9 슬라이드 크기(12192000 x 6858000 EMU) 기준으로 잡되, 실제 슬라이드 크기가
 * 다르면 호출부가 비율로 다시 계산한다(이 상수는 "기본 비율"의 근거로만 쓴다). */
export const NEW_TEXTBOX_WIDTH_RATIO = 0.6;
export const NEW_TEXTBOX_HEIGHT_RATIO = 0.2;
```

`src/routes/edit.ts`가 메시지 kind로 스위치하는 부분(`room.ts`의 `MESSAGE_GRID_STRUCTURE_OP = 4`
바로 아래, 같은 로컬 상수 패턴 — 이 코드베이스는 메시지 kind 상수를 공유 모듈 하나로 묶지
않고 클라이언트/서버 양쪽에 각자 같은 숫자로 선언해 두는 기존 관례를 그대로 따른다)에
**둘 다** 추가:

```ts
const MESSAGE_SLIDE_CONTENT_SUBSCRIBE = 5;
```

(`src/collab/room.ts`의 `MESSAGE_GRID_STRUCTURE_OP = 4;` 바로 아래 한 줄, 그리고 나중
Task 8에서 만들 `src/client/editor-slides.ts`에도 같은 상수 — 지금은 room.ts 쪽만 추가해도
빌드는 깨지지 않는다, 클라이언트 쪽은 Task 8에서 추가).

- [ ] **Step 2: `src/pptx/textSchema.ts` 작성**

`src/docx/schema.ts`의 `marks`(strong/em/underline/textColor/fontSize)와 정확히 같은
정의를 재사용하되, 노드는 `doc`/`paragraph`/`text`만 남긴 훨씬 작은 스키마를 만든다(표·
목록·제목·이미지 없음 — pptx 텍스트박스는 문단의 나열일 뿐이라 필요 없다):

```ts
import { Schema, type MarkSpec, type NodeSpec } from 'prosemirror-model';

/**
 * pptx 텍스트박스 전용 최소 ProseMirror 스키마. docx의 docxSchema(src/docx/schema.ts)와
 * 마크 정의(강조 종류·속성)는 동일하게 맞춘다 — import/export 시 같은 pt/색 표현을 그대로
 * 재사용하기 위해서다. 다만 텍스트박스는 표·목록·제목이 필요 없어 노드는 doc/paragraph/text만
 * 남긴다(docxSchema 그대로 재사용하면 안 쓸 커맨드·키맵까지 끌고 들어온다).
 */
const nodes: Record<string, NodeSpec> = {
  doc: { content: 'paragraph+' },
  paragraph: {
    content: 'inline*',
    group: 'block',
    parseDOM: [{ tag: 'p' }],
    toDOM: () => ['p', 0],
  },
  text: { group: 'inline' },
};

const marks: Record<string, MarkSpec> = {
  strong: {
    parseDOM: [{ tag: 'strong' }, { tag: 'b' }],
    toDOM: () => ['strong', 0],
  },
  em: {
    parseDOM: [{ tag: 'em' }, { tag: 'i' }],
    toDOM: () => ['em', 0],
  },
  underline: {
    parseDOM: [{ tag: 'u' }],
    toDOM: () => ['u', 0],
  },
  textColor: {
    attrs: { color: {} },
    parseDOM: [{ style: 'color', getAttrs: (value) => ({ color: value as string }) }],
    toDOM: (mark) => ['span', { style: `color: ${mark.attrs.color as string}` }, 0],
  },
  fontSize: {
    /** 포인트(pt) 단위 — pptx의 <a:rPr sz="..">(100분의 1 포인트)와는 import/export 시 변환한다. */
    attrs: { size: {} },
    parseDOM: [{ style: 'font-size', getAttrs: (value) => ({ size: parseFloat(value as string) }) }],
    toDOM: (mark) => ['span', { style: `font-size: ${mark.attrs.size as number}pt` }, 0],
  },
};

export const pptxTextSchema = new Schema({ nodes, marks });
```

- [ ] **Step 3: 스키마가 실제로 문서를 만들 수 있는지 확인하는 테스트**

```ts
// tests/pptxTextSchema.test.ts
import { describe, expect, test } from 'vitest';
import { pptxTextSchema as schema } from '../src/pptx/textSchema.js';

describe('pptxTextSchema', () => {
  test('builds a paragraph with mixed bold/plain runs', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('Hello '),
        schema.text('World', [schema.marks.strong.create()]),
      ]),
    ]);
    expect(doc.textContent).toBe('Hello World');
    const boldRun = doc.child(0).child(1);
    expect(boldRun.marks.map((m) => m.type.name)).toEqual(['strong']);
  });

  test('textColor and fontSize marks carry their attrs', () => {
    const mark = schema.marks.textColor.create({ color: '#FF0000' });
    expect(mark.attrs.color).toBe('#FF0000');
    const sizeMark = schema.marks.fontSize.create({ size: 18 });
    expect(sizeMark.attrs.size).toBe(18);
  });
});
```

- [ ] **Step 4: 테스트 실행**

Run: `npx vitest run tests/pptxTextSchema.test.ts`
Expected: PASS (2개)

- [ ] **Step 5: 타입체크**

Run: `npx tsc --noEmit`
Expected: 클린

- [ ] **Step 6: 커밋**

```bash
git add src/collab/pptxShared.ts src/pptx/textSchema.ts tests/pptxTextSchema.test.ts
git commit -m "feat(pptx): add slide-content data model types and a minimal ProseMirror text schema"
```

---

### Task 2: Import — 슬라이드 XML의 텍스트박스를 ProseMirror 문서로 파싱

**Files:**
- Create: `src/pptx/slideContent.ts`
- Test: `tests/pptxSlideContent.test.ts`

**Interfaces:**
- Consumes: `pptxTextSchema`(Task 1).
- Produces: `interface ParsedShape { shapeId: string; x: number | null; y: number | null; w: number | null; h: number | null; doc: PMNode }`,
  `function parseSlideShapes(slideXml: string): ParsedShape[]` — 이후 Task 5(서버 시딩)가 그대로 쓴다.

- [ ] **Step 1: 실패하는 테스트 먼저 작성**

```ts
// tests/pptxSlideContent.test.ts
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
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `npx vitest run tests/pptxSlideContent.test.ts`
Expected: FAIL(`parseSlideShapes` 정의 안 됨)

- [ ] **Step 3: 구현**

```ts
// src/pptx/slideContent.ts
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
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `npx vitest run tests/pptxSlideContent.test.ts`
Expected: PASS (7개). 실패하면 `fast-xml-parser`의 실제 출력 구조(배열/객체 여부, 속성
접두어)를 `console.log(JSON.stringify(parser.parse(xml), null, 2))`로 직접 찍어 확인하고
파서 옵션이나 접근 경로를 맞춘다 — 이 스텝의 코드는 방향을 보여주는 것이지 한 글자도
안 틀렸다는 보장은 아니다.

- [ ] **Step 5: 타입체크 + 전체 테스트**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 클린, 전부 PASS

- [ ] **Step 6: 커밋**

```bash
git add src/pptx/slideContent.ts tests/pptxSlideContent.test.ts
git commit -m "feat(pptx): parse slide XML text boxes into ProseMirror documents"
```

---

### Task 3: Export — ProseMirror 문서를 슬라이드 XML로 되돌리기

**Files:**
- Modify: `src/pptx/slideContent.ts`
- Test: `tests/pptxSlideContent.test.ts`

**Interfaces:**
- Consumes: `ParsedShape`, `pptxTextSchema`(Task 1, 2).
- Produces: `function serializeTxBodyXml(doc: PMNode): string`,
  `function spliceShapeTxBody(slideXml: string, shapeId: string, txBodyXml: string): string`,
  `function appendNewShape(slideXml: string, meta: SlideShapeMeta, txBodyXml: string): string`
  — Task 6(서버 export 통합)이 그대로 쓴다.

- [ ] **Step 1: 라운드트립 테스트 먼저 작성(Task 2 파일에 이어서)**

`tests/pptxSlideContent.test.ts`에 추가:

```ts
import { appendNewShape, parseSlideShapes, serializeTxBodyXml, spliceShapeTxBody } from '../src/pptx/slideContent.js';
import { pptxTextSchema as schema } from '../src/pptx/textSchema.js';

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
});
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `npx vitest run tests/pptxSlideContent.test.ts`
Expected: FAIL(`serializeTxBodyXml`/`spliceShapeTxBody`/`appendNewShape` 정의 안 됨)

- [ ] **Step 3: 구현** (`src/pptx/slideContent.ts`에 이어서 추가)

```ts
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
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `npx vitest run tests/pptxSlideContent.test.ts`
Expected: PASS (11개 전부). `spliceShapeTxBody`의 정규식이 중첩된 `<p:sp>`(그룹 도형 안)를
잘못 매칭하면 실패할 수 있다 — 이번 범위(그룹 도형 미지원)에서는 최상위 `<p:sp>`만 다루므로
실패하면 테스트 픽스처가 그룹 도형을 쓰고 있는 게 아닌지부터 확인한다.

- [ ] **Step 5: 전체 테스트 + 타입체크**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 클린, 전부 PASS

- [ ] **Step 6: 커밋**

```bash
git add src/pptx/slideContent.ts tests/pptxSlideContent.test.ts
git commit -m "feat(pptx): serialize ProseMirror docs back into slide txBody XML"
```

---

### Task 4: `exportPptxFromSlides`가 살아남은 항목의 옛 id를 같이 돌려주게 확장

**이 태스크가 왜 중요한가**: `flushToStore`(Task 5)가 저장마다 모든 SlideEntry에 새
`randomUUID()`를 발급한다(기존 동작, 확인 완료). `slideContents`를 `SlideEntry.id`로
키를 잡는 이 설계 전체가, 저장 후 그 키를 새 id로 옮길 방법이 있어야만 유지된다 — 이
태스크 없이 뒤 태스크들을 진행하면 "저장 한 번 하면 방금 쓴 텍스트가 사라지는" 치명적
버그가 된다(스펙의 Global Constraints: 원본 유실 금지).

**Files:**
- Modify: `src/pptx/collab.ts`
- Test: `tests/pptxCollab.test.ts`

**Interfaces:**
- Produces: `interface ExportedPptx { bytes: Uint8Array; finalParts: readonly string[]; survivingEntryIds: readonly string[] }`
  (`survivingEntryIds[i]`는 `finalParts[i]`를 만들어낸 SlideEntry의 **저장 전** id) —
  Task 5의 `flushToStore` 재작성 블록이 그대로 쓴다.

- [ ] **Step 1: 실패하는 테스트 먼저 작성**

`tests/pptxCollab.test.ts`에 추가(기존 `exportPptxFromSlides` 테스트들 근처):

```ts
test('ExportedPptx.survivingEntryIds matches finalParts 1:1, in the same order, using pre-save ids', () => {
  const bytes = makeSamplePptx(['First', 'Second']);
  const ydoc = new Y.Doc();
  const slidesArray = slidesArrayFrom(ydoc);
  seedSlidesFromPptx(bytes, slidesArray);
  const idsBeforeExport = slidesArray.toArray().map((e) => e.id);

  const result = exportPptxFromSlides(bytes, slidesArray);

  expect(result.survivingEntryIds).toEqual(idsBeforeExport);
  expect(result.survivingEntryIds).toHaveLength(result.finalParts.length);
});

test('survivingEntryIds drops ids of entries filterValidEntries rejects (broken sourceSlidePart)', () => {
  const bytes = makeSamplePptx(['First', 'Second']);
  const ydoc = new Y.Doc();
  const slidesArray = slidesArrayFrom(ydoc);
  seedSlidesFromPptx(bytes, slidesArray);
  // 두 번째 항목을 존재하지 않는 슬라이드 파트를 가리키도록 깨뜨린다.
  const broken = { ...slidesArray.get(1), sourceSlidePart: 'slide999.xml' };
  ydoc.transact(() => {
    slidesArray.delete(1, 1);
    slidesArray.insert(1, [broken]);
  });
  const survivingIdBefore = slidesArray.get(0).id;

  const result = exportPptxFromSlides(bytes, slidesArray);

  expect(result.survivingEntryIds).toEqual([survivingIdBefore]); // 깨진 두 번째 항목은 빠짐
});
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `npx vitest run tests/pptxCollab.test.ts`
Expected: FAIL(`survivingEntryIds`가 `undefined`)

- [ ] **Step 3: 구현**

`src/pptx/collab.ts`의 `ExportedPptx` 인터페이스와 `exportPptxFromSlides` 함수를 찾아
수정한다(정확한 현재 줄 번호는 구현 시 `grep -n "interface ExportedPptx\|entryList ="
src/pptx/collab.ts`로 다시 확인 — 이 태스크 작성 시점 기준 `entryList`는 158번째 줄
근처에서 `filterValidEntries(rawEntryList, zipEntries)`로 만들어진다):

```ts
export interface ExportedPptx {
  readonly bytes: Uint8Array;
  readonly finalParts: readonly string[];
  /** finalParts[i]를 만들어낸 SlideEntry의 저장 전 id — flushToStore가 slideContents를
   * 새로 발급되는 id로 재키잉할 때 쓴다(옛 id -> 새 id 매핑의 "옛 id" 쪽). */
  readonly survivingEntryIds: readonly string[];
}

export function exportPptxFromSlides(pptxBytes: Uint8Array, slidesArray: Y.Array<SlideEntry>): ExportedPptx {
  const zipEntries = { ...unzipSync(pptxBytes) };
  const rawEntryList = slidesArray.toArray();
  const entryList = filterValidEntries(rawEntryList, zipEntries);
  if (entryList.length === 0) {
    throw new Error('슬라이드가 하나도 남지 않아 저장할 수 없습니다.');
  }
  const survivingEntryIds = entryList.map((e) => e.id); // <- 추가

  // ...(기존 로직 그대로, finalParts를 만드는 for 루프 등 안 바뀜)...

  return { bytes, finalParts, survivingEntryIds }; // <- 반환문에 survivingEntryIds 추가
}
```

기존 `exportPptxFromSlides` 호출부(`room.ts`의 `flushToStore`)는 Task 5에서 같이 고친다
— 지금은 반환 타입만 넓히므로 기존 호출부가 `survivingEntryIds`를 안 써도 컴파일은 깨지지
않는다(추가 필드는 구조적 타이핑상 안전).

- [ ] **Step 4: 테스트 통과 확인**

Run: `npx vitest run tests/pptxCollab.test.ts`
Expected: PASS 전부(기존 테스트 포함)

- [ ] **Step 5: 타입체크 + 전체 테스트**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 클린

- [ ] **Step 6: 커밋**

```bash
git add src/pptx/collab.ts tests/pptxCollab.test.ts
git commit -m "feat(pptx): exportPptxFromSlides also returns the pre-save id each surviving slide came from"
```

---

### Task 5: 서버 — `flushToStore`가 저장 시 `slideContents`를 새 id로 재키잉

**Files:**
- Modify: `src/collab/room.ts:1165-1218`(현재 pptx `flushToStore` 분기 — 실제 줄 번호는
  구현 시 `grep -n "kind === 'pptx'" src/collab/room.ts`로 재확인)
- Test: `tests/roomPptxContent.test.ts`(신규)

**Interfaces:**
- Consumes: `ExportedPptx.survivingEntryIds`(Task 4), `PPTX_SLIDE_CONTENTS_MAP`(Task 1).
- Produces: 저장 후에도 `slideContents`가 새 `SlideEntry.id`로 안전하게 유지된다는 보장 —
  이후 모든 태스크가 이 보장을 전제로 한다.

- [ ] **Step 1: 실패하는 테스트 먼저 작성**

```ts
// tests/roomPptxContent.test.ts
import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeSyncStep1 } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { makeSamplePptx } from './helpers.js';

const MESSAGE_SYNC = 0;

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  readonly mirror = new Y.Doc();
  send(data: Uint8Array, cb?: (err?: Error) => void): void {
    const decoder = decoding.createDecoder(data);
    if (decoding.readVarUint(decoder) === MESSAGE_SYNC) {
      const syncType = decoding.readVarUint(decoder);
      if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) readUpdate(decoder, this.mirror, this);
    }
    cb?.();
  }
  close(): void {}
  ping(): void {}
}

async function flushEventLoop(turns = 30): Promise<void> {
  for (let i = 0; i < turns; i += 1) await new Promise((resolve) => setImmediate(resolve));
}

describe('room.ts: pptx flushToStore re-keys slideContents to the new post-save ids', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-pptx-content-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.useRealTimers();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  test('content written before a save is still readable by the same SlideEntry after autosave renumbers ids', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
    const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');

    const docId = randomUUID();
    await store.saveOriginal(docId, 'pptx', makeSamplePptx(['Slide A']));
    await store.writeMeta({
      id: docId,
      name: 'x.pptx',
      ext: 'pptx',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
      progress: 100,
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'pptx', [], false);

    const ydoc = getRoomYdocForTest(docId)!;
    const slidesArray = ydoc.getArray(PPTX_SLIDES_ARRAY);
    const entryId = slidesArray.get(0).id as string;
    const slideContents = ydoc.getMap(PPTX_SLIDE_CONTENTS_MAP);
    // 이 테스트는 room.ts의 실제 시딩 로직(Task 8)을 아직 안 쓰고, "id로 찾을 수 있다"만
    // 검증하는 것이 목적이라 슬라이드 콘텐츠를 직접 흉내 낸 최소 구조로 채운다.
    ydoc.transact(() => {
      const shapeMap = new Y.Map();
      shapeMap.set('content', 'PLACEHOLDER'); // 이 태스크 시점엔 Y.XmlFragment 배선이 아직 없음 — 값 종류는 중요하지 않다, 키가 살아남는지만 본다
      const perSlide = new Y.Map();
      perSlide.set('shape-1', shapeMap);
      slideContents.set(entryId, perSlide);
    });

    // 자동저장을 강제로 한 번 돌린다(8초 유휴 타이머를 건너뛰고 직접 flushToStore 경로를 태운다).
    await vi.advanceTimersByTimeAsync(9000);
    await flushEventLoop();

    const newEntryId = slidesArray.get(0).id as string;
    expect(newEntryId).not.toBe(entryId); // id는 정말 바뀌었다(기존 동작 확인)
    expect(slideContents.has(entryId)).toBe(false); // 옛 키는 사라졌고
    expect(slideContents.has(newEntryId)).toBe(true); // 새 키로 옮겨졌다
    expect((slideContents.get(newEntryId) as Y.Map<unknown>).get('shape-1')).toBeDefined();
  });
});
```

**주의**: `getRoomYdocForTest`가 아직 없으면(테스트 전용 export) `room.ts`에 xlsx 쪽의
`getRoomLazyGridForTest`와 같은 패턴으로 추가한다:

```ts
/** 테스트 전용 — 룸의 Y.Doc을 직접 들여다본다. */
export function getRoomYdocForTest(docId: string): Y.Doc | null {
  return rooms.get(docId)?.ydoc ?? null;
}
```

(이미 `room.ts`에 이런 함수가 있으면 — `grep -n "getRoomYdocForTest\|room.ydoc"
src/collab/room.ts로 확인 — 새로 안 만들고 재사용한다.)

- [ ] **Step 2: 테스트 실패 확인**

Run: `npx vitest run tests/roomPptxContent.test.ts`
Expected: FAIL(옛 키가 여전히 남아 있거나, `PPTX_SLIDE_CONTENTS_MAP`/`getRoomYdocForTest`가 없어 임포트 에러)

- [ ] **Step 3: 구현**

`src/collab/room.ts`의 pptx `flushToStore` 분기를 수정한다:

```ts
} else if (kind === 'pptx') {
  const original = await store.readOriginal(docId, meta.ext);
  const slidesArray = ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);
  const slideContents = ydoc.getMap<Y.Map<unknown>>(PPTX_SLIDE_CONTENTS_MAP); // <- 추가
  const snapshotBeforeSave = JSON.stringify(slidesArray.toArray());
  const { bytes: newPptxBytes, finalParts, survivingEntryIds } = exportPptxFromSlides(original, slidesArray); // <- survivingEntryIds 구조분해 추가
  await store.saveOriginal(docId, meta.ext, Buffer.from(newPptxBytes));

  if (JSON.stringify(slidesArray.toArray()) === snapshotBeforeSave) {
    room.suppressAutosaveReschedule = true;
    try {
      ydoc.transact(() => {
        const newIds = finalParts.map(() => randomUUID());
        // slideContents 재키잉을 배열 재작성과 같은 트랜잭션 안에서 한다 — 그래야 "배열은 새
        // id인데 slideContents는 아직 옛 id"인 순간이 다른 클라이언트에 보이지 않는다.
        for (let i = 0; i < survivingEntryIds.length; i += 1) {
          const oldId = survivingEntryIds[i];
          const content = slideContents.get(oldId);
          if (content !== undefined) {
            slideContents.set(newIds[i], content);
            slideContents.delete(oldId);
          }
        }
        slidesArray.delete(0, slidesArray.length);
        slidesArray.insert(
          0,
          finalParts.map((part, index) => ({
            id: newIds[index],
            kind: 'original' as const,
            sourceSlidePart: part,
            sourcePageNumber: index + 1,
          })),
        );
      });
    } finally {
      room.suppressAutosaveReschedule = false;
    }
  }

  const converter = findConverter(meta.ext);
  // ...(이하 기존 그대로)
```

`PPTX_SLIDE_CONTENTS_MAP`을 room.ts 상단 import에 추가한다(`import { ...,
PPTX_SLIDE_CONTENTS_MAP, ... } from './pptxShared.js';`).

- [ ] **Step 4: 테스트 통과 확인**

Run: `npx vitest run tests/roomPptxContent.test.ts`
Expected: PASS

- [ ] **Step 5: 전체 테스트 + 타입체크**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 클린, 전부 PASS(특히 `tests/pptxCollab.test.ts`, 기존 pptx 관련 테스트가 전부
그대로 통과해야 한다 — 이 태스크는 저장 로직을 바꾸는 것이라 회귀 위험이 크다)

- [ ] **Step 6: 커밋**

```bash
git add src/collab/room.ts tests/roomPptxContent.test.ts
git commit -m "fix(pptx): re-key slideContents to the new post-save SlideEntry ids in the same transaction"
```

---

### Task 6: 서버 — 지연 시딩(`seedSlideContentIfNeeded`) + WS 메시지 처리 + export 통합

**Files:**
- Modify: `src/pptx/collab.ts`(시딩 함수 추가), `src/collab/room.ts`(WS 메시지 처리, export
  통합), `src/routes/edit.ts`(있다면 메시지 라우팅 확인 — 이 프로젝트는 `handleMessage`가
  `room.ts` 안에 있으므로 실제로는 `room.ts`만 고치면 될 가능성이 높다, 구현 시
  `grep -n "MESSAGE_PAGE_SUBSCRIBE" src/routes/edit.ts src/collab/room.ts`로 실제 위치 재확인)
- Test: `tests/roomPptxContent.test.ts`

**Interfaces:**
- Consumes: `parseSlideShapes`(Task 2), `serializeTxBodyXml`/`spliceShapeTxBody`/
  `appendNewShape`(Task 3), `PPTX_SLIDE_CONTENTS_MAP`/`MESSAGE_SLIDE_CONTENT_SUBSCRIBE`(Task 1).
- Produces: `async function seedSlideContentIfNeeded(room: Room, slideEntryId: string, duplicateOfId?: string): Promise<void>` —
  Task 8(클라이언트)이 이 함수를 호출하도록 만드는 WS 메시지를 보낸다.

- [ ] **Step 1: 실패하는 테스트 먼저 작성**

`tests/roomPptxContent.test.ts`에 이어서(위 describe 블록 안, 또는 새 describe 블록):

```ts
test('opening a slide for the first time seeds its text boxes from the original XML', async () => {
  const { store } = await import('../src/storage.js');
  const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
  const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');

  const docId = randomUUID();
  await store.saveOriginal(docId, 'pptx', makeSamplePptx(['Hello World']));
  await store.writeMeta({
    id: docId, name: 'x.pptx', ext: 'pptx', status: 'ready', pageCount: 1,
    createdAt: new Date().toISOString(), progress: 100,
  });
  const socket = new FakeSocket();
  await joinEditSession(socket as never, {} as never, docId, 'pptx', [], false);

  const ydoc = getRoomYdocForTest(docId)!;
  const entryId = ydoc.getArray(PPTX_SLIDES_ARRAY).get(0).id as string;
  const slideContents = ydoc.getMap(PPTX_SLIDE_CONTENTS_MAP);
  expect(slideContents.has(entryId)).toBe(false); // 열기 전엔 비어 있어야 한다(지연 로딩)

  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, 5); // MESSAGE_SLIDE_CONTENT_SUBSCRIBE
  encoding.writeVarString(encoder, JSON.stringify({ slideEntryId: entryId }));
  socket.emit('message', encoding.toUint8Array(encoder));
  await flushEventLoop();

  const perSlide = slideContents.get(entryId) as Y.Map<unknown> | undefined;
  expect(perSlide).toBeDefined();
  expect(perSlide!.size).toBeGreaterThan(0); // "Hello World" 텍스트박스 하나가 채워졌다
});

test('opening the same slide twice does not overwrite already-edited content', async () => {
  const { store } = await import('../src/storage.js');
  const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
  const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');

  const docId = randomUUID();
  await store.saveOriginal(docId, 'pptx', makeSamplePptx(['Hello World']));
  await store.writeMeta({
    id: docId, name: 'x.pptx', ext: 'pptx', status: 'ready', pageCount: 1,
    createdAt: new Date().toISOString(), progress: 100,
  });
  const socket = new FakeSocket();
  await joinEditSession(socket as never, {} as never, docId, 'pptx', [], false);
  const ydoc = getRoomYdocForTest(docId)!;
  const entryId = ydoc.getArray(PPTX_SLIDES_ARRAY).get(0).id as string;
  const slideContents = ydoc.getMap(PPTX_SLIDE_CONTENTS_MAP);

  const subscribe = () => {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, 5);
    encoding.writeVarString(encoder, JSON.stringify({ slideEntryId: entryId }));
    socket.emit('message', encoding.toUint8Array(encoder));
  };
  subscribe();
  await flushEventLoop();

  // 사용자가 뭔가 편집했다고 흉내낸다 — 표식용 필드를 하나 심는다.
  ydoc.transact(() => {
    (slideContents.get(entryId) as Y.Map<unknown>).set('__edited_marker', true);
  });

  subscribe(); // 같은 슬라이드를 다시 "연다"
  await flushEventLoop();

  expect((slideContents.get(entryId) as Y.Map<unknown>).get('__edited_marker')).toBe(true); // 안 덮어써짐
});

test('a slide that was never opened is exported byte-for-byte identical to its original XML', async () => {
  const { store } = await import('../src/storage.js');
  const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
  const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');
  const { exportPptxFromSlides } = await import('../src/pptx/collab.js');
  const { unzipSync, strFromU8 } = await import('fflate');

  const docId = randomUUID();
  const original = makeSamplePptx(['Slide One', 'Slide Two']);
  await store.saveOriginal(docId, 'pptx', original);
  await store.writeMeta({
    id: docId, name: 'x.pptx', ext: 'pptx', status: 'ready', pageCount: 2,
    createdAt: new Date().toISOString(), progress: 100,
  });
  const socket = new FakeSocket();
  await joinEditSession(socket as never, {} as never, docId, 'pptx', [], false);
  const ydoc = getRoomYdocForTest(docId)!;
  const slidesArray = ydoc.getArray(PPTX_SLIDES_ARRAY);
  const firstEntryId = slidesArray.get(0).id as string;

  // 첫 번째 슬라이드만 연다(시딩) — 두 번째 슬라이드는 한 번도 구독하지 않는다.
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, 5);
  encoding.writeVarString(encoder, JSON.stringify({ slideEntryId: firstEntryId }));
  socket.emit('message', encoding.toUint8Array(encoder));
  await flushEventLoop();

  const slideContents = ydoc.getMap(PPTX_SLIDE_CONTENTS_MAP);
  const result = exportPptxFromSlides(original, slidesArray, slideContents);

  const originalEntries = unzipSync(original);
  const secondSourcePart = (slidesArray.get(1) as { sourceSlidePart: string }).sourceSlidePart;
  const exportedEntries = unzipSync(result.bytes);
  const secondFinalPart = result.finalParts[1];

  expect(strFromU8(exportedEntries[`ppt/slides/${secondFinalPart}`])).toBe(
    strFromU8(originalEntries[`ppt/slides/${secondSourcePart}`]),
  );
});
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `npx vitest run tests/roomPptxContent.test.ts`
Expected: FAIL(메시지를 처리하는 핸들러가 없어 `slideContents`가 계속 비어 있음)

- [ ] **Step 3: 구현**

`src/pptx/collab.ts`에 시딩 함수 추가(`prosemirrorToYXmlFragment`는 `y-prosemirror`에서
가져온다, `room.ts`가 docx export에 이미 쓰고 있는 것과 같은 함수):

```ts
import { prosemirrorToYXmlFragment } from 'y-prosemirror';
import * as Y from 'yjs';
import { parseSlideShapes } from './slideContent.js';
import { unzipSync, strFromU8 } from 'fflate';
import type { SlideShapeMeta } from '../collab/pptxShared.js';

/** 한 슬라이드(entry)의 텍스트박스를 slideContents에 채운다. 이미 채워져 있으면 아무 것도
 * 안 한다(호출부가 먼저 has() 체크를 하지만, 동시에 두 클라이언트가 같은 슬라이드를 처음
 * 여는 경쟁을 대비해 여기서도 한 번 더 확인한다). */
export function seedSlideContentFromXml(
  slideContents: Y.Map<Y.Map<Y.Map<unknown>>>,
  slideEntryId: string,
  pptxBytes: Uint8Array,
  sourceSlidePart: string,
): void {
  if (slideContents.has(slideEntryId)) return;
  const zipEntries = unzipSync(pptxBytes);
  const slideXml = strFromU8(zipEntries[`ppt/slides/${sourceSlidePart}`]);
  const shapes = parseSlideShapes(slideXml);

  const perSlide = new Y.Map<Y.Map<unknown>>();
  for (const shape of shapes) {
    const shapeMap = new Y.Map<unknown>();
    const meta: SlideShapeMeta = { shapeId: shape.shapeId, x: shape.x, y: shape.y, w: shape.w, h: shape.h };
    shapeMap.set('meta', meta);
    const fragment = new Y.XmlFragment();
    shapeMap.set('content', fragment);
    perSlide.set(shape.shapeId, shapeMap);
    // prosemirrorToYXmlFragment는 fragment가 이미 어떤 Y.Doc에 붙어 있어야 동작하므로,
    // 먼저 shapeMap에 꽂아 부모 트랜잭션 트리에 편입시킨 뒤 채운다(순서 중요 — 반대로 하면
    // "아직 Y.Doc에 속하지 않은 타입" 에러가 난다. 실제 구현 시 y-prosemirror 버전의 정확한
    // 요구사항을 테스트로 먼저 확인할 것 — Step 2의 실패 테스트가 이걸 그대로 검증한다).
    prosemirrorToYXmlFragment(shape.doc, fragment);
  }
  slideContents.set(slideEntryId, perSlide);
}

/** 복제된 슬라이드의 콘텐츠를 소스 슬라이드에서 깊은 복사한다(원본 XML이 아니라 "지금까지의
 * 편집 상태"). Y.Map/Y.XmlFragment는 clone()이 없으므로 JSON으로 왕복해 새 인스턴스를 만든다. */
export function duplicateSlideContent(
  slideContents: Y.Map<Y.Map<Y.Map<unknown>>>,
  sourceEntryId: string,
  newEntryId: string,
): void {
  const source = slideContents.get(sourceEntryId);
  if (!source) return; // 소스가 아직 안 열려 비어 있으면 복제본도 빈 채로 둔다 — 호출부가 먼저 소스를 seed해 둔다
  const perSlide = new Y.Map<Y.Map<unknown>>();
  for (const [shapeId, shapeMap] of source.entries()) {
    const newShapeMap = new Y.Map<unknown>();
    newShapeMap.set('meta', shapeMap.get('meta'));
    const newFragment = new Y.XmlFragment();
    newShapeMap.set('content', newFragment);
    perSlide.set(shapeId, newShapeMap);
    const sourceFragment = shapeMap.get('content') as Y.XmlFragment;
    // 얕은 clone: Y.XmlFragment.toArray()로 자식을 꺼내 새 fragment에 옮긴다. 각 자식도
    // Y.XmlElement/Y.XmlText라 clone()을 지원한다(yjs 공식 API).
    for (const child of sourceFragment.toArray()) {
      newFragment.push([(child as { clone(): typeof child }).clone()]);
    }
  }
  slideContents.set(newEntryId, perSlide);
}
```

`src/collab/room.ts`에 `seedSlideContentIfNeeded` 추가하고 `handleMessage`의
`switch`에 새 case를 단다. 아래 코드는 `room.ydoc`/`room.docId`를 쓰는데, 이 두 필드명이
`Room` 타입의 실제 필드명과 일치하는지 구현 시 `grep -n "interface Room" -A 20
src/collab/room.ts`로 먼저 확인한다(이미 `flushToStore` 등 같은 파일의 기존 함수들이 같은
필드를 쓰고 있을 가능성이 높다 — 다르면 그 이름으로 맞춘다):

```ts
async function seedSlideContentIfNeeded(room: Room, slideEntryId: string, duplicateOfId?: string): Promise<void> {
  const slideContents = room.ydoc.getMap<Y.Map<Y.Map<unknown>>>(PPTX_SLIDE_CONTENTS_MAP);
  if (slideContents.has(slideEntryId)) return;
  const slidesArray = room.ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);
  const entry = slidesArray.toArray().find((e) => e.id === slideEntryId);
  if (!entry) return; // 이미 삭제된 슬라이드를 늦게 요청한 경우 — 조용히 무시

  if (duplicateOfId) {
    await seedSlideContentIfNeeded(room, duplicateOfId); // 소스 먼저 채운다
    duplicateSlideContent(slideContents, duplicateOfId, slideEntryId);
    return;
  }
  if (entry.kind === 'blank') {
    slideContents.set(slideEntryId, new Y.Map());
    return;
  }
  if (!entry.sourceSlidePart) return;
  // xlsx Room은 원본 바이트를 room.lazyGrid 쪽에 들고 있지만, pptx Room에는 그런 캐시 필드가
  // 없다 — 매번 store에서 다시 읽는다(이미 flushToStore도 매 저장마다 같은 방식으로 읽는다).
  const bytes = await store.readOriginal(room.docId, 'pptx');
  seedSlideContentFromXml(slideContents, slideEntryId, bytes, entry.sourceSlidePart);
}
```

`handleMessage`의 `switch (messageType)`에 추가:

```ts
case MESSAGE_SLIDE_CONTENT_SUBSCRIBE: {
  try {
    const payload = decoding.readVarString(decoder);
    const parsed = JSON.parse(payload) as { slideEntryId?: unknown; duplicateOfId?: unknown };
    if (typeof parsed.slideEntryId !== 'string') {
      console.error('슬라이드 콘텐츠 구독 메시지 거부(형식 오류):', payload.slice(0, 200));
      break;
    }
    const duplicateOfId = typeof parsed.duplicateOfId === 'string' ? parsed.duplicateOfId : undefined;
    void seedSlideContentIfNeeded(room, parsed.slideEntryId, duplicateOfId).catch((error: unknown) => {
      console.error(`슬라이드 콘텐츠 시딩 실패 (문서 ${room.docId}):`, error);
    });
  } catch (error) {
    console.error(`슬라이드 콘텐츠 구독 메시지 처리 오류 (문서 ${room.docId}):`, error);
  }
  break;
}
```

`export` 쪽도 통합한다 — `exportPptxFromSlides`가 각 slide 파트를 만든 뒤, 그 slide의
`slideContents`가 있으면 도형별로 `spliceShapeTxBody`/`appendNewShape`를 적용하도록
`src/pptx/collab.ts`의 export 루프 끝(파트별 zip 엔트리를 확정하는 지점)에서 한 단계
더 처리한다. 정확한 삽입 지점은 구현 시 현재 함수 본문을 다시 읽고 "zipEntries[해당 파트
경로] = strToU8(xml)"로 확정되는 줄 바로 앞에 끼워 넣는다:

```ts
// exportPptxFromSlides 안, finalParts/zipEntries가 확정된 뒤 bytes를 zipSync하기 전에 추가
for (let i = 0; i < entryList.length; i += 1) {
  const perSlide = slideContents?.get(entryList[i].id);
  if (!perSlide) continue; // 한 번도 안 연 슬라이드 — 원본 그대로 둔다
  const partPath = `ppt/slides/${finalParts[i]}`;
  let slideXml = strFromU8(zipEntries[partPath]);
  for (const [shapeId, shapeMap] of perSlide.entries()) {
    const doc = yXmlFragmentToProseMirrorRootNode(shapeMap.get('content') as Y.XmlFragment, pptxTextSchema);
    const txBodyXml = serializeTxBodyXml(doc);
    const meta = shapeMap.get('meta') as SlideShapeMeta;
    if (shapeId.startsWith('new-')) {
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
```

이 블록을 넣으려면 `exportPptxFromSlides`가 `slideContents`를 매개변수로 받아야 한다 —
함수 시그니처를 `exportPptxFromSlides(pptxBytes, slidesArray, slideContents?:
Y.Map<Y.Map<Y.Map<unknown>>>)`로 넓힌다(선택 인자로 둬서, Task 4에서 이미 고친 테스트들이
3번째 인자 없이 불러도 그대로 통과하게 한다 — 그 테스트들은 콘텐츠 편집을 안 다루므로
`undefined`면 위 for 루프가 전부 `continue`해서 기존 동작과 100% 동일하다).

`flushToStore`(Task 5) 호출부도 `exportPptxFromSlides(original, slidesArray,
slideContents)`로 세 번째 인자를 넘기도록 고친다.

- [ ] **Step 4: 테스트 통과 확인**

Run: `npx vitest run tests/roomPptxContent.test.ts`
Expected: PASS. `prosemirrorToYXmlFragment`/`yXmlFragmentToProseMirrorRootNode`의 정확한
호출 규약(예: fragment를 먼저 부모에 붙여야 하는지, 독립적으로 만들어도 되는지)은
y-prosemirror 실제 타입 선언과 `node_modules/y-prosemirror`의 관련 함수 구현을 직접 열어
확인한다 — docx 쪽(`room.ts`의 기존 `prosemirrorToYXmlFragment(doc, fragment)` 호출부,
`seedRoom`의 docx 분기)이 이미 검증된 실사용 예시이므로 그 순서를 그대로 따라간다.

- [ ] **Step 5: 전체 테스트 + 타입체크**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 클린, 전부 PASS(`tests/pptxCollab.test.ts`의 기존 export 테스트들 포함 — 3번째
인자를 안 넘기던 기존 호출부가 그대로 동작해야 한다)

- [ ] **Step 6: 커밋**

```bash
git add src/pptx/collab.ts src/collab/room.ts tests/roomPptxContent.test.ts
git commit -m "feat(pptx): lazy-seed slide text-box content on first open, merge it back into exported slide XML"
```

---

### Task 7: 서버 — "슬라이드 복제" 액션이 콘텐츠도 같이 복제하도록 흐름 변경

**Files:**
- Modify: `src/collab/room.ts`(또는 `src/routes/edit.ts`, 실제 "복제" 처리가 서버 쪽에 있는지
  확인 필요 — 현재는 `setupCardActions()`가 순수 클라이언트에서 `slidesArray.insert`만 하므로
  서버는 그 배열 변경을 관찰하지 않는다. 이 태스크는 **클라이언트가 복제 직후 곧바로
  `MESSAGE_SLIDE_CONTENT_SUBSCRIBE`를 `duplicateOfId`와 함께 보내는 것**으로 처리한다 —
  새 서버 로직은 필요 없다, Task 6에서 이미 `duplicateOfId` 처리를 만들어 뒀다)
- Test: `tests/roomPptxContent.test.ts`

**Interfaces:**
- Consumes: Task 6의 `seedSlideContentIfNeeded(room, slideEntryId, duplicateOfId)`.
- Produces: 회귀 테스트로 "복제 후 각자 독립 편집"을 못박아 둔다 — Task 8(클라이언트)의
  실제 UI 동작이 이 계약을 지키는지 검증하는 기준이 된다.

- [ ] **Step 1: 실패하는 테스트 먼저 작성**

```ts
test('duplicating a slide copies its content independently — editing the copy does not touch the original', async () => {
  const { store } = await import('../src/storage.js');
  const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
  const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');

  const docId = randomUUID();
  await store.saveOriginal(docId, 'pptx', makeSamplePptx(['Original Text']));
  await store.writeMeta({
    id: docId, name: 'x.pptx', ext: 'pptx', status: 'ready', pageCount: 1,
    createdAt: new Date().toISOString(), progress: 100,
  });
  const socket = new FakeSocket();
  await joinEditSession(socket as never, {} as never, docId, 'pptx', [], false);
  const ydoc = getRoomYdocForTest(docId)!;
  const slidesArray = ydoc.getArray(PPTX_SLIDES_ARRAY);
  const originalId = slidesArray.get(0).id as string;

  // 1) 원본을 연다(시딩).
  const subscribe = (slideEntryId: string, duplicateOfId?: string) => {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, 5);
    encoding.writeVarString(encoder, JSON.stringify({ slideEntryId, duplicateOfId }));
    socket.emit('message', encoding.toUint8Array(encoder));
  };
  subscribe(originalId);
  await flushEventLoop();

  // 2) 클라이언트가 하는 것과 같은 방식으로 복제 항목을 배열에 넣는다.
  const dupId = randomUUID();
  ydoc.transact(() => {
    slidesArray.insert(1, [{ id: dupId, kind: 'duplicate', sourceSlidePart: 'slide1.xml', sourcePageNumber: null }]);
  });
  subscribe(dupId, originalId); // 3) 복제 콘텐츠 시딩 요청
  await flushEventLoop();

  const slideContents = ydoc.getMap(PPTX_SLIDE_CONTENTS_MAP);
  const dupShapes = slideContents.get(dupId) as Y.Map<Y.Map<unknown>>;
  expect(dupShapes.size).toBeGreaterThan(0);

  // 4) 복제본의 첫 도형 내용을 지운다(편집 흉내).
  const [shapeId] = [...dupShapes.keys()];
  ydoc.transact(() => {
    const fragment = (dupShapes.get(shapeId) as Y.Map<unknown>).get('content') as Y.XmlFragment;
    fragment.delete(0, fragment.length);
  });

  // 5) 원본은 그대로여야 한다.
  const originalShapes = slideContents.get(originalId) as Y.Map<Y.Map<unknown>>;
  const originalFragment = (originalShapes.get(shapeId) as Y.Map<unknown>).get('content') as Y.XmlFragment;
  expect(originalFragment.length).toBeGreaterThan(0);
});
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `npx vitest run tests/roomPptxContent.test.ts`
Expected: Task 6를 이미 끝냈다면 이 테스트는 **이미 통과할 가능성이 높다**(Task 6에서
`duplicateOfId` 처리를 이미 구현했으므로) — 이 태스크는 사실상 그 계약을 명시적으로
검증·고정하는 회귀 테스트 태스크다. 만약 실패하면 Task 6의 `duplicateSlideContent`
구현(특히 `Y.XmlFragment` clone 부분)을 다시 확인한다.

- [ ] **Step 3~4: (Task 6에서 이미 구현됐다면 생략, 실패 시 여기서 수정)**

- [ ] **Step 5: 전체 테스트 + 타입체크**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 클린, 전부 PASS

- [ ] **Step 6: 커밋**

```bash
git add tests/roomPptxContent.test.ts
git commit -m "test(pptx): lock in duplicate-slide content independence as a regression test"
```

---

### Task 8: 클라이언트 — 두 단(왼쪽 레일 · 오른쪽 콘텐츠 캔버스) 레이아웃 + 왼쪽 레일 재배치

**Files:**
- Modify: `public/editor-slides.html`
- Modify: `public/css/style.css`
- Modify: `src/client/editor-slides.ts`(왼쪽 레일 관련 함수만 — `cardFor`, `renderSlides`,
  `setupDragAndDrop`, `setupCardActions`는 대부분 그대로 재사용, DOM 컨테이너 id만 바뀜)

**Interfaces:**
- Consumes: 기존 `SlideEntry`/`slidesArray`(안 바뀜).
- Produces: 오른쪽 콘텐츠 영역을 담을 빈 컨테이너(`#slide-canvas`) — Task 9가 그 안을 채운다.
  클릭 시 `activeSlideId`(모듈 전역 상태)가 바뀌고 `onActiveSlideChange(entryId)` 콜백이
  불린다 — Task 9가 이 콜백에 실제 렌더링을 건다.

- [ ] **Step 1: HTML 구조 변경**

`public/editor-slides.html`의 `<main id="slides-grid" class="slides-editor-grid"
aria-label="슬라이드 목록"></main>` 한 줄을 아래로 교체:

```html
<div class="slides-editor-body">
  <aside id="slides-rail" class="slides-rail" aria-label="슬라이드 목록"></aside>
  <main id="slide-canvas-wrap" class="slide-canvas-wrap">
    <div id="slide-canvas" class="slide-canvas" aria-label="선택한 슬라이드 콘텐츠"></div>
  </main>
</div>
```

`<template id="slide-card-template">`는 그대로 두되(카드 자체 마크업은 안 바뀜), 그 안의
`.slide-card`에 클릭 시 활성 표시를 위한 클래스 훅만 이미 있으면 재사용(없으면
`.slide-card.active { ... }` CSS를 새로 추가).

- [ ] **Step 2: CSS — 좌우 분할 레이아웃**

`public/css/style.css`에 추가(기존 `.slides-editor-grid` 블록 근처 — 위치는
`grep -n "slides-editor-grid" public/css/style.css`로 확인):

```css
.slides-editor-body {
  display: flex;
  height: calc(100vh - var(--toolbar-h, 96px)); /* 실제 툴바 높이 변수/값은 기존 .grid-editor-page 등의 계산 방식을 그대로 따른다 */
  overflow: hidden;
}

.slides-rail {
  width: 220px;
  flex: none;
  overflow-y: auto;
  padding: 0.75rem;
  display: flex;
  flex-direction: column;
  gap: 0.75rem;
  background: var(--v-glass);
  border-right: 1px solid var(--v-glass-line);
}

/* 레일 안 카드는 기존 격자 카드보다 작게 — 썸네일 폭 기준으로 눌러 준다 */
.slides-rail .slide-card { width: 100%; }
.slides-rail .slide-card.active .slide-card-thumb { outline: 2px solid var(--v-accent); outline-offset: -2px; }

.slide-canvas-wrap {
  flex: 1;
  overflow: auto;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 2rem;
  background: #e9eaed;
}

.slide-canvas {
  position: relative;
  background: #fff;
  box-shadow: 0 12px 40px rgba(0, 0, 0, 0.25);
  /* 실제 폭/높이는 JS가 슬라이드 크기(EMU)에 맞춰 인라인 style로 설정한다(Task 9) */
}
```

(`--toolbar-h`가 이 코드베이스에 실제로 없으면, 기존 `.grid-editor-page`가 툴바 높이를
어떻게 계산해 두는지 `grep -n "toolbar" public/css/style.css`로 확인하고 같은 방식을
따른다 — 매직 넘버를 새로 만들지 않는다.)

- [ ] **Step 3: `editor-slides.ts` — 컨테이너 id 갱신 + 활성 슬라이드 상태 추가**

```ts
// 기존
const gridEl = document.getElementById('slides-grid') as HTMLElement;
// 아래로 교체
const railEl = document.getElementById('slides-rail') as HTMLElement;
const canvasEl = document.getElementById('slide-canvas') as HTMLElement;

let activeSlideId: string | null = null;
let onActiveSlideChange: ((entryId: string) => void) | null = null; // Task 9가 채운다
```

`cardFor`/`renderSlides`/`setupDragAndDrop`가 `gridEl`을 참조하던 부분을 전부 `railEl`로
바꾼다(단순 이름 교체, 로직 불변). `renderSlides()` 끝에 활성 카드 표시를 추가:

```ts
function renderSlides(): void {
  const entries = slidesArray.toArray();
  railEl.replaceChildren(...entries.map((entry, index) => cardFor(entry, index)));
  if (!activeSlideId && entries.length > 0) selectSlide(entries[0].id);
  for (const card of railEl.querySelectorAll('.slide-card')) {
    card.classList.toggle('active', (card as HTMLElement).dataset.id === activeSlideId);
  }
}

function selectSlide(entryId: string): void {
  activeSlideId = entryId;
  for (const card of railEl.querySelectorAll('.slide-card')) {
    card.classList.toggle('active', (card as HTMLElement).dataset.id === entryId);
  }
  onActiveSlideChange?.(entryId);
}
```

`setupCardActions()`의 클릭 핸들러 맨 앞에 "복제/삭제 버튼이 아니라 카드 자체를 클릭했으면
`selectSlide`" 분기를 추가:

```ts
gridEl.addEventListener('click', (event) => {
  const target = event.target as HTMLElement;
  const card = target.closest<HTMLElement>('.slide-card');
  if (!card) return;
  if (!target.closest('.slide-card-duplicate') && !target.closest('.slide-card-delete')) {
    selectSlide(card.dataset.id!);
    return; // 카드 몸통 클릭은 선택만 하고 끝 — 아래 복제/삭제 로직으로 안 내려간다
  }
  // ...(기존 복제/삭제 로직 그대로, gridEl -> railEl 이름만 교체)
});
```

- [ ] **Step 4: 빌드 + 수동 확인**

Run: `node scripts/build-editor.mjs`
Expected: 에러 없이 `public/js/editor-slides.bundle.js` 재생성.

dev 서버를 띄우고 실제 pptx 문서로 "동시편집 열기"를 열어, 왼쪽에 세로 슬라이드 목록,
오른쪽에 빈 캔버스 영역이 보이는지, 슬라이드 카드를 클릭하면 `.active` 표시가 옮겨가는지
확인한다(이 태스크 시점엔 오른쪽이 아직 비어 있는 게 정상 — Task 9가 채운다).

- [ ] **Step 5: 커밋**

```bash
git add public/editor-slides.html public/css/style.css src/client/editor-slides.ts
git commit -m "feat(pptx): restructure the slide editor into a left rail + right canvas layout"
```

---

### Task 9: 클라이언트 — 오른쪽 캔버스에 텍스트박스 배치 + ProseMirror 마운트(읽기 전용 포함)

**Files:**
- Modify: `src/client/editor-slides.ts`

**Interfaces:**
- Consumes: Task 8의 `onActiveSlideChange` 훅, `pptxTextSchema`(Task 1),
  `PPTX_SLIDE_CONTENTS_MAP`/`MESSAGE_SLIDE_CONTENT_SUBSCRIBE`(Task 1), y-prosemirror.
- Produces: 슬라이드를 열면 그 슬라이드의 텍스트박스가 실제 위치에 보이고, `canEdit`이면
  클릭해서 타이핑 가능 — Task 10(서식 툴바)이 여기서 만든 활성 `EditorView`를 대상으로 동작한다.

- [ ] **Step 1: 슬라이드 크기(EMU) 읽어 캔버스 스케일 계산**

`fetchViewInfo`가 돌려주는 `ViewInfo`에 슬라이드 크기가 없다 — `/api/view/info` 응답을
확장해야 하는지, 아니면 `slidesArray`와 별개로 룸이 따로 내려주는 값이 있는지 먼저
확인한다(`grep -n "sldSz\|slideWidth\|slideSize" src/routes/viewer.ts src/collab/room.ts`).
없으면 표준 16:9(`12192000 x 6858000` EMU) 기본값을 쓰고, 실제 프레젠테이션 크기는 이후
개선 과제로 남긴다(이번 태스크에서 서버 API를 넓히지 않는다 — 범위 밖 확장 자제).

```ts
const SLIDE_W_EMU = 12192000;
const SLIDE_H_EMU = 6858000;
const EMU_PER_PX = SLIDE_W_EMU / 960; // 캔버스 폭을 960px 기준으로 스케일(반응형은 이후 과제)

function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}
```

- [ ] **Step 2: WS로 슬라이드 콘텐츠 구독 + 응답 관찰**

```ts
const MESSAGE_SLIDE_CONTENT_SUBSCRIBE = 5;

function subscribeSlideContent(entryId: string, duplicateOfId?: string): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SLIDE_CONTENT_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify({ slideEntryId: entryId, duplicateOfId }));
  provider.ws?.send(encoding.toUint8Array(encoder)); // provider.ws는 y-websocket이 내부적으로 들고 있는 실제 WebSocket — WebsocketProvider의 실제 필드명은 구현 시 node_modules/y-websocket 타입 선언으로 재확인(비공개 필드일 수 있음, 그 경우 provider가 노출하는 공개 send 경로를 대신 쓴다)
}
```

**주의**: `WebsocketProvider`가 원시 소켓 전송을 공개 API로 노출하지 않을 수 있다(비공개
`ws` 필드일 가능성). 구현 시작 전에 `node_modules/y-websocket/src/y-websocket.js`를 열어
`MESSAGE_PAGE_SUBSCRIBE`를 이미 이렇게 보내고 있는 `grid-editor.ts`의 실제 전송 코드
(`grep -n "MESSAGE_PAGE_SUBSCRIBE" -A 10 src/client/grid-editor.ts`)를 그대로 베낀다 —
이미 검증된 패턴이 있다.

`PPTX_SLIDE_CONTENTS_MAP` Y.Map을 관찰해 슬라이드가 시딩되면 캔버스를 다시 그린다:

```ts
const slideContents = ydoc.getMap<Y.Map<Y.Map<unknown>>>(PPTX_SLIDE_CONTENTS_MAP);
slideContents.observe(() => {
  if (activeSlideId) renderCanvasForSlide(activeSlideId);
});
```

- [ ] **Step 3: 텍스트박스 배치 + ProseMirror 마운트/언마운트**

```ts
const mountedViews = new Map<string, EditorView>(); // shapeId -> EditorView, 슬라이드를 벗어나면 전부 destroy

function renderCanvasForSlide(entryId: string): void {
  for (const view of mountedViews.values()) view.destroy();
  mountedViews.clear();
  canvasEl.style.width = `${emuToPx(SLIDE_W_EMU)}px`;
  canvasEl.style.height = `${emuToPx(SLIDE_H_EMU)}px`;
  canvasEl.replaceChildren();

  const perSlide = slideContents.get(entryId);
  if (!perSlide) {
    subscribeSlideContent(entryId); // 아직 안 열렸다 — 요청하고 observe가 다시 불러줄 때까지 대기
    return;
  }
  for (const [shapeId, shapeMap] of perSlide.entries()) {
    const meta = shapeMap.get('meta') as SlideShapeMeta;
    const box = document.createElement('div');
    box.className = 'slide-textbox';
    box.style.left = `${emuToPx(meta.x ?? SLIDE_W_EMU * 0.2)}px`;
    box.style.top = `${emuToPx(meta.y ?? SLIDE_H_EMU * 0.4)}px`;
    box.style.width = `${emuToPx(meta.w ?? SLIDE_W_EMU * NEW_TEXTBOX_WIDTH_RATIO)}px`;
    box.style.height = `${emuToPx(meta.h ?? SLIDE_H_EMU * NEW_TEXTBOX_HEIGHT_RATIO)}px`;
    canvasEl.append(box);

    const fragment = shapeMap.get('content') as Y.XmlFragment;
    const view = new EditorView(box, {
      state: EditorState.create({
        schema: pptxTextSchema,
        plugins: [ySyncPlugin(fragment), yCursorPlugin(provider.awareness), yUndoPlugin()],
      }),
      editable: () => canEdit,
    });
    mountedViews.set(shapeId, view);
  }
}
```

`onActiveSlideChange = renderCanvasForSlide;`를 `main()` 초기화부에 연결한다(Task 8이
만든 훅).

**CSS 추가**(`public/css/style.css`):

```css
.slide-textbox {
  position: absolute;
  padding: 4px;
  outline: 1px dashed transparent;
}
.slide-textbox:hover, .slide-textbox:focus-within { outline-color: var(--v-glass-line); }
.slide-textbox .ProseMirror { outline: none; min-height: 1em; }
```

- [ ] **Step 4: 빌드 + 실제 pptx로 수동 확인**

Run: `node scripts/build-editor.mjs`

dev 서버에서 여러 텍스트박스가 있는 실제 pptx를 "동시편집 열기"로 열어: 슬라이드를 클릭하면
그 슬라이드의 텍스트박스들이 원래 위치 근처에 나타나는지, 클릭해서 타이핑하면 글자가
바뀌는지, 다른 슬라이드로 갔다 돌아와도 내용이 유지되는지(=서버에 저장됐는지, 새로고침
후에도 남아 있는지) 확인한다. 두 브라우저 탭을 열어 같은 텍스트박스에 동시에 타이핑해도
둘 다 살아남는지(CRDT 병합)도 확인한다.

- [ ] **Step 5: 커밋**

```bash
git add src/client/editor-slides.ts public/css/style.css
git commit -m "feat(pptx): render and live-edit text boxes on the right-hand slide canvas"
```

---

### Task 10: 클라이언트 — 서식 툴바(굵게/기울임/밑줄/색/크기)

**Files:**
- Modify: `public/editor-slides.html`(툴바 버튼 마크업)
- Modify: `src/client/editor-slides.ts`

**Interfaces:**
- Consumes: Task 9의 `mountedViews`(현재 포커스된 `EditorView`를 알아야 함).

- [ ] **Step 1: 포커스된 EditorView 추적**

```ts
let focusedView: EditorView | null = null;
// renderCanvasForSlide 안, new EditorView(...) 옵션에 추가:
//   dispatchTransaction(tr) { const newState = view.state.apply(tr); view.updateState(newState); },
//   handleDOMEvents: { focus: () => { focusedView = view; return false; } }
```

- [ ] **Step 2: 툴바 버튼 + 커맨드 연결**

`public/editor-slides.html`의 `.toolbar-format`(빈 슬라이드 추가 버튼이 있던 자리)에
버튼 4개 추가: `btn-bold`, `btn-italic`, `btn-underline`, `input#text-color-picker`
(`type="color"`). `src/client/editor-slides.ts`:

```ts
import { toggleMark } from 'prosemirror-commands';

function bindFormatButton(id: string, markName: keyof typeof pptxTextSchema.marks): void {
  document.getElementById(id)?.addEventListener('click', () => {
    if (!focusedView) return;
    toggleMark(pptxTextSchema.marks[markName])(focusedView.state, focusedView.dispatch);
    focusedView.focus();
  });
}
bindFormatButton('btn-bold', 'strong');
bindFormatButton('btn-italic', 'em');
bindFormatButton('btn-underline', 'underline');

document.getElementById('text-color-picker')?.addEventListener('input', (event) => {
  if (!focusedView) return;
  const color = (event.target as HTMLInputElement).value;
  const { from, to } = focusedView.state.selection;
  const tr = focusedView.state.tr.addMark(from, to, pptxTextSchema.marks.textColor.create({ color }));
  focusedView.dispatch(tr);
});
```

- [ ] **Step 3: 빌드 + 수동 확인**

Run: `node scripts/build-editor.mjs`

텍스트를 선택하고 굵게/기울임/밑줄/색 버튼을 눌러 실제로 적용되는지, 저장 후 다시 열어도
유지되는지(export 라운드트립) 확인한다.

- [ ] **Step 4: 커밋**

```bash
git add public/editor-slides.html src/client/editor-slides.ts
git commit -m "feat(pptx): add a formatting toolbar (bold/italic/underline/color) for the active text box"
```

---

### Task 11: 클라이언트 — 새 텍스트박스 추가/삭제 + 읽기 전용 모드

**Files:**
- Modify: `public/editor-slides.html`
- Modify: `src/client/editor-slides.ts`

**Interfaces:**
- Consumes: Task 9의 `slideContents`/`renderCanvasForSlide`.

- [ ] **Step 1: "텍스트 추가" 버튼**

```ts
document.getElementById('btn-add-textbox')?.addEventListener('click', () => {
  if (!canEdit || !activeSlideId) return;
  const perSlide = slideContents.get(activeSlideId) ?? new Y.Map();
  const shapeId = `new-${crypto.randomUUID()}`;
  ydoc.transact(() => {
    const shapeMap = new Y.Map();
    const meta: SlideShapeMeta = {
      shapeId,
      x: Math.round(SLIDE_W_EMU * (1 - NEW_TEXTBOX_WIDTH_RATIO) / 2),
      y: Math.round(SLIDE_H_EMU * (1 - NEW_TEXTBOX_HEIGHT_RATIO) / 2),
      w: Math.round(SLIDE_W_EMU * NEW_TEXTBOX_WIDTH_RATIO),
      h: Math.round(SLIDE_H_EMU * NEW_TEXTBOX_HEIGHT_RATIO),
    };
    shapeMap.set('meta', meta);
    shapeMap.set('content', new Y.XmlFragment());
    perSlide.set(shapeId, shapeMap);
    if (!slideContents.has(activeSlideId!)) slideContents.set(activeSlideId!, perSlide);
  });
});
```

- [ ] **Step 2: 텍스트박스 삭제(호버 시 작은 × 버튼)**

`renderCanvasForSlide`의 텍스트박스 생성 루프에서 `box`에 삭제 버튼을 같이 붙인다:

```ts
if (canEdit) {
  const deleteBtn = document.createElement('button');
  deleteBtn.className = 'slide-textbox-delete';
  deleteBtn.textContent = '×';
  deleteBtn.addEventListener('click', (event) => {
    event.stopPropagation();
    ydoc.transact(() => { perSlide.delete(shapeId); });
  });
  box.append(deleteBtn);
}
```

CSS: `.slide-textbox-delete { position: absolute; top: -10px; right: -10px; display: none; }`
`.slide-textbox:hover .slide-textbox-delete { display: block; }`

- [ ] **Step 3: 읽기 전용 모드**

`canEdit === false`면: `#btn-add-textbox`/삭제 버튼 숨김(이미 Task 1의 기존 패턴 —
`if (!canEdit) { document.getElementById('btn-add-blank')?.setAttribute('hidden', ''); }`
근처에 같이 추가), `EditorView`는 이미 Task 9에서 `editable: () => canEdit`로 막혀 있다.

- [ ] **Step 4: 빌드 + 수동 확인**

Run: `node scripts/build-editor.mjs`

빈 슬라이드를 추가하고 "텍스트 추가"를 눌러 새 텍스트박스가 중앙에 생기는지, 타이핑 후
저장·재오픈해도 남아 있는지, 삭제가 되는지, 읽기 전용 토큰으로 열었을 때 그 버튼들이
안 보이고 타이핑도 안 되는지 확인한다.

- [ ] **Step 5: 커밋**

```bash
git add public/editor-slides.html public/css/style.css src/client/editor-slides.ts
git commit -m "feat(pptx): support adding/removing text boxes on shapeless slides; wire up read-only mode"
```

---

### Task 12: 통합 테스트 + 실제 pptx 파일 end-to-end 확인

**Files:** 없음(검증 전용) — 필요 시 `tests/roomPptxContent.test.ts`에 통합 테스트 추가.

- [ ] **Step 1: 두 클라이언트가 같은 텍스트박스를 동시에 타이핑해도 둘 다 살아남는지**

`tests/roomPptxContent.test.ts`에 추가(두 개의 `FakeSocket` + 두 개의 mirror `Y.Doc`으로
같은 룸에 join, 같은 `Y.XmlFragment` 경로에 각각 다른 글자를 동시 삽입 — 패턴은
`tests/roomStructureSync.test.ts`의 "두 구조 연산이 연쇄로 합성되는지" 테스트를 참고해
서로 다른 위치에 텍스트를 삽입하는 두 트랜잭션을 순서대로 적용하고 둘 다 최종 문서에
남아 있는지 확인).

- [ ] **Step 2: 전체 테스트 스위트**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 클린, 전부 PASS

- [ ] **Step 3: 실제 pptx 파일로 수동 end-to-end**

로컬에서 실제 파워포인트 파일(여러 슬라이드, 굵게/색/크기가 섞인 텍스트박스 포함)을 업로드
→ "동시편집 열기" → 왼쪽 레일에서 슬라이드 추가/복제/삭제/드래그 순서바꾸기 → 오른쪽에서
텍스트 수정 + 서식 적용 → 저장 → 다운로드해서 실제 PowerPoint(또는 LibreOffice Impress)로
다시 열어 내용·서식·다른 슬라이드가 전부 멀쩡한지 확인. 특히:
- 편집 안 한 슬라이드(한 번도 안 연 슬라이드)가 원본 그대로인지(바이트 비교까지는 아니어도
  최소 육안 확인).
- 표·이미지가 있는 슬라이드가 그대로 보존되는지.
- 두 번 이상 저장(자동저장 포함)해도 내용이 안 사라지는지(Task 5의 id 재키잉 검증).

- [ ] **Step 4: `docs/TODO.md`에 이 작업 완료 기록**

기존 항목 형식(이번 세션의 다른 완료 항목들 참고)을 따라 "pptx 텍스트박스 콘텐츠 동시편집"
섹션을 추가하고, 실제로 확인한 내용(위 Step 3에서 무엇을 확인했는지)을 요약한다.

```bash
git add docs/TODO.md
git commit -m "docs: mark pptx text-box content editing as complete in TODO.md"
```

---

### Task 13: 삭제한 텍스트박스가 저장 파일에 실제로 반영되지 않는 버그 수정

**이 태스크가 왜 있는가**: Task 12의 실제 파워포인트 파일 end-to-end 검증 중 발견된 실제
버그다. Task 11에서 만든 텍스트박스 삭제 버튼은 `perSlide.delete(shapeId)`로 클라이언트
Yjs 상태에서는 정확히 지워지고 UI에도 정확히 사라지지만, `exportPptxFromSlides`의 병합
루프(`src/pptx/collab.ts`)는 **현재 `perSlide`에 남아있는 도형만** 순회하며 갱신/추가할 뿐,
"원본 XML에는 있었는데 지금 `perSlide`에는 없는" 도형을 실제로 제거하는 로직이 어디에도
없다 — Task 6(병합 루프를 만든 태스크)은 Task 11(삭제 버튼을 만든 태스크)보다 먼저 만들어져
애초에 "삭제"라는 개념 자체가 없었고, Task 11도 클라이언트 쪽만 다뤘다. 결과: 삭제 버튼을
누르면 UI는 "삭제됨"을 보여주고 저장 상태는 "저장됨"을 보여주지만, 실제 저장된 파일에는
그 도형이 그대로 남아있다 — 더 나쁘게는, 시딩이 1회성(`has()` 체크)이라 서버 재시작 후 그
슬라이드가 다시 시딩되면 사용자가 명시적으로 지운 도형이 되살아날 수 있다. "원본 유실 금지"
원칙의 반대 방향 실패(유실이 아니라 사용자가 원치 않는 내용의 "부활")이지만, 이 도구의
핵심 신뢰 계약("편집한 대로 저장된다")을 똑같이 어기므로 반드시 고친다.

**Files:**
- Modify: `src/pptx/slideContent.ts` (새 `removeShape` 함수 추가)
- Modify: `src/pptx/collab.ts` (`exportPptxFromSlides`의 병합 루프 뒤에 제거 단계 추가)
- Test: `tests/pptxSlideContent.test.ts`, `tests/roomPptxContent.test.ts`

**Interfaces:**
- Consumes: `hasShape`/`spliceShapeTxBody`의 기존 `<p:sp>` 블록 탐지 정규식(Task 3/6에서
  이미 검증됨), `exportPptxFromSlides`에 이미 존재하는 `baselineShapes` 맵(Task 6의 "수정
  안 된 도형은 건드리지 않기" 최적화용으로 이미 만들어져 있음 — 이번엔 "삭제된 도형 찾기"
  용도로 재사용한다).
- Produces: `function removeShape(slideXml: string, shapeId: string): string` — 해당
  shapeId의 `<p:sp>...</p:sp>` 블록 전체를 제거한다. 못 찾으면 throw(이 파일의 기존 "절대
  조용히 실패하지 않는다" 원칙과 동일).

- [ ] **Step 1: 실패하는 테스트 먼저 작성**

`tests/pptxSlideContent.test.ts`에 추가(기존 `describe('serializeTxBodyXml + spliceShapeTxBody (export)', ...)` 블록 안, 또는 새 `describe` 블록):

```ts
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
});
```

`removeShape`/`appendNewShape`/`parseSlideShapes`/`serializeTxBodyXml`/`pptxTextSchema as schema`를
이 파일 상단 import 목록에 추가한다(이미 대부분 import돼 있을 것 — `removeShape`만 새로 추가).

- [ ] **Step 2: 테스트 실패 확인**

Run: `npx vitest run tests/pptxSlideContent.test.ts`
Expected: FAIL(`removeShape` 정의 안 됨)

- [ ] **Step 3: `removeShape` 구현** (`src/pptx/slideContent.ts`에 추가)

```ts
/** shapeId에 해당하는 <p:sp> 블록 전체를 슬라이드 XML에서 제거한다(사용자가 텍스트박스를
 * 삭제했을 때 씀). spliceShapeTxBody와 같은 블록 탐지 정규식을 쓰되, txBody만 바꾸는 대신
 * 블록 전체를 들어낸다. 다른 도형은 전혀 건드리지 않는다. */
export function removeShape(slideXml: string, shapeId: string): string {
  const spBlockRe = /<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g;
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
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `npx vitest run tests/pptxSlideContent.test.ts`
Expected: PASS(새 3개 포함 전부)

- [ ] **Step 5: `exportPptxFromSlides`의 병합 루프에 제거 단계 추가**

`src/pptx/collab.ts`를 읽고(`grep -n "baselineShapes\|perSlide.entries" src/pptx/collab.ts`로
정확한 현재 줄 위치 재확인) 기존 도형별 update/append 루프가 끝난 직후, 그 슬라이드의 XML을
확정하기 전(`zipEntries[partPath] = strToU8(slideXml)` 하기 전)에 추가:

```ts
// perSlide에 이미 있는 도형들은 위 루프에서 처리됐다. 여기서는 "원본/baseline에는 있었는데
// 지금 perSlide에는 없는" 도형 — 즉 사용자가 삭제한 도형 — 을 실제로 슬라이드 XML에서
// 제거한다. baselineShapes가 null인 경우(kind:'blank', 애초에 baseline이 없음)는 삭제
// 개념 자체가 성립하지 않으므로 건너뛴다.
if (baselineShapes) {
  for (const baselineShapeId of baselineShapes.keys()) {
    if (!perSlide.has(baselineShapeId)) {
      slideXml = removeShape(slideXml, baselineShapeId);
    }
  }
}
```

(`baselineShapes`의 정확한 타입이 `Map`인지 `Record`인지, `perSlide`가 정확히 어떤 변수명인지는
실제 코드를 읽고 맞춘다 — 위 스니펫은 방향을 보여주는 것.) `removeShape`를 이 파일 상단의
`slideContent.js` import에 추가한다.

- [ ] **Step 6: 서버 통합 테스트 추가**

`tests/roomPptxContent.test.ts`에 추가(기존 파일의 FakeSocket/flushEventLoop/subscribeSlideContent
헬퍼 재사용, 새로 선언하지 않음):

1. **삭제 후 export하면 실제로 사라지는지**: 슬라이드를 열어(seed) 도형이 2개 이상 있는
   슬라이드를 만들고, 그 중 하나를 `perSlide.delete(shapeId)`로 지운 뒤 `exportPptxFromSlides`를
   호출 — 결과 XML에 그 도형이 없는지, 다른 도형은 그대로인지 확인.
2. **두 번째 저장에서도 되살아나지 않는지**: 위 상태에서 `flushToStore`를 한 번 더 트리거(또는
   `exportPptxFromSlides`를 다시 호출)해서, 삭제된 도형이 여전히 없는지 재확인(재시딩되지
   않는 한 baseline 비교 결과가 안정적이어야 함).
3. **한 번도 저장 안 된 new- 도형을 삭제하면 그냥 아무 일도 안 일어나는지(no-op)**: `new-`
   도형을 추가만 하고 export를 한 번도 안 한 상태에서 바로 삭제하면, baseline에 애초에 없던
   도형이므로 `baselineShapes.has(...)`가 false — 제거 루프가 아무것도 안 하는 게 맞다(이미
   존재하지 않는 도형을 지우려는 게 아니라, 애초에 export된 적 없는 도형이라 제거할 대상
   자체가 없는 경우). 이 케이스가 에러 없이 조용히 넘어가는지 확인.
4. **슬라이드의 모든 도형을 다 지운 경우**: 도형이 하나도 안 남은 슬라이드를 export해도
   에러 없이 빈 `<p:spTree>`(또는 최소 구조)로 저장되는지 확인.

- [ ] **Step 7: 전체 테스트 + 타입체크**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 클린, 전부 PASS(특히 Task 6/12에서 이미 검증된 "수정 안 된 도형은 안 건드림"
동작이 이번 변경으로 깨지지 않았는지 — 기존 byte-for-byte 테스트들이 여전히 통과해야 함)

- [ ] **Step 8: 실제 pptx로 재확인**

Task 12에서 쓴 것과 같은(또는 비슷한) 실제 pptx 파일로: 텍스트박스 하나를 삭제 → 저장 →
다운로드 → 재파싱해서 그 도형이 실제로 없는지 확인. 두 번 저장해도 되살아나지 않는지도
확인.

- [ ] **Step 9: `docs/TODO.md` 갱신**

Task 12에서 "알려진 문제"로 적어둔 삭제 버그 항목을 "해결됨"으로 갱신한다.

- [ ] **Step 10: 커밋**

```bash
git add src/pptx/slideContent.ts src/pptx/collab.ts tests/pptxSlideContent.test.ts tests/roomPptxContent.test.ts docs/TODO.md
git commit -m "fix(pptx): actually remove a deleted text box's shape from the exported slide XML"
```
