# PDF 뷰어 마크업(리뷰 코멘트) 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 원본 파일은 건드리지 않고, 뷰어에서 페이지 위 사각형 영역에 리뷰 코멘트(마크업)를 남기고 공유해서 볼 수 있게 하며, 출력/다운로드 시 옵션으로 포함할 수 있게 한다.

**Architecture:** 문서당 마크업 배열 하나를 기존 저장 계층(`DerivedKind`)에 새 종류로 추가해 JSON으로 영속화한다. 기존 열람 토큰(`canEdit`) 인증 방식을 그대로 재사용하는 3개의 새 라우트(`GET/POST/DELETE /api/view/markup...`)로 CRUD를 노출하고, 기존 출력(`/api/view/print`)·다운로드(`/api/view/download`) 라우트에는 opt-in 쿼리 파라미터(`includeMarkup=1`)만 추가한다. 프론트엔드는 이미 있는 "거리 측정" 도구(`measure` 객체, 클릭 인터랙션)와 같은 구조를 그대로 따르되, 드래그로 사각형을 그리는 인터랙션과 코멘트 팝오버를 추가한다.

**Tech Stack:** Fastify(TS) 서버, `pdf-lib`(PDF에 사각형·텍스트 그리기), vitest, 바닐라 JS 프론트엔드(빌드 없음, `public/js/viewer.js` 직접 서빙).

**Spec:** `docs/superpowers/specs/2026-09-01-pdf-viewer-markup-design.md`

## Global Constraints

- 원본 파일(`data/originals/*`)은 절대 쓰지 않는다 — 마크업은 항상 별도 파생 데이터.
- 마크업 조회는 열람 가능한 모든 사람에게 허용, 추가/삭제는 열람 토큰의 `canEdit`이 있어야 한다.
- 적용 범위는 `viewKind: 'pages'` 문서로 한정한다 (xlsx 그리드 뷰 제외).
- `includeMarkup` 쿼리 파라미터가 없으면 기존 출력/다운로드 동작과 100% 동일해야 한다 (opt-in, 회귀 없음).
- 마크업 좌표는 페이지 크기와 무관한 0~1 비율로 저장한다 (기존 `measure`/`search-hit-overlay`와 같은 관례).
- 새 커밋은 매 태스크 끝에 한 번씩, 이 계획의 태스크 경계와 일치하게 만든다.

---

### Task 1: 저장 계층에 `markup` 파생 종류 추가

**Files:**
- Modify: `src/storage.ts:11` (`DerivedKind`), `src/storage.ts:59-62` (`DiskStore.derivedPath`), `src/storage.ts:124-133` (`DiskStore.deleteDocument`), `src/storage.ts:158-171` (`MemoryStore.enforceBudget`), `src/storage.ts:215-220` (`MemoryStore.deleteDocument`)
- Test: `tests/storage-markup.test.ts`

**Interfaces:**
- Consumes: 없음 (기존 `DocumentStore` 인터페이스, `DiskStore`, `MemoryStore`, `newDocumentId`, `ensureDataDirs`를 그대로 씀)
- Produces: `DerivedKind`가 `'pdf' | 'grid' | 'markup'`이 됨. `store.saveDerived(id, 'markup', buf)` / `store.readDerived(id, 'markup')` / `store.hasDerived(id, 'markup')`가 `.markup.json` 파일로 동작. `store.deleteDocument(id)`가 markup 파일도 함께 지움. (다음 태스크의 `src/markup.ts`가 이 API를 그대로 씀)

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/storage-markup.test.ts` 새로 작성:

```ts
import { mkdtempSync, rmSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, test, vi } from 'vitest';

async function loadStoreWithTempDir() {
  const dir = mkdtempSync(path.join(os.tmpdir(), 'docview-storage-test-'));
  vi.resetModules();
  process.env.DOCVIEW_DATA_DIR = dir;
  process.env.DOCVIEW_STORAGE = 'disk';
  const storageModule = await import('../src/storage.js');
  await storageModule.ensureDataDirs();
  return { dir, storageModule };
}

describe('storage: markup derived kind', () => {
  let tempDirs: string[] = [];

  afterEach(() => {
    delete process.env.DOCVIEW_DATA_DIR;
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
    for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true });
    tempDirs = [];
  });

  test('markup and grid derived kinds do not collide on disk', async () => {
    const { dir, storageModule } = await loadStoreWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();
    await store.writeMeta({
      id,
      name: 'a.xlsx',
      ext: 'xlsx',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
    });
    await store.saveDerived(id, 'grid', Buffer.from('{"grid":true}', 'utf8'));
    await store.saveDerived(id, 'markup', Buffer.from('[]', 'utf8'));
    const grid = await store.readDerived(id, 'grid');
    const markup = await store.readDerived(id, 'markup');
    expect(grid.toString('utf8')).toBe('{"grid":true}');
    expect(markup.toString('utf8')).toBe('[]');
  });

  test('hasDerived reports false before save and true after', async () => {
    const { dir, storageModule } = await loadStoreWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();
    await store.writeMeta({
      id,
      name: 'a.pdf',
      ext: 'pdf',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
    });
    expect(await store.hasDerived(id, 'markup')).toBe(false);
    await store.saveDerived(id, 'markup', Buffer.from('[]', 'utf8'));
    expect(await store.hasDerived(id, 'markup')).toBe(true);
  });

  test('deleteDocument removes the markup derived file along with original/meta', async () => {
    const { dir, storageModule } = await loadStoreWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();
    await store.writeMeta({
      id,
      name: 'a.pdf',
      ext: 'pdf',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
    });
    await store.saveOriginal(id, 'pdf', Buffer.from('%PDF-1.4', 'utf8'));
    await store.saveDerived(id, 'markup', Buffer.from('[]', 'utf8'));
    await store.deleteDocument(id);
    expect(await store.readMeta(id)).toBeNull();
    expect(await store.hasDerived(id, 'markup')).toBe(false);
  });
});
```

- [ ] **Step 2: 테스트 실행해서 실패 확인**

Run: `npx vitest run tests/storage-markup.test.ts`
Expected: FAIL — 첫 번째 테스트(`markup and grid derived kinds do not collide`)가 실패한다. 현재 `derivedPath`가 `kind === 'pdf' ? ... : ...`로 `'grid'`와 `'markup'`을 똑같이 `${id}.grid.json`에 매핑하기 때문에, `markup`으로 저장한 `'[]'`이 `grid`로 저장한 값을 덮어써 `grid.toString('utf8')`이 `'{"grid":true}'`가 아니라 `'[]'`로 나온다.

- [ ] **Step 3: 최소 구현**

`src/storage.ts:11`의 `DerivedKind` 수정:

```ts
/** 변환 산출물 종류: pdf(페이지 뷰·출력용) / grid(스프레드시트 그리드 뷰 JSON) / markup(리뷰 코멘트 JSON) */
export type DerivedKind = 'pdf' | 'grid' | 'markup';
```

`src/storage.ts:59-62`의 `DiskStore.derivedPath` 수정:

```ts
  private derivedPath(id: string, kind: DerivedKind): string {
    assertValidId(id);
    const suffix = kind === 'pdf' ? 'pdf' : kind === 'grid' ? 'grid.json' : 'markup.json';
    return path.join(config.convertedDir, `${id}.${suffix}`);
  }
```

`src/storage.ts:124-133`의 `DiskStore.deleteDocument` 수정 (markup 파일 삭제 추가):

```ts
  async deleteDocument(id: string): Promise<void> {
    const meta = await this.readMeta(id);
    if (!meta) return;
    await Promise.allSettled([
      fs.rm(this.originalPath(id, meta.ext), { force: true }),
      fs.rm(this.derivedPath(id, 'pdf'), { force: true }),
      fs.rm(this.derivedPath(id, 'grid'), { force: true }),
      fs.rm(this.derivedPath(id, 'markup'), { force: true }),
      fs.rm(this.metaPath(id), { force: true }),
    ]);
  }
```

`src/storage.ts:158-171`의 `MemoryStore.enforceBudget` 수정 (markup 키 정리 추가):

```ts
  private enforceBudget(protectedId: string): void {
    while (this.totalBytes() > this.budgetBytes) {
      const oldest = sortByCreatedAtDesc([...this.metas.values()])
        .reverse()
        .find((meta) => meta.id !== protectedId);
      if (!oldest) return;
      this.originals.delete(oldest.id);
      this.derived.delete(`${oldest.id}#pdf`);
      this.derived.delete(`${oldest.id}#grid`);
      this.derived.delete(`${oldest.id}#markup`);
      this.metas.delete(oldest.id);
      console.warn(`메모리 예산 초과로 오래된 문서를 제거했습니다: ${oldest.name} (${oldest.id})`);
    }
  }
```

`src/storage.ts:215-220`의 `MemoryStore.deleteDocument` 수정:

```ts
  async deleteDocument(id: string): Promise<void> {
    this.originals.delete(id);
    this.derived.delete(`${id}#pdf`);
    this.derived.delete(`${id}#grid`);
    this.derived.delete(`${id}#markup`);
    this.metas.delete(id);
  }
```

- [ ] **Step 4: 테스트 실행해서 통과 확인**

Run: `npx vitest run tests/storage-markup.test.ts`
Expected: PASS (3 tests)

- [ ] **Step 5: 타입체크**

Run: `npx tsc --noEmit`
Expected: 에러 없음

- [ ] **Step 6: 커밋**

```bash
git add src/storage.ts tests/storage-markup.test.ts
git commit -m "Add markup derived-kind storage path, deletion cleanup"
```

---

### Task 2: 마크업 도메인 모듈 (`src/markup.ts`)

**Files:**
- Create: `src/markup.ts`
- Test: `tests/markup.test.ts`

**Interfaces:**
- Consumes: `store` (`src/storage.ts`, Task 1의 `'markup'` 파생 종류), `DocumentMeta`(`src/types.ts`)
- Produces:
  - `interface MarkupItem { id, page, x, y, w, h, text, author, createdAt }` (모두 `readonly`, `page/x/y/w/h`는 `number`, 나머지 `string`)
  - `interface AddMarkupInput { page, x, y, w, h, text, author }`
  - `class MarkupValidationError extends Error`
  - `listMarkup(docId: string): Promise<readonly MarkupItem[]>`
  - `addMarkup(docId: string, meta: DocumentMeta, input: AddMarkupInput): Promise<MarkupItem>`
  - `deleteMarkup(docId: string, markupId: string): Promise<boolean>` (지워졌으면 true, 없었으면 false)

  다음 태스크(`src/rendering.ts`)가 `MarkupItem`을, 그 다음 태스크(`src/routes/viewer.ts`)가 이 5개를 전부 그대로 가져다 쓴다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/markup.test.ts` 새로 작성:

```ts
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import type { DocumentMeta } from '../src/types.js';

async function freshMarkupModule() {
  vi.resetModules();
  process.env.DOCVIEW_STORAGE = 'memory';
  return import('../src/markup.js');
}

function makeMeta(overrides: Partial<DocumentMeta> = {}): DocumentMeta {
  return {
    id: '11111111-1111-4111-8111-111111111111',
    name: 'sample.pdf',
    ext: 'pdf',
    status: 'ready',
    pageCount: 3,
    createdAt: new Date().toISOString(),
    ...overrides,
  };
}

describe('markup', () => {
  afterEach(() => {
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
  });

  test('listMarkup on a document with no markup returns an empty array', async () => {
    const { listMarkup } = await freshMarkupModule();
    const items = await listMarkup(makeMeta().id);
    expect(items).toEqual([]);
  });

  test('addMarkup stores a new item with a generated id and timestamp, listMarkup returns it', async () => {
    const { addMarkup, listMarkup } = await freshMarkupModule();
    const meta = makeMeta();
    const item = await addMarkup(meta.id, meta, {
      page: 2,
      x: 0.1,
      y: 0.2,
      w: 0.3,
      h: 0.1,
      text: '이 부분을 수정해주세요',
      author: '검토자A',
    });
    expect(item.id).toBeTruthy();
    expect(item.page).toBe(2);
    expect(item.text).toBe('이 부분을 수정해주세요');
    expect(item.author).toBe('검토자A');
    expect(new Date(item.createdAt).toString()).not.toBe('Invalid Date');

    const items = await listMarkup(meta.id);
    expect(items).toHaveLength(1);
    expect(items[0]).toEqual(item);
  });

  test('addMarkup rejects a page number outside the document range', async () => {
    const { addMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta({ pageCount: 3 });
    await expect(
      addMarkup(meta.id, meta, { page: 4, x: 0, y: 0, w: 0.1, h: 0.1, text: 'x', author: 'a' }),
    ).rejects.toThrow(MarkupValidationError);
  });

  test('addMarkup rejects a zero-area rectangle', async () => {
    const { addMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta();
    await expect(
      addMarkup(meta.id, meta, { page: 1, x: 0.5, y: 0.5, w: 0, h: 0.1, text: 'x', author: 'a' }),
    ).rejects.toThrow(MarkupValidationError);
  });

  test('addMarkup rejects a rectangle that overflows the page bounds', async () => {
    const { addMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta();
    await expect(
      addMarkup(meta.id, meta, { page: 1, x: 0.9, y: 0.1, w: 0.5, h: 0.1, text: 'x', author: 'a' }),
    ).rejects.toThrow(MarkupValidationError);
  });

  test('addMarkup rejects empty comment text', async () => {
    const { addMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta();
    await expect(
      addMarkup(meta.id, meta, { page: 1, x: 0, y: 0, w: 0.1, h: 0.1, text: '   ', author: 'a' }),
    ).rejects.toThrow(MarkupValidationError);
  });

  test('deleteMarkup removes the item and returns true; returns false for an unknown id', async () => {
    const { addMarkup, deleteMarkup, listMarkup } = await freshMarkupModule();
    const meta = makeMeta();
    const item = await addMarkup(meta.id, meta, {
      page: 1,
      x: 0,
      y: 0,
      w: 0.1,
      h: 0.1,
      text: 'x',
      author: 'a',
    });
    expect(await deleteMarkup(meta.id, item.id)).toBe(true);
    expect(await listMarkup(meta.id)).toEqual([]);
    expect(await deleteMarkup(meta.id, item.id)).toBe(false);
  });

  test('two markups added concurrently on the same document are both persisted (no lost update)', async () => {
    const { addMarkup, listMarkup } = await freshMarkupModule();
    const meta = makeMeta();
    await Promise.all([
      addMarkup(meta.id, meta, { page: 1, x: 0, y: 0, w: 0.1, h: 0.1, text: 'one', author: 'a' }),
      addMarkup(meta.id, meta, { page: 1, x: 0.2, y: 0.2, w: 0.1, h: 0.1, text: 'two', author: 'b' }),
    ]);
    const items = await listMarkup(meta.id);
    expect(items.map((i) => i.text).sort()).toEqual(['one', 'two']);
  });
});
```

- [ ] **Step 2: 테스트 실행해서 실패 확인**

Run: `npx vitest run tests/markup.test.ts`
Expected: FAIL — `src/markup.ts` 파일이 없어 `Cannot find module '../src/markup.js'`

- [ ] **Step 3: 최소 구현**

`src/markup.ts` 새로 작성:

```ts
import { randomUUID } from 'node:crypto';
import { store } from './storage.js';
import type { DocumentMeta } from './types.js';

export interface MarkupItem {
  readonly id: string;
  readonly page: number;
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly text: string;
  readonly author: string;
  readonly createdAt: string;
}

export interface AddMarkupInput {
  readonly page: number;
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly text: string;
  readonly author: string;
}

export class MarkupValidationError extends Error {}

const MAX_TEXT_LENGTH = 2000;
/** 좌표 비교 시 부동소수점 오차를 흡수하기 위한 여유값 */
const BOUNDS_EPSILON = 0.0001;

function assertValidInput(input: AddMarkupInput, pageCount: number): void {
  if (!Number.isInteger(input.page) || input.page < 1 || input.page > pageCount) {
    throw new MarkupValidationError(`잘못된 페이지 번호입니다: ${input.page} (전체 ${pageCount}페이지)`);
  }
  for (const [key, value] of Object.entries({ x: input.x, y: input.y, w: input.w, h: input.h })) {
    if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || value > 1) {
      throw new MarkupValidationError(`잘못된 좌표값입니다: ${key}=${value}`);
    }
  }
  if (input.w <= 0 || input.h <= 0) {
    throw new MarkupValidationError('마크업 영역의 폭과 높이는 0보다 커야 합니다.');
  }
  if (input.x + input.w > 1 + BOUNDS_EPSILON || input.y + input.h > 1 + BOUNDS_EPSILON) {
    throw new MarkupValidationError('마크업 영역이 페이지를 벗어납니다.');
  }
  if (input.text.trim().length === 0) {
    throw new MarkupValidationError('코멘트 내용을 입력해 주세요.');
  }
  if (input.text.trim().length > MAX_TEXT_LENGTH) {
    throw new MarkupValidationError(`코멘트는 ${MAX_TEXT_LENGTH}자를 넘을 수 없습니다.`);
  }
}

/**
 * 문서 ID별 순차 실행 큐. 같은 문서에 대한 추가/삭제가 동시에 들어와도
 * "읽기 → 수정 → 쓰기"가 서로 겹쳐 한쪽의 변경을 덮어쓰지 않게 한다.
 * (src/converters/libreOffice.ts의 프로필 슬롯 큐와 동일한 패턴)
 */
const docLocks = new Map<string, Promise<unknown>>();

async function withDocLock<T>(docId: string, task: () => Promise<T>): Promise<T> {
  const previous = docLocks.get(docId) ?? Promise.resolve();
  let release!: () => void;
  const current = new Promise<void>((resolve) => {
    release = resolve;
  });
  docLocks.set(docId, previous.then(() => current));
  await previous.catch(() => {});
  try {
    return await task();
  } finally {
    release();
  }
}

async function readAll(docId: string): Promise<MarkupItem[]> {
  if (!(await store.hasDerived(docId, 'markup'))) return [];
  const buffer = await store.readDerived(docId, 'markup');
  return JSON.parse(buffer.toString('utf8')) as MarkupItem[];
}

async function writeAll(docId: string, items: readonly MarkupItem[]): Promise<void> {
  await store.saveDerived(docId, 'markup', Buffer.from(JSON.stringify(items), 'utf8'));
}

export async function listMarkup(docId: string): Promise<readonly MarkupItem[]> {
  return readAll(docId);
}

export async function addMarkup(
  docId: string,
  meta: DocumentMeta,
  input: AddMarkupInput,
): Promise<MarkupItem> {
  assertValidInput(input, meta.pageCount);
  return withDocLock(docId, async () => {
    const items = await readAll(docId);
    const item: MarkupItem = {
      id: randomUUID(),
      page: input.page,
      x: input.x,
      y: input.y,
      w: input.w,
      h: input.h,
      text: input.text.trim(),
      author: input.author,
      createdAt: new Date().toISOString(),
    };
    await writeAll(docId, [...items, item]);
    return item;
  });
}

export async function deleteMarkup(docId: string, markupId: string): Promise<boolean> {
  return withDocLock(docId, async () => {
    const items = await readAll(docId);
    const next = items.filter((item) => item.id !== markupId);
    if (next.length === items.length) return false;
    await writeAll(docId, next);
    return true;
  });
}
```

- [ ] **Step 4: 테스트 실행해서 통과 확인**

Run: `npx vitest run tests/markup.test.ts`
Expected: PASS (8 tests)

- [ ] **Step 5: 타입체크**

Run: `npx tsc --noEmit`
Expected: 에러 없음

- [ ] **Step 6: 커밋**

```bash
git add src/markup.ts tests/markup.test.ts
git commit -m "Add markup domain module: CRUD, validation, per-doc write lock"
```

---

### Task 3: PDF에 마크업 그려 넣기 (`src/rendering.ts`)

**Files:**
- Modify: `src/rendering.ts` (기존 `buildPrintPdf` 아래에 함수 추가)
- Test: `tests/rendering-markup.test.ts`

**Interfaces:**
- Consumes: `MarkupItem`(Task 2의 `src/markup.ts`), `pdf-lib`의 `PDFDocument`/`StandardFonts`/`rgb` (이미 `rendering.ts`에서 import 중)
- Produces: `burnMarkupIntoPdf(pdfBytes: Buffer, items: readonly MarkupItem[]): Promise<Uint8Array>` — 다음 태스크(`src/routes/viewer.ts`)가 출력/다운로드 라우트에서 그대로 씀

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/rendering-markup.test.ts` 새로 작성. 최소 유효 PDF를 `pdf-lib`로 직접 만들어 씀(다른 fixture 불필요):

```ts
import { PDFDocument } from 'pdf-lib';
import { describe, expect, test } from 'vitest';
import { burnMarkupIntoPdf } from '../src/rendering.js';
import type { MarkupItem } from '../src/markup.js';

async function makeTwoPagePdf(): Promise<Buffer> {
  const doc = await PDFDocument.create();
  doc.addPage([200, 300]);
  doc.addPage([200, 300]);
  return Buffer.from(await doc.save());
}

function makeItem(overrides: Partial<MarkupItem> = {}): MarkupItem {
  return {
    id: 'm1',
    page: 1,
    x: 0.1,
    y: 0.1,
    w: 0.3,
    h: 0.1,
    text: '이 부분을 수정해주세요',
    author: '검토자A',
    createdAt: new Date().toISOString(),
    ...overrides,
  };
}

describe('burnMarkupIntoPdf', () => {
  test('returns a valid PDF with the same page count when no items are given', async () => {
    const original = await makeTwoPagePdf();
    const result = await burnMarkupIntoPdf(original, []);
    const reopened = await PDFDocument.load(result);
    expect(reopened.getPageCount()).toBe(2);
  });

  test('burns a markup item into the correct page without changing page count', async () => {
    const original = await makeTwoPagePdf();
    const result = await burnMarkupIntoPdf(original, [makeItem({ page: 2 })]);
    const reopened = await PDFDocument.load(result);
    expect(reopened.getPageCount()).toBe(2);
    // 원본보다 산출물이 확실히 커야 한다(사각형+텍스트가 실제로 그려졌다는 근거)
    expect(result.length).toBeGreaterThan(original.length);
  });

  test('handles multiple items across multiple pages', async () => {
    const original = await makeTwoPagePdf();
    const items = [
      makeItem({ id: 'm1', page: 1, text: '첫 페이지 코멘트' }),
      makeItem({ id: 'm2', page: 2, text: '둘째 페이지 코멘트' }),
    ];
    const result = await burnMarkupIntoPdf(original, items);
    const reopened = await PDFDocument.load(result);
    expect(reopened.getPageCount()).toBe(2);
  });

  test('ignores an item whose page number is out of range instead of throwing', async () => {
    const original = await makeTwoPagePdf();
    await expect(burnMarkupIntoPdf(original, [makeItem({ page: 99 })])).resolves.toBeInstanceOf(
      Uint8Array,
    );
  });
});
```

- [ ] **Step 2: 테스트 실행해서 실패 확인**

Run: `npx vitest run tests/rendering-markup.test.ts`
Expected: FAIL — `burnMarkupIntoPdf`가 `src/rendering.ts`에 없어 `SyntaxError`/`does not provide an export named 'burnMarkupIntoPdf'`

- [ ] **Step 3: 최소 구현**

`src/rendering.ts` 맨 끝(`buildPrintPdf` 함수 뒤)에 추가:

```ts
import type { MarkupItem } from './markup.js';

/**
 * 마크업(사각형+코멘트)을 PDF에 그려 넣는다. 원본 파일은 건드리지 않고, 출력/다운로드용으로
 * 새로 만드는 PDF 바이트에만 적용한다(호출부: src/routes/viewer.ts).
 * 좌표는 페이지 크기와 무관한 0~1 비율(src/markup.ts)이라 이 함수가 실제 pt 단위로 환산한다.
 */
export async function burnMarkupIntoPdf(
  pdfBytes: Buffer,
  items: readonly MarkupItem[],
): Promise<Uint8Array> {
  const doc = await PDFDocument.load(toOwnedBytes(pdfBytes), { ignoreEncryption: true });
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const pages = doc.getPages();

  const byPage = new Map<number, MarkupItem[]>();
  for (const item of items) {
    const list = byPage.get(item.page) ?? [];
    list.push(item);
    byPage.set(item.page, list);
  }

  const fontSize = 8;
  const lineHeight = fontSize * 1.3;
  const boxPad = 3;
  const maxCharsPerLine = 42;

  for (const [pageNumber, pageItems] of byPage) {
    const page = pages[pageNumber - 1];
    if (!page) continue; // 범위를 벗어난 페이지 번호는 조용히 건너뛴다
    const { width, height } = page.getSize();

    for (const item of pageItems) {
      const boxX = item.x * width;
      const boxW = item.w * width;
      const boxH = item.h * height;
      // pdf-lib 좌표계는 좌하단이 원점이라, 위에서부터 재는 비율 y를 뒤집는다.
      const boxY = height - (item.y + item.h) * height;

      page.drawRectangle({
        x: boxX,
        y: boxY,
        width: boxW,
        height: boxH,
        borderColor: rgb(0.91, 0.42, 0.02),
        borderWidth: 1.5,
        color: rgb(0.91, 0.42, 0.02),
        opacity: 0.08,
        borderOpacity: 0.9,
      });

      const text = toWinAnsiSafe(item.text);
      const lines: string[] = [];
      for (let i = 0; i < text.length; i += maxCharsPerLine) {
        lines.push(text.slice(i, i + maxCharsPerLine));
      }
      const textBlockHeight = lines.length * lineHeight + boxPad * 2;
      const textY = Math.max(0, boxY - textBlockHeight);
      const textWidthPx =
        Math.max(boxW, ...lines.map((line) => font.widthOfTextAtSize(line, fontSize))) + boxPad * 2;

      page.drawRectangle({
        x: boxX,
        y: textY,
        width: textWidthPx,
        height: textBlockHeight,
        color: rgb(1, 0.97, 0.88),
        borderColor: rgb(0.91, 0.42, 0.02),
        borderWidth: 1,
        opacity: 0.95,
      });

      lines.forEach((line, i) => {
        page.drawText(line, {
          x: boxX + boxPad,
          y: textY + textBlockHeight - boxPad - (i + 1) * lineHeight + (lineHeight - fontSize) / 2,
          size: fontSize,
          font,
          color: rgb(0.2, 0.15, 0.02),
        });
      });
    }
  }

  return doc.save();
}
```

`src/rendering.ts` 최상단의 import에 `MarkupItem` 타입 import를 추가(파일 맨 위, 기존 `import { PDFDocument, StandardFonts, degrees, rgb } from 'pdf-lib';` 다음 줄):

```ts
import type { MarkupItem } from './markup.js';
```

(주의: 위 함수 코드 블록에 있는 `import type { MarkupItem } from './markup.js';`는 실제로는 파일 맨 위 import 구역에 넣는다 — 함수 본문 중간이 아니라 파일 최상단에 다른 import들과 함께 둔다.)

- [ ] **Step 4: 테스트 실행해서 통과 확인**

Run: `npx vitest run tests/rendering-markup.test.ts`
Expected: PASS (4 tests)

- [ ] **Step 5: 타입체크**

Run: `npx tsc --noEmit`
Expected: 에러 없음

- [ ] **Step 6: 커밋**

```bash
git add src/rendering.ts tests/rendering-markup.test.ts
git commit -m "Add burnMarkupIntoPdf: render markup rectangles+comments onto a PDF"
```

---

### Task 4: API 라우트 (`src/routes/viewer.ts`)

**Files:**
- Modify: `src/routes/viewer.ts`

**Interfaces:**
- Consumes: `listMarkup`/`addMarkup`/`deleteMarkup`/`MarkupValidationError`(Task 2), `burnMarkupIntoPdf`(Task 3), 기존 `resolveView`/`recordAudit`/`ensurePdf`/`buildPrintPdf`
- Produces:
  - `GET /api/view/markup?token=` → `{success, data: MarkupItem[]}`
  - `POST /api/view/markup?token=` body `{page,x,y,w,h,text}` → `{success, data: MarkupItem}` (201) — `canEdit` 필요
  - `DELETE /api/view/markup/:markupId?token=` → `{success, data:{id}}` — `canEdit` 필요
  - `GET /api/view/print?token=&includeMarkup=1`, `GET /api/view/download?token=&includeMarkup=1` — 기존 라우트에 opt-in 파라미터 추가, 생략 시 기존과 동일

이 태스크는 이 저장소에 라우트 레벨 자동 테스트 선례가 없어(기존 `documents.ts`/`viewer.ts` 라우트들도 자동 테스트가 없고 수동/Playwright로만 검증됨), Task 2/3에서 이미 비즈니스 로직을 커버했다는 전제로 TDD 없이 진행하고 Step 마지막에 실행 중인 서버로 직접 curl 검증한다.

- [ ] **Step 1: import 추가**

`src/routes/viewer.ts` 상단 import 블록(9번째 줄 `import { PageOutOfRangeError, buildPrintPdf, getPageDimensions, renderPage } from '../rendering.js';` 를 아래로 교체, 그리고 새 줄 추가):

```ts
import {
  PageOutOfRangeError,
  buildPrintPdf,
  burnMarkupIntoPdf,
  getPageDimensions,
  renderPage,
} from '../rendering.js';
```

그 아래 다른 import들 사이(예: `docxHasUnsupportedContent` import 다음 줄)에 추가:

```ts
import { MarkupValidationError, addMarkup, deleteMarkup, listMarkup } from '../markup.js';
```

- [ ] **Step 2: 스키마·타입 추가**

`src/routes/viewer.ts`의 기존 `tokenQuerySchema`/`TokenQuery` 정의(약 196~199번째 줄, 226~228번째 줄) 바로 아래에 추가:

```ts
const tokenWithMarkupQuerySchema = {
  type: 'object',
  properties: { token: { type: 'string' }, includeMarkup: { type: 'string' } },
} as const;

interface TokenWithMarkupQuery extends TokenQuery {
  readonly includeMarkup?: string;
}

const addMarkupBodySchema = {
  type: 'object',
  required: ['page', 'x', 'y', 'w', 'h', 'text'],
  additionalProperties: false,
  properties: {
    page: { type: 'integer', minimum: 1 },
    x: { type: 'number', minimum: 0, maximum: 1 },
    y: { type: 'number', minimum: 0, maximum: 1 },
    w: { type: 'number', minimum: 0, maximum: 1 },
    h: { type: 'number', minimum: 0, maximum: 1 },
    text: { type: 'string', minLength: 1, maxLength: 2000 },
  },
} as const;

interface AddMarkupBody {
  readonly page: number;
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly text: string;
}

const markupParamsSchema = {
  type: 'object',
  required: ['markupId'],
  properties: { markupId: { type: 'string' } },
} as const;
```

- [ ] **Step 3: 마크업 CRUD 라우트 추가**

`src/routes/viewer.ts`의 `/api/view/search` 라우트(376~397번째 줄) 끝, `/api/view/print` 라우트 시작 전에 삽입:

```ts
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/markup',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      const items = await listMarkup(view.meta.id);
      return { success: true, data: items };
    },
  );

  app.post<{ Querystring: TokenQuery; Body: AddMarkupBody }>(
    '/api/view/markup',
    { schema: { querystring: tokenQuerySchema, body: addMarkupBodySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canEdit) {
        await recordAudit({ type: 'markup_add_denied', docId: view.meta.id, viewer: view.claims.viewer });
        return reply.code(403).send({ success: false, error: '마크업 추가 권한이 없습니다.' });
      }
      if ((view.meta.viewKind ?? 'pages') !== 'pages') {
        return reply.code(400).send({ success: false, error: '이 문서 형식은 마크업을 지원하지 않습니다.' });
      }
      try {
        const item = await addMarkup(view.meta.id, view.meta, {
          ...request.body,
          author: view.claims.viewer,
        });
        await recordAudit({ type: 'markup_added', docId: view.meta.id, viewer: view.claims.viewer, detail: item.id });
        return await reply.code(201).send({ success: true, data: item });
      } catch (error) {
        if (error instanceof MarkupValidationError) {
          return reply.code(400).send({ success: false, error: error.message });
        }
        request.log.error(error, 'markup add failed');
        return reply.code(500).send({ success: false, error: '마크업 추가에 실패했습니다.' });
      }
    },
  );

  app.delete<{ Querystring: TokenQuery; Params: { markupId: string } }>(
    '/api/view/markup/:markupId',
    { schema: { querystring: tokenQuerySchema, params: markupParamsSchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canEdit) {
        await recordAudit({ type: 'markup_delete_denied', docId: view.meta.id, viewer: view.claims.viewer });
        return reply.code(403).send({ success: false, error: '마크업 삭제 권한이 없습니다.' });
      }
      const deleted = await deleteMarkup(view.meta.id, request.params.markupId);
      if (!deleted) {
        return reply.code(404).send({ success: false, error: '마크업을 찾을 수 없습니다.' });
      }
      await recordAudit({
        type: 'markup_deleted',
        docId: view.meta.id,
        viewer: view.claims.viewer,
        detail: request.params.markupId,
      });
      return { success: true, data: { id: request.params.markupId } };
    },
  );

```

- [ ] **Step 4: 출력 라우트에 `includeMarkup` 추가**

`src/routes/viewer.ts`의 기존 `/api/view/print` 라우트(399~424번째 줄)를 통째로 교체:

```ts
  app.get<{ Querystring: TokenWithMarkupQuery }>(
    '/api/view/print',
    { schema: { querystring: tokenWithMarkupQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canPrint) {
        await recordAudit({ type: 'print_denied', docId: view.meta.id, viewer: view.claims.viewer });
        return reply.code(403).send({ success: false, error: '출력 권한이 없습니다.' });
      }
      let pdfSource: Buffer;
      try {
        pdfSource = await ensurePdf(view.meta);
      } catch (error) {
        request.log.error(error, 'print pdf generation failed');
        return reply.code(500).send({ success: false, error: '출력용 PDF 생성에 실패했습니다.' });
      }
      let pdfBytes: Buffer | Uint8Array = pdfSource;
      if (request.query.includeMarkup === '1') {
        const items = await listMarkup(view.meta.id);
        if (items.length > 0) {
          pdfBytes = await burnMarkupIntoPdf(Buffer.from(pdfBytes), items);
        }
      }
      const watermark = `${view.claims.viewer}  ${new Date().toISOString().slice(0, 10)}`;
      const finalBytes = await buildPrintPdf(Buffer.from(pdfBytes), watermark);
      await recordAudit({ type: 'printed', docId: view.meta.id, viewer: view.claims.viewer });
      return reply
        .type('application/pdf')
        .header('content-disposition', 'inline; filename="print.pdf"')
        .send(Buffer.from(finalBytes));
    },
  );
```

- [ ] **Step 5: 다운로드 라우트에 `includeMarkup` 추가**

`src/routes/viewer.ts`의 기존 `/api/view/download` 라우트(426~447번째 줄)를 통째로 교체:

```ts
  app.get<{ Querystring: TokenWithMarkupQuery }>(
    '/api/view/download',
    { schema: { querystring: tokenWithMarkupQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canDownload) {
        await recordAudit({ type: 'download_denied', docId: view.meta.id, viewer: view.claims.viewer });
        return reply.code(403).send({ success: false, error: '저장(다운로드) 권한이 없습니다.' });
      }

      if (request.query.includeMarkup === '1') {
        const items = await listMarkup(view.meta.id);
        if (items.length > 0) {
          let pdfSource: Buffer;
          try {
            pdfSource = await ensurePdf(view.meta);
          } catch (error) {
            request.log.error(error, 'markup download pdf generation failed');
            return reply.code(500).send({ success: false, error: '마크업 포함 PDF 생성에 실패했습니다.' });
          }
          const merged = await burnMarkupIntoPdf(pdfSource, items);
          await recordAudit({
            type: 'downloaded',
            docId: view.meta.id,
            viewer: view.claims.viewer,
            detail: 'with markup',
          });
          const baseName = view.meta.name.replace(/\.[^./]+$/, '');
          const encodedName = encodeURIComponent(`${baseName}-markup.pdf`);
          return reply
            .type('application/pdf')
            .header('content-disposition', `attachment; filename*=UTF-8''${encodedName}`)
            .send(Buffer.from(merged));
        }
      }

      // 지금 이 문서를 편집 중인 동시편집 세션이 떠 있으면(=유휴 자동저장 8초 debounce가 아직 안
      // 지났을 수 있음) 디스크의 원본이 최신이 아닐 수 있다 — 그 경우 메모리의 Y.Doc에서 바로 최신
      // 바이트를 만든다. 세션이 없으면(이미 디스크에 반영 완료) 원본을 그대로 쓴다.
      const data = (await getCurrentBytes(view.meta.id)) ?? (await store.readOriginal(view.meta.id, view.meta.ext));
      await recordAudit({ type: 'downloaded', docId: view.meta.id, viewer: view.claims.viewer });
      const encodedName = encodeURIComponent(view.meta.name);
      return reply
        .type('application/octet-stream')
        .header('content-disposition', `attachment; filename*=UTF-8''${encodedName}`)
        .send(data);
    },
  );
```

- [ ] **Step 6: 타입체크**

Run: `npx tsc --noEmit`
Expected: 에러 없음

- [ ] **Step 7: 실행 중인 개발 서버로 수동 curl 검증**

서버가 안 떠 있으면 `npm run dev`로 띄운다(백그라운드). 그다음:

```bash
# 1. pdf/이미지 문서 하나 업로드해서 문서 id를 얻는다 (이미 업로드된 pages-view 문서가 있으면 재사용)
curl -s -X POST -F "file=@<로컬 pdf 경로>;filename=markup-test.pdf" http://localhost:8000/api/documents

# 2. canEdit:true 토큰 발급 (문서 id를 위 응답의 data.id로 치환)
curl -s -X POST http://localhost:8000/api/documents/<docId>/tokens \
  -H "content-type: application/json" \
  -d '{"viewer":"검토자A","canPrint":true,"canDownload":true,"canCopy":true,"canEdit":true}'

# 3. 마크업 추가 (위 응답의 data.token으로 치환)
curl -s -X POST "http://localhost:8000/api/view/markup?token=<token>" \
  -H "content-type: application/json" \
  -d '{"page":1,"x":0.1,"y":0.1,"w":0.3,"h":0.1,"text":"이 부분을 수정해주세요"}'

# 4. 목록 조회 — 방금 추가한 항목이 보여야 한다
curl -s "http://localhost:8000/api/view/markup?token=<token>"

# 5. includeMarkup=1로 출력 PDF 생성 — 응답이 정상 PDF(바이너리, %PDF로 시작)여야 한다
curl -s "http://localhost:8000/api/view/print?token=<token>&includeMarkup=1" -o /tmp/print-with-markup.pdf
head -c 8 /tmp/print-with-markup.pdf

# 6. 삭제 (위 3번 응답의 data.id로 치환)
curl -s -X DELETE "http://localhost:8000/api/view/markup/<markupId>?token=<token>"

# 7. 목록 다시 조회 — 빈 배열이어야 한다
curl -s "http://localhost:8000/api/view/markup?token=<token>"
```

Expected: 3번 201 + 생성된 항목, 4번에 그 항목 포함, 5번 파일이 `%PDF`로 시작, 6번 `{"success":true,...}`, 7번 `{"success":true,"data":[]}`. 검증이 끝나면 curl로 만든 테스트 문서는 `DELETE /api/documents/<docId>`로 정리한다.

- [ ] **Step 8: 커밋**

```bash
git add src/routes/viewer.ts
git commit -m "Add markup CRUD routes and includeMarkup option to print/download"
```

---

### Task 5: 뷰어 HTML — 툴바 버튼 추가

**Files:**
- Modify: `public/viewer.html`

**Interfaces:**
- Consumes: 없음 (정적 HTML)
- Produces: `#markup-toggle-btn`, `#markup-add-btn`, `#markup-include-wrap`(내부에 `#markup-include-checkbox`) — Task 7의 `public/js/viewer.js`가 `document.getElementById`로 참조

- [ ] **Step 1: 툴바 버튼 추가**

`public/viewer.html`의 `#invert-btn` 버튼(33~41번째 줄) 바로 다음에 삽입:

```html
      <button id="markup-toggle-btn" class="icon-btn" title="마크업 보기/숨기기" aria-label="마크업 보기 전환" aria-pressed="true">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="16" height="16" rx="2"/><path d="M9 9h6v6H9z"/></svg>
      </button>
      <button id="markup-add-btn" class="icon-btn" title="마크업 추가 (영역을 드래그)" aria-label="마크업 추가" aria-pressed="false" hidden>
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="16" height="16" rx="2"/><path d="M12 9v6"/><path d="M9 12h6"/></svg>
      </button>
```

같은 파일의 `#print-btn` 버튼(42~45번째 줄) 바로 앞에 삽입:

```html
      <label id="markup-include-wrap" class="markup-include-toggle" hidden>
        <input type="checkbox" id="markup-include-checkbox" /> 마크업 포함
      </label>
```

- [ ] **Step 2: 브라우저에서 육안 확인**

`npm run dev`로 서버가 떠 있으면(안 떠 있으면 띄운다), 아무 pages-view 문서를 열어 툴바에 새 버튼 2개(마크업 보기 토글, 숨겨진 마크업 추가)와 체크박스(숨김 상태)가 DOM에 존재하는지 브라우저 개발자 도구로 확인한다. 아직 JS를 안 붙였으므로 클릭해도 동작은 없다 — 존재 여부만 확인.

- [ ] **Step 3: 커밋**

```bash
git add public/viewer.html
git commit -m "Add markup toolbar buttons and include-markup checkbox to viewer.html"
```

---

### Task 6: 뷰어 CSS — 마크업 스타일

**Files:**
- Modify: `public/css/style.css`

**Interfaces:**
- Consumes: 없음
- Produces: `.markup-layer`, `.markup-rect`, `.markup-rect-draft`, `.markup-tooltip*`, `.markup-popover*`, `.markup-include-toggle`, `.pages-container.markup-drawing`, `.icon-btn[hidden]` — Task 7의 JS가 이 클래스들을 토글

- [ ] **Step 1: CSS 추가**

`public/css/style.css`의 `.measure-label { ... }` 규칙(716~730번째 줄) 바로 다음에 새 섹션 추가:

```css
/* ---------- 마크업(리뷰 코멘트) ---------- */

/* .icon-btn이 display:inline-flex를 선언해 브라우저 기본 [hidden]{display:none}보다
   우선해버리는 문제 — 이 파일의 .toolbar-group[hidden]/.editor-toolbar[hidden]과 같은 이유로
   markup-add-btn을 canEdit 아닐 때 실제로 숨기려면 이 규칙이 필요하다. */
.icon-btn[hidden] { display: none; }

.pages-container.markup-drawing,
.pages-container.markup-drawing .page-frame { cursor: crosshair; }

.markup-layer {
  position: absolute;
  inset: 0;
  z-index: 6;
}

.markup-rect {
  position: absolute;
  border: 1.5px solid #e8590c;
  background: rgba(232, 89, 12, 0.12);
  border-radius: 2px;
  pointer-events: auto;
  cursor: pointer;
}

.markup-rect-draft {
  pointer-events: none;
  border-style: dashed;
}

.markup-tooltip { display: none; }
.markup-rect:hover .markup-tooltip,
.markup-rect:focus-within .markup-tooltip { display: flex; }

.markup-tooltip {
  position: absolute;
  left: 0;
  top: 100%;
  margin-top: 6px;
  min-width: 160px;
  max-width: 260px;
  flex-direction: column;
  gap: 4px;
  padding: 0.5rem 0.6rem;
  background: #1c1b1a;
  color: #fff;
  border-radius: 8px;
  font-size: 0.78rem;
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.4);
  z-index: 7;
}

.markup-tooltip-meta { color: rgba(255, 255, 255, 0.6); font-size: 0.7rem; }
.markup-tooltip-text { white-space: pre-wrap; word-break: break-word; }

.markup-tooltip-delete {
  align-self: flex-end;
  border: none;
  background: transparent;
  color: #ff8787;
  font-size: 0.72rem;
  cursor: pointer;
  padding: 0.15rem 0.3rem;
}
.markup-tooltip-delete:hover { text-decoration: underline; }

.markup-popover {
  position: absolute;
  z-index: 8;
  width: 220px;
  display: flex;
  flex-direction: column;
  gap: 0.4rem;
  padding: 0.5rem;
  background: var(--v-glass);
  border: 1px solid var(--v-glass-line);
  border-radius: 10px;
  backdrop-filter: blur(20px) saturate(1.5);
  -webkit-backdrop-filter: blur(20px) saturate(1.5);
  box-shadow: 0 12px 40px rgba(0, 0, 0, 0.45);
}

.markup-popover textarea {
  width: 100%;
  min-height: 64px;
  resize: vertical;
  font-family: var(--font-sans);
  font-size: 0.8rem;
  padding: 0.4rem;
  border-radius: 6px;
  border: 1px solid var(--v-glass-line);
  background: rgba(255, 255, 255, 0.06);
  color: var(--v-text);
}

.markup-popover-actions {
  display: flex;
  justify-content: flex-end;
  gap: 0.4rem;
}

.markup-include-toggle {
  display: flex;
  align-items: center;
  gap: 0.3rem;
  font-size: 0.78rem;
  color: var(--v-text-dim);
  white-space: nowrap;
}
.markup-include-toggle[hidden] { display: none; }
```

- [ ] **Step 2: 브라우저에서 육안 확인**

문서를 다시 열어 "마크업 추가" 버튼이 (아직 JS가 canEdit 체크를 안 하므로) `hidden` 상태로 실제로 안 보이는지 확인한다 — `.icon-btn[hidden]` 규칙이 없으면 이 시점에 버튼이 계속 보이는 회귀가 생긴다.

- [ ] **Step 3: 커밋**

```bash
git add public/css/style.css
git commit -m "Add markup overlay/popover/tooltip styles to viewer"
```

---

### Task 7: 뷰어 JS — 마크업 인터랙션

**Files:**
- Modify: `public/js/viewer.js`

**Interfaces:**
- Consumes: Task 4의 3개 API, Task 5의 DOM 요소, Task 6의 CSS 클래스, 기존 `fetchJsonView`/`tokenQuery`/`pagesContainer`/`permissions`
- Produces: `markup` 상태 객체, `loadMarkup()`, `setMarkupVisible()`, `setMarkupDrawMode()` — `init()`이 호출

- [ ] **Step 1: `fetchJsonView`를 POST/DELETE도 지원하도록 확장**

`public/js/viewer.js`의 기존 함수(631~636번째 줄)를 교체:

```js
async function fetchJsonView(url, options = {}) {
  const res = await fetch(url, { signal: viewAbort.signal, ...options });
  const json = await res.json();
  if (!res.ok || !json.success) throw new Error(json.error || `요청 실패 (HTTP ${res.status})`);
  return json.data;
}
```

(기존 호출부는 전부 두 번째 인자 없이 부르므로 그대로 동작한다.)

- [ ] **Step 2: 마크업 상태·렌더링 함수 추가**

`public/js/viewer.js`의 `measure` 관련 코드 블록 전체(795~907번째 줄, `/* ==================== 도면 도구: 커서 기준 줌 · 패닝 · 거리 측정 ==================== */`부터 `setMeasureMode` 함수 끝까지) 바로 다음, `setupDrawingTools` 함수(909번째 줄) 시작 전에 새 블록을 삽입:

```js
/* ==================== 마크업(리뷰 코멘트) ==================== */

const markup = {
  visible: true,
  drawMode: false,
  items: [],
  dragging: null, // {frame, page, fx, fy, rectEl, box}
  draftRect: null, // 팝오버가 열려 있는 동안 유지되는 사각형 엘리먼트
};

function markupLayer(frame) {
  let layer = frame.querySelector('.markup-layer');
  if (!layer) {
    layer = document.createElement('div');
    layer.className = 'markup-layer';
    frame.append(layer);
  }
  return layer;
}

function clearMarkupOverlays() {
  for (const frame of pagesContainer.children) {
    const layer = frame.querySelector('.markup-layer');
    if (layer) layer.replaceChildren();
  }
}

async function deleteMarkupItem(item, rectEl) {
  if (!confirm('이 마크업을 삭제할까요?')) return;
  try {
    await fetchJsonView(`/api/view/markup/${item.id}?${tokenQuery}`, { method: 'DELETE' });
    markup.items = markup.items.filter((it) => it.id !== item.id);
    rectEl.remove();
    updateMarkupIncludeVisibility();
  } catch (error) {
    alert(error.message || '삭제에 실패했습니다.');
  }
}

function renderMarkupItem(item) {
  const frame = pagesContainer.children[item.page - 1];
  if (!frame) return;
  const layer = markupLayer(frame);
  const rectEl = document.createElement('div');
  rectEl.className = 'markup-rect';
  rectEl.style.left = `${item.x * 100}%`;
  rectEl.style.top = `${item.y * 100}%`;
  rectEl.style.width = `${item.w * 100}%`;
  rectEl.style.height = `${item.h * 100}%`;
  rectEl.tabIndex = 0;
  rectEl.dataset.markupId = item.id;

  const tooltip = document.createElement('div');
  tooltip.className = 'markup-tooltip';
  const metaLine = document.createElement('div');
  metaLine.className = 'markup-tooltip-meta';
  metaLine.textContent = `${item.author} · ${new Date(item.createdAt).toLocaleString('ko-KR')}`;
  const body = document.createElement('div');
  body.className = 'markup-tooltip-text';
  body.textContent = item.text;
  tooltip.append(metaLine, body);

  if (permissions.canEdit) {
    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'markup-tooltip-delete';
    del.textContent = '삭제';
    del.addEventListener('click', (event) => {
      event.stopPropagation();
      deleteMarkupItem(item, rectEl);
    });
    tooltip.append(del);
  }

  rectEl.append(tooltip);
  layer.append(rectEl);
}

function renderAllMarkup() {
  clearMarkupOverlays();
  if (!markup.visible) return;
  for (const item of markup.items) renderMarkupItem(item);
}

function updateMarkupIncludeVisibility() {
  document.getElementById('markup-include-wrap').hidden = markup.items.length === 0;
}

async function loadMarkup() {
  try {
    markup.items = await fetchJsonView(`/api/view/markup?${tokenQuery}`);
    renderAllMarkup();
    updateMarkupIncludeVisibility();
  } catch {
    /* 마크업 로딩 실패는 열람 자체를 막지 않는다 */
  }
}

function setMarkupVisible(visible) {
  markup.visible = visible;
  document.getElementById('markup-toggle-btn').setAttribute('aria-pressed', String(visible));
  renderAllMarkup();
}

function closeMarkupPopover() {
  const existing = document.querySelector('.markup-popover');
  if (existing) existing.remove();
}

function cancelMarkupDraft() {
  if (markup.dragging) {
    markup.dragging.rectEl.remove();
    markup.dragging = null;
  }
  if (markup.draftRect) {
    markup.draftRect.remove();
    markup.draftRect = null;
  }
  closeMarkupPopover();
}

function setMarkupDrawMode(on) {
  markup.drawMode = on;
  document.getElementById('markup-add-btn').setAttribute('aria-pressed', String(on));
  pagesContainer.classList.toggle('markup-drawing', on);
  if (!on) cancelMarkupDraft();
}

function openMarkupPopover(frame, page, box) {
  closeMarkupPopover();
  const popover = document.createElement('div');
  popover.className = 'markup-popover';
  popover.style.left = `${Math.min(box.x + box.w, 0.7) * 100}%`;
  popover.style.top = `${box.y * 100}%`;

  const textarea = document.createElement('textarea');
  textarea.placeholder = '코멘트를 입력하세요';
  textarea.maxLength = 2000;

  const actions = document.createElement('div');
  actions.className = 'markup-popover-actions';
  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.className = 'btn-pill';
  cancelBtn.textContent = '취소';
  const saveBtn = document.createElement('button');
  saveBtn.type = 'button';
  saveBtn.className = 'btn-pill btn-pill-primary';
  saveBtn.textContent = '저장';
  actions.append(cancelBtn, saveBtn);
  popover.append(textarea, actions);
  frame.append(popover);
  textarea.focus();

  cancelBtn.addEventListener('click', () => cancelMarkupDraft());

  saveBtn.addEventListener('click', async () => {
    const text = textarea.value.trim();
    if (!text) {
      textarea.focus();
      return;
    }
    saveBtn.disabled = true;
    try {
      const item = await fetchJsonView(`/api/view/markup?${tokenQuery}`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ page, x: box.x, y: box.y, w: box.w, h: box.h, text }),
      });
      markup.items.push(item);
      cancelMarkupDraft();
      renderAllMarkup();
      updateMarkupIncludeVisibility();
    } catch (error) {
      alert(error.message || '마크업 저장에 실패했습니다.');
      saveBtn.disabled = false;
    }
  });
}

function setupMarkupTool() {
  document.getElementById('markup-toggle-btn').addEventListener('click', () => {
    setMarkupVisible(!markup.visible);
  });

  const addBtn = document.getElementById('markup-add-btn');
  if (permissions.canEdit) {
    addBtn.hidden = false;
    addBtn.addEventListener('click', () => setMarkupDrawMode(!markup.drawMode));
  }

  pagesContainer.addEventListener('mousedown', (event) => {
    if (!markup.drawMode) return;
    const frame = event.target.closest('.page-frame');
    if (!frame) return;
    event.preventDefault();
    const rect = frame.getBoundingClientRect();
    const fx = (event.clientX - rect.left) / rect.width;
    const fy = (event.clientY - rect.top) / rect.height;
    const rectEl = document.createElement('div');
    rectEl.className = 'markup-rect markup-rect-draft';
    markupLayer(frame).append(rectEl);
    markup.dragging = {
      frame,
      page: [...pagesContainer.children].indexOf(frame) + 1,
      fx,
      fy,
      rectEl,
      box: null,
    };
  });

  document.addEventListener('mousemove', (event) => {
    if (!markup.dragging) return;
    const { frame, fx: startFx, fy: startFy, rectEl } = markup.dragging;
    const rect = frame.getBoundingClientRect();
    const curFx = Math.min(Math.max((event.clientX - rect.left) / rect.width, 0), 1);
    const curFy = Math.min(Math.max((event.clientY - rect.top) / rect.height, 0), 1);
    const x = Math.min(startFx, curFx);
    const y = Math.min(startFy, curFy);
    const w = Math.abs(curFx - startFx);
    const h = Math.abs(curFy - startFy);
    rectEl.style.left = `${x * 100}%`;
    rectEl.style.top = `${y * 100}%`;
    rectEl.style.width = `${w * 100}%`;
    rectEl.style.height = `${h * 100}%`;
    markup.dragging.box = { x, y, w, h };
  });

  document.addEventListener('mouseup', () => {
    if (!markup.dragging) return;
    const { frame, page, rectEl, box } = markup.dragging;
    markup.dragging = null;
    const MIN_SIZE = 0.01;
    if (!box || box.w < MIN_SIZE || box.h < MIN_SIZE) {
      rectEl.remove();
      return;
    }
    markup.draftRect = rectEl;
    openMarkupPopover(frame, page, box);
  });
}
```

- [ ] **Step 3: 인쇄/다운로드 핸들러가 `includeMarkup`을 붙이도록 수정**

`public/js/viewer.js`의 `setupToolbar` 함수 안, 인쇄 버튼 클릭 핸들러(296~316번째 줄)를 교체:

```js
  if (permissions.canPrint) {
    printBtn.addEventListener('click', () => {
      const includeMarkup = document.getElementById('markup-include-checkbox')?.checked ? '&includeMarkup=1' : '';
      // 팝업 차단을 피하려면 사용자 클릭 스택 안에서 창을 먼저 열어둬야 한다
      const printWindow = window.open('', '_blank');
      runButtonTask(printBtn, '생성', async (onProgress) => {
        const chunks = await fetchWithProgress(`/api/view/print?${tokenQuery}${includeMarkup}`, onProgress);
        const blob = new Blob(chunks, { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        if (printWindow && !printWindow.closed) {
          printWindow.location.href = url;
        } else {
          const anchor = document.createElement('a');
          anchor.href = url;
          anchor.target = '_blank';
          anchor.rel = 'noopener';
          anchor.click();
        }
        setTimeout(() => URL.revokeObjectURL(url), 60_000);
      }).catch(() => {
        if (printWindow && !printWindow.closed) printWindow.close();
      });
    });
  }
```

다운로드 버튼 클릭 핸들러(319~334번째 줄)를 교체:

```js
  downloadBtn.disabled = !permissions.canDownload;
  downloadBtn.title = permissions.canDownload ? '원본 파일 다운로드' : '저장 권한이 없습니다';
  if (permissions.canDownload) {
    downloadBtn.addEventListener('click', () => {
      const includeMarkup = document.getElementById('markup-include-checkbox')?.checked ? '&includeMarkup=1' : '';
      runButtonTask(downloadBtn, '저장', async (onProgress) => {
        const chunks = await fetchWithProgress(`/api/view/download?${tokenQuery}${includeMarkup}`, onProgress);
        const blob = new Blob(chunks);
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = docName;
        anchor.click();
        setTimeout(() => URL.revokeObjectURL(url), 60_000);
      }).catch(() => {});
    });
  }
```

- [ ] **Step 4: Esc로 마크업 그리기 모드 종료**

`public/js/viewer.js`의 `setupDeterrents` 함수 안 `escape` 처리(358~361번째 줄)를 교체:

```js
    if (key === 'escape') {
      if (measure.mode) setMeasureMode(false);
      else if (markup.drawMode) setMarkupDrawMode(false);
      else closeSearchBar();
    }
```

- [ ] **Step 5: 그리드 뷰에서 마크업 UI 숨기기**

`public/js/viewer.js`의 `initGridView` 함수 시작부(748~756번째 줄), 기존 `measure-btn` hidden 처리 다음 줄에 추가:

```js
  document.getElementById('markup-toggle-btn').hidden = true;
  document.getElementById('markup-add-btn').hidden = true;
  document.getElementById('markup-include-wrap').hidden = true;
```

- [ ] **Step 6: `init()`에서 새 함수 호출**

`public/js/viewer.js`의 `init()` 함수 끝부분(1026~1030번째 줄)을 교체:

```js
  pageIndicator.textContent = `1 / ${pageCount}`;
  applyZoom();
  buildPages();
  trackCurrentPage();
  setupDrawingTools();
  setupMarkupTool();
  loadMarkup();
```

- [ ] **Step 7: 타입체크 겸 정적 오류 확인**

이 프로젝트의 `public/js/*.js`는 esbuild 번들 대상이 아니라 직접 서빙되는 순수 JS라 별도 빌드 스텝이 없다. 브라우저 콘솔에 문법 오류가 없는지 확인하는 것으로 대체한다: 서버(`npm run dev`)가 떠 있는 상태에서 뷰어 페이지를 열고 개발자 도구 콘솔에 빨간 에러가 없는지 확인한다.

- [ ] **Step 8: 커밋**

```bash
git add public/js/viewer.js
git commit -m "Wire up markup drag-to-draw, popover, toggle, and print/download options"
```

---

### Task 8: 종단 간(End-to-End) 검증

**Files:**
- Create (임시, 커밋 안 함): 프로젝트 루트의 `scripts/_tmp_markup_e2e.mjs` — 검증 후 삭제

**Interfaces:**
- Consumes: Task 1~7의 전체 기능
- Produces: 없음 (검증 전용)

- [ ] **Step 1: Playwright(Node) 스크립트로 실제 브라우저 흐름 검증**

`npm run dev`로 서버를 띄운 상태에서(꺼져 있으면 백그라운드로 시작), 프로젝트 루트에 임시 스크립트를 만든다(이미 `devDependencies`에 `playwright`가 있으므로 `node_modules` 해석을 위해 프로젝트 안에 둔다):

```js
// scripts/_tmp_markup_e2e.mjs
import { chromium } from 'playwright';
import { readFileSync, writeFileSync } from 'node:fs';

const PDF_BYTES = Buffer.from(
  '%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 300 300]>>endobj\nxref\n0 4\n0000000000 65535 f \ntrailer<</Size 4/Root 1 0 R>>\nstartxref\n0\n%%EOF',
);
writeFileSync('scripts/_tmp_markup_e2e.pdf', PDF_BYTES);

const form = new FormData();
form.append('file', new Blob([PDF_BYTES], { type: 'application/pdf' }), 'markup-e2e.pdf');
const uploadRes = await fetch('http://localhost:8000/api/documents', { method: 'POST', body: form });
const uploaded = (await uploadRes.json()).data;

const tokenRes = await fetch(`http://localhost:8000/api/documents/${uploaded.id}/tokens`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify({ viewer: '검토자A', canPrint: true, canDownload: true, canCopy: true, canEdit: true }),
});
const issued = (await tokenRes.json()).data;

const browser = await chromium.launch({ headless: true });
const page = await browser.newPage();
const consoleErrors = [];
page.on('console', (msg) => { if (msg.type() === 'error') consoleErrors.push(msg.text()); });
page.on('pageerror', (err) => consoleErrors.push(String(err)));

await page.goto(`http://localhost:8000${issued.viewerUrl}`);
await page.waitForLoadState('networkidle');
await page.waitForSelector('.page-frame');

// 마크업 추가 모드 켜고 사각형 드래그
await page.click('#markup-add-btn');
const frame = await page.locator('.page-frame').first().boundingBox();
await page.mouse.move(frame.x + 40, frame.y + 40);
await page.mouse.down();
await page.mouse.move(frame.x + 140, frame.y + 90);
await page.mouse.up();
await page.waitForSelector('.markup-popover textarea');
await page.fill('.markup-popover textarea', '이 부분을 수정해주세요');
await page.click('.markup-popover .btn-pill-primary');
await page.waitForSelector('.markup-rect:not(.markup-rect-draft)');

// 새로고침해도 남아있는지 확인
await page.reload();
await page.waitForLoadState('networkidle');
await page.waitForSelector('.markup-rect');
const afterReloadCount = await page.locator('.markup-rect').count();

// 토글로 숨기기
await page.click('#markup-toggle-btn');
await page.waitForTimeout(200);
const hiddenCount = await page.locator('.markup-layer .markup-rect').count();
await page.click('#markup-toggle-btn');

// 삭제
await page.hover('.markup-rect');
await page.click('.markup-tooltip-delete');
await page.waitForTimeout(200);
const afterDeleteCount = await page.locator('.markup-rect').count();

console.log(JSON.stringify({ afterReloadCount, hiddenCount, afterDeleteCount, consoleErrors }));

await browser.close();
await fetch(`http://localhost:8000/api/documents/${uploaded.id}`, { method: 'DELETE' });
```

Run: `node scripts/_tmp_markup_e2e.mjs`
Expected: 출력 JSON이 `{"afterReloadCount":1,"hiddenCount":0,"afterDeleteCount":0,"consoleErrors":[]}` — 즉 새로고침 후에도 마크업이 남아있고(서버 저장 확인), 토글로 숨기면 오버레이가 0개가 되고, 삭제 후에는 0개가 된다.

- [ ] **Step 2: 정리**

```bash
rm scripts/_tmp_markup_e2e.mjs scripts/_tmp_markup_e2e.pdf
```

- [ ] **Step 3: 전체 테스트 스위트 + 타입체크 최종 확인**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 타입 에러 없음, 모든 테스트(기존 + 이번에 추가한 `storage-markup`/`markup`/`rendering-markup`) PASS

- [ ] **Step 4: 최종 상태 확인 (커밋할 것 없음 — Task 1~7에서 이미 다 커밋됨)**

```bash
git status --short
```

Expected: 임시 스크립트를 지웠으니 깨끗해야 한다(다른 병행 작업 중인 파일 변경은 무시).
