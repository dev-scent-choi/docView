# 뷰어 첫 화면 지연 개선 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 큰 문서를 "바로보기"로 열 때 첫 화면(1페이지/첫 그리드 행)이 뜨기까지 걸리는 시간을 줄인다.

**Architecture:** PDF 페이지 크기와 xlsx 행 오프셋 인덱스를 변환/추출 시점에 미리 계산해
`store.saveDerived`로 영구 저장하고(`'pagesizes'`/`'gridindex'` 파생 산출물), 조회 시점엔 그
파일만 읽는다. 이 기능 배포 이전 문서는 파생 산출물이 없으므로, 조회 시 없으면 기존처럼
그 자리에서 계산하되 계산 결과를 그 자리에서 저장해 다음부터는 빠른 경로를 타게 한다
(자가치유, 별도 마이그레이션 불필요). 클라이언트(`viewer.js`)는 1페이지 이미지 요청을 나머지
페이지 프레임 구성과 병렬로 시작한다.

**Tech Stack:** TypeScript(Fastify 서버), pdf-lib, Vitest, 순수 JS 클라이언트(`public/js/viewer.js`, 번들 없음)

**Spec:** [docs/superpowers/specs/2026-09-07-viewer-first-paint-design.md](../specs/2026-09-07-viewer-first-paint-design.md)

## Global Constraints

- `DerivedKind`(`src/storage.ts`)을 `'pdf' | 'grid' | 'markup' | 'pagesizes' | 'gridindex'`로 확장한다.
- `pagesizes` 파생 산출물의 JSON 모양: `Array<{ width: number; height: number }>` (pt 단위,
  반올림하지 않은 raw 값 — 반올림은 소비 시점인 `getPageDimensions`가 한다).
- `gridindex` 파생 산출물의 JSON 모양: `{ lineOffsets: number[] }`.
- `src/rendering.ts`는 `src/storage.ts`를 직접 import하지 않는다 — 영구 저장소 읽기/쓰기는
  항상 호출자가 콜백(`PersistedPageSizes`)으로 주입한다(기존 `loadBytes: BytesProvider` 패턴과 동일).
- 자가치유: 파생 산출물이 없으면 기존과 동일하게 그 자리에서 계산하고, 계산 결과를 그 자리에서
  저장한다. 별도 백필/마이그레이션 스크립트는 만들지 않는다.
- `DiskStore.deleteDocument`, `MemoryStore.deleteDocument`, `MemoryStore.enforceBudget`의
  정리(삭제) 목록에 새 파생 종류 두 가지를 반드시 추가한다 — 빠뜨리면 디스크에 파일이 남거나
  메모리 예산 계산이 새 종류를 누락한다.
- 범위 밖(이번 계획에서 다루지 않음): docx 전용 `/api/view/info` 부가 필드
  (`pageSetup`/`defaultFontSizePt`/`defaultFontFamily`/`editHasUnsupportedContent`) 최적화,
  협업 편집기(editor.html 등)의 로딩 방식, PDF 렌더링(래스터화) 자체의 캐시 정책.

---

### Task 1: storage.ts — 파생 산출물 종류 확장

**Files:**
- Modify: `src/storage.ts:11` (`DerivedKind`), `:59-63` (`derivedPath`), `:132-142`
  (`DiskStore.deleteDocument`), `:168-181` (`MemoryStore.enforceBudget`), `:225-231`
  (`MemoryStore.deleteDocument`)
- Test: `tests/storage-derivedkinds.test.ts` (신규)

**Interfaces:**
- Produces: `DerivedKind` 타입에 `'pagesizes' | 'gridindex'` 추가 — Task 2~6이 이 두 종류로
  `store.saveDerived`/`readDerived`/`hasDerived`를 호출한다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/storage-derivedkinds.test.ts` (전체 새 파일, `tests/storage-markup.test.ts`와 동일한
임시 디스크 디렉터리 패턴을 그대로 따른다):

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

describe('storage: pagesizes/gridindex derived kinds', () => {
  let tempDirs: string[] = [];

  afterEach(() => {
    delete process.env.DOCVIEW_DATA_DIR;
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
    for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true });
    tempDirs = [];
  });

  test('pagesizes and gridindex derived kinds do not collide with each other or with grid/pdf/markup', async () => {
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
    await store.saveDerived(id, 'pagesizes', Buffer.from('[{"width":595,"height":842}]', 'utf8'));
    await store.saveDerived(id, 'gridindex', Buffer.from('{"lineOffsets":[0,10]}', 'utf8'));
    expect((await store.readDerived(id, 'grid')).toString('utf8')).toBe('{"grid":true}');
    expect((await store.readDerived(id, 'pagesizes')).toString('utf8')).toBe(
      '[{"width":595,"height":842}]',
    );
    expect((await store.readDerived(id, 'gridindex')).toString('utf8')).toBe(
      '{"lineOffsets":[0,10]}',
    );
  });

  test('hasDerived reports false before save and true after, for both new kinds', async () => {
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
    expect(await store.hasDerived(id, 'pagesizes')).toBe(false);
    expect(await store.hasDerived(id, 'gridindex')).toBe(false);
    await store.saveDerived(id, 'pagesizes', Buffer.from('[]', 'utf8'));
    await store.saveDerived(id, 'gridindex', Buffer.from('{"lineOffsets":[]}', 'utf8'));
    expect(await store.hasDerived(id, 'pagesizes')).toBe(true);
    expect(await store.hasDerived(id, 'gridindex')).toBe(true);
  });

  test('deleteDocument removes pagesizes and gridindex derived files along with original/meta', async () => {
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
    await store.saveDerived(id, 'pagesizes', Buffer.from('[]', 'utf8'));
    await store.saveDerived(id, 'gridindex', Buffer.from('{"lineOffsets":[]}', 'utf8'));
    await store.deleteDocument(id);
    expect(await store.readMeta(id)).toBeNull();
    expect(await store.hasDerived(id, 'pagesizes')).toBe(false);
    expect(await store.hasDerived(id, 'gridindex')).toBe(false);
  });
});
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/storage-derivedkinds.test.ts`
Expected: FAIL — `saveDerived`/`readDerived`가 `'pagesizes'`/`'gridindex'`를 `DerivedKind`
타입으로 받아들이지 못해 TypeScript 컴파일 자체가 실패하거나(타입 오류), 런타임에서
`derivedPath`의 삼항 매핑이 `'markup.json'`으로 잘못 떨어져 세 파일이 서로 덮어써서
assertion이 깨진다.

- [ ] **Step 3: `src/storage.ts` 수정**

Line 11:
```ts
export type DerivedKind = 'pdf' | 'grid' | 'markup' | 'pagesizes' | 'gridindex';
```

`derivedPath` (기존 `:59-63`을 대체):
```ts
  private derivedPath(id: string, kind: DerivedKind): string {
    assertValidId(id);
    const suffix =
      kind === 'pdf' ? 'pdf'
      : kind === 'grid' ? 'grid.json'
      : kind === 'markup' ? 'markup.json'
      : kind === 'pagesizes' ? 'pagesizes.json'
      : 'gridindex.json';
    return path.join(config.convertedDir, `${id}.${suffix}`);
  }
```

`DiskStore.deleteDocument` (기존 `:132-142`을 대체):
```ts
  async deleteDocument(id: string): Promise<void> {
    const meta = await this.readMeta(id);
    if (!meta) return;
    await Promise.allSettled([
      fs.rm(this.originalPath(id, meta.ext), { force: true }),
      fs.rm(this.derivedPath(id, 'pdf'), { force: true }),
      fs.rm(this.derivedPath(id, 'grid'), { force: true }),
      fs.rm(this.derivedPath(id, 'markup'), { force: true }),
      fs.rm(this.derivedPath(id, 'pagesizes'), { force: true }),
      fs.rm(this.derivedPath(id, 'gridindex'), { force: true }),
      fs.rm(this.metaPath(id), { force: true }),
    ]);
  }
```

`MemoryStore.enforceBudget` (기존 `:168-181`을 대체 — 삭제 목록 두 줄만 추가):
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
      this.derived.delete(`${oldest.id}#pagesizes`);
      this.derived.delete(`${oldest.id}#gridindex`);
      this.metas.delete(oldest.id);
      console.warn(`메모리 예산 초과로 오래된 문서를 제거했습니다: ${oldest.name} (${oldest.id})`);
    }
  }
```

`MemoryStore.deleteDocument` (기존 `:225-231`을 대체):
```ts
  async deleteDocument(id: string): Promise<void> {
    this.originals.delete(id);
    this.derived.delete(`${id}#pdf`);
    this.derived.delete(`${id}#grid`);
    this.derived.delete(`${id}#markup`);
    this.derived.delete(`${id}#pagesizes`);
    this.derived.delete(`${id}#gridindex`);
    this.metas.delete(id);
  }
```

- [ ] **Step 4: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/storage-derivedkinds.test.ts`
Expected: PASS (3 tests)

- [ ] **Step 5: 기존 전체 테스트 스위트가 깨지지 않았는지 확인**

Run: `npx vitest run && npx tsc --noEmit`
Expected: 모두 PASS (특히 `tests/storage-markup.test.ts`)

- [ ] **Step 6: 커밋**

```bash
git add src/storage.ts tests/storage-derivedkinds.test.ts
git commit -m "feat: 파생 산출물에 pagesizes/gridindex 종류 추가"
```

---

### Task 2: spreadsheet.ts — 줄 오프셋 계산 + 그리드 저장 헬퍼

**Files:**
- Modify: `src/spreadsheet.ts:1` (import 추가), `:90` 뒤에 새 함수 두 개 추가
- Test: `tests/spreadsheet-derived.test.ts` (신규)

**Interfaces:**
- Consumes: `Task 1`이 확장한 `DerivedKind`(`'gridindex'`), `DocumentStore`(`src/storage.ts`)의
  `saveDerived(id: string, kind: DerivedKind, data: Buffer): Promise<void>` 시그니처.
- Produces:
  - `computeLineOffsets(buffer: Buffer): readonly number[]` — Task 6(`viewer.ts`)이 재사용한다.
  - `saveGridDerived(store: Pick<DocumentStore, 'saveDerived'>, docId: string, grid: GridDocument): Promise<void>`
    — Task 4(`documents.ts`)와 Task 5(`room.ts`)가 재사용한다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/spreadsheet-derived.test.ts` (신규 파일):

```ts
import { describe, expect, test, vi } from 'vitest';
import { computeLineOffsets, saveGridDerived, type GridDocument } from '../src/spreadsheet.js';

describe('spreadsheet: computeLineOffsets', () => {
  test('returns byte offsets for each line, starting with 0', () => {
    const buffer = Buffer.from('line one\nline two\nline three', 'utf8');
    const offsets = computeLineOffsets(buffer);
    expect(offsets).toEqual([0, 9, 18]);
    expect(buffer.toString('utf8', offsets[1], offsets[2] - 1)).toBe('line two');
  });

  test('a buffer with no newline yields just the initial 0 offset', () => {
    const buffer = Buffer.from('single line, no newline', 'utf8');
    expect(computeLineOffsets(buffer)).toEqual([0]);
  });
});

function makeTestGrid(): GridDocument {
  return {
    styles: [],
    sheets: [
      {
        name: 'Sheet1',
        rowCount: 2,
        colCount: 1,
        colWidths: [],
        merges: [],
        freezeRows: 0,
        freezeCols: 0,
        rows: [[{ v: 'a' }], [{ v: 'b' }]],
      },
    ],
  };
}

describe('spreadsheet: saveGridDerived', () => {
  test('saves both grid and gridindex derived kinds, with offsets matching the serialized grid bytes', async () => {
    const saveDerived = vi.fn(async () => undefined);
    await saveGridDerived({ saveDerived }, 'doc-1', makeTestGrid());

    expect(saveDerived).toHaveBeenCalledTimes(2);
    const [gridCall, indexCall] = saveDerived.mock.calls;
    expect(gridCall[0]).toBe('doc-1');
    expect(gridCall[1]).toBe('grid');
    const serializedBytes = gridCall[2] as Buffer;

    expect(indexCall[0]).toBe('doc-1');
    expect(indexCall[1]).toBe('gridindex');
    const { lineOffsets } = JSON.parse((indexCall[2] as Buffer).toString('utf8'));
    expect(lineOffsets).toEqual(computeLineOffsets(serializedBytes));
  });
});
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/spreadsheet-derived.test.ts`
Expected: FAIL — `computeLineOffsets`/`saveGridDerived`가 아직 없어 `'../src/spreadsheet.js'`에서
import 오류.

- [ ] **Step 3: `src/spreadsheet.ts` 수정**

파일 최상단(`import ExcelJS from 'exceljs';` 다음 줄)에 타입 전용 import 추가:

```ts
import type { DocumentStore } from './storage.js';
```

`serializeGrid` 함수(현재 `:78-90`) 바로 뒤에 다음 두 함수를 추가:

```ts
/**
 * JSONL 버퍼에서 각 줄이 시작하는 바이트 오프셋 배열을 만든다(0번째는 항상 0 = 헤더 시작줄).
 * src/routes/viewer.ts의 그리드 조회(sliceRows/searchGrid)가 필요한 범위만 파싱할 수 있게
 * 해주는 인덱스 — 수십만 행이면 이 스캔 자체가 O(파일 크기)라, 그리드를 만드는 시점에 미리
 * 계산해 saveGridDerived로 함께 저장해 두면 조회 시점에 매번 다시 스캔하지 않아도 된다.
 */
export function computeLineOffsets(buffer: Buffer): readonly number[] {
  const lineOffsets: number[] = [0];
  for (let i = buffer.indexOf(10); i !== -1; i = buffer.indexOf(10, i + 1)) {
    lineOffsets.push(i + 1);
  }
  return lineOffsets;
}

/**
 * 그리드를 JSONL로 직렬화해 저장하고(derived kind 'grid'), 같은 바이트로 만든 줄 오프셋
 * 인덱스도 함께 저장한다(derived kind 'gridindex'). 업로드 시 최초 추출
 * (routes/documents.ts의 runGridExtraction)과 동시편집 저장 후 재생성(collab/room.ts)
 * 양쪽에서 재사용해, 그리드 내용이 바뀔 때마다 인덱스가 항상 최신 상태로 유지되게 한다
 * (그리드가 바뀌지 않는 한 인덱스도 따로 무효화할 필요가 없다).
 */
export async function saveGridDerived(
  store: Pick<DocumentStore, 'saveDerived'>,
  docId: string,
  grid: GridDocument,
): Promise<void> {
  const serialized = serializeGrid(grid);
  await store.saveDerived(docId, 'grid', serialized);
  await store.saveDerived(
    docId,
    'gridindex',
    Buffer.from(JSON.stringify({ lineOffsets: computeLineOffsets(serialized) })),
  );
}
```

- [ ] **Step 4: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/spreadsheet-derived.test.ts`
Expected: PASS (3 tests)

- [ ] **Step 5: 타입 체크**

Run: `npx tsc --noEmit`
Expected: PASS (특히 `DocumentStore` 순환 import 없음 확인 — `storage.ts`는 `spreadsheet.ts`를
import하지 않으므로 문제 없어야 한다)

- [ ] **Step 6: 커밋**

```bash
git add src/spreadsheet.ts tests/spreadsheet-derived.test.ts
git commit -m "feat: 그리드 JSONL 줄 오프셋 사전 계산 헬퍼 추가"
```

---

### Task 3: rendering.ts — PDF 페이지 크기 사전 계산 지원

**Files:**
- Modify: `src/rendering.ts:51-54`(`PageSize`), `:93-139`(`getPageSizes`/`getPageDimensions`/
  `getPageCountFromBytes`), `:165-178`(`renderPage`)
- Test: `tests/rendering-pagesizes.test.ts` (신규)

**Interfaces:**
- Consumes: 없음(이 파일 자체가 최하위 계층 — `storage.ts`를 import하지 않는다).
- Produces:
  - `export interface PageSize { readonly width: number; readonly height: number; }`
  - `export interface PersistedPageSizes { readonly load: () => Promise<readonly PageSize[] | null>; readonly save: (sizes: readonly PageSize[]) => Promise<void>; }`
  - `export interface PdfPageInfo { readonly pageCount: number; readonly sizes: readonly PageSize[]; }`
  - `export async function getPdfPageInfo(pdfBytes: Buffer): Promise<PdfPageInfo>` — Task 4가
    기존 `getPageCountFromBytes` 대신 이 함수를 쓴다.
  - `export async function getPageDimensions(docId: string, loadBytes: BytesProvider, persisted: PersistedPageSizes): Promise<ReadonlyArray<{ w: number; h: number }>>`
    — 시그니처에 `persisted` 매개변수 추가(필수, 선택 아님). Task 6이 호출한다.
  - `export async function renderPage(docId: string, pageNumber: number, requestedScale: number, loadBytes: BytesProvider, persisted: PersistedPageSizes): Promise<Buffer>`
    — 시그니처에 `persisted` 매개변수 추가. Task 6이 호출한다.
  - `getPageCountFromBytes`는 삭제된다(오직 `src/routes/documents.ts:113` 한 곳에서만 쓰이며,
    Task 4가 `getPdfPageInfo`로 교체한다).

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/rendering-pagesizes.test.ts` (신규 파일 — `tests/helpers.ts`의 `makeSamplePdf` 재사용,
`tests/rendering-markup.test.ts`와 같은 순수 함수 테스트 스타일):

```ts
import { describe, expect, test, vi } from 'vitest';
import { getPageDimensions, getPdfPageInfo, renderPage, type PageSize } from '../src/rendering.js';
import { makeSamplePdf } from './helpers.js';

function nullPersistence() {
  return {
    load: vi.fn(async () => null as readonly PageSize[] | null),
    save: vi.fn(async () => undefined),
  };
}

describe('rendering: getPdfPageInfo', () => {
  test('returns page count and per-page sizes from a single parse', async () => {
    const pdf = await makeSamplePdf(3);
    const info = await getPdfPageInfo(pdf);
    expect(info.pageCount).toBe(3);
    expect(info.sizes).toEqual([
      { width: 595, height: 842 },
      { width: 595, height: 842 },
      { width: 595, height: 842 },
    ]);
  });
});

describe('rendering: getPageDimensions with persisted page sizes', () => {
  test('uses persisted sizes without ever calling loadBytes when they exist', async () => {
    const persisted = {
      load: vi.fn(async () => [{ width: 100, height: 200 }] as readonly PageSize[]),
      save: vi.fn(async () => undefined),
    };
    const loadBytes = vi.fn(async () => {
      throw new Error('loadBytes should not be called when persisted sizes exist');
    });
    const dims = await getPageDimensions(`doc-${Math.random()}`, loadBytes, persisted);
    expect(dims).toEqual([{ w: 100, h: 200 }]);
    expect(loadBytes).not.toHaveBeenCalled();
    expect(persisted.save).not.toHaveBeenCalled();
  });

  test('falls back to parsing bytes and self-heals by calling persisted.save when nothing is persisted yet', async () => {
    const pdf = await makeSamplePdf(2);
    const persisted = nullPersistence();
    const loadBytes = vi.fn(async () => pdf);
    const dims = await getPageDimensions(`doc-${Math.random()}`, loadBytes, persisted);
    expect(dims).toEqual([
      { w: 595, h: 842 },
      { w: 595, h: 842 },
    ]);
    expect(loadBytes).toHaveBeenCalledTimes(1);
    expect(persisted.save).toHaveBeenCalledTimes(1);
    expect(persisted.save).toHaveBeenCalledWith([
      { width: 595, height: 842 },
      { width: 595, height: 842 },
    ]);
  });
});

describe('rendering: renderPage with persisted page sizes', () => {
  test('renders using persisted sizes, calling loadBytes only once (for rasterization, not the size lookup)', async () => {
    const pdf = await makeSamplePdf(1);
    const persisted = {
      load: vi.fn(async () => [{ width: 595, height: 842 }] as readonly PageSize[]),
      save: vi.fn(async () => undefined),
    };
    // loadBytes는 renderPage 내부의 실제 래스터화(openDocument)에는 여전히 필요하므로 진짜
    // PDF를 준다 — 대신 호출 횟수를 세어, persisted가 있으면 "크기 조회" 몫으로는 단 한 번도
    // 더 불리지 않는다는 것(래스터화용 1회만 호출됨)을 확인한다.
    const loadBytes = vi.fn(async () => pdf);
    const png = await renderPage(`doc-${Math.random()}`, 1, 1, loadBytes, persisted);
    expect(png.length).toBeGreaterThan(0);
    expect(persisted.load).toHaveBeenCalledTimes(1);
    expect(loadBytes).toHaveBeenCalledTimes(1);
  });
});
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/rendering-pagesizes.test.ts`
Expected: FAIL — `getPdfPageInfo`가 없고, `getPageDimensions`/`renderPage`가 세 번째/다섯 번째
인자를 받지 않아 타입 오류.

- [ ] **Step 3: `src/rendering.ts` 수정**

기존 `PageSize` 인터페이스(`:51-54`)를 export로 바꾸고 바로 아래에 새 타입 두 개를 추가:

```ts
export interface PageSize {
  readonly width: number;
  readonly height: number;
}

/** getPageSizes가 매번 PDF를 새로 파싱하지 않도록, 영구 저장된 페이지 크기를 읽고/쓰는
 * 콜백 쌍 — loadBytes와 같은 기존 패턴대로 이 파일은 storage.ts를 직접 import하지 않고
 * 항상 호출자(routes/*.ts)가 만들어 주입한다. */
export interface PersistedPageSizes {
  readonly load: () => Promise<readonly PageSize[] | null>;
  readonly save: (sizes: readonly PageSize[]) => Promise<void>;
}
```

기존 `getPageCountFromBytes`(`:133-139`)를 통째로 지우고 그 자리에:

```ts
export interface PdfPageInfo {
  readonly pageCount: number;
  readonly sizes: readonly PageSize[];
}

/**
 * PDF 바이트를 한 번만 파싱해 페이지 수와 페이지별 크기를 함께 얻는다. 변환 직후(업로드,
 * 동시편집 저장 후 재변환) routes/documents.ts의 convertToPdf가 이 함수로 페이지 수를 얻는
 * 김에 페이지 크기도 함께 영구 저장해 두면, 조회 시점(getPageSizes)엔 다시 파싱할 필요가 없다.
 */
export async function getPdfPageInfo(pdfBytes: Buffer): Promise<PdfPageInfo> {
  const doc = await PDFDocument.load(toOwnedBytes(pdfBytes), {
    ignoreEncryption: true,
    updateMetadata: false,
  });
  const sizes = doc.getPages().map((page) => {
    const { width, height } = page.getSize();
    return { width, height };
  });
  return { pageCount: doc.getPageCount(), sizes };
}
```

기존 `getPageSizes`(`:110-122`)를 대체:

```ts
function getPageSizes(
  docId: string,
  loadBytes: BytesProvider,
  persisted: PersistedPageSizes,
): Promise<readonly PageSize[]> {
  return pageSizeCache.getOrLoad(docId, async () => {
    const existing = await persisted.load();
    if (existing) return existing;
    const bytes = await loadBytes();
    const { sizes } = await getPdfPageInfo(bytes);
    await persisted.save(sizes);
    return sizes;
  });
}
```

기존 `getPageDimensions`(`:125-131`)를 대체:

```ts
/** 페이지 크기 목록(pt) — 뷰어가 로딩 전 프레임 크기를 고정해 정확한 스크롤 점프를 보장한다 */
export async function getPageDimensions(
  docId: string,
  loadBytes: BytesProvider,
  persisted: PersistedPageSizes,
): Promise<ReadonlyArray<{ w: number; h: number }>> {
  const sizes = await getPageSizes(docId, loadBytes, persisted);
  return sizes.map((size) => ({ w: Math.round(size.width), h: Math.round(size.height) }));
}
```

기존 `renderPage`(`:165-178`)를 대체:

```ts
export async function renderPage(
  docId: string,
  pageNumber: number,
  requestedScale: number,
  loadBytes: BytesProvider,
  persisted: PersistedPageSizes,
): Promise<Buffer> {
  const sizes = await getPageSizes(docId, loadBytes, persisted);
  if (pageNumber < 1 || pageNumber > sizes.length) {
    throw new PageOutOfRangeError(pageNumber, sizes.length);
  }
  const scale = clampScale(requestedScale ?? config.renderScale, sizes[pageNumber - 1]);
  const document = await openDocument(docId, scale, loadBytes);
  return document.getPage(pageNumber);
}
```

- [ ] **Step 4: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/rendering-pagesizes.test.ts`
Expected: PASS (4 tests)

- [ ] **Step 5: 타입 체크 — 아직 documents.ts/viewer.ts를 안 고쳤으므로 실패가 정상**

Run: `npx tsc --noEmit`
Expected: FAIL — `src/routes/documents.ts:113`이 삭제된 `getPageCountFromBytes`를 참조하고,
`src/routes/viewer.ts`의 `getPageDimensions`/`renderPage` 호출이 인자 개수가 안 맞는다는
오류가 나야 정상이다(Task 4/6에서 고친다). 이 시점에 이 오류들 **외의** 다른 새 오류가
없는지만 확인한다.

- [ ] **Step 6: 커밋**

```bash
git add src/rendering.ts tests/rendering-pagesizes.test.ts
git commit -m "feat: PDF 페이지 크기를 영구 저장소에서 우선 조회하도록 rendering.ts 확장"
```

(주의: 이 커밋 시점엔 `tsc --noEmit`이 아직 실패한다 — Task 4/6이 호출부를 마저 고쳐야
전체가 다시 초록불이 된다. Step 5에서 예상한 오류와 실제 오류가 일치하는지 반드시 확인하고
커밋한다.)

---

### Task 4: documents.ts — 변환/추출 시점에 파생 산출물 저장 연결

**Files:**
- Modify: `src/routes/documents.ts:10`(import), `:12`(import), `:66-91`(`runGridExtraction`),
  `:101-118`(`convertToPdf`)
- Test: `tests/documents-derived.test.ts` (신규)

**Interfaces:**
- Consumes: Task 2의 `saveGridDerived(store, docId, grid)`, Task 3의
  `getPdfPageInfo(pdfBytes): Promise<PdfPageInfo>`.
- Produces: 없음(외부에서 재사용하는 새 함수 없음 — 기존 `convertToPdf`의 반환 타입은
  `{ pdfBytes, pageCount }`로 그대로 유지해 `src/collab/room.ts`의 기존 호출부를 깨지 않는다).

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/documents-derived.test.ts` (신규 파일 — `MemoryStore`를 직접 써서 파일 I/O 없이
`convertToPdf`/`runGridExtraction`이 파생 산출물을 저장하는지만 검증한다. LibreOffice 등
실제 변환기는 필요 없다 — `convertToPdf`는 `Converter` 인터페이스만 요구하므로 가짜
converter를 준다):

```ts
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, test, vi } from 'vitest';
import { makeSamplePdf } from './helpers.js';
import type { Converter } from '../src/converters/types.js';

async function loadDocumentsModuleWithTempDir() {
  const dir = mkdtempSync(path.join(os.tmpdir(), 'docview-documents-test-'));
  vi.resetModules();
  process.env.DOCVIEW_DATA_DIR = dir;
  process.env.DOCVIEW_STORAGE = 'disk';
  const storageModule = await import('../src/storage.js');
  await storageModule.ensureDataDirs();
  const documentsModule = await import('../src/routes/documents.js');
  return { dir, storageModule, documentsModule };
}

describe('documents.ts: convertToPdf persists pagesizes', () => {
  let tempDirs: string[] = [];

  afterEach(() => {
    delete process.env.DOCVIEW_DATA_DIR;
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
    for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true });
    tempDirs = [];
  });

  test('after convertToPdf, pagesizes derived data matches the converted PDF page sizes', async () => {
    const { dir, storageModule, documentsModule } = await loadDocumentsModuleWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();
    await store.saveOriginal(id, 'docx', Buffer.from('fake docx bytes', 'utf8'));
    await store.writeMeta({
      id,
      name: 'a.docx',
      ext: 'docx',
      status: 'processing',
      pageCount: 0,
      createdAt: new Date().toISOString(),
    });
    const meta = (await store.readMeta(id))!;

    const samplePdf = await makeSamplePdf(2);
    const fakeConverter: Converter = {
      name: 'fake',
      extensions: ['docx'],
      convert: async (_source, targetPdfPath) => {
        writeFileSync(targetPdfPath, samplePdf);
      },
    };

    const { pageCount } = await documentsModule.convertToPdf(meta, fakeConverter);
    expect(pageCount).toBe(2);
    expect(await store.hasDerived(id, 'pagesizes')).toBe(true);
    const sizes = JSON.parse((await store.readDerived(id, 'pagesizes')).toString('utf8'));
    expect(sizes).toEqual([
      { width: 595, height: 842 },
      { width: 595, height: 842 },
    ]);
  });
});
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/documents-derived.test.ts`
Expected: FAIL — `hasDerived(id, 'pagesizes')`가 `false` (아직 저장 안 함).

- [ ] **Step 3: `src/routes/documents.ts` 수정**

Line 10, import 교체:
```ts
import { getPdfPageInfo, invalidateRenderCache } from '../rendering.js';
```

Line 12, import 교체(`serializeGrid` 대신 `saveGridDerived`):
```ts
import { parseCsvToGrid, parseXlsxToGrid, saveGridDerived, type GridDocument } from '../spreadsheet.js';
```

`runGridExtraction`(기존 `:66-91`) 안의 이 한 줄:
```ts
    await store.saveDerived(meta.id, 'grid', serializeGrid(grid));
```
을 다음으로 교체:
```ts
    await saveGridDerived(store, meta.id, grid);
```

`convertToPdf`(기존 `:101-118`)를 대체:
```ts
export async function convertToPdf(
  meta: DocumentMeta,
  converter: Converter,
): Promise<{ pdfBytes: Buffer; pageCount: number }> {
  const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-convert-'));
  try {
    const sourcePath = path.join(workDir, `source.${meta.ext}`);
    await fs.writeFile(sourcePath, await store.readOriginal(meta.id, meta.ext));
    const targetPdfPath = path.join(workDir, 'converted.pdf');
    await converter.convert(sourcePath, targetPdfPath, meta.name);
    const pdfBytes = await fs.readFile(targetPdfPath);
    await store.saveDerived(meta.id, 'pdf', pdfBytes);
    const { pageCount, sizes } = await getPdfPageInfo(pdfBytes);
    await store.saveDerived(meta.id, 'pagesizes', Buffer.from(JSON.stringify(sizes)));
    return { pdfBytes, pageCount };
  } finally {
    await fs.rm(workDir, { recursive: true, force: true });
  }
}
```

- [ ] **Step 4: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/documents-derived.test.ts`
Expected: PASS (1 test)

- [ ] **Step 5: 타입 체크 — documents.ts 관련 오류는 사라져야 한다**

Run: `npx tsc --noEmit`
Expected: `src/routes/documents.ts` 관련 오류는 모두 사라지고, `src/routes/viewer.ts` 관련
오류만 남아야 한다(Task 6에서 해결).

- [ ] **Step 6: 커밋**

```bash
git add src/routes/documents.ts tests/documents-derived.test.ts
git commit -m "feat: 업로드 변환/그리드 추출 시 페이지 크기·행 인덱스를 함께 저장"
```

---

### Task 5: room.ts — 동시편집 저장 후에도 gridindex를 최신으로 유지

**Files:**
- Modify: `src/collab/room.ts:33`(import), `:153`(호출부)
- Test: `tests/roomGridIndex.test.ts` (신규)

**Interfaces:**
- Consumes: Task 2의 `saveGridDerived(store, docId, grid)`.

**왜 필요한가:** xlsx 동시편집 저장(`flushToStore`의 `kind === 'xlsx'` 분기)은
`runGridExtraction`을 거치지 않고 직접 `store.saveDerived(docId, 'grid', serializeGrid(grid))`를
호출해 그리드를 재생성한다. Task 4에서 `runGridExtraction`만 `saveGridDerived`로 바꾸면, 최초
업로드 때는 `gridindex`가 저장되지만 **동시편집으로 셀을 바꿔 재저장한 뒤에는 `gridindex`가
갱신되지 않고 예전 오프셋 그대로 남는다** — 그리드 JSONL 바이트 내용이 바뀌었으므로 오프셋도
달라졌을 수 있어, 다음 조회가 낡은 오프셋으로 새 파일을 잘라 읽는 데이터 손상 버그가 된다.
이 태스크는 `room.ts`도 같은 헬퍼로 바꿔 이 창을 없앤다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/roomGridIndex.test.ts` (신규 파일 — 이 세션에서 이미 쓴
`tests/roomAutosaveLoop.test.ts`의 FakeSocket/fake-timer 패턴을 그대로 재사용한다. xlsx
문서를 열고 셀 하나를 편집해 자동저장을 한 번 돌린 뒤, `gridindex`가 방금 저장된 `grid`
바이트에 대해 실제로 올바른 값으로 갱신됐는지 확인한다):

```ts
import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeUpdate } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import ExcelJS from 'exceljs';

const MESSAGE_SYNC = 0;

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  readonly mirror = new Y.Doc();

  send(data: Uint8Array, cb?: (err?: Error) => void): void {
    const decoder = decoding.createDecoder(data);
    const messageType = decoding.readVarUint(decoder);
    if (messageType === MESSAGE_SYNC) {
      const syncType = decoding.readVarUint(decoder);
      if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) {
        readUpdate(decoder, this.mirror, this);
      }
      // messageYjsSyncStep1은 이 테스트에서 응답할 필요가 없다(셀 편집만 서버로 보내면 된다).
    }
    cb?.();
  }
  close(): void {}
  ping(): void {}
}

async function flushMicrotasks(times = 50): Promise<void> {
  for (let i = 0; i < times; i += 1) await Promise.resolve();
}

async function makeSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 'Name';
  sheet.getCell('A2').value = 'Alice';
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
}

describe('room.ts: xlsx collab save keeps gridindex in sync with the grid it just wrote', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-gridindex-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
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

  test('after an xlsx cell edit autosaves, gridindex offsets match the freshly saved grid bytes', async () => {
    vi.useFakeTimers();

    const { store } = await import('../src/storage.js');
    const { joinEditSession } = await import('../src/collab/room.js');
    const { GRID_CELLS_MAP, gridCellKey } = await import('../src/collab/gridShared.js');
    const { computeLineOffsets } = await import('../src/spreadsheet.js');

    const docId = randomUUID();
    const original = await makeSampleXlsx();
    await store.saveOriginal(docId, 'xlsx', original);
    await store.writeMeta({
      id: docId,
      name: 'x.xlsx',
      ext: 'xlsx',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
      progress: 100,
      viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);

    // 실제 원격 편집처럼, 별도 Y.Doc에서 셀 하나를 바꾼 업데이트를 서버로 보낸다.
    const scratch = new Y.Doc();
    scratch.getMap<string>(GRID_CELLS_MAP).set(gridCellKey(0, 1, 0), 'Alice Kim');
    const update = Y.encodeStateAsUpdate(scratch);
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeUpdate(encoder, update);
    socket.emit('message', encoding.toUint8Array(encoder));

    await vi.advanceTimersByTimeAsync(8_100);
    await flushMicrotasks();

    expect(await store.hasDerived(docId, 'gridindex')).toBe(true);
    const savedGrid = await store.readDerived(docId, 'grid');
    const { lineOffsets } = JSON.parse((await store.readDerived(docId, 'gridindex')).toString('utf8'));
    expect(lineOffsets).toEqual(computeLineOffsets(savedGrid));
  });
});
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/roomGridIndex.test.ts`
Expected: FAIL — `store.hasDerived(docId, 'gridindex')`가 `false` (room.ts가 아직 `'grid'`만
저장하고 `'gridindex'`는 저장하지 않는다).

- [ ] **Step 3: `src/collab/room.ts` 수정**

Line 33, import 교체:
```ts
import { parseXlsxToGrid, saveGridDerived } from '../spreadsheet.js';
```

Line 153의 이 한 줄:
```ts
      await store.saveDerived(docId, 'grid', serializeGrid(grid));
```
을 다음으로 교체:
```ts
      await saveGridDerived(store, docId, grid);
```

- [ ] **Step 4: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/roomGridIndex.test.ts`
Expected: PASS (1 test)

- [ ] **Step 5: 기존 xlsx 동시편집 테스트가 깨지지 않았는지 확인**

Run: `npx vitest run tests/xlsx-collab.test.ts tests/roomAutosaveLoop.test.ts && npx tsc --noEmit`
Expected: 모두 PASS

- [ ] **Step 6: 커밋**

```bash
git add src/collab/room.ts tests/roomGridIndex.test.ts
git commit -m "fix: xlsx 동시편집 저장 후에도 gridindex가 최신 상태로 갱신되게 함"
```

---

### Task 6: viewer.ts — 조회 경로가 영구 저장소를 우선 사용하도록 배선

**Files:**
- Modify: `src/routes/viewer.ts:9-15`(import), `:28`(import), `:71-122`(`LoadedGrid`/`indexGrid`/
  `loadGrid`), `:380`(`getPageDimensions` 호출), `:477-479`(`renderPage` 호출)
- Test: `tests/viewer-derived-cache.test.ts` (신규)

**Interfaces:**
- Consumes: Task 3의 `PageSize`/`PersistedPageSizes` 타입과 새 `getPageDimensions`/`renderPage`
  시그니처, Task 2의 `computeLineOffsets`.
- Produces:
  - `export function pageSizePersistence(docId: string): PersistedPageSizes` — 테스트가 직접
    호출해 검증한다.
  - `export function loadGrid(docId: string): Promise<LoadedGrid>` — 기존에는 export가
    아니었다; 테스트를 위해 export로 바꾼다(외부 라우트 동작은 그대로).

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/viewer-derived-cache.test.ts` (신규 파일 — `tests/storage-markup.test.ts`와 같은
임시 디스크 디렉터리 패턴):

```ts
import { mkdtempSync, rmSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, test, vi } from 'vitest';

async function loadViewerModuleWithTempDir() {
  const dir = mkdtempSync(path.join(os.tmpdir(), 'docview-viewer-test-'));
  vi.resetModules();
  process.env.DOCVIEW_DATA_DIR = dir;
  process.env.DOCVIEW_STORAGE = 'disk';
  const storageModule = await import('../src/storage.js');
  await storageModule.ensureDataDirs();
  const viewerModule = await import('../src/routes/viewer.js');
  return { dir, storageModule, viewerModule };
}

function sampleGridHeader() {
  return {
    styles: [],
    sheets: [
      { name: 'S1', rowCount: 2, colCount: 1, colWidths: [], merges: [], freezeRows: 0, freezeCols: 0 },
    ],
  };
}

describe('viewer.ts: loadGrid self-heals and reuses a persisted gridindex', () => {
  let tempDirs: string[] = [];

  afterEach(() => {
    delete process.env.DOCVIEW_DATA_DIR;
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
    for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true });
    tempDirs = [];
  });

  test('computes and persists gridindex when missing, matching a manual scan', async () => {
    const { dir, storageModule, viewerModule } = await loadViewerModuleWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();
    const jsonl = [
      JSON.stringify(sampleGridHeader()),
      JSON.stringify([{ v: 'a' }]),
      JSON.stringify([{ v: 'b' }]),
    ].join('\n');
    await store.saveDerived(id, 'grid', Buffer.from(jsonl, 'utf8'));

    expect(await store.hasDerived(id, 'gridindex')).toBe(false);
    const grid = await viewerModule.loadGrid(id);
    expect(grid.header.sheets[0].name).toBe('S1');
    expect(await store.hasDerived(id, 'gridindex')).toBe(true);
    const persisted = JSON.parse((await store.readDerived(id, 'gridindex')).toString('utf8'));
    expect(persisted.lineOffsets).toEqual(grid.lineOffsets);
  });

  test('uses a persisted gridindex as-is instead of rescanning the buffer', async () => {
    const { dir, storageModule, viewerModule } = await loadViewerModuleWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();
    const jsonl = [
      JSON.stringify(sampleGridHeader()),
      JSON.stringify([{ v: 'a' }]),
      JSON.stringify([{ v: 'b' }]),
    ].join('\n');
    await store.saveDerived(id, 'grid', Buffer.from(jsonl, 'utf8'));
    // 실제 스캔이라면 절대 나올 수 없는 표시용 오프셋 — loadGrid가 이 값을 그대로 쓴다면
    // (다시 스캔하지 않는다면) 아래 assertion에서 바로 드러난다.
    await store.saveDerived(id, 'gridindex', Buffer.from(JSON.stringify({ lineOffsets: [0, 999] }), 'utf8'));

    const grid = await viewerModule.loadGrid(id);
    expect(grid.lineOffsets).toEqual([0, 999]);
  });
});

describe('viewer.ts: pageSizePersistence', () => {
  let tempDirs: string[] = [];

  afterEach(() => {
    delete process.env.DOCVIEW_DATA_DIR;
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
    for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true });
    tempDirs = [];
  });

  test('load() returns null when no pagesizes derived data exists', async () => {
    const { dir, storageModule, viewerModule } = await loadViewerModuleWithTempDir();
    tempDirs.push(dir);
    const { newDocumentId } = storageModule;
    const id = newDocumentId();
    expect(await viewerModule.pageSizePersistence(id).load()).toBeNull();
  });

  test('save() then load() round-trips the page sizes', async () => {
    const { dir, storageModule, viewerModule } = await loadViewerModuleWithTempDir();
    tempDirs.push(dir);
    const { newDocumentId } = storageModule;
    const id = newDocumentId();
    const persistence = viewerModule.pageSizePersistence(id);
    await persistence.save([{ width: 595, height: 842 }]);
    expect(await persistence.load()).toEqual([{ width: 595, height: 842 }]);
  });
});
```

- [ ] **Step 2: 테스트 실행해 실패 확인**

Run: `npx vitest run tests/viewer-derived-cache.test.ts`
Expected: FAIL — `viewerModule.loadGrid`/`viewerModule.pageSizePersistence`가 `undefined`
(아직 export 안 됨).

- [ ] **Step 3: `src/routes/viewer.ts` 수정**

Import 블록(`:9-15`)을 대체:
```ts
import {
  PageOutOfRangeError,
  buildPrintPdf,
  burnMarkupIntoPdf,
  getPageDimensions,
  renderPage,
  type PageSize,
  type PersistedPageSizes,
} from '../rendering.js';
```

Line 28을 대체:
```ts
import { computeLineOffsets, type GridCell, type GridHeader } from '../spreadsheet.js';
```

`indexGrid`/`loadGrid`(기존 `:104-122`)를 대체 — `buildLoadedGridFromOffsets`를 새로 뽑아내고
`loadGrid`가 영구 저장소를 확인하게 한다:

```ts
function buildLoadedGridFromOffsets(buffer: Buffer, lineOffsets: readonly number[]): LoadedGrid {
  const headerEnd = lineOffsets.length > 1 ? lineOffsets[1] - 1 : buffer.length;
  const header = JSON.parse(buffer.toString('utf8', 0, headerEnd)) as GridHeader;
  const sheetStartLine: number[] = [];
  let line = 1;
  for (const sheet of header.sheets) {
    sheetStartLine.push(line);
    line += sheet.rowCount;
  }
  return { header, buffer, lineOffsets, sheetStartLine };
}

function indexGrid(buffer: Buffer): LoadedGrid {
  return buildLoadedGridFromOffsets(buffer, computeLineOffsets(buffer));
}

async function loadGridUncached(docId: string): Promise<LoadedGrid> {
  const buffer = await store.readDerived(docId, 'grid');
  if (await store.hasDerived(docId, 'gridindex')) {
    const raw = await store.readDerived(docId, 'gridindex');
    const { lineOffsets } = JSON.parse(raw.toString('utf8')) as { lineOffsets: number[] };
    return buildLoadedGridFromOffsets(buffer, lineOffsets);
  }
  const grid = indexGrid(buffer);
  await store.saveDerived(
    docId,
    'gridindex',
    Buffer.from(JSON.stringify({ lineOffsets: grid.lineOffsets })),
  );
  return grid;
}

export function loadGrid(docId: string): Promise<LoadedGrid> {
  return gridCache.getOrLoad(docId, () => loadGridUncached(docId));
}

/**
 * PDF 페이지 크기를 영구 저장소에서 읽고/쓰는 콜백 쌍 — rendering.ts의 getPageSizes에 주입한다.
 * 변환 시점(convertToPdf)에 이미 저장돼 있으면 그 파일만 읽고 PDF를 다시 파싱하지 않는다.
 * 없으면(이 기능 배포 전에 업로드된 문서) 기존처럼 그 자리에서 계산한 뒤 이 save로 저장해,
 * 다음부터는 그 문서도 빠른 경로를 타게 한다(자가치유).
 */
export function pageSizePersistence(docId: string): PersistedPageSizes {
  return {
    load: async () => {
      if (!(await store.hasDerived(docId, 'pagesizes'))) return null;
      const raw = await store.readDerived(docId, 'pagesizes');
      return JSON.parse(raw.toString('utf8')) as readonly PageSize[];
    },
    save: async (sizes) => {
      await store.saveDerived(docId, 'pagesizes', Buffer.from(JSON.stringify(sizes)));
    },
  };
}
```

`/api/view/info` 핸들러 안의 `getPageDimensions` 호출(기존 `:378-383`)을 대체:
```ts
      const pages =
        viewKind === 'pages'
          ? await getPageDimensions(
              view.meta.id,
              () => store.readDerived(view.meta.id, 'pdf'),
              pageSizePersistence(view.meta.id),
            ).catch(() => undefined)
          : undefined;
```

`/api/view/pages/:page` 핸들러 안의 `renderPage` 호출(기존 `:477-479`)을 대체:
```ts
        const png = await renderPage(
          view.meta.id,
          request.params.page,
          request.query.scale,
          () => store.readDerived(view.meta.id, 'pdf'),
          pageSizePersistence(view.meta.id),
        );
```

- [ ] **Step 4: 테스트 실행해 통과 확인**

Run: `npx vitest run tests/viewer-derived-cache.test.ts`
Expected: PASS (4 tests)

- [ ] **Step 5: 전체 타입 체크와 테스트 스위트 — 이제 전부 초록불이어야 한다**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 모두 PASS. `tsc`에 Task 3 완료 시점부터 남아 있던 오류가 이제 전부 사라져야 한다.

- [ ] **Step 6: 커밋**

```bash
git add src/routes/viewer.ts tests/viewer-derived-cache.test.ts
git commit -m "feat: 뷰어 조회 경로가 사전 계산된 페이지 크기·그리드 인덱스를 우선 사용"
```

---

### Task 7: viewer.js — 1페이지 이미지 요청을 프레임 구성과 병렬로 시작

**Files:**
- Modify: `public/js/viewer.js:310-343` (`buildPages`)

**Interfaces:**
- Consumes: 이미 정의된 `ensureLoaded(canvas)`(`:286-291`) — 변경 없음, 그대로 재사용.

이 태스크는 순수 클라이언트 타이밍 변경이라 이 프로젝트에 `viewer.js`용 자동 테스트 하네스가
없다(빌드 파이프라인도 없이 `public/js/`에 직접 두는 손으로 작성한 파일 — 세션 관례상 이런
UI/타이밍 변경은 살아있는 서버에 대해 브라우저로 직접 확인한다). Step 3에서 Playwright로
직접 검증한다.

- [ ] **Step 1: `public/js/viewer.js`의 `buildPages` 수정**

기존 `buildPages()`(`:310-343`)를 대체:

```js
function buildPages() {
  // 화면 근처(±1200px)에 들어오면 로딩, 멀어지면 비트맵 해제 (가상화)
  const observer = new IntersectionObserver((entries) => {
    for (const entry of entries) {
      const canvas = entry.target.querySelector('canvas');
      if (entry.isIntersecting) {
        ensureLoaded(canvas);
      } else {
        releaseCanvas(canvas);
      }
    }
  }, { rootMargin: '1200px 0px' });

  const frames = [];
  for (let page = 1; page <= pageCount; page += 1) {
    const frame = document.createElement('div');
    frame.className = 'page-frame page-loading';
    // 페이지 크기를 미리 알면 로딩 전에도 프레임 높이가 정확해 스크롤 점프가 어긋나지 않는다
    const dims = pageDimensions && pageDimensions[page - 1];
    if (dims && dims.w > 0) {
      frame.style.aspectRatio = `${dims.w} / ${dims.h}`;
      frame.style.minHeight = '0';
    }
    const canvas = document.createElement('canvas');
    canvas.dataset.page = String(page);
    const label = document.createElement('span');
    label.className = 'page-number';
    label.textContent = String(page);
    frame.append(canvas, label);
    frames.push(frame);
    if (page === 1) {
      // 나머지 프레임을 다 만들 때까지 기다리지 않고 1페이지 이미지 요청을 먼저 흘려보낸다 —
      // IntersectionObserver 콜백은 이 함수가 끝나고 다음 틱에야 배달되므로, 그대로 두면
      // "전체 프레임 구성"이 "1페이지 요청"보다 항상 먼저 끝나야 하는 불필요한 직렬 의존이
      // 생긴다. ensureLoaded는 이미 requestedScale 가드가 있어(loadPageInto가 그 값을
      // 동기적으로 세팅) 잠시 뒤 옵저버가 같은 캔버스에 대해 또 호출해도 중복 요청은 없다.
      ensureLoaded(canvas);
    }
    observer.observe(frame);
  }
  pagesContainer.replaceChildren(...frames);
}
```

- [ ] **Step 2: dev 서버 기동 확인**

Run: `npm run dev` (백그라운드로 띄워 둔다 — 이미 8000번 포트에 떠 있다면 이 단계는 건너뛴다)

- [ ] **Step 3: Playwright로 요청 순서 직접 검증**

여러 페이지짜리 PDF-뷰류 문서(이미 `data/`에 있는 `pageCount`가 큰 docx/pptx 문서의 view
토큰을 `/api/documents/:id/tokens`로 발급받아 사용하거나, 새로 하나 업로드한다) 하나를 골라
그 view token으로 아래 스크립트를 프로젝트 루트에서 실행한다(스크립트는 검증 후 지운다):

```js
// _verify-eager-page1.mjs
import { chromium } from 'playwright';

const VIEWER_URL = process.argv[2]; // 예: http://127.0.0.1:8000/viewer.html?token=...
const browser = await chromium.launch();
const page = await browser.newPage();
const events = [];
page.on('request', (req) => {
  if (req.url().includes('/api/view/pages/1')) events.push({ t: Date.now(), what: 'page1-request' });
});
await page.exposeFunction('reportFramesBuilt', () => events.push({ t: Date.now(), what: 'frames-built' }));
await page.addInitScript(() => {
  const original = Element.prototype.replaceChildren;
  Element.prototype.replaceChildren = function (...args) {
    if (this.id === 'pages') window.reportFramesBuilt();
    return original.apply(this, args);
  };
});
await page.goto(VIEWER_URL, { waitUntil: 'networkidle' });
await page.waitForTimeout(500);
console.log(events);
const page1Idx = events.findIndex((e) => e.what === 'page1-request');
const framesIdx = events.findIndex((e) => e.what === 'frames-built');
console.log(
  page1Idx !== -1 && framesIdx !== -1 && page1Idx <= framesIdx
    ? 'PASS: 1페이지 요청이 프레임 구성 완료와 동시에 또는 그 전에 나감'
    : 'FAIL: 1페이지 요청이 프레임 구성보다 늦음',
);
await browser.close();
```

Run: `node _verify-eager-page1.mjs "http://127.0.0.1:8000/viewer.html?token=<발급받은 토큰>"`
Expected: `PASS: 1페이지 요청이 프레임 구성 완료와 동시에 또는 그 전에 나감` 출력.

검증이 끝나면 스크립트를 지운다: `rm _verify-eager-page1.mjs`

- [ ] **Step 4: 회귀 확인 — 페이지가 여전히 정상적으로 다 로드/해제되는지**

같은 브라우저 세션에서 수동으로(또는 Playwright로) 스크롤을 끝까지 내려 마지막 페이지까지
정상적으로 로드되는지, 다시 위로 스크롤할 때 중간 페이지들이 다시 로드되는지 확인한다
(가상화 동작 자체는 이번 변경으로 건드리지 않았지만, 1페이지를 다르게 다루기 시작했으므로
회귀가 없는지 한 번은 직접 확인한다).

- [ ] **Step 5: 커밋**

```bash
git add public/js/viewer.js
git commit -m "perf: 뷰어가 1페이지 이미지 요청을 전체 프레임 구성과 병렬로 시작"
```

---

## 최종 확인

모든 태스크가 끝나면:

```bash
npx tsc --noEmit && npx vitest run
```

모두 초록불이어야 하며, `npm run build:editor`(관련 없는 편집기 번들이지만 회귀가 없는지)도
한 번 더 돌려 확인한다.
