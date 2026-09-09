# xlsx 지연(Lazy) Canonical Grid Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace `Room.canonicalGrid: GridDocument`(파일 전체를 파싱해 메모리에 통째로 올리는
구조)를, 편집된 칸만 메모리에 두고 나머지는 온디스크 원본 행 저장소에서 그때그때 읽어오는
`LazyGridStore`로 교체한다.

**Architecture:** 새 독립 모듈(`src/xlsx/lazyGridStore.ts`)이 원본 xlsx를 스트리밍으로 한 번
훑어 문서별 임시 JSONL 파일(원본 행 데이터) + 바이트 오프셋 인덱스를 만든다. 읽기는 "더티
오버레이(편집된 칸) → 읽기 캐시(최근 조회한 원본 행) → 디스크"의 3단 순서를 따르고, 쓰기는
항상 더티 오버레이에만 반영된다. 행/열 삽입·삭제는 실제 데이터를 옮기는 대신 "논리 좌표 →
원본 좌표" 변환 배열만 스플라이스한다.

**Tech Stack:** TypeScript, ExcelJS `stream.xlsx.WorkbookReader`(스트리밍 리더), `jszip`(테마
XML 단독 추출), Node `fs`(온디스크 저장/읽기), Vitest.

**Spec:** `docs/superpowers/specs/2026-09-09-xlsx-lazy-canonical-grid-design.md`

## Global Constraints

- 모든 xlsx 문서(크기 무관)에 통일 적용한다 — 작은 파일 전용 별도 경로를 만들지 않는다(스펙
  "범위" 절).
- `dirtyOverlay`에 값이 있는 칸이 유일한 진실이다 — 읽기 캐시는 언제 밀려나도 안전해야 한다
  (편집된 값을 절대 읽기 캐시에만 두지 않는다).
- 캐시 단위는 칸이 아니라 **행**이다(`"sheetIdx:원본행번호" -> 파싱된 행 전체`).
- `readCache` 기본 상한은 스토어 전체 5,000행(상수로 두고 튜닝 가능하게 한다).
- 온디스크 임시 파일 이름은 `docId` + `randomUUID()`를 포함해 충돌을 피한다(기존
  `exportStreaming.ts`의 임시 파일 명명 관례와 동일).
- 기존 구조연산 회귀 테스트(`tests/roomStructureSync.test.ts`, `tests/roomRowColShift.test.ts`)는
  이 재설계 이후에도 전부 그대로 통과해야 한다 — 이 계획의 각 관련 태스크는 이 테스트들을
  마지막에 재실행해 확인한다(Task 13).
- fast-formula-parser의 `onCell`/`onRange` 콜백은 동기 전용이다 — 콜백 안에서 예외를 던지지
  않는다(라이브러리가 재포장해버림). 대신 클로저 플래그 + `parser.parse()` 이후 확인 패턴을
  쓴다(기존 `rangeTooLarge`와 동일한 원리, `xlsx/formulas.ts` 기존 코드 참고).

---

## Phase 1 — `LazyGridStore` 독립 모듈 (아직 room.ts에 연결 안 함)

### Task 1: `spreadsheet.ts`의 재사용 가능한 헬퍼를 export로 전환

**Files:**
- Modify: `src/spreadsheet.ts:250` (`resolveColor` — private 유지, `extractStyle` 내부에서만
  쓰임, export 불필요)
- Modify: `src/spreadsheet.ts:288` (`displayValue` → `export function displayValue`)
- Modify: `src/spreadsheet.ts:314` (`parseMergeRange` → `export function parseMergeRange`)
- Modify: `src/spreadsheet.ts:326` (`class StylePalette` → `export class StylePalette`)
- Modify: `src/spreadsheet.ts:367` (`extractStyle` → `export function extractStyle`)
- Test: `tests/spreadsheet-derived.test.ts` (기존 파일에 이 함수들이 여전히 같은 결과를 내는지
  확인하는 테스트 추가)

**Interfaces:**
- Produces: `displayValue(cell: ExcelJS.Cell): string`, `parseMergeRange(range: string): GridMerge | null`,
  `class StylePalette { register(style): number | undefined; toArray(): GridStyle[] }`,
  `extractStyle(cell: ExcelJS.Cell, palette: StylePalette, themeColors: readonly string[], sheetProtected: boolean): number | undefined`
  — 이 네 가지 이름과 시그니처를 Task 3(스트리밍 ingestion)이 그대로 가져다 쓴다.

- [ ] **Step 1: 이 네 선언 앞에 `export`를 추가한다(로직은 전혀 안 바꾼다)**

```ts
// src/spreadsheet.ts
export function displayValue(cell: ExcelJS.Cell): string { /* 기존 그대로 */ }
export function parseMergeRange(range: string): GridMerge | null { /* 기존 그대로 */ }
export class StylePalette { /* 기존 그대로 */ }
export function extractStyle(
  cell: ExcelJS.Cell,
  palette: StylePalette,
  themeColors: readonly string[],
  sheetProtected: boolean,
): number | undefined { /* 기존 그대로 */ }
```

- [ ] **Step 2: 회귀 테스트 추가 — 기존 `parseXlsxToGrid`가 이 함수들을 통해 계산한 결과와
  직접 호출한 결과가 같은지(순수히 export 여부만 바뀌었다는 걸 확인하는 안전망)**

```ts
// tests/spreadsheet-derived.test.ts 에 추가
import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { displayValue, extractStyle, parseMergeRange, StylePalette } from '../src/spreadsheet.js';

describe('spreadsheet.ts: newly-exported helpers behave identically to before', () => {
  test('displayValue/extractStyle/parseMergeRange work when called directly (not just via parseXlsxToGrid)', () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.getCell('A1').value = 'hello';
    sheet.getCell('A1').font = { bold: true };
    sheet.mergeCells('A1:B1');

    expect(displayValue(sheet.getCell('A1'))).toBe('hello');

    const palette = new StylePalette();
    const styleIndex = extractStyle(sheet.getCell('A1'), palette, [], false);
    expect(styleIndex).toBe(0);
    expect(palette.toArray()[0]).toEqual({ b: 1 });

    expect(parseMergeRange('A1:B1')).toEqual({ r: 1, c: 1, rs: 1, cs: 2 });
  });
});
```

- [ ] **Step 3: 테스트 실행 확인**

Run: `npx vitest run tests/spreadsheet-derived.test.ts`
Expected: PASS (기존 테스트 포함 전부)

- [ ] **Step 4: 전체 타입체크·테스트**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 전부 그대로 통과(로직 변경 없음, export만 추가했으므로 회귀 없어야 함)

- [ ] **Step 5: 커밋**

```bash
git add src/spreadsheet.ts tests/spreadsheet-derived.test.ts
git commit -m "refactor(spreadsheet): export displayValue/extractStyle/parseMergeRange/StylePalette for reuse in streaming ingestion"
```

---

### Task 2: `jszip` 의존성 추가 + 테마 색상 추출 헬퍼

**Files:**
- Modify: `package.json` (dependencies에 `jszip` 추가)
- Create: `src/xlsx/themeExtract.ts`
- Test: `tests/xlsxThemeExtract.test.ts`

**Interfaces:**
- Consumes: `parseThemeColors(themeXml: string | undefined): readonly string[]`
  (`src/spreadsheet.ts`, 이미 export됨, 변경 없음)
- Produces: `extractThemeColorsFromZip(xlsxBytes: Buffer): Promise<readonly string[]>` — Task 3이
  이 함수를 부른다.

- [ ] **Step 1: `jszip`을 직접 의존성으로 추가**

```bash
npm install jszip
```

- [ ] **Step 2: 실패하는 테스트를 먼저 쓴다**

```ts
// tests/xlsxThemeExtract.test.ts
import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { extractThemeColorsFromZip } from '../src/xlsx/themeExtract.js';

describe('xlsx/themeExtract.ts: extractThemeColorsFromZip', () => {
  test('extracts the same theme colors that ExcelJS.Workbook.xlsx.load would expose via workbook.model.themes', async () => {
    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet('Sheet1').getCell('A1').value = 'x';
    const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);

    const viaZip = await extractThemeColorsFromZip(bytes);

    const reloaded = new ExcelJS.Workbook();
    await reloaded.xlsx.load(bytes as unknown as ArrayBuffer);
    const themes = (reloaded.model as unknown as { themes?: Record<string, string> }).themes;
    const { parseThemeColors } = await import('../src/spreadsheet.js');
    const viaFullLoad = parseThemeColors(themes?.theme1);

    expect(viaZip).toEqual(viaFullLoad);
    expect(viaZip.length).toBeGreaterThan(0);
  });

  test('returns an empty array for a zip with no theme1.xml entry (never throws)', async () => {
    const JSZip = (await import('jszip')).default;
    const zip = new JSZip();
    zip.file('empty.txt', 'nothing here');
    const bytes = Buffer.from(await zip.generateAsync({ type: 'nodebuffer' }));

    await expect(extractThemeColorsFromZip(bytes)).resolves.toEqual([]);
  });
});
```

- [ ] **Step 3: 테스트 실패 확인**

Run: `npx vitest run tests/xlsxThemeExtract.test.ts`
Expected: FAIL — `Cannot find module '../src/xlsx/themeExtract.js'`

- [ ] **Step 4: 최소 구현**

```ts
// src/xlsx/themeExtract.ts
import JSZip from 'jszip';
import { parseThemeColors } from '../spreadsheet.js';

/**
 * ExcelJS의 스트리밍 리더(ExcelJS.stream.xlsx.WorkbookReader)는 xl/theme/theme1.xml을 파싱하지
 * 않는다(ExcelJS 소스 확인 완료 — 알려진 항목만 처리하고 나머지는 버림). 테마 색을 참조하는 셀
 * 서식을 스트리밍 ingestion에서도 올바르게 해석하려면 이 파일이 필요하므로, jszip으로 이 항목
 * 하나만 별도로 꺼낸다(작은 파일이라 압축 해제 비용 무시할 만함).
 */
export async function extractThemeColorsFromZip(xlsxBytes: Buffer): Promise<readonly string[]> {
  const zip = await JSZip.loadAsync(xlsxBytes);
  const entry = zip.file('xl/theme/theme1.xml');
  if (!entry) return [];
  const xml = await entry.async('text');
  return parseThemeColors(xml);
}
```

- [ ] **Step 5: 테스트 통과 확인**

Run: `npx vitest run tests/xlsxThemeExtract.test.ts`
Expected: PASS (2 tests)

- [ ] **Step 6: 커밋**

```bash
git add package.json package-lock.json src/xlsx/themeExtract.ts tests/xlsxThemeExtract.test.ts
git commit -m "feat(xlsx): add jszip dependency + standalone theme1.xml color extraction for streaming ingestion"
```

---

### Task 3: `LazyGridStore` 데이터 모델 + 스트리밍 ingestion

**Files:**
- Create: `src/xlsx/lazyGridStore.ts`
- Test: `tests/lazyGridStore.test.ts`

**Interfaces:**
- Consumes: `displayValue`/`parseMergeRange`/`StylePalette`/`extractStyle`(Task 1),
  `extractThemeColorsFromZip`(Task 2), `MAX_ROWS`/`MAX_COLS`/`YIELD_EVERY_N_ROWS`/
  `yieldToEventLoop`(이미 export된 `spreadsheet.ts`의 기존 상수/함수, 변경 없음),
  `GridCell`/`GridStyle`/`GridMerge`(`spreadsheet.ts`, 이미 export됨)
- Produces:
  - `interface LazyGridSheetMeta { name: string; rowCount: number; colCount: number; colWidths: readonly number[]; merges: readonly GridMerge[]; freezeRows: number; freezeCols: number }`
  - `interface LazyGridStore { readonly docId: string; readonly rowStorePath: string; readonly sheetMeta: LazyGridSheetMeta[]; readonly styles: GridStyle[]; readonly rowIndex: Array<Array<{ byteOffset: number; byteLength: number } | undefined>>; readonly logicalToOriginalRow: number[][]; readonly logicalToOriginalCol: number[][]; readonly dirtyOverlay: Map<string, GridCell | 0>; readonly readCache: Map<string, ReadonlyArray<GridCell | 0>>; readonly readCacheOrder: string[] }`
  - `async function ingestXlsxLazily(docId: string, xlsxBytes: Buffer): Promise<LazyGridStore>`
  - 이후 태스크(4, 5, 6)가 이 `LazyGridStore`에 메서드처럼 동작하는 독립 함수들(`getCell`,
    `getRow`, `setDirtyCell`, `applyRowShift`, `applyColShift`, `growSheet`, `closeLazyGridStore`)
    을 이 파일에 계속 추가한다.

이 태스크는 **ingestion(수집)만** 구현한다 — 읽기/쓰기/구조연산은 Task 4-6.

- [ ] **Step 1: 작은 픽스처로 기본 ingestion을 검증하는 실패 테스트**

```ts
// tests/lazyGridStore.test.ts
import { randomUUID } from 'node:crypto';
import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { ingestXlsxLazily } from '../src/xlsx/lazyGridStore.js';

async function makeSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 'Name';
  sheet.getCell('B1').value = 'Score';
  sheet.getCell('A2').value = 'Alice';
  sheet.getCell('B2').value = 10;
  sheet.getRow(1).getCell(1).font = { bold: true };
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

describe('xlsx/lazyGridStore.ts: ingestXlsxLazily', () => {
  test('produces correct sheet metadata and an on-disk row store for a small file', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);

    expect(store.sheetMeta).toHaveLength(1);
    expect(store.sheetMeta[0].name).toBe('Sheet1');
    expect(store.sheetMeta[0].rowCount).toBe(2);
    expect(store.sheetMeta[0].colCount).toBe(2);
    expect(store.styles.length).toBeGreaterThanOrEqual(1); // A1의 bold 서식

    // 항등 매핑으로 초기화돼야 한다
    expect(store.logicalToOriginalRow[0]).toEqual([0, 1]);
    expect(store.logicalToOriginalCol[0]).toEqual([0, 1]);

    // 온디스크 저장소에 실제로 두 행이 기록됐어야 한다
    expect(store.rowIndex[0][0]).toBeDefined();
    expect(store.rowIndex[0][1]).toBeDefined();

    await import('node:fs/promises').then((fs) => fs.rm(store.rowStorePath, { force: true }));
  });

  test('trims trailing formatting-only rows the same way parseXlsxToGrid does', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.getCell('A1').value = 'hi';
    sheet.getRow(10).height = 30; // 값 없이 서식(행 높이)만 — 꼬리 행
    const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);

    const store = await ingestXlsxLazily(randomUUID(), bytes);
    // parseXlsxToGrid와 동일하게: 실제 값이 있는 마지막 행(1)까지만 rowCount로 잡는다.
    expect(store.sheetMeta[0].rowCount).toBe(1);

    await import('node:fs/promises').then((fs) => fs.rm(store.rowStorePath, { force: true }));
  });

  test('resolves theme colors the same way as the non-streaming parser', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.getCell('A1').value = 'x';
    sheet.getCell('A1').font = { color: { theme: 1 } }; // dk1 테마 색
    const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);

    const { parseXlsxToGrid } = await import('../src/spreadsheet.js');
    const viaFullParse = await parseXlsxToGrid(bytes);
    const cellViaFullParse = viaFullParse.sheets[0].rows[0][0];
    const styleViaFullParse =
      cellViaFullParse !== 0 && cellViaFullParse.s !== undefined
        ? viaFullParse.styles[cellViaFullParse.s]
        : undefined;

    const store = await ingestXlsxLazily(randomUUID(), bytes);
    const fs = await import('node:fs/promises');
    const raw = await fs.readFile(store.rowStorePath);
    const entry = store.rowIndex[0][0]!;
    const row = JSON.parse(raw.subarray(entry.byteOffset, entry.byteOffset + entry.byteLength).toString('utf8'));
    const cellViaStreaming = row[0];
    const styleViaStreaming =
      cellViaStreaming !== 0 && cellViaStreaming.s !== undefined ? store.styles[cellViaStreaming.s] : undefined;

    expect(styleViaStreaming).toEqual(styleViaFullParse);
    await fs.rm(store.rowStorePath, { force: true });
  });
});
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `npx vitest run tests/lazyGridStore.test.ts`
Expected: FAIL — `Cannot find module '../src/xlsx/lazyGridStore.js'`

- [ ] **Step 3: 구현**

```ts
// src/xlsx/lazyGridStore.ts
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import {
  displayValue,
  extractStyle,
  parseMergeRange,
  StylePalette,
  yieldToEventLoop,
  YIELD_EVERY_N_ROWS,
  MAX_ROWS,
  MAX_COLS,
  type GridCell,
  type GridMerge,
  type GridStyle,
} from '../spreadsheet.js';
import { extractThemeColorsFromZip } from './themeExtract.js';

export interface LazyGridSheetMeta {
  name: string;
  rowCount: number;
  colCount: number;
  colWidths: readonly number[];
  merges: readonly GridMerge[];
  freezeRows: number;
  freezeCols: number;
}

export interface LazyRowIndexEntry {
  readonly byteOffset: number;
  readonly byteLength: number;
}

export interface LazyGridStore {
  readonly docId: string;
  readonly rowStorePath: string;
  readonly sheetMeta: LazyGridSheetMeta[];
  readonly styles: GridStyle[];
  /** sheetIdx -> 원본 행 번호(ingestion 시점 기준, 0-based) -> 디스크 위치. ingestion 이후 불변. */
  readonly rowIndex: Array<Array<LazyRowIndexEntry | undefined>>;
  /** sheetIdx -> 논리 행 -> 원본 행(-1 = 원본에 없던 새 행). */
  readonly logicalToOriginalRow: number[][];
  readonly logicalToOriginalCol: number[][];
  /** gridCellKey -> 실제로 편집된 칸의 현재 값. */
  readonly dirtyOverlay: Map<string, GridCell | 0>;
  /** "sheetIdx:원본행번호" -> 파싱된 행(캐시 단위는 행). */
  readonly readCache: Map<string, ReadonlyArray<GridCell | 0>>;
  /** readCache의 LRU 순서(맨 앞이 가장 오래됨) — Map은 삽입 순서를 유지하므로 재사용한다. */
  readCacheLimit: number;
}

const DEFAULT_READ_CACHE_ROWS = 5000;

function identityArray(length: number): number[] {
  return Array.from({ length }, (_, i) => i);
}

export async function ingestXlsxLazily(docId: string, xlsxBytes: Buffer): Promise<LazyGridStore> {
  const themeColors = await extractThemeColorsFromZip(xlsxBytes);
  const rowStorePath = path.join(os.tmpdir(), `docview-lazygrid-${docId}-${randomUUID()}.jsonl`);
  const writeHandle = await fs.open(rowStorePath, 'w');

  const sheetMeta: LazyGridSheetMeta[] = [];
  const rowIndex: Array<Array<LazyRowIndexEntry | undefined>> = [];
  const palette = new StylePalette();
  let byteOffset = 0;

  try {
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(xlsxBytes, {
      worksheets: 'emit',
      sharedStrings: 'cache',
      hyperlinks: 'ignore',
      styles: 'cache',
      entries: 'ignore',
    });

    let sheetIdx = -1;
    for await (const worksheetReader of reader) {
      sheetIdx += 1;
      const thisSheetIndex: Array<LazyRowIndexEntry | undefined> = [];
      rowIndex.push(thisSheetIndex);

      let maxRowWithValue = 0;
      let maxColWithValue = 0;
      let originalRow = -1;
      // sheetProtection은 스트리밍 리더에서 신뢰성 있게 얻을 수 있는 API가 없어 false로 둔다
      // (알려진 스코프 제한 — 시트 보호 잠금 서식(lk)이 스트리밍 경로에서는 반영되지 않음).
      const sheetProtected = false;

      for await (const row of worksheetReader) {
        originalRow += 1;
        if (originalRow >= MAX_ROWS) break;
        const cells: Array<GridCell | 0> = [];
        let rowHasValue = false;
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          if (colNumber > MAX_COLS) return;
          const text = displayValue(cell);
          const styleRef = extractStyle(cell, palette, themeColors, sheetProtected);
          if (text === '' && styleRef === undefined) return;
          if (colNumber > maxColWithValue) maxColWithValue = colNumber;
          while (cells.length < colNumber) cells.push(0);
          cells[colNumber - 1] = styleRef === undefined ? { v: text } : { v: text, s: styleRef };
          if (text !== '') rowHasValue = true;
        });
        if (rowHasValue) maxRowWithValue = originalRow + 1;

        const json = JSON.stringify(cells);
        const buf = Buffer.from(json, 'utf8');
        await writeHandle.write(buf, 0, buf.length, byteOffset);
        thisSheetIndex[originalRow] = { byteOffset, byteLength: buf.length };
        byteOffset += buf.length;

        if ((originalRow + 1) % YIELD_EVERY_N_ROWS === 0) await yieldToEventLoop();
      }

      const rowCount = Math.min(maxRowWithValue, MAX_ROWS);
      const colCount = Math.min(maxColWithValue, MAX_COLS);
      const rawSheet = worksheetReader as unknown as {
        model?: { merges?: readonly string[] };
        views?: ReadonlyArray<{ state?: string; ySplit?: number; xSplit?: number }>;
        columns?: ReadonlyArray<{ width?: number } | undefined> | null;
      };
      const merges = (rawSheet.model?.merges ?? [])
        .map(parseMergeRange)
        .filter((m): m is GridMerge => m !== null && m.r <= rowCount && m.c <= colCount);
      const frozenView = rawSheet.views?.find((v) => v.state === 'frozen');
      const colWidths: number[] = [];
      for (let c = 1; c <= colCount; c += 1) {
        const width = rawSheet.columns?.[c - 1]?.width;
        colWidths.push(width ? Math.round(width * 7.2) : 100);
      }

      sheetMeta.push({
        name: worksheetReader.name,
        rowCount,
        colCount,
        colWidths,
        merges,
        freezeRows: frozenView?.ySplit ?? 0,
        freezeCols: frozenView?.xSplit ?? 0,
      });
    }
  } finally {
    await writeHandle.close();
  }

  return {
    docId,
    rowStorePath,
    sheetMeta,
    styles: palette.toArray(),
    rowIndex,
    logicalToOriginalRow: sheetMeta.map((_, i) => identityArray(rowIndex[i]?.length ?? 0)),
    logicalToOriginalCol: sheetMeta.map((meta) => identityArray(meta.colCount)),
    dirtyOverlay: new Map(),
    readCache: new Map(),
    readCacheLimit: DEFAULT_READ_CACHE_ROWS,
  };
}
```

**참고 — 구현 중 실제 ExcelJS 스트리밍 리더 API를 검증하며 다음을 확인할 것**(이 항목들은
`node_modules/exceljs/lib/stream/xlsx/worksheet-reader.js`/`workbook-reader.js` 소스를 직접
읽어 이미 한 번 검증됐지만, 버전이 다를 수 있으니 재확인 권장):
- `new ExcelJS.stream.xlsx.WorkbookReader(input, options)`의 `input`으로 `Buffer`를 바로 줄 수
  있는지, 아니면 `Readable` 스트림으로 감싸야 하는지(`_getStream`이 `Buffer`를 처리 못 하면
  `Readable.from(xlsxBytes)`로 감싼다).
- `for await (const worksheetReader of reader)`가 `WorksheetReader` 인스턴스를 준다는 것,
  그리고 `for await (const row of worksheetReader)`가 각 행의 `Row` 객체(`.eachCell`,
  `.height` 있음, 일반 API와 동일)를 준다는 것(둘 다 `Symbol.asyncIterator` 구현 확인됨,
  위 소스 88-113행).
- `worksheetReader.name`, `.model.merges`, `.views`, `.columns`가 비streaming API와 같은
  모양으로 채워지는지 — 안 채워지면 `rawSheet` 캐스팅 부분을 실제 API에 맞게 조정.

- [ ] **Step 4: 테스트 실행, 실패하면 위 API 가정을 실제 동작에 맞게 조정**

Run: `npx vitest run tests/lazyGridStore.test.ts`
Expected: 처음엔 API 가정이 틀려 실패할 수 있음 — 에러 메시지를 보고 위 "참고" 항목대로
`lazyGridStore.ts`의 스트리밍 API 사용 부분만 조정(테스트 자체는 바꾸지 않는다). 최종적으로
PASS.

- [ ] **Step 5: 실제 대용량 파일로 수동 검증(자동 테스트에는 포함하지 않음 — 로컬 전용 파일 경로에
  의존하므로)**

로컬에서 임시 스크립트로 `ingestXlsxLazily`를 292,244행 파일에 돌려, (a) 파싱 중 메모리가
안정적인지(피크가 수백MB 이내), (b) `sheetMeta[0].rowCount`가 292,244와 일치하는지 확인한다.
검증 후 스크립트는 지운다(이 프로젝트의 기존 관례 — `scratch-*.mjs`로 만들고 확인 후 삭제).

- [ ] **Step 6: 전체 테스트 + 타입체크**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 전부 PASS

- [ ] **Step 7: 커밋**

```bash
git add src/xlsx/lazyGridStore.ts tests/lazyGridStore.test.ts
git commit -m "feat(xlsx): add LazyGridStore streaming ingestion (LazyGridStore is not yet wired into room.ts)"
```

---

### Task 4: 읽기 API — `getCell`/`getRow`

**Files:**
- Modify: `src/xlsx/lazyGridStore.ts` (Task 3의 파일에 이어서 추가)
- Test: `tests/lazyGridStore.test.ts` (이어서 추가)

**Interfaces:**
- Consumes: Task 3의 `LazyGridStore` 타입, `gridCellKey`/`parseGridCellKey`(`src/collab/gridShared.ts`, 이미 export됨)
- Produces:
  - `async function getRow(store: LazyGridStore, sheetIdx: number, logicalRow: number): Promise<ReadonlyArray<GridCell | 0>>`
  - `async function getCell(store: LazyGridStore, sheetIdx: number, logicalRow: number, logicalCol: number): Promise<GridCell | 0>`
  - Task 8(페이지 시딩)과 Task 9(수식 계산)이 이 두 함수를 소비한다.

- [ ] **Step 1: 실패하는 테스트**

```ts
// tests/lazyGridStore.test.ts 에 추가
import { getCell, getRow } from '../src/xlsx/lazyGridStore.js';
import { gridCellKey } from '../src/collab/gridShared.js';

describe('xlsx/lazyGridStore.ts: getCell/getRow', () => {
  test('reads an unedited cell from disk on first access, then from cache on repeat access', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);

    expect(await getCell(store, 0, 0, 0)).toEqual({ v: 'Name', s: expect.any(Number) });
    expect(await getCell(store, 0, 1, 0)).toEqual({ v: 'Alice' });
    // 캐시가 채워졌는지 — 같은 원본 행에 대해 readCache에 항목이 생겼어야 한다.
    expect(store.readCache.has('0:0')).toBe(true);

    await fs.rm(store.rowStorePath, { force: true });
  });

  test('a dirty overlay entry always wins over the on-disk original, even if cached', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    await getCell(store, 0, 1, 0); // 먼저 캐시를 데운다
    store.dirtyOverlay.set(gridCellKey(0, 1, 0), { v: 'Alice Kim' });

    expect(await getCell(store, 0, 1, 0)).toEqual({ v: 'Alice Kim' });

    await fs.rm(store.rowStorePath, { force: true });
  });

  test('a coordinate beyond logicalToOriginalRow/Col bounds (newly inserted, unedited) returns 0 without touching disk', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    store.logicalToOriginalRow[0].splice(1, 0, -1); // 1번 위치에 "새 행" 삽입

    expect(await getCell(store, 0, 1, 0)).toBe(0); // 새로 삽입된 빈 행
    expect(await getCell(store, 0, 2, 0)).toEqual({ v: 'Alice' }); // 밀려난 원래 2번째 행

    await fs.rm(store.rowStorePath, { force: true });
  });

  test('readCache evicts the oldest row once over the configured limit, and re-reads correctly after eviction', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    store.readCacheLimit = 1;

    await getRow(store, 0, 0);
    expect(store.readCache.has('0:0')).toBe(true);
    await getRow(store, 0, 1);
    expect(store.readCache.has('0:0')).toBe(false); // 밀려남
    expect(store.readCache.has('0:1')).toBe(true);

    // 밀려난 뒤에도 다시 읽으면 같은 값
    expect(await getCell(store, 0, 0, 0)).toEqual({ v: 'Name', s: expect.any(Number) });

    await fs.rm(store.rowStorePath, { force: true });
  });
});
```

(테스트 파일 상단에 `import fs from 'node:fs/promises';`, `import { randomUUID } from 'node:crypto';`
가 이미 Task 3에서 추가돼 있어야 함 — 없으면 추가한다.)

- [ ] **Step 2: 테스트 실패 확인**

Run: `npx vitest run tests/lazyGridStore.test.ts`
Expected: FAIL — `getCell`/`getRow` not exported

- [ ] **Step 3: 구현**

```ts
// src/xlsx/lazyGridStore.ts 에 추가
import { gridCellKey } from '../collab/gridShared.js';

async function readRowFromDisk(store: LazyGridStore, sheetIdx: number, originalRow: number): Promise<ReadonlyArray<GridCell | 0>> {
  const entry = store.rowIndex[sheetIdx]?.[originalRow];
  if (!entry) return [];
  const handle = await fs.open(store.rowStorePath, 'r');
  try {
    const buf = Buffer.alloc(entry.byteLength);
    await handle.read(buf, 0, entry.byteLength, entry.byteOffset);
    return JSON.parse(buf.toString('utf8')) as Array<GridCell | 0>;
  } finally {
    await handle.close();
  }
}

function cacheKey(sheetIdx: number, originalRow: number): string {
  return `${sheetIdx}:${originalRow}`;
}

function touchCache(store: LazyGridStore, key: string, row: ReadonlyArray<GridCell | 0>): void {
  // Map은 삽입 순서를 유지한다 — 다시 넣으면(delete 후 set) 맨 뒤(=가장 최근)로 이동한다.
  store.readCache.delete(key);
  store.readCache.set(key, row);
  while (store.readCache.size > store.readCacheLimit) {
    const oldest = store.readCache.keys().next().value;
    if (oldest === undefined) break;
    store.readCache.delete(oldest);
  }
}

export async function getRow(
  store: LazyGridStore,
  sheetIdx: number,
  logicalRow: number,
): Promise<ReadonlyArray<GridCell | 0>> {
  const originalRow = store.logicalToOriginalRow[sheetIdx]?.[logicalRow];
  if (originalRow === undefined || originalRow < 0) return [];
  const key = cacheKey(sheetIdx, originalRow);
  const cached = store.readCache.get(key);
  if (cached) {
    touchCache(store, key, cached); // LRU 최신화
    return cached;
  }
  const row = await readRowFromDisk(store, sheetIdx, originalRow);
  touchCache(store, key, row);
  return row;
}

export async function getCell(
  store: LazyGridStore,
  sheetIdx: number,
  logicalRow: number,
  logicalCol: number,
): Promise<GridCell | 0> {
  const key = gridCellKey(sheetIdx, logicalRow, logicalCol);
  if (store.dirtyOverlay.has(key)) return store.dirtyOverlay.get(key)!;

  const originalCol = store.logicalToOriginalCol[sheetIdx]?.[logicalCol];
  if (originalCol === undefined || originalCol < 0) return 0;

  const row = await getRow(store, sheetIdx, logicalRow);
  return row[originalCol] ?? 0;
}
```

- [ ] **Step 4: 테스트 실행**

Run: `npx vitest run tests/lazyGridStore.test.ts`
Expected: PASS (전체)

- [ ] **Step 5: 타입체크 + 전체 테스트**

Run: `npx tsc --noEmit && npx vitest run`
Expected: PASS

- [ ] **Step 6: 커밋**

```bash
git add src/xlsx/lazyGridStore.ts tests/lazyGridStore.test.ts
git commit -m "feat(xlsx): add LazyGridStore getCell/getRow read path (dirty overlay -> read cache -> disk)"
```

---

### Task 5: 쓰기 API + 더티 키 이동(구조연산 대비)

**Files:**
- Modify: `src/xlsx/lazyGridStore.ts`
- Test: `tests/lazyGridStore.test.ts`

**Interfaces:**
- Consumes: 기존 `shiftDirtyCellKeys`의 로직(`src/collab/room.ts`, 현재 `Set<string>` 전용) —
  이 태스크는 그 로직을 **일반화**해 `LazyGridStore.dirtyOverlay`(`Map`)에도 적용할 수 있게
  만든다. `room.ts`의 기존 함수는 이 태스크에서 건드리지 않는다(Task 11에서 room.ts 쪽을
  이 새 버전으로 교체).
- Produces:
  - `function setDirtyCell(store: LazyGridStore, sheetIdx: number, row: number, col: number, value: GridCell | 0): void`
  - `function shiftDirtyOverlayKeys(overlay: Map<string, GridCell | 0>, sheetIdx: number, axis: 'row' | 'col', at: number, count: number, insert: boolean): void`

- [ ] **Step 1: 실패하는 테스트**

```ts
// tests/lazyGridStore.test.ts 에 추가
import { setDirtyCell, shiftDirtyOverlayKeys } from '../src/xlsx/lazyGridStore.js';

describe('xlsx/lazyGridStore.ts: setDirtyCell/shiftDirtyOverlayKeys', () => {
  test('setDirtyCell makes getCell return the new value immediately', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    setDirtyCell(store, 0, 1, 0, { v: 'Alice Kim' });
    expect(await getCell(store, 0, 1, 0)).toEqual({ v: 'Alice Kim' });
    await fs.rm(store.rowStorePath, { force: true });
  });

  test('shiftDirtyOverlayKeys moves an edited coordinate down when a row is inserted above it', () => {
    const overlay = new Map<string, GridCell | 0>([[gridCellKey(0, 5, 2), { v: 'edited' }]]);
    shiftDirtyOverlayKeys(overlay, 0, 'row', 2, 1, true); // 2번 위치에 행 1개 삽입
    expect(overlay.has(gridCellKey(0, 5, 2))).toBe(false);
    expect(overlay.get(gridCellKey(0, 6, 2))).toEqual({ v: 'edited' });
  });

  test('shiftDirtyOverlayKeys drops an edited coordinate that falls inside a deleted range', () => {
    const overlay = new Map<string, GridCell | 0>([[gridCellKey(0, 5, 2), { v: 'edited' }]]);
    shiftDirtyOverlayKeys(overlay, 0, 'row', 4, 3, false); // 4,5,6번 행 삭제(5가 그 안에 포함)
    expect(overlay.has(gridCellKey(0, 5, 2))).toBe(false);
    expect(overlay.size).toBe(0);
  });

  test('shiftDirtyOverlayKeys leaves coordinates in other sheets untouched', () => {
    const overlay = new Map<string, GridCell | 0>([[gridCellKey(1, 5, 2), { v: 'other sheet' }]]);
    shiftDirtyOverlayKeys(overlay, 0, 'row', 0, 1, true);
    expect(overlay.get(gridCellKey(1, 5, 2))).toEqual({ v: 'other sheet' });
  });
});
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `npx vitest run tests/lazyGridStore.test.ts`
Expected: FAIL — 함수 없음

- [ ] **Step 3: 구현** (`shiftDirtyOverlayKeys`는 `room.ts`의 기존 `shiftDirtyCellKeys`와
  정확히 같은 규칙을 쓰되 `Set` 대신 `Map`을 순회한다)

```ts
// src/xlsx/lazyGridStore.ts 에 추가
import { parseGridCellKey } from '../collab/gridShared.js';

export function setDirtyCell(
  store: LazyGridStore,
  sheetIdx: number,
  row: number,
  col: number,
  value: GridCell | 0,
): void {
  store.dirtyOverlay.set(gridCellKey(sheetIdx, row, col), value);
}

/** room.ts의 shiftDirtyCellKeys(Set 전용)와 동일한 규칙 — Map(dirtyOverlay)에 적용하는 버전. */
export function shiftDirtyOverlayKeys(
  overlay: Map<string, GridCell | 0>,
  sheetIdx: number,
  axis: 'row' | 'col',
  at: number,
  count: number,
  insert: boolean,
): void {
  if (overlay.size === 0) return;
  const shifted: Array<[string, GridCell | 0]> = [];
  for (const [key, value] of overlay) {
    const parsed = parseGridCellKey(key);
    if (!parsed || parsed.sheetIdx !== sheetIdx) continue;
    let { row, col } = parsed;
    const coord = axis === 'row' ? row : col;
    overlay.delete(key);
    if (insert) {
      if (coord >= at) {
        if (axis === 'row') row += count;
        else col += count;
      }
    } else {
      if (coord >= at && coord < at + count) continue;
      if (coord >= at + count) {
        if (axis === 'row') row -= count;
        else col -= count;
      }
    }
    shifted.push([gridCellKey(sheetIdx, row, col), value]);
  }
  for (const [key, value] of shifted) overlay.set(key, value);
}
```

- [ ] **Step 4: 테스트 실행 + 타입체크**

Run: `npx vitest run tests/lazyGridStore.test.ts && npx tsc --noEmit`
Expected: PASS

- [ ] **Step 5: 커밋**

```bash
git add src/xlsx/lazyGridStore.ts tests/lazyGridStore.test.ts
git commit -m "feat(xlsx): add LazyGridStore write path (setDirtyCell) and dirty-key shifting for structural ops"
```

---

### Task 6: 구조연산(행/열 삽입·삭제) + 직접 rowCount/colCount 확장

**Files:**
- Modify: `src/xlsx/lazyGridStore.ts`
- Test: `tests/lazyGridStore.test.ts`

**Interfaces:**
- Produces:
  - `function applyRowShift(store: LazyGridStore, sheetIdx: number, at: number, count: number, insert: boolean): void`
  - `function applyColShift(store: LazyGridStore, sheetIdx: number, at: number, count: number, insert: boolean): void`
  - `function growSheet(store: LazyGridStore, sheetIdx: number, newRowCount: number, newColCount: number): void`
  — Task 11/12(room.ts의 `applyGridStructureOp`/`applyStructureChangeToCanonical`)가 이 세
  함수를 부른다.

- [ ] **Step 1: 실패하는 테스트**

```ts
// tests/lazyGridStore.test.ts 에 추가
import { applyColShift, applyRowShift, growSheet } from '../src/xlsx/lazyGridStore.js';

describe('xlsx/lazyGridStore.ts: applyRowShift/applyColShift/growSheet', () => {
  test('inserting a row shifts logicalToOriginalRow and dirtyOverlay together, keeping getCell correct', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    setDirtyCell(store, 0, 1, 0, { v: 'Alice Kim' }); // 논리 1행(원래 Alice)을 편집

    applyRowShift(store, 0, 1, 1, true); // 1번 위치에 행 1개 삽입 — Alice가 2번째로 밀림

    expect(store.logicalToOriginalRow[0]).toEqual([0, -1, 1]);
    expect(await getCell(store, 0, 1, 0)).toBe(0); // 새로 삽입된 빈 행
    expect(await getCell(store, 0, 2, 0)).toEqual({ v: 'Alice Kim' }); // 편집 내용이 함께 밀림

    await fs.rm(store.rowStorePath, { force: true });
  });

  test('deleting a row removes the mapping and any dirty overlay entry inside the deleted range', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    applyRowShift(store, 0, 0, 1, false); // 0번째 행(Name/Score 헤더) 삭제

    expect(store.logicalToOriginalRow[0]).toEqual([1]);
    expect(await getCell(store, 0, 0, 0)).toEqual({ v: 'Alice' });

    await fs.rm(store.rowStorePath, { force: true });
  });

  test('applyColShift mirrors applyRowShift for the column axis', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    applyColShift(store, 0, 0, 1, true); // 맨 앞에 열 1개 삽입 — Name/Alice가 오른쪽으로 밀림

    expect(store.logicalToOriginalCol[0]).toEqual([-1, 0, 1]);
    expect(await getCell(store, 0, 0, 0)).toBe(0);
    expect(await getCell(store, 0, 0, 1)).toEqual({ v: 'Name', s: expect.any(Number) });

    await fs.rm(store.rowStorePath, { force: true });
  });

  test('growSheet extends logicalToOriginalRow/Col with -1 entries without touching existing mappings', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    growSheet(store, 0, 5, 4); // 2행×2열 시트를 5행×4열로 직접 확장(삽입 연산 없이)

    expect(store.logicalToOriginalRow[0]).toEqual([0, 1, -1, -1, -1]);
    expect(store.logicalToOriginalCol[0]).toEqual([0, 1, -1, -1]);
    expect(await getCell(store, 0, 4, 3)).toBe(0); // 새로 확장된 영역

    await fs.rm(store.rowStorePath, { force: true });
  });
});
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `npx vitest run tests/lazyGridStore.test.ts`
Expected: FAIL — 함수 없음

- [ ] **Step 3: 구현**

```ts
// src/xlsx/lazyGridStore.ts 에 추가

export function applyRowShift(
  store: LazyGridStore,
  sheetIdx: number,
  at: number,
  count: number,
  insert: boolean,
): void {
  const mapping = store.logicalToOriginalRow[sheetIdx];
  if (insert) {
    mapping.splice(at, 0, ...new Array(count).fill(-1));
  } else {
    mapping.splice(at, count);
  }
  shiftDirtyOverlayKeys(store.dirtyOverlay, sheetIdx, 'row', at, count, insert);
  const meta = store.sheetMeta[sheetIdx];
  meta.rowCount = insert ? meta.rowCount + count : Math.max(1, meta.rowCount - count);
}

export function applyColShift(
  store: LazyGridStore,
  sheetIdx: number,
  at: number,
  count: number,
  insert: boolean,
): void {
  const mapping = store.logicalToOriginalCol[sheetIdx];
  if (insert) {
    mapping.splice(at, 0, ...new Array(count).fill(-1));
  } else {
    mapping.splice(at, count);
  }
  shiftDirtyOverlayKeys(store.dirtyOverlay, sheetIdx, 'col', at, count, insert);
  const meta = store.sheetMeta[sheetIdx];
  meta.colCount = insert ? meta.colCount + count : Math.max(1, meta.colCount - count);
}

/** 삽입·삭제 연산 없이 rowCount/colCount만 직접 늘어나는 경로(그리드 에디터 툴바의 "행/열 추가"
 * 버튼) 전용 — 줄어들 때는 이 함수를 쓰지 않는다(room.ts의 기존 규칙과 동일: 배열은 그대로 두고
 * sheetMeta의 count만 낮추면 getCell의 경계 체크가 자연히 뒤쪽을 가려낸다). */
export function growSheet(store: LazyGridStore, sheetIdx: number, newRowCount: number, newColCount: number): void {
  const meta = store.sheetMeta[sheetIdx];
  const rowMapping = store.logicalToOriginalRow[sheetIdx];
  while (rowMapping.length < newRowCount) rowMapping.push(-1);
  const colMapping = store.logicalToOriginalCol[sheetIdx];
  while (colMapping.length < newColCount) colMapping.push(-1);
  meta.rowCount = Math.max(meta.rowCount, newRowCount);
  meta.colCount = Math.max(meta.colCount, newColCount);
}
```

- [ ] **Step 4: 테스트 실행 + 타입체크**

Run: `npx vitest run tests/lazyGridStore.test.ts && npx tsc --noEmit`
Expected: PASS

- [ ] **Step 5: 커밋**

```bash
git add src/xlsx/lazyGridStore.ts tests/lazyGridStore.test.ts
git commit -m "feat(xlsx): add LazyGridStore row/col structural shift + direct grow support"
```

---

### Task 7: 정리(cleanup) — 임시 파일 삭제 + 서버 시작 시 orphan 스윕

**Files:**
- Modify: `src/xlsx/lazyGridStore.ts`
- Modify: `src/server.ts` (또는 `ensureDataDirs`가 정의된 파일 — 실제 위치는 `Grep -n "ensureDataDirs" src/*.ts`로 확인 후 그 파일에 스윕 호출 추가)
- Test: `tests/lazyGridStore.test.ts`

**Interfaces:**
- Produces:
  - `async function closeLazyGridStore(store: LazyGridStore): Promise<void>` — room.ts의
    `closeConn`/`forceCloseRoom` 경로(Task 11)가 부른다.
  - `async function sweepOrphanedLazyGridFiles(): Promise<void>` — 서버 시작 시 한 번 호출.

- [ ] **Step 1: 실패하는 테스트**

```ts
// tests/lazyGridStore.test.ts 에 추가
import { closeLazyGridStore, sweepOrphanedLazyGridFiles } from '../src/xlsx/lazyGridStore.js';

describe('xlsx/lazyGridStore.ts: cleanup', () => {
  test('closeLazyGridStore deletes the on-disk row store file', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    await closeLazyGridStore(store);
    await expect(fs.access(store.rowStorePath)).rejects.toThrow();
  });

  test('sweepOrphanedLazyGridFiles removes leftover docview-lazygrid-* files in the OS temp dir', async () => {
    const os = await import('node:os');
    const path = await import('node:path');
    const orphanPath = path.join(os.tmpdir(), `docview-lazygrid-orphan-${randomUUID()}.jsonl`);
    await fs.writeFile(orphanPath, 'leftover');

    await sweepOrphanedLazyGridFiles();

    await expect(fs.access(orphanPath)).rejects.toThrow();
  });

  test('sweepOrphanedLazyGridFiles does not touch unrelated temp files', async () => {
    const os = await import('node:os');
    const path = await import('node:path');
    const unrelatedPath = path.join(os.tmpdir(), `not-ours-${randomUUID()}.txt`);
    await fs.writeFile(unrelatedPath, 'leave me alone');

    await sweepOrphanedLazyGridFiles();

    await expect(fs.access(unrelatedPath)).resolves.toBeUndefined();
    await fs.rm(unrelatedPath, { force: true });
  });
});
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `npx vitest run tests/lazyGridStore.test.ts`
Expected: FAIL — 함수 없음

- [ ] **Step 3: 구현**

```ts
// src/xlsx/lazyGridStore.ts 에 추가
export async function closeLazyGridStore(store: LazyGridStore): Promise<void> {
  await fs.rm(store.rowStorePath, { force: true, maxRetries: 5, retryDelay: 100 }).catch(() => {
    /* 정리 실패는 치명적이지 않다 — OS 임시 폴더 청소 또는 다음 서버 시작 시 sweep에 맡긴다 */
  });
}

/** 서버 비정상 종료로 남은 이 프로젝트의 임시 행 저장소 파일을 정리한다 — 정상 종료라면 이미
 * closeLazyGridStore가 다 지웠을 것이므로, 이건 그 경우를 복구하는 안전망일 뿐이다. */
export async function sweepOrphanedLazyGridFiles(): Promise<void> {
  const dir = os.tmpdir();
  let entries: string[];
  try {
    entries = await fs.readdir(dir);
  } catch {
    return;
  }
  await Promise.all(
    entries
      .filter((name) => name.startsWith('docview-lazygrid-'))
      .map((name) => fs.rm(path.join(dir, name), { force: true }).catch(() => undefined)),
  );
}
```

- [ ] **Step 4: 테스트 실행 + 타입체크**

Run: `npx vitest run tests/lazyGridStore.test.ts && npx tsc --noEmit`
Expected: PASS

- [ ] **Step 5: 서버 시작 시점에 스윕을 연결한다**

`ensureDataDirs`(`src/storage.ts:251`)가 서버 시작 시 불리는 지점을 확인하고(`grep -rn
"ensureDataDirs" src/`), 그 호출부(보통 `src/server.ts`)에서 `sweepOrphanedLazyGridFiles()`도
함께 await 한다:

```ts
// 호출부(실제 위치는 grep 결과에 맞춰 조정)
import { sweepOrphanedLazyGridFiles } from './xlsx/lazyGridStore.js';
// ...
await ensureDataDirs();
await sweepOrphanedLazyGridFiles();
```

- [ ] **Step 6: 전체 테스트 + 타입체크**

Run: `npx tsc --noEmit && npx vitest run`
Expected: PASS

- [ ] **Step 7: 커밋**

```bash
git add src/xlsx/lazyGridStore.ts tests/lazyGridStore.test.ts src/server.ts
git commit -m "feat(xlsx): add LazyGridStore cleanup (per-room delete + startup orphan sweep)"
```

---

## Phase 1 완료 시점의 상태

`LazyGridStore`는 이제 완전히 독립적으로 동작하고 테스트된 모듈이다 — **아직 `room.ts`에
연결되지 않았다**(기존 xlsx 협업 편집 동작은 전혀 바뀌지 않았고, 회귀 위험이 없다). 여기까지
머지해도 안전하다.

---

## Phase 2 — 룸의 읽기 경로 전환

### Task 8: `seedGridPageFromXlsx`/`seedGridStylesFromXlsx`를 `LazyGridStore` 기반으로 전환

**Files:**
- Modify: `src/xlsx/collab.ts:256` (`seedGridPageFromXlsx`)
- Modify: `src/xlsx/collab.ts:320` (`seedGridStylesFromXlsx`)
- Modify: `src/collab/room.ts` — 이 두 함수를 부르는 모든 지점(`seedRoom`,
  `refreshSeededPagesAfterShift`, `reconcile*` 계열) — 호출부는 아직 `room.canonicalGrid:
  GridDocument`를 넘기고 있으므로, 이 태스크에서는 **함수 시그니처를 `LazyGridStore`로 바꾸되
  호출부는 Task 11까지 임시로 `store.canonicalGrid`(옛 필드, 이 태스크 시점엔 아직 존재)를
  감싸는 어댑터를 통해 호출**한다 — 상세는 아래 "구현 순서 메모" 참고.
- Test: `tests/roomWindowedSeeding.test.ts`(기존 파일에 회귀 테스트 추가),
  `tests/roomStructureSync.test.ts`/`tests/roomRowColShift.test.ts`(기존 스위트, 이 태스크
  이후에도 반드시 전부 통과해야 함 — Global Constraints 참고)

**구현 순서 메모(이 태스크를 시작하기 전에 실제로 재확인할 것)**: `seedGridPageFromXlsx`는
현재 `canonicalGrid: GridDocument`를 받아 `sheet.rows[r][c]`를 동기로 읽는다
(`src/xlsx/collab.ts:256-268` 근방, 이 계획 작성 시점 기준). 이 함수를 async 배치 읽기로
바꾸려면:

1. 시그니처를 `seedGridPageFromXlsx(ydoc: Y.Doc, store: LazyGridStore, sheetIdx: number, page: number): Promise<void>`로 바꾼다.
2. 함수 안에서 `sheet.rows[r][c]` 직접 인덱싱 대신, 페이지 범위(`pageStart`~`pageEnd`)의 각 행을
   `await getRow(store, sheetIdx, r)`로 가져와 그 배열을 인덱싱한다. 청크(`YIELD_EVERY_N_ROWS`)
   경계마다 있던 기존 양보 지점은 그대로 유지하되, 그 청크 안에서 `getRow` 호출들을
   `Promise.all`로 모아 한 번에 기다리면(이미 캐시가 데워진 행은 즉시 resolve, 새로 읽어야 하는
   행만 실제 디스크 I/O) 청크당 디스크 왕복이 병렬화된다.
3. `mergeCoveredCells(sheet.merges)` 호출은 `sheetMeta[sheetIdx].merges`로 바뀐다(둘 다 이미
   같은 `GridMerge[]` 모양).
4. `seedGridStylesFromXlsx`도 같은 방식으로 바꾸되, 이 함수는 **시트 전체**를 순회한다(윈도잉
   대상이 아님, 기존 주석 참고) — `for (원본 행 전체) getRow(...)`를 청크+양보로 돈다. 이건
   여전히 O(시트 크기) 디스크 읽기를 room 열 때 한 번 하게 된다는 뜻이지만(스펙의 알려진 특성 —
   서식은 어차피 한 번은 전량 읽어야 한다), 결과를 JS 힙에 오래 들고 있지 않고 Yjs로 흘려보낸
   뒤 버린다는 점이 메모리 이득의 핵심이다.

- [ ] **Step 1: `tests/roomWindowedSeeding.test.ts`에 실패하는 테스트를 먼저 추가** — 정확한
  테스트 코드는 이 태스크를 시작하는 구현자가 Task 3-6에서 확정된 `LazyGridStore`의 실제 export
  목록을 보고 작성한다. 최소한 다음을 검증해야 한다: (a) 특정 페이지를 구독하면 그 페이지 범위의
  셀 값이 Yjs `GRID_CELLS_MAP`에 정확히 채워지는지(기존 동작과 동일), (b) 캐시에 없던 행을
  구독해도(디스크 첫 접근) 정확한 값이 채워지는지, (c) 서식이 페이지 구독과 무관하게 항상 전량
  채워지는지(기존 동작 유지).

- [ ] **Step 2~5: TDD 사이클 반복** — 실패 확인 → 구현 → 통과 확인 → 커밋. 구현자는 Task 3-6의
  정확한 함수 시그니처(`getRow`, `getCell`)를 그대로 가져다 쓴다(이 계획 문서 Task 4의 "Produces"
  참고).

- [ ] **Step 6: 기존 구조연산 회귀 스위트 재실행(필수 — Global Constraints)**

Run: `npx vitest run tests/roomStructureSync.test.ts tests/roomRowColShift.test.ts tests/roomWindowedSeeding.test.ts`
Expected: 전부 PASS. 실패하면 이 태스크의 변경이 페이지 시딩의 기존 동작을 바꾼 것이므로,
`seedGridPageFromXlsx`/`seedGridStylesFromXlsx`의 값 채우기 로직이 원래와 정확히 같은 결과를
내는지 다시 확인한다.

- [ ] **Step 7: 커밋**

```bash
git add src/xlsx/collab.ts src/collab/room.ts tests/roomWindowedSeeding.test.ts
git commit -m "refactor(xlsx): seedGridPageFromXlsx/seedGridStylesFromXlsx read through LazyGridStore instead of a fully-resident GridDocument"
```

---

### Task 9: 수식 계산을 "캐시 미스 시 재시도" 방식으로 전환

**Files:**
- Modify: `src/xlsx/formulas.ts` (`rawCellText`, `computeGridCellDisplay`)
- Test: `tests/xlsxFormulaRecompute.test.ts`(기존 파일에 추가), `tests/xlsxFormulas.test.ts`(기존
  10개 테스트가 그대로 통과해야 함 — 순수 계산 로직은 안 바뀌므로 회귀 없어야 함)

**Interfaces:**
- Consumes: `getRow`(Task 4)
- Produces: `async function computeGridCellDisplayAsync(store: LazyGridStore, sheetIdx: number, row: number, col: number, evalGuard?: Set<string>): Promise<unknown>` — 기존 동기 `computeGridCellDisplay`는 **그대로 남겨둔다**(순수 GridDocument 기반 호출부가 남아 있는 동안은 유지, Task 11에서 최종 정리).

- [ ] **Step 1: 실패하는 테스트**

```ts
// tests/xlsxFormulaRecompute.test.ts 에 추가
import { computeGridCellDisplayAsync } from '../src/xlsx/formulas.js';
import { ingestXlsxLazily, setDirtyCell } from '../src/xlsx/lazyGridStore.js';

test('a formula referencing a row far outside any cached page is still evaluated correctly (forces at least one cache-miss retry)', async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = { formula: 'A1000' }; // 1000번째 행을 참조
  sheet.getCell('A1000').value = 42;
  const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);

  const store = await ingestXlsxLazily(randomUUID(), bytes);
  // A1의 원본 텍스트는 파싱 시 '=A1000'으로 저장돼 있어야 함(이미 검증된 기존 규칙) — 캐시는
  // 비어 있는 상태로 시작(강제로 미스가 나게 함).
  expect(store.readCache.size).toBe(0);

  const result = await computeGridCellDisplayAsync(store, 0, 0, 0);
  expect(result).toBe(42);

  await fs.rm(store.rowStorePath, { force: true });
});

test('IF() with a value-dependent branch resolves only the branch actually taken, without prefetching the other branch', async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 1; // 조건이 참이 되게
  sheet.getCell('B1').value = { formula: 'IF(A1>0, C1, D1)' };
  sheet.getCell('C1').value = 'yes';
  sheet.getCell('D1').value = 'no';
  const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);

  const store = await ingestXlsxLazily(randomUUID(), bytes);
  const result = await computeGridCellDisplayAsync(store, 0, 0, 1); // B1
  expect(result).toBe('yes');

  await fs.rm(store.rowStorePath, { force: true });
});
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `npx vitest run tests/xlsxFormulaRecompute.test.ts`
Expected: FAIL — `computeGridCellDisplayAsync` not exported

- [ ] **Step 3: 구현** — 기존 동기 `rawCellText`/`computeGridCellDisplay`(`xlsx/formulas.ts:16-85`)
  의 그대로 두고, `LazyGridStore` 기반의 async 래퍼를 옆에 추가한다. 핵심 아이디어: 동기 평가를
  **캐시로 미리 채워진 상태에서만** 실행하고, 평가 도중 캐시에 없는 좌표를 만나면 클로저 플래그에
  기록만 해두고 그 자리는 `''`(빈 문자열)로 넘긴 뒤, 평가가 끝나면 그 플래그를 보고 필요한 행을
  채워 처음부터 다시 평가한다(스펙 문서 "수식 계산" 절 참고).

```ts
// src/xlsx/formulas.ts 에 추가
import type { LazyGridStore } from './lazyGridStore.js';
import { getRow } from './lazyGridStore.js';

/** computeGridCellDisplay와 완전히 같은 평가 로직이지만, 데이터 소스가 GridDocument 동기
 * 인덱싱이 아니라 "이미 캐시에 채워져 있어야 하는" LazyGridStore 동기 조회다. 캐시에 없는
 * 좌표를 만나면 던지지 않고(fast-formula-parser가 동기 콜백 안의 예외를 재포장해버림 —
 * rangeTooLarge와 같은 이유) onMiss 콜백으로 알린다. */
function computeGridCellDisplaySync(
  store: LazyGridStore,
  sheetIdx: number,
  row: number,
  col: number,
  evalGuard: Set<string>,
  onMiss: (sheetIdx: number, row: number) => void,
): unknown {
  const rawCellTextSync = (r: number, c: number): string => {
    const originalCol = store.logicalToOriginalCol[sheetIdx]?.[c];
    const key = gridCellKey(sheetIdx, r, c);
    if (store.dirtyOverlay.has(key)) {
      const cell = store.dirtyOverlay.get(key);
      return cell === 0 || cell === undefined ? '' : cell.v;
    }
    if (originalCol === undefined || originalCol < 0) return '';
    const originalRow = store.logicalToOriginalRow[sheetIdx]?.[r];
    if (originalRow === undefined || originalRow < 0) return '';
    const cacheKeyStr = `${sheetIdx}:${originalRow}`;
    const cachedRow = store.readCache.get(cacheKeyStr);
    if (!cachedRow) {
      onMiss(sheetIdx, r);
      return '';
    }
    const cell = cachedRow[originalCol];
    return cell === 0 || cell === undefined ? '' : cell.v;
  };
  // 이하 로직은 기존 computeGridCellDisplay(xlsx/formulas.ts:29-85)와 동일하되,
  // rawCellText(grid, ...) 호출을 위 rawCellTextSync(...)로, 재귀 호출을
  // computeGridCellDisplaySync(store, ..., evalGuard, onMiss)로 바꾼 것 — 실제 구현 시
  // 기존 함수 본문을 복사해 이 두 곳만 치환한다(로직 자체는 바뀌지 않는다).
  // [실제 구현자가 xlsx/formulas.ts:29-85의 본문을 그대로 옮겨와 두 치환점만 바꿀 것]
  throw new Error('구현 시 기존 computeGridCellDisplay 본문을 이 함수 안으로 옮기고 두 치환점만 바꿀 것');
}

export async function computeGridCellDisplayAsync(
  store: LazyGridStore,
  sheetIdx: number,
  row: number,
  col: number,
  evalGuard: Set<string> = new Set(),
): Promise<unknown> {
  const MAX_RETRIES = 10_000; // 한 수식이 참조할 수 있는 서로 다른 행 수의 넉넉한 상한(안전망)
  for (let attempt = 0; attempt < MAX_RETRIES; attempt += 1) {
    let missed: { sheetIdx: number; row: number } | null = null;
    const result = computeGridCellDisplaySync(store, sheetIdx, row, col, new Set(evalGuard), (missSheet, missRow) => {
      if (!missed) missed = { sheetIdx: missSheet, row: missRow }; // 이번 시도의 첫 미스만 채우고 재시도
    });
    if (!missed) return result;
    await getRow(store, missed.sheetIdx, missed.row); // 캐시를 채운다
  }
  throw new Error(`수식 계산이 ${MAX_RETRIES}번 재시도 후에도 끝나지 않음 (문서 손상 의심)`);
}
```

**구현자 참고**: `computeGridCellDisplaySync`의 본문(주석 처리된 부분)은 반드시 기존
`computeGridCellDisplay`(`xlsx/formulas.ts:29-85`)의 실제 코드를 그대로 복사해 옮겨야 한다 —
이 계획 문서에 그 전체를 다시 옮겨 적지 않은 이유는 로직이 하나도 안 바뀌기 때문이다(치환점
두 곳: `rawCellText(grid, sheetIdx, row, col)` → `rawCellTextSync(row, col)`, 재귀 호출의
`computeGridCellDisplay(grid, sheetIdx, ref.row-1, ref.col-1, evalGuard)` →
`computeGridCellDisplaySync(store, sheetIdx, ref.row-1, ref.col-1, evalGuard, onMiss)`).

- [ ] **Step 4: 테스트 실행, 실패 시 위 치환이 정확한지 확인**

Run: `npx vitest run tests/xlsxFormulaRecompute.test.ts tests/xlsxFormulas.test.ts`
Expected: PASS (새 테스트 + 기존 10개 그대로)

- [ ] **Step 5: 전체 테스트 + 타입체크**

Run: `npx tsc --noEmit && npx vitest run`
Expected: PASS

- [ ] **Step 6: 커밋**

```bash
git add src/xlsx/formulas.ts tests/xlsxFormulaRecompute.test.ts
git commit -m "feat(xlsx): add computeGridCellDisplayAsync — cache-miss-triggers-retry formula evaluation over LazyGridStore"
```

---

### Task 10: `recomputeFormulasForPage`가 새 비동기 평가를 쓰도록 전환

**Files:**
- Modify: `src/collab/room.ts` (`recomputeFormulasForPage`)
- Test: `tests/xlsxFormulaRecompute.test.ts`(room.ts 통합 시나리오 추가)

**Interfaces:**
- Consumes: `computeGridCellDisplayAsync`(Task 9)

- [ ] **Step 1: 실패하는 테스트** — 구독 중인 페이지 밖의 셀을 참조하는 수식이 편집 후 정확히
  재계산되는지, 기존 `tests/xlsxFormulaRecompute.test.ts`의 room.ts 통합 테스트 패턴을 따라
  작성한다(이미 이 파일에 "formula cells recompute into GRID_COMPUTED_MAP" 스위트가 있음 —
  그 패턴 그대로 새 케이스 추가).

- [ ] **Step 2~4: TDD 사이클** — `recomputeFormulasForPage`의 `computeGridCellDisplay(room.canonicalGrid!, ...)` 호출을 `await computeGridCellDisplayAsync(room.lazyGrid!, ...)`로 교체(단,
  `room.lazyGrid` 필드는 Task 11에서 정식으로 추가되므로, 이 태스크 시점에는 room.ts에 이미
  존재하는 임시 필드/어댑터를 통해 접근한다 — Task 8의 "구현 순서 메모"와 같은 원칙).

- [ ] **Step 5: 기존 회귀 스위트 재실행**

Run: `npx vitest run tests/xlsxFormulaRecompute.test.ts tests/roomStructureSync.test.ts tests/roomRowColShift.test.ts`
Expected: PASS

- [ ] **Step 6: 커밋**

```bash
git add src/collab/room.ts tests/xlsxFormulaRecompute.test.ts
git commit -m "refactor(collab): recomputeFormulasForPage uses computeGridCellDisplayAsync over LazyGridStore"
```

---

## Phase 3 — 구조연산 마이그레이션 (재설계됨, 2026-09-09)

**이 Phase는 Task 8 이후 실제 시도에서 발견된 진짜 아키텍처 문제로 재설계됐다 — 아래 "왜
재설계했는가" 절 먼저 읽을 것.** 원래 계획은 Task 11(구조연산)/12(직접 확장)/13(export 연동)/
14(진짜 배선)를 별도 태스크로 나눴는데, Task 11을 실제로 시도한 두 명의 독립된 구현자가
**같은 결론**에 도달했다: `applyGridStructureOp`가 `applyRowShift`/`applyColShift`(Task 6)를
부르려면 그 대상이 **진짜로 룸이 사는 동안 유지되는 하나의 `LazyGridStore` 인스턴스**여야
하는데, 지금은 Task 8이 만든 **임시 어댑터**(`lazyGridStoreAdapterForCanonicalGrid`,
`room.ts:278-299`)만 있다 — 이 어댑터는 매 호출마다 `room.canonicalGrid`(여전히 진짜 소스)에서
**새로 값-복사한 일회용 객체**를 만들어 넘긴다(`sheetMeta`도 `styles`도 매번 새 배열/객체). 구조
시프트가 이 어댑터의 `sheetMeta[i].rowCount`를 바꿔도 그 변경은 함수가 끝나는 즉시 버려지고
실제 `room.canonicalGrid`는 한 글자도 안 바뀐다 — **조용히 아무 일도 안 일어난 것처럼 보이면서
저장된 파일과 다른 참여자 화면은 실제로는 어긋난다.** 두 구현자 모두 이걸 직접 실험으로
재현했다(`tests/roomStructureSync.test.ts`+`tests/roomRowColShift.test.ts`가 34개 중 12개
실패)하고 정직하게 커밋 없이 되돌렸다 — **계획의 결함이지 구현자의 실수가 아니다.**

**근본 원인**: `Room`은 지금 "진짜 소스"를 **두 벌** 갖고 있다 — 항상 완전히 메모리에 상주하는
`canonicalGrid: GridDocument`(모든 실제 편집이 쌓이는 곳: `applyCellsChangeToCanonical`,
`applyStylesChangeToCanonical`)와, Task 3-7에서 만든 진짜 `LazyGridStore`(아직 아무 데도 안 쓰임,
룸에 상주하는 인스턴스가 없음). Task 8-10은 **읽기**만 하므로 어댑터로 눈속임이 가능했지만,
구조연산은 **쓰기**라 어댑터로는 안 된다. 셀 편집·서식 편집·구조연산·시트 추가·저장, 이 다섯 곳이
전부 **같은 하나의 살아있는 객체**를 봐야 서로 어긋나지 않는다 — 그래서 이걸 따로따로 마이그레이션할
수 없고(그러면 그 사이 기간에 두 소스가 어긋난다), **한 태스크로 묶어야 한다.** 이게 아래 새
Task 11이 원래 계획의 Task 11+12+13+14(진짜 배선 부분)를 모두 흡수한 이유다.

### Task 11: 진짜 상주 `Room.lazyGrid` 배선 + 모든 쓰기 경로(셀·서식 편집, 구조연산, 직접 확장,
시트 추가, 저장)를 한 번에 이전

**이 태스크는 크고 위험도가 높다(critical) — 이 세션에서 이미 두 번 실패를 겪은 바로 그 지점이다.
서두르지 말고, 각 단계마다 실제로 테스트를 돌려 확인하며 진행할 것.**

**Files:**
- Modify: `src/collab/room.ts` — `Room` 인터페이스, `seedRoom`, `applyCellsChangeToCanonical`,
  `applyStylesChangeToCanonical`, `ensureCanonicalCellSlot`(삭제 대상), `ensureCanonicalSheetSlot`,
  `applyStructureChangeToCanonical`, `applyGridStructureOp`, `flushToStore`, `getCurrentBytes`,
  `exportXlsxWithFallback`, `closeConn`, `forceCloseRoom`, 그리고 Task 8/10이 만든
  `lazyGridStoreAdapterForCanonicalGrid`/`GridDocumentBackedRowCache`(이 태스크가 끝나면 완전히
  불필요해진다 — 삭제)
- Modify: `src/xlsx/collab.ts` (`exportXlsxFromGrid`)
- Modify: `src/xlsx/exportStreaming.ts` (`exportXlsxFromGridStreaming`)
- Test: `tests/roomStructureSync.test.ts`, `tests/roomRowColShift.test.ts`(34개, 이 태스크의
  실질적 완료 기준), `tests/xlsx-collab.test.ts`, `tests/xlsxExportStreaming.test.ts`,
  `tests/roomExportFallback.test.ts`, `tests/roomWindowedSeeding.test.ts`,
  `tests/xlsxFormulaRecompute.test.ts`(전부 그대로 통과해야 함 — Task 8-10이 만든 어댑터 기반
  코드를 이 태스크가 실제 `lazyGrid`로 바꾸므로)

**시작하기 전에 반드시 실제 현재 코드를 읽을 것** — 아래는 방향과 이유를 설명하는 것이지 줄 단위
정답이 아니다(Task 8-10이 이미 이 파일들을 상당히 바꿔놓았다):

1. **`Room` 인터페이스에 `lazyGrid: LazyGridStore | null` 필드를 추가**한다. `canonicalGrid:
   GridDocument | null` 필드는 **이 태스크 안에서 완전히 제거**한다(아래 모든 단계가 끝나야
   안전하게 지울 수 있으므로, 실제로는 이 태스크의 마지막에 지운다 — 그 전까지는 남겨둬도
   무방하지만 더 이상 아무도 읽거나 써서는 안 된다).

2. **`seedRoom`을 `ingestXlsxLazily(docId, original)`(Task 3)를 부르도록 바꾼다** —
   `parseXlsxToGrid` 호출과 `room.canonicalGrid = grid` 대입을 제거하고 `room.lazyGrid = store`로
   교체한다. `seedGridStructureOnly`/`seedGridStylesFromXlsx`/`seedGridPageFromXlsx` 등 시딩
   경로가 지금 `GridDocument`나 어댑터를 받고 있다면 `LazyGridStore`(`room.lazyGrid`)를 직접
   받도록 시그니처를 정리한다(Task 8이 이미 `seedGridPageFromXlsx`/`seedGridStylesFromXlsx`
   자체는 `LazyGridStore` 시그니처로 바꿔뒀으므로, 여기서는 **어댑터로 감싸던 호출부만** 실제
   `room.lazyGrid`를 직접 넘기도록 고치면 된다).

3. **`applyCellsChangeToCanonical`을 `setDirtyCell`(Task 5) 기반으로 다시 쓴다.** 지금은
   Yjs `cells` 맵의 변경을 `canonicalGrid.sheets[i].rows[r][c]`에 직접 대입한다(빈 값이면
   `ensureCanonicalCellSlot`으로 배열을 늘려가며). 새 버전은 그 칸이 논리적으로 존재하는지
   (`row < lazyGrid.sheetMeta[sheetIdx].rowCount && col < ...colCount` — 상한 밖이면 지금처럼
   무시)만 확인하고, `setDirtyCell(room.lazyGrid, sheetIdx, row, col, value)`을 부른다. **삭제
   (`change.action === 'delete'`)는 `setDirtyCell(..., 0)`으로 명시적으로 빈 칸을 오버레이에
   써야 한다** — 그냥 `dirtyOverlay.delete(key)`를 하면 안 된다(그러면 "편집 안 된 칸"으로
   되돌아가 원본 값이 다시 보이는 버그가 된다 — dirtyOverlay는 "값이 없다"와 "이 칸이 지워졌다"를
   구분해야 하므로 `0`을 명시적으로 저장한다, `Room.dirtyCells`의 기존 주석에 있던 "왜 항상
   기억해야 하는가"와 같은 이유). `ensureCanonicalCellSlot`은 이제 필요 없다(이 함수 전체를
   호출부에서 지운다 — `Map`은 미리 크기를 맞출 필요가 없다는 스펙의 원래 설계 그대로).

4. **`applyStylesChangeToCanonical`도 같은 방식으로**: 서식 변경을 `setDirtyCell`로 반영하되,
   "값은 그대로 두고 서식만 지운다"는 기존 의미를 지키려면 **현재 그 칸의 값을 먼저 읽어와야
   한다**(`await getCell(room.lazyGrid, sheetIdx, row, col)`) — 이 함수가 `async`가 아니었다면
   `async`로 바꾸고 호출부(Yjs `styles.observe` 콜백)도 그에 맞게 조정한다(Yjs observe 콜백은
   동기여야 하므로, `void applyStylesChangeToCanonical(...)`처럼 흘려보내거나, 이미 Task 10이
   같은 문제를 해결한 패턴 — 청크 분류 후 `Promise.all`로 비동기 읽기를 먼저 끝내고 마지막에
   동기 `ydoc.transact`로 쓰기 — 을 참고해 room.ts 안에서 일관된 방식을 쓴다).

5. **`ensureCanonicalSheetSlot`(시트 추가)을 `lazyGrid.sheetMeta`에 새 항목을 push하도록**
   바꾼다(빈 시트 메타: rowCount/colCount 1, 빈 병합/열너비, freeze 0 — Task 3의
   `LazyGridSheetMeta` 모양 그대로). 이 새 시트는 원본에 없던 시트이므로
   `lazyGrid.logicalToOriginalRow/Col`에도 그 시트 자리에 **빈 매핑 배열**([])을 추가해야
   한다(길이가 늘면서 `growSheet`로 필요한 만큼 `-1`을 채워 넣는다) — 원본이 없는 시트는 모든
   좌표가 항상 "새 칸"(-1 매핑)이어야 `getCell`이 디스크를 보려다 실패하지 않는다.

6. **`applyGridStructureOp`(구조연산 메시지 처리)를 `applyRowShift`/`applyColShift`(Task 6)로
   전환** — 지금 `insertCanonicalRows`/`deleteCanonicalRows`/`insertCanonicalCols`/
   `deleteCanonicalCols`(`gridShift.ts`)를 부르는 지점을 `applyRowShift(room.lazyGrid, op.sheet,
   op.at, op.count, insert)` / `applyColShift(...)`로 1:1 교체한다. **`room.structuralOps.push(...)`
   호출(export 재생용 로그, 이번 세션 앞부분에서 이미 만들어짐)은 그대로 둔다** — `LazyGridStore`와
   무관한, export가 원본 워크북에 구조연산을 재생할 때 쓰는 별도 기록이라 건드릴 필요 없다. 두
   기록(실시간 읽기용 `logicalToOriginalRow/Col`과 export 재생용 `room.structuralOps`)은 같은
   이벤트에서 함께 갱신되기만 하면 되고, 하나로 합칠 필요는 없다. `shiftMerges`/`shiftGridStyleKeys`
   호출은 원래 계획대로 그대로 둔다(이 태스크 대상이 아님).

7. **직접 rowCount/colCount 확장 경로(`applyStructureChangeToCanonical`)를 `growSheet`(Task 6)로
   전환** — `sheet.rowCount = rowCount`로 직접 늘리던 지점을 `growSheet(room.lazyGrid, sheetIdx,
   rowCount, colCount)` 호출로 바꾼다. 줄어드는 경우는 `growSheet`를 부르지 않고
   `lazyGrid.sheetMeta[sheetIdx].rowCount/colCount`만 낮춘다(매핑 배열은 그대로 — Task 6의
   `growSheet` 주석과 동일한 원칙).

8. **`exportXlsxFromGrid`/`exportXlsxFromGridStreaming`의 인자를 정리**: 지금
   `(originalXlsxBytes, canonicalGrid, baselineSheetExtents, dirtyCells, structuralOps, options)`
   / `(originalXlsxBytes, canonicalGrid, dirtyCells, structuralOps, options)`인데, `canonicalGrid`
   대신 `room.lazyGrid`를 읽도록 바꾼다. 구체적으로:
   - 시트별 rowCount/colCount는 `lazyGrid.sheetMeta[i].rowCount/colCount`에서 읽는다 —
     `baselineSheetExtents` 인자는 **완전히 제거**한다(더 이상 필요 없다: "baseline과 비교해
     줄었는지" 판정 자체가, `sheetMeta`가 이미 "지금 논리적으로 몇 행/열인지"를 정확히 담고 있어
     불필요해진다).
   - 값 diff 루프는 `lazyGrid.dirtyOverlay`를 직접 순회한다(키별로 좌표를 파싱해 원본 워크북의
     해당 셀에 오버레이 값을 쓴다) — `dirtyCells`(옛 `Set<string>`, `room.dirtyCells`) 인자도
     이제 필요 없다(오버레이의 키 집합이 이미 그 정보다, 값까지 포함해서).
   - `structuralOps` 인자는 **그대로 유지**한다(`room.structuralOps`, 위 6번에서 설명한 대로
     export 재생 전용 기록이라 이 리팩터와 무관하게 계속 필요하다).
   - `room.ts`의 `exportXlsxWithFallback`/`flushToStore`/`getCurrentBytes` 호출부를 새 시그니처에
     맞게 고친다.

9. **정리**: `lazyGridStoreAdapterForCanonicalGrid`/`GridDocumentBackedRowCache`/
   `identityIndexArray`(room.ts:239-299, Task 8이 만든 임시 어댑터 전체)를 삭제하고, 그걸 쓰던
   4곳(`grep -rn "lazyGridStoreAdapterForCanonicalGrid" src/collab/room.ts`로 확인)을 전부
   `room.lazyGrid`를 직접 넘기도록 바꾼다. `Room.canonicalGrid` 필드 자체도 이제 삭제한다.

10. **`closeConn`/`forceCloseRoom`에 정리 호출 추가**:

```ts
// room.ts의 closeConn, 룸이 완전히 닫히는 분기(그리고 forceCloseRoom 경로)에 추가
if (room.lazyGrid) void closeLazyGridStore(room.lazyGrid);
```

- [x] **검증 순서(각 단계마다, 특히 3·4·6·8을 마칠 때마다 아래를 돌려 회귀를 조기에 잡을 것)**

Run: `npx vitest run tests/roomStructureSync.test.ts tests/roomRowColShift.test.ts`
Expected: PASS (34개 전부) — 이게 이 태스크의 진짜 완료 기준이다. 이전 두 번의 시도가 정확히
이 스위트에서 12/34 실패로 문제를 발견했으니, 여기서 전부 통과해야만 어댑터 문제가 실제로
해결된 것이다.

결과: 34/34 PASS. 통과 과정에서 실제 버그 둘을 발견해 함께 고쳤다 — (1) `applyStructureChangeToCanonical`
호출을 `enqueueCanonicalGrowth`로 감쌌더니 시트 이름변경·추가·삭제·순서변경 4개 테스트가 깨져
직접 호출로 되돌림(주석도 실제 동작에 맞게 수정), (2) `tests/roomRowColShift.test.ts`의 "malformed
or out-of-bounds" 테스트가 `parseXlsxToGrid`(actualRowCount, "값 있는 행 개수")와
`ingestXlsxLazily`(진짜 최대 행 번호)의 rowCount 계산 차이 때문에 우연히 통과하던 낡은 경계값을
쓰고 있어 실제 크기(1501)에 맞게 교정.

- [x] **전체 회귀**

Run: `npx tsc --noEmit && npx vitest run`
Expected: 클린, 전부 PASS(특히 위에 나열한 6개 테스트 파일)

결과: `npx tsc --noEmit` 클린. `npx vitest run` 41 files / 301 tests 전부 PASS. (테스트 파일
목록이 처음 계획과 달라진 이유: `tests/roomExportFallback.test.ts`는 이미 LazyGridStore
시그니처로 작성돼 있었고, 이번에 함께 손볼 필요가 있었던 건 `tests/roomCanonicalGrid.test.ts`·
`tests/roomCanonicalGrowthYielding.test.ts`·`tests/xlsx-collab.test.ts`·
`tests/xlsxExportStreaming.test.ts`·`tests/xlsxExportYielding.test.ts`였다 — 전부 옮김.)

- [ ] **실제 대용량 파일로 재검증** — 이전 세션에서 스트리밍 저장 폴백을 검증했던 것과 같은 방식
  (임시 scratch 스크립트, 확인 후 삭제)으로 292,244행 파일을 열람 → 페이지 스크롤 → 셀 편집 →
  행 삽입 → 저장 → 재오픈까지 end-to-end 확인한다. 이 파일 자체가 이번 세션 전체 조사의 발단이었던
  실제 사용자 파일이다. **Task 13(최종 회귀 검증 + 실제 대용량 파일 end-to-end)으로 이월** — Task 13이
  같은 검증을 다시 요구하므로 중복 작업을 피하기 위해 거기서 함께 한다.

- [x] **커밋** — 이 태스크는 하나의 큰 커밋이어도 되고(여러 파일이 동시에 일관돼야 서로 안
  깨지므로), 만약 3~10을 각각 독립적으로 커밋할 수 있는 지점이 실제로 있다면(예: 3·4를 끝내고
  구조연산은 아직 어댑터를 쓰는 중간 상태가 실제로 테스트를 통과한다면) 더 잘게 나눠도 좋다 —
  단, **중간 커밋이 기존 34개 구조연산 테스트를 깨는 상태로 남아서는 절대 안 된다.**

커밋: `8d5c9a5`.

```bash
git add -A
git commit -m "refactor(collab): wire a real persistent Room.lazyGrid; migrate all canonical write paths (cell/style edits, structural ops, direct growth, sheet add, export) to it; remove the Task 8 adapter"
```

---

### Task 12: 죽은 코드 정리

**Files:**
- Modify: `src/xlsx/gridShift.ts` — `insertCanonicalRows`/`deleteCanonicalRows`/
  `insertCanonicalCols`/`deleteCanonicalCols`가 더 이상 어디서도 안 불리면 파일째(또는 함수만)
  삭제(`grep -rn "insertCanonicalRows\|deleteCanonicalRows\|insertCanonicalCols\|
  deleteCanonicalCols" src/`로 확인 — `shiftMerges`는 Task 11에서도 계속 쓰이므로 남긴다)
- Modify: `src/collab/room.ts` — `growCanonicalRows`/`widenCanonicalRows`(옛 배열 패딩 함수,
  Task 11 이후 안 쓰이면 삭제), `Room.dirtyCells: Set<string>`과
  `Room.baselineSheetExtents`(둘 다 Task 11에서 `lazyGrid.dirtyOverlay`/`lazyGrid.sheetMeta`로
  대체돼 죽은 필드가 됐을 것 — `grep -rn "\.dirtyCells\|\.baselineSheetExtents" src/`로 확인 후
  필드 선언과 초기화 지점(`getOrCreateRoom`)까지 함께 삭제), `shiftDirtyCellKeys`(room.ts의 옛
  `Set<string>` 전용 버전 — `lazyGridStore.ts`의 `shiftDirtyOverlayKeys`로 완전히 대체됐다면 삭제)

**이 태스크를 시작하기 전에 Task 11이 실제로 완료·커밋돼 있는지 먼저 확인**(`git log`로 Task 11의
커밋 존재 확인) — Task 11 없이 이 태스크만 진행하면 아직 쓰이고 있는 코드를 지우게 된다.

- [ ] **Step 1**: 위 각 후보에 대해 `grep -rn`으로 실제 호출부가 남아있지 않은지 확인 후 삭제.
  하나씩 지우고 그때마다 타입체크를 돌려, 어딘가 여전히 참조하고 있었다면(즉 Task 11이 그
  참조를 놓쳤다면) 이 태스크에서 바로 잡을지 Task 11로 돌아가 고칠지 판단한다(단순 참조 누락이면
  이 자리에서 마저 고치는 게 낫다).

Run: `npx tsc --noEmit` (삭제할 때마다)
Expected: 클린 — 에러가 나면 그 참조부를 실제 대체 함수로 먼저 고친 뒤 다시 삭제 시도

- [ ] **Step 2: 전체 테스트**

Run: `npx vitest run`
Expected: 전부 PASS, 특히 `tests/roomStructureSync.test.ts`/`tests/roomRowColShift.test.ts`(34개)

- [ ] **커밋**

```bash
git add -A
git commit -m "refactor: delete dead code superseded by LazyGridStore (old array-splice structural ops, Room.dirtyCells/baselineSheetExtents, old Set-based shiftDirtyCellKeys)"
```

---

### Task 13: 최종 회귀 검증 + 실제 대용량 파일 end-to-end

**Files:** 없음(검증 전용)

- [ ] **Step 1: 전체 테스트 스위트**

Run: `npx vitest run`
Expected: 전부 PASS, 특히 `tests/roomStructureSync.test.ts`/`tests/roomRowColShift.test.ts`(34개)

- [ ] **Step 2: 타입체크**

Run: `npx tsc --noEmit`
Expected: 클린

- [ ] **Step 3: 실제 292,244행 파일로 열람 → 페이지 스크롤 → 셀 편집 → 구독 밖 셀을 참조하는
  수식 입력 → 저장 → 재오픈까지 실제 서버(`npm run dev`)로 수동 확인, 그리고 이 과정 동안
  서버 프로세스 메모리(`process.memoryUsage()` 로그 또는 OS 작업 관리자)가 이전 세션에 측정한
  535MB보다 뚜렷이 낮게 유지되는지 확인(정확한 목표 수치는 없지만 "파일 크기에 비례하지 않고
  안정적으로 유지"가 확인 기준).

- [ ] **Step 4: `docs/TODO.md` 갱신** — "남은 항목"의 지연 materialization 항목을 완료로 표시.

```bash
git add docs/TODO.md
git commit -m "docs: mark lazy canonical grid materialization as complete in TODO.md"
```
