import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { Readable } from 'node:stream';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
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
import { gridCellKey, parseGridCellKey } from '../collab/gridShared.js';

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
  /** readCache 상한(칸이 아니라 행 개수) — Map은 삽입 순서를 유지하므로 LRU 축출에 그대로 재사용한다. */
  readCacheLimit: number;
}

const DEFAULT_READ_CACHE_ROWS = 5000;

function identityArray(length: number): number[] {
  return Array.from({ length }, (_, i) => i);
}

/**
 * 워크북 XML의 sheetId -> 물리 워크시트 경로("worksheets/sheetN.xml") 매핑.
 * ExcelJS.stream.xlsx.WorkbookReader는 mergeCell/sheetView(틀 고정) XML을 전혀 해석하지 않는다
 * (node_modules/exceljs/lib/stream/xlsx/worksheet-reader.js 확인 완료 — `case 'mergeCell': break`
 * 처럼 그냥 버림, `.model`/`.views` 게터 자체가 없음). 그래서 merges/freezeRows/freezeCols를
 * 스트리밍 경로에서도 잃지 않으려면 jszip으로 각 시트의 원본 XML을 별도로 읽어야 한다. 이때 어떤
 * WorksheetReader가 어떤 sheetN.xml에 대응하는지는 "출현 순서" 같은 추측에 의존하지 않고, 워크북
 * 리더가 채워두는 `model.sheets`(sheetId/rId)와 `workbookRels`(rId -> Target)를 그대로 따라가
 * 확정한다(worksheet-reader.js의 `_parseWorksheet`가 동일한 방식으로 이름을 확정하는 것과 같은
 * 원리). 아래 preSeedWorkbookReaderState의 이유로 이 두 필드는 이제 내부 엔진이 아니라 우리가
 * 직접, 결정적으로 채운다.
 */
interface RawWorkbookReaderInternals {
  model?: { sheets?: Array<{ id: number; name: string; state?: string; rId: string }> };
  workbookRels?: Array<{ Id: string; Target: string }>;
  /**
   * 아래 세 메서드는 ExcelJS.stream.xlsx.WorkbookReader의 private(밑줄 접두사) 메서드다 — 공개
   * API가 아니지만, preSeedWorkbookReaderState가 왜 이 메서드들을 직접 호출하는지는 그 함수의
   * 주석에 설명했다. exceljs 버전이 바뀌면(현재 package.json은 ^4.4.0 — 4.4.0이 최신) 이 메서드들의
   * 존재/시그니처가 바뀔 수 있으니, exceljs 업그레이드 시 반드시 재검증할 것.
   */
  _parseRels(entry: NodeJS.ReadableStream): Promise<void>;
  _parseWorkbook(entry: NodeJS.ReadableStream): Promise<void>;
  _parseStyles(entry: NodeJS.ReadableStream): Promise<void>;
  _parseSharedStrings(entry: NodeJS.ReadableStream): AsyncIterable<unknown>;
}

async function loadWorksheetXmlsByPath(zip: JSZip): Promise<Map<string, string>> {
  const reads: Array<Promise<void>> = [];
  const byPath = new Map<string, string>();
  zip.forEach((relPath, entry) => {
    if (/^xl\/worksheets\/sheet\d+\.xml$/.test(relPath)) {
      reads.push(entry.async('text').then((text) => { byPath.set(relPath, text); }));
    }
  });
  await Promise.all(reads);
  return byPath;
}

/**
 * 실험으로 확인된 exceljs(4.4.0, 이 프로젝트에 고정된 버전) 스트리밍 리더의 심각한 레이스 컨디션을
 * 우회한다 — 반드시 ingestXlsxLazily에서 워크북 리더를 순회하기 *전에* 호출해야 한다.
 *
 * 무슨 문제인가: `ExcelJS.stream.xlsx.WorkbookReader`는 xl/workbook.xml, xl/_rels/workbook.xml.rels,
 * xl/styles.xml, xl/sharedStrings.xml을 "내부 zip 스트림을 훑다가 마주치는 대로" 비동기적으로
 * 파싱해 `this.model`/`this.workbookRels`/`this.styles`/`this.sharedStrings`에 채워 넣는다. 이
 * 프로젝트에서 실제로 재현·격리한 버그: 이 채워 넣기가 각 워크시트 파싱 시작 시점보다 늦게(또는
 * 아예 실행되지 않고) 끝나는 경우가 실제로 존재한다 — 순수히 파일이 작을 때(약 11KB 미만 zip에서
 * 100% 재현, 그 이상에서는 관찰 못 함 — 그러나 타이밍 레이스라 크기와 무관하게 이론상 발생 가능)
 * 자주 벌어진다. 결과: (a) 워크시트 메타(`model.sheets`) 조회 시 `Cannot read properties of
 * undefined (reading 'sheets')`로 즉시 죽거나, (b) 더 위험하게, sharedStrings/styles가 비어 있는
 * 채로 조용히 넘어가 문자열 칸이 `{sharedString: idx}` placeholder로, 서식이 빈 채로 남을 수
 * 있다(오탐 없이 통과하는 silent corruption — 이 레이스가 스타일에도 똑같이 적용된다는 것도 별도
 * 스크립트로 확인함).
 *
 * 왜 이렇게 고치나: 이 네 조각(rels/workbook/styles/sharedStrings)은 파일 크기와 무관하게 항상
 * 작고, docId당 한 번만 필요하며, 이미 열어 둔 jszip 인스턴스에서 그대로 꺼낼 수 있다. 이 함수는
 * 그 원본 XML을 워크북 리더의 **바로 그 private 파싱 메서드**(`_parseRels`/`_parseWorkbook`/
 * `_parseStyles`/`_parseSharedStrings`)에 직접, 순서대로, 완전히 await해 넘긴다 — 즉 파싱 로직
 * 자체는 한 글자도 새로 만들지 않고 exceljs 코드를 그대로 재사용하며, 다만 "내부 zip 스트림이
 * 알아서 순서대로 넘겨줄 때"가 아니라 "우리가 결정적으로 먼저 채워 넣을 때" 실행되게 만들 뿐이다.
 * 이후 메인 순회 중 내부 엔진이 (레이스에서 이겨) 같은 필드를 다시 덮어써도 완전히 같은 내용이라
 * 무해하다. 80회 연속 실행 + 다양한 파일 크기(2~292,244행 규모까지)로 이 우회가 레이스를 없앤다는
 * 것을 별도 스크립트로 검증했다(스크립트는 이 프로젝트 관례대로 확인 후 삭제함).
 */
async function preSeedWorkbookReaderState(
  reader: ExcelJS.stream.xlsx.WorkbookReader,
  zip: JSZip,
): Promise<void> {
  const rawReader = reader as unknown as RawWorkbookReaderInternals;

  const relsEntry = zip.file('xl/_rels/workbook.xml.rels');
  if (relsEntry) {
    await rawReader._parseRels(Readable.from(await relsEntry.async('nodebuffer')));
  }
  const workbookEntry = zip.file('xl/workbook.xml');
  if (workbookEntry) {
    await rawReader._parseWorkbook(Readable.from(await workbookEntry.async('nodebuffer')));
  }
  const stylesEntry = zip.file('xl/styles.xml');
  if (stylesEntry) {
    await rawReader._parseStyles(Readable.from(await stylesEntry.async('nodebuffer')));
  }
  const sharedStringsEntry = zip.file('xl/sharedStrings.xml');
  if (sharedStringsEntry) {
    // 'cache' 모드에서는 아무것도 yield하지 않는 async generator다 — 끝까지 돌려서(for await) 부수
    // 효과(this.sharedStrings 채우기)만 얻으면 된다.
    for await (const _unused of rawReader._parseSharedStrings(
      Readable.from(await sharedStringsEntry.async('nodebuffer')),
    )) {
      // 'cache' 모드에서는 절대 오지 않는 분기 — 안전망으로만 존재.
      void _unused;
    }
  }
}

function extractMergeRefsFromSheetXml(xml: string): string[] {
  const refs: string[] = [];
  for (const m of xml.matchAll(/<mergeCell\s+ref="([^"]+)"/g)) refs.push(m[1]);
  return refs;
}

/** OOXML의 pane state는 "frozen"(완전 고정)과 "frozenSplit"(고정+분할) 둘 다 "틀 고정"으로 친다 —
 * ExcelJS의 sheet-view-xform.js VIEW_STATES 매핑과 동일한 규칙. */
function extractFrozenSplitFromSheetXml(xml: string): { xSplit: number; ySplit: number } | undefined {
  const paneTag = xml.match(/<pane\b[^>]*\/?>/)?.[0];
  if (!paneTag) return undefined;
  const state = paneTag.match(/state="([^"]*)"/)?.[1];
  if (state !== 'frozen' && state !== 'frozenSplit') return undefined;
  const xSplit = Number(paneTag.match(/xSplit="([^"]*)"/)?.[1] ?? '0');
  const ySplit = Number(paneTag.match(/ySplit="([^"]*)"/)?.[1] ?? '0');
  return { xSplit, ySplit };
}

interface RawColSpec {
  min: number;
  max: number;
  width: number | undefined;
}

/**
 * 스트리밍 WorksheetReader의 `.columns` 게터는 이 프로젝트에서 고정한 exceljs@4.4.0에서 항상 빈
 * 배열을 반환한다 — worksheet-reader.js가 `this._columns = Column.fromModel(cols)`를 인자 1개로
 * 호출하는데(267행), `Column.fromModel(worksheet, cols)`(doc/column.js:295)는 인자 2개를 기대해서
 * 실제 `cols` 배열이 `worksheet` 자리로 밀려 들어가고 진짜 `cols`는 사라진다(`cols = cols || []`가
 * `undefined || []`로 항상 `[]`가 됨). 그래서 컬럼 너비는 `worksheetReader.columns`를 믿지 않고,
 * merges/freeze pane과 같은 방식으로 원본 시트 XML의 `<col min max width>`를 직접 파싱해서 얻는다.
 */
function extractColSpecsFromSheetXml(xml: string): RawColSpec[] {
  const specs: RawColSpec[] = [];
  for (const m of xml.matchAll(/<col\s+([^>]*?)\/?>/g)) {
    const attrs = m[1];
    const min = Number(attrs.match(/\bmin="([^"]*)"/)?.[1] ?? '0');
    const max = Number(attrs.match(/\bmax="([^"]*)"/)?.[1] ?? '0');
    if (!(min > 0) || !(max > 0)) continue;
    const widthStr = attrs.match(/\bwidth="([^"]*)"/)?.[1];
    specs.push({ min, max, width: widthStr === undefined ? undefined : Number(widthStr) });
  }
  return specs;
}

/**
 * exceljs의 `Column.fromModel(worksheet, cols)`(doc/column.js:295-312)와 동일한 순서로 min 기준
 * 정렬된 col 스펙을 순회하며 컬럼별 너비를 배정한다 — 겹치지 않는 정상 XML에서는 결과가 같고,
 * 다만 `count`가 `colCount`를 넘어서면 더 볼 필요가 없으니 그 지점에서 순회를 멈춘다(원본은
 * col.max까지, 예: 16384까지 무조건 순회함).
 */
function computeColWidthsPx(colSpecs: readonly RawColSpec[], colCount: number): number[] {
  const sorted = [...colSpecs].sort((a, b) => a.min - b.min);
  const widthAt: Array<number | undefined> = new Array(colCount + 1).fill(undefined); // 1-based
  let count = 1;
  for (const spec of sorted) {
    if (count > colCount) break;
    while (count < spec.min && count <= colCount) count += 1;
    while (count <= spec.max && count <= colCount) {
      widthAt[count] = spec.width;
      count += 1;
    }
  }
  const colWidths: number[] = [];
  for (let c = 1; c <= colCount; c += 1) {
    const width = widthAt[c];
    colWidths.push(width ? Math.round(width * 7.2) : 100);
  }
  return colWidths;
}

export async function ingestXlsxLazily(docId: string, xlsxBytes: Buffer): Promise<LazyGridStore> {
  const themeColors = await extractThemeColorsFromZip(xlsxBytes);
  const zip = await JSZip.loadAsync(xlsxBytes);
  const worksheetXmlsByPath = await loadWorksheetXmlsByPath(zip);
  const rowStorePath = path.join(os.tmpdir(), `docview-lazygrid-${docId}-${randomUUID()}.jsonl`);
  const writeHandle = await fs.open(rowStorePath, 'w');

  const sheetMeta: LazyGridSheetMeta[] = [];
  const rowIndex: Array<Array<LazyRowIndexEntry | undefined>> = [];
  const palette = new StylePalette();
  let byteOffset = 0;

  try {
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from(xlsxBytes), {
      worksheets: 'emit',
      sharedStrings: 'cache',
      hyperlinks: 'ignore',
      styles: 'cache',
      entries: 'ignore',
    });
    // 아래 preSeedWorkbookReaderState의 큰 주석 참고 — exceljs 4.4.0 스트리밍 리더 자체의 레이스
    // 컨디션을 우회하기 위해 반드시 순회 시작 전에 호출해야 한다.
    await preSeedWorkbookReaderState(reader, zip);
    const rawReader = reader as unknown as RawWorkbookReaderInternals;

    for await (const worksheetReader of reader) {
      const thisSheetIndex: Array<LazyRowIndexEntry | undefined> = [];
      rowIndex.push(thisSheetIndex);

      // exceljs의 @types 선언(index.d.ts)은 스트리밍 WorksheetReader에 실제로 있는 .name/.id/
      // .columns를 선언하지 않는다(선언된 columns()는 메서드로 돼 있는데 런타임은 getter 배열이다 —
      // 타입 선언과 런타임이 어긋나는 부분, worksheet-reader.js 소스로 확인 완료). 그래서 이 시트
      // 전용으로 한 번만 캐스팅해 재사용한다.
      const rawWorksheetReader = worksheetReader as unknown as {
        name: string;
        id?: number;
        dataValidations?: { find(address: string): unknown };
      };

      // 스트리밍 워크시트 리더는 데이터 유효성 검사(xl/worksheets/sheetN.xml의 <dataValidations>)를
      // 전혀 파싱하지 않는다(worksheet-reader.js에 해당 케이스 자체가 없음) — 그런데 Cell.prototype
      // .dataValidation 게터(doc/cell.js)는 `this.worksheet.dataValidations.find(...)`를 무조건
      // 호출해서, 이 프로퍼티가 없으면 (extractStyle이 읽자마자) 곧바로 던진다. sheetProtected와
      // 같은 이유로 알려진 스코프 제한으로 받아들이고 — 항상 "검사 없음"으로 답하는 최소 구현을
      // 붙여 크래시만 막는다(nu 스타일 비트가 스트리밍 경로에서는 항상 비어 있게 됨).
      rawWorksheetReader.dataValidations = { find: () => undefined };

      // 이 시트의 mergeCell/틀고정 정보가 담긴 원본 XML을 찾는다(위 설명대로 sheetId -> rId -> 물리
      // 경로 순으로 확정 — 출현 순서를 신뢰하지 않는다).
      const wsId = rawWorksheetReader.id;
      const sheetModelEntry = rawReader.model?.sheets?.find((s) => s.id === wsId);
      const rel = sheetModelEntry
        ? rawReader.workbookRels?.find((r) => r.Id === sheetModelEntry.rId)
        : undefined;
      const sheetXml = rel ? worksheetXmlsByPath.get(`xl/${rel.Target}`) : undefined;

      let maxRowWithValue = 0;
      let maxColWithValue = 0;
      let rowsProcessed = 0;
      // sheetProtection은 스트리밍 리더에서 신뢰성 있게 얻을 수 있는 API가 없어 false로 둔다
      // (알려진 스코프 제한 — 시트 보호 잠금 서식(lk)이 스트리밍 경로에서는 반영되지 않음).
      const sheetProtected = false;

      for await (const row of worksheetReader) {
        // ExcelJS 스트리밍 리더는 XML에 실제로 등장하는 <row> 엘리먼트만 순서대로 내보낸다 — 값이나
        // 서식이 전혀 없는 행은 XML 자체에 없을 수 있어(예: 1행과 300행만 있고 그 사이는 완전히
        // 비어 있는 시트) 내보내지는 행 개수만큼 단순히 0,1,2...로 세면 원본 행 번호와 어긋난다
        // (칸이 조용히 잘못된 위치에 저장되는 버그). 그래서 반드시 row.number(1-based, 실제 원본
        // 행 번호)를 그대로 써야 한다.
        const originalRow = row.number - 1;
        if (originalRow < 0) continue;
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
        if (rowHasValue) maxRowWithValue = row.number;

        const json = JSON.stringify(cells);
        const buf = Buffer.from(json, 'utf8');
        await writeHandle.write(buf, 0, buf.length, byteOffset);
        thisSheetIndex[originalRow] = { byteOffset, byteLength: buf.length };
        byteOffset += buf.length;

        rowsProcessed += 1;
        if (rowsProcessed % YIELD_EVERY_N_ROWS === 0) await yieldToEventLoop();
      }

      let rowCount = Math.min(maxRowWithValue, MAX_ROWS);
      let colCount = Math.min(maxColWithValue, MAX_COLS);

      // 병합된 칸 중 마스터(좌상단)가 아닌 칸은 실제 xlsx XML에 <c> 엘리먼트가 아예 없는 경우가
      // 흔하다(값은 마스터 칸에만 있고, 병합 자체는 <mergeCells>로만 기록됨) — 그래서 셀 스캔만으로는
      // 이런 칸이 있는 행/열을 못 본다. 반면 워크북을 통째로 읽는 parseXlsxToGrid는 ExcelJS가 로드
      // 시점에 병합 범위 전체에 Cell(type=Merge)을 실제로 채워 넣어(worksheet.js의 `model` setter가
      // 병합마다 mergeCellsWithoutStyle을 호출) actualColumnCount/actualRowCount에 그 칸까지
      // 잡히게 만든다(직접 실측·확인함 — 병합된 헤더처럼 마스터 칸 하나만 값이 있는 표에서
      // colCount가 어긋나는 걸 재현). 스트리밍 경로에서도 병합 범위만큼 rowCount/colCount를
      // 넓혀 이 값을 맞춘다.
      const rawMerges = sheetXml
        ? extractMergeRefsFromSheetXml(sheetXml).map(parseMergeRange).filter((m): m is GridMerge => m !== null)
        : [];
      for (const m of rawMerges) {
        rowCount = Math.max(rowCount, Math.min(m.r + m.rs - 1, MAX_ROWS));
        colCount = Math.max(colCount, Math.min(m.c + m.cs - 1, MAX_COLS));
      }
      // parseXlsxToGrid와 동일하게, 병합의 좌상단이 최종 rowCount/colCount 안에 있는 병합만 남긴다.
      const merges = rawMerges.filter((m) => m.r <= rowCount && m.c <= colCount);
      const frozen = sheetXml ? extractFrozenSplitFromSheetXml(sheetXml) : undefined;

      const colSpecs = sheetXml ? extractColSpecsFromSheetXml(sheetXml) : [];
      const colWidths = computeColWidthsPx(colSpecs, colCount);

      sheetMeta.push({
        name: rawWorksheetReader.name,
        rowCount,
        colCount,
        colWidths,
        merges,
        freezeRows: frozen?.ySplit ?? 0,
        freezeCols: frozen?.xSplit ?? 0,
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
    logicalToOriginalRow: sheetMeta.map((meta) => identityArray(meta.rowCount)),
    logicalToOriginalCol: sheetMeta.map((meta) => identityArray(meta.colCount)),
    dirtyOverlay: new Map(),
    readCache: new Map(),
    readCacheLimit: DEFAULT_READ_CACHE_ROWS,
  };
}

async function readRowFromDisk(
  store: LazyGridStore,
  sheetIdx: number,
  originalRow: number,
): Promise<ReadonlyArray<GridCell | 0>> {
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
