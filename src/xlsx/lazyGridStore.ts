import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import type { FileHandle } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { Readable } from 'node:stream';
import { setTimeout as timerSleep } from 'node:timers/promises';
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

/** openRowStoreWithRetry의 최대 재시도 횟수·간격 — closeLazyGridStore/exportStreaming.ts의
 * fs.rm(..., { maxRetries: 5, retryDelay: 100 })와 같은 값을 그대로 맞춘다(일관된 관용구). */
const ROW_STORE_OPEN_MAX_RETRIES = 5;
const ROW_STORE_OPEN_RETRY_DELAY_MS = 100;

/**
 * `fs.open(rowStorePath, 'r')`가 `ENOENT`로 실패하면 최대 5번, 100ms 간격으로 재시도한다.
 *
 * **왜 필요한가**: 실제 29만 행급 파일로 room.ts의 실제 시딩 경로(당시의 seedGridStylesFromXlsx가
 * 청크당 최대 2000개 행을 `Promise.all`로 동시에 읽던 실 사용 패턴 — 그 함수는 이후 삭제됐고 지금은
 * flushRowReads가 그 배치를 핸들 하나로 처리한다)를 검증하는 중, 파일이 방금 정상적으로
 * 쓰이고 닫혔는데도 이후의 `fs.open`이 간헐적으로 `ENOENT`를 던지는 현상을 실측했다 — 실패 시점이
 * 실행마다 달라(같은 재현 스크립트를 반복해도 일관된 행 번호가 아니었다) 이 코드의 결정적 경쟁이라기
 * 보다는 이 개발 환경(짧은 경로 표기 `OPENMI~1`이 나타나는 Windows 프로필)의 OS/파일시스템 계층
 * 불안정성으로 보인다 — 같은 임시 폴더를 대상으로 한 기본적인 `fs.watch` 호출조차 별개로 libuv
 * 내부 단언 실패로 죽는 것을 확인했다. 재시도 대기를 크게 늘려봐도(최대 약 1.85초) 근본적으로
 * 해결되지 않았다(대기 중에도 반복해서 실패해 결국 메모리 누적 후 OOM으로 이어짐) — 즉 이 재시도가
 * 그 환경 문제 자체를 고치지는 못한다. 그래도 이 코드베이스의 다른 파일 I/O(fs.rm의 maxRetries/
 * retryDelay)가 이미 같은 부류의 Windows 파일 I/O 불안정성에 짧은 재시도로 대응하는 것과 같은
 * 관용구를 유지하는 편이 맨 첫 시도에서 바로 포기하는 것보다는 낫다고 판단해 남긴다. 진짜로 파일이
 * 없는 경우(예: closeLazyGridStore로 이미 정리된 룸을 읽으려는 버그)는 재시도가 소진된 뒤 그대로
 * 던져 조용히 감춰지지 않는다.
 */
async function openRowStoreWithRetry(rowStorePath: string): Promise<FileHandle> {
  for (let attempt = 0; ; attempt += 1) {
    try {
      return await fs.open(rowStorePath, 'r');
    } catch (error) {
      const isEnoent = (error as NodeJS.ErrnoException).code === 'ENOENT';
      if (!isEnoent || attempt >= ROW_STORE_OPEN_MAX_RETRIES) throw error;
      await sleepIgnoringFakeTimers(ROW_STORE_OPEN_RETRY_DELAY_MS);
    }
  }
}

/**
 * 재시도 사이의 대기 — **전역 `setTimeout`을 쓰지 않는다**(리뷰 지적 I2).
 *
 * 왜: 이 재시도는 테스트가 기다리는 백그라운드 작업(페이지 시딩 → getRow → 이 open) 한가운데에
 * 있고, 룸 테스트들의 `flushEventLoop`는 이제 그 작업의 **완료 자체**를 await한다
 * (drainRoomBackgroundWorkForTest). 그 테스트 중 여럿이 가짜 타이머를 깔아 두는데
 * (`vi.useFakeTimers({ toFake: ['setTimeout', ...] })`), 가짜 타이머는 저절로 흐르지 않으므로
 * `setTimeout` 기반 대기가 이 경로에서 한 번이라도 걸리면 그 타이머는 **영원히 깨어나지 않고**
 * 테스트는 타임아웃까지 멈춘다(그리고 나중에 진짜 타이머로 되돌려도 이미 잠든 대기는 살아나지
 * 않는다). 이 환경은 실제로 간헐적 ENOENT가 관측되는 곳이라(위 docblock 참고) 도달 불가능한
 * 가정이 아니다.
 *
 * `node:timers/promises`의 `setTimeout`은 전역이 아니라 Node 내부 타이머를 직접 쓰기 때문에
 * 가짜 타이머 설치(@sinonjs/fake-timers는 `globalThis`만 교체한다)의 영향을 받지 않는다 —
 * 이 저장소의 vitest(3.2.7)에서 `toFake: ['setTimeout', ...]`와 기본 전량 가짜 타이머 양쪽 모두에
 * 대해 실제로 실측 확인했다. 그래서 진짜 타이머일 때의 동작(정확히 이 지연만큼 기다린다)은 그대로
 * 유지하면서, 가짜 타이머 아래에서도 CPU를 태우는 스핀 없이 제때 깨어난다.
 */
function sleepIgnoringFakeTimers(ms: number): Promise<void> {
  return timerSleep(ms);
}

function cacheKey(sheetIdx: number, originalRow: number): string {
  return `${sheetIdx}:${originalRow}`;
}

interface PendingRowRead {
  /** cacheKey(sheetIdx, originalRow) — inFlight 표에서 자신을 지우기 위해 들고 다닌다. */
  readonly key: string;
  readonly entry: LazyRowIndexEntry;
  readonly resolve: (row: ReadonlyArray<GridCell | 0>) => void;
  readonly reject: (error: unknown) => void;
}

interface RowReadBatch {
  /** 아직 핸들이 열리지 않은, 이번 배치로 묶일 행 읽기 요청들. */
  queue: PendingRowRead[];
  /** 플러시가 이미 예약돼 있는지 — 예약된 동안 들어온 요청은 그 플러시에 그대로 합류한다. */
  flushScheduled: boolean;
  /** cacheKey -> 지금 큐에 있거나 디스크에서 읽히는 중인 행의 결과(같은 행 중복 읽기 방지). */
  readonly inFlight: Map<string, Promise<ReadonlyArray<GridCell | 0>>>;
}

/**
 * 스토어별 행 읽기 배치 상태 — `LazyGridStore`의 **공개 데이터 모양은 건드리지 않는다**(그쪽은
 * ingestion 결과와 캐시만 담는 순수 데이터다). 스토어가 GC되면 이 항목도 같이 사라지도록 WeakMap에
 * 둔다. 배치 상태는 이벤트 루프 한 턴 안에서만 살아 있는 과도 상태라 스토어의 수명과 무관하다.
 */
const rowReadBatches = new WeakMap<LazyGridStore, RowReadBatch>();

function rowReadBatchFor(store: LazyGridStore): RowReadBatch {
  let batch = rowReadBatches.get(store);
  if (!batch) {
    batch = { queue: [], flushScheduled: false, inFlight: new Map() };
    rowReadBatches.set(store, batch);
  }
  return batch;
}

/**
 * 한 배치로 묶인 행 읽기를 **파일 핸들 하나로** 처리한다 — `fs.open` 1회 + `handle.read()` N회 +
 * `close` 1회.
 *
 * **왜 배치인가**: 예전에는 행 하나를 읽을 때마다 open→read→close를 새로 했다. 낮은 볼륨에서는
 * 괜찮지만, 페이지 시딩(`seedGridPageFromXlsx`가 500행을 `Promise.all`로 한꺼번에 읽는다)과 수식
 * 재계산(`recomputeFormulasForPage`가 청크당 125행 × 열 수만큼 `getCell`을 한꺼번에 흘린다)처럼
 * 한 번에 수백~수천 건이 몰리는 실제 경로에서는 이 파일 핸들 생명주기 syscall이 읽기 자체보다
 * 훨씬 비쌌다. 이 개발 환경에서 실측: 500회 open+read+close ≈ 145~200ms vs 같은 500회 읽기를
 * 핸들 하나로 ≈ 4ms(약 36배). Node 기본 libuv 스레드풀이 4개뿐이라 수백 개의 동시 open/close가
 * 작은 배치로 직렬화되면서 syscall 오버헤드를 매번 다 무는 것이 원인으로 보인다.
 *
 * **핸들 수명은 여전히 짧게 유지한다**: 스토어 전체 수명 동안 핸들을 들고 있지 않고 **배치 하나**
 * (= 대개 마이크로태스크 한 턴에 몰린 요청들)가 끝나면 바로 닫는다 — 예전 패턴과 같은 자릿수의
 * 아주 짧은 수명이다. 그래서 룸이 닫힐 때 `closeLazyGridStore`의 `fs.rm`이 Windows에서 "사용 중인
 * 파일"에 걸릴 위험은 예전과 같은 수준으로 유지되고(게다가 그 `fs.rm`은 원래부터 재시도한다),
 * `sweepOrphanedLazyGridFiles`의 나이 기반 안전망도 그대로 유효하다.
 *
 * 열기는 반드시 `openRowStoreWithRetry`를 거친다 — 이 환경의 간헐적 `ENOENT`에 대한 방어를
 * 배치화로 잃지 않기 위해서다(그 함수의 주석 참고). 배치당 한 번만 여니 재시도도 배치당 한 번이다.
 *
 * 같은 스토어에 배치가 동시에 여럿 뜨는 것(예: 두 페이지를 동시에 시딩)은 정상이고 안전하다 —
 * 각 배치가 자기 핸들을 따로 열고, 읽기는 모두 명시적 `position`을 넘기므로(파일 오프셋을 공유·
 * 변경하지 않는 pread 방식) 서로 간섭하지 않는다. 행 단위 중복은 `inFlight`가 막는다.
 */
async function flushRowReads(store: LazyGridStore, batch: RowReadBatch): Promise<void> {
  let handle: FileHandle;
  try {
    handle = await openRowStoreWithRetry(store.rowStorePath);
  } catch (error) {
    // 열기가 (재시도까지 소진하고) 실패하면 이 배치로 묶인 요청을 전부 같은 오류로 실패시킨다 —
    // 예전에 행마다 따로 열던 때와 같은 오류가 같은 호출부로 그대로 전달된다.
    const pending = batch.queue;
    batch.queue = [];
    batch.flushScheduled = false;
    for (const p of pending) {
      batch.inFlight.delete(p.key);
      p.reject(error);
    }
    return;
  }
  // 핸들이 열리기를 기다리는 동안 들어온 요청까지 같은 핸들로 함께 처리한다(그 요청들은
  // flushScheduled가 아직 true라 새 플러시를 예약하지 않았다). 여기서 비우고 나서야 플래그를
  // 내리므로, 이 시점 이후의 요청은 자기 배치를 새로 연다.
  const pending = batch.queue;
  batch.queue = [];
  batch.flushScheduled = false;
  try {
    await Promise.all(
      pending.map(async (p) => {
        try {
          const buf = Buffer.alloc(p.entry.byteLength);
          await handle.read(buf, 0, p.entry.byteLength, p.entry.byteOffset);
          p.resolve(JSON.parse(buf.toString('utf8')) as Array<GridCell | 0>);
        } catch (error) {
          p.reject(error);
        } finally {
          batch.inFlight.delete(p.key);
        }
      }),
    );
  } finally {
    await handle.close();
  }
}

/**
 * 행 하나를 디스크에서 읽는다 — 실제 읽기는 `flushRowReads`가 배치로 묶어 처리한다.
 *
 * **동기적으로** 큐에 넣고 마이크로태스크로 플러시를 예약하는 것이 핵심이다: 호출부들은 이미
 * `Promise.all(rows.map(...))` 모양으로 한 턴에 수백 건을 몰아 보내므로(seedGridPageFromXlsx,
 * recomputeFormulasForPage, applyCellsChangeToCanonical …), 그 요청들이 자연스럽게 한 배치 =
 * 핸들 하나로 묶인다. 호출부 시그니처를 바꾸지 않으므로 시딩 쪽 로직은 한 글자도 건드리지 않는다.
 */
function readRowFromDisk(
  store: LazyGridStore,
  sheetIdx: number,
  originalRow: number,
): Promise<ReadonlyArray<GridCell | 0>> {
  const entry = store.rowIndex[sheetIdx]?.[originalRow];
  if (!entry) return Promise.resolve([]);

  const batch = rowReadBatchFor(store);
  const key = cacheKey(sheetIdx, originalRow);
  // 같은 행을 동시에 여러 번 요청하면(예: 한 행의 여러 칸을 getCell로 동시에 훑는 재계산 패스)
  // 디스크 읽기는 한 번만 나가고 나머지는 같은 Promise를 공유한다. 행 데이터는 ingestion 이후
  // 불변이므로 공유해도 안전하다.
  const already = batch.inFlight.get(key);
  if (already) return already;

  let resolve!: (row: ReadonlyArray<GridCell | 0>) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<ReadonlyArray<GridCell | 0>>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  batch.inFlight.set(key, promise);
  batch.queue.push({ key, entry, resolve, reject });

  if (!batch.flushScheduled) {
    batch.flushScheduled = true;
    queueMicrotask(() => {
      void flushRowReads(store, batch).catch((error: unknown) => {
        // 개별 행 읽기 실패는 이미 각 요청으로 전달됐다 — 여기까지 오는 건 핸들 닫기 실패 정도라
        // 관찰할 요청이 남아 있지 않으므로 로그만 남긴다.
        console.error('행 저장소 배치 읽기 정리 실패:', error);
      });
    });
  }
  return promise;
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

export async function closeLazyGridStore(store: LazyGridStore): Promise<void> {
  await fs.rm(store.rowStorePath, { force: true, maxRetries: 5, retryDelay: 100 }).catch(() => {
    /* 정리 실패는 치명적이지 않다 — OS 임시 폴더 청소 또는 다음 서버 시작 시 sweep에 맡긴다 */
  });
}

/** sweepOrphanedLazyGridFiles가 "충분히 오래됐다"고 볼 나이(ms) — 이보다 최근에 수정된 파일은
 * 건드리지 않는다. 서버 비정상 종료로 남은 파일은 이 서버 프로세스가 시작하기 전부터 있었던
 * 것이라 항상 이 값보다 훨씬 오래됐다(재시작까지 걸린 시간만큼) — 그러니 안전망 성격은 그대로
 * 유지된다. 반면 같은 프로세스(또는 같은 테스트 스위트의 다른 워커)가 방금 막 만든, 지금 실제로
 * 쓰이고 있는 파일은 이 값보다 훨씬 최근이라 스윕 대상에서 제외된다 — 읽기 배치 사이사이에 핸들을
 * 열고 닫으므로(flushRowReads) Windows에서도 파일이 항상 잠겨 있지는 않아, 나이 확인 없이 무조건
 * 지우면 다른 워커가 마침 그 순간 사이에 읽고 있던 파일이 사라지는 경쟁이 실제로 있었다. 이 스윕은
 * 열린 핸들이 아니라 `fs.stat`의 mtime만 보므로 배치 읽기 여부와 무관하게 그대로 동작한다.
 */
const ORPHAN_SWEEP_MIN_AGE_MS = 60_000;

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
  const now = Date.now();
  await Promise.all(
    entries
      .filter((name) => name.startsWith('docview-lazygrid-'))
      .map(async (name) => {
        const filePath = path.join(dir, name);
        try {
          const stat = await fs.stat(filePath);
          if (now - stat.mtimeMs < ORPHAN_SWEEP_MIN_AGE_MS) return; // 최근 파일 — 지금 쓰이고 있을 수 있다
          await fs.rm(filePath, { force: true });
        } catch {
          /* 이미 없어졌거나 stat/rm이 실패해도 이 안전망은 다음 서버 시작 때 다시 시도된다 */
        }
      }),
  );
}
