import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import type { FastifyInstance, FastifyReply } from 'fastify';
import { recordAudit } from '../audit.js';
import { SizedLruCache } from '../cache.js';
import { config } from '../config.js';
import { findConverter } from '../converters/index.js';
import {
  PageOutOfRangeError,
  buildPrintPdf,
  burnMarkupIntoPdf,
  getPageDimensions,
  renderPage,
  type PageSize,
  type PersistedPageSizes,
} from '../rendering.js';
import { InvalidTokenError, verifyViewToken } from '../security.js';
import { extractPdfText, searchPdfText, type PageText } from '../search.js';
import { detectNearbyTextStyle } from '../textStyle.js';
import {
  docxHasUnsupportedContent,
  getDocxDefaultFontFamily,
  getDocxDefaultFontSizePt,
  getDocxPageSetup,
} from '../docx/import.js';
import { MarkupValidationError, addMarkup, deleteMarkup, listMarkup, updateMarkup } from '../markup.js';
import { touchPresence } from '../presence.js';
import { flushNow, getCurrentBytes } from '../collab/room.js';
import { computeLineOffsets, type GridCell, type GridHeader } from '../spreadsheet.js';
import { store } from '../storage.js';
import type { DocumentMeta, ViewClaims } from '../types.js';

export interface ViewContext {
  readonly claims: ViewClaims;
  readonly meta: DocumentMeta;
}

/**
 * 토큰 검증 + 문서 조회. 실패 시 reply를 직접 전송하고 null을 반환한다.
 * 토큰은 경로가 아닌 쿼리 파라미터로 받는다(JWT의 `.`이 경로 라우팅을 깨뜨리고
 * 액세스 로그에 토큰이 남는 문제를 피하기 위함).
 *
 * `src/routes/edit.ts`(동시편집 WebSocket 라우트)도 동일한 검증이 필요해 이 함수를 재사용한다.
 */
export async function resolveView(token: string | undefined, reply: FastifyReply): Promise<ViewContext | null> {
  if (!token) {
    await reply.code(401).send({ success: false, error: '열람 토큰이 없습니다.' });
    return null;
  }
  let claims: ViewClaims;
  try {
    claims = verifyViewToken(token);
  } catch (error) {
    if (error instanceof InvalidTokenError) {
      await reply.code(401).send({ success: false, error: error.message });
      return null;
    }
    throw error;
  }
  const meta = await store.readMeta(claims.docId);
  if (!meta || meta.status !== 'ready') {
    await reply.code(404).send({ success: false, error: '문서를 찾을 수 없습니다.' });
    return null;
  }
  return { claims, meta };
}

/**
 * JSONL 그리드 인덱스: 헤더만 파싱하고 행 데이터는 줄 오프셋으로 색인해
 * 요청된 범위의 줄만 그때그때 파싱한다 (수십만 행 문서 대응).
 */
interface LoadedGrid {
  readonly header: GridHeader;
  readonly buffer: Buffer;
  readonly lineOffsets: readonly number[];
  readonly sheetStartLine: readonly number[];
}

/** 그리드·PDF텍스트 캐시는 렌더링 문서 캐시와 나머지 예산을 나눠 쓴다 */
const gridCache = new SizedLruCache<LoadedGrid>({
  maxBytes: Math.floor(config.cacheBudgetBytes * 0.25),
  maxEntries: 4,
  sizeOf: (grid) => grid.buffer.length + grid.lineOffsets.length * 8,
});

const pdfTextCache = new SizedLruCache<PageText[]>({
  maxBytes: Math.floor(config.cacheBudgetBytes * 0.1),
  maxEntries: 6,
  sizeOf: (pages) =>
    pages.reduce(
      (sum, page) => sum + page.items.reduce((acc, item) => acc + item.str.length * 2 + 48, 0),
      0,
    ),
});

/**
 * 동시편집 저장(src/collab/room.ts)이 문서를 다시 생성한 뒤 호출한다. docx/pptx는 새 PDF로
 * 재변환된 뒤 이전 텍스트 레이어가 검색 결과에 남지 않도록, xlsx는 그리드(+gridindex)가 다시
 * 저장된 뒤 이전 LoadedGrid가 캐시에 남아 낡은 내용을 서빙하지 않도록 각각 무효화한다.
 */
export function invalidateViewCaches(docId: string): void {
  pdfTextCache.invalidate(docId);
  gridCache.invalidate(docId);
}

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

/**
 * 영구 저장된 gridindex가 구조적으로 말이 되는지 값싸게(오프셋 배열만 훑고, 버퍼 바이트는
 * 건드리지 않는다) 검사한다 — 실제 스캔 결과와 바이트 단위로 똑같은지까지는 확인하지 않는다
 * (그러면 인덱스를 신뢰하는 의미가 없어진다). 이 검사의 목적은 딱 하나: Task 6에서 실제로
 * 겪은 버그 클래스 — 마지막 오프셋이 버퍼 길이를 넘어서 Buffer.toString이 end를 조용히
 * clamp해 여러 줄이 뒤섞인 문자열을 반환하고 JSON.parse가 깨지는 것 — 을 막는다.
 */
function offsetsLookValid(buffer: Buffer, lineOffsets: readonly number[]): boolean {
  if (lineOffsets.length === 0 || lineOffsets[0] !== 0) return false;
  for (let i = 1; i < lineOffsets.length; i += 1) {
    if (lineOffsets[i] <= lineOffsets[i - 1]) return false;
  }
  return lineOffsets[lineOffsets.length - 1] <= buffer.length;
}

async function loadGridUncached(docId: string): Promise<LoadedGrid> {
  const buffer = await store.readDerived(docId, 'grid');
  try {
    const raw = await store.readDerived(docId, 'gridindex');
    const { lineOffsets } = JSON.parse(raw.toString('utf8')) as { lineOffsets: number[] };
    if (offsetsLookValid(buffer, lineOffsets)) {
      return buildLoadedGridFromOffsets(buffer, lineOffsets);
    }
  } catch {
    // gridindex가 없거나(정상 — 최초 조회), 손상됐거나, 구조적으로 말이 안 되면(offsetsLookValid
    // 실패) 여기로 떨어진다 — 아래에서 다시 스캔해 올바른 값으로 덮어써 자가치유한다.
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
      try {
        const raw = await store.readDerived(docId, 'pagesizes');
        const sizes = JSON.parse(raw.toString('utf8')) as unknown;
        return Array.isArray(sizes) ? (sizes as readonly PageSize[]) : null;
      } catch {
        return null; // 없음/손상 모두 캐시 미스로 취급 — getPageSizes가 다시 계산해 덮어써 자가치유한다.
      }
    },
    save: async (sizes) => {
      await store.saveDerived(docId, 'pagesizes', Buffer.from(JSON.stringify(sizes)));
    },
  };
}

/** PDF 텍스트 추출 결과 캐시 (문서당 1회 추출) */
function loadPdfText(docId: string): Promise<PageText[]> {
  return pdfTextCache.getOrLoad(docId, () => store.readDerived(docId, 'pdf').then(extractPdfText));
}

interface GridSearchHit {
  readonly sheet: number;
  readonly r: number;
  readonly c: number;
  readonly snippet: string;
}

/** JSONL 그리드에서 부분 문자열을 검색한다. 줄 단위 1차 필터 후 매칭 줄만 파싱한다. */
function searchGrid(grid: LoadedGrid, query: string, limit: number): { hits: GridSearchHit[]; total: number } {
  const needle = query.toLowerCase();
  const hits: GridSearchHit[] = [];
  let total = 0;
  for (let sheetIdx = 0; sheetIdx < grid.header.sheets.length; sheetIdx += 1) {
    const sheet = grid.header.sheets[sheetIdx];
    const startLine = grid.sheetStartLine[sheetIdx];
    for (let r = 1; r <= sheet.rowCount; r += 1) {
      const line = startLine + r - 1;
      const begin = grid.lineOffsets[line];
      const finish =
        line + 1 < grid.lineOffsets.length ? grid.lineOffsets[line + 1] - 1 : grid.buffer.length;
      const raw = grid.buffer.toString('utf8', begin, finish);
      if (!raw.toLowerCase().includes(needle)) continue;
      const cells = JSON.parse(raw) as Array<GridCell | 0>;
      for (let c = 0; c < cells.length; c += 1) {
        const cell = cells[c];
        if (cell === 0 || !cell.v.toLowerCase().includes(needle)) continue;
        total += 1;
        if (hits.length < limit) {
          hits.push({ sheet: sheetIdx, r, c: c + 1, snippet: cell.v.slice(0, 120) });
        }
      }
    }
  }
  return { hits, total };
}

function sliceRows(
  grid: LoadedGrid,
  sheetIdx: number,
  from: number,
  count: number,
): Array<Array<GridCell | 0>> {
  const sheet = grid.header.sheets[sheetIdx];
  const start = grid.sheetStartLine[sheetIdx] + from - 1;
  const end = Math.min(start + count, grid.sheetStartLine[sheetIdx] + sheet.rowCount);
  const rows: Array<Array<GridCell | 0>> = [];
  for (let line = start; line < end; line += 1) {
    const begin = grid.lineOffsets[line];
    const finish = line + 1 < grid.lineOffsets.length ? grid.lineOffsets[line + 1] - 1 : grid.buffer.length;
    rows.push(JSON.parse(grid.buffer.toString('utf8', begin, finish)));
  }
  return rows;
}

/**
 * 그리드 문서의 PDF는 업로드 시 만들지 않고 출력 요청 시 생성한다.
 * (대형 리스트의 수 분짜리 LibreOffice 변환을 열람 경로에서 제거)
 */
async function ensurePdf(meta: DocumentMeta): Promise<Buffer> {
  if (await store.hasDerived(meta.id, 'pdf')) {
    return store.readDerived(meta.id, 'pdf');
  }
  const converter = findConverter(meta.ext);
  if (!converter) {
    throw new Error(`PDF 변환기를 찾을 수 없습니다: ${meta.ext}`);
  }
  const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-print-'));
  try {
    const sourcePath = path.join(workDir, `source.${meta.ext}`);
    await fs.writeFile(sourcePath, await store.readOriginal(meta.id, meta.ext));
    const targetPdfPath = path.join(workDir, 'converted.pdf');
    await converter.convert(sourcePath, targetPdfPath, meta.name);
    const pdfBytes = await fs.readFile(targetPdfPath);
    await store.saveDerived(meta.id, 'pdf', pdfBytes);
    return pdfBytes;
  } finally {
    await fs.rm(workDir, { recursive: true, force: true });
  }
}

// token 누락/무효는 resolveView가 401로 일관 처리하도록 스키마에서 required로 강제하지 않는다.
const tokenQuerySchema = {
  type: 'object',
  properties: { token: { type: 'string' } },
} as const;

const presenceBodySchema = {
  type: 'object',
  required: ['sessionId'],
  additionalProperties: false,
  properties: { sessionId: { type: 'string', minLength: 1, maxLength: 100 } },
} as const;

interface PresenceBody {
  readonly sessionId: string;
}

const pageRequestSchema = {
  params: {
    type: 'object',
    required: ['page'],
    properties: { page: { type: 'integer', minimum: 1 } },
  },
  querystring: {
    type: 'object',
    properties: {
      token: { type: 'string' },
      scale: { type: 'integer', minimum: 1, maximum: 16, default: 2 },
    },
  },
} as const;

const gridRowsSchema = {
  type: 'object',
  properties: {
    token: { type: 'string' },
    sheet: { type: 'integer', minimum: 0, default: 0 },
    from: { type: 'integer', minimum: 1, default: 1 },
    count: { type: 'integer', minimum: 1, maximum: 500, default: 200 },
  },
} as const;

interface TokenQuery {
  readonly token?: string;
}

const tokenWithMarkupQuerySchema = {
  type: 'object',
  properties: { token: { type: 'string' }, includeMarkup: { type: 'string' } },
} as const;

interface TokenWithMarkupQuery extends TokenQuery {
  readonly includeMarkup?: string;
}

const addMarkupBodySchema = {
  type: 'object',
  required: ['page', 'x', 'y', 'text'],
  additionalProperties: false,
  properties: {
    kind: { type: 'string', enum: ['comment', 'text'] },
    page: { type: 'integer', minimum: 1 },
    x: { type: 'number', minimum: 0, maximum: 1 },
    y: { type: 'number', minimum: 0, maximum: 1 },
    w: { type: 'number', minimum: 0, maximum: 1 },
    h: { type: 'number', minimum: 0, maximum: 1 },
    fontSizePt: { type: 'number', minimum: 1, maximum: 400 },
    bold: { type: 'boolean' },
    fontFamily: { type: 'string', enum: ['sans-serif', 'serif', 'monospace'] },
    text: { type: 'string', minLength: 1, maxLength: 2000 },
  },
  if: { properties: { kind: { const: 'text' } }, required: ['kind'] },
  then: { required: ['fontSizePt'] },
  else: { required: ['w', 'h'] },
} as const;

interface AddMarkupBody {
  readonly kind?: 'comment' | 'text';
  readonly page: number;
  readonly x: number;
  readonly y: number;
  readonly w?: number;
  readonly h?: number;
  readonly fontSizePt?: number;
  readonly bold?: boolean;
  readonly fontFamily?: 'sans-serif' | 'serif' | 'monospace';
  readonly text: string;
}

const updateMarkupBodySchema = {
  type: 'object',
  required: ['x', 'y'],
  additionalProperties: false,
  properties: {
    x: { type: 'number', minimum: 0, maximum: 1 },
    y: { type: 'number', minimum: 0, maximum: 1 },
    fontSizePt: { type: 'number', minimum: 1, maximum: 400 },
  },
} as const;

interface UpdateMarkupBody {
  readonly x: number;
  readonly y: number;
  readonly fontSizePt?: number;
}

const markupParamsSchema = {
  type: 'object',
  required: ['markupId'],
  properties: { markupId: { type: 'string' } },
} as const;

const textStyleSchema = {
  params: {
    type: 'object',
    required: ['page'],
    properties: { page: { type: 'integer', minimum: 1 } },
  },
  querystring: {
    type: 'object',
    required: ['x', 'y'],
    properties: {
      token: { type: 'string' },
      x: { type: 'number', minimum: 0, maximum: 1 },
      y: { type: 'number', minimum: 0, maximum: 1 },
    },
  },
} as const;

interface TextStyleQuery extends TokenQuery {
  readonly x: number;
  readonly y: number;
}

interface PageQuery extends TokenQuery {
  readonly scale: number;
}

interface GridRowsQuery extends TokenQuery {
  readonly sheet: number;
  readonly from: number;
  readonly count: number;
}

const searchQuerySchema = {
  type: 'object',
  required: ['q'],
  properties: {
    token: { type: 'string' },
    q: { type: 'string', minLength: 1, maxLength: 200 },
    limit: { type: 'integer', minimum: 1, maximum: 500, default: 200 },
  },
} as const;

interface SearchQuery extends TokenQuery {
  readonly q: string;
  readonly limit: number;
}

export async function viewerRoutes(app: FastifyInstance): Promise<void> {
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/info',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      await recordAudit({ type: 'opened', docId: view.meta.id, viewer: view.claims.viewer });
      const viewKind = view.meta.viewKind ?? 'pages';
      // 페이지 크기를 미리 내려 뷰어가 로딩 전 프레임 크기를 고정한다 (스크롤 점프 정확도)
      const pages =
        viewKind === 'pages'
          ? await getPageDimensions(
              view.meta.id,
              () => store.readDerived(view.meta.id, 'pdf'),
              pageSizePersistence(view.meta.id),
            ).catch(() => undefined)
          : undefined;
      // 편집기가 열리기 전 "저장하면 이미지가 사라집니다" 경고를 띄우기 위한 사전 검사
      // (canEdit 토큰 + docx일 때만 — 원본을 다시 읽는 비용이 있으므로 불필요하면 건너뛴다)
      const editHasUnsupportedContent =
        view.claims.perms.canEdit && view.meta.ext === 'docx'
          ? await store
              .readOriginal(view.meta.id, view.meta.ext)
              .then(docxHasUnsupportedContent)
              .catch(() => false)
          : false;
      // docx 에디터가 실제 Word 페이지 크기·여백 비율로 화면을 그리기 위한 값(src/client/editor.ts).
      // 원본을 다시 읽어야 해서(canEdit 여부와 무관하게) docx일 때만 계산한다.
      const pageSetup =
        view.meta.ext === 'docx'
          ? await store.readOriginal(view.meta.id, view.meta.ext).then(getDocxPageSetup).catch(() => null)
          : null;
      // 직접 서식도 스타일도 없는 실행문이 상속받는 문서 기본 폰트 크기(w:docDefaults) —
      // 없으면 브라우저 기본값(16px)이 적용되어 Word가 실제로 보여주는 크기와 달라진다.
      const defaultFontSizePt =
        view.meta.ext === 'docx'
          ? await store
              .readOriginal(view.meta.id, view.meta.ext)
              .then(getDocxDefaultFontSizePt)
              .catch(() => null)
          : null;
      // 문서 기본 서체(word/theme/theme1.xml로 해석) — 없으면 null(에디터가 기존 폰트를 그대로 씀).
      const defaultFontFamily =
        view.meta.ext === 'docx'
          ? await store
              .readOriginal(view.meta.id, view.meta.ext)
              .then(getDocxDefaultFontFamily)
              .catch(() => null)
          : null;
      return {
        success: true,
        data: {
          // 편집기(editor.js)가 협업 WebSocket 룸 이름으로 사용한다 — 토큰이 이미 이 문서로
          // 범위가 고정돼 있으므로 docId 자체를 별도로 노출해도 새로운 권한 노출은 없다.
          docId: view.meta.id,
          name: view.meta.name,
          ext: view.meta.ext,
          pageCount: view.meta.pageCount,
          viewKind,
          pages,
          permissions: view.claims.perms,
          viewer: view.claims.viewer,
          editHasUnsupportedContent,
          pageSetup,
          defaultFontSizePt,
          defaultFontFamily,
        },
      };
    },
  );

  /**
   * "지금 이 문서를 보고 있는 사람" 표시용 하트비트 — pages/grid 등 viewKind와 무관하게 viewer.html을
   * 여는 모든 문서에 동일하게 적용된다(src/presence.ts). 클라이언트가 주기적으로 호출하며, 응답에
   * 갱신된 현재 접속자 목록을 실어 보내 별도 조회 요청이 필요 없게 한다.
   */
  app.post<{ Querystring: TokenQuery; Body: PresenceBody }>(
    '/api/view/presence',
    { schema: { querystring: tokenQuerySchema, body: presenceBodySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      const viewers = touchPresence(view.meta.id, request.body.sessionId, view.claims.viewer);
      return { success: true, data: { viewers } };
    },
  );

  // Ctrl+S — 8초 유휴 자동저장을 기다리지 않고 지금 바로 저장한다(Office 365의 수동 저장과 동일한
  // 역할이지만, 이 앱은 이미 계속 자동저장하므로 사실상 "저장 큐를 지금 비워라"에 가깝다).
  app.post<{ Querystring: TokenQuery }>(
    '/api/view/save',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canEdit) {
        return reply.code(403).send({ success: false, error: '편집 권한이 없습니다.' });
      }
      const saved = await flushNow(view.meta.id);
      return { success: true, data: { saved } };
    },
  );

  app.get<{ Params: { page: number }; Querystring: PageQuery }>(
    '/api/view/pages/:page',
    { schema: pageRequestSchema },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      try {
        const png = await renderPage(
          view.meta.id,
          request.params.page,
          request.query.scale,
          () => store.readDerived(view.meta.id, 'pdf'),
          pageSizePersistence(view.meta.id),
        );
        return await reply.type('image/png').header('cache-control', 'private, no-store').send(png);
      } catch (error) {
        if (error instanceof PageOutOfRangeError) {
          return reply.code(404).send({ success: false, error: error.message });
        }
        request.log.error(error, 'page rendering failed');
        return reply.code(500).send({ success: false, error: '페이지 렌더링에 실패했습니다.' });
      }
    },
  );

  /** 그리드 메타: 시트 목록·크기·병합·스타일 팔레트 (셀 데이터 제외) */
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/grid/meta',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if ((view.meta.viewKind ?? 'pages') !== 'grid') {
        return reply.code(400).send({ success: false, error: '그리드 뷰 문서가 아닙니다.' });
      }
      const grid = await loadGrid(view.meta.id);
      return { success: true, data: grid.header };
    },
  );

  /** 그리드 행 데이터를 범위 단위로 반환한다 (가상 스크롤용) */
  app.get<{ Querystring: GridRowsQuery }>(
    '/api/view/grid/rows',
    { schema: { querystring: gridRowsSchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if ((view.meta.viewKind ?? 'pages') !== 'grid') {
        return reply.code(400).send({ success: false, error: '그리드 뷰 문서가 아닙니다.' });
      }
      const grid = await loadGrid(view.meta.id);
      if (!grid.header.sheets[request.query.sheet]) {
        return reply.code(404).send({ success: false, error: '시트를 찾을 수 없습니다.' });
      }
      const from = request.query.from;
      const rows = sliceRows(grid, request.query.sheet, from, request.query.count);
      return { success: true, data: { from, rows } };
    },
  );

  /**
   * 문서 내 문자 검색.
   * grid 문서: 셀 값 검색 → {kind:'grid', hits:[{sheet,r,c,snippet}]}
   * pages 문서: PDF 텍스트 레이어 검색 → {kind:'pages', hits:[{page,rects,snippet}]}
   * rects는 페이지 크기 대비 비율(0~1)이라 클라이언트가 배율과 무관하게 하이라이트를 그린다.
   */
  app.get<{ Querystring: SearchQuery }>(
    '/api/view/search',
    { schema: { querystring: searchQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      const { q, limit } = request.query;
      try {
        if ((view.meta.viewKind ?? 'pages') === 'grid') {
          const grid = await loadGrid(view.meta.id);
          const { hits, total } = searchGrid(grid, q, limit);
          return { success: true, data: { kind: 'grid', hits, total } };
        }
        const pages = await loadPdfText(view.meta.id);
        const { hits, total } = searchPdfText(pages, q, limit);
        return { success: true, data: { kind: 'pages', hits, total } };
      } catch (error) {
        request.log.error(error, 'search failed');
        return reply.code(500).send({ success: false, error: '검색에 실패했습니다.' });
      }
    },
  );

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
        const body = request.body;
        const item = await addMarkup(
          view.meta.id,
          view.meta,
          body.kind === 'text'
            ? {
                kind: 'text',
                page: body.page,
                x: body.x,
                y: body.y,
                fontSizePt: body.fontSizePt ?? 12,
                bold: body.bold ?? false,
                fontFamily: body.fontFamily,
                text: body.text,
                author: view.claims.viewer,
              }
            : {
                kind: 'comment',
                page: body.page,
                x: body.x,
                y: body.y,
                w: body.w ?? 0,
                h: body.h ?? 0,
                text: body.text,
                author: view.claims.viewer,
              },
        );
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

  app.patch<{ Querystring: TokenQuery; Params: { markupId: string }; Body: UpdateMarkupBody }>(
    '/api/view/markup/:markupId',
    { schema: { querystring: tokenQuerySchema, params: markupParamsSchema, body: updateMarkupBodySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canEdit) {
        await recordAudit({ type: 'markup_move_denied', docId: view.meta.id, viewer: view.claims.viewer });
        return reply.code(403).send({ success: false, error: '마크업 이동 권한이 없습니다.' });
      }
      try {
        const updated = await updateMarkup(view.meta.id, request.params.markupId, request.body);
        if (!updated) {
          return reply.code(404).send({ success: false, error: '마크업을 찾을 수 없습니다.' });
        }
        await recordAudit({
          type: 'markup_moved',
          docId: view.meta.id,
          viewer: view.claims.viewer,
          detail: updated.id,
        });
        return { success: true, data: updated };
      } catch (error) {
        if (error instanceof MarkupValidationError) {
          return reply.code(400).send({ success: false, error: error.message });
        }
        request.log.error(error, 'markup move failed');
        return reply.code(500).send({ success: false, error: '마크업 이동에 실패했습니다.' });
      }
    },
  );

  /**
   * 클릭 지점 근처의 기존 텍스트에서 글자 크기·굵기·베이스라인을 추정한다 — "텍스트 추가" 도구가
   * 새 텍스트를 주변 폰트와 어울리게 자동으로 놓기 위해 쓴다(src/textStyle.ts).
   */
  app.get<{ Params: { page: number }; Querystring: TextStyleQuery }>(
    '/api/view/pages/:page/text-style',
    { schema: textStyleSchema },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if ((view.meta.viewKind ?? 'pages') !== 'pages') {
        return reply.code(400).send({ success: false, error: '이 문서 형식은 지원하지 않습니다.' });
      }
      try {
        const pages = await loadPdfText(view.meta.id);
        const pageText = pages[request.params.page - 1];
        if (!pageText) {
          return reply.code(404).send({ success: false, error: '페이지를 찾을 수 없습니다.' });
        }
        const style = detectNearbyTextStyle(pageText, request.query.x, request.query.y);
        return { success: true, data: style };
      } catch (error) {
        request.log.error(error, 'text style detection failed');
        return reply.code(500).send({ success: false, error: '폰트 스타일 감지에 실패했습니다.' });
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
}
