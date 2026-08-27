import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import type { FastifyInstance, FastifyReply } from 'fastify';
import { recordAudit } from '../audit.js';
import { SizedLruCache } from '../cache.js';
import { config } from '../config.js';
import { findConverter } from '../converters/index.js';
import { PageOutOfRangeError, buildPrintPdf, getPageDimensions, renderPage } from '../rendering.js';
import { InvalidTokenError, verifyViewToken } from '../security.js';
import { extractPdfText, searchPdfText, type PageText } from '../search.js';
import { docxHasUnsupportedContent } from '../docx/import.js';
import type { GridCell, GridHeader } from '../spreadsheet.js';
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
 * docx 동시편집 저장(src/collab/room.ts)이 문서를 새 PDF로 재변환한 뒤 호출한다.
 * 이전 PDF 기준으로 추출된 텍스트 레이어가 검색 결과에 남지 않도록 지운다.
 */
export function invalidateViewCaches(docId: string): void {
  pdfTextCache.invalidate(docId);
  gridCache.invalidate(docId);
}

function indexGrid(buffer: Buffer): LoadedGrid {
  const lineOffsets: number[] = [0];
  for (let i = buffer.indexOf(10); i !== -1; i = buffer.indexOf(10, i + 1)) {
    lineOffsets.push(i + 1);
  }
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

function loadGrid(docId: string): Promise<LoadedGrid> {
  return gridCache.getOrLoad(docId, () => store.readDerived(docId, 'grid').then(indexGrid));
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
          ? await getPageDimensions(view.meta.id, () => store.readDerived(view.meta.id, 'pdf')).catch(
              () => undefined,
            )
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
      return {
        success: true,
        data: {
          // 편집기(editor.js)가 협업 WebSocket 룸 이름으로 사용한다 — 토큰이 이미 이 문서로
          // 범위가 고정돼 있으므로 docId 자체를 별도로 노출해도 새로운 권한 노출은 없다.
          docId: view.meta.id,
          name: view.meta.name,
          pageCount: view.meta.pageCount,
          viewKind,
          pages,
          permissions: view.claims.perms,
          viewer: view.claims.viewer,
          editHasUnsupportedContent,
        },
      };
    },
  );

  app.get<{ Params: { page: number }; Querystring: PageQuery }>(
    '/api/view/pages/:page',
    { schema: pageRequestSchema },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      try {
        const png = await renderPage(view.meta.id, request.params.page, request.query.scale, () =>
          store.readDerived(view.meta.id, 'pdf'),
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
    '/api/view/print',
    { schema: { querystring: tokenQuerySchema } },
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
      const watermark = `${view.claims.viewer}  ${new Date().toISOString().slice(0, 10)}`;
      const pdfBytes = await buildPrintPdf(pdfSource, watermark);
      await recordAudit({ type: 'printed', docId: view.meta.id, viewer: view.claims.viewer });
      return reply
        .type('application/pdf')
        .header('content-disposition', 'inline; filename="print.pdf"')
        .send(Buffer.from(pdfBytes));
    },
  );

  app.get<{ Querystring: TokenQuery }>(
    '/api/view/download',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canDownload) {
        await recordAudit({ type: 'download_denied', docId: view.meta.id, viewer: view.claims.viewer });
        return reply.code(403).send({ success: false, error: '저장(다운로드) 권한이 없습니다.' });
      }
      const data = await store.readOriginal(view.meta.id, view.meta.ext);
      await recordAudit({ type: 'downloaded', docId: view.meta.id, viewer: view.claims.viewer });
      const encodedName = encodeURIComponent(view.meta.name);
      return reply
        .type('application/octet-stream')
        .header('content-disposition', `attachment; filename*=UTF-8''${encodedName}`)
        .send(data);
    },
  );
}
