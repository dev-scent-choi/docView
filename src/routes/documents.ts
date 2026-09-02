import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import type { FastifyInstance } from 'fastify';
import { requireApiKey, type ApiKeyAuthenticatedRequest } from '../apiKeyAuth.js';
import { recordAudit } from '../audit.js';
import { config } from '../config.js';
import { findConverter, supportedExtensions } from '../converters/index.js';
import { ConversionError, type Converter } from '../converters/types.js';
import { getPageCountFromBytes, invalidateRenderCache } from '../rendering.js';
import { issueViewToken } from '../security.js';
import { parseCsvToGrid, parseXlsxToGrid, serializeGrid, type GridDocument } from '../spreadsheet.js';
import { newDocumentId, store } from '../storage.js';
import type { DocumentMeta } from '../types.js';
import { invalidateViewCaches } from './viewer.js';

const tokenBodySchema = {
  type: 'object',
  required: ['viewer'],
  additionalProperties: false,
  properties: {
    viewer: { type: 'string', minLength: 1, maxLength: 100 },
    canPrint: { type: 'boolean', default: false },
    canDownload: { type: 'boolean', default: false },
    canCopy: { type: 'boolean', default: false },
    canEdit: { type: 'boolean', default: false },
  },
} as const;

interface TokenBody {
  readonly viewer: string;
  readonly canPrint: boolean;
  readonly canDownload: boolean;
  readonly canCopy: boolean;
  readonly canEdit: boolean;
}

/** LibreOffice가 중간 진행률을 제공하지 않으므로 단계 완료 시점마다 진행률을 기록한다. */
const PROGRESS_CONVERTING = 25;
const PROGRESS_PDF_DONE = 85;

/** 그리드 뷰로 여는 스프레드시트 확장자 (PDF는 출력 시점에 지연 생성) */
const GRID_VIEW_EXTS = new Set(['xlsx', 'csv']);

/** 실시간 동시편집(에디터)을 지원하는 확장자 — docx는 문서 편집기, xlsx는 그리드 편집기로 연결된다 */
const EDITABLE_EXTS = new Set(['docx', 'xlsx']);

function editorPathFor(ext: string): string {
  return ext === 'xlsx' ? '/editor-grid.html' : '/editor.html';
}

async function buildGrid(meta: DocumentMeta): Promise<GridDocument> {
  const original = await store.readOriginal(meta.id, meta.ext);
  if (meta.ext === 'csv') {
    return parseCsvToGrid(original.toString('utf8'));
  }
  return parseXlsxToGrid(original);
}

/**
 * 스프레드시트는 PDF 변환 없이 셀 데이터만 추출해 즉시 열람 가능하게 한다.
 * (대형 리스트의 수 분짜리 PDF 변환 대기 제거 — PDF는 출력 요청 시 생성)
 */
async function runGridExtraction(app: FastifyInstance, meta: DocumentMeta): Promise<void> {
  try {
    await store.writeMeta({ ...meta, progress: PROGRESS_CONVERTING });
    const grid = await buildGrid(meta);
    if (grid.sheets.length === 0) {
      throw new ConversionError('표시할 시트가 없습니다. 파일이 비어 있거나 지원되지 않는 내용입니다.');
    }
    await store.saveDerived(meta.id, 'grid', serializeGrid(grid));
    await store.writeMeta({
      ...meta,
      status: 'ready',
      viewKind: 'grid',
      pageCount: grid.sheets.length,
      progress: 100,
    });
    await recordAudit({ type: 'converted', docId: meta.id, detail: `${grid.sheets.length} sheets (grid)` });
  } catch (error) {
    app.log.error(error, 'grid extraction failed');
    const userMessage =
      error instanceof ConversionError
        ? error.message
        : '스프레드시트를 읽지 못했습니다. 파일이 손상되었거나 지원되지 않는 형식입니다.';
    await store.writeMeta({ ...meta, status: 'failed', error: userMessage });
    await recordAudit({ type: 'conversion_failed', docId: meta.id, detail: meta.name });
  }
}

/**
 * 문서 원본을 변환기로 PDF화해 저장하고 페이지 수를 반환한다. 파일 경로가 필요한
 * 외부 도구(LibreOffice 등) 특성상 임시 작업 폴더에 꺼냈다가 변환 후 즉시 삭제한다
 * (memory 모드에서도 변환 순간 외에는 디스크에 남는 것이 없다).
 *
 * `src/collab/room.ts`의 동시편집 저장도 재변환에 이 함수를 그대로 재사용한다 —
 * 진행률·상태 갱신은 호출자(runConversion 또는 collab room)가 각자의 흐름에 맞게 처리한다.
 */
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
    const pageCount = await getPageCountFromBytes(pdfBytes);
    return { pdfBytes, pageCount };
  } finally {
    await fs.rm(workDir, { recursive: true, force: true });
  }
}

async function runConversion(app: FastifyInstance, converter: Converter, meta: DocumentMeta): Promise<void> {
  try {
    await store.writeMeta({ ...meta, progress: PROGRESS_CONVERTING });
    const { pageCount } = await convertToPdf(meta, converter);
    await store.writeMeta({ ...meta, progress: PROGRESS_PDF_DONE });
    await store.writeMeta({ ...meta, status: 'ready', viewKind: 'pages', pageCount, progress: 100 });
    await recordAudit({ type: 'converted', docId: meta.id, detail: `${pageCount} pages` });
  } catch (error) {
    app.log.error(error, 'document conversion failed');
    const userMessage =
      error instanceof ConversionError
        ? error.message
        : '문서 변환에 실패했습니다. 파일이 손상되었거나 지원되지 않는 내용입니다.';
    await store.writeMeta({ ...meta, status: 'failed', error: userMessage });
    await recordAudit({ type: 'conversion_failed', docId: meta.id, detail: meta.name });
  }
}

export async function documentRoutes(app: FastifyInstance): Promise<void> {
  app.post('/api/documents', { preHandler: requireApiKey }, async (request, reply) => {
    const callerProject = (request as ApiKeyAuthenticatedRequest).callerProject;
    let file;
    let data: Buffer;
    try {
      file = await request.file();
      if (!file) {
        return await reply.code(400).send({ success: false, error: '업로드된 파일이 없습니다.' });
      }
      data = await file.toBuffer();
    } catch (error) {
      const code = (error as { code?: string }).code;
      if (code === 'FST_REQ_FILE_TOO_LARGE') {
        return reply.code(413).send({ success: false, error: '파일이 업로드 용량 제한을 초과했습니다.' });
      }
      request.log.error(error, 'upload parsing failed');
      return reply.code(400).send({ success: false, error: '업로드 요청을 처리할 수 없습니다.' });
    }

    const ext = path.extname(file.filename).slice(1).toLowerCase();
    const converter = findConverter(ext);
    if (!converter) {
      return reply.code(415).send({
        success: false,
        error: `지원하지 않는 형식입니다: .${ext || '(없음)'} — 지원 형식: ${supportedExtensions.join(', ')}`,
      });
    }

    const id = newDocumentId();
    try {
      await store.saveOriginal(id, ext, data);
      const meta: DocumentMeta = {
        id,
        name: file.filename,
        ext,
        status: 'processing',
        pageCount: 0,
        createdAt: new Date().toISOString(),
        progress: 10,
      };
      await store.writeMeta(meta);
      await recordAudit({
        type: 'uploaded',
        docId: id,
        detail: callerProject ? `${file.filename} (caller=${callerProject})` : file.filename,
      });
      // 변환은 백그라운드로 진행하고 즉시 응답한다. 클라이언트는 상태를 폴링한다.
      if (GRID_VIEW_EXTS.has(ext)) {
        void runGridExtraction(app, meta);
      } else {
        void runConversion(app, converter, meta);
      }
      return await reply.code(201).send({ success: true, data: meta });
    } catch (error) {
      request.log.error(error, 'upload persisting failed');
      return reply.code(500).send({ success: false, error: '업로드 저장에 실패했습니다.' });
    }
  });

  app.get<{ Params: { id: string } }>('/api/documents/:id', async (request, reply) => {
    const meta = await store.readMeta(request.params.id);
    if (!meta) {
      return reply.code(404).send({ success: false, error: '문서를 찾을 수 없습니다.' });
    }
    return { success: true, data: meta };
  });

  app.get('/api/documents', async () => {
    const documents = await store.listMeta();
    return { success: true, data: documents };
  });

  app.get('/api/formats', async () => ({ success: true, data: supportedExtensions }));

  app.delete<{ Params: { id: string } }>(
    '/api/documents/:id',
    { preHandler: requireApiKey },
    async (request, reply) => {
      const meta = await store.readMeta(request.params.id);
      if (!meta) {
        return reply.code(404).send({ success: false, error: '문서를 찾을 수 없습니다.' });
      }
      await store.deleteDocument(meta.id);
      invalidateRenderCache(meta.id);
      invalidateViewCaches(meta.id);
      await recordAudit({ type: 'deleted', docId: meta.id, detail: meta.name });
      return { success: true, data: { id: meta.id } };
    },
  );

  app.post<{ Params: { id: string }; Body: TokenBody }>(
    '/api/documents/:id/tokens',
    { schema: { body: tokenBodySchema }, preHandler: requireApiKey },
    async (request, reply) => {
      const callerProject = (request as ApiKeyAuthenticatedRequest).callerProject;
      const meta = await store.readMeta(request.params.id);
      if (!meta || meta.status !== 'ready') {
        return reply.code(404).send({ success: false, error: '문서를 찾을 수 없습니다.' });
      }
      const { viewer, canPrint, canDownload, canCopy, canEdit } = request.body;
      const isEditableExt = EDITABLE_EXTS.has(meta.ext);
      // canEdit은 두 가지 용도로 쓰인다: (1) docx/xlsx 실시간 동시편집 참여 허용 — 이 경우 해당
      // 확장자용 변환기가 실제로 있어야 한다(편집 UI는 있는데 저장이 막히는 상황을 막기 위함).
      // (2) PDF 등 pages 뷰 문서의 마크업(리뷰 코멘트) 추가/삭제 허용 — 동시편집 WebSocket과 무관한
      // 뷰어 전용 권한이라 변환기 유무와 상관없이 요청한 대로 발급한다. 어느 쪽이든 동시편집
      // WebSocket(src/routes/edit.ts)은 docx/xlsx가 아니면 canEdit 값과 무관하게 접속 자체를 거부해
      // canEdit이 마크업 용도로 켜진 문서가 잘못 동시편집 룸에 들어가 원본을 덮어쓰는 일을 막는다.
      const effectiveCanEdit = canEdit && (!isEditableExt || findConverter(meta.ext) !== null);
      const token = issueViewToken(
        {
          docId: meta.id,
          viewer,
          perms: { canPrint, canDownload, canCopy, canEdit: effectiveCanEdit },
        },
        config.tokenTtlSeconds,
      );
      await recordAudit({
        type: 'token_issued',
        docId: meta.id,
        viewer,
        detail: `print=${canPrint} download=${canDownload} copy=${canCopy} edit=${effectiveCanEdit}${callerProject ? ` caller=${callerProject}` : ''}`,
      });
      // 그리드(xlsx) 편집기는 canEdit:false여도 읽기 전용으로 그대로 열 수 있다(collab/room.ts가
      // 그 연결의 실제 쓰기를 서버에서 막는다) — 과거 연도 접속현황처럼 "현재 연도와 같은 화면인데
      // 손은 못 대는" 문서가 이 편집기를 그대로 재사용할 수 있게, xlsx는 편집 권한과 무관하게 항상
      // editorUrl을 내려준다. docx는 editor.ts가 아직 읽기 전용을 지원하지 않아 기존대로 둔다.
      // (isEditableExt로 한 번 더 감싸는 이유: effectiveCanEdit이 이제 pdf 등에서도 true일 수 있는데,
      // 그런 문서는 editor.html/editor-grid.html이 아니라 markup 기능이 있는 viewer.html로만 열려야 한다.)
      const isGrid = meta.ext === 'xlsx';
      return {
        success: true,
        data: {
          token,
          viewerUrl: `/viewer.html?token=${encodeURIComponent(token)}`,
          editorUrl: isEditableExt && (effectiveCanEdit || isGrid)
            ? `${editorPathFor(meta.ext)}?token=${encodeURIComponent(token)}`
            : undefined,
          expiresInSeconds: config.tokenTtlSeconds,
        },
      };
    },
  );
}
