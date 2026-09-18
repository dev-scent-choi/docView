import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import type { FastifyInstance } from 'fastify';
import { requireApiKey, type ApiKeyAuthenticatedRequest } from '../apiKeyAuth.js';
import { recordAudit } from '../audit.js';
import { forceCloseRoom } from '../collab/room.js';
import { config } from '../config.js';
import { convertDwgToDxf } from '../converters/dwgConverter.js';
import { findConverter, supportedExtensions } from '../converters/index.js';
import { convertDxfToSvg, convertLegacySpreadsheetToXlsx } from '../converters/libreOffice.js';
import { ConversionError, type Converter } from '../converters/types.js';
import { clearPresence } from '../presence.js';
import { getPdfPageInfo, invalidateRenderCache } from '../rendering.js';
import { issueViewToken } from '../security.js';
import { buildGridFromOriginal, saveGridDerived, type GridDocument } from '../spreadsheet.js';
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

/** 실시간 동시편집(에디터)을 지원하는 확장자 — docx는 문서 편집기, xlsx는 그리드 편집기,
 * pptx는 슬라이드 구성 편집기로 연결된다. hwp/hwpx는 다른 셋과 달리 Yjs 동시편집이 아니라
 * @rhwp/editor(단일 사용자 편집, 명시적 저장) 기반 전용 편집기로 연결된다 — 원본 .hwp/.hwpx
 * 확장자·바이트를 그대로 유지한 채 편집하기 위해서다(README 참고). */
export const EDITABLE_EXTS = new Set(['docx', 'xlsx', 'pptx', 'hwp', 'hwpx']);
export function editorPathFor(ext: string): string {
  if (ext === 'xlsx') return '/editor-grid.html';
  if (ext === 'pptx') return '/editor-slides.html';
  if (ext === 'hwp' || ext === 'hwpx') return '/editor-hwp.html';
  return '/editor.html';
}

async function buildGrid(meta: DocumentMeta): Promise<GridDocument> {
  return buildGridFromOriginal(store, meta);
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
    await saveGridDerived(store, meta.id, grid);
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
    const { pageCount, sizes } = await getPdfPageInfo(pdfBytes);
    await store.saveDerived(meta.id, 'pagesizes', Buffer.from(JSON.stringify(sizes)));
    return { pdfBytes, pageCount };
  } finally {
    // recursive:true일 때만 적용되는 옵션 — Windows에서 백신 등이 방금 만든 임시 파일을 잠깐
    // 스캔하는 동안 삭제가 EBUSY로 실패하는 경우가 있어, 짧게 재시도한다.
    await fs.rm(workDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}

/** 도면(dxf/dwg) 전용 부가 산출물 — 확대해도 흐려지지 않는 벡터 뷰를 위해 PDF와 별개로
 * SVG도 만든다. PDF 변환은 이미 끝난 뒤라 이게 실패해도(LibreOffice DXF→SVG 필터가 특정
 * 도면을 못 그리거나 등) 문서 열람 자체는 막히지 않는다 — 그냥 기존 pages(PDF) 뷰로 남는다. */
const VECTOR_SVG_EXTS = new Set(['dxf', 'dwg']);

async function tryGenerateVectorSvg(meta: DocumentMeta): Promise<boolean> {
  const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-svg-'));
  try {
    const sourcePath = path.join(workDir, `source.${meta.ext}`);
    await fs.writeFile(sourcePath, await store.readOriginal(meta.id, meta.ext));
    const dxfPath = meta.ext === 'dwg' ? await convertDwgToDxf(sourcePath, workDir) : sourcePath;
    const targetSvgPath = path.join(workDir, 'out.svg');
    await convertDxfToSvg(dxfPath, targetSvgPath);
    await store.saveDerived(meta.id, 'svg', await fs.readFile(targetSvgPath));
    return true;
  } catch (error) {
    console.warn(`도면 SVG 생성 실패 (PDF 뷰로 폴백, docId=${meta.id}):`, error);
    return false;
  } finally {
    await fs.rm(workDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}

/** 업로드 핸들러가 findConverter를 부르기 전에 쓴다 — 성공하면 xls/ods 업로드를 xlsx가 올라온
 * 것과 완전히 동일하게 취급해(ext까지 바꿔치기) 그리드 뷰·동시편집이 그대로 켜지게 한다. 실패하면
 * (LibreOffice 미설치, 손상된 파일 등) null을 돌려주고 호출부가 원래 ext로 계속 진행한다 —
 * xls/ods는 여전히 spreadsheetConverter(PDF 전용 열람)로 폴백 가능하다(그 경로도 LibreOffice가
 * 필요해, LibreOffice가 아예 없으면 거기서도 결국 415로 거부된다 — 일관된 결과). */
async function tryNormalizeLegacySpreadsheetToXlsx(
  data: Buffer,
  ext: 'xls' | 'ods',
): Promise<Buffer | null> {
  const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-legacyxlsx-'));
  try {
    const sourcePath = path.join(workDir, `source.${ext}`);
    await fs.writeFile(sourcePath, data);
    const targetXlsxPath = path.join(workDir, 'out.xlsx');
    await convertLegacySpreadsheetToXlsx(sourcePath, ext, targetXlsxPath);
    return await fs.readFile(targetXlsxPath);
  } catch (error) {
    console.warn(`.${ext} → xlsx 정규화 실패 (PDF 전용 열람으로 폴백):`, error);
    return null;
  } finally {
    await fs.rm(workDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}

async function runConversion(app: FastifyInstance, converter: Converter, meta: DocumentMeta): Promise<void> {
  try {
    await store.writeMeta({ ...meta, progress: PROGRESS_CONVERTING });
    const { pageCount } = await convertToPdf(meta, converter);
    await store.writeMeta({ ...meta, progress: PROGRESS_PDF_DONE });
    const viewKind =
      VECTOR_SVG_EXTS.has(meta.ext) && (await tryGenerateVectorSvg(meta)) ? 'svg' : 'pages';
    await store.writeMeta({ ...meta, status: 'ready', viewKind, pageCount, progress: 100 });
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

interface UploadResult {
  readonly code: number;
  readonly body: { readonly success: boolean; readonly data?: DocumentMeta; readonly error?: string };
}

/** 업로드 진입점 두 곳(multipart 폼 필드 / 원시 바이트 스트림)이 공유하는 실제 처리 로직 —
 * 파일명과 바이트만 받으면 나머지(확장자 판별, xls/ods 정규화, 변환기 조회, 저장,
 * 백그라운드 변환 착수)는 완전히 동일하다. 반환값을 라우트 핸들러가 그대로 reply.code(...).send(...)에
 * 싣는다. */
async function ingestUpload(
  app: FastifyInstance,
  originalFilename: string,
  data: Buffer,
  callerProject: string | undefined,
): Promise<UploadResult> {
  let ext = path.extname(originalFilename).slice(1).toLowerCase();
  let displayName = originalFilename;
  // xls/ods(레거시 바이너리·ODF 스프레드시트)는 이 서비스의 그리드 뷰·동시편집 엔진이 전제하는
  // xlsx 구조가 아니라서, 지금까지는 PDF 전용 열람만 가능했고 편집은 아예 지원되지 않았다.
  // 여기서 xlsx로 정규화해 두면 이후 로직 전체(findConverter부터 GRID_VIEW_EXTS 분기까지)가
  // 원래 xlsx가 올라온 것처럼 그대로 동작해 그리드 뷰·동시편집이 자동으로 켜진다 — 실패하면
  // (LibreOffice 미설치, 손상된 파일 등) 원래 ext로 계속 진행해 기존 PDF 전용 열람으로 폴백한다.
  if (ext === 'xls' || ext === 'ods') {
    const converted = await tryNormalizeLegacySpreadsheetToXlsx(data, ext);
    if (converted) {
      data = converted;
      // 표시 이름·다운로드 파일명(meta.name)의 확장자도 같이 바꾼다 — 안 바꾸면 실제로는
      // xlsx 바이트인데 파일명만 "report.xls"로 남아, 다운로드한 사용자가 확장자와 실제 내용이
      // 어긋난 파일을 받게 된다(뷰어 화면의 route.ts 다운로드 라우트가 meta.name을 그대로
      // Content-Disposition 파일명으로 쓴다).
      displayName = `${displayName.replace(/\.[^./]+$/, '')}.xlsx`;
      ext = 'xlsx';
    }
  }
  const converter = findConverter(ext);
  if (!converter) {
    return {
      code: 415,
      body: {
        success: false,
        error: `지원하지 않는 형식입니다: .${ext || '(없음)'} — 지원 형식: ${supportedExtensions.join(', ')}`,
      },
    };
  }

  const id = newDocumentId();
  try {
    await store.saveOriginal(id, ext, data);
    const meta: DocumentMeta = {
      id,
      name: displayName,
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
      detail: callerProject ? `${originalFilename} (caller=${callerProject})` : originalFilename,
    });
    // 변환은 백그라운드로 진행하고 즉시 응답한다. 클라이언트는 상태를 폴링한다.
    if (GRID_VIEW_EXTS.has(ext)) {
      void runGridExtraction(app, meta);
    } else {
      void runConversion(app, converter, meta);
    }
    return { code: 201, body: { success: true, data: meta } };
  } catch (error) {
    app.log.error(error, 'upload persisting failed');
    return { code: 500, body: { success: false, error: '업로드 저장에 실패했습니다.' } };
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

    const result = await ingestUpload(app, file.filename, data, callerProject);
    return await reply.code(result.code).send(result.body);
  });

  app.post<{ Querystring: { filename: string } }>(
    '/api/documents/stream',
    {
      preHandler: requireApiKey,
      schema: {
        querystring: {
          type: 'object',
          required: ['filename'],
          properties: { filename: { type: 'string', minLength: 1, maxLength: 300 } },
        },
      },
    },
    async (request, reply) => {
      const callerProject = (request as ApiKeyAuthenticatedRequest).callerProject;
      // Content-Type: application/octet-stream 요청만 여기 닿는다(app.ts에 등록한 원시 바이트
      // 파서). multipart 경계를 만들 필요 없이 파일 바이트를 요청 본문에 그대로 실어 보내는
      // 쪽을 위한 경로 — 같은 서버에 파일이 이미 있는 다른 프로젝트가 그 바이트를 파이프로
      // 그대로 흘려보낼 때 특히 유용하다(디스크에 다시 쓰지 않고 스트리밍).
      const data = request.body as Buffer;
      if (!data || data.length === 0) {
        return reply.code(400).send({ success: false, error: '업로드된 파일이 없습니다.' });
      }
      const result = await ingestUpload(app, request.query.filename, data, callerProject);
      return await reply.code(result.code).send(result.body);
    },
  );

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
      // 반드시 삭제 뒤에 부른다 — forceCloseRoom 주석 참고(순서가 바뀌면 삭제 직후 파일이
      // 자동저장으로 되살아날 수 있다).
      forceCloseRoom(meta.id);
      invalidateRenderCache(meta.id);
      invalidateViewCaches(meta.id);
      // 삭제된 문서는 다시 열릴 수 없어(뷰어를 열 방법이 없음) presence가 스스로 정리될 기회가
      // 없다 — 명시적으로 지운다(presence.ts의 clearPresence 문서 주석 참고).
      clearPresence(meta.id);
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
      // 문서(docx)·그리드(xlsx)·슬라이드(pptx) 편집기 전부 canEdit:false여도 읽기 전용으로 그대로
      // 열 수 있다(collab/room.ts가 그 연결의 실제 쓰기를 서버에서 막는다) — 과거 연도 접속현황처럼
      // "현재 연도와 같은 화면인데 손은 못 대는" 문서가 이 편집기를 그대로 재사용할 수 있게, 편집
      // 권한과 무관하게 항상 editorUrl을 내려준다.
      // (isEditableExt로 한 번 더 감싸는 이유: effectiveCanEdit이 이제 pdf 등에서도 true일 수 있는데,
      // 그런 문서는 editor.html/editor-grid.html이 아니라 markup 기능이 있는 viewer.html로만 열려야 한다.)
      const supportsReadOnlyEditor = meta.ext === 'docx' || meta.ext === 'xlsx' || meta.ext === 'pptx';
      return {
        success: true,
        data: {
          token,
          viewerUrl: `/viewer.html?token=${encodeURIComponent(token)}`,
          editorUrl: isEditableExt && (effectiveCanEdit || supportsReadOnlyEditor)
            ? `${editorPathFor(meta.ext)}?token=${encodeURIComponent(token)}`
            : undefined,
          expiresInSeconds: config.tokenTtlSeconds,
        },
      };
    },
  );
}
