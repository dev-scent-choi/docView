import path from 'node:path';

const dataRoot = process.env.DOCVIEW_DATA_DIR ?? path.resolve(process.cwd(), 'data');

const DEFAULT_PORT = 8000;
const DEFAULT_TOKEN_TTL_SECONDS = 60 * 60;
const MAX_UPLOAD_BYTES = 100 * 1024 * 1024;
const RENDER_SCALE = 2;
const DEFAULT_MEMORY_BUDGET_MB = 1024;
const DEFAULT_CACHE_BUDGET_MB = 512;

export type StorageMode = 'disk' | 'memory';

const storageMode: StorageMode = process.env.DOCVIEW_STORAGE === 'memory' ? 'memory' : 'disk';

export const config = {
  port: Number(process.env.DOCVIEW_PORT ?? DEFAULT_PORT),
  host: process.env.DOCVIEW_HOST ?? '127.0.0.1',
  dataRoot,
  originalsDir: path.join(dataRoot, 'originals'),
  convertedDir: path.join(dataRoot, 'converted'),
  metaDir: path.join(dataRoot, 'meta'),
  auditDir: path.join(dataRoot, 'audit'),
  tokenTtlSeconds: Number(process.env.DOCVIEW_TOKEN_TTL ?? DEFAULT_TOKEN_TTL_SECONDS),
  maxUploadBytes: MAX_UPLOAD_BYTES,
  renderScale: RENDER_SCALE,
  /** disk: 파일 저장(기본) / memory: 서버 메모리에만 보관 (재시작 시 소멸, 디스크 비잔류) */
  storageMode,
  memoryBudgetBytes:
    Number(process.env.DOCVIEW_MEMORY_BUDGET_MB ?? DEFAULT_MEMORY_BUDGET_MB) * 1024 * 1024,
  /** 렌더링·파싱 캐시(그리드/PDF 텍스트/페이지 이미지)가 공유하는 총 메모리 상한 */
  cacheBudgetBytes:
    Number(process.env.DOCVIEW_CACHE_BUDGET_MB ?? DEFAULT_CACHE_BUDGET_MB) * 1024 * 1024,
} as const;
