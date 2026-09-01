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

/** "이름:키" 쌍을 콤마로 나열한다(예: "maintenance-management:abc123,other-project:def456").
 * 키 → 호출자 이름 맵으로 뒤집어 둔다 — 같은 키가 두 번 나오면 뒤에 오는 이름이 이긴다. */
function parseApiKeys(raw: string | undefined): ReadonlyMap<string, string> {
  const map = new Map<string, string>();
  for (const pair of (raw ?? '').split(',')) {
    const trimmed = pair.trim();
    if (!trimmed) continue;
    const sep = trimmed.indexOf(':');
    if (sep === -1) continue;
    const name = trimmed.slice(0, sep).trim();
    const key = trimmed.slice(sep + 1).trim();
    if (name && key) map.set(key, name);
  }
  return map;
}

function parseAllowedOrigins(raw: string | undefined): readonly string[] {
  return (raw ?? '')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}

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
  /** 업로드/토큰 발급처럼 "누구나 호출하면 안 되는" 관리용 API를 호출할 수 있는 프로젝트별 키.
   * 비어 있으면(미설정) 이 인증은 꺼진 채로 동작한다(개발 편의 — DOCVIEW_SECRET의 임시 시크릿
   * 폴백과 같은 철학) — 운영 환경에서는 반드시 설정해야 한다. */
  apiKeys: parseApiKeys(process.env.DOCVIEW_API_KEYS),
  /** 이 뷰어 화면을 iframe으로 띄워도 되는 출처 목록. 비어 있으면 프레임 제한을 걸지 않는다
   * (개발 편의 — 운영 환경에서는 반드시 실제로 이 화면을 임베드할 사이트 출처로 설정해야 한다). */
  allowedOrigins: parseAllowedOrigins(process.env.DOCVIEW_ALLOWED_ORIGINS),
} as const;
