import fs from 'node:fs/promises';
import path from 'node:path';
import type { FastifyInstance } from 'fastify';
import { requireAdminAuth } from '../adminAuth.js';
import { isApiKeyAuthEnabled } from '../apiKeyAuth.js';
import { config } from '../config.js';
import { dwg2dxfPath, odaConverterPath } from '../converters/dwgConverter.js';
import { detectH2Orestart, sofficePath } from '../converters/libreOffice.js';
import { store } from '../storage.js';

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[ch]!);
}

function yesNo(value: boolean, yes = '설정됨', no = '미설정'): string {
  return value ? yes : no;
}

/** originalsDir/convertedDir 안 파일 크기 합(바이트) — 둘 다 1단계 평면 구조(하위 폴더 없음)라
 * 재귀 없이 readdir + stat만으로 충분하다. 디렉터리 자체가 없거나(memory 모드) 읽기 실패하면
 * 0으로 취급한다(관리자 페이지가 그 이유로 통째로 에러 화면이 되면 안 된다). */
async function dirSizeBytes(dir: string): Promise<number> {
  try {
    const entries = await fs.readdir(dir);
    let total = 0;
    for (const entry of entries) {
      const stat = await fs.stat(path.join(dir, entry)).catch(() => null);
      if (stat?.isFile()) total += stat.size;
    }
    return total;
  } catch {
    return 0;
  }
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  const units = ['KB', 'MB', 'GB', 'TB'];
  let value = bytes / 1024;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value.toFixed(1)} ${units[unitIndex]}`;
}

function statusRow(label: string, value: string): string {
  return `<tr><th>${escapeHtml(label)}</th><td>${value}</td></tr>`;
}

export async function adminRoutes(app: FastifyInstance): Promise<void> {
  app.get('/admin', { preHandler: requireAdminAuth }, async (_request, reply) => {
    const docs = await store.listMeta();
    const originalsBytes = config.storageMode === 'disk' ? await dirSizeBytes(config.originalsDir) : 0;
    const convertedBytes = config.storageMode === 'disk' ? await dirSizeBytes(config.convertedDir) : 0;
    const h2orestart = await detectH2Orestart();
    const h2orestartText =
      h2orestart === null ? '확인 불가' : yesNo(h2orestart, '설치됨', '미설치(또는 확인 안 됨)');

    const apiKeyProjects = [...new Set(config.apiKeys.values())];

    const html = `<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>DocView — 관리자</title>
<style>
  body { font-family: -apple-system, "Malgun Gothic", sans-serif; background: #f3f4f7; color: #1b1b1b; margin: 0; padding: 2rem; }
  h1 { font-size: 1.3rem; margin: 0 0 0.25rem; }
  .sub { color: #666; font-size: 0.85rem; margin-bottom: 1.5rem; }
  section { background: #fff; border: 1px solid #e0e0e0; border-radius: 8px; padding: 1.25rem 1.5rem; margin-bottom: 1.25rem; }
  h2 { font-size: 1rem; margin: 0 0 0.75rem; color: #333; }
  table { width: 100%; border-collapse: collapse; font-size: 0.9rem; }
  th { text-align: left; width: 40%; padding: 0.4rem 0.5rem 0.4rem 0; color: #555; font-weight: 500; vertical-align: top; }
  td { padding: 0.4rem 0; }
  .ok { color: #1a7f37; font-weight: 600; }
  .warn { color: #b35900; font-weight: 600; }
  .bad { color: #c0392b; font-weight: 600; }
  .muted { color: #888; }
  code { background: #f0f0f0; padding: 0.1rem 0.35rem; border-radius: 4px; font-size: 0.85em; }
</style>
</head>
<body>
  <h1>DocView 관리자 — 환경설정 확인</h1>
  <div class="sub">읽기 전용입니다. 값을 바꾸려면 서버의 환경변수를 고치고 재시작하세요(README 참고).</div>

  <section>
    <h2>서버</h2>
    <table>
      ${statusRow('바인딩 주소', `<code>${escapeHtml(config.host)}:${config.port}</code>`)}
      ${statusRow('저장 모드', config.storageMode === 'disk' ? '디스크(재시작해도 유지)' : '<span class="warn">메모리(재시작 시 전부 소멸)</span>')}
      ${config.storageMode === 'disk' ? statusRow('데이터 폴더', `<code>${escapeHtml(config.dataRoot)}</code>`) : ''}
      ${statusRow('토큰 유효기간', `${config.tokenTtlSeconds}초`)}
    </table>
  </section>

  <section>
    <h2>보안</h2>
    <table>
      ${statusRow(
        '업로드/토큰 API 키 인증',
        isApiKeyAuthEnabled
          ? `<span class="ok">켜짐</span> — 등록된 프로젝트: ${apiKeyProjects.map((p) => `<code>${escapeHtml(p)}</code>`).join(', ')}`
          : '<span class="bad">꺼짐 — 누구나 업로드/토큰 발급 API를 호출할 수 있습니다</span>',
      )}
      ${statusRow(
        'iframe 임베드 출처 제한',
        config.allowedOrigins.length > 0
          ? `<span class="ok">켜짐</span> — ${config.allowedOrigins.map((o) => `<code>${escapeHtml(o)}</code>`).join(', ')}`
          : '<span class="bad">꺼짐 — 아무 사이트에서나 이 화면을 iframe으로 띄울 수 있습니다</span>',
      )}
      ${statusRow('관리자 페이지 허용 IP', [...config.adminAllowedIps].map((ip) => `<code>${escapeHtml(ip)}</code>`).join(', '))}
    </table>
  </section>

  <section>
    <h2>문서 변환기</h2>
    <table>
      ${statusRow('LibreOffice', sofficePath ? `<span class="ok">감지됨</span> — <code>${escapeHtml(sofficePath)}</code>` : '<span class="bad">미감지</span> — docx/pptx/hwp 등 대부분의 변환이 동작하지 않습니다')}
      ${statusRow('H2Orestart(HWP 5.0/HWPX 임포트)', h2orestartText)}
      ${statusRow(
        'DWG → DXF 변환기',
        dwg2dxfPath || odaConverterPath
          ? `<span class="ok">감지됨</span> — ${[dwg2dxfPath ? 'LibreDWG' : null, odaConverterPath ? 'ODA File Converter' : null].filter(Boolean).join(', ')}`
          : '<span class="muted">미감지 — dwg 업로드가 거부됩니다(dxf는 영향 없음)</span>',
      )}
    </table>
  </section>

  <section>
    <h2>저장 용량</h2>
    <table>
      ${statusRow('업로드된 문서 수', `${docs.length}개`)}
      ${
        config.storageMode === 'disk'
          ? statusRow('디스크 사용량', `원본 ${formatBytes(originalsBytes)} + 변환본 ${formatBytes(convertedBytes)} = ${formatBytes(originalsBytes + convertedBytes)}`)
          : statusRow('디스크 사용량', '<span class="muted">해당 없음(메모리 모드)</span>')
      }
    </table>
  </section>
</body>
</html>`;

    return await reply.type('text/html; charset=utf-8').send(html);
  });
}
