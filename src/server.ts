import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { isApiKeyAuthEnabled } from './apiKeyAuth.js';
import { buildApp } from './app.js';
import { config } from './config.js';
import { supportedExtensions } from './converters/index.js';
import { sofficePath, warmUpProfiles } from './converters/libreOffice.js';
import { isEphemeralSecret } from './security.js';
import { ensureDataDirs } from './storage.js';
import { sweepOrphanedLazyGridFiles } from './xlsx/lazyGridStore.js';

/** public/rhwp/는 git에 커밋하지 않고 scripts/mirror-rhwp-studio.mjs로 받는다(tools/의 DWG
 * 변환기 바이너리와 같은 이유 — README "한글(hwp/hwpx) 편집" 참고). 없으면 hwp/hwpx의 PDF
 * 전용 열람(LibreOffice 경로)은 그대로 동작하지만 편집기(editor-hwp.html)는 로드에 실패한다. */
const rhwpStudioMirrored = existsSync(
  path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'public', 'rhwp', 'index.html'),
);

async function main(): Promise<void> {
  await ensureDataDirs();
  await sweepOrphanedLazyGridFiles();
  const app = await buildApp();
  if (isEphemeralSecret) {
    app.log.warn(
      'DOCVIEW_SECRET 환경변수가 없어 임시 시크릿을 생성했습니다. 서버 재시작 시 발급된 열람 토큰이 모두 무효화됩니다.',
    );
  }
  if (!isApiKeyAuthEnabled) {
    app.log.warn(
      'DOCVIEW_API_KEYS가 설정되지 않아 업로드/토큰 발급 API가 인증 없이 열려 있습니다. 운영 환경에서는 반드시 설정하세요.',
    );
  }
  if (config.allowedOrigins.length === 0) {
    app.log.warn(
      'DOCVIEW_ALLOWED_ORIGINS가 설정되지 않아 이 화면을 아무 출처에서나 iframe으로 띄울 수 있습니다. 운영 환경에서는 반드시 설정하세요.',
    );
  }
  if (!sofficePath) {
    app.log.warn('LibreOffice가 없어 Office/DXF 변환이 비활성화되었습니다. (DOCVIEW_SOFFICE로 경로 지정 가능)');
  }
  if (!rhwpStudioMirrored) {
    app.log.warn(
      'public/rhwp/가 없어 한글(hwp/hwpx) 편집기가 열리지 않습니다. node scripts/mirror-rhwp-studio.mjs로 받으세요.',
    );
  }
  app.log.info(
    config.adminPassword
      ? `관리자 페이지(/admin) 켜짐 — 허용 IP: ${[...config.adminAllowedIps].join(', ')}`
      : '관리자 페이지(/admin) 꺼짐 (DOCVIEW_ADMIN_PASSWORD 미설정)',
  );
  app.log.info(`저장 모드: ${config.storageMode} / 지원 포맷: ${supportedExtensions.join(', ')}`);
  await app.listen({ port: config.port, host: config.host });
  // 첫 변환의 콜드 스타트를 없애기 위해 백그라운드로 프로필을 웜업한다
  void warmUpProfiles().then(() => app.log.info('LibreOffice 프로필 웜업 완료'));
}

main().catch((error) => {
  console.error('서버 시작 실패:', error);
  process.exit(1);
});
