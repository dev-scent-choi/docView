import { buildApp } from './app.js';
import { config } from './config.js';
import { supportedExtensions } from './converters/index.js';
import { sofficePath, warmUpProfiles } from './converters/libreOffice.js';
import { isEphemeralSecret } from './security.js';
import { ensureDataDirs } from './storage.js';

async function main(): Promise<void> {
  await ensureDataDirs();
  const app = await buildApp();
  if (isEphemeralSecret) {
    app.log.warn(
      'DOCVIEW_SECRET 환경변수가 없어 임시 시크릿을 생성했습니다. 서버 재시작 시 발급된 열람 토큰이 모두 무효화됩니다.',
    );
  }
  if (!sofficePath) {
    app.log.warn('LibreOffice가 없어 Office/DXF 변환이 비활성화되었습니다. (DOCVIEW_SOFFICE로 경로 지정 가능)');
  }
  app.log.info(`저장 모드: ${config.storageMode} / 지원 포맷: ${supportedExtensions.join(', ')}`);
  await app.listen({ port: config.port, host: config.host });
  // 첫 변환의 콜드 스타트를 없애기 위해 백그라운드로 프로필을 웜업한다
  void warmUpProfiles().then(() => app.log.info('LibreOffice 프로필 웜업 완료'));
}

main().catch((error) => {
  console.error('서버 시작 실패:', error);
  process.exit(1);
});
