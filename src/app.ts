import path from 'node:path';
import { fileURLToPath } from 'node:url';
import multipart from '@fastify/multipart';
import fastifyStatic from '@fastify/static';
import fastifyWebsocket from '@fastify/websocket';
import Fastify, { type FastifyInstance } from 'fastify';
import { config } from './config.js';
import { editRoutes } from './routes/edit.js';
import { documentRoutes } from './routes/documents.js';
import { viewerRoutes } from './routes/viewer.js';

const publicDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'public');

export async function buildApp(): Promise<FastifyInstance> {
  const app = Fastify({
    logger: { level: process.env.LOG_LEVEL ?? 'info' },
    bodyLimit: config.maxUploadBytes,
  });

  // 이 화면(뷰어/편집기)을 iframe으로 띄울 수 있는 출처를 제한한다 — 설정 안 하면(개발 편의)
  // 아무 출처에서나 임베드할 수 있으니, 운영 환경에서는 DOCVIEW_ALLOWED_ORIGINS를 반드시 채워야
  // 한다(server.ts가 부팅 시 경고). JSON 응답 등 프레임과 무관한 응답에 붙어도 무해하므로
  // 라우트를 가리지 않고 전역으로 건다.
  if (config.allowedOrigins.length > 0) {
    const frameAncestors = `frame-ancestors ${config.allowedOrigins.join(' ')}`;
    app.addHook('onSend', async (_request, reply, payload) => {
      reply.header('Content-Security-Policy', frameAncestors);
      return payload;
    });
  }

  await app.register(multipart, {
    limits: { fileSize: config.maxUploadBytes, files: 1 },
  });
  await app.register(fastifyWebsocket);
  await app.register(fastifyStatic, {
    root: publicDir,
    // 뷰어/관리 화면 JS가 브라우저에 캐시되어 구버전이 돌지 않도록 항상 재검증시킨다.
    setHeaders(res) {
      res.setHeader('cache-control', 'no-cache');
    },
  });
  await app.register(documentRoutes);
  await app.register(viewerRoutes);
  await app.register(editRoutes);

  return app;
}
