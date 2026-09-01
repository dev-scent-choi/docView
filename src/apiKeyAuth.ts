import type { FastifyReply, FastifyRequest } from 'fastify';
import { config } from './config.js';

/** DOCVIEW_API_KEYS가 하나도 설정되지 않았으면(개발 환경 등) 이 인증 자체가 꺼져 있는 것이다. */
export const isApiKeyAuthEnabled = config.apiKeys.size > 0;

/** 유효한 키에 매핑된 호출자 이름을 request에 심어 감사 로그 등에서 "누가 호출했는지" 남길 수
 * 있게 한다 — 이 API를 호출하는 쪽은 어차피 서버(브라우저 아님)이므로 이름 자체는 신뢰 정보다. */
export interface ApiKeyAuthenticatedRequest extends FastifyRequest {
  callerProject?: string;
}

/**
 * 업로드(`POST /api/documents`)·토큰 발급(`POST /api/documents/:id/tokens`)처럼 아무나 호출하면
 * 안 되는 관리용 API에 붙이는 preHandler. `Authorization: Bearer <키>` 헤더를 요구한다.
 *
 * DOCVIEW_API_KEYS가 비어 있으면(미설정) 검사 자체를 건너뛴다 — 로컬 개발 편의를 위해서다.
 * 부팅 시 한 번 경고를 남기므로(server.ts), 운영 배포 시 설정을 빠뜨려도 로그에서 확인할 수 있다.
 */
export async function requireApiKey(request: FastifyRequest, reply: FastifyReply): Promise<void> {
  if (!isApiKeyAuthEnabled) return;

  const header = request.headers.authorization;
  const key = header?.startsWith('Bearer ') ? header.slice('Bearer '.length).trim() : undefined;
  const caller = key ? config.apiKeys.get(key) : undefined;

  if (!caller) {
    await reply.code(401).send({ success: false, error: 'API 키가 없거나 유효하지 않습니다.' });
    return;
  }

  (request as ApiKeyAuthenticatedRequest).callerProject = caller;
}
