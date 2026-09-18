import { timingSafeEqual } from 'node:crypto';
import type { FastifyReply, FastifyRequest } from 'fastify';
import { config } from './config.js';

/** DOCVIEW_ADMIN_PASSWORD가 비어 있으면 관리자 페이지 자체가 꺼져 있는 것이다. */
export const isAdminEnabled = config.adminPassword !== null;

/** 길이가 다르면 timingSafeEqual 자체가 예외를 던지므로, 길이 비교부터 안전하게 처리한다
 * (길이가 다르다는 사실만으로 타이밍 차이가 나는 것은 애초에 값 추측에 쓸모없는 정보라 무해하다 —
 * 위험한 건 "같은 길이 안에서 몇 바이트째에 틀렸는지"가 새는 것뿐이다). */
function safeEquals(a: string, b: string): boolean {
  const bufA = Buffer.from(a);
  const bufB = Buffer.from(b);
  if (bufA.length !== bufB.length) return false;
  return timingSafeEqual(bufA, bufB);
}

/** 프록시 없이 이 프로세스가 직접 0.0.0.0에 바인딩되는 배포 형태를 전제로 request.ip(소켓의 실제
 * 원격 주소)를 그대로 신뢰한다 — 리버스 프록시를 앞에 두게 되면 Fastify trustProxy 설정과 함께
 * 다시 검토해야 한다(그 경우 클라이언트가 X-Forwarded-For를 위조해 IP 검사를 우회할 수 있다). */
function isIpAllowed(request: FastifyRequest): boolean {
  return config.adminAllowedIps.has(request.ip);
}

/**
 * `/admin`, `/api/admin/*`에 붙이는 preHandler. 두 단계로 막는다:
 * 1. IP 허용 목록에 없으면 404(관리자 페이지가 있다는 사실 자체를 감춘다 — 403은 "여기 뭔가
 *    있는데 너는 못 본다"는 정보를 흘린다).
 * 2. IP는 통과했지만 HTTP Basic Auth 자격증명이 없거나 틀리면 401 + WWW-Authenticate 헤더로
 *    브라우저가 기본 로그인 팝업을 띄우게 한다.
 *
 * DOCVIEW_ADMIN_PASSWORD가 아예 비어 있으면(미설정) 첫 단계 이전에 404 — "설정 안 하면 페이지
 * 자체가 없는 것"이 기본값이다(실수로 보호 없이 열리는 것을 막기 위함).
 */
export async function requireAdminAuth(request: FastifyRequest, reply: FastifyReply): Promise<void> {
  if (!isAdminEnabled) {
    await reply.code(404).send({ success: false, error: 'Not Found' });
    return;
  }
  if (!isIpAllowed(request)) {
    await reply.code(404).send({ success: false, error: 'Not Found' });
    return;
  }

  const header = request.headers.authorization;
  const base64 = header?.startsWith('Basic ') ? header.slice('Basic '.length).trim() : undefined;
  const decoded = base64 ? Buffer.from(base64, 'base64').toString('utf-8') : '';
  const providedPassword = decoded.includes(':') ? decoded.slice(decoded.indexOf(':') + 1) : '';

  if (!config.adminPassword || !safeEquals(providedPassword, config.adminPassword)) {
    await reply
      .code(401)
      .header('WWW-Authenticate', 'Basic realm="DocView Admin"')
      .send({ success: false, error: '인증이 필요합니다.' });
  }
}
