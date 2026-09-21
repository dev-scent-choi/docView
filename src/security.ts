import { randomBytes } from 'node:crypto';
import jwt from 'jsonwebtoken';
import type { Permissions, ViewClaims } from './types.js';

const secretFromEnv = process.env.DOCVIEW_SECRET;

/** 환경변수 미설정 시 부팅마다 새 시크릿이 생성되어 기존 토큰이 모두 무효화된다. */
export const isEphemeralSecret = !secretFromEnv;

const jwtSecret = secretFromEnv ?? randomBytes(32).toString('hex');

export class InvalidTokenError extends Error {
  constructor(message = '만료되었거나 유효하지 않은 열람 토큰입니다.') {
    super(message);
    this.name = 'InvalidTokenError';
  }
}

/** 서명·검증 양쪽에 알고리즘을 명시로 고정한다. jsonwebtoken 9.x는 alg:"none"(서명 없음)
 * 토큰을 옵션 없이도 이미 기본으로 거부하고, 지금 이 앱은 대칭키(HS256) 하나만 쓰기 때문에
 * RS256/HS256를 오가며 공개키를 대칭키로 오인시키는 전형적인 algorithm-confusion 공격도 적용될
 * 여지가 없다 — 즉 지금 당장 이 고정이 막아주는 실제 구멍은 없다. 그래도 향후 비대칭키 로그인
 * 등을 추가하며 verify 호출부가 바뀔 때를 대비해 알고리즘을 명시로 강제해 두는 값싼 보험이다. */
const JWT_ALGORITHM = 'HS256';

export function issueViewToken(claims: ViewClaims, ttlSeconds: number): string {
  return jwt.sign(
    { docId: claims.docId, viewer: claims.viewer, perms: claims.perms },
    jwtSecret,
    { expiresIn: ttlSeconds, algorithm: JWT_ALGORITHM },
  );
}

export function verifyViewToken(token: string): ViewClaims {
  let payload: unknown;
  try {
    payload = jwt.verify(token, jwtSecret, { algorithms: [JWT_ALGORITHM] });
  } catch {
    throw new InvalidTokenError();
  }
  if (typeof payload !== 'object' || payload === null) {
    throw new InvalidTokenError();
  }
  const { docId, viewer, perms } = payload as Record<string, unknown>;
  if (typeof docId !== 'string' || typeof viewer !== 'string' || typeof perms !== 'object' || perms === null) {
    throw new InvalidTokenError();
  }
  const raw = perms as Record<string, unknown>;
  const permissions: Permissions = {
    canPrint: raw.canPrint === true,
    canDownload: raw.canDownload === true,
    canCopy: raw.canCopy === true,
    canEdit: raw.canEdit === true,
  };
  return { docId, viewer, perms: permissions };
}
