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

export function issueViewToken(claims: ViewClaims, ttlSeconds: number): string {
  return jwt.sign(
    { docId: claims.docId, viewer: claims.viewer, perms: claims.perms },
    jwtSecret,
    { expiresIn: ttlSeconds },
  );
}

export function verifyViewToken(token: string): ViewClaims {
  let payload: unknown;
  try {
    payload = jwt.verify(token, jwtSecret);
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
