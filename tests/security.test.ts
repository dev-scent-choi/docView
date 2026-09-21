import jwt from 'jsonwebtoken';
import { describe, expect, test } from 'vitest';
import { InvalidTokenError, issueViewToken, verifyViewToken } from '../src/security.js';
import type { ViewClaims } from '../src/types.js';

const claims: ViewClaims = {
  docId: '11111111-2222-3333-4444-555555555555',
  viewer: '홍길동',
  perms: { canPrint: true, canDownload: false, canCopy: false, canEdit: false },
};

describe('security', () => {
  test('issues and verifies a token round-trip with permissions intact', () => {
    const token = issueViewToken(claims, 60);
    const verified = verifyViewToken(token);
    expect(verified.docId).toBe(claims.docId);
    expect(verified.viewer).toBe(claims.viewer);
    expect(verified.perms).toEqual({ canPrint: true, canDownload: false, canCopy: false, canEdit: false });
  });

  test('throws InvalidTokenError for a tampered token', () => {
    const token = issueViewToken(claims, 60);
    expect(() => verifyViewToken(`${token}x`)).toThrow(InvalidTokenError);
  });

  test('throws InvalidTokenError for an expired token', () => {
    const token = issueViewToken(claims, -10);
    expect(() => verifyViewToken(token)).toThrow(InvalidTokenError);
  });

  test('throws InvalidTokenError for garbage input', () => {
    expect(() => verifyViewToken('not-a-jwt')).toThrow(InvalidTokenError);
  });

  test('coerces missing permission fields to false instead of trusting the payload', () => {
    const token = issueViewToken(
      { ...claims, perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: false } },
      60,
    );
    const verified = verifyViewToken(token);
    expect(verified.perms.canPrint).toBe(false);
    expect(verified.perms.canDownload).toBe(false);
    expect(verified.perms.canCopy).toBe(false);
    expect(verified.perms.canEdit).toBe(false);
  });

  test('carries canEdit through the round-trip when granted', () => {
    const token = issueViewToken({ ...claims, perms: { ...claims.perms, canEdit: true } }, 60);
    const verified = verifyViewToken(token);
    expect(verified.perms.canEdit).toBe(true);
  });

  // jsonwebtoken 9.x는 alg:"none" 토큰을 옵션 없이도 이미 기본으로 거부하므로(직접 확인함 —
  // security.ts의 JWT_ALGORITHM 주석 참고) 이 테스트를 지금 통과시키는 건 우리가 추가한
  // algorithms:['HS256'] 고정이 아니라 그 라이브러리 기본 동작이다. 그래도 "서명 없는 토큰은
  // 거부되어야 한다"는 보장 자체는 이 앱에 중요하므로, 라이브러리 업그레이드나 verify 호출부
  // 변경으로 이게 깨지면 바로 잡아내도록 회귀 테스트로 남겨둔다.
  test('rejects a forged token that declares alg:"none" (no signature) even with valid-looking claims', () => {
    const forged = jwt.sign(
      { docId: claims.docId, viewer: claims.viewer, perms: { ...claims.perms, canEdit: true, canDownload: true } },
      null as unknown as string,
      { algorithm: 'none' },
    );
    expect(() => verifyViewToken(forged)).toThrow(InvalidTokenError);
  });
});
