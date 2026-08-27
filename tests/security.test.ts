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
});
