import { afterEach, describe, expect, test, vi } from 'vitest';

/** config.ts가 import 시점에 환경변수를 한 번만 읽으므로, adminAuth.ts도 함께 매번
 * 모듈 캐시를 초기화하고 동적으로 다시 import해야 서로 다른 설정으로 테스트할 수 있다. */
async function loadWithEnv(password: string | undefined, allowedIps: string | undefined) {
  vi.resetModules();
  if (password === undefined) delete process.env.DOCVIEW_ADMIN_PASSWORD;
  else process.env.DOCVIEW_ADMIN_PASSWORD = password;
  if (allowedIps === undefined) delete process.env.DOCVIEW_ADMIN_ALLOWED_IPS;
  else process.env.DOCVIEW_ADMIN_ALLOWED_IPS = allowedIps;
  return import('../src/adminAuth.js');
}

function fakeReply() {
  const calls: { code?: number; header?: Record<string, string>; body?: unknown } = { header: {} };
  return {
    code(status: number) {
      calls.code = status;
      return this;
    },
    header(name: string, value: string) {
      calls.header![name] = value;
      return this;
    },
    async send(body: unknown) {
      calls.body = body;
    },
    calls,
  };
}

function fakeRequest(ip: string, authHeader?: string) {
  return { ip, headers: { authorization: authHeader } } as unknown as Parameters<
    typeof import('../src/adminAuth.js').requireAdminAuth
  >[0];
}

function basicAuthHeader(username: string, password: string): string {
  return `Basic ${Buffer.from(`${username}:${password}`).toString('base64')}`;
}

describe('adminAuth', () => {
  const originalPassword = process.env.DOCVIEW_ADMIN_PASSWORD;
  const originalAllowedIps = process.env.DOCVIEW_ADMIN_ALLOWED_IPS;

  afterEach(() => {
    if (originalPassword === undefined) delete process.env.DOCVIEW_ADMIN_PASSWORD;
    else process.env.DOCVIEW_ADMIN_PASSWORD = originalPassword;
    if (originalAllowedIps === undefined) delete process.env.DOCVIEW_ADMIN_ALLOWED_IPS;
    else process.env.DOCVIEW_ADMIN_ALLOWED_IPS = originalAllowedIps;
    vi.resetModules();
  });

  test('when DOCVIEW_ADMIN_PASSWORD is unset, the admin page is disabled and hidden behind a 404', async () => {
    const { requireAdminAuth, isAdminEnabled } = await loadWithEnv(undefined, undefined);
    expect(isAdminEnabled).toBe(false);

    const request = fakeRequest('127.0.0.1', basicAuthHeader('admin', 'whatever'));
    const reply = fakeReply();
    await requireAdminAuth(request, reply as any);

    expect(reply.calls.code).toBe(404);
  });

  test('rejects a request from an IP outside the allowed set with a 404 (not 403, to avoid revealing the route)', async () => {
    const { requireAdminAuth } = await loadWithEnv('secret123', undefined);

    const request = fakeRequest('203.0.113.5', basicAuthHeader('admin', 'secret123'));
    const reply = fakeReply();
    await requireAdminAuth(request, reply as any);

    expect(reply.calls.code).toBe(404);
  });

  test('localhost is always allowed even without DOCVIEW_ADMIN_ALLOWED_IPS set', async () => {
    const { requireAdminAuth } = await loadWithEnv('secret123', undefined);

    const request = fakeRequest('127.0.0.1', undefined);
    const reply = fakeReply();
    await requireAdminAuth(request, reply as any);

    // IP는 통과했으니 인증 부재로 401 — 404(IP 차단)이 아니어야 한다.
    expect(reply.calls.code).toBe(401);
  });

  test('an IP added via DOCVIEW_ADMIN_ALLOWED_IPS is allowed through to the auth check', async () => {
    const { requireAdminAuth } = await loadWithEnv('secret123', '203.0.113.5');

    const request = fakeRequest('203.0.113.5', undefined);
    const reply = fakeReply();
    await requireAdminAuth(request, reply as any);

    expect(reply.calls.code).toBe(401);
  });

  test('rejects a request with no Authorization header', async () => {
    const { requireAdminAuth } = await loadWithEnv('secret123', undefined);

    const request = fakeRequest('127.0.0.1', undefined);
    const reply = fakeReply();
    await requireAdminAuth(request, reply as any);

    expect(reply.calls.code).toBe(401);
    expect(reply.calls.header?.['WWW-Authenticate']).toMatch(/^Basic /);
  });

  test('rejects a wrong password', async () => {
    const { requireAdminAuth } = await loadWithEnv('secret123', undefined);

    const request = fakeRequest('127.0.0.1', basicAuthHeader('admin', 'wrong-password'));
    const reply = fakeReply();
    await requireAdminAuth(request, reply as any);

    expect(reply.calls.code).toBe(401);
  });

  test('accepts the correct password regardless of username (username is not checked)', async () => {
    const { requireAdminAuth } = await loadWithEnv('secret123', undefined);

    const request = fakeRequest('127.0.0.1', basicAuthHeader('anyone-at-all', 'secret123'));
    const reply = fakeReply();
    await requireAdminAuth(request, reply as any);

    expect(reply.calls.code).toBeUndefined();
  });
});
