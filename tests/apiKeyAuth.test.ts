import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

/** apiKeyAuth.ts는 config.ts를 import 시점에 한 번만 읽으므로, 서로 다른 DOCVIEW_API_KEYS
 * 값으로 테스트하려면 매번 모듈 캐시를 초기화하고 env를 바꾼 뒤 동적으로 다시 import해야 한다. */
async function loadWithEnv(apiKeys: string | undefined) {
  vi.resetModules();
  if (apiKeys === undefined) delete process.env.DOCVIEW_API_KEYS;
  else process.env.DOCVIEW_API_KEYS = apiKeys;
  return import('../src/apiKeyAuth.js');
}

function fakeReply() {
  const calls: { code?: number; body?: unknown } = {};
  return {
    code(status: number) {
      calls.code = status;
      return this;
    },
    async send(body: unknown) {
      calls.body = body;
    },
    calls,
  };
}

function fakeRequest(authHeader?: string) {
  return { headers: { authorization: authHeader } } as unknown as Parameters<
    typeof import('../src/apiKeyAuth.js').requireApiKey
  >[0];
}

describe('apiKeyAuth', () => {
  const originalEnv = process.env.DOCVIEW_API_KEYS;

  afterEach(() => {
    if (originalEnv === undefined) delete process.env.DOCVIEW_API_KEYS;
    else process.env.DOCVIEW_API_KEYS = originalEnv;
    vi.resetModules();
  });

  test('when DOCVIEW_API_KEYS is unset, auth is disabled and every request passes through', async () => {
    const { requireApiKey, isApiKeyAuthEnabled } = await loadWithEnv(undefined);
    expect(isApiKeyAuthEnabled).toBe(false);

    const request = fakeRequest(undefined);
    const reply = fakeReply();
    await requireApiKey(request, reply as any);

    expect(reply.calls.code).toBeUndefined();
  });

  test('rejects a request with no Authorization header when keys are configured', async () => {
    const { requireApiKey } = await loadWithEnv('maintenance-management:secret-1');

    const request = fakeRequest(undefined);
    const reply = fakeReply();
    await requireApiKey(request, reply as any);

    expect(reply.calls.code).toBe(401);
  });

  test('rejects a request with an unknown key', async () => {
    const { requireApiKey } = await loadWithEnv('maintenance-management:secret-1');

    const request = fakeRequest('Bearer wrong-key');
    const reply = fakeReply();
    await requireApiKey(request, reply as any);

    expect(reply.calls.code).toBe(401);
  });

  test('accepts a valid key and tags the request with the caller project name', async () => {
    const { requireApiKey } = await loadWithEnv('maintenance-management:secret-1,other-project:secret-2');

    const request = fakeRequest('Bearer secret-2');
    const reply = fakeReply();
    await requireApiKey(request, reply as any);

    expect(reply.calls.code).toBeUndefined();
    expect((request as any).callerProject).toBe('other-project');
  });

  test('a leaked key from one project does not grant a different project name (still valid though)', async () => {
    // 키 자체가 다르면 당연히 거부되지만, 여러 프로젝트 키가 동시에 설정되어 있어도
    // 서로 다른 키는 서로 다른 caller로 정확히 구분되어야 한다 (책임 추적/개별 회수 목적).
    const { requireApiKey } = await loadWithEnv('project-a:key-a,project-b:key-b');

    const requestA = fakeRequest('Bearer key-a');
    const replyA = fakeReply();
    await requireApiKey(requestA, replyA as any);
    expect((requestA as any).callerProject).toBe('project-a');

    const requestB = fakeRequest('Bearer key-b');
    const replyB = fakeReply();
    await requireApiKey(requestB, replyB as any);
    expect((requestB as any).callerProject).toBe('project-b');
  });
});
