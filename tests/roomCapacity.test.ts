import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

/**
 * 동시에 열 수 있는 "대용량 xlsx 룸" 개수 상한(room.ts의 checkRoomCapacity) 검증.
 *
 * 진단된 문제: 대용량 xlsx 룸 하나가 canonicalGrid 전체를 메모리에 들고 있어(29만 행급 파일
 * 기준 수백MB~1GB대), 동시 접속자가 늘면 그만큼 서버 메모리가 곱해져 OOM 위험이 커진다. 이 상한은
 * "새로 열어야 하는" 대용량 문서만 막는다 — 이미 열려 있는 문서에 대한 추가 접속자는 항상 받는다
 * (그 룸은 이미 한 번만 세어져 있으므로 추가 메모리 비용이 없다).
 *
 * checkRoomCapacity는 파일 바이트 크기만 보고(파싱 없이) 판단하므로, 대부분의 테스트는 실제
 * xlsx 내용 없이 크기만 맞춘 더미 버퍼로 충분하다 — 실제 룸을 열고 닫아 슬롯이 반환되는지 보는
 * 마지막 테스트만 joinEditSession을 통해 실제 연결 수명주기를 탄다.
 */

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  send(_data: Uint8Array, cb?: (err?: Error) => void): void {
    cb?.();
  }
  close(): void {}
  ping(): void {}
}

const LARGE = Buffer.alloc(6 * 1024 * 1024); // 5MB 기준을 넘는 크기 — 내용은 무관(바이트 크기만 봄)
const SMALL = Buffer.alloc(1024);

async function seedLargeDoc(docId: string, bytes: Buffer): Promise<void> {
  const { store } = await import('../src/storage.js');
  await store.saveOriginal(docId, 'xlsx', bytes);
  await store.writeMeta({
    id: docId,
    name: 'big.xlsx',
    ext: 'xlsx',
    status: 'ready',
    pageCount: 1,
    createdAt: new Date().toISOString(),
    progress: 100,
    viewKind: 'grid',
  });
}

describe('room.ts: concurrent large-xlsx room capacity cap', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-capacity-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  test('small xlsx files are never capped, no matter how many are checked', async () => {
    const { checkRoomCapacity } = await import('../src/collab/room.js');
    for (let i = 0; i < 10; i += 1) {
      const docId = randomUUID();
      await seedLargeDoc(docId, SMALL);
      await expect(checkRoomCapacity(docId, 'xlsx')).resolves.toEqual({ ok: true });
    }
  });

  test('non-xlsx documents are never capped even if the file is large', async () => {
    const { checkRoomCapacity } = await import('../src/collab/room.js');
    for (let i = 0; i < 10; i += 1) {
      const docId = randomUUID();
      await seedLargeDoc(docId, LARGE); // ext는 아래서 'docx'로 넘겨 확장자만 다르게 취급
      await expect(checkRoomCapacity(docId, 'docx')).resolves.toEqual({ ok: true });
    }
  });

  test('the 4th concurrent large xlsx room is rejected with the approved message', async () => {
    const { checkRoomCapacity } = await import('../src/collab/room.js');
    const docIds = await Promise.all(
      Array.from({ length: 4 }, async () => {
        const docId = randomUUID();
        await seedLargeDoc(docId, LARGE);
        return docId;
      }),
    );

    for (let i = 0; i < 3; i += 1) {
      await expect(checkRoomCapacity(docIds[i], 'xlsx')).resolves.toEqual({ ok: true });
    }
    await expect(checkRoomCapacity(docIds[3], 'xlsx')).resolves.toEqual({
      ok: false,
      reason: '사용자가 많습니다. 잠시 후 다시 시도해주세요.',
    });
  });

  test('re-checking an already-reserved docId does not consume an extra slot', async () => {
    const { checkRoomCapacity } = await import('../src/collab/room.js');
    const docIds = await Promise.all(
      Array.from({ length: 3 }, async () => {
        const docId = randomUUID();
        await seedLargeDoc(docId, LARGE);
        return docId;
      }),
    );

    for (let i = 0; i < 3; i += 1) {
      await checkRoomCapacity(docIds[i], 'xlsx');
    }
    // 같은 docId를 여러 번 다시 체크해도(재접속 등) 슬롯을 새로 소비하지 않는다.
    await expect(checkRoomCapacity(docIds[0], 'xlsx')).resolves.toEqual({ ok: true });
    await expect(checkRoomCapacity(docIds[1], 'xlsx')).resolves.toEqual({ ok: true });

    const extraDocId = randomUUID();
    await seedLargeDoc(extraDocId, LARGE);
    await expect(checkRoomCapacity(extraDocId, 'xlsx')).resolves.toEqual({
      ok: false,
      reason: '사용자가 많습니다. 잠시 후 다시 시도해주세요.',
    });
  });

  test('closing the last connection of a large room frees its slot for a new large document', async () => {
    const { checkRoomCapacity, joinEditSession } = await import('../src/collab/room.js');

    const docIds = await Promise.all(
      Array.from({ length: 3 }, async () => {
        const docId = randomUUID();
        await seedLargeDoc(docId, LARGE);
        return docId;
      }),
    );

    const sockets: FakeSocket[] = [];
    for (const docId of docIds) {
      const socket = new FakeSocket();
      // checkRoomCapacity가 아니라 joinEditSession(실제 룸 생성)까지 거쳐야 closeConn이 슬롯을
      // 관리하는 실제 흐름을 탄다 — edit.ts가 실제로 하는 순서 그대로.
      const capacity = await checkRoomCapacity(docId, 'xlsx');
      expect(capacity.ok).toBe(true);
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
      sockets.push(socket);
    }

    const fourthDocId = randomUUID();
    await seedLargeDoc(fourthDocId, LARGE);
    await expect(checkRoomCapacity(fourthDocId, 'xlsx')).resolves.toEqual({
      ok: false,
      reason: '사용자가 많습니다. 잠시 후 다시 시도해주세요.',
    });

    // 첫 번째 문서의 마지막 연결이 끊기면(closeConn) 그 룸이 완전히 닫히며 슬롯도 반환된다.
    sockets[0].emit('close');
    await expect(checkRoomCapacity(fourthDocId, 'xlsx')).resolves.toEqual({ ok: true });
  });
});
