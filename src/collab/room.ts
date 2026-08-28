import type { IncomingMessage } from 'node:http';
import type { WebSocket } from 'ws';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import {
  Awareness,
  applyAwarenessUpdate,
  encodeAwarenessUpdate,
  removeAwarenessStates,
} from 'y-protocols/awareness';
import {
  messageYjsSyncStep1,
  messageYjsSyncStep2,
  messageYjsUpdate,
  readSyncStep1,
  readSyncStep2,
  readUpdate,
  writeSyncStep1,
  writeUpdate,
} from 'y-protocols/sync';
import { prosemirrorToYXmlFragment, yXmlFragmentToProseMirrorRootNode } from 'y-prosemirror';
import { recordAudit } from '../audit.js';
import { findConverter } from '../converters/index.js';
import { docxSchema } from '../docx/schema.js';
import { exportDocx } from '../docx/export.js';
import { importDocx } from '../docx/import.js';
import { invalidateRenderCache } from '../rendering.js';
import { convertToPdf } from '../routes/documents.js';
import { invalidateViewCaches } from '../routes/viewer.js';
import { store } from '../storage.js';
import { parseXlsxToGrid, serializeGrid } from '../spreadsheet.js';
import { exportXlsxFromGrid, reparseGridForViewer, seedGridFromXlsx } from '../xlsx/collab.js';

/**
 * y-websocket이 제공하는 참조 서버 유틸(bin/utils.cjs)을 그대로 쓰지 않고 같은 프로토콜을
 * 직접 구현한다. 이유: 그 유틸은 CommonJS라 내부에서 `require('yjs')`로 yjs를 불러오는데,
 * 이 프로젝트(y-prosemirror 포함)는 ESM으로 `import`해 yjs를 불러온다 — package.json exports의
 * import/require 조건이 서로 다른 번들(dist/yjs.mjs vs dist/yjs.cjs)을 가리켜 같은 프로세스 안에
 * yjs 모듈이 두 벌 로드되고, Y.Doc/Y.XmlElement의 instanceof 검사가 깨져 동기화가 조용히
 * 실패할 수 있다(실제로 개발 중 "Yjs was already imported" 경고로 확인됨, yjs#438).
 * 아래는 순수 ESM 경로(yjs, y-protocols/sync, y-protocols/awareness, lib0/*)만 사용해
 * 이 문제를 원천 차단한다. 메시지 타입 값과 처리 흐름은 y-websocket의 참조 구현과 동일하게
 * 맞춰 두었다(클라이언트는 표준 `y-websocket`의 WebsocketProvider를 그대로 쓴다).
 */

/** ProseMirror 문서를 담는 Y.XmlFragment의 고정 키 이름 (클라이언트의 ySyncPlugin과 일치해야 한다) */
export const XML_FRAGMENT_NAME = 'prosemirror';

const MESSAGE_SYNC = 0;
const MESSAGE_AWARENESS = 1;

const AUTOSAVE_IDLE_MS = 8_000;
const PING_INTERVAL_MS = 30_000;

/** 룸이 다루는 문서 종류 — 시딩·저장 로직만 갈라지고 Yjs 동기화 프로토콜 자체는 공통이다. */
type DocKind = 'docx' | 'xlsx';

function kindOfExt(ext: string): DocKind {
  return ext === 'xlsx' ? 'xlsx' : 'docx';
}

interface Room {
  readonly docId: string;
  readonly kind: DocKind;
  readonly ydoc: Y.Doc;
  readonly awareness: Awareness;
  /** 연결 → 그 연결이 소유한(생성한) awareness clientId 집합 */
  readonly conns: Map<WebSocket, Set<number>>;
  /** 읽기 전용(canEdit:false)으로 참여한 연결 — 이 연결이 보낸 syncStep2/update는 무시한다.
   *  syncStep1(자기 상태 벡터를 보내고 서버 diff를 요청하는 것뿐)은 그대로 응답해야 문서 내용을
   *  받을 수 있으므로 막지 않는다 — 실제 "쓰기"에 해당하는 메시지만 서버에서 직접 걸러낸다
   *  (클라이언트 쪽에서 편집 UI를 숨기는 것과 별개로, 조작된 소켓 메시지로 과거 연도 데이터가
   *  바뀌는 일을 서버에서 원천 차단한다). */
  readonly readOnlyConns: Set<WebSocket>;
  autosaveTimer: NodeJS.Timeout | null;
  /** 문서 시딩(원본 → Y.Doc)이 끝났는지 — 첫 연결에 syncStep1을 보내기 전에 기다린다 */
  readonly ready: Promise<void>;
}

const rooms = new Map<string, Room>();

function send(room: Room, conn: WebSocket, message: Uint8Array): void {
  if (conn.readyState !== conn.CONNECTING && conn.readyState !== conn.OPEN) {
    closeConn(room, conn);
    return;
  }
  try {
    conn.send(message, (err) => {
      if (err) closeConn(room, conn);
    });
  } catch {
    closeConn(room, conn);
  }
}

function broadcast(room: Room, message: Uint8Array): void {
  for (const conn of room.conns.keys()) send(room, conn, message);
}

async function seedRoom(docId: string, kind: DocKind, ydoc: Y.Doc): Promise<void> {
  try {
    const meta = await store.readMeta(docId);
    if (!meta) return;
    const original = await store.readOriginal(docId, meta.ext);
    if (kind === 'xlsx') {
      const grid = await parseXlsxToGrid(original);
      seedGridFromXlsx(ydoc, grid);
    } else {
      const { doc } = importDocx(original);
      prosemirrorToYXmlFragment(doc, ydoc.getXmlFragment(XML_FRAGMENT_NAME));
    }
  } catch (error) {
    console.error(`동시편집 세션 초기화 실패 (문서 ${docId}):`, error);
  }
}

/**
 * 현재 협업 상태(Y.Doc)를 원본 포맷으로 직렬화해 저장하고, 뷰어용 캐시(PDF 또는 그리드)를
 * 재생성하며, 렌더링·검색 캐시를 무효화한다. 유휴 자동 저장과 "마지막 참여자 퇴장" 양쪽에서 호출된다.
 */
async function flushToStore(docId: string, kind: DocKind, ydoc: Y.Doc): Promise<void> {
  const meta = await store.readMeta(docId);
  if (!meta) return;

  if (kind === 'xlsx') {
    const original = await store.readOriginal(docId, meta.ext);
    const newXlsxBytes = await exportXlsxFromGrid(original, ydoc);
    await store.saveOriginal(docId, meta.ext, Buffer.from(newXlsxBytes));
    try {
      const grid = await reparseGridForViewer(newXlsxBytes);
      await store.saveDerived(docId, 'grid', serializeGrid(grid));
      await store.writeMeta({
        ...meta,
        pageCount: grid.sheets.length,
        status: 'ready',
        viewKind: 'grid',
        progress: 100,
      });
    } catch (error) {
      console.error(`동시편집 저장 후 그리드 재생성 실패 (문서 ${docId}):`, error);
    }
  } else {
    const fragment = ydoc.getXmlFragment(XML_FRAGMENT_NAME);
    const pmDoc = yXmlFragmentToProseMirrorRootNode(fragment, docxSchema);
    const original = await store.readOriginal(docId, meta.ext);
    const newDocxBytes = exportDocx(original, pmDoc);
    await store.saveOriginal(docId, meta.ext, Buffer.from(newDocxBytes));

    const converter = findConverter(meta.ext);
    if (converter) {
      try {
        const { pageCount } = await convertToPdf(meta, converter);
        await store.writeMeta({ ...meta, pageCount, status: 'ready', viewKind: 'pages', progress: 100 });
      } catch (error) {
        console.error(`동시편집 저장 후 PDF 재변환 실패 (문서 ${docId}):`, error);
      }
    }
  }
  invalidateRenderCache(docId);
  invalidateViewCaches(docId);
  await recordAudit({ type: 'edit_saved', docId });
}

function scheduleAutosave(room: Room): void {
  if (room.autosaveTimer) clearTimeout(room.autosaveTimer);
  room.autosaveTimer = setTimeout(() => {
    room.autosaveTimer = null;
    void flushToStore(room.docId, room.kind, room.ydoc);
  }, AUTOSAVE_IDLE_MS);
}

function getOrCreateRoom(docId: string, ext: string): Room {
  const existing = rooms.get(docId);
  if (existing) return existing;

  const kind = kindOfExt(ext);
  const ydoc = new Y.Doc({ gc: true });
  const awareness = new Awareness(ydoc);
  awareness.setLocalState(null);
  const conns = new Map<WebSocket, Set<number>>();

  const room: Room = {
    docId,
    kind,
    ydoc,
    awareness,
    conns,
    readOnlyConns: new Set<WebSocket>(),
    autosaveTimer: null,
    ready: seedRoom(docId, kind, ydoc),
  };
  rooms.set(docId, room);

  ydoc.on('update', (update: Uint8Array) => {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeUpdate(encoder, update);
    broadcast(room, encoding.toUint8Array(encoder));
    scheduleAutosave(room);
  });

  awareness.on(
    'update',
    ({ added, updated, removed }: { added: number[]; updated: number[]; removed: number[] }, conn: unknown) => {
      const changed = added.concat(updated, removed);
      if (conn instanceof Object && room.conns.has(conn as WebSocket)) {
        const controlled = room.conns.get(conn as WebSocket)!;
        added.forEach((id) => controlled.add(id));
        removed.forEach((id) => controlled.delete(id));
      }
      const encoder = encoding.createEncoder();
      encoding.writeVarUint(encoder, MESSAGE_AWARENESS);
      encoding.writeVarUint8Array(encoder, encodeAwarenessUpdate(awareness, changed));
      broadcast(room, encoding.toUint8Array(encoder));
    },
  );

  return room;
}

function closeConn(room: Room, conn: WebSocket): void {
  const controlled = room.conns.get(conn);
  if (controlled) {
    room.conns.delete(conn);
    room.readOnlyConns.delete(conn);
    removeAwarenessStates(room.awareness, [...controlled], null);
    if (room.conns.size === 0) {
      if (room.autosaveTimer) {
        clearTimeout(room.autosaveTimer);
        room.autosaveTimer = null;
      }
      rooms.delete(room.docId);
      void room.ready
        .then(() => flushToStore(room.docId, room.kind, room.ydoc))
        .catch((error) => console.error(`동시편집 종료 저장 실패 (문서 ${room.docId}):`, error))
        .finally(() => room.ydoc.destroy());
    }
  }
  try {
    conn.close();
  } catch {
    /* 이미 닫힌 연결 */
  }
}

function handleMessage(room: Room, conn: WebSocket, message: Uint8Array): void {
  try {
    const encoder = encoding.createEncoder();
    const decoder = decoding.createDecoder(message);
    const messageType = decoding.readVarUint(decoder);
    switch (messageType) {
      case MESSAGE_SYNC: {
        encoding.writeVarUint(encoder, MESSAGE_SYNC);
        // syncStep1은 "내 상태 벡터를 보낼 테니 diff를 달라"는 읽기 요청이라 읽기 전용 연결도
        // 반드시 처리해야 문서 내용을 받는다 — syncStep2/update만 실제 쓰기이므로 읽기 전용
        // 연결에서 온 것이면 조용히 버린다(클라이언트를 신뢰하지 않고 서버가 직접 막는다).
        const syncMessageType = decoding.readVarUint(decoder);
        const readOnly = room.readOnlyConns.has(conn);
        if (syncMessageType === messageYjsSyncStep1) {
          readSyncStep1(decoder, encoder, room.ydoc);
        } else if (!readOnly) {
          if (syncMessageType === messageYjsSyncStep2) {
            readSyncStep2(decoder, room.ydoc, conn);
          } else if (syncMessageType === messageYjsUpdate) {
            readUpdate(decoder, room.ydoc, conn);
          }
        }
        if (encoding.length(encoder) > 1) send(room, conn, encoding.toUint8Array(encoder));
        break;
      }
      case MESSAGE_AWARENESS: {
        applyAwarenessUpdate(room.awareness, decoding.readVarUint8Array(decoder), conn);
        break;
      }
      default:
        break;
    }
  } catch (error) {
    console.error('동시편집 메시지 처리 오류:', error);
  }
}

function toUint8Array(data: ArrayBuffer | Buffer | ArrayBufferView): Uint8Array {
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  return new Uint8Array(data);
}

/**
 * 인증·권한 검증(src/routes/edit.ts)을 마친 WebSocket 연결을 문서별 협업 룸에 연결한다.
 *
 * 리스너 등록 순서가 중요하다: room.ready(docx → Y.Doc 시딩)를 기다리는 동안에도 클라이언트는
 * 연결이 열리자마자 자신의 syncStep1을 보낸다. `socket.on('message', ...)`를 await 뒤로 미루면
 * 그 사이 도착한 메시지가 유실되어(리스너가 아직 없으므로 EventEmitter가 그냥 버림) 동기화가
 * 영원히 완료되지 않는 문제가 실제로 있었다 — 그래서 리스너는 항상 먼저 붙이고, room.ready가
 * 끝나기 전에 도착한 메시지는 큐에 쌓아뒀다가 준비되는 즉시 순서대로 처리한다.
 *
 * earlyMessages: 호출자(edit.ts)가 인증 확인(그 자체도 비동기)을 하는 동안 이미 도착해 버린
 * 메시지들 — 같은 이유로 여기서 만드는 큐에 먼저 합쳐 순서대로 처리한다.
 */
export async function joinEditSession(
  socket: WebSocket,
  _rawRequest: IncomingMessage,
  docId: string,
  docExt: string,
  earlyMessages: ReadonlyArray<ArrayBuffer | Buffer> = [],
  readOnly = false,
): Promise<void> {
  const room = getOrCreateRoom(docId, docExt);

  socket.binaryType = 'arraybuffer';
  room.conns.set(socket, new Set());
  if (readOnly) room.readOnlyConns.add(socket);

  let seeded = false;
  const pendingMessages: Uint8Array[] = earlyMessages.map(toUint8Array);

  socket.on('message', (data: ArrayBuffer | Buffer) => {
    const bytes = toUint8Array(data);
    if (!seeded) {
      pendingMessages.push(bytes);
      return;
    }
    handleMessage(room, socket, bytes);
  });

  let pongReceived = true;
  const pingInterval = setInterval(() => {
    if (!pongReceived) {
      if (room.conns.has(socket)) closeConn(room, socket);
      clearInterval(pingInterval);
      return;
    }
    if (room.conns.has(socket)) {
      pongReceived = false;
      try {
        socket.ping();
      } catch {
        closeConn(room, socket);
        clearInterval(pingInterval);
      }
    }
  }, PING_INTERVAL_MS);

  socket.on('close', () => {
    closeConn(room, socket);
    clearInterval(pingInterval);
  });
  socket.on('pong', () => {
    pongReceived = true;
  });

  await room.ready;
  if (!room.conns.has(socket)) return; // 대기 중 연결이 이미 끊겼으면 조용히 종료
  seeded = true;
  for (const bytes of pendingMessages) handleMessage(room, socket, bytes);

  const syncEncoder = encoding.createEncoder();
  encoding.writeVarUint(syncEncoder, MESSAGE_SYNC);
  writeSyncStep1(syncEncoder, room.ydoc);
  send(room, socket, encoding.toUint8Array(syncEncoder));

  const states = room.awareness.getStates();
  if (states.size > 0) {
    const awarenessEncoder = encoding.createEncoder();
    encoding.writeVarUint(awarenessEncoder, MESSAGE_AWARENESS);
    encoding.writeVarUint8Array(awarenessEncoder, encodeAwarenessUpdate(room.awareness, [...states.keys()]));
    send(room, socket, encoding.toUint8Array(awarenessEncoder));
  }
}

/** 현재 메모리에 열려 있는 협업 룸 수 (진단·테스트용) */
export function activeRoomCount(): number {
  return rooms.size;
}
