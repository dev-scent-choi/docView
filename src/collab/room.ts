import { randomUUID } from 'node:crypto';
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
import { exportPptxFromSlides, seedSlidesFromPptx } from '../pptx/collab.js';
import { PPTX_SLIDES_ARRAY, type SlideEntry } from './pptxShared.js';
import { DOC_META_MAP, DOC_META_SAVED_AT_KEY } from './gridShared.js';

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
type DocKind = 'docx' | 'xlsx' | 'pptx';

function kindOfExt(ext: string): DocKind {
  if (ext === 'xlsx') return 'xlsx';
  if (ext === 'pptx') return 'pptx';
  return 'docx';
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
    } else if (kind === 'pptx') {
      seedSlidesFromPptx(original, ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY));
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
  } else if (kind === 'pptx') {
    const original = await store.readOriginal(docId, meta.ext);
    const slidesArray = ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);
    // store.saveOriginal(디스크 쓰기, await 동안 이벤트 루프가 다른 클라이언트의 편집을 그대로
    // 받아들인다)이 끝난 뒤 배열이 이 시점과 달라졌는지 비교하기 위한 스냅샷 — 진짜 위험 구간은
    // exportPptxFromSlides의 동기적 zip 조작(끼어들 틈 없음)이 아니라 바로 이 await다.
    const snapshotBeforeSave = JSON.stringify(slidesArray.toArray());
    const { bytes: newPptxBytes, finalParts } = exportPptxFromSlides(original, slidesArray);
    await store.saveOriginal(docId, meta.ext, Buffer.from(newPptxBytes));

    // 저장이 끝난 직후, Yjs 배열을 "방금 디스크에 쓴 내용"과 다시 일치시킨다: 모든 항목을
    // finalParts 순서의 'original' 항목으로 재구성한다. 이렇게 하지 않으면 (1) duplicate 항목이
    // sourcePageNumber: null인 채로 남아 썸네일이 계속 깨진 채고, (2) 그 duplicate가 가리키던
    // sourceSlidePart를 이후에 원본 쪽에서 삭제하면 다음 저장이 그 항목을 조용히 떨어뜨리거나
    // (필터로 막긴 했지만) 의도와 다른 결과를 낳으며, (3) 매 저장마다 새 slideN.xml 번호가 계속
    // 새로 발급되어 파트 번호가 무한정 커진다. 저장 직후 배열을 "디스크 = 배열"로 리셋해 두면
    // 두 번째 저장은 완전히 멱등이고, 복제된 슬라이드도 다음 렌더부터 진짜 썸네일을 가진다.
    //
    // 단, store.saveOriginal이 진행되는 동안(디스크 쓰기라 수십~수백ms 걸릴 수 있다) 다른
    // 클라이언트가 새 편집을 넣었다면, 그 시점의 배열은 이미 export가 반영한 상태와 다르다 —
    // 그런데도 그대로 리셋하면 방금 들어온 새 편집이 통째로 덮어써져 사라진다. 그래서 저장 시작
    // 전 스냅샷과 지금(저장 끝난 시점)을 비교해, 그 사이에 아무도 배열을 건드리지 않았을 때만
    // 재작성한다 — 누군가 건드렸다면 이번엔 리셋을 건너뛰고, 그 새 편집은 다음 자동저장
    // 주기(8초 후)에서 자연스럽게 반영된다(그동안 sourcePageNumber가 한 번 더 낡은 채로 남는
    // 정도의 대가일 뿐, 편집 유실은 없다).
    if (JSON.stringify(slidesArray.toArray()) === snapshotBeforeSave) {
      ydoc.transact(() => {
        slidesArray.delete(0, slidesArray.length);
        slidesArray.insert(
          0,
          finalParts.map((part, index) => ({
            id: randomUUID(),
            kind: 'original' as const,
            sourceSlidePart: part,
            sourcePageNumber: index + 1,
          })),
        );
      });
    }

    const converter = findConverter(meta.ext);
    if (converter) {
      try {
        const { pageCount } = await convertToPdf(meta, converter);
        await store.writeMeta({ ...meta, pageCount, status: 'ready', viewKind: 'pages', progress: 100 });
      } catch (error) {
        console.error(`동시편집 저장 후 PDF 재변환 실패 (문서 ${docId}):`, error);
      }
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
  // 일반 Y.Map 쓰기라 기존 sync 프로토콜을 그대로 타고 모든 클라이언트에 전파된다 — 커스텀
  // 메시지 없이도 클라이언트가 "저장됨" 표시를 실시간으로 띄울 수 있는 이유.
  ydoc.getMap<number>(DOC_META_MAP).set(DOC_META_SAVED_AT_KEY, Date.now());
}

/**
 * Ctrl+S 등 클라이언트가 명시적으로 "지금 저장" 요청을 보냈을 때 쓴다 — 8초 유휴 자동저장을
 * 기다리지 않고 즉시 flushToStore를 돈다. 그 문서를 편집 중인 룸이 메모리에 없으면(아무도 안
 * 열어봤거나 이미 유휴 저장까지 끝난 상태) 저장할 미변경 사항이 없다는 뜻이라 false를 돌려준다.
 */
export async function flushNow(docId: string): Promise<boolean> {
  const room = rooms.get(docId);
  if (!room) return false;
  await room.ready;
  if (room.autosaveTimer) {
    clearTimeout(room.autosaveTimer);
    room.autosaveTimer = null;
  }
  await flushToStore(room.docId, room.kind, room.ydoc);
  return true;
}

/**
 * 지금 그 문서를 편집 중인 세션(룸)이 메모리에 떠 있으면, 아직 유휴 자동저장(8초 debounce) 전이라
 * 디스크의 원본은 최신이 아닐 수 있다 — 그런 경우 디스크에 쓰지는 않고 그 자리에서만 최신 Y.Doc
 * 상태로 바이트를 만들어 돌려준다("방금 입력한 값을 바로 다운로드"가 안 되던 문제의 원인). 룸이
 * 없으면(=마지막 참여자가 나가며 이미 flushToStore로 디스크에 반영됨) null을 돌려주고, 호출측은
 * store.readOriginal을 그대로 쓰면 된다.
 */
export async function getCurrentBytes(docId: string): Promise<Uint8Array | null> {
  const room = rooms.get(docId);
  if (!room) return null;
  await room.ready;
  const meta = await store.readMeta(docId);
  if (!meta) return null;
  const original = await store.readOriginal(docId, meta.ext);
  if (room.kind === 'xlsx') {
    return exportXlsxFromGrid(original, room.ydoc);
  }
  if (room.kind === 'pptx') {
    // 읽기 전용 헬퍼라 Yjs 배열은 절대 건드리지 않는다 — flushToStore의 저장-후 배열 재작성과
    // 달리, 여기서는 바이트만 필요하다.
    return exportPptxFromSlides(original, room.ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY)).bytes;
  }
  const fragment = room.ydoc.getXmlFragment(XML_FRAGMENT_NAME);
  const pmDoc = yXmlFragmentToProseMirrorRootNode(fragment, docxSchema);
  return exportDocx(original, pmDoc);
}

function scheduleAutosave(room: Room): void {
  if (room.autosaveTimer) clearTimeout(room.autosaveTimer);
  room.autosaveTimer = setTimeout(() => {
    room.autosaveTimer = null;
    // flushToStore가 던지면(예: 저장 도중 발견한 데이터 이상) .catch() 없이 그냥 버려진 Promise의
    // reject는 Node에서 unhandledRejection이 되어 프로세스 전체를 죽인다 — 그러면 이 문서만이
    // 아니라 그 순간 다른 모든 문서(docx/xlsx 포함)를 편집 중이던 세션까지 함께 끊긴다. 반드시
    // 여기서 잡아 로그만 남긴다.
    flushToStore(room.docId, room.kind, room.ydoc).catch((error) => {
      console.error(`유휴 자동저장 실패 (문서 ${room.docId}):`, error);
    });
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
