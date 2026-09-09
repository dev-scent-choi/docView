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
import {
  saveGridDerived,
  yieldToEventLoop,
  MAX_COLS,
  MAX_ROWS,
  YIELD_EVERY_N_ROWS,
  type GridCell,
  type GridDocument,
  type GridSheet,
} from '../spreadsheet.js';
import { exportXlsxFromGridStreaming } from '../xlsx/exportStreaming.js';
import {
  exportXlsxFromGrid,
  findOrAddGridStyleIndex,
  fromCellStyleAttrs,
  isAutosaveExemptOrigin,
  reparseGridForViewer,
  seedGridPageFromXlsx,
  seedGridStructureOnly,
  seedGridStylesFromXlsx,
  shiftGridStyleKeys,
  unseedGridPageFromXlsx,
  GRID_COMPUTED_WRITE_ORIGIN,
  GRID_STRUCTURE_SHIFT_ORIGIN,
  GRID_WINDOW_SEED_ORIGIN,
} from '../xlsx/collab.js';
import { shiftMerges } from '../xlsx/gridShift.js';
import {
  applyColShift,
  applyRowShift,
  closeLazyGridStore,
  getCell,
  getRow,
  growSheet,
  ingestXlsxLazily,
  setDirtyCell,
  type LazyGridSheetMeta,
  type LazyGridStore,
} from '../xlsx/lazyGridStore.js';
import { exportPptxFromSlides, seedSlidesFromPptx } from '../pptx/collab.js';
import { computeGridCellDisplayAsync, formatComputedValueForStorage } from '../xlsx/formulas.js';
import { PPTX_SLIDES_ARRAY, type SlideEntry } from './pptxShared.js';
import {
  DOC_META_MAP,
  DOC_META_SAVED_AT_KEY,
  DOC_META_SAVE_ERROR_AT_KEY,
  GRID_CELLS_MAP,
  GRID_COMPUTED_MAP,
  GRID_MERGES_MAP,
  GRID_STRUCTURE_MAP,
  GRID_STYLES_MAP,
  gridCellKey,
  parseGridCellKey,
  PAGE_SIZE_ROWS,
  type CellStyleAttrs,
  type GridMergeShape,
  type GridStructureOpKind,
  type GridStructureOpRecord,
  type SheetStructure,
} from './gridShared.js';
import { runSerialized } from './saveSerializer.js';

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
const MESSAGE_PAGE_SUBSCRIBE = 2;
/**
 * 행/열 삽입·삭제를 **명시적인 연산으로** 알리는 메시지 — 클라이언트가 Yjs 셀 키를 직접 옮기는
 * 대신 "이 시트의 이 자리에 행 N개를 넣어 달라"만 보내고, 실제 이동은 서버가 시트 전체를 담고
 * 있는 canonical 그리드에서 수행한다(gridShift.ts / applyGridStructureOp).
 *
 * **왜 3이 아니라 4인가**: 0~3은 y-websocket의 표준 프로토콜이 이미 쓰고 있다
 * (messageSync=0, messageAwareness=1, messageAuth=2, messageQueryAwareness=3 —
 * node_modules/y-websocket/src/y-websocket.js). 특히 3(messageQueryAwareness)은 클라이언트가
 * 접속 직후 **실제로 서버에 보내는** 메시지라, 3을 쓰면 그 메시지가 이 분기로 흘러들어와 JSON
 * 파싱 오류 로그만 쌓인다. 기존 MESSAGE_PAGE_SUBSCRIBE=2가 messageAuth와 값이 겹치는 것은
 * messageAuth가 서버→클라이언트 전용이라 실무상 무해해서 그대로 뒀지만(계획에 기록된 알려진
 * 항목), 새로 추가하는 값은 그 범위를 아예 벗어나게 둔다.
 */
const MESSAGE_GRID_STRUCTURE_OP = 4;

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
  ready: Promise<void>;
  /**
   * true인 동안 발생한 ydoc 'update'는 scheduleAutosave를 다시 예약하지 않는다.
   * flushToStore 자신이 저장 결과를 Y.Doc에 반영(pptx 배열 재작성, 저장 시각 기록)할 때만
   * 짧게 켠다 — 이 자기 자신의 쓰기까지 "새 편집"으로 취급해 또 자동저장을 예약하면, 그
   * 자동저장이 끝나며 다시 같은 쓰기를 하고 또 예약하는 식으로 영원히 반복된다(실사용자
   * 편집이 하나도 없어도 8초마다 계속 저장·PDF 재변환이 도는 문제의 원인이었다). 진짜
   * 사용자/원격 편집이 도착하는 시점(await 중, 이 플래그가 꺼져 있는 동안)에는 그대로
   * scheduleAutosave가 걸려야 하므로, 함수 전체가 아니라 자기 자신의 동기적 쓰기 구간에만
   * 좁게 두른다.
   */
  suppressAutosaveReschedule: boolean;
  /**
   * xlsx 룸에서만 쓰인다 — 룸이 사는 동안 진짜로 상주하는 지연(lazy) 그리드 저장소(Task 3-7).
   * 값·서식 편집은 전부 `dirtyOverlay`에만 반영되고, 편집되지 않은 원본 셀은 필요할 때만
   * 디스크에서 읽는다. 저장·다운로드·수식 계산·구조연산 등 canonical을 읽거나 쓰는 모든 경로가
   * 이제 이 필드 **하나**를 기준으로 한다(Task 11 이전에는 항상 전체가 메모리에 상주하는
   * `canonicalGrid: GridDocument`가 이 역할이었다 — 이 필드가 그걸 완전히 대체한다). docx/pptx
   * 룸에서는 항상 null.
   */
  lazyGrid: LazyGridStore | null;
  /**
   * xlsx 룸에서만 쓰인다 — 룸이 처음 시딩될 때(`ingestXlsxLazily` 직후, 편집이 반영되기 전) 각
   * 시트의 (rowCount, colCount)만 남겨 둔 가벼운 스냅샷("원본 그대로"의 크기만). 협업 중 새로
   * 추가된 시트는 이 배열에 자리가 없다(원래 존재하지 않았으므로 — undefined로 판별).
   *
   * **왜 필요한가 (`lazyGrid.sheetMeta[i].rowCount/colCount`로 대체할 수 없는 이유)**: `sheetMeta`는
   * 지금까지의 모든 편집(구조연산 + 직접 확장/축소)이 반영된 **현재** 값이라 계속 바뀐다. 반면
   * `exportXlsxFromGrid`가 구조연산(`structuralOps`)을 원본 워크북에 재생할 때 그 재생의
   * **출발점**으로 삼아야 하는 건 "룸이 열렸을 때 원래 몇 행/열이었는지"라는 고정된 기준선이다
   * (직접 확장/축소는 `structuralOps`에 기록되지 않으므로 현재 `sheetMeta` 값에서 거꾸로 계산할
   * 수도 없다). ExcelJS의 raw `sheet.rowCount`(서식만 있는 꼬리 행까지 세어 부풀려질 수 있음)도
   * 대신 쓸 수 없다 — 그게 바로 이 필드가 막는 버그다(`tests/xlsx-collab.test.ts`의 "export does
   * not call spliceRows/spliceColumns when nothing was actually edited" 참고).
   */
  baselineSheetExtents: ReadonlyArray<{ readonly rowCount: number; readonly colCount: number }> | null;
  /**
   * xlsx 룸에서만 쓰인다 — 협업 중 적용된 행/열 삽입·삭제 연산(MESSAGE_GRID_STRUCTURE_OP)을
   * 순서대로 쌓아 둔 기록. exportXlsxFromGrid가 원본을 다시 로드한 워크북에 이 기록을 같은
   * 순서로 재생(spliceRows/spliceColumns)해 lazyGrid와 같은 물리적 행/열 위치를 맞춘다.
   * 룸이 여러 번 저장돼도(자동저장마다) 항상 같은 원본(originalXlsxBytes)에서 다시 시작하므로
   * 이 기록은 룸이 사는 동안 계속 누적돼야 멱등이 유지된다(originalXlsxBytes 주석과 같은 이유).
   */
  readonly structuralOps: GridStructureOpRecord[];
  /**
   * xlsx 룸에서만 쓰인다 — 룸을 처음 열 때 읽은 **원본 바이트 그대로**. 이후 모든 내보내기
   * (자동저장·다운로드)는 디스크의 현재 파일이 아니라 항상 이 바이트를 기준으로 한다.
   *
   * **왜 매번 store.readOriginal을 다시 읽으면 안 되는가**: baselineSheetExtents/structuralOps/
   * lazyGrid.dirtyOverlay(= 이 바이트를 기준으로 한 상대적 변경 기록)와 removedSheetIndices/sheetDisplayOrder는
   * 전부 "룸이 열린 시점의 원본"을 기준으로 한
   * 상대적 표현이다. 자동저장이 한 번 돌고 나면 디스크의 파일은 이미 그 변경이 반영된 상태라,
   * 다음 저장에서 같은 removedSheetIndices를 다시 적용하면 **이미 지워진 시트 대신 그 뒤의
   * 멀쩡한 시트를 지운다**(두 번째 저장에서 데이터가 조용히 사라지는 버그). 항상 같은 원본에서
   * 같은 diff를 다시 적용하면 몇 번을 저장해도 결과가 동일하다(멱등).
   */
  originalXlsxBytes: Uint8Array | null;
  /**
   * xlsx 룸에서만 쓰인다 — 협업 중 삭제된 시트의 "자리 번호"(lazyGrid.sheetMeta의 인덱스).
   * lazyGrid.sheetMeta에서 splice하지 않는 이유는 XlsxExportOptions의 주석 참고(인덱스가 곧
   * Yjs 셀 키의 시트 번호라 밀리면 셀 데이터 전체가 다른 시트를 가리킨다).
   */
  readonly removedSheetIndices: Set<number>;
  /** xlsx 룸에서만 쓰인다 — 시트 자리 번호 → 탭 표시 순서(SheetStructure.order). */
  readonly sheetDisplayOrder: Map<number, number>;
  /** xlsx 룸에서만 쓰인다 — 연결별로 "지금 이 시트의 이 페이지들을 보고 있다"를 추적한다.
   * 연결이 끊기면(closeConn) 이 맵에서 즉시 제거한다. */
  readonly pageSubscriptions: Map<WebSocket, { sheet: number; pages: readonly number[] }>;
  /**
   * xlsx 룸에서만 쓰인다 — canonical 그리드의 **큰 배열 성장**을 순서대로 이어 붙이는 백그라운드
   * 체인. 구조 변경(rowCount/colCount 확장)이 요구하는 배열 할당은 상한(30만 행)까지 갈 수 있어
   * 동기로 돌면 이벤트 루프를 통째로 멈춘다 — 이 프로젝트가 이미 여러 번 고친 그 버그 클래스다.
   *
   * Task 11 이전에는 canonical이 시트 전체를 실제 `GridCell` 배열로 담고 있어(`GridDocument`),
   * 이 체인이 그 배열 자체를 청크+양보로 채우는 무거운 작업(`growCanonicalRows`/`widenCanonicalRows`)
   * 을 순서대로 이어 붙이는 용도였다. Task 11부터는 `LazyGridStore.growSheet`(정수 매핑 배열만
   * `-1`로 패딩)가 그 자리를 대신하므로 이 작업 자체가 항상 가볍다 — 하지만 체인은 그대로
   * 유지한다: 행/열 삽입·삭제(`applyGridStructureOp`)도 같은 시트의 매핑 배열을 건드리므로,
   * 두 종류의 작업이 겹치면 순서가 어긋날 수 있다(계획 문서 Task 11 참고).
   *
   * 하나의 체인으로 묶는 이유는 recomputeSubscribedPagesInBackground와 같다 — 따로 흘려보내면
   * 각 호출이 자기 첫 await까지 동기로 실행돼 양보 없는 블록이 연달아 붙는다.
   */
  canonicalGrowth: Promise<void>;
}

const rooms = new Map<string, Room>();

/**
 * 대용량으로 취급하는 xlsx 원본 파일 크기 기준(바이트) — 이 이상이면 동시에 열 수 있는 룸 개수에
 * 상한을 둔다. 작은 파일은 룸당 메모리 사용이 미미해 상한 대상이 아니다.
 *
 * **왜 행 수가 아니라 바이트 크기인가**: 상한 검사는 룸을 실제로 열기(seedRoom으로 파싱) 전에
 * 해야 의미가 있다 — 파싱까지 해봐야 행 수를 아는 거라면 "상한을 검사하려고 이미 상한을 넘는
 * 메모리를 써버리는" 모순이 생긴다. 파일 바이트 크기는 파싱 없이 디스크에서 바로 알 수 있는
 * 유일한 저비용 대리 지표다.
 */
const LARGE_XLSX_BYTE_THRESHOLD = 5 * 1024 * 1024;

/** 동시에 열 수 있는 "대용량 xlsx 룸" 개수 상한 — 이미 열려 있는 문서에 대한 추가 접속자(같은
 * docId)는 상한과 무관하게 항상 받는다(그 룸은 이미 세어져 있으므로 추가 메모리 비용이 없다). */
const MAX_CONCURRENT_LARGE_XLSX_ROOMS = 3;

/** 지금 대용량으로 카운트되고 있는 xlsx 문서 id 집합 — 룸이 완전히 닫히면(closeConn) 함께 뺀다. */
const largeRoomDocIds = new Set<string>();

/**
 * 새 편집/열람 세션(joinEditSession)을 시작하기 전에 호출한다 — 이미 열려 있는 문서면 항상
 * 허용하고, 새로 열어야 하는 대용량 xlsx 문서인데 이미 상한만큼 열려 있으면 거부한다.
 * 허용되는 경우 그 자리에서 카운트를 예약한다(같은 docId를 두 번 세지는 않는다 — Set).
 */
export async function checkRoomCapacity(
  docId: string,
  ext: string,
): Promise<{ readonly ok: true } | { readonly ok: false; readonly reason: string }> {
  if (rooms.has(docId) || largeRoomDocIds.has(docId)) return { ok: true };
  if (ext !== 'xlsx') return { ok: true };
  const size = await store.originalSize(docId, ext);
  if (size === null || size < LARGE_XLSX_BYTE_THRESHOLD) return { ok: true };
  if (largeRoomDocIds.size >= MAX_CONCURRENT_LARGE_XLSX_ROOMS) {
    return { ok: false, reason: '사용자가 많습니다. 잠시 후 다시 시도해주세요.' };
  }
  largeRoomDocIds.add(docId);
  return { ok: true };
}

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

/**
 * canonical(lazyGrid)이 스스로 넓어질 수 있는 상한. 클라이언트가 보내는 셀 키는 신뢰할 수 없는
 * 입력이므로(조작된 소켓 메시지로 "3:99999999:99999999" 같은 키를 보낼 수 있다), 아래 성장
 * 로직에 반드시 상한이 필요하다 — 없으면 키 하나로 수억 칸짜리 배열을 만들어 서버 메모리를
 * 통째로 날릴 수 있다.
 *
 * **두 값 모두 파서의 상한(spreadsheet.ts의 MAX_ROWS/MAX_COLS)을 그대로 쓴다.** 예전에는 열만
 * 1024로 따로 뒀는데(클라이언트 밀기 루프의 상한 1000을 넉넉히 덮으려는 의도였다), 그건 이
 * 코드베이스가 실제로 파싱하는 열 수(120)의 8.5배라 조작된 메시지 한 통으로 약 2.4GB를
 * 할당시킬 수 있는 창구였다(실측: 300k행 × 1024열 = 2.4GB / 1.1초 정지).
 *
 * 열 상한을 파서와 맞춰도 잃는 것이 없다: canonical이 120열을 넘겨 들고 있어도 저장 후
 * reparseGridForViewer와 다음 룸 열기가 모두 MAX_COLS에서 잘라내므로, 120열 너머는 애초에
 * 편집 세션이 지속적으로 표현할 수 있는 영역이 아니다. 반대로 상한을 파서보다 넉넉히 두면
 * "파일에서는 절대 나올 수 없는 크기"를 서버 메모리에만 만들어낼 수 있게 된다.
 */
const MAX_CANONICAL_ROWS = MAX_ROWS;
const MAX_CANONICAL_COLS = MAX_COLS;
/** 한 워크북이 가질 수 있는 시트 자리 수 상한 — 위와 같은 이유(구조 맵 키도 클라이언트 입력이다). */
const MAX_CANONICAL_SHEETS = 512;

/** row/col 둘 다 신뢰할 수 있는 좌표 범위 안에 있는지 — 셀·서식 관찰자가 각자 시트를 넓히기
 * 전에 반드시 거치는 보안 경계다(위 MAX_CANONICAL_ROWS/COLS 주석 참고, 옛 ensureCanonicalCellSlot의
 * 바로 그 검사). */
function isWithinCanonicalBounds(row: number, col: number): boolean {
  return (
    Number.isInteger(row) &&
    Number.isInteger(col) &&
    row >= 0 &&
    col >= 0 &&
    row < MAX_CANONICAL_ROWS &&
    col < MAX_CANONICAL_COLS
  );
}

/**
 * xlsx 룸의 GRID_CELLS_MAP(Y.Map)에 생긴 변화 하나하나를 lazyGrid.dirtyOverlay에 반영한다.
 * 반드시 `cells.observe(...)`로 등록해서 써야 한다 — Y.Doc의 'update' 이벤트를 훑어 매번 맵
 * 전체를 다시 스캔하는 방식은 삭제(`cells.delete(key)`)를 놓친다: 삭제된 키는 그 시점 맵의
 * `for...of` 순회에 아예 나타나지 않으므로 "이 칸이 지워졌다"는 사실 자체를 관찰할 수 없다.
 * `YMapEvent.changes.keys`는 각 키가 add/update/delete 중 무엇이었는지 직접 알려주므로,
 * 삭제만 별도로 dirtyOverlay에 명시적 0(빈 칸 sentinel)을 써서 되돌리고, 그 외에는 현재 맵
 * 값을 그대로 읽어 반영한다(`dirtyOverlay.delete(key)`가 아니다 — "편집 안 된 칸"으로
 * 되돌아가면 원본 값이 다시 보이는 버그가 된다).
 *
 * **왜 async인가**: 텍스트만 바뀐 편집이라도 그 칸에 이미 걸려 있던 서식(.s)은 보존해야 한다
 * (옛 keptStyle과 동일한 의도) — 그러려면 그 칸의 "현재" 값을 알아야 하는데, LazyGridStore의
 * 읽기(getCell)는 아직 캐시되지 않은 원본 행을 디스크에서 읽어야 할 수도 있어 본질적으로
 * 비동기다. Yjs의 observe 콜백 자체는 동기여야 하므로(중간에 await할 수 없다), 호출부(seedRoom)는
 * 이 Promise를 `void`로 흘려보내고 끝난 뒤에 재계산을 잇는다(Task 10이 recomputeFormulasForPage에
 * 쓴 것과 같은 패턴 — 청크 분류 → Promise.all로 비동기 읽기를 먼저 끝냄 → 동기 쓰기).
 *
 * **왜 관찰자 안에서 즉석으로 넓히는가 (C2의 핵심)**: 클라이언트는 원본 범위 밖 칸에 입력할 때
 * `cells.set(...)`과 `structure.set(...)`(rowCount/colCount 확장)을 **한 트랜잭션 안에서** 하는데,
 * Yjs는 그 트랜잭션의 관찰자를 "그 트랜잭션에서 먼저 건드려진 타입" 순서로 부른다. commitCell은
 * 셀 맵을 먼저 건드리므로 **cells.observe가 structure.observe보다 먼저** 불린다 — 즉 "구조가
 * 먼저 넓어지기를 기다린다"는 설계는 항상 한 박자 늦어 그 편집을 통째로 놓친다. 그래서 순서에
 * 기대지 않고 셀·서식 관찰자가 각자 필요한 만큼(`growSheet`) 직접 넓힌다.
 */
async function applyCellsChangeToCanonical(
  store: LazyGridStore,
  cells: Y.Map<string>,
  event: Y.YMapEvent<string>,
): Promise<void> {
  interface PendingCellChange {
    readonly key: string;
    readonly sheetIdx: number;
    readonly row: number;
    readonly col: number;
    readonly deleted: boolean;
  }
  const changes: PendingCellChange[] = [];
  for (const [key, change] of event.changes.keys) {
    const parsed = parseGridCellKey(key);
    if (!parsed) continue;
    const { sheetIdx, row, col } = parsed;
    if (sheetIdx < 0 || sheetIdx >= store.sheetMeta.length) continue;
    if (!isWithinCanonicalBounds(row, col)) continue;
    changes.push({ key, sheetIdx, row, col, deleted: change.action === 'delete' });
  }
  if (changes.length === 0) return;

  // 값 쓰기(삭제 아님)만 시트를 넓힌다 — 삭제는 절대 시트를 넓히지 않는다(원래 없던 칸이
  // 지워졌다는 건 아무 일도 아니다). Map(dirtyOverlay)은 미리 크기를 맞출 필요가 없으므로,
  // 여기서 늘리는 건 오직 sheetMeta.rowCount/colCount(와 그에 딸린 매핑 배열 길이)뿐이다.
  for (const c of changes) {
    if (!c.deleted) growSheet(store, c.sheetIdx, c.row + 1, c.col + 1);
  }

  // 지금 그 칸의 값을 먼저 읽어 둔다 — 텍스트만 바꾸는 편집이 기존 서식(.s)을 지우지 않기
  // 위함이다. 삭제는 이 값을 쓰지 않으므로 굳이 읽지 않는다(0으로 채워 Promise.all 모양만 맞춘다).
  const currentCells = await Promise.all(
    changes.map((c) => (c.deleted ? Promise.resolve<GridCell | 0>(0) : getCell(store, c.sheetIdx, c.row, c.col))),
  );

  for (let i = 0; i < changes.length; i += 1) {
    const c = changes[i];
    if (c.deleted) {
      const meta = store.sheetMeta[c.sheetIdx];
      if (c.row >= meta.rowCount || c.col >= meta.colCount) continue; // 원래 있던 칸이 아니면 할 일 없음
      setDirtyCell(store, c.sheetIdx, c.row, c.col, 0);
      continue;
    }
    const value = cells.get(c.key) ?? '';
    const current = currentCells[i];
    const keptStyle = current !== 0 && current !== undefined ? current.s : undefined;
    setDirtyCell(
      store,
      c.sheetIdx,
      c.row,
      c.col,
      value === '' ? 0 : { v: value, ...(keptStyle !== undefined ? { s: keptStyle } : {}) },
    );
  }
}

/**
 * canonical 그리드를 크게 건드리는 작업 하나를 룸의 백그라운드 체인 끝에 이어 붙인다
 * (Room.canonicalGrowth 주석 참고).
 *
 * 배열 패딩(성장)뿐 아니라 **행/열 삽입·삭제(applyGridStructureOp)도 반드시 이 체인을 탄다.**
 * 열 방향 연산은 청크마다 이벤트 루프에 양보하므로, 두 연산이 겹치면 한쪽이 아직 밀지 않은 행을
 * 다른 쪽이 밀어 데이터가 어긋난다. 행 방향 연산 자체는 동기지만, 진행 중인 열 연산이 행 배열을
 * 인덱스로 순회하는 중간에 splice로 행을 끼워 넣으면 같은 문제가 생긴다 — 그래서 네 연산과 패딩
 * 전부를 한 줄로 세운다.
 */
function enqueueCanonicalGrowth(room: Room, work: () => Promise<void>): void {
  room.canonicalGrowth = room.canonicalGrowth.then(work).catch((error: unknown) => {
    console.error(`canonical 그리드 갱신 실패 (문서 ${room.docId}):`, error);
  });
}

/**
 * 지금 예약돼 있는 canonical 갱신 작업이 전부 끝날 때까지 기다린다.
 *
 * **내보내기(exportXlsxFromGrid) 직전에는 반드시 이걸 먼저 await해야 한다** — 리뷰 지적 I-1.
 * 열 방향 연산은 청크마다 이벤트 루프에 양보하므로(gridShift.ts의 insertCanonicalCols /
 * deleteCanonicalCols), 그 사이에 끼어든 내보내기는 canonical 그리드를 **앞쪽 행은 이미 밀렸고
 * 뒤쪽 행은 아직 안 밀린** 반쪽 상태로 읽는다. 그렇게 나간 파일은 행마다 열이 다르게 밀려 있다 —
 * 이 계획 전체가 없애려던 바로 그 어긋남이, 화면에는 정상으로 보이는 채로 사용자 파일에 저장된다.
 *
 * 체인이 도는 도중에 새 작업이 덧붙을 수 있으므로(한 구조 연산이 서식 키 재매핑·페이지 재시딩·수식
 * 재계산을 다시 예약한다) 더 이상 늘어나지 않을 때까지 따라간다.
 */
async function drainCanonicalGrowth(room: Room): Promise<void> {
  let seen: Promise<void> | null = null;
  while (seen !== room.canonicalGrowth) {
    seen = room.canonicalGrowth;
    await seen;
  }
}

/**
 * GRID_STYLES_MAP(서식 레이어)의 변화를 lazyGrid.dirtyOverlay에 반영한다 — 리뷰 지적 C1.
 *
 * 이게 없으면 **값을 한 글자도 바꾸지 않고 서식만 바꾼 편집**(굵게·색·표시 형식)이 Yjs에만 남고
 * 저장 대상에는 전혀 도달하지 않아, 저장된 파일에서 조용히 사라진다(다른 참여자 화면에는 보이므로
 * 사용자는 저장됐다고 믿는다 — 가장 나쁜 종류의 손실).
 *
 * 셀은 서식 객체가 아니라 공용 팔레트(store.styles)의 **인덱스**만 들고 있으므로(GridCell.s),
 * 들어온 CellStyleAttrs를 GridStyle로 변환한 뒤 팔레트에서 같은 서식을 찾거나 새로 추가해 그
 * 인덱스를 셀에 심는다. applyCellsChangeToCanonical과 같은 이유로 async다 — "값은 그대로 두고
 * 서식만 지운다/붙인다"는 기존 의미를 지키려면 먼저 그 칸의 현재 값을 읽어야 한다.
 */
async function applyStylesChangeToCanonical(
  store: LazyGridStore,
  styles: Y.Map<CellStyleAttrs>,
  event: Y.YMapEvent<CellStyleAttrs>,
): Promise<void> {
  interface PendingStyleChange {
    readonly key: string;
    readonly sheetIdx: number;
    readonly row: number;
    readonly col: number;
    readonly attrs: CellStyleAttrs | undefined;
  }
  const isMeaningful = (attrs: CellStyleAttrs | undefined): attrs is CellStyleAttrs =>
    !!attrs && typeof attrs === 'object' && Object.keys(attrs).length > 0;

  const changes: PendingStyleChange[] = [];
  for (const [key, change] of event.changes.keys) {
    const parsed = parseGridCellKey(key);
    if (!parsed) continue;
    const { sheetIdx, row, col } = parsed;
    if (sheetIdx < 0 || sheetIdx >= store.sheetMeta.length) continue;
    if (!isWithinCanonicalBounds(row, col)) continue;
    const attrs = change.action === 'delete' ? undefined : styles.get(key);
    changes.push({ key, sheetIdx, row, col, attrs });
  }
  if (changes.length === 0) return;

  // 새 서식이 실제로 붙는 경우만 시트를 넓힌다(옛 ensureCanonicalCellSlot의 grow 역할, 이제
  // growSheet) — 서식을 지우는 쪽은 절대 시트를 넓히지 않는다(아래 루프의 주석 참고).
  for (const c of changes) {
    if (isMeaningful(c.attrs)) growSheet(store, c.sheetIdx, c.row + 1, c.col + 1);
  }

  const currentCells = await Promise.all(changes.map((c) => getCell(store, c.sheetIdx, c.row, c.col)));

  for (let i = 0; i < changes.length; i += 1) {
    const c = changes[i];
    const current = currentCells[i];
    if (!isMeaningful(c.attrs)) {
      // 서식이 지워졌을 때(또는 빈 서식이 들어왔을 때)는 **값은 그대로 두고 서식 참조만 뗀다** —
      // cells.observe의 삭제가 칸을 통째로 0으로 되돌리는 것과 의도적으로 다르다. 저기서는
      // "내용을 지웠다"지만 여기서는 "서식만 지웠다"이므로, 칸을 0으로 만들면 멀쩡한 텍스트까지
      // 날아간다. 값도 원래 비어 있던 칸(서식만 있던 칸)만 완전한 빈 칸(0)으로 되돌린다.
      const meta = store.sheetMeta[c.sheetIdx];
      if (c.row >= meta.rowCount || c.col >= meta.colCount) continue;
      if (current === 0 || current === undefined) continue;
      setDirtyCell(store, c.sheetIdx, c.row, c.col, current.v === '' ? 0 : { v: current.v });
      continue;
    }
    const styleIndex = findOrAddGridStyleIndex(store, fromCellStyleAttrs(c.attrs));
    setDirtyCell(store, c.sheetIdx, c.row, c.col, {
      v: current === 0 || current === undefined ? '' : current.v,
      s: styleIndex,
    });
  }
}

/** 새로 추가된 시트 자리를 채우는 최소한의 빈 시트 메타. */
function makeEmptyLazySheetMeta(name: string): LazyGridSheetMeta {
  return { name, rowCount: 1, colCount: 1, colWidths: [], merges: [], freezeRows: 0, freezeCols: 0 };
}

/**
 * lazyGrid.sheetMeta가 sheetIdx 자리를 갖도록 (필요하면 중간 구멍까지) 늘린다.
 * 구멍(구조 맵에 대응하는 키가 없는 중간 자리)은 "삭제된 자리"로 표시해 내보내기에서 빠지게 한다 —
 * 배열에서 splice하면 그 뒤 시트 번호가 전부 밀려 Yjs 셀 키가 다른 시트를 가리키기 때문이다.
 *
 * 새로 만든 시트 자리는 원본에 없던 시트이므로 logicalToOriginalRow/Col(그리고 rowIndex)에도
 * 그 자리에 빈 매핑 배열을 push해 자리를 맞추고, growSheet(1, 1)로 그 매핑 배열을 즉시
 * rowCount/colCount(각 1)와 같은 길이까지 채운다 — 안 그러면 매핑 배열 길이가 rowCount보다
 * 짧은 채로 남아, 이 시트를 대상으로 한 이후의 applyRowShift/applyColShift(Array.prototype.splice
 * 기반)가 짧은 배열 길이에 clamp되어 엉뚱한 위치에 삽입할 수 있다("매핑 배열 길이 >=
 * rowCount/colCount"는 이 코드베이스가 항상 지키는 불변식이다 — ingestXlsxLazily의 항등 매핑,
 * growSheet의 패딩 모두 이 불변식을 유지한다).
 */
function ensureLazySheetSlot(
  room: Room,
  structure: Y.Map<SheetStructure>,
  sheetIdx: number,
): LazyGridSheetMeta | null {
  const store = room.lazyGrid;
  if (!store || sheetIdx < 0 || sheetIdx >= MAX_CANONICAL_SHEETS) return null;
  while (store.sheetMeta.length <= sheetIdx) {
    const idx = store.sheetMeta.length;
    store.sheetMeta.push(makeEmptyLazySheetMeta(`Sheet${idx + 1}`));
    store.rowIndex.push([]);
    store.logicalToOriginalRow.push([]);
    store.logicalToOriginalCol.push([]);
    growSheet(store, idx, 1, 1);
    if (idx !== sheetIdx && !structure.has(String(idx))) room.removedSheetIndices.add(idx);
  }
  return store.sheetMeta[sheetIdx];
}

/**
 * GRID_STRUCTURE_MAP(시트 구조)의 변화를 lazyGrid에 반영한다 — 리뷰 지적 C2.
 *
 * 시트 이름 변경·추가·삭제·순서 변경과 행/열 수 증감을 모두 여기서 받는다. 삭제와 순서는
 * lazyGrid를 재배치하지 않고 room의 별도 집합/맵에 모아 뒀다가 내보내기 시점에
 * XlsxExportOptions로 넘긴다(인덱스 = 시트 식별자라 배열을 건드릴 수 없다 — 그쪽 주석 참고).
 *
 * 행/열 **데이터의 이동**(행 삽입 시 아래로 밀기 등)은 여기서 하지 않는다 — 그건 별도의
 * MESSAGE_GRID_STRUCTURE_OP 메시지와 applyGridStructureOp가 처리한다. 이 함수가 받는
 * rowCount/colCount 변경은 "사용자가 시트 끝을 넘겨 입력해서 시트가 넓어졌다"(commitCell / 툴바의
 * 행·열 추가 버튼)에 해당하는 **순수한 확장**뿐이고, 기존 데이터의 자리는 그대로다. 두 경로는
 * 서로 겹치지 않는다: 확장은 여기서, 자리를 옮기는 삽입·삭제는 저기서.
 *
 * **Task 11 이전과 달라진 점**: 그때는 배열 패딩(rows 채우기)이 최대 300MB급 무거운 작업이라
 * 반드시 canonicalGrowth 체인에 예약해야 했다(growCanonicalRows/widenCanonicalRows, 청크+양보).
 * `growSheet`(LazyGridStore)는 정수 매핑 배열만 채우므로 이제 항상 동기·즉시 완료된다 — 그래서
 * 이 함수는 **더 이상 `enqueueCanonicalGrowth`로 감싸지 않고 structure.observe에서 직접, 동기로**
 * 부른다(실제로 감싸봤더니 `tests/roomStructureSync.test.ts`의 시트 이름 변경·추가·삭제·순서
 * 변경 테스트가 깨졌다 — 원인을 더 파기보다, 검증된 대로 직접 호출로 되돌렸다). 이 함수 자체가
 * 통째로 동기이므로 실행이 시작되면 중간에 다른 작업이 끼어들 수 없다(JS는 단일 스레드) — 그래서
 * `applyGridStructureOp`(행/열 삽입·삭제, enqueueCanonicalGrowth로 감싸는 별도 비동기 체인)와
 * 뒤섞일 수 있는 건 오직 "어느 쪽이 먼저 완전히 끝나느냐"라는 순서 문제뿐이고, 그 순서는 애초에
 * Yjs 트랜잭션이 도착한 순서를 그대로 따른다(구조 변경 메시지와 확장 메시지가 같은 트랜잭션 안에
 * 섞여 오지 않는 한 겹칠 일이 없다).
 */
function applyStructureChangeToCanonical(
  room: Room,
  structure: Y.Map<SheetStructure>,
  event: Y.YMapEvent<SheetStructure>,
): void {
  if (!room.lazyGrid) return;
  const store = room.lazyGrid;
  for (const [key, change] of event.changes.keys) {
    const sheetIdx = Number(key);
    if (!Number.isInteger(sheetIdx) || sheetIdx < 0 || sheetIdx >= MAX_CANONICAL_SHEETS) continue;
    if (change.action === 'delete') {
      if (sheetIdx < store.sheetMeta.length) room.removedSheetIndices.add(sheetIdx);
      continue;
    }
    const value = structure.get(key);
    if (!value || typeof value !== 'object') continue;
    // 지웠던 자리를 다시 쓰면(되돌리기 등) 삭제 표시를 거둔다.
    room.removedSheetIndices.delete(sheetIdx);
    const meta = ensureLazySheetSlot(room, structure, sheetIdx);
    if (!meta) continue;
    if (typeof value.name === 'string' && value.name.trim() !== '') meta.name = value.name;
    if (Number.isInteger(value.rowCount) && value.rowCount > 0) {
      const rowCount = Math.min(value.rowCount, MAX_CANONICAL_ROWS);
      // 행이 늘어난 경우에만 매핑 배열을 growSheet로 채운다 — 줄어든 경우엔 rowCount만
      // 낮춘다(내보내기가 그 뒤를 무시하고 ExcelJS 쪽 꼬리 행도 잘라낸다). 매핑 배열을 실제로
      // 자르면 그 사이 도착한 편집(dirtyOverlay)의 좌표 체계가 깨진다 — growSheet 주석과 동일한 원칙.
      if (rowCount > meta.rowCount) {
        growSheet(store, sheetIdx, rowCount, meta.colCount);
      } else {
        meta.rowCount = rowCount;
      }
    }
    if (Number.isInteger(value.colCount) && value.colCount > 0) {
      const colCount = Math.min(value.colCount, MAX_CANONICAL_COLS);
      if (colCount > meta.colCount) {
        growSheet(store, sheetIdx, meta.rowCount, colCount);
      } else {
        meta.colCount = colCount;
      }
    }
    if (Number.isInteger(value.freezeRows)) meta.freezeRows = value.freezeRows as number;
    if (Number.isInteger(value.freezeCols)) meta.freezeCols = value.freezeCols as number;
  }
  rebuildSheetDisplayOrder(room, structure);
}

/** 구조 맵의 order 필드를 모아 "시트 자리 번호 → 탭 순서" 맵을 다시 만든다 — 시트 수가 많아야
 * 수십 개라 변경 하나마다 전부 다시 만들어도 싸고, 부분 갱신보다 어긋날 여지가 없다. */
function rebuildSheetDisplayOrder(room: Room, structure: Y.Map<SheetStructure>): void {
  room.sheetDisplayOrder.clear();
  for (const [key, value] of structure.entries()) {
    const sheetIdx = Number(key);
    if (!Number.isInteger(sheetIdx) || sheetIdx < 0) continue;
    const order = value?.order;
    if (typeof order === 'number' && Number.isFinite(order)) room.sheetDisplayOrder.set(sheetIdx, order);
  }
}

/** 지금 룸 상태로 xlsx를 내보낼 때 함께 넘길 구조 변경 정보(시트 삭제/탭 순서). */
function exportOptionsFor(room: Room) {
  return {
    removedSheetIndices: room.removedSheetIndices,
    sheetDisplayOrder: room.sheetDisplayOrder,
  };
}

/**
 * xlsx 내보내기의 유일한 진입점 — 정상 경로(exportXlsxFromGrid)를 먼저 시도하고, 그게 실패하면
 * (ExcelJS의 대용량 재직렬화 한계, 예: `RangeError: Invalid array length` — exportStreaming.ts
 * 주석 참고) 스트리밍 폴백으로 자동 재시도한다.
 *
 * **왜 사전 임계값이 아니라 "실패하면 재시도"인가**: 실패 여부는 행 수뿐 아니라 서식 밀도 등
 * 여러 요인에 좌우돼 미리 정확한 기준선을 잡기 어렵다. 이 순서는 절대다수인 "정상 크기" 파일에는
 * 전혀 영향을 주지 않고(항상 검증된 기존 경로만 탄다), 실제로 그 경로가 감당 못하는 극단적으로
 * 큰 파일에서만 스트리밍 경로가 관여한다.
 *
 * 스트리밍 폴백은 스코프가 좁다(구조 연산 미지원 등, exportStreaming.ts 주석 참고) — 그것도
 * 실패하면(예: 구조 연산이 있어서) 그 에러를 그대로 던진다. 호출부(flushToStore/getCurrentBytes)
 * 는 이미 실패를 정상적으로 처리한다(저장 실패 표시, 에러 로그).
 */
async function exportXlsxWithFallback(room: Room, originalXlsxBytes: Uint8Array): Promise<Uint8Array> {
  const store = room.lazyGrid!;
  const options = exportOptionsFor(room);
  try {
    return await exportXlsxFromGrid(
      originalXlsxBytes,
      store,
      room.baselineSheetExtents!,
      room.structuralOps,
      options,
    );
  } catch (error) {
    console.error(
      `xlsx 내보내기 실패, 스트리밍 폴백으로 재시도 (문서 ${room.docId}):`,
      error,
    );
    return exportXlsxFromGridStreaming(originalXlsxBytes, store, room.structuralOps, options);
  }
}

/**
 * 시트 하나의 **페이지 하나**(PAGE_SIZE_ROWS행) 안에 있는 모든 수식 셀을 canonical 그리드 기준으로
 * 다시 계산해 GRID_COMPUTED_MAP에 반영한다. 원본 셀 값(GRID_CELLS_MAP)은 절대 건드리지 않는다 —
 * 계산된 표시값만 별도 맵에 쓴다. 더 이상 수식이 아니거나(빈 칸 포함) 존재하지 않는 셀의 옛
 * computed 값은 지운다 — 계산값이 GRID_CELLS_MAP으로 새지 않고, 낡은 값도 남지 않는다.
 *
 * **범위가 페이지 단위인 이유**(윈도잉 전에는 시트 전체를 돌았다): 계산 결과를 쓰는 곳이
 * GRID_COMPUTED_MAP, 즉 "지금 구독 중인 페이지만 담는 창"이다. 아무도 보고 있지 않은 페이지의
 * 계산값을 써 넣으면 창 밖의 데이터가 Yjs에 다시 쌓여 이 태스크가 없애려는 그 메모리·전파
 * 비용이 그대로 돌아온다(편집 한 번마다 시트 전체를 다시 계산하는 CPU 비용도 함께). 단, 수식은
 * 다른 페이지의 셀을 참조할 수 있으므로 "편집된 셀이 속한 페이지"만 다시 계산하는 것으로는
 * 부족하다 — 호출부(cells.observe)는 **구독 중인 모든 페이지**를 대상으로 이 함수를 부른다.
 * 계산 자체는 항상 canonical 전체를 참조하므로(computeGridCellDisplay), 창 밖의 값을 참조하는
 * 수식도 정확히 계산된다.
 *
 * 넓은 범위를 참조하는 수식이 많은 시트를 훑는 동안 이벤트 루프를 막지 않도록
 * YIELD_EVERY_N_ROWS행마다 실제 이벤트 루프에 양보한다 — 대형 xlsx가 서버를 통째로 멈추던
 * 버그(6355a17)가 이 반복문에서 재발하지 않도록 하는 필수 조치.
 *
 * 반드시 `cells.observe(...)`(아래 seedRoom에 등록)가 호출해야 한다 — Y.Doc의 'update'
 * 이벤트가 아니다. 이 함수 자신이 GRID_COMPUTED_MAP에 쓰는 매 `ydoc.transact(...)`도 그
 * Y.Doc의 'update'를 발생시키므로, 만약 'update' 핸들러 안에서 이 함수를 다시 트리거했다면
 * "재계산 결과가 스스로 또 재계산을 유발"하는 무한 루프가 됐을 것이다(자동저장 무한 루프와
 * 같은 클래스의 버그). `cells.observe`는 GRID_CELLS_MAP이 실제로 바뀔 때만 불리고
 * GRID_COMPUTED_MAP 쓰기에는 반응하지 않으므로, 이 재귀 트리거가 원천적으로 없다 — 그래서
 * 호출부에서도 room.suppressAutosaveReschedule로 감쌀 필요가 없다(각 computed 쓰기가
 * scheduleAutosave를 다시 예약하는 건 8초 디바운스로 자연히 합쳐지는 무해한 부작용일 뿐이고,
 * 반대로 "안전하게" 감쌌다가는 재계산이 진행되는 동안 도착한 진짜 편집의 자동저장 예약을
 * 놓치는 별개의 버그를 새로 만든다).
 */
/**
 * 페이지 하나를 재계산할 때 쓰는 양보 간격 — 리뷰 지적 I3.
 *
 * 이 루프에 YIELD_EVERY_N_ROWS(2000)를 그대로 쓰면 페이지 크기(500)가 그보다 작아 **한 페이지 =
 * 양보 한 번 없는 500행짜리 단일 동기 블록**이 된다. 페이지 사이의 양보(Task 7에서 이미 고친 것)는
 * 페이지를 여러 개 구독했을 때만 의미가 있고, 페이지 하나만 구독한 클라이언트가 넓은 범위를
 * 참조하는 수식이 가득한 500행을 편집하면 그 계산이 끝날 때까지 서버 전체가 그대로 멈춘다.
 * 페이지 크기의 1/4로 잘라 페이지 **안에서도** 반드시 몇 번은 쉬게 한다. YIELD_EVERY_N_ROWS와
 * min을 취하는 이유는 두 상수 중 어느 쪽이 나중에 바뀌어도 "더 촘촘한 쪽"이 이기게 하기 위함이다.
 */
const RECOMPUTE_YIELD_EVERY_N_ROWS = Math.max(1, Math.min(YIELD_EVERY_N_ROWS, Math.ceil(PAGE_SIZE_ROWS / 4)));

async function recomputeFormulasForPage(room: Room, sheetIdx: number, page: number): Promise<void> {
  const store = room.lazyGrid;
  if (!store) return;
  const meta = store.sheetMeta[sheetIdx];
  if (!meta) return;
  const computed = room.ydoc.getMap<string>(GRID_COMPUTED_MAP);
  const pageStart = page * PAGE_SIZE_ROWS;
  const pageEnd = Math.min(pageStart + PAGE_SIZE_ROWS, meta.rowCount);
  for (let start = pageStart; start < pageEnd; start += RECOMPUTE_YIELD_EVERY_N_ROWS) {
    const end = Math.min(start + RECOMPUTE_YIELD_EVERY_N_ROWS, pageEnd);
    // "이 칸이 수식인가"를 판정하려면 원본 텍스트가 필요하다 — getCell은 더티 오버레이를 먼저
    // 보고 없으면 디스크/캐시에서 읽으므로, 편집된 칸이든 원본 그대로인 칸이든 항상 정확한
    // "지금" 텍스트를 준다(seedGridPageFromXlsx의 resolveGridCell과 같은 규칙, xlsx/collab.ts 참고).
    // 청크 안의 모든 좌표를 한 번에 Promise.all로 흘려보내 디스크 왕복을 겹치게 한다.
    const coords: Array<{ r: number; c: number }> = [];
    for (let r = start; r < end; r += 1) {
      for (let c = 0; c < meta.colCount; c += 1) coords.push({ r, c });
    }
    const cellsAtCoords = await Promise.all(coords.map(({ r, c }) => getCell(store, sheetIdx, r, c)));

    const nonFormulaKeys: string[] = [];
    const formulaCells: Array<{ r: number; c: number; key: string }> = [];
    for (let i = 0; i < coords.length; i += 1) {
      const { r, c } = coords[i];
      const cell = cellsAtCoords[i];
      const key = gridCellKey(sheetIdx, r, c);
      if (cell === 0 || cell === undefined || !cell.v.startsWith('=')) {
        nonFormulaKeys.push(key);
      } else {
        formulaCells.push({ r, c, key });
      }
    }
    // Yjs의 transact() 콜백은 반드시 동기여야 한다(비동기 콜백을 넘기면 첫 await에서 콜백이
    // 일시 정지하는 순간 트랜잭션이 이미 닫혀버려, 그 뒤의 쓰기가 트랜잭션 밖에서 개별적으로
    // 일어난다) — 그래서 computeGridCellDisplayAsync는 반드시 transact 밖에서 먼저 끝낸 뒤,
    // 이미 계산된 값만 동기 콜백 안에서 쓴다(seedGridPageFromXlsx가 getRow를 Promise.all로 먼저
    // 기다린 뒤 transact하는 것과 같은 패턴, xlsx/collab.ts 참고).
    const values = await Promise.all(
      formulaCells.map(({ r, c }) => computeGridCellDisplayAsync(store, sheetIdx, r, c)),
    );
    room.ydoc.transact(() => {
      for (const key of nonFormulaKeys) {
        if (computed.has(key)) computed.delete(key);
      }
      for (let i = 0; i < formulaCells.length; i += 1) {
        computed.set(formulaCells[i].key, formatComputedValueForStorage(values[i]));
      }
    }, GRID_COMPUTED_WRITE_ORIGIN);
    await yieldToEventLoop();
  }
}

/**
 * 편집 하나가 유발하는 "구독 중인 모든 페이지 재계산"을, 페이지마다 **순서대로 await하는 하나의
 * 백그라운드 체인**으로 돌린다.
 *
 * **왜 페이지별로 `void`를 따로 붙이면 안 되는가** — 이 프로젝트가 이미 두 번 고친 "이벤트 루프
 * 통째로 멈춤" 버그가 그대로 되살아난다. recomputeFormulasForPage는 첫 청크를 **동기로** 돌고
 * 나서야 첫 `await yieldToEventLoop()`에 닿는데, 페이지 크기(PAGE_SIZE_ROWS=500)가
 * YIELD_EVERY_N_ROWS(2000)보다 작아 한 페이지의 재계산은 내부적으로 한 번도 양보하지 않는
 * 500행짜리 단일 동기 블록이다. 그런 호출을 `void f(); void f(); ...`처럼 동기 루프에서 줄줄이
 * 시작하면 각 호출이 자기 첫 await까지 **동기로** 실행되므로, N페이지를 구독한 클라이언트가 셀
 * 하나를 편집하는 순간 N×500행의 수식 평가가 양보 지점 하나 없이 연달아 붙어 돈다(= 구독 범위
 * 전체를 한 호흡에 재계산하는 동안 서버 전체가 멈춘다).
 *
 * 하나의 체인으로 묶어 `await`하면 앞 페이지의 Promise가 (그 페이지 자신의
 * `await yieldToEventLoop()`을 거쳐) 실제로 해소된 뒤에 다음 페이지가 시작하므로, 페이지의 동기
 * 블록 사이마다 이벤트 루프가 확실히 한 턴 돈다. 바깥의 `void`는 여전히 하나뿐이라 호출부
 * (cells.observe 콜백)는 예전과 똑같이 즉시 반환한다 — 여러 페이지를 시딩할 때
 * reconcilePageSubscriptions가 이미 쓰고 있는 것과 같은 패턴이다.
 */
function recomputeSubscribedPagesInBackground(room: Room, sheetIndices: Iterable<number>): void {
  void (async () => {
    for (const sheetIdx of sheetIndices) {
      for (const page of getSubscribedPageUnion(room, sheetIdx)) {
        // 페이지 하나가 실패해도 나머지 페이지는 계속 재계산한다 — 그리고 버려진 Promise의
        // reject가 프로세스를 죽이지 않도록 반드시 여기서 잡는다(scheduleAutosave의 같은 주석 참고).
        await recomputeFormulasForPage(room, sheetIdx, page).catch((error: unknown) => {
          console.error(
            `수식 재계산 실패 (문서 ${room.docId}, 시트 ${sheetIdx}, 페이지 ${page}):`,
            error,
          );
        });
      }
    }
  })().catch((error: unknown) => {
    console.error(`수식 재계산 체인 실패 (문서 ${room.docId}):`, error);
  });
}

/** MESSAGE_GRID_STRUCTURE_OP의 페이로드 — "이 시트의 at 자리에 행/열 count개를 넣거나 뺀다".
 * gridShared.ts의 GridStructureOpRecord와 모양이 같다(같은 타입을 재사용해도 되지만, 이건
 * "지금 막 도착한, 아직 적용 안 된 메시지"라는 의미를 이름으로 구분해 둔다). */
interface GridStructureOp {
  readonly sheet: number;
  readonly kind: GridStructureOpKind;
  readonly at: number;
  readonly count: number;
}

const GRID_STRUCTURE_OP_KINDS: readonly string[] = ['insertRows', 'deleteRows', 'insertCols', 'deleteCols'];

/**
 * 한 메시지가 한 번에 요구할 수 있는 행/열 개수 상한.
 *
 * 상한이 **반드시** 필요하다: count는 신뢰할 수 없는 입력이고, 삽입 경로는 그 수만큼 새 행 배열을
 * 동기로 만든다 — count에 1e9가 들어오면 메시지 한 통으로 서버 메모리를 통째로 날릴 수 있다
 * (이 프로젝트가 MESSAGE_PAGE_SUBSCRIBE에서 이미 한 번 막은 것과 같은 클래스의 창구).
 * 실제 UI는 항상 1을 보내고(우클릭 "위에 행 삽입"), 여러 행이 필요하면 여러 번 부르면 된다.
 */
const MAX_STRUCTURE_OP_COUNT = 500;

/**
 * MESSAGE_GRID_STRUCTURE_OP 페이로드 검증 — MESSAGE_PAGE_SUBSCRIBE(isValidPageSubscription)와
 * 완전히 같은 원칙이다: **신뢰할 수 없는 입력**이고, 여기서 통과한 숫자가 그대로 canonical 그리드
 * splice의 인자와 배열 할당 크기가 된다. 하나라도 어긋나면 메시지 전체를 거부한다(보정하지 않는다 —
 * 보정하면 클라이언트가 요청한 것과 서버가 한 것이 조용히 달라진다).
 *
 * 여기서는 **페이로드 자체의 형식**만 본다. `at`이 시트 안쪽인지 같은 룸 상태에 달린 검사는
 * applyGridStructureOp가 실제로 적용하기 직전에 한다 — 이 연산들은 하나의 백그라운드 체인에
 * 줄을 서므로(enqueueCanonicalGrowth) 메시지가 도착한 시점의 rowCount로 검사해 두면 앞선
 * 연산이 그 값을 바꿔 놓은 뒤에 적용될 수 있다.
 */
function isValidGridStructureOp(parsed: unknown): parsed is GridStructureOp {
  if (!parsed || typeof parsed !== 'object') return false;
  const { sheet, kind, at, count } = parsed as { sheet?: unknown; kind?: unknown; at?: unknown; count?: unknown };
  if (typeof kind !== 'string' || !GRID_STRUCTURE_OP_KINDS.includes(kind)) return false;
  // Number.isInteger는 NaN·Infinity·비정수·문자열을 한 번에 걸러 준다.
  if (!Number.isInteger(sheet) || (sheet as number) < 0 || (sheet as number) >= MAX_CANONICAL_SHEETS) return false;
  if (!Number.isInteger(at) || (at as number) < 0) return false;
  if (!Number.isInteger(count) || (count as number) < 1 || (count as number) > MAX_STRUCTURE_OP_COUNT) return false;
  return true;
}

/** 지금 실제로 Yjs에 셀이 올라와 있는 페이지 번호 집합(그 시트에 대해). */
function seededPagesInYdoc(room: Room, sheetIdx: number): Set<number> {
  const pages = new Set<number>();
  for (const key of room.ydoc.getMap<string>(GRID_CELLS_MAP).keys()) {
    const parsed = parseGridCellKey(key);
    if (parsed && parsed.sheetIdx === sheetIdx) pages.add(Math.floor(parsed.row / PAGE_SIZE_ROWS));
  }
  return pages;
}

/**
 * canonical이 밀린 뒤, Yjs에 올라와 있는 그 시트의 페이지들을 **새 canonical 기준으로 다시** 채운다.
 *
 * 왜 통째로 다시 채우는가: 행이 하나 끼워지면 그 아래 **모든** 행의 내용이 한 칸씩 달라지고,
 * 열이 끼워지면 모든 페이지의 모든 행이 달라진다 — 어떤 키가 어떻게 바뀌었는지 따로 계산하는 것보다
 * "이 페이지를 비우고 새 진실에서 다시 채운다"가 훨씬 단순하고 어긋날 여지가 없다. 비우기·채우기·
 * 재계산 세 원시 연산은 모두 Task 7이 이미 쓰고 있는 것 그대로이고, 각자 내부에서 이벤트 루프에
 * 양보한다.
 *
 * 대상은 "구독 중인 페이지"만이 아니라 **지금 Yjs에 올라와 있는 모든 페이지**다: 구독에서 빠졌지만
 * 아직 유예 시간(PAGE_UNSEED_GRACE_MS) 안이라 남아 있는 페이지도 옛 데이터를 들고 있으므로,
 * 그대로 두면 클라이언트가 유예 시간 안에 그 페이지로 되돌아갔을 때 밀리기 전 내용을 본다.
 * 그런 페이지는 지금 비우기만 하고 다시 채우지 않는다(아무도 안 보고 있으므로 — 다시 구독되는
 * 순간 reconcilePageSubscriptions가 새 canonical에서 채운다).
 */
async function refreshSeededPagesAfterShift(
  room: Room,
  sheetIdx: number,
  rowExtentBefore: number,
): Promise<void> {
  const store = room.lazyGrid;
  if (!store) return;
  const meta = store.sheetMeta[sheetIdx];
  if (!meta) return;
  const subscribed = getSubscribedPageUnion(room, sheetIdx);
  const pages = [...new Set<number>([...subscribed, ...seededPagesInYdoc(room, sheetIdx)])].sort((a, b) => a - b);
  // 비울 행 범위는 삽입·삭제 **전후 중 더 넓은 쪽**까지 잡아야 한다 — 삭제로 시트가 짧아진 경우
  // 새 길이로 자르면 옛 꼬리 행의 키가 Yjs에 남아, 이제 존재하지 않는 행이 화면에 계속 보인다.
  // (Task 11 이전에는 sheet.rows.length도 함께 봤다 — 그건 canonicalGrid가 배열을 미리 늘려둔
  // 만큼 여유가 있을 수 있어서였다. LazyGridStore는 그런 여분의 물리적 배열이 없으므로
  // rowExtentBefore/현재 rowCount 두 값의 더 큰 쪽이면 충분하다.)
  const rowExtent = Math.max(rowExtentBefore, meta.rowCount);

  for (const page of pages) {
    await unseedGridPageFromXlsx(room.ydoc, sheetIdx, page, rowExtent);
    if (rooms.get(room.docId) !== room) return;
    if (!subscribed.has(page)) {
      // 방금 비웠으니 예약돼 있던 지연 제거는 할 일이 없다 — 파괴된 Y.Doc을 건드리는 타이머를
      // 남기지 않도록 취소한다(cancelPendingUnseeds와 같은 취지).
      cancelPendingUnseed(room.docId, sheetIdx, page);
      continue;
    }
    await seedGridPageFromXlsx(room.ydoc, store, sheetIdx, page);
    await recomputeFormulasForPage(room, sheetIdx, page);
    if (rooms.get(room.docId) !== room) return;
  }
}

/**
 * 행/열 삽입·삭제를 canonical 그리드에 적용하고, 그 결과를 Yjs에도 반영한다 — 이 함수가 이번
 * 재설계의 핵심이다.
 *
 * 순서에 의도가 있다: **구조 메타데이터(rowCount/colCount)를 맨 마지막에** 쓴다. 클라이언트의
 * structure.observe는 표를 통째로 다시 그리므로, 그걸 먼저 보내면 "행 수는 늘었는데 데이터는 아직
 * 옛 자리"인 상태가 한 번 화면에 보인다. 셀·서식·병합을 먼저 맞춰 두고 마지막에 구조를 알리면
 * 그 재렌더가 곧 완성된 상태의 재렌더가 된다.
 *
 * 자동저장은 별도로 예약하지 않는다 — GRID_STRUCTURE_SHIFT_ORIGIN은 (창 시딩과 달리)
 * isAutosaveExemptOrigin에 들어 있지 않으므로, 위 쓰기들이 그대로 ydoc 'update'를 타고
 * scheduleAutosave를 건다.
 */
async function applyGridStructureOp(room: Room, op: GridStructureOp): Promise<void> {
  if (rooms.get(room.docId) !== room) return; // 룸이 이미 닫혔다(체인에 줄 서 있는 동안)
  const store = room.lazyGrid;
  if (!store) return;
  const meta = store.sheetMeta[op.sheet];
  if (!meta) return;
  if (room.removedSheetIndices.has(op.sheet)) return; // 협업 중 삭제된 시트 자리

  const onRows = op.kind === 'insertRows' || op.kind === 'deleteRows';
  const insert = op.kind === 'insertRows' || op.kind === 'insertCols';
  const extent = onRows ? meta.rowCount : meta.colCount;
  const limit = onRows ? MAX_CANONICAL_ROWS : MAX_CANONICAL_COLS;

  if (insert) {
    // at === extent(시트 맨 끝에 붙이기)는 허용한다 — 그보다 뒤는 정상 클라이언트가 보낼 수 없다.
    if (op.at > extent) return;
    // 상한을 넘기면서 **마지막 행/열의 데이터를 조용히 떨어뜨리는** 대신 연산 자체를 거부한다.
    // (열 상한 MAX_CANONICAL_COLS는 파서의 MAX_COLS=120과 같은 값이라, 정확히 120열까지 꽉 찬
    //  시트에서는 열 삽입이 거부된다 — 그 너머는 저장해도 다음 파싱에서 잘려 나가는 영역이다.)
    if (extent + op.count > limit) return;
  } else {
    if (op.at >= extent) return;
    if (op.at + op.count > extent) return;
    if (extent - op.count < 1) return; // 마지막 남은 행/열은 지울 수 없다(클라이언트 가드와 동일)
  }

  const rowExtentBefore = meta.rowCount;

  // logicalToOriginalRow/Col 매핑과 dirtyOverlay 키 이동, rowCount/colCount 갱신은
  // applyRowShift/applyColShift(lazyGridStore.ts, Task 6)가 전부 동기로 처리한다 — 더 이상
  // 청크+양보가 필요 없다(정수 매핑 배열 스플라이스일 뿐이므로, Room.canonicalGrowth 주석 참고).
  // 다만 그 두 함수는 merges/colWidths까지는 건드리지 않으므로(시트 데이터 이동만 책임진다),
  // 이 자리에서 gridShift.ts의 shiftMerges를 그대로 재사용해 직접 보정한다 — Task 11 이전에
  // insertCanonicalRows 등이 내부적으로 하던 바로 그 일이다.
  if (onRows) {
    applyRowShift(store, op.sheet, op.at, op.count, insert);
  } else {
    applyColShift(store, op.sheet, op.at, op.count, insert);
  }
  meta.merges = shiftMerges(meta.merges, onRows ? 'row' : 'col', op.at, op.count, insert);
  if (!onRows) {
    // 열 너비도 gridShift.ts의 옛 insertCanonicalCols/deleteCanonicalCols와 같은 규칙으로 직접
    // 보정한다 — 새로 끼워지는 열은 밀려나는 그 열의 너비를 그대로 물려받는다(엑셀과 같은 동작).
    const widths = meta.colWidths;
    if (insert) {
      if (widths.length > op.at) {
        meta.colWidths = [...widths.slice(0, op.at), ...new Array<number>(op.count).fill(widths[op.at]), ...widths.slice(op.at)];
      }
    } else if (widths.length > op.at) {
      meta.colWidths = [...widths.slice(0, op.at), ...widths.slice(op.at + op.count)];
    }
  }

  // 이 연산을 기록한다 — 내보내기(exportXlsxFromGrid)가 원본을 다시 로드한 워크북에 같은 순서로
  // 재생하기 위한 로그다.
  room.structuralOps.push({ sheet: op.sheet, kind: op.kind, at: op.at, count: op.count });

  // 1) 병합 — canonical 쪽은 바로 위에서 이미 보정했고, 클라이언트가 렌더링에 쓰는 Yjs 맵도
  //    같은 값으로 맞춘다(이 맵은 윈도잉 대상이 아니라 항상 전체가 들어 있다).
  const mergesMap = room.ydoc.getMap<readonly GridMergeShape[]>(GRID_MERGES_MAP);
  const mergesKey = String(op.sheet);
  room.ydoc.transact(() => {
    if (meta.merges.length > 0) mergesMap.set(mergesKey, meta.merges);
    else if (mergesMap.has(mergesKey)) mergesMap.delete(mergesKey);
  }, GRID_STRUCTURE_SHIFT_ORIGIN);

  // 2) 서식 키 재매핑 — 서식 맵은 윈도잉하지 않으므로 페이지 재시딩으로는 옮겨지지 않는다.
  await shiftGridStyleKeys(
    room.ydoc,
    op.sheet,
    onRows ? 'row' : 'col',
    op.at,
    op.count,
    insert,
    MAX_CANONICAL_ROWS,
    MAX_CANONICAL_COLS,
  );
  if (rooms.get(room.docId) !== room) return;

  // 3) 셀 값·계산값 — Yjs에 올라와 있는 페이지를 새 canonical 기준으로 다시 채운다.
  await refreshSeededPagesAfterShift(room, op.sheet, rowExtentBefore);
  if (rooms.get(room.docId) !== room) return;

  // 4) 구조 메타데이터(맨 마지막 — 위 주석 참고). 기존 항목을 펼쳐 이름·틀고정·탭 순서를 보존하고
  //    이 연산이 실제로 바꾼 것(행/열 수, 열 너비)만 덮어쓴다.
  const structure = room.ydoc.getMap<SheetStructure>(GRID_STRUCTURE_MAP);
  const previous = structure.get(mergesKey);
  const next: SheetStructure = {
    ...(previous ?? {
      name: meta.name,
      freezeRows: meta.freezeRows,
      freezeCols: meta.freezeCols,
    }),
    rowCount: meta.rowCount,
    colCount: meta.colCount,
    colWidths: meta.colWidths,
  };
  room.ydoc.transact(() => {
    structure.set(mergesKey, next);
  }, GRID_STRUCTURE_SHIFT_ORIGIN);
}

async function seedRoom(room: Room): Promise<void> {
  const { docId, kind, ydoc } = room;
  try {
    const meta = await store.readMeta(docId);
    if (!meta) return;
    const original = await store.readOriginal(docId, meta.ext);
    if (kind === 'xlsx') {
      const lazyStore = await ingestXlsxLazily(docId, original);
      room.lazyGrid = lazyStore;
      // 이후 모든 내보내기의 기준 원본 — 왜 매번 디스크를 다시 읽지 않는지는 Room.originalXlsxBytes 주석 참고.
      room.originalXlsxBytes = new Uint8Array(original);
      // 전체 그리드가 아니라 시트별 크기만 남긴다 — Room.baselineSheetExtents 주석 참고.
      room.baselineSheetExtents = lazyStore.sheetMeta.map((s) => ({ rowCount: s.rowCount, colCount: s.colCount }));
      // 셀 데이터는 여기서 채우지 않는다(윈도잉 컷오버) — 구조/병합 메타데이터만 전량 채우고,
      // 실제 셀은 클라이언트가 "이 페이지들을 보고 있다"고 알려올 때 그 페이지만 채운다
      // (reconcilePageSubscriptions). 룸이 막 생긴 이 시점에는 구독자가 하나도 없으므로 채울
      // 페이지도 없다. 큰 xlsx를 열 때 전체 셀을 Yjs에 심으며 서버가 멈추던 문제와, 그 전체를
      // 모든 클라이언트에 전파하던 비용이 여기서 사라진다.
      await seedGridStructureOnly(ydoc, lazyStore.sheetMeta);
      // 서식도 구조 메타데이터와 같은 부류다 — **절대 윈도잉하지 않고** 여기서 전량 한 번만 채운
      // 뒤 어떤 unseed 경로에서도 지우지 않는다. 서식을 셀 값처럼 윈도잉하면, 서버에 서식 편집을
      // canonical로 되반영하는 경로가 없기 때문에 페이지가 비워졌다 다시 채워지는 순간 사용자가
      // 적용해 둔 굵게·색·표시 형식이 조용히 원본 서식으로 되돌아간다
      // (seedGridStylesFromXlsx 주석 참고). LazyGridStore 기준으로는 이 한 번의 전량 시딩이
      // 시트 전체를 디스크에서 훑는 유일한 지점이다(스펙 문서의 알려진 특성 — 서식은 어차피
      // 한 번은 전량 읽어야 한다, 다 읽고 나면 Yjs로 흘려보내고 버린다).
      await seedGridStylesFromXlsx(ydoc, lazyStore);
      const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
      cells.observe((event) => {
        // 서버 자신이 창을 채우거나(seed) 비운(unseed) 변경은 사용자 편집이 아니므로 canonical에
        // 되반영하지 않는다 — 특히 unseed는 "키 삭제"라, 걸러내지 않으면 아래
        // applyCellsChangeToCanonical이 "사용자가 그 500행을 다 지웠다"로 해석해 canonical을
        // 비우고 그게 그대로 파일에 저장된다(자세한 설명은 GRID_WINDOW_SEED_ORIGIN 주석 참고).
        // 새로 채운 페이지의 수식 계산은 reconcilePageSubscriptions가 직접 부르므로 여기서
        // 다시 걸 필요도 없다.
        if (event.transaction.origin === GRID_WINDOW_SEED_ORIGIN) return;
        // applyCellsChangeToCanonical은 이제 async다(LazyGridStore.getCell이 디스크를 읽어야 할
        // 수도 있어서) — Yjs observe 콜백 자체는 동기여야 하므로 흘려보내고, 재계산은 그 완료를
        // 기다린 뒤 잇는다(재계산이 더티 오버레이가 아직 안 채워진 값을 볼 수 없도록).
        void applyCellsChangeToCanonical(lazyStore, cells, event)
          .then(() => {
            // 이번 배치에서 실제로 바뀐 셀이 속한 시트만 재계산한다 — event.changes.keys가
            // "시트:행:열" 키별로 add/update/delete를 알려주므로 거기서 시트 인덱스를 뽑아낸다.
            // 재계산 대상은 그 시트 전체가 아니라 "지금 구독 중인 페이지들"이다 — 수식이 창 밖의
            // 셀을 참조할 수 있으니 편집된 페이지 하나로는 부족하지만, 아무도 안 보는 페이지까지
            // 계산해 Yjs에 쌓을 이유는 없다(recomputeFormulasForPage 주석 참고).
            const touchedSheets = new Set<number>();
            for (const key of event.changes.keys.keys()) {
              const parsed = parseGridCellKey(key);
              if (parsed) touchedSheets.add(parsed.sheetIdx);
            }
            // 페이지마다 따로 흘려보내지 않고 **하나의 체인**으로 순서대로 돈다 — 그러지 않으면
            // 여러 페이지의 동기 재계산 블록이 양보 없이 연달아 붙어 이벤트 루프가 멈춘다
            // (recomputeSubscribedPagesInBackground 주석 참고).
            recomputeSubscribedPagesInBackground(room, touchedSheets);
          })
          .catch((error: unknown) => {
            console.error(`셀 변경 반영 실패 (문서 ${room.docId}):`, error);
          });
      });
      // 서식(C1)과 구조(C2)도 셀 값과 똑같이 canonical에 되반영한다 — 셀 값만 관찰하던 것이
      // "서식만 바꾼 편집"과 "원본 범위 밖 편집·모든 구조 변경"이 저장된 파일에서 조용히 사라지던
      // 원인이었다. 최초 전량 시딩(seedGridStructureOnly/seedGridStylesFromXlsx)은 이미 위에서
      // 끝났고 그 트랜잭션에는 GRID_WINDOW_SEED_ORIGIN이 붙어 있으므로, 등록 시점과 origin 양쪽
      // 으로 "서버 자신의 시딩"을 사용자 편집으로 오해할 일이 없다(cells.observe와 동일한 방어).
      const styles = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_MAP);
      styles.observe((event) => {
        if (event.transaction.origin === GRID_WINDOW_SEED_ORIGIN) return;
        // 행/열 삽입이 서식 키를 다시 매긴 것도 되반영하지 않는다 — canonical의 서식은 셀 객체
        // 안의 인덱스(GridCell.s)라 행/열을 옮기면 이미 함께 따라온 상태다. 여기서 또 반영하면
        // 같은 이동이 두 번 적용된다(shiftGridStyleKeys / GRID_STRUCTURE_SHIFT_ORIGIN 주석 참고).
        if (event.transaction.origin === GRID_STRUCTURE_SHIFT_ORIGIN) return;
        void applyStylesChangeToCanonical(lazyStore, styles, event).catch((error: unknown) => {
          console.error(`서식 변경 반영 실패 (문서 ${room.docId}):`, error);
        });
      });
      const structure = ydoc.getMap<SheetStructure>(GRID_STRUCTURE_MAP);
      structure.observe((event) => {
        if (event.transaction.origin === GRID_WINDOW_SEED_ORIGIN) return;
        // 위와 같은 이유 — 행/열 삽입의 rowCount/colCount는 applyGridStructureOp가 canonical에
        // 직접 쓴 값을 그대로 클라이언트에 알리는 것뿐이므로 되반영할 것이 없다.
        if (event.transaction.origin === GRID_STRUCTURE_SHIFT_ORIGIN) return;
        applyStructureChangeToCanonical(room, structure, event);
      });
      // 예전에는 여기서 모든 시트를 한 번 통째로 재계산했다 — 윈도잉 이후에는 그 시점에 구독
      // 중인 페이지가 하나도 없어(=Yjs에 아무 셀도 없어) 계산해 둘 곳도, 그걸 볼 사람도 없다.
      // 원본 파일에 이미 있던 수식의 최초 계산값은 그 페이지가 처음 구독되는 순간
      // reconcilePageSubscriptions가 셀 데이터와 함께 채운다.
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
async function flushToStore(room: Room): Promise<void> {
  const { docId, kind, ydoc } = room;
  const meta = await store.readMeta(docId);
  if (!meta) return;

  if (kind === 'xlsx') {
    // 디스크의 현재 파일이 아니라 룸을 열 때의 원본을 쓴다 — 이유는 Room.originalXlsxBytes 주석 참고
    // (안 그러면 두 번째 저장부터 시트 삭제가 엉뚱한 시트에 적용된다).
    const original = room.originalXlsxBytes ?? (await store.readOriginal(docId, meta.ext));
    if (!room.lazyGrid || !room.baselineSheetExtents) {
      console.error(`동시편집 저장 실패 — canonical 그리드 없음 (문서 ${docId})`);
    } else {
      // 진행 중인 행/열 이동이 끝난 뒤에 읽는다 — 반쪽 상태를 저장하지 않기 위해(리뷰 지적 I-1,
      // drainCanonicalGrowth 주석 참고).
      await drainCanonicalGrowth(room);
      const newXlsxBytes = await exportXlsxWithFallback(room, original);
      await store.saveOriginal(docId, meta.ext, Buffer.from(newXlsxBytes));
      try {
        const grid = await reparseGridForViewer(newXlsxBytes);
        await saveGridDerived(store, docId, grid);
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
      room.suppressAutosaveReschedule = true;
      try {
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
      } finally {
        room.suppressAutosaveReschedule = false;
      }
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
  // 메시지 없이도 클라이언트가 "저장됨" 표시를 실시간으로 띄울 수 있는 이유. suppressAutosaveReschedule로
  // 감싸는 이유는 Room 인터페이스 정의부의 설명 참고 — 이 쓰기 자체가 다음 자동저장을 또
  // 예약해 버리면 실제 편집이 하나도 없어도 8초마다 영원히 재저장이 도는 루프가 생긴다.
  room.suppressAutosaveReschedule = true;
  try {
    ydoc.getMap<number>(DOC_META_MAP).set(DOC_META_SAVED_AT_KEY, Date.now());
  } finally {
    room.suppressAutosaveReschedule = false;
  }
}

/**
 * 유휴 자동저장 타이머 / Ctrl+S 즉시저장 / 마지막 참여자 퇴장, 이 세 지점이 같은 문서를 저장할 때
 * 서로 겹치지 않도록 docId별로 한 줄로 세운다 — saveSerializer.ts 참고.
 *
 * flushToStore가 실패하면(예: pptx가 슬라이드를 전부 지워 내보낼 기준 레이아웃이 없는 경우처럼
 * 구조적으로 되풀이 실패하는 상태) 그동안은 서버 콘솔에만 로그를 남기고 클라이언트에는 아무
 * 신호도 가지 않아, "저장 중…" 표시가 영원히 멈춰 있는 것처럼 보였다 — 실패 시각을 docMeta에
 * 기록해 다른 성공 신호(savedAt)와 같은 방식으로 모든 참여자 화면에 실시간 전파한다.
 */
function queueFlush(room: Room): Promise<void> {
  return runSerialized(room.docId, () =>
    flushToStore(room).catch((error: unknown) => {
      room.suppressAutosaveReschedule = true;
      try {
        room.ydoc.getMap<number>(DOC_META_MAP).set(DOC_META_SAVE_ERROR_AT_KEY, Date.now());
      } finally {
        room.suppressAutosaveReschedule = false;
      }
      throw error;
    }),
  );
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
  await queueFlush(room);
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
    if (!room.lazyGrid || !room.baselineSheetExtents) return null;
    // flushToStore와 같은 이유로 진행 중인 행/열 이동을 먼저 기다린다 — 이 경로는 사용자가 방금
    // 누른 다운로드라 반쪽 상태가 그대로 손에 들어간다(리뷰 지적 I-1).
    await drainCanonicalGrowth(room);
    // flushToStore와 같은 이유로 룸을 열 때의 원본을 기준으로 내보낸다(Room.originalXlsxBytes 주석).
    return exportXlsxWithFallback(room, room.originalXlsxBytes ?? original);
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
    queueFlush(room).catch((error) => {
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
    ready: Promise.resolve(), // 아래에서 즉시 덮어씀 — seedRoom(room)이 room 자신을 참조해야 해서
    suppressAutosaveReschedule: false,
    lazyGrid: null,
    baselineSheetExtents: null,
    structuralOps: [],
    originalXlsxBytes: null,
    removedSheetIndices: new Set<number>(),
    sheetDisplayOrder: new Map<number, number>(),
    pageSubscriptions: new Map(),
    canonicalGrowth: Promise.resolve(),
  };
  room.ready = seedRoom(room);
  rooms.set(docId, room);

  ydoc.on('update', (update: Uint8Array, origin: unknown) => {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeUpdate(encoder, update);
    // 전파는 origin과 무관하게 항상 한다 — 방금 채운 페이지가 클라이언트에 도착하는 경로가
    // 바로 이 broadcast이고, 비운 페이지가 클라이언트 메모리에서도 함께 줄어드는 경로이기도 하다.
    broadcast(room, encoding.toUint8Array(encoder));
    // canonical 그리드 동기화는 이제 cells.observe(...)(seedRoom, 시딩 완료 후 등록)가 맡는다 —
    // 여기서 다시 처리하면 삭제(cells.delete)를 놓치는 문제가 있었다(findings 참고).
    // 저장 결과를 반영하는 flushToStore 자신의 쓰기는 예외 — 위 suppressAutosaveReschedule 설명 참고.
    //
    // 창을 채우고 비우는 것(GRID_WINDOW_SEED_ORIGIN)도 자동저장을 예약하지 않는다: 그건 사용자
    // 편집이 아니라 "누가 어디를 보고 있는지"에 따른 캐시 조작일 뿐이고, 저장의 유일한 기준인
    // lazyGrid는 그로 인해 한 글자도 바뀌지 않는다. 예약을 걸면 아무도 편집하지 않은 문서를
    // 스크롤만 해도 8초마다 파일 전체를 다시 내보내고 뷰어 캐시를 재생성한다(편집 없는 무한
    // 재저장 루프와 같은 클래스의 낭비). 이 억제는 플래그가 아니라 트랜잭션 origin으로 판별하므로
    // 시딩 도중(청크 사이 양보 틈)에 도착한 진짜 편집의 자동저장 예약을 놓치는 일은 없다.
    //
    // 수식 계산값 쓰기(GRID_COMPUTED_WRITE_ORIGIN)도 같은 이유로 자동저장을 예약하지 않는다:
    // exportXlsxFromGrid는 lazyGrid/structuralOps만 읽고 computed 맵은 아예 보지 않으므로
    // 그 맵에 무엇이 쓰여도 저장될 바이트는 달라지지 않는다. 표식이 없으면 새 페이지를 스크롤로
    // 열기만 해도(구독 → 그 페이지 재계산 → computed 쓰기) 자동저장이 걸려, 아무도 편집하지 않은
    // 문서를 읽는 동안 전체 export + 재파싱이 반복된다. 진짜 편집은 GRID_CELLS_MAP 쓰기로
    // 도착하고 그건 이 표식이 없으므로 그대로 자동저장을 예약한다.
    if (isAutosaveExemptOrigin(origin)) return;
    if (!room.suppressAutosaveReschedule) scheduleAutosave(room);
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
    // **순서가 미묘하다**: 아래 delete가 끝나면 이 연결이 어느 시트를 보고 있었는지 알 방법이
    // 없고, "이전 합집합"도 이미 그 연결이 빠진 값이 되어 버린다 — 그러면 이 연결만 보던
    // 페이지가 차이(before - after)에 나타나지 않아 영원히 Yjs에 남는다. 그래서 구독 정보와
    // 그 시트의 합집합을 delete "전에" 확보하고, after는 delete "후에" 구한다.
    // 한 연결은 동시에 한 시트만 구독하므로 그 시트 하나만 reconcile하면 된다.
    const subscription = room.pageSubscriptions.get(conn);
    const before = subscription === undefined ? null : getSubscribedPageUnion(room, subscription.sheet);

    room.conns.delete(conn);
    room.readOnlyConns.delete(conn);
    room.pageSubscriptions.delete(conn);
    removeAwarenessStates(room.awareness, [...controlled], null);
    if (room.conns.size === 0) {
      if (room.autosaveTimer) {
        clearTimeout(room.autosaveTimer);
        room.autosaveTimer = null;
      }
      // 룸이 닫히면 남은 지연 제거 예약은 전부 취소한다 — 이제 Yjs를 비울 이유가 없고(문서
      // 전체가 메모리에서 사라진다), 살아 있는 타이머는 10초 뒤 이미 destroy된 Y.Doc을
      // 건드리려 하며 그동안 Room과 lazyGrid를 GC에서 붙잡아 둔다.
      cancelPendingUnseeds(room.docId);
      rooms.delete(room.docId);
      largeRoomDocIds.delete(room.docId);
      if (room.lazyGrid) void closeLazyGridStore(room.lazyGrid);
      void room.ready
        .then(() => queueFlush(room))
        .catch((error) => console.error(`동시편집 종료 저장 실패 (문서 ${room.docId}):`, error))
        .finally(() => room.ydoc.destroy());
    } else if (subscription !== undefined && before !== null) {
      // 룸이 계속 살아 있는 경우에만 reconcile한다 — 위 분기에서는 곧 Y.Doc이 destroy되므로
      // 채우거나 비울 대상 자체가 없다.
      reconcileInBackground(room, subscription.sheet, before, getSubscribedPageUnion(room, subscription.sheet));
    }
  }
  try {
    conn.close();
  } catch {
    /* 이미 닫힌 연결 */
  }
}

/**
 * 문서 삭제(`DELETE /api/documents/:id`) 직후 호출한다 — 그 문서를 메모리상 편집 중이던 Room이
 * 있으면 모든 연결을 강제로 끊어 즉시 정리한다.
 *
 * **왜 필요했는가**: 이 호출이 없으면 삭제 API가 디스크의 원본·파생 파일만 지우고, 이미 열려
 * 있던 협업 Room(lazyGrid 등 대용량 메모리를 그대로 붙잡고 있음)은 마지막 참여자가 스스로
 * 나갈 때까지 계속 살아 있었다 — 대용량 xlsx 룸이라면 삭제 후에도 한동안(또는 브라우저 탭을 안
 * 닫으면 무기한) 메모리를 계속 점유하고, `checkRoomCapacity`의 상한 슬롯도 계속 차지한다.
 *
 * **왜 순서가 중요한가(호출부의 계약)**: 반드시 `store.deleteDocument`가 끝난 **뒤**에 불러야
 * 한다. 남은 연결을 여기서 강제로 끊으면 closeConn의 "마지막 참여자 퇴장" 경로가 그대로 타서
 * `queueFlush(room)`(→ `flushToStore`)가 예약되는데, `flushToStore`는 `store.readMeta`가 이미
 * null(문서가 지워짐)이면 즉시 아무 것도 안 하고 반환한다 — 삭제가 먼저 끝나 있어야 이 안전한
 * no-op 경로를 타고, 순서가 바뀌면 삭제 직후에 방금 지운 파일이 자동저장으로 되살아날 수 있다.
 *
 * 4410(HTTP 410 Gone과 같은 취지)로 닫아, 클라이언트가 재연결을 한 번 더 시도하는 왕복 없이
 * 곧바로 "문서가 삭제되었습니다"를 보여줄 수 있게 한다(grid-editor.ts의 connection-close 처리는
 * 4400번대 코드를 전부 영구 거부로 취급해 재연결을 멈추고 이유를 그대로 보여준다).
 */
export function forceCloseRoom(docId: string, reason = '문서가 삭제되었습니다.'): void {
  const room = rooms.get(docId);
  if (!room) return;
  for (const conn of [...room.conns.keys()]) {
    try {
      conn.close(4410, reason);
    } catch {
      /* 이미 닫힌 연결 */
    }
    closeConn(room, conn);
  }
}

/**
 * 한 시트가 가질 수 있는 페이지 번호의 상한 — canonical 그리드가 자랄 수 있는 최대 행 수를 페이지
 * 크기로 나눈 값이다(300_000 / 500 = 600). 두 상수 중 어느 쪽이 나중에 바뀌어도 자동으로 따라오도록
 * 상수를 직접 쓰지 않고 계산한다. 이 범위 밖 페이지 번호는 채울 데이터가 존재할 수 없으므로 정상
 * 클라이언트가 보낼 이유가 없다.
 */
const MAX_SUBSCRIBABLE_PAGES = Math.ceil(MAX_CANONICAL_ROWS / PAGE_SIZE_ROWS);

/**
 * 한 연결이 **동시에** 구독할 수 있는 페이지 수의 상한. 구독은 "지금 화면에 보이는 창 + 앞뒤 여유"를
 * 뜻하므로 실제로는 한 자릿수다 — 32는 그보다 훨씬 넉넉하면서도, 조작된 소켓 메시지 하나로 수천
 * 페이지 시딩(= 수백만 셀을 Yjs에 밀어 넣는 작업)을 걸어 서버를 묶어 버리는 것은 막는다.
 */
const MAX_SUBSCRIBED_PAGES_PER_CONN = 32;

/**
 * MESSAGE_PAGE_SUBSCRIBE 페이로드 검증 — 리뷰 지적 I2.
 *
 * 이 메시지는 **신뢰할 수 없는 입력**이다: 여기서 통과한 숫자가 그대로 페이지 시딩 루프의 행 범위와
 * getSubscribedPageUnion의 Set 내용이 된다. 예전 검증은 `typeof === 'number'`뿐이라 0.5·-1·1e9·
 * NaN 같은 값과 길이 5000짜리 배열이 전부 통과했다(비정수 페이지는 행 범위 계산을 깨고, 거대한
 * 배열은 그 자체로 값싼 DoS다).
 *
 * 하나라도 어긋나면 **메시지 전체를 거부**한다 — 잘라내거나 보정하면 클라이언트가 요청한 창과 서버가
 * 채운 창이 조용히 어긋나, 화면의 빈 칸을 나중에 추적하기가 훨씬 어려워진다.
 */
function isValidPageSubscription(parsed: unknown): parsed is { sheet: number; pages: readonly number[] } {
  if (!parsed || typeof parsed !== 'object') return false;
  const { sheet, pages } = parsed as { sheet?: unknown; pages?: unknown };
  // Number.isInteger는 NaN·Infinity·비정수를 한 번에 걸러 준다(typeof === 'number'가 놓치던 것들).
  if (!Number.isInteger(sheet) || (sheet as number) < 0 || (sheet as number) >= MAX_CANONICAL_SHEETS) return false;
  if (!Array.isArray(pages) || pages.length > MAX_SUBSCRIBED_PAGES_PER_CONN) return false;
  return pages.every((page) => Number.isInteger(page) && page >= 0 && page < MAX_SUBSCRIBABLE_PAGES);
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
      case MESSAGE_PAGE_SUBSCRIBE: {
        try {
          const payload = decoding.readVarString(decoder);
          const parsed: unknown = JSON.parse(payload);
          if (!isValidPageSubscription(parsed)) {
            // 잘못된 메시지는 **통째로** 버린다(잘라내거나 보정하지 않는다) — 일부만 받아들이면
            // 클라이언트는 자기가 요청한 창이 채워졌다고 믿는데 서버는 다른 창을 채우게 되어,
            // 화면에 빈 칸이 보이는 원인을 추적하기 어려워진다. 형식이 깨진 JSON을 그냥 로그만
            // 남기고 무시하는 아래 catch와 같은 처리다.
            console.error('페이지 구독 메시지 거부(형식/범위 오류):', payload.slice(0, 200));
          } else {
            // 합집합은 이 연결의 구독을 덮어쓰기 "전"과 "후"로 각각 구해야 한다 — 그 차이가
            // 곧 "채울 페이지"와 "비울 페이지"다(reconcilePageSubscriptions).
            //
            // 시트를 갈아탄 경우(시트 탭 전환)에는 새 시트만 reconcile하면 이전 시트의 페이지가
            // 아무도 보지 않는데도 Yjs에 영원히 남는다 — 이전 시트도 함께 reconcile해서 유예
            // 시간 뒤에 비워지도록 한다. 한 연결은 동시에 한 시트만 구독하므로 이전 시트는 최대
            // 하나다(스펙의 "시트 전환" 절).
            const previous = room.pageSubscriptions.get(conn);
            const previousSheet =
              previous !== undefined && previous.sheet !== parsed.sheet ? previous.sheet : null;
            const beforeNew = getSubscribedPageUnion(room, parsed.sheet);
            const beforeOld = previousSheet === null ? null : getSubscribedPageUnion(room, previousSheet);

            room.pageSubscriptions.set(conn, { sheet: parsed.sheet, pages: parsed.pages });

            if (previousSheet !== null && beforeOld !== null) {
              reconcileInBackground(room, previousSheet, beforeOld, getSubscribedPageUnion(room, previousSheet));
            }
            reconcileInBackground(room, parsed.sheet, beforeNew, getSubscribedPageUnion(room, parsed.sheet));
          }
        } catch (error) {
          console.error('페이지 구독 메시지 처리 오류:', error);
        }
        break;
      }
      case MESSAGE_GRID_STRUCTURE_OP: {
        try {
          const payload = decoding.readVarString(decoder);
          if (room.readOnlyConns.has(conn)) {
            // 행/열 삽입·삭제는 **쓰기**다 — 읽기 전용으로 들어온 연결이 보낸 것은 버린다.
            // (페이지 구독 메시지는 "지금 여기를 보고 있다"는 읽기 요청이라 읽기 전용도 허용하는
            //  것과 의도적으로 다르다. 클라이언트가 편집 UI를 숨기는 것과 별개로 서버가 직접 막는다.)
            // 페이로드는 로그에 남기지 않는다 — 조작된 연결이 로그를 불릴 수 있는 창구를 줄인다.
            console.error('행/열 구조 변경 메시지 거부(읽기 전용 연결)');
          } else {
            const parsed: unknown = JSON.parse(payload);
            if (!isValidGridStructureOp(parsed)) {
              // 구독 메시지와 같은 원칙 — 형식이 어긋나면 통째로 버린다(보정하지 않는다).
              console.error('행/열 구조 변경 메시지 거부(형식/범위 오류):', payload.slice(0, 200));
            } else {
              // 실제 적용은 canonical 갱신 체인에서 순서대로 돈다(enqueueCanonicalGrowth 주석) —
              // 여기서 await하면 소켓 메시지 처리 자체가 막힌다.
              enqueueCanonicalGrowth(room, () => applyGridStructureOp(room, parsed));
            }
          }
        } catch (error) {
          console.error('행/열 구조 변경 메시지 처리 오류:', error);
        }
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

/**
 * 구독에서 빠진 페이지를 곧바로 비우지 않고 기다리는 유예 시간. 스크롤은 본질적으로 왕복이 잦아
 * (한 칸 내렸다 바로 올리거나, 두 페이지 경계를 오가며 흔들리는 경우), 빠지는 즉시 비우면 같은
 * 500행을 계속 지웠다 다시 채우며 그 전량을 모든 클라이언트에 다시 전파하게 된다. 10초 정도
 * 남겨두면 그 왕복은 전부 "이미 채워져 있음"으로 흡수된다.
 */
const PAGE_UNSEED_GRACE_MS = 10_000;

/**
 * 지연 제거가 예약된 페이지의 타이머 — 문서별로 묶어 둔다(바깥 키 docId, 안쪽 키 "시트:페이지").
 * 문서별로 묶는 이유는 룸이 닫힐 때 그 문서의 예약을 한 번에 취소해야 하기 때문이다
 * (cancelPendingUnseeds).
 */
const pendingPageUnseeds = new Map<string, Map<string, NodeJS.Timeout>>();

function pageUnseedKey(sheet: number, page: number): string {
  return `${sheet}:${page}`;
}

/** 예약된 지연 제거 하나를 취소한다 — 유예 시간 안에 그 페이지가 다시 구독될 때 쓴다. */
function cancelPendingUnseed(docId: string, sheet: number, page: number): void {
  const perDoc = pendingPageUnseeds.get(docId);
  if (!perDoc) return;
  const key = pageUnseedKey(sheet, page);
  const timer = perDoc.get(key);
  if (!timer) return;
  clearTimeout(timer);
  perDoc.delete(key);
  if (perDoc.size === 0) pendingPageUnseeds.delete(docId);
}

/**
 * 그 문서에 예약된 지연 제거를 전부 취소한다 — 마지막 참여자가 나가 룸을 닫을 때 반드시 불러야
 * 한다. 남겨두면 (1) 10초 뒤 이미 destroy된 Y.Doc에 쓰기를 시도하고, (2) 그동안 살아 있는
 * 타이머가 Room과 lazyGrid 전체를 GC 대상에서 붙잡아 둔다.
 */
function cancelPendingUnseeds(docId: string): void {
  const perDoc = pendingPageUnseeds.get(docId);
  if (!perDoc) return;
  for (const timer of perDoc.values()) clearTimeout(timer);
  pendingPageUnseeds.delete(docId);
}

/**
 * 구독 합집합이 before → after로 바뀐 것을 실제 Yjs 내용에 반영한다 — 이 함수가 이 태스크의
 * 핵심이다(윈도우를 실제로 움직이는 유일한 지점).
 *
 * - 새로 들어온 페이지: 즉시 셀 데이터를 채우고, 그 페이지의 수식 계산값도 같이 채운다(둘이
 *   함께 도착해야 클라이언트가 수식 칸을 빈 칸으로 잠깐 보는 일이 없다).
 * - 빠진 페이지: PAGE_UNSEED_GRACE_MS 뒤에 "그때도 여전히 아무도 안 보고 있으면" 비운다.
 */
async function reconcilePageSubscriptions(
  room: Room,
  sheet: number,
  before: Set<number>,
  after: Set<number>,
): Promise<void> {
  if (!room.lazyGrid) return;
  const lazyStore = room.lazyGrid;

  // 1) 취소는 먼저, 동기적으로 전부 처리한다. 아래 채우기 루프는 페이지마다 await로 양보하므로,
  //    취소를 그 루프 안에서 하면 "앞 페이지를 채우는 동안 뒤 페이지의 제거 타이머가 발동"하는
  //    틈이 생긴다(유예 10초에 비하면 극히 좁은 틈이지만, 없앨 수 있는 틈을 남길 이유가 없다).
  for (const page of after) cancelPendingUnseed(room.docId, sheet, page);

  // 2) 새로 구독된 페이지를 채운다.
  for (const page of after) {
    if (before.has(page)) continue;
    await seedGridPageFromXlsx(room.ydoc, lazyStore, sheet, page);
    await recomputeFormulasForPage(room, sheet, page);
  }

  // 3) 빠진 페이지의 지연 제거를 예약한다.
  for (const page of before) {
    if (after.has(page)) continue;
    const perDoc = pendingPageUnseeds.get(room.docId) ?? new Map<string, NodeJS.Timeout>();
    pendingPageUnseeds.set(room.docId, perDoc);
    const key = pageUnseedKey(sheet, page);
    clearTimeout(perDoc.get(key)); // 같은 페이지에 이미 예약이 있으면 뒤 것으로 갱신
    perDoc.set(
      key,
      setTimeout(() => {
        perDoc.delete(key);
        if (perDoc.size === 0) pendingPageUnseeds.delete(room.docId);
        // 룸이 이미 닫혔으면(마지막 참여자 퇴장 → rooms에서 제거 + ydoc.destroy()) 비울 것도,
        // 비울 대상 Y.Doc도 없다 — 파괴된 Y.Doc을 건드리지 않도록 여기서 끝낸다. 정상 경로에서는
        // closeConn이 cancelPendingUnseeds로 미리 취소하므로 여기까지 오지 않는다.
        if (rooms.get(room.docId) !== room || !room.lazyGrid) return;
        // 유예 시간 동안 누군가 이 페이지를 다시 구독했을 수 있다 — 예약 취소를 놓친 경로가
        // 있어도 여기서 한 번 더 확인해 "보고 있는 페이지를 비우는" 사고를 막는다.
        if (getSubscribedPageUnion(room, sheet).has(page)) return;
        const rowCount = room.lazyGrid.sheetMeta[sheet]?.rowCount ?? 0;
        void unseedGridPageFromXlsx(room.ydoc, sheet, page, rowCount).catch((error: unknown) => {
          console.error(`페이지 비우기 실패 (문서 ${room.docId}, 시트 ${sheet}, 페이지 ${page}):`, error);
        });
      }, PAGE_UNSEED_GRACE_MS),
    );
  }
}

/**
 * reconcilePageSubscriptions를 기다리지 않고 흘려보낸다 — 호출부(소켓 메시지 처리, 연결 종료)는
 * 둘 다 동기 함수라 await할 수 없고, 페이지 채우기는 청크마다 이벤트 루프에 양보하므로 애초에
 * 기다려서도 안 된다(그 사이 다른 클라이언트의 메시지가 막힌다).
 *
 * `.catch`가 반드시 필요하다: 버려진 Promise의 reject는 Node에서 unhandledRejection이 되어
 * 프로세스 전체를 죽이고, 그러면 이 문서만이 아니라 그 순간 열려 있던 모든 편집 세션이 함께
 * 끊긴다(scheduleAutosave의 같은 주석 참고).
 */
function reconcileInBackground(room: Room, sheet: number, before: Set<number>, after: Set<number>): void {
  void reconcilePageSubscriptions(room, sheet, before, after).catch((error: unknown) => {
    console.error(`페이지 구독 반영 실패 (문서 ${room.docId}, 시트 ${sheet}):`, error);
  });
}

/**
 * 특정 시트에 대해, 지금 그 룸에 연결된 모든 클라이언트가 구독 중인 페이지 번호의 합집합을
 * 계산한다 — 이 합집합이 곧 "Yjs가 담고 있어야 하는 창"이고, 이 값이 바뀔 때마다
 * reconcilePageSubscriptions가 실제 Yjs 내용을 그에 맞춘다.
 */
function getSubscribedPageUnion(room: Room, sheet: number): Set<number> {
  const union = new Set<number>();
  for (const subscription of room.pageSubscriptions.values()) {
    if (subscription.sheet !== sheet) continue;
    for (const page of subscription.pages) union.add(page);
  }
  return union;
}

/** 테스트 전용 */
export function getSubscribedPageUnionForTest(docId: string, sheet: number): Set<number> {
  const room = rooms.get(docId);
  if (!room) return new Set();
  return getSubscribedPageUnion(room, sheet);
}

/**
 * 테스트 전용 — "지금 실제로 Yjs에 채워져 있는 페이지 집합"을 들여다본다.
 *
 * getSubscribedPageUnionForTest(누가 무엇을 보고 있다고 말했는지)와 다르다: 이쪽은 그 의도가
 * 실제 Yjs 내용에 반영됐는지, 특히 유예 시간 뒤 정말로 비워졌는지를 본다. 유예 제거는
 * lazyGrid에는 아무 영향이 없어(편집된 칸은 dirtyOverlay에 남고, 나머지는 언제든 디스크에서 다시
 * 읽을 수 있다) 클라이언트 mirror만으로는 "지워짐"을 관찰하기 어렵기 때문에 서버 내부 상태를
 * 직접 확인할 창구가 필요하다.
 *
 * GRID_CELLS_MAP의 키에서 행 번호를 뽑아 페이지 번호로 환산한다 — 값이 있는 셀이 하나도 없는
 * 페이지는(예: 전부 빈 칸인 500행) 채워져 있어도 여기 나타나지 않으므로, 테스트 픽스처는 페이지마다
 * 최소 한 칸은 값이 있도록 만들어야 한다.
 */
export function getSeededPagesForTest(docId: string, sheet: number): Set<number> {
  const room = rooms.get(docId);
  if (!room) return new Set();
  return seededPagesInYdoc(room, sheet);
}

/** 테스트 전용 — 룸의 지연(lazy) 그리드 저장소를 직접 들여다본다. 프로덕션 코드에서는 쓰지 않는다. */
export function getRoomLazyGridForTest(docId: string): LazyGridStore | null {
  return rooms.get(docId)?.lazyGrid ?? null;
}

/**
 * 테스트 전용 — 지금 예약돼 있는 canonical 성장 작업이 전부 끝날 때까지 기다린다.
 * 성장은 이벤트 루프에 양보하며 청크 단위로 도는 백그라운드 체인이라(canonicalGrowth 주석 참고)
 * 고정된 틱 수를 흘려보내는 것으로는 완료를 보장할 수 없다.
 */
export async function awaitCanonicalGrowthForTest(docId: string): Promise<void> {
  const room = rooms.get(docId);
  if (!room) return;
  // 내보내기 경로가 쓰는 것과 **같은 배수 로직**을 쓴다(drainCanonicalGrowth) — 두 벌로 두면
  // 한쪽만 고쳐지는 일이 생긴다.
  await drainCanonicalGrowth(room);
}
