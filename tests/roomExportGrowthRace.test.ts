import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeSyncStep1 } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

/**
 * 이 파일이 붙잡는 회귀 — 리뷰 지적 I-1: **내보내기가 반쪽 상태(torn)의 canonical 그리드를 읽는다.**
 *
 * 행/열 삽입·삭제는 canonical 갱신 체인(Room.canonicalGrowth)에서 돌고, 열 방향 연산은
 * YIELD_EVERY_N_ROWS(2000)행마다 실제로 이벤트 루프에 양보한다(gridShift.ts의 insertCanonicalCols).
 * 즉 그 연산이 도는 동안 canonical 그리드는 **앞쪽 행은 이미 밀렸고 뒤쪽 행은 아직 안 밀린** 상태로
 * 관찰될 수 있다.
 *
 * 그런데 flushToStore(자동저장/Ctrl+S)와 getCurrentBytes(즉시 다운로드)는 그 체인을 기다리지 않고
 * exportXlsxFromGrid에 canonicalGrid를 바로 넘겼다 — 다운로드나 저장이 청크 진행 중에 끼어들면
 * **정확히 이 계획 전체가 없애려던 그 어긋남**(행별로 열이 다르게 밀린 파일)이 사용자 파일로
 * 나간다. 조용한 데이터 손상이고, 화면에는 정상으로 보인다.
 *
 * 검증은 반드시 **실제로 내보낸 바이트를 다시 열어보는** 것으로 한다(roomRowColShift.test.ts와 같은
 * 원칙). 핵심 단정은 두 가지다:
 *   (1) 비-torn — 시트의 앞·중간·끝 행이 **모두 같은 만큼** 밀려 있다.
 *   (2) 체인을 실제로 비웠다 — 밀린 양이 예약된 연산 전부(6번)를 반영한다.
 */

const MESSAGE_SYNC = 0;
const MESSAGE_PAGE_SUBSCRIBE = 2;
const MESSAGE_GRID_STRUCTURE_OP = 4;

/** 열 삽입 루프의 청크 경계(spreadsheet.ts의 YIELD_EVERY_N_ROWS)와 픽스처 표식 위치의 근거. */
const YIELD_EVERY_N_ROWS = 2000;

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  readonly mirror = new Y.Doc();

  send(data: Uint8Array, cb?: (err?: Error) => void): void {
    const decoder = decoding.createDecoder(data);
    if (decoding.readVarUint(decoder) === MESSAGE_SYNC) {
      const syncType = decoding.readVarUint(decoder);
      if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) readUpdate(decoder, this.mirror, this);
    }
    cb?.();
  }
  close(): void {}
  ping(): void {}
}

function sendSubscribe(socket: FakeSocket, payload: unknown): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_PAGE_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify(payload));
  socket.emit('message', encoding.toUint8Array(encoder));
}

function sendStructureOp(socket: FakeSocket, payload: unknown): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_GRID_STRUCTURE_OP);
  encoding.writeVarString(encoder, JSON.stringify(payload));
  socket.emit('message', encoding.toUint8Array(encoder));
}

async function flushEventLoop(turns = 80): Promise<void> {
  for (let i = 0; i < turns; i += 1) await new Promise((resolve) => setImmediate(resolve));
}

function syncMirrorWithRoom(socket: FakeSocket): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeSyncStep1(encoder, socket.mirror);
  socket.emit('message', encoding.toUint8Array(encoder));
}

/**
 * 5000행 픽스처 — 열 삽입 루프가 **여러 청크에 걸쳐** 돌아야 반쪽 상태가 관찰 가능하다
 * (5000행 = 2000 + 2000 + 1000, 청크마다 양보).
 *
 *   B열: 1~5000행 전부 'b'  ← parseXlsxToGrid의 rowCount는 "값이 있는 행의 개수"라 성긴 시트를
 *        쓰면 시트가 잘려 경계 행 자체가 사라진다(roomRowColShift.test.ts의 같은 이유).
 *   A열 표식(괄호는 0-based):
 *     A1    (0)    'top'      ← 첫 청크
 *     A2501 (2500) 'mid2500'  ← 두 번째 청크
 *     A4501 (4500) 'tail4500' ← 마지막 청크 — 반쪽 상태에서 **아직 밀리지 않은** 쪽
 */
async function makeTallXlsx(rows = 5000): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  for (let excelRow = 1; excelRow <= rows; excelRow += 1) sheet.getCell(excelRow, 2).value = 'b';
  sheet.getCell('A1').value = 'top';
  sheet.getCell(`A${YIELD_EVERY_N_ROWS + 501}`).value = 'mid2500';
  sheet.getCell(`A${YIELD_EVERY_N_ROWS * 2 + 501}`).value = 'tail4500';
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

async function seedDoc(docId: string, original: Buffer): Promise<void> {
  const { store } = await import('../src/storage.js');
  await store.saveOriginal(docId, 'xlsx', original);
  await store.writeMeta({
    id: docId,
    name: 'x.xlsx',
    ext: 'xlsx',
    status: 'ready',
    pageCount: 1,
    createdAt: new Date().toISOString(),
    progress: 100,
    viewKind: 'grid',
  });
}

/** 그 행에서 표식이 놓인 열 번호(1-based) — 없으면 -1. "얼마나 밀렸는지"를 직접 읽기 위한 헬퍼. */
function columnOf(sheet: ExcelJS.Worksheet, row: number, marker: string): number {
  const width = Math.max(sheet.columnCount, 16);
  for (let c = 1; c <= width; c += 1) {
    if (sheet.getCell(row, c).value === marker) return c;
  }
  return -1;
}

async function openWorkbook(bytes: Uint8Array): Promise<ExcelJS.Worksheet> {
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(bytes as unknown as ArrayBuffer);
  return reopened.getWorksheet('Sheet1')!;
}

/**
 * 내보낸 시트가 **비-torn**이고 예약된 연산 전부를 반영하는지.
 * expectedShift = 삽입한 열 개수(0-based 0열에 삽입하므로 A열 표식은 그만큼 오른쪽으로 간다).
 */
function expectFullyShifted(sheet: ExcelJS.Worksheet, insertedCols: number): void {
  const top = columnOf(sheet, 1, 'top');
  const mid = columnOf(sheet, YIELD_EVERY_N_ROWS + 501, 'mid2500');
  const tail = columnOf(sheet, YIELD_EVERY_N_ROWS * 2 + 501, 'tail4500');
  // (1) 비-torn: 세 행이 모두 같은 만큼 밀렸다(청크 경계로 갈라지지 않았다).
  expect({ mid, tail }).toEqual({ mid: top, tail: top });
  // (2) 체인을 실제로 비운 뒤 내보냈다.
  expect(top).toBe(1 + insertedCols);
  // 옛 자리는 비어 있다(반쪽만 밀린 채 옛 값이 남지 않았다).
  for (let c = 1; c <= insertedCols; c += 1) {
    expect(sheet.getCell(1, c).value).toBeFalsy();
    expect(sheet.getCell(YIELD_EVERY_N_ROWS * 2 + 501, c).value).toBeFalsy();
  }
}

describe('room.ts: export must not read a half-shifted canonical grid (review finding I-1)', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-export-race-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
    // setImmediate는 반드시 real로 남긴다 — ExcelJS의 zip/stream 내부가 진짜 setImmediate 턴에
    // 의존해서, 전부 fake로 두면 내보내기가 영원히 끝나지 않는다(계획의 Global Constraints).
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.useRealTimers();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  async function openRoom(original: Buffer): Promise<{ docId: string; socket: FakeSocket }> {
    const { joinEditSession } = await import('../src/collab/room.js');
    const docId = randomUUID();
    await seedDoc(docId, original);
    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false);
    sendSubscribe(socket, { sheet: 0, pages: [0] });
    await flushEventLoop();
    syncMirrorWithRoom(socket);
    return { docId, socket };
  }

  /**
   * 체인이 **분명히 아직 돌고 있는** 상태를 만든다: 열 삽입 6번을 연달아 예약한다. 하나하나가
   * 5000행짜리 청크 루프 + 서식 키 재매핑 + 페이지 재시딩 + 수식 재계산을 거치므로, 내보내기가
   * 시작되는 시점(ExcelJS의 zip 로드가 끝나기도 전)에 체인은 절대 비어 있지 않다.
   */
  function enqueueSixColumnInserts(socket: FakeSocket): number {
    for (let i = 0; i < 6; i += 1) sendStructureOp(socket, { sheet: 0, kind: 'insertCols', at: 0, count: 1 });
    return 6;
  }

  test('getCurrentBytes (immediate download) waits for the in-flight canonical shift chain', async () => {
    const { getCurrentBytes } = await import('../src/collab/room.js');
    const { docId, socket } = await openRoom(await makeTallXlsx());

    const inserted = enqueueSixColumnInserts(socket);
    // 일부러 settle하지 않는다 — 청크 루프가 도는 중에 다운로드가 끼어드는 실제 상황이다.
    const bytes = await getCurrentBytes(docId);
    expect(bytes).not.toBeNull();

    expectFullyShifted(await openWorkbook(bytes!), inserted);
  }, 180_000);

  test('flushNow → flushToStore (Ctrl+S save) waits for the in-flight canonical shift chain', async () => {
    const { flushNow } = await import('../src/collab/room.js');
    const { store } = await import('../src/storage.js');
    const { docId, socket } = await openRoom(await makeTallXlsx());

    const inserted = enqueueSixColumnInserts(socket);
    // 같은 이유로 settle하지 않는다 — 유휴 자동저장/Ctrl+S가 청크 루프 중간에 도는 상황이다.
    await expect(flushNow(docId)).resolves.toBe(true);

    expectFullyShifted(await openWorkbook(await store.readOriginal(docId, 'xlsx')), inserted);
  }, 180_000);
});
