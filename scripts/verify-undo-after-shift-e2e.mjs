// 실행 중인 서버를 상대로, **실제 브라우저에서** "셀 편집 → 행 삽입 → Ctrl+Z"를 눌러
// undo가 **엉뚱한 행을 덮어쓰지 않는지** 확인한다(리뷰 지적 I-4).
//
// 겨냥하는 결함: Y.UndoManager의 스택에 쌓인 로컬 셀 편집은 편집 당시의 `sheet:row:col` 키로
// 되돌린다. 그 뒤에 행/열 삽입·삭제가 일어나면 그 키는 **다른 행**을 가리키므로, Ctrl+Z가
// 사용자가 손댄 적도 없는 칸에 옛 값을 조용히 써버린다.
//
// 세 가지 모드:
//
//   control      구조 변경 **없이** 편집 → Ctrl+Z. 대조군 — Ctrl+Z 경로 자체가 이 브라우저에서
//                실제로 동작하는지 확인한다(편집이 정말로 되돌려져야 한다).
//
//   shift        편집 → (내가) 행 삽입 → Ctrl+Z. 편집한 칸이 **구독 중인 페이지**에 있는 경우다.
//                여기서는 Yjs 자체가 한 겹 막아 준다: 서버가 구조 변경 뒤 구독 페이지를 다시
//                채우면서 그 키에 **더 새로운 항목**을 쓰므로, redoItem이 "다른 클라이언트의
//                변경과 충돌한다"며 복원을 거부한다(yjs/dist/yjs.cjs의 redoItem, 기본
//                ignoreRemoteMapChanges=false). 그래서 이 모드는 **고치기 전에도 통과한다** —
//                회귀 방지용으로만 의미가 있다.
//
//   unsubscribed 편집 → 편집한 칸의 페이지가 **더 이상 구독되지 않도록 멀리 스크롤** →
//                그 페이지를 구독하지 않는 **원격 협업자**(원시 WebSocket 클라이언트)가 행 삽입 →
//                Ctrl+Z. 이때는 그 키에 더 새로운 항목이 없어(구독 밖 페이지는 다시 채워지지
//                않는다) Yjs의 위 보호가 걸리지 않는다 — **낡은 키로 옛 값이 그대로 복원되어**
//                밀려 들어온 다른 행의 값을 조용히 덮어쓰고, 그게 canonical을 거쳐 파일로 나간다.
//                이게 I-4의 실제 재현 경로다.
//
// 검증은 **실제 다운로드 API가 내려주는 바이트**를 다시 열어 본다(verify-rowcol-shift-e2e.mjs와 같은 원칙).
//
// 사용법(서버를 먼저 같은 DATA_DIR/SECRET/PORT로 띄워 둘 것):
//   DOCVIEW_DATA_DIR=... DOCVIEW_SECRET=... DOCVIEW_PORT=... \
//     node scripts/verify-undo-after-shift-e2e.mjs [control|shift|unsubscribed]
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import jwt from 'jsonwebtoken';
import { chromium } from 'playwright';
import WebSocket from 'ws';
import * as encoding from 'lib0/encoding';
import * as decoding from 'lib0/decoding';
import * as syncProtocol from 'y-protocols/sync';
import * as Y from 'yjs';

const DATA_DIR = process.env.DOCVIEW_DATA_DIR;
const SECRET = process.env.DOCVIEW_SECRET;
const PORT = process.env.DOCVIEW_PORT ?? '8099';

const MODE = process.argv[2] ?? 'shift';
if (!['control', 'shift', 'unsubscribed'].includes(MODE)) {
  console.error(`unknown mode: ${MODE} (control | shift | unsubscribed)`);
  process.exit(1);
}
console.log(`=== live e2e: undo after structural shift (mode: ${MODE}) ===`);

if (!DATA_DIR || !SECRET) {
  console.error('DOCVIEW_DATA_DIR / DOCVIEW_SECRET 환경변수가 필요합니다.');
  process.exit(1);
}

const ROW_HEIGHT = 27;
const EDIT_ROW = 2; // 0-based — 편집할 행(페이지 0)
// unsubscribed 모드는 페이지 경계를 넘어야 하므로 여러 페이지짜리 시트가 필요하다(PAGE_SIZE_ROWS=500).
const ROWS = MODE === 'unsubscribed' ? 1600 : 10;
const FAR_ROW = 1300; // 0-based — 페이지 0이 구독에서 빠지도록 멀리 스크롤할 지점

/** A열 0~9행에 'keep0'..'keep9', B열은 전 행 'b'(parseXlsxToGrid의 rowCount가 잘리지 않게). */
async function makeFixture() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  for (let r = 1; r <= ROWS; r += 1) sheet.getCell(r, 2).value = 'b';
  for (let r = 1; r <= Math.min(10, ROWS); r += 1) sheet.getCell(r, 1).value = `keep${r - 1}`;
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

const id = randomUUID();
await fs.mkdir(path.join(DATA_DIR, 'originals'), { recursive: true });
await fs.mkdir(path.join(DATA_DIR, 'meta'), { recursive: true });
await fs.writeFile(path.join(DATA_DIR, 'originals', `${id}.xlsx`), await makeFixture());
await fs.writeFile(
  path.join(DATA_DIR, 'meta', `${id}.json`),
  JSON.stringify({
    id,
    name: 'undo-after-shift-e2e.xlsx',
    ext: 'xlsx',
    status: 'ready',
    pageCount: 1,
    createdAt: new Date().toISOString(),
    viewKind: 'grid',
    progress: 100,
  }),
);

const token = jwt.sign(
  { docId: id, viewer: 'e2e', perms: { canPrint: true, canDownload: true, canCopy: true, canEdit: true } },
  SECRET,
  { expiresIn: 1800 },
);
const base = `http://127.0.0.1:${PORT}`;
const editorUrl = `${base}/editor-grid.html?token=${encodeURIComponent(token)}`;

const failures = [];

/**
 * 페이지 구독을 하지 않는 **원격 협업자** — 원시 WebSocket으로 붙어 구조 연산 메시지만 보낸다.
 * 구독하지 않는 것이 핵심이다: 구독하면 서버가 구조 변경 뒤 그 페이지를 다시 채워(=그 셀 키에 더
 * 새로운 항목을 써서) Yjs의 redoItem 보호가 되살아나고, 재현하려는 상황이 사라진다.
 */
async function connectRemoteCollaborator() {
  const mirror = new Y.Doc();
  const ws = new WebSocket(`ws://127.0.0.1:${PORT}/ws/edit/${id}?token=${encodeURIComponent(token)}`);
  ws.binaryType = 'arraybuffer';
  await new Promise((resolve, reject) => {
    ws.on('open', resolve);
    ws.on('error', reject);
  });
  ws.on('message', (data) => {
    const bytes = data instanceof ArrayBuffer ? new Uint8Array(data) : new Uint8Array(data);
    const decoder = decoding.createDecoder(bytes);
    if (decoding.readVarUint(decoder) !== 0) return; // messageSync만 본다
    const syncType = decoding.readVarUint(decoder);
    if (syncType === syncProtocol.messageYjsSyncStep2 || syncType === syncProtocol.messageYjsUpdate) {
      syncProtocol.readUpdate(decoder, mirror, ws);
    }
  });
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, 0); // messageSync
  syncProtocol.writeSyncStep1(encoder, mirror);
  ws.send(encoding.toUint8Array(encoder));
  await new Promise((resolve) => setTimeout(resolve, 800));
  return {
    /** 클라이언트의 sendGridStructureOp와 같은 인코딩(varUint 4 + varString JSON). */
    sendStructureOp(payload) {
      const e = encoding.createEncoder();
      encoding.writeVarUint(e, 4); // MESSAGE_GRID_STRUCTURE_OP
      encoding.writeVarString(e, JSON.stringify(payload));
      ws.send(encoding.toUint8Array(e));
    },
    /** 룸이 지금 Yjs에 들고 있는 그 셀 키의 값(없으면 undefined) — 유령 쓰기를 직접 관찰한다. */
    cellValue(key) {
      return mirror.getMap('cells').get(key); // gridShared.ts의 GRID_CELLS_MAP
    },
    close() {
      ws.close();
    },
  };
}

const browser = await chromium.launch();
const page = await browser.newPage();
page.on('pageerror', (err) => {
  console.log('PAGEERROR:', err.message);
  failures.push(`pageerror: ${err.message}`);
});

await page.goto(editorUrl, { waitUntil: 'domcontentloaded' });
await page.waitForSelector('#grid-edit-table td[data-row]', { timeout: 30000 });
await page.waitForTimeout(800); // 페이지 구독 디바운스(200ms) + 최초 시딩

// ---------------------------------------------------------------- 1) 로컬 셀 편집
await page.click(`#grid-edit-table td[data-row="${EDIT_ROW}"][data-col="0"]`);
await page.keyboard.type('EDITED');
await page.keyboard.press('Enter');
await page.waitForTimeout(800);
const afterEdit = await page.evaluate(
  (r) => document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="0"]`)?.textContent,
  EDIT_ROW,
);
console.log(`edit applied: (r${EDIT_ROW}, A) = ${JSON.stringify(afterEdit)}`);
if (afterEdit !== 'EDITED') failures.push(`셀 편집이 적용되지 않았다: ${JSON.stringify(afterEdit)}`);

// ---------------------------------------------------------------- 2) 구조 변경
let remote = null;
if (MODE === 'unsubscribed') {
  // (a) 편집한 칸의 페이지(0)가 구독에서 빠지도록 멀리 스크롤 — 서버가 그 페이지의 셀을 Yjs에서
  //     지운다(윈도잉). 이때 그 키에는 **더 새로운 항목이 생기지 않는다**(삭제 표시만 붙는다).
  await page.locator('.grid-editor-scroll').evaluate((el, y) => { el.scrollTop = y; }, FAR_ROW * ROW_HEIGHT);
  await page.waitForTimeout(2000);
  // (b) 그 페이지를 구독하지 않는 원격 협업자가 행 삽입을 보낸다.
  remote = await connectRemoteCollaborator();
  console.log(`remote collaborator sees cell 0:${EDIT_ROW}:0 =`, JSON.stringify(remote.cellValue(`0:${EDIT_ROW}:0`)));
  remote.sendStructureOp({ sheet: 0, kind: 'insertRows', at: 0, count: 1 });
  await page.waitForTimeout(3000);
} else if (MODE === 'shift') {
  // 실제 UI 경로: 행 머리글 우클릭 → 메뉴.
  await page.click('#grid-edit-table th.g-edit-rownum[data-row="0"]', { button: 'right' });
  await page.waitForSelector('.grid-context-menu', { timeout: 5000 });
  await page.click('.grid-context-menu button:has-text("위에 행 삽입")');
  await page.waitForTimeout(2500);
  const afterInsert = await page.evaluate(() => {
    const t = (r) => document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="0"]`)?.textContent ?? null;
    return { r0: t(0), r1: t(1), r2: t(2), r3: t(3) };
  });
  console.log('DOM after insert:', JSON.stringify(afterInsert));
  if (afterInsert.r3 !== 'EDITED') failures.push(`행 삽입이 편집을 아래로 밀지 않았다: ${JSON.stringify(afterInsert)}`);
}

// ---------------------------------------------------------------- 3) Ctrl+Z
// 편집모드가 아닌 상태(칸 하나 클릭만)에서 누른다 — 편집 중이면 브라우저 기본 undo로 넘어간다.
const clickRow = MODE === 'unsubscribed' ? FAR_ROW + 2 : 6;
await page.click(`#grid-edit-table td[data-row="${clickRow}"][data-col="0"]`);
await page.keyboard.press('Control+z');
await page.waitForTimeout(1500);
if (remote) {
  console.log(`after Ctrl+Z, room holds cell 0:${EDIT_ROW}:0 =`, JSON.stringify(remote.cellValue(`0:${EDIT_ROW}:0`)));
}

// ---------------------------------------------------------------- 4) 진짜 검증: 다운로드 바이트
const res = await fetch(`${base}/api/view/download?token=${encodeURIComponent(token)}`);
if (!res.ok) {
  failures.push(`download failed: ${res.status}`);
} else {
  const bytes = Buffer.from(await res.arrayBuffer());
  console.log(`downloaded ${bytes.length} bytes`);
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(bytes);
  const sheet = reopened.getWorksheet('Sheet1');
  const colA = [];
  for (let r = 1; r <= Math.min(12, sheet.rowCount); r += 1) colA.push(sheet.getCell(r, 1).value ?? null);
  console.log('col A (first rows) of downloaded file:', JSON.stringify(colA));

  // 기대치는 전부 **다운로드한 파일**의 엑셀 주소(1-based) 기준.
  const expect =
    MODE === 'control'
      ? [
          ['A1', 'keep0', '그대로'],
          ['A3', 'keep2', '핵심(대조군): Ctrl+Z가 편집을 되돌려 옛 값이 돌아왔다'],
          ['A4', 'keep3', '그대로'],
        ]
      : [
          ['A1', null, '새로 삽입된 빈 행'],
          ['A2', 'keep0', '0-based 0 → 1'],
          ['A3', 'keep1', '핵심: undo가 이 칸을 덮어쓰지 않았다(고치기 전에는 keep2가 여기 써졌다)'],
          ['A4', 'EDITED', '내 편집은 밀린 자리에 그대로 살아 있다'],
          ['A5', 'keep3', '그 아래도 정상'],
        ];
  for (const [addr, want, why] of expect) {
    const actual = sheet.getCell(addr).value;
    const ok = want === null ? !actual : actual === want;
    console.log(`${ok ? 'OK  ' : 'FAIL'} ${addr} = ${JSON.stringify(actual)} (expect ${JSON.stringify(want)}) — ${why}`);
    if (!ok) failures.push(`${addr}: expected ${JSON.stringify(want)}, got ${JSON.stringify(actual)}`);
  }

  const revived = colA.filter((v) => v === 'keep2').length;
  console.log(`'keep2' occurrences: ${revived} (expect ${MODE === 'control' ? 1 : 0})`);
  if (MODE === 'control') {
    if (revived !== 1) {
      failures.push(`대조군인데 Ctrl+Z가 편집을 되돌리지 않았다 — Ctrl+Z 경로 자체가 안 먹었다는 뜻`);
    }
  } else {
    // 'keep2'는 내가 덮어썼으므로 파일 어디에도 다시 나타나선 안 된다 — 되살아나면 그게 곧
    // "undo가 낡은 키로 옛 값을 되살렸다"는 증거다.
    if (revived !== 0) failures.push(`'keep2'가 ${revived}번 되살아났다 — 낡은 undo 스택이 엉뚱한 행을 덮어썼다`);
    const keep1 = colA.filter((v) => v === 'keep1').length;
    if (keep1 !== 1) failures.push(`'keep1'이 ${keep1}번 — 조용히 사라졌다`);
  }
}

remote?.close();
await browser.close();
if (failures.length > 0) {
  console.log(`\nE2E FAILED (mode: ${MODE}, ${failures.length} problems):`);
  for (const f of failures) console.log(' -', f);
  process.exit(1);
}
console.log(`\nE2E PASSED (mode: ${MODE}) — 구조 변경 뒤의 Ctrl+Z가 엉뚱한 행을 덮어쓰지 않았다.`);
