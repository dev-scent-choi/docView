// 실행 중인 서버를 상대로, **실제 브라우저에서 우클릭 → "위에 행 삽입"**을 눌러 1600행짜리 xlsx에
// 행을 삽입하고, 그 뒤 **실제 다운로드 API가 내려주는 바이트**를 다시 열어 데이터 정합성을 확인한다.
//
// 이 확인이 겨냥하는 결함: 예전에는 클라이언트가 Yjs 셀 키를 직접 옮겼고 그 루프의 상한이
// 하드코딩된 1000이라, 1001행이 넘는 시트에서 행을 하나 넣으면 행 1000의 데이터가 사라지고
// 1001행 이후는 밀리지 않아 시트 나머지 전체가 한 칸씩 어긋났다. 게다가 Yjs는 지금 구독 중인
// 500행 페이지만 담으므로 창 밖의 행은 애초에 옮길 수가 없다 — 그래서 브라우저가 행 100 부근만
// 보고 있는 상태에서 삽입해도 행 999/1000/1500이 정확히 밀려야 한다는 것이 이 검증의 핵심이다.
//
// 사용법(서버를 먼저 같은 DATA_DIR/SECRET/PORT로 띄워 둘 것):
//   DOCVIEW_DATA_DIR=... DOCVIEW_SECRET=... DOCVIEW_PORT=... \
//     node scripts/verify-rowcol-shift-e2e.mjs [insert-row|delete-row|insert-col|delete-col]
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';
import jwt from 'jsonwebtoken';
import { chromium } from 'playwright';

const DATA_DIR = process.env.DOCVIEW_DATA_DIR;
const SECRET = process.env.DOCVIEW_SECRET;
const PORT = process.env.DOCVIEW_PORT ?? '8099';
const ROWS = 1600;
const AT_ROW = 100; // 0-based — 삽입/삭제 지점(옛 하드코딩 상한 1000과 한참 떨어진 평범한 자리)

/**
 * 네 가지 UI 동작 각각에 대해 "어떤 메뉴를 어떻게 열어 무엇을 누르고, 그 결과 파일이 어떤 모양이어야
 * 하는지"를 한자리에 적어 둔다. expect의 주소는 전부 **다운로드한 파일** 기준이다.
 */
const OPS = {
  'insert-row': {
    menu: 'row',
    item: '위에 행 삽입',
    expect: [
      ['A1', 'top', '삽입 지점 위 — 그대로'],
      ['A101', null, '새로 삽입된 빈 행'],
      ['A502', 'mid500', '0-based 500 → 501'],
      ['A1001', 'row999', '0-based 999 → 1000'],
      ['A1002', 'row1000', '핵심: 옛 코드가 덮어써서 날려버린 행'],
      ['A1502', 'row1500', '핵심: 옛 코드가 아예 밀지 않던 행'],
    ],
    markerCount: 1,
    colBCount: ROWS,
    colBIndex: 2,
  },
  'delete-row': {
    menu: 'row',
    item: '행 삭제',
    expect: [
      ['A1', 'top', '삭제 지점 위 — 그대로'],
      ['A500', 'mid500', '0-based 500 → 499'],
      ['A999', 'row999', '0-based 999 → 998'],
      ['A1000', 'row1000', '핵심: 창 밖의 행도 정확히 당겨졌다'],
      ['A1500', 'row1500', '핵심: 시트 끝까지 어긋남 없이 당겨졌다'],
    ],
    markerCount: 1,
    colBCount: ROWS - 1,
    colBIndex: 2,
  },
  'insert-col': {
    menu: 'col',
    item: '왼쪽에 열 삽입',
    expect: [
      ['A1', null, '새로 삽입된 빈 열'],
      ['B1', 'top', 'A → B'],
      ['C1', 'b', 'B → C'],
      ['B1000', 'row999', '열 삽입은 행 번호를 바꾸지 않는다'],
      ['B1001', 'row1000', '핵심: 창 밖의 행에서도 열이 밀렸다'],
      ['B1501', 'row1500', '핵심: 시트 끝 행에서도 열이 밀렸다'],
    ],
    markerCount: 1,
    markerCol: 2,
    colBCount: ROWS,
    colBIndex: 3, // 'b'가 B → C로 밀렸다
  },
  'delete-col': {
    menu: 'col',
    item: '열 삭제',
    expect: [
      ['A1', 'b', 'B열이 A로 당겨졌다'],
      ['A1501', 'b', '창 밖의 행에서도 당겨졌다'],
    ],
    markerCount: 0, // A열(표식들)이 통째로 지워졌다
    markerCol: 1,
    colBCount: ROWS,
    colBIndex: 1,
  },
};

const opName = process.argv[2] ?? 'insert-row';
const op = OPS[opName];
if (!op) {
  console.error(`unknown op: ${opName} (${Object.keys(OPS).join(' | ')})`);
  process.exit(1);
}
console.log(`=== live e2e: ${opName} ("${op.item}") ===`);

if (!DATA_DIR || !SECRET) {
  console.error('DOCVIEW_DATA_DIR / DOCVIEW_SECRET 환경변수가 필요합니다.');
  process.exit(1);
}

/** tests/roomRowColShift.test.ts의 makeTallXlsx와 같은 픽스처(같은 표식·같은 기하). */
async function makeTallXlsx() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  for (let r = 1; r <= ROWS; r += 1) sheet.getCell(r, 2).value = 'b';
  sheet.getCell('A1').value = 'top';
  sheet.getCell('A501').value = 'mid500';
  sheet.getCell('A1000').value = 'row999';
  sheet.getCell('A1001').value = 'row1000';
  sheet.getCell('A1501').value = 'row1500';
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

const id = randomUUID();
await fs.mkdir(path.join(DATA_DIR, 'originals'), { recursive: true });
await fs.mkdir(path.join(DATA_DIR, 'meta'), { recursive: true });
await fs.writeFile(path.join(DATA_DIR, 'originals', `${id}.xlsx`), await makeTallXlsx());
await fs.writeFile(
  path.join(DATA_DIR, 'meta', `${id}.json`),
  JSON.stringify({
    id,
    name: 'rowcol-shift-e2e.xlsx',
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

const browser = await chromium.launch();
const page = await browser.newPage();
const failures = [];
page.on('pageerror', (err) => {
  console.log('PAGEERROR:', err.message);
  failures.push(`pageerror: ${err.message}`);
});

await page.goto(editorUrl, { waitUntil: 'domcontentloaded' });
await page.waitForSelector('#grid-edit-table td[data-row]', { timeout: 30000 });

const initialRowCount = await page.evaluate(() =>
  document.querySelectorAll('#grid-edit-table tr').length,
);
console.log(`editor loaded; rendered <tr> count = ${initialRowCount} (가상 스크롤이라 1600보다 훨씬 작아야 한다)`);

// 대상 지점(행 100)이 화면에 보이도록 스크롤한다 — 행 높이 27px. 열 연산도 같은 자리에서
// 수행해, "지금 보고 있지 않은 페이지의 행들"이 함께 밀리는지가 그대로 시험대에 오르게 한다.
await page.locator('.grid-editor-scroll').evaluate((el) => { el.scrollTop = 100 * 27 - 200; });
await page.waitForTimeout(600); // 구독 디바운스(200ms) + 재렌더

// 실제 UI 경로: 머리글 우클릭 → 컨텍스트 메뉴 항목 클릭.
if (op.menu === 'row') {
  await page.waitForSelector(`#grid-edit-table th.g-edit-rownum[data-row="${AT_ROW}"]`, { timeout: 20000 });
  await page.click(`#grid-edit-table th.g-edit-rownum[data-row="${AT_ROW}"]`, { button: 'right' });
} else {
  // 열 머리글은 첫 열(A)을 대상으로 한다.
  await page.waitForSelector('#grid-edit-table thead th[data-col="0"]', { timeout: 20000 });
  await page.click('#grid-edit-table thead th[data-col="0"]', { button: 'right' });
}
await page.waitForSelector('.grid-context-menu', { timeout: 5000 });
const menuItems = await page.$$eval('.grid-context-menu button', (els) => els.map((e) => e.textContent));
console.log('context menu items:', JSON.stringify(menuItems));
await page.click(`.grid-context-menu button:has-text("${op.item}")`);

// 서버가 canonical을 밀고 구독 페이지를 다시 채운 뒤 구조 메타데이터를 보낸다 — 그 왕복을 기다린다.
await page.waitForTimeout(2500);

// 화면(가상 렌더)에도 반영됐는지 — 편집이 화면에서 조용히 사라지지 않았는지 최소 확인.
const domAfter = await page.evaluate((at) => {
  const cellText = (r, c) => {
    const td = document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="${c}"]`);
    return td ? td.textContent : null;
  };
  return { atPoint: cellText(at, 1), below: cellText(at + 1, 1), rendered: document.querySelectorAll('#grid-edit-table tr').length };
}, AT_ROW);
console.log('DOM after op:', JSON.stringify(domAfter));

// **진짜 검증**: 실제 다운로드 API의 바이트를 다시 열어 본다.
const res = await fetch(`${base}/api/view/download?token=${encodeURIComponent(token)}`);
if (!res.ok) {
  failures.push(`download failed: ${res.status}`);
} else {
  const bytes = Buffer.from(await res.arrayBuffer());
  console.log(`downloaded ${bytes.length} bytes`);
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(bytes);
  const sheet = reopened.getWorksheet('Sheet1');

  for (const [addr, expected, why] of op.expect) {
    const actual = sheet.getCell(addr).value;
    const ok = expected === null ? !actual : actual === expected;
    console.log(`${ok ? 'OK  ' : 'FAIL'} ${addr} = ${JSON.stringify(actual)} (expect ${JSON.stringify(expected)}) — ${why}`);
    if (!ok) failures.push(`${addr}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
  }

  // 복제·유실이 없는지 — 표식은 정확히 기대 횟수만큼, 'b'는 정확히 기대 칸수만큼.
  const markerCol = op.markerCol ?? 1;
  const counts = {};
  let bCount = 0;
  for (let r = 1; r <= sheet.rowCount; r += 1) {
    const a = sheet.getCell(r, markerCol).value;
    if (typeof a === 'string' && a !== 'b') counts[a] = (counts[a] ?? 0) + 1;
    if (sheet.getCell(r, op.colBIndex).value === 'b') bCount += 1;
  }
  console.log('marker occurrences:', JSON.stringify(counts), `'b' count in col ${op.colBIndex}:`, bCount);
  for (const marker of ['top', 'mid500', 'row999', 'row1000', 'row1500']) {
    const seen = counts[marker] ?? 0;
    if (seen !== op.markerCount) {
      failures.push(`${marker} occurs ${seen} times (expect exactly ${op.markerCount})`);
    }
  }
  if (bCount !== op.colBCount) {
    failures.push(`'b' appears ${bCount} times (expect ${op.colBCount} — a duplicate or loss shows up here)`);
  }
}

await browser.close();
if (failures.length > 0) {
  console.log(`\nE2E FAILED (${failures.length}):`);
  for (const f of failures) console.log(' -', f);
  process.exit(1);
}
console.log(`\nE2E PASSED (${opName}) — 1600행 시트에서 브라우저 UI로 조작해도 다운로드한 파일이 정확히 한 칸씩 밀렸다(창 밖 행 포함).`);
