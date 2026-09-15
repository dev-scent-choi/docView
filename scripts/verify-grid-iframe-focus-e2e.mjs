// 실행 중인 서버를 상대로, **호스트 페이지 + 그 안에 iframe으로 박은 그리드 편집기**라는
// 이 뷰어의 실제 사용 형태(DOCVIEW_ALLOWED_ORIGINS)를 그대로 만들어 두고, 진짜 마우스 휠/클릭/
// 타이핑으로 두 가지 "조용히 잘못되는" 결함을 확인한다.
//
//  iframe-scroll  — 스크롤 패치의 포커스 복원(restoreScrollWindowFocus)이 document.activeElement가
//                   <body>인지만 보고 되돌리면, "이 문서 안에 포커스가 없다"와 "포커스가 아예 이
//                   문서 밖(호스트 페이지)에 있다"를 구분하지 못한다. 사용자가 iframe을 클릭한 적도
//                   없이 호스트 입력칸에 타이핑하면서 그리드 위에 휠만 굴려도 그리드가 포커스를
//                   빼앗아, 이어지는 타자가 스프레드시트 칸에 들어가 그대로 커밋된다.
//  iframe-rebuild — 같은 구멍이 더 오래된 전체 재빌드 경로(renderTableVirtualized 끝의 무조건
//                   focus())에도 있다. 창 크기 변경/원격 구조 변경 등으로 재빌드가 돌면, 역시
//                   호스트의 포커스를 가로챈다.
//  same-doc       — 위 가드가 정작 고치려던 경우까지 막아버리면 안 된다: iframe 자신이 포커스를
//                   갖고 있을 때(칸을 클릭한 뒤)는 휠 왕복 후에도 포커스가 그 칸으로 돌아오고
//                   타이핑이 그 칸에 들어가야 한다(35f704f가 세운 동작).
//  sheet-sel      — 선택 상태(selAnchor~selFocus/headerSelection)는 시트 번호를 들고 있지 않다.
//                   시트를 바꾼 뒤 스크롤 한 번이면 이전 시트의 선택이 새 시트에 칠해지고, 더
//                   나쁘게는 툴바 서식 버튼이 새 시트의 같은 좌표 칸들에 조용히 먹는다.
//
// 사용법(서버를 먼저 같은 SECRET/PORT로, 그리고 HOST_ORIGIN을 DOCVIEW_ALLOWED_ORIGINS에 넣어 띄울 것):
//   DOCVIEW_SECRET=... DOCVIEW_PORT=8791 HOST_PORT=8792 node scripts/verify-grid-iframe-focus-e2e.mjs
import http from 'node:http';
import ExcelJS from 'exceljs';
import jwt from 'jsonwebtoken';
import { chromium } from 'playwright';

const SECRET = process.env.DOCVIEW_SECRET;
const PORT = process.env.DOCVIEW_PORT ?? '8791';
const HOST_PORT = Number(process.env.HOST_PORT ?? 8792);
const ROWS = 1200;
const COLS = 10;
const FOCUS_ROW = 8; // 0-based — 스크롤 0에서 실제로 보이는 자리
const FOCUS_COL = 1;
const ROW_H = 27; // grid-editor.ts의 ROW_HEIGHT_PX와 같아야 한다

if (!SECRET) {
  console.error('DOCVIEW_SECRET 환경변수가 필요합니다.');
  process.exit(1);
}

const base = `http://127.0.0.1:${PORT}`;
const failures = [];
function check(ok, label, detail) {
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
  if (!ok) failures.push(`${label}${detail ? ` — ${detail}` : ''}`);
}

/** 시트 2장짜리 픽스처 — 시트 간 선택 상태 누수를 보려면 두 시트가 같은 좌표에 값을 갖고 있어야
 * "엉뚱한 시트에 서식이 먹었다"를 좌표로 확인할 수 있다. */
async function makeFixture() {
  const workbook = new ExcelJS.Workbook();
  for (const name of ['Sheet1', 'Sheet2']) {
    const sheet = workbook.addWorksheet(name);
    for (let r = 1; r <= ROWS; r += 1) {
      for (let c = 1; c <= COLS; c += 1) sheet.getCell(r, c).value = `${name}-r${r - 1}c${c - 1}`;
    }
  }
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

async function uploadFixture() {
  const bytes = await makeFixture();
  const form = new FormData();
  form.append('file', new Blob([bytes]), 'grid-iframe-focus-e2e.xlsx');
  const res = await fetch(`${base}/api/documents`, { method: 'POST', body: form });
  if (!res.ok) throw new Error(`upload failed: ${res.status} ${await res.text()}`);
  const { data } = await res.json();
  for (let i = 0; i < 160; i += 1) {
    const meta = await (await fetch(`${base}/api/documents/${data.id}`)).json();
    if (meta.data?.status === 'ready') return data.id;
    if (meta.data?.status === 'failed') throw new Error('conversion failed');
    await new Promise((r) => setTimeout(r, 250));
  }
  throw new Error('conversion timed out');
}

/** 호스트 페이지를 진짜 http 출처에서 띄운다 — about:blank/data: URL은 출처가 "null"이라
 * frame-ancestors 검사에 걸려 실제 임베드 상황을 재현하지 못한다. */
function startHostServer(editorUrl) {
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>host page</title></head>
<body style="margin:0;font:14px sans-serif;padding:8px">
  <div><label>host input: <input id="host-input" style="width:520px;height:26px"></label></div>
  <iframe id="grid" src="${editorUrl.replace(/&/g, '&amp;')}" style="width:1240px;height:700px;border:1px solid #888"></iframe>
</body></html>`;
  const server = http.createServer((_req, res) => {
    res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
    res.end(html);
  });
  return new Promise((resolve) => server.listen(HOST_PORT, '127.0.0.1', () => resolve(server)));
}

/** iframe 안 좌표를 최상위 페이지 좌표로 옮긴다 — page.mouse는 최상위 페이지 기준이라
 * iframe의 화면상 위치를 더해야 진짜 그 칸 위에 커서가 놓인다. */
async function iframeOffset(page) {
  const box = await page.locator('#grid').boundingBox();
  if (!box) throw new Error('iframe not found');
  return box;
}

async function cellPoint(frame, page, row, col) {
  const rect = await frame.evaluate(
    ({ r, c }) => {
      const td = document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="${c}"]`);
      if (!td) return null;
      const b = td.getBoundingClientRect();
      return { x: b.left + b.width / 2, y: b.top + b.height / 2 };
    },
    { r: row, c: col },
  );
  if (!rect) throw new Error(`td(${row},${col}) not rendered`);
  const off = await iframeOffset(page);
  return { x: off.x + rect.x, y: off.y + rect.y };
}

/** 진짜 마우스 휠로 그리드를 굴린다(프로그램적 scrollTop 대입이 아니다 — 이 버그는 "클릭 없이
 * 휠만 굴렸을 때"가 핵심이라 입력 경로 자체가 진짜여야 의미가 있다). */
async function wheelOverGrid(page, frame, totalDelta, { steps = 14 } = {}) {
  const off = await iframeOffset(page);
  // 그리드 본문 한가운데 — 스티키 머리글/툴바가 아닌 스크롤 영역 위에 커서를 둔다.
  await page.mouse.move(off.x + off.width / 2, off.y + off.height / 2);
  const per = Math.round(totalDelta / steps);
  for (let i = 0; i < steps; i += 1) {
    await page.mouse.wheel(0, per);
    await page.waitForTimeout(40);
  }
  await page.waitForTimeout(500);
  return frame.evaluate(() => document.querySelector('.grid-editor-scroll').scrollTop);
}

const activeInfo = (ctx) =>
  ctx.evaluate(() => {
    const el = document.activeElement;
    if (!el) return { tag: 'null' };
    return { tag: el.tagName, id: el.id ?? '', row: el.dataset?.row ?? null, col: el.dataset?.col ?? null };
  });

const cellText = (frame, r, c) =>
  frame.evaluate(
    ({ r, c }) => document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="${c}"]`)?.textContent ?? null,
    { r, c },
  );

const paintCounts = (ctx) =>
  ctx.evaluate(() => ({
    range: document.querySelectorAll('#grid-edit-table .g-range-selected').length,
    col: document.querySelectorAll('#grid-edit-table .g-col-selected').length,
    row: document.querySelectorAll('#grid-edit-table .g-row-selected').length,
  }));

async function main() {
  const id = await uploadFixture();
  const token = jwt.sign(
    { docId: id, viewer: 'e2e', perms: { canPrint: true, canDownload: true, canCopy: true, canEdit: true } },
    SECRET,
    { expiresIn: 1800 },
  );
  const editorUrl = `${base}/editor-grid.html?token=${encodeURIComponent(token)}`;
  const hostServer = await startHostServer(editorUrl);

  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  page.on('pageerror', (err) => {
    console.log('PAGEERROR:', err.message);
    failures.push(`pageerror: ${err.message}`);
  });
  await page.goto(`http://127.0.0.1:${HOST_PORT}/`, { waitUntil: 'domcontentloaded' });
  const frameEl = await page.waitForSelector('#grid');
  const frame = await frameEl.contentFrame();
  await frame.waitForSelector('#grid-edit-table td[data-row]', { timeout: 30000 });
  await page.waitForTimeout(1000);
  check(true, '[setup] 호스트 페이지 안의 iframe으로 편집기가 실제로 떴다', `frame=${frame.url().slice(0, 48)}...`);

  // ------------------------------------------------------- iframe-scroll (핵심)
  console.log('\n=== [iframe-scroll] 호스트 입력칸에 타이핑하는 중 iframe 위에서 휠만 굴렸을 때 ===');
  // 1) 먼저 그리드 칸을 한 번 클릭해 둔다(activeCellKey가 생기는 조건 — 사용자가 아까 그리드를
  //    쓰다가 호스트 쪽으로 옮겨간 상황).
  const p = await cellPoint(frame, page, FOCUS_ROW, FOCUS_COL);
  await page.mouse.click(p.x, p.y);
  await page.waitForTimeout(250);
  const clicked = await activeInfo(frame);
  check(
    clicked.tag === 'TD' && clicked.row === String(FOCUS_ROW),
    '[iframe-scroll] 전제: iframe 안의 칸이 실제로 선택됐다',
    JSON.stringify(clicked),
  );
  const originalText = await cellText(frame, FOCUS_ROW, FOCUS_COL);

  // 2) 호스트 페이지의 입력칸으로 옮겨 타이핑한다(여기서부터 iframe은 클릭조차 하지 않는다).
  await page.click('#host-input');
  await page.keyboard.type('HOST-TYPED-1');
  await page.waitForTimeout(150);
  const hostFocused = await activeInfo(page);
  check(hostFocused.id === 'host-input', '[iframe-scroll] 전제: 포커스가 호스트 입력칸에 있다', JSON.stringify(hostFocused));

  // 3) iframe 위에서 "진짜 휠"만 굴린다 — 활성 칸이 창 밖으로 나갔다가 돌아오게(포커스 복원 경로가
  //    실제로 도는 조건). 클릭은 한 번도 하지 않는다.
  const away = await wheelOverGrid(page, frame, 600 * ROW_H);
  const gone = await frame.evaluate(
    ({ r, c }) => !document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="${c}"]`),
    { r: FOCUS_ROW, c: FOCUS_COL },
  );
  check(away > 1000 && gone, '[iframe-scroll] 전제: 휠로 실제로 스크롤됐고 활성 칸의 행이 창 밖으로 제거됐다', `scrollTop=${away} removed=${gone}`);
  const home = await wheelOverGrid(page, frame, -600 * ROW_H);
  check(home === 0, '[iframe-scroll] 전제: 휠로 원래 자리까지 되돌아왔다', `scrollTop=${home}`);

  const afterWheel = await Promise.all([activeInfo(page), activeInfo(frame)]);
  console.log(`  휠 직후 — 호스트 activeElement=${JSON.stringify(afterWheel[0])} / iframe activeElement=${JSON.stringify(afterWheel[1])}`);
  check(
    afterWheel[0].id === 'host-input',
    '[iframe-scroll] 휠만 굴려도 호스트 입력칸이 포커스를 유지한다',
    JSON.stringify(afterWheel[0]),
  );
  check(
    !(afterWheel[1].tag === 'TD'),
    '[iframe-scroll] 그리드가 (클릭도 없이) 포커스를 가져가지 않았다',
    `iframe activeElement=${JSON.stringify(afterWheel[1])}`,
  );

  // 4) 진짜 중요한 것: 이어서 친 글자가 어디로 가는가.
  await page.keyboard.type('-AFTER-WHEEL');
  await page.keyboard.press('Enter');
  await page.waitForTimeout(500);
  const hostValue = await page.inputValue('#host-input');
  const cellNow = await cellText(frame, FOCUS_ROW, FOCUS_COL);
  console.log(`  호스트 입력칸 = ${JSON.stringify(hostValue)} / 셀(${FOCUS_ROW},${FOCUS_COL}) = ${JSON.stringify(cellNow)}`);
  check(
    hostValue === 'HOST-TYPED-1-AFTER-WHEEL',
    '[iframe-scroll] 핵심: 휠 뒤에 친 글자가 호스트 입력칸에 그대로 들어간다',
    `실제=${JSON.stringify(hostValue)}`,
  );
  check(
    cellNow === originalText,
    '[iframe-scroll] 핵심: 스프레드시트 칸의 값이 바뀌지 않았다(엉뚱한 곳에 커밋되지 않았다)',
    `실제=${JSON.stringify(cellNow)} 기대=${JSON.stringify(originalText)}`,
  );

  // ------------------------------------------------------- iframe-rebuild
  console.log('\n=== [iframe-rebuild] 호스트에 포커스가 있는 채로 전체 재빌드가 돌 때 ===');
  // 전체 재빌드(renderTableVirtualized) 경로도 같은 구멍이 있었다 — 여기서는 창 크기 변경으로
  // 재빌드를 유발한다(window resize → scheduleViewportRefit → renderTableVirtualized).
  const p2 = await cellPoint(frame, page, FOCUS_ROW, FOCUS_COL + 1);
  await page.mouse.click(p2.x, p2.y);
  await page.waitForTimeout(250);
  const original2 = await cellText(frame, FOCUS_ROW, FOCUS_COL + 1);
  await page.click('#host-input');
  await page.fill('#host-input', 'HOST-TYPED-2');
  await page.waitForTimeout(150);
  // 전체 재빌드가 실제로 돌았는지 세어 둔다 — tableEl에 <colgroup>이 새로 꽂히는 것이 전체
  // 재빌드(renderTableVirtualized)의 표식이다(증분 패치는 tbody만 건드린다).
  await frame.evaluate(() => {
    const table = document.getElementById('grid-edit-table');
    window.__rebuilds = 0;
    window.__mo = new MutationObserver((recs) => {
      for (const r of recs) {
        if (r.target === table && [...r.addedNodes].some((n) => n.tagName === 'COLGROUP')) window.__rebuilds += 1;
      }
    });
    window.__mo.observe(table, { childList: true });
  });
  // iframe "자체"의 크기를 호스트 쪽에서 바꾼다 — 그래야 iframe 안 window가 resize를 받아
  // scheduleViewportRefit → renderTableVirtualized(전체 재빌드)가 돈다. (최상위 뷰포트만 줄이면
  // 픽셀로 고정된 iframe의 내부 크기는 그대로라 재빌드가 아예 일어나지 않아 시험이 무의미해진다.)
  // iframe은 여전히 한 번도 클릭하지 않는다.
  await page.evaluate(() => {
    document.getElementById('grid').style.height = '520px';
  });
  await page.waitForTimeout(900);
  const rebuilds = await frame.evaluate(() => {
    window.__mo.disconnect();
    return window.__rebuilds;
  });
  check(rebuilds > 0, '[iframe-rebuild] 전제: 창 크기 변경으로 전체 재빌드가 실제로 돌았다', `${rebuilds}회`);
  const rebuiltActive = await Promise.all([activeInfo(page), activeInfo(frame)]);
  console.log(`  재빌드 직후 — 호스트=${JSON.stringify(rebuiltActive[0])} / iframe=${JSON.stringify(rebuiltActive[1])}`);
  check(
    rebuiltActive[0].id === 'host-input' && rebuiltActive[1].tag !== 'TD',
    '[iframe-rebuild] 전체 재빌드가 호스트의 포커스를 빼앗지 않는다',
    `host=${JSON.stringify(rebuiltActive[0])} iframe=${JSON.stringify(rebuiltActive[1])}`,
  );
  await page.keyboard.type('-AFTER-REBUILD');
  await page.keyboard.press('Enter');
  await page.waitForTimeout(500);
  const hostValue2 = await page.inputValue('#host-input');
  const cellNow2 = await cellText(frame, FOCUS_ROW, FOCUS_COL + 1);
  check(
    hostValue2 === 'HOST-TYPED-2-AFTER-REBUILD' && cellNow2 === original2,
    '[iframe-rebuild] 재빌드 뒤에 친 글자도 호스트 입력칸에 들어간다(셀은 그대로)',
    `host=${JSON.stringify(hostValue2)} cell=${JSON.stringify(cellNow2)}`,
  );
  await page.evaluate(() => {
    document.getElementById('grid').style.height = '700px';
  });
  await page.waitForTimeout(700);

  // ------------------------------------------------------- same-doc (회귀 방지)
  console.log('\n=== [same-doc] iframe 자신이 포커스를 가진 경우에는 원래대로 포커스가 돌아와야 한다 ===');
  const p3 = await cellPoint(frame, page, FOCUS_ROW + 2, FOCUS_COL);
  await page.mouse.click(p3.x, p3.y);
  await page.waitForTimeout(250);
  const before3 = await activeInfo(frame);
  check(
    before3.tag === 'TD' && before3.row === String(FOCUS_ROW + 2),
    '[same-doc] 전제: 칸이 선택됐다(= iframe이 포커스를 갖고 있다)',
    JSON.stringify(before3),
  );
  await wheelOverGrid(page, frame, 600 * ROW_H);
  const backTop = await wheelOverGrid(page, frame, -600 * ROW_H);
  const back3 = await activeInfo(frame);
  console.log(`  휠 왕복 뒤(scrollTop=${backTop}) iframe activeElement=${JSON.stringify(back3)}`);
  check(
    back3.tag === 'TD' && back3.row === String(FOCUS_ROW + 2) && back3.col === String(FOCUS_COL),
    '[same-doc] 휠 왕복 뒤 활성 칸에 포커스가 복원된다(35f704f가 고친 동작)',
    JSON.stringify(back3),
  );
  await page.keyboard.type('KEEPME');
  await page.keyboard.press('Enter');
  await page.waitForTimeout(600);
  const typed3 = await cellText(frame, FOCUS_ROW + 2, FOCUS_COL);
  check(typed3 === 'KEEPME', '[same-doc] 휠 왕복 직후 친 값이 그 칸에 그대로 들어간다', `실제=${JSON.stringify(typed3)}`);

  // ------------------------------------------------------- sheet-sel
  console.log('\n=== [sheet-sel] 시트를 바꾼 뒤 스크롤하면 이전 시트의 선택이 따라오는가 ===');
  const sheetPage = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  sheetPage.on('pageerror', (err) => {
    console.log('PAGEERROR(sheet):', err.message);
    failures.push(`pageerror(sheet): ${err.message}`);
  });
  await sheetPage.goto(editorUrl, { waitUntil: 'domcontentloaded' });
  await sheetPage.waitForSelector('#grid-edit-table td[data-row]', { timeout: 30000 });
  await sheetPage.waitForTimeout(800);

  // Sheet1에서 C열 전체를 선택한다(머리글 클릭 — selFocusRow가 시트 끝까지 잡히는 가장 넓은 선택).
  await sheetPage.click('#grid-edit-table thead th[data-col="2"]');
  await sheetPage.waitForTimeout(300);
  const onSheet1 = await paintCounts(sheetPage);
  console.log(`  Sheet1 열 선택 직후: ${JSON.stringify(onSheet1)}`);
  check(onSheet1.range > 0 && onSheet1.col > 0, '[sheet-sel] 전제: Sheet1에서 열 전체 선택이 칠해졌다', JSON.stringify(onSheet1));

  // Sheet2로 전환
  await sheetPage.locator('.grid-sheet-tab').nth(1).click();
  await sheetPage.waitForTimeout(700);
  const justSwitched = await paintCounts(sheetPage);
  console.log(`  Sheet2 전환 직후: ${JSON.stringify(justSwitched)}`);
  check(
    justSwitched.range === 0 && justSwitched.col === 0 && justSwitched.row === 0,
    '[sheet-sel] 전제: 시트 전환 직후에는 아무것도 칠해져 있지 않다',
    JSON.stringify(justSwitched),
  );

  // Sheet2에서 진짜 휠로 한 번 스크롤 — 여기서 이전 시트의 선택이 되살아나면 결함이다.
  await sheetPage.mouse.move(640, 500);
  for (let i = 0; i < 6; i += 1) {
    await sheetPage.mouse.wheel(0, 400);
    await sheetPage.waitForTimeout(50);
  }
  await sheetPage.waitForTimeout(600);
  const afterScroll = await paintCounts(sheetPage);
  console.log(`  Sheet2에서 스크롤 뒤: ${JSON.stringify(afterScroll)} (결함이 있으면 수십~수백 칸)`);
  check(
    afterScroll.range === 0 && afterScroll.col === 0 && afterScroll.row === 0,
    '[sheet-sel] 핵심: 다른 시트에서 잡아 둔 선택이 새 시트에 칠해지지 않는다',
    JSON.stringify(afterScroll),
  );

  // 눈에 보이는 것보다 더 나쁜 쪽 — 툴바 서식이 "보이지 않는 이전 시트 범위"에 먹는지.
  await sheetPage.click('#btn-cell-bold');
  await sheetPage.waitForTimeout(500);
  const bolded = await sheetPage.evaluate(
    () => [...document.querySelectorAll('#grid-edit-table td[data-row]')].filter((td) => td.style.fontWeight === '700').length,
  );
  console.log(`  Sheet2에서 굵게를 눌렀을 때 실제로 굵어진 칸 = ${bolded}`);
  check(bolded === 0, '[sheet-sel] 핵심: 시트 전환 뒤 툴바 서식이 엉뚱한 칸에 조용히 먹지 않는다', `${bolded}칸`);

  // 같은 시트 안에서의 선택 유지(35f704f가 세운 동작)는 그대로여야 한다. 칠해진 칸 수는 "지금
  // 렌더된 행 수"에 비례하므로(가상 스크롤), 반드시 **같은 스크롤 위치**로 돌아와서 비교한다 —
  // 위쪽 끝(scrollTop=0)은 위쪽 오버스캔이 잘려 렌더 행 자체가 더 적다.
  const sameSheetScroll = await sheetPage.evaluate(() => document.querySelector('.grid-editor-scroll').scrollTop);
  await sheetPage.click('#grid-edit-table thead th[data-col="1"]');
  await sheetPage.waitForTimeout(300);
  const sameSheetBefore = await paintCounts(sheetPage);
  await sheetPage.mouse.move(640, 500);
  for (let i = 0; i < 10; i += 1) {
    await sheetPage.mouse.wheel(0, 600);
    await sheetPage.waitForTimeout(40);
  }
  await sheetPage.waitForTimeout(400);
  for (let i = 0; i < 10; i += 1) {
    await sheetPage.mouse.wheel(0, -600);
    await sheetPage.waitForTimeout(40);
  }
  await sheetPage.waitForTimeout(600);
  const sameSheetScrollBack = await sheetPage.evaluate(() => document.querySelector('.grid-editor-scroll').scrollTop);
  const sameSheetAfter = await paintCounts(sheetPage);
  console.log(
    `  같은 시트 휠 왕복(scrollTop ${sameSheetScroll} → ${sameSheetScrollBack}): ${JSON.stringify(sameSheetBefore)} → ${JSON.stringify(sameSheetAfter)}`,
  );
  check(
    sameSheetScrollBack === sameSheetScroll,
    '[sheet-sel] 전제: 휠 왕복이 정확히 같은 스크롤 위치로 돌아왔다(칠해진 칸 수를 비교할 수 있는 조건)',
    `${sameSheetScroll} → ${sameSheetScrollBack}`,
  );
  check(
    sameSheetAfter.range === sameSheetBefore.range && sameSheetAfter.col === sameSheetBefore.col && sameSheetBefore.range > 0,
    '[sheet-sel] 회귀 방지: 같은 시트 안에서는 휠 왕복 뒤에도 선택이 그대로 남는다',
    `${JSON.stringify(sameSheetBefore)} → ${JSON.stringify(sameSheetAfter)}`,
  );

  // ------------------------------------------------------- sheet-ops (나머지 시트 전환 경로 스모크)
  // 활성 시트가 바뀌는 경로는 탭 클릭 말고도 셋 더 있다(시트 추가/삭제/검색 결과 점프) — 같은
  // 선택 초기화를 붙였으므로, 그 경로들이 예외 없이 돌고 새 시트에 선택이 남지 않는지 확인한다.
  console.log('\n=== [sheet-ops] 시트 추가·검색 점프·시트 삭제에서도 선택 초기화가 안전한가 ===');
  await sheetPage.click('.grid-sheet-tab-add');
  await sheetPage.waitForTimeout(1200);
  const added = {
    tabs: await sheetPage.locator('.grid-sheet-tab').count(),
    paint: await paintCounts(sheetPage),
  };
  console.log(`  시트 추가 후: 탭 ${added.tabs}개 / ${JSON.stringify(added.paint)}`);
  check(
    added.tabs === 3 && added.paint.range === 0 && added.paint.col === 0 && added.paint.row === 0,
    '[sheet-ops] 시트 추가(+) 뒤 새 시트로 전환되고 이전 시트의 선택이 따라오지 않는다',
    JSON.stringify(added),
  );

  await sheetPage.click('#search-btn');
  await sheetPage.fill('#search-input', 'Sheet1-r3c4');
  await sheetPage.click('#search-next');
  await sheetPage.waitForTimeout(1200);
  const jumped = {
    active: await sheetPage.locator('.grid-sheet-tab.active').textContent(),
    paint: await paintCounts(sheetPage),
  };
  console.log(`  검색 점프 후: 활성 탭=${jumped.active} / ${JSON.stringify(jumped.paint)}`);
  check(
    jumped.active === 'Sheet1' && jumped.paint.range === 0,
    '[sheet-ops] 검색 결과로 다른 시트에 건너뛴 뒤에도 이전 시트의 선택이 따라오지 않는다',
    JSON.stringify(jumped),
  );
  await sheetPage.click('#search-close');

  sheetPage.once('dialog', (d) => d.accept());
  await sheetPage.locator('.grid-sheet-tab').nth(2).click({ button: 'right' });
  await sheetPage.waitForTimeout(200);
  await sheetPage.locator('.sheet-context-menu button', { hasText: '삭제' }).first().click();
  await sheetPage.waitForTimeout(1500);
  const deleted = {
    tabs: await sheetPage.locator('.grid-sheet-tab').count(),
    paint: await paintCounts(sheetPage),
  };
  console.log(`  시트 삭제 후: 탭 ${deleted.tabs}개 / ${JSON.stringify(deleted.paint)}`);
  check(deleted.tabs === 2, '[sheet-ops] 시트 삭제 경로도 예외 없이 동작한다', JSON.stringify(deleted));

  await browser.close();
  hostServer.close();
}

await main();
if (failures.length > 0) {
  console.log(`\nFAILED (${failures.length}):`);
  for (const f of failures) console.log(' -', f);
  process.exit(1);
}
console.log('\nALL CHECKS PASSED');
