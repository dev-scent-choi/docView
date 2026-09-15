// 실행 중인 서버를 상대로, **실제 브라우저에서 클릭·스크롤·타이핑**을 그대로 흉내 내
// 그리드 편집기의 세 가지 결함을 한 번에 확인한다(전부 "조용히 잘못되는" 부류라 화면만 봐서는
// 티가 안 나고, 실제 DOM 상태를 직접 재 봐야만 드러난다).
//
//  spill  — 값이 있는 칸의 넘침 표시용 span(.g-cell-spill)이 오른쪽 빈 칸들 "위"까지 넓어지는데,
//           pointer-events가 살아 있으면 그 빈 칸을 클릭해도 클릭이 span에게 먹혀
//           closest('td')가 **왼쪽(원래 값이 있는) 칸**으로 해석된다 → 사용자가 친 값이 엉뚱한
//           칸에 들어간다. (동시편집 중 상대가 옆 칸을 채웠다 지웠다 하면 스필 폭이 실시간으로
//           바뀌므로 "협업할 때만 생기는 것처럼" 보였다.)
//  focus  — 스크롤 전용 증분 렌더(updateScrollWindow)가 활성 칸의 <tr>을 창 밖으로 내보냈다가
//           다시 넣을 때 포커스를 되돌려주지 않아, 돌아온 뒤 타이핑이 <body>로 흘러 통째로 사라진다.
//  range  — 같은 함수가 선택 범위 하이라이트도 다시 칠하지 않아, 화면에는 아무것도 선택돼 있지
//           않은데 내부 선택 상태(selAnchor~selFocus)는 그대로 남는다 → 굵게 등 서식 버튼이
//           "보이지 않는" 범위 전체에 조용히 적용된다.
//  perf   — 위 수정이 81be4d8의 스크롤 성능 개선(전체 재빌드 → 증분 패치)을 되돌리지 않았는지
//           longtask 기준으로 다시 잰다.
//
// 사용법(서버를 먼저 같은 SECRET/PORT로 띄워 둘 것 — 저장소는 memory여도 된다):
//   DOCVIEW_SECRET=... DOCVIEW_PORT=8123 node scripts/verify-grid-scroll-selection-e2e.mjs [all|spill|focus|range|perf]
import ExcelJS from 'exceljs';
import jwt from 'jsonwebtoken';
import { chromium } from 'playwright';

const SECRET = process.env.DOCVIEW_SECRET;
const PORT = process.env.DOCVIEW_PORT ?? '8123';
const ROWS = 1200;
// 열이 많을수록 "전체 재빌드"와 "증분 패치"의 차이가 커진다(재빌드 비용은 행당 <td> 수에 비례) —
// perf 측정이 두 경로를 확실히 구분할 수 있을 만큼 넓게 잡는다.
const COLS = 25;
const SPILL_ROW = 5; // 0-based — 이 행만 A열에 값이 있고 나머지 열은 비어 있다(스필이 생기는 조건)
const SPILL_TEXT = 'SPILLSOURCE-넘침-테스트-값';
const FOCUS_ROW = 12; // 0-based — 포커스 복원 시험 대상(스크롤 0에서 실제로 화면에 보이는 자리여야 한다)
const FOCUS_COL = 1;
const RANGE_R0 = 105;
const RANGE_R1 = 115;
const RANGE_C0 = 0;
const RANGE_C1 = 3;
const RANGE_CELLS = (RANGE_R1 - RANGE_R0 + 1) * (RANGE_C1 - RANGE_C0 + 1);
const ROW_H = 27; // grid-editor.ts의 ROW_HEIGHT_PX와 같아야 한다

const which = process.argv[2] ?? 'all';
const wants = (name) => which === 'all' || which === name;

if (!SECRET) {
  console.error('DOCVIEW_SECRET 환경변수가 필요합니다.');
  process.exit(1);
}

/** A~E 5열 × 400행을 전부 채운다 — 단 SPILL_ROW만 A열 하나만 값이 있고 B~E는 비운다.
 * "전부 채움"이 기본인 이유: 값이 있는 칸은 오른쪽이 비어 있을 때만 스필 span이 넓어지므로,
 * 다른 시험(focus/range)의 클릭이 스필에 가로채이지 않게 하려면 주변을 다 채워 두는 쪽이 안전하다. */
async function makeFixture() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  for (let r = 1; r <= ROWS; r += 1) {
    for (let c = 1; c <= COLS; c += 1) {
      if (r - 1 === SPILL_ROW) continue;
      sheet.getCell(r, c).value = `r${r - 1}c${c - 1}`;
    }
  }
  sheet.getCell(SPILL_ROW + 1, 1).value = SPILL_TEXT; // B~E는 비운 채로 둔다
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

const base = `http://127.0.0.1:${PORT}`;

/** 실제 업로드 API로 문서를 넣는다 — DOCVIEW_STORAGE=memory에서도 그대로 동작한다(디스크 시딩은 안 됨). */
async function uploadFixture() {
  const bytes = await makeFixture();
  const form = new FormData();
  form.append('file', new Blob([bytes]), 'grid-scroll-selection-e2e.xlsx');
  const res = await fetch(`${base}/api/documents`, { method: 'POST', body: form });
  if (!res.ok) throw new Error(`upload failed: ${res.status} ${await res.text()}`);
  const { data } = await res.json();
  for (let i = 0; i < 120; i += 1) {
    const meta = await (await fetch(`${base}/api/documents/${data.id}`)).json();
    if (meta.data?.status === 'ready') return data.id;
    if (meta.data?.status === 'failed') throw new Error('conversion failed');
    await new Promise((r) => setTimeout(r, 250));
  }
  throw new Error('conversion timed out');
}

const failures = [];
const notes = [];
function check(ok, label, detail) {
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
  if (!ok) failures.push(`${label}${detail ? ` — ${detail}` : ''}`);
}

/** scrollTop을 옮기고, 스크롤 리스너의 rAF 재렌더가 끝날 때까지 기다린다. */
async function scrollTo(page, top) {
  await page.locator('.grid-editor-scroll').evaluate((el, t) => { el.scrollTop = t; }, top);
  await page.waitForTimeout(400);
}

/** td의 기하학적 중심을 실제 마우스로 클릭한다 — Playwright의 click()은 "가려진 요소"를
 * actionability 검사에서 걸러버려, 지금 확인하려는 "가려져서 엉뚱한 칸이 눌린다"는 현상 자체를
 * 재현할 수 없다. 좌표를 직접 계산해 mouse.click을 쓴다. */
async function clickCellCenter(page, row, col) {
  const box = await cellCenter(page, row, col);
  await page.mouse.click(box.x, box.y);
  await page.waitForTimeout(150);
}

/** 칸의 화면 중심 좌표 — 가상 스크롤 때문에 "DOM에는 있지만(오버스캔) 화면 밖"인 칸이 흔해서,
 * 그런 자리를 클릭하면 엉뚱한 요소가 눌려 시험 자체가 무의미해진다. 실제 눈에 보이는 자리로
 * 먼저 스크롤을 맞춰 두고(행 높이·머리글 높이를 어림하는 대신 실제 rect 차이만큼 밀어서 — 스티키
 * thead/틀고정이 있어도 정확하다), 그래도 안 보이면 바로 실패시킨다.
 * lastRow를 주면 그 행까지 한 화면에 다 들어오도록 맞춘다(범위 드래그용). */
async function cellCenter(page, row, col, lastRow = row) {
  const measure = () =>
    page.evaluate(
      ({ r, c, lr }) => {
        const host = document.querySelector('.grid-editor-scroll');
        const td = document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="${c}"]`);
        const lastTd = document.querySelector(`#grid-edit-table td[data-row="${lr}"][data-col="${c}"]`);
        if (!td) return null;
        const b = td.getBoundingClientRect();
        const lb = (lastTd ?? td).getBoundingClientRect();
        const h = host.getBoundingClientRect();
        return {
          x: b.left + b.width / 2,
          y: b.top + b.height / 2,
          top: b.top,
          bottom: lb.bottom,
          hostTop: h.top,
          hostBottom: h.bottom,
          left: b.left,
          right: b.right,
          hostLeft: h.left,
          hostRight: h.right,
          scrollTop: host.scrollTop,
        };
      },
      { r: row, c: col, lr: lastRow },
    );

  const MARGIN = 30; // 스티키 머리글/틀고정에 살짝 가려지는 자리를 피한다
  for (let attempt = 0; attempt < 5; attempt += 1) {
    let m = await measure();
    if (!m) {
      // 아예 렌더되지 않았으면(창 밖) 일단 그 행 근처로 스크롤해 렌더시킨다.
      await scrollTo(page, Math.max(0, row * ROW_H - 200));
      m = await measure();
      if (!m) throw new Error(`td(${row},${col}) not rendered even after scrolling near it`);
    }
    const fitsV = m.top >= m.hostTop + MARGIN && m.bottom <= m.hostBottom - MARGIN;
    const fitsH = m.left >= m.hostLeft && m.right <= m.hostRight;
    if (fitsV && fitsH) return m;
    if (!fitsH) throw new Error(`td(${row},${col}) is horizontally off-screen — 시험 좌표를 바꿔야 한다`);
    // 세로로만 어긋났으면 그 차이만큼 정확히 밀어준다.
    const delta = m.top - (m.hostTop + MARGIN);
    await scrollTo(page, Math.max(0, m.scrollTop + delta));
  }
  throw new Error(`td(${row},${col}) could not be scrolled into view`);
}

/** 지금 포커스를 갖고 있는 요소를 "행/열"로 알려준다(칸이 아니면 태그 이름만). */
const activeCellInfo = (page) =>
  page.evaluate(() => {
    const el = document.activeElement;
    if (!el) return { tag: 'null' };
    return {
      tag: el.tagName,
      row: el.dataset?.row ?? null,
      col: el.dataset?.col ?? null,
    };
  });

const cellText = (page, r, c) =>
  page.evaluate(
    ({ r, c }) => document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="${c}"]`)?.textContent ?? null,
    { r, c },
  );

async function main() {
  const id = await uploadFixture();
  const token = jwt.sign(
    { docId: id, viewer: 'e2e', perms: { canPrint: true, canDownload: true, canCopy: true, canEdit: true } },
    SECRET,
    { expiresIn: 1800 },
  );
  const editorUrl = `${base}/editor-grid.html?token=${encodeURIComponent(token)}`;

  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
  page.on('pageerror', (err) => {
    console.log('PAGEERROR:', err.message);
    failures.push(`pageerror: ${err.message}`);
  });
  await page.goto(editorUrl, { waitUntil: 'domcontentloaded' });
  await page.waitForSelector('#grid-edit-table td[data-row]', { timeout: 30000 });
  await page.waitForTimeout(800);

  // ---------------------------------------------------------------- spill
  if (wants('spill')) {
    console.log('\n=== [spill] 넘침 span이 오른쪽 빈 칸의 클릭을 가로채는가 ===');
    await scrollTo(page, 0);
    const spillWidth = await page.evaluate(
      ({ r }) => {
        const td = document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="0"]`);
        const span = td?.querySelector('.g-cell-spill');
        return span ? { spanW: span.getBoundingClientRect().width, tdW: td.getBoundingClientRect().width } : null;
      },
      { r: SPILL_ROW },
    );
    console.log(`  spill span width=${spillWidth?.spanW} vs td width=${spillWidth?.tdW} (span이 더 넓어야 이 시험이 성립한다)`);
    check(
      spillWidth !== null && spillWidth.spanW > spillWidth.tdW + 1,
      '[spill] 전제: 넘침 span이 자기 칸보다 실제로 넓다',
      JSON.stringify(spillWidth),
    );

    // 스필에 덮인 이웃 칸(C열, 0-based 2)의 한가운데를 실제로 클릭한다.
    await clickCellCenter(page, SPILL_ROW, 2);
    const focused = await activeCellInfo(page);
    console.log(`  클릭 후 activeElement: ${JSON.stringify(focused)} (기대: row=${SPILL_ROW}, col=2)`);
    check(
      focused.tag === 'TD' && focused.row === String(SPILL_ROW) && focused.col === '2',
      '[spill] 스필에 덮인 칸을 클릭하면 그 칸 자신이 선택된다',
      `실제 row=${focused.row} col=${focused.col}`,
    );

    // 그 상태로 바로 타이핑 → 값이 어디에 들어가는가(사용자가 겪은 증상 그대로).
    await page.keyboard.type('TYPEDHERE');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(400);
    const at02 = await cellText(page, SPILL_ROW, 2);
    const at00 = await cellText(page, SPILL_ROW, 0);
    console.log(`  타이핑 후: (${SPILL_ROW},2)=${JSON.stringify(at02)}  (${SPILL_ROW},0)=${JSON.stringify(at00)}`);
    check(at02 === 'TYPEDHERE', '[spill] 친 값이 클릭한 칸에 들어간다', `(${SPILL_ROW},2)=${JSON.stringify(at02)}`);
    check(
      at00 === SPILL_TEXT,
      '[spill] 왼쪽(스필 소유) 칸의 값은 덮이지 않았다',
      `(${SPILL_ROW},0)=${JSON.stringify(at00)}`,
    );
    // 다음 시험을 위해 되돌린다.
    await clickCellCenter(page, SPILL_ROW, 2);
    await page.keyboard.press('Delete');
    await page.waitForTimeout(300);
  }

  // ---------------------------------------------------------------- focus
  if (wants('focus')) {
    console.log('\n=== [focus] 스크롤 왕복 뒤 활성 칸에 포커스가 돌아오는가 ===');
    await scrollTo(page, 0);
    await clickCellCenter(page, FOCUS_ROW, FOCUS_COL);
    const before = await activeCellInfo(page);
    check(
      before.tag === 'TD' && before.row === String(FOCUS_ROW),
      '[focus] 전제: 대상 칸이 실제로 선택됐다',
      JSON.stringify(before),
    );

    await scrollTo(page, 600 * ROW_H); // 활성 칸이 창 밖으로 완전히 나가게
    const away = await page.evaluate(
      ({ r, c }) => ({
        rendered: !!document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="${c}"]`),
        active: document.activeElement?.tagName ?? 'null',
      }),
      { r: FOCUS_ROW, c: FOCUS_COL },
    );
    console.log(`  멀리 스크롤한 뒤: ${JSON.stringify(away)} (기대: rendered=false, active=BODY)`);
    check(!away.rendered, '[focus] 전제: 활성 칸의 행이 실제로 창 밖으로 제거됐다', JSON.stringify(away));

    await scrollTo(page, 0); // 다시 돌아온다
    const back = await activeCellInfo(page);
    console.log(`  돌아온 뒤 activeElement: ${JSON.stringify(back)} (기대: row=${FOCUS_ROW}, col=${FOCUS_COL})`);
    check(
      back.tag === 'TD' && back.row === String(FOCUS_ROW) && back.col === String(FOCUS_COL),
      '[focus] 스크롤 왕복 뒤 활성 칸에 포커스가 복원된다',
      JSON.stringify(back),
    );

    // 진짜 중요한 것: 그 상태로 친 글자가 실제로 그 칸에 들어가는가(아니면 통째로 사라지는가).
    await page.keyboard.type('KEEPME');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(400);
    const typed = await cellText(page, FOCUS_ROW, FOCUS_COL);
    console.log(`  타이핑 후 (${FOCUS_ROW},${FOCUS_COL}) = ${JSON.stringify(typed)}`);
    check(typed === 'KEEPME', '[focus] 스크롤 왕복 직후 친 값이 사라지지 않는다', `실제=${JSON.stringify(typed)}`);
  }

  // ---------------------------------------------------------------- range
  if (wants('range')) {
    console.log('\n=== [range] 스크롤 왕복 뒤 선택 범위가 다시 칠해지는가 (보이는 것 == 적용되는 것) ===');
    await scrollTo(page, 100 * ROW_H);
    // 드래그로 범위 선택 — mousedown/move/up을 실제 마우스로. 범위의 첫 행부터 마지막 행까지가
    // 한 화면에 다 들어오게 맞춘 뒤(cellCenter의 lastRow), 그 스크롤 위치를 왕복의 기준으로 삼는다.
    const p0 = await cellCenter(page, RANGE_R0, RANGE_C0, RANGE_R1);
    const p1 = await cellCenter(page, RANGE_R1, RANGE_C1, RANGE_R1);
    const homeScroll = await page.locator('.grid-editor-scroll').evaluate((el) => el.scrollTop);
    await page.mouse.move(p0.x, p0.y);
    await page.mouse.down();
    await page.mouse.move((p0.x + p1.x) / 2, (p0.y + p1.y) / 2, { steps: 5 });
    await page.mouse.move(p1.x, p1.y, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(250);

    const selectedBefore = await page.evaluate(() => document.querySelectorAll('#grid-edit-table .g-range-selected').length);
    console.log(`  드래그 직후 하이라이트된 칸 수 = ${selectedBefore} (기대 ${RANGE_CELLS})`);
    check(selectedBefore === RANGE_CELLS, '[range] 전제: 드래그 선택이 실제로 칠해졌다', `${selectedBefore}`);

    await scrollTo(page, 600 * ROW_H);
    const goneMid = await page.evaluate(
      ({ r }) => !document.querySelector(`#grid-edit-table td[data-row="${r}"]`),
      { r: RANGE_R0 },
    );
    check(goneMid, '[range] 전제: 선택 범위의 행이 실제로 창 밖으로 제거됐다', `removed=${goneMid}`);
    await scrollTo(page, homeScroll);
    const selectedAfter = await page.evaluate(() => document.querySelectorAll('#grid-edit-table .g-range-selected').length);
    console.log(`  스크롤 왕복 뒤 하이라이트된 칸 수 = ${selectedAfter} (기대 ${RANGE_CELLS}, 결함이 있으면 0)`);
    check(
      selectedAfter === RANGE_CELLS,
      '[range] 스크롤 왕복 뒤에도 선택 범위가 그대로 보인다',
      `${selectedAfter}/${RANGE_CELLS}`,
    );

    // 이제 굵게를 눌러, "실제로 서식이 먹은 칸 수"와 "화면에 선택돼 보이던 칸 수"가 일치하는지 본다.
    await page.click('#btn-cell-bold');
    await page.waitForTimeout(400);
    const bolded = await page.evaluate(() => {
      const tds = [...document.querySelectorAll('#grid-edit-table td[data-row]')];
      const hit = tds.filter((td) => td.style.fontWeight === '700');
      const rows = hit.map((td) => Number(td.dataset.row));
      const cols = hit.map((td) => Number(td.dataset.col));
      return {
        count: hit.length,
        r0: Math.min(...rows), r1: Math.max(...rows),
        c0: Math.min(...cols), c1: Math.max(...cols),
      };
    });
    console.log(`  굵게 적용된 칸 수 = ${bolded.count} (범위 r${bolded.r0}..${bolded.r1} c${bolded.c0}..${bolded.c1})`);
    check(
      bolded.count === selectedAfter,
      '[range] 핵심: 서식이 실제로 먹은 칸 수 == 화면에 선택돼 보이던 칸 수',
      `적용 ${bolded.count} vs 표시 ${selectedAfter} — 다르면 "보이지 않는 범위에 조용히 서식이 적용"되는 상태다`,
    );
    check(
      bolded.count === RANGE_CELLS && bolded.r0 === RANGE_R0 && bolded.r1 === RANGE_R1,
      '[range] 원래 선택했던 범위에 정확히(그 이상도 이하도 아니게) 적용됐다',
      JSON.stringify(bolded),
    );
  }

  // ---------------------------------------------------------------- column
  // 열 머리글 클릭(selectColumn)은 선택 범위를 selFocusRow = rowCount-1, 즉 시트 끝까지로 잡는다 —
  // "보이지 않는 범위에 서식이 먹는" 위험이 가장 큰 경우이자, 스크롤 프레임마다 그 범위를 순진하게
  // 다시 칠하면 성능이 가장 크게 망가지는 경우이기도 하다. 둘 다 한 번에 확인한다.
  if (wants('column')) {
    console.log('\n=== [column] 열 전체 선택 상태로 스크롤 왕복 ===');
    await scrollTo(page, 0);
    await page.click('#grid-edit-table thead th[data-col="2"]');
    await page.waitForTimeout(300);
    const paintedBefore = await page.evaluate(() => ({
      range: document.querySelectorAll('#grid-edit-table td.g-range-selected').length,
      col: document.querySelectorAll('#grid-edit-table td.g-col-selected').length,
    }));
    console.log(`  열 선택 직후: 범위표시 ${paintedBefore.range}칸 / 열표시 ${paintedBefore.col}칸`);
    check(paintedBefore.range > 0 && paintedBefore.col > 0, '[column] 전제: 열 전체 선택이 실제로 칠해졌다', JSON.stringify(paintedBefore));

    await scrollTo(page, 600 * ROW_H);
    await scrollTo(page, 0);
    const paintedAfter = await page.evaluate(() => ({
      range: document.querySelectorAll('#grid-edit-table td.g-range-selected').length,
      col: document.querySelectorAll('#grid-edit-table td.g-col-selected').length,
    }));
    console.log(`  스크롤 왕복 뒤: 범위표시 ${paintedAfter.range}칸 / 열표시 ${paintedAfter.col}칸`);
    check(
      paintedAfter.range === paintedBefore.range,
      '[column] 스크롤 왕복 뒤에도 열 선택 범위가 그대로 보인다',
      `${paintedAfter.range} vs ${paintedBefore.range}`,
    );
    check(
      paintedAfter.col === paintedBefore.col,
      '[column] 스크롤 왕복 뒤에도 열 머리글 선택 표시(g-col-selected)가 남아 있다',
      `${paintedAfter.col} vs ${paintedBefore.col}`,
    );
  }

  // ---------------------------------------------------------------- perf
  if (wants('perf')) {
    console.log('\n=== [perf] 스크롤 중 longtask (81be4d8의 개선이 유지되는지) ===');
    // 측정 구간을 서버 구독 페이지 하나(0~499행) 안으로 묶는다 — 시트 전체를 훑어 내리면 스크롤이
    // scheduleSubscriptionUpdate로 새 페이지를 요청하고, 그때 서버가 밀어 넣는 Yjs 갱신이
    // structure/cells observer를 통해 **전체 재빌드**를 부른다. 그건 렌더 경로가 아니라 서버 왕복
    // 비용이라, 여기서 재려는 "스크롤 프레임당 렌더 비용"을 통째로 가려버린다(실측: 같은 코드가
    // 5열 시트에서 0건, 25열 시트 전체 훑기에서 30건으로 나왔다 — 차이는 전부 이 왕복이었다).
    const PERF_STEPS = 40;
    const PERF_SPAN = 9500; // ≈ 행 0~350 — 첫 구독 페이지(500행) 안
    // 워밍업: 측정 구간을 한 번 왕복해 그 구간의 셀 페이지를 미리 받아 둔다.
    await scrollTo(page, PERF_SPAN);
    await scrollTo(page, 0);
    await page.waitForTimeout(1200);
    await page.evaluate(() => {
      window.__lt = [];
      window.__ltObs = new PerformanceObserver((list) => {
        for (const e of list.getEntries()) window.__lt.push(e.duration);
      });
      window.__ltObs.observe({ entryTypes: ['longtask'] });
      // 81be4d8이 산 것의 본질은 "스크롤만으로는 전체 재빌드를 하지 않는다"이다 — longtask ms는
      // 기계·시트 폭에 따라 크게 흔들리지만 이 불변식은 그렇지 않아, 회귀 판정의 기준으로 더 낫다.
      // 전체 재빌드는 tableEl.replaceChildren()으로 <colgroup>을 새로 넣는 것으로 알아볼 수 있다.
      const table = document.getElementById('grid-edit-table');
      window.__rebuilds = 0;
      window.__mo = new MutationObserver((recs) => {
        for (const r of recs) {
          if (r.target === table && [...r.addedNodes].some((n) => n.tagName === 'COLGROUP')) window.__rebuilds += 1;
        }
      });
      window.__mo.observe(table, { childList: true });
    });
    // 같은 구간을 톱니 모양으로 오르내린다 — 한 스텝이 500px(≈18행)라 매 프레임 창 가장자리에서
    // 실제로 행이 들고 난다(증분 패치가 일해야 하는 조건). 매번 rAF 재렌더가 돌도록 한 프레임씩 양보한다.
    await page.evaluate(
      async ({ steps, span }) => {
        const host = document.querySelector('.grid-editor-scroll');
        const stride = 500;
        const legs = Math.round(span / stride);
        for (let i = 0; i < steps; i += 1) {
          const leg = i % (legs * 2);
          host.scrollTop = (leg < legs ? leg : legs * 2 - leg) * stride;
          await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
        }
      },
      { steps: PERF_STEPS, span: PERF_SPAN },
    );
    await page.waitForTimeout(600);
    const lt = await page.evaluate(() => {
      window.__ltObs.disconnect();
      window.__mo.disconnect();
      const total = window.__lt.reduce((a, b) => a + b, 0);
      return {
        count: window.__lt.length,
        total: Math.round(total),
        max: Math.round(Math.max(0, ...window.__lt)),
        rebuilds: window.__rebuilds,
      };
    });
    console.log(`  longtasks: count=${lt.count} totalMs=${lt.total} maxMs=${lt.max} / 전체 재빌드 ${lt.rebuilds}회`);
    notes.push(`perf: ${lt.count} longtasks / ${lt.total}ms total / ${lt.max}ms max / ${lt.rebuilds} full rebuilds`);
    // 핵심 회귀 판정 — 스크롤"만"으로는 전체 재빌드가 단 한 번도 일어나면 안 된다(81be4d8 이전
    // 코드로 재면 40프레임에 40회 가까이 나온다).
    check(lt.rebuilds === 0, '[perf] 스크롤만으로는 전체 재빌드가 일어나지 않는다 (81be4d8의 불변식)', `${lt.rebuilds}회`);
    // 절대 시간은 기계·시트 폭을 많이 타므로 느슨한 안전망으로만 쓴다(이 25열 픽스처의 실측:
    // 81be4d8 이전 3553ms / 81be4d8 이후 약 2000ms).
    check(lt.total <= 3000, '[perf] 스크롤 longtask 총합이 안전망 안에 있다', `${lt.total}ms`);
  }

  await browser.close();
}

await main();
console.log('');
for (const n of notes) console.log(`NOTE ${n}`);
if (failures.length > 0) {
  console.log(`\nFAILED (${failures.length}):`);
  for (const f of failures) console.log(' -', f);
  process.exit(1);
}
console.log('\nALL CHECKS PASSED');
