// 두 브라우저 컨텍스트(=두 사용자)로 동시 접속해 (1) 아바타 프레즌스가 서로에게 보이는지,
// (2) 상대가 지금 어디를 편집 중인지(docx 원격 커서 이름표 / xlsx 셀 하이라이트)가 실시간으로
// 반영되는지, (3) 툴바 버튼의 눌린(aria-pressed) 상태가 실제 서식을 따라가는지를 검증한다.
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { chromium } from 'playwright';
import { strToU8, zipSync } from 'fflate';
import ExcelJS from 'exceljs';
import jwt from 'jsonwebtoken';

const BASE = `http://127.0.0.1:${process.env.DOCVIEW_PORT}`;
const DATA_DIR = process.env.DOCVIEW_DATA_DIR;
const SECRET = process.env.DOCVIEW_SECRET;
const SHOT_DIR = path.resolve('scripts/.screenshots');

function makeSampleDocx() {
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>';
  const rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:body><w:p><w:r><w:t xml:space="preserve">동시 접속 확인용 문단입니다.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>';
  return Buffer.from(
    zipSync({
      '[Content_Types].xml': strToU8(contentTypes),
      '_rels/.rels': strToU8(rootRels),
      'word/document.xml': strToU8(documentXml),
    }),
  );
}

async function makeSampleXlsx() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = '이름';
  sheet.getCell('B1').value = '점수';
  sheet.getCell('A2').value = '동시접속확인';
  sheet.getCell('B2').value = 10;
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

async function seedDoc(ext, bytes) {
  const id = randomUUID();
  await fs.mkdir(path.join(DATA_DIR, 'originals'), { recursive: true });
  await fs.mkdir(path.join(DATA_DIR, 'meta'), { recursive: true });
  await fs.writeFile(path.join(DATA_DIR, 'originals', `${id}.${ext}`), bytes);
  await fs.writeFile(
    path.join(DATA_DIR, 'meta', `${id}.json`),
    JSON.stringify({
      id,
      name: `presence-check.${ext}`,
      ext,
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
      viewKind: ext === 'xlsx' ? 'grid' : 'pages',
      progress: 100,
    }),
  );
  return id;
}

function tokenFor(docId, viewer) {
  return jwt.sign(
    { docId, viewer, perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: true } },
    SECRET,
    { expiresIn: 300 },
  );
}

async function main() {
  await fs.mkdir(SHOT_DIR, { recursive: true });
  const browser = await chromium.launch();
  const ctxA = await browser.newContext();
  const ctxB = await browser.newContext();
  const pageA = await ctxA.newPage();
  const pageB = await ctxB.newPage();
  const errors = [];
  for (const p of [pageA, pageB]) {
    p.on('console', (msg) => {
      if (msg.type() === 'error') errors.push(msg.text());
    });
    p.on('pageerror', (err) => errors.push(`pageerror: ${err.message}`));
  }

  console.log('--- docx: 두 사용자 동시 접속 ---');
  const docxId = await seedDoc('docx', makeSampleDocx());
  const tokenA = tokenFor(docxId, '김철수');
  const tokenB = tokenFor(docxId, '이영희');
  await pageA.goto(`${BASE}/editor.html?token=${encodeURIComponent(tokenA)}`, { waitUntil: 'domcontentloaded' });
  await pageB.goto(`${BASE}/editor.html?token=${encodeURIComponent(tokenB)}`, { waitUntil: 'domcontentloaded' });
  await pageA.waitForSelector('.ProseMirror', { timeout: 20000 });
  await pageB.waitForSelector('.ProseMirror', { timeout: 20000 });
  // 서로의 아바타가 보이는지 (본인 포함 2명)
  await pageA.waitForFunction(() => document.querySelectorAll('#presence-list .presence-badge').length === 2, {
    timeout: 10000,
  });
  await pageB.waitForFunction(() => document.querySelectorAll('#presence-list .presence-badge').length === 2, {
    timeout: 10000,
  });
  const avatarTitlesOnA = await pageA.evaluate(() =>
    [...document.querySelectorAll('#presence-list .presence-badge')].map((el) => el.title),
  );
  if (!avatarTitlesOnA.includes('이영희')) {
    throw new Error(`A 화면에 B(이영희) 아바타가 보이지 않습니다: ${JSON.stringify(avatarTitlesOnA)}`);
  }
  console.log('OK: 아바타 프레즌스 상호 확인 —', avatarTitlesOnA);

  // B가 타이핑하면 A 화면에 y-prosemirror 원격 커서 이름표가 뜨는지
  await pageB.click('.ProseMirror');
  await pageB.keyboard.press('End');
  await pageB.keyboard.type(' [이영희가 입력]');
  await pageA.waitForFunction(
    () => document.querySelector('.ProseMirror')?.textContent?.includes('[이영희가 입력]'),
    { timeout: 10000 },
  );
  await pageA.waitForFunction(() => !!document.querySelector('.ProseMirror-yjs-cursor'), { timeout: 10000 });
  await pageA.screenshot({ path: path.join(SHOT_DIR, '11-docx-remote-cursor.png') });
  console.log('OK: B의 원격 커서/이름표가 A 화면에 표시됨, screenshot 11-docx-remote-cursor.png');

  // 활성 상태(aria-pressed) 확인: 전체 선택 후 굵게 -> 버튼이 눌린 상태로 표시되는지
  // (B의 원격 커서 데코레이션이 막 갱신된 직후라 잠깐의 안정화 시간을 준다)
  await pageA.waitForTimeout(300);
  await pageA.click('.ProseMirror');
  await pageA.keyboard.press('Control+A');
  await pageA.click('#btn-bold');
  await pageA.waitForFunction(() => document.getElementById('btn-bold')?.getAttribute('aria-pressed') === 'true', {
    timeout: 5000,
  });
  await pageA.click('#btn-align-center');
  await pageA.waitForFunction(
    () => document.getElementById('btn-align-center')?.getAttribute('aria-pressed') === 'true',
    { timeout: 5000 },
  );
  const alignLeftPressed = await pageA.getAttribute('#btn-align-left', 'aria-pressed');
  if (alignLeftPressed === 'true') throw new Error('가운데 정렬 후에도 왼쪽 정렬 버튼이 눌린 상태로 남아있습니다.');
  console.log('OK: 굵게/가운데 정렬 버튼의 active-state(aria-pressed)가 선택 영역을 따라감');

  console.log('--- xlsx: 두 사용자 동시 접속 ---');
  const xlsxId = await seedDoc('xlsx', await makeSampleXlsx());
  const gTokenA = tokenFor(xlsxId, '김철수');
  const gTokenB = tokenFor(xlsxId, '이영희');
  await pageA.goto(`${BASE}/editor-grid.html?token=${encodeURIComponent(gTokenA)}`, { waitUntil: 'domcontentloaded' });
  await pageB.goto(`${BASE}/editor-grid.html?token=${encodeURIComponent(gTokenB)}`, { waitUntil: 'domcontentloaded' });
  await pageA.waitForSelector('#grid-edit-table td[contenteditable]', { timeout: 20000 });
  await pageB.waitForSelector('#grid-edit-table td[contenteditable]', { timeout: 20000 });
  await pageA.waitForFunction(() => document.querySelectorAll('#presence-list .presence-badge').length === 2, {
    timeout: 10000,
  });

  // B가 특정 셀을 포커스하면 A 화면에 "B가 여기 편집 중" 표시가 뜨는지
  await pageB.click('#grid-edit-table td[data-row="1"][data-col="1"]');
  await pageA.waitForFunction(
    () => {
      const td = document.querySelector('#grid-edit-table td[data-row="1"][data-col="1"]');
      return td?.classList.contains('g-remote-cell') && td.querySelector('.g-remote-tag')?.textContent === '이영희';
    },
    { timeout: 10000 },
  );
  await pageA.screenshot({ path: path.join(SHOT_DIR, '12-grid-remote-cell.png') });
  console.log('OK: B가 편집 중인 셀이 A 화면에 색 테두리+이름표로 표시됨, screenshot 12-grid-remote-cell.png');

  // B가 다른 셀로 옮기면 이전 셀의 표시가 사라지는지
  await pageB.click('#grid-edit-table td[data-row="0"][data-col="0"]');
  await pageA.waitForFunction(
    () => !document.querySelector('#grid-edit-table td[data-row="1"][data-col="1"]')?.classList.contains('g-remote-cell'),
    { timeout: 10000 },
  );
  await pageA.waitForFunction(
    () => document.querySelector('#grid-edit-table td[data-row="0"][data-col="0"]')?.classList.contains('g-remote-cell'),
    { timeout: 10000 },
  );
  console.log('OK: 셀 이동 시 원격 편집 표시가 따라 이동함');

  // 셀 서식 active-state: A가 자기 셀을 골라 굵게 누르면 버튼이 눌린 상태로 표시되는지
  await pageA.click('#grid-edit-table td[data-row="0"][data-col="1"]');
  await pageA.click('#btn-cell-bold');
  await pageA.waitForFunction(
    () => document.getElementById('btn-cell-bold')?.getAttribute('aria-pressed') === 'true',
    { timeout: 5000 },
  );
  console.log('OK: 엑셀 셀 서식 버튼의 active-state가 선택 셀을 따라감');

  console.log('--- console errors ---');
  if (errors.length === 0) console.log('없음');
  else for (const e of errors) console.log('ERROR:', e);

  await browser.close();
  console.log('\nALL PRESENCE CHECKS PASSED');
  if (errors.length > 0) process.exit(1);
}

main().catch((error) => {
  console.error('PRESENCE CHECK FAILED:', error);
  process.exit(1);
});
