// 헤드리스 브라우저로 실제 UI를 구동해 확인한다: 업로드→뷰어, docx 동시편집기(editor.html),
// xlsx 동시편집기(editor-grid.html). 이 환경에는 LibreOffice가 없어 docx는 서식 없는 텍스트
// 미리보기 폴백 컨버터(src/converters/docxTextPreview.ts)로 변환된다 — 그래서 이제 실제 업로드
// API로도 docx/xlsx 둘 다 끝까지(변환 → 편집기 진입 → 저장) 검증할 수 있다(단계 7, 8).
// 단계 3~6은 여전히 원본 파일을 디스크에 직접 시딩 + canEdit 토큰을 직접 서명해 편집기 페이지에
// 곧장 진입하는 방식을 쓴다(업로드 단계를 건너뛰고 협업 룸 로직만 빠르게 반복 확인하기 위함).
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { chromium } from 'playwright';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { strToU8, zipSync } from 'fflate';
import jwt from 'jsonwebtoken';
import ExcelJS from 'exceljs';

const BASE = `http://127.0.0.1:${process.env.DOCVIEW_PORT}`;
const DATA_DIR = process.env.DOCVIEW_DATA_DIR;
const SECRET = process.env.DOCVIEW_SECRET;
const SHOT_DIR = path.resolve('scripts/.screenshots');

async function makeSamplePdf() {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.HelveticaBold);
  const page = doc.addPage([595, 842]);
  page.drawText('DocView Browser Check', { x: 60, y: 760, size: 24, font });
  return Buffer.from(await doc.save());
}

function makeSampleDocx(paragraphs) {
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
  const body = paragraphs.map((t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`).join('');
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}<w:sectPr/></w:body></w:document>`;
  return Buffer.from(
    zipSync({
      '[Content_Types].xml': strToU8(contentTypes),
      '_rels/.rels': strToU8(rootRels),
      'word/document.xml': strToU8(documentXml),
    }),
  );
}

async function seedEditableDoc() {
  const id = randomUUID();
  await fs.mkdir(path.join(DATA_DIR, 'originals'), { recursive: true });
  await fs.mkdir(path.join(DATA_DIR, 'meta'), { recursive: true });
  await fs.writeFile(path.join(DATA_DIR, 'originals', `${id}.docx`), makeSampleDocx(['브라우저 확인용 문단입니다.']));
  await fs.writeFile(
    path.join(DATA_DIR, 'meta', `${id}.json`),
    JSON.stringify({
      id,
      name: 'browser-check.docx',
      ext: 'docx',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
      viewKind: 'pages',
      progress: 100,
    }),
  );
  const token = jwt.sign(
    { docId: id, viewer: '브라우저확인', perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: true } },
    SECRET,
    { expiresIn: 300 },
  );
  return token;
}

async function makeSampleXlsx() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = '이름';
  sheet.getCell('B1').value = '점수';
  sheet.getCell('A2').value = '브라우저확인';
  sheet.getCell('B2').value = 10;
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

async function seedEditableXlsx() {
  const id = randomUUID();
  await fs.mkdir(path.join(DATA_DIR, 'originals'), { recursive: true });
  await fs.mkdir(path.join(DATA_DIR, 'meta'), { recursive: true });
  await fs.writeFile(path.join(DATA_DIR, 'originals', `${id}.xlsx`), await makeSampleXlsx());
  await fs.writeFile(
    path.join(DATA_DIR, 'meta', `${id}.json`),
    JSON.stringify({
      id,
      name: 'browser-check.xlsx',
      ext: 'xlsx',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
      viewKind: 'grid',
      progress: 100,
    }),
  );
  const token = jwt.sign(
    { docId: id, viewer: '브라우저확인', perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: true } },
    SECRET,
    { expiresIn: 300 },
  );
  return token;
}

async function main() {
  await fs.mkdir(SHOT_DIR, { recursive: true });
  const browser = await chromium.launch();
  const context = await browser.newContext();
  const page = await context.newPage();
  const consoleErrors = [];
  page.on('console', (msg) => {
    if (msg.type() === 'error') consoleErrors.push(msg.text());
  });
  page.on('pageerror', (err) => consoleErrors.push(`pageerror: ${err.message}`));

  console.log('--- 1) 문서 목록 페이지 ---');
  await page.goto(BASE, { waitUntil: 'domcontentloaded' });
  await page.waitForSelector('#docs-heading');
  await page.screenshot({ path: path.join(SHOT_DIR, '01-index.png') });
  console.log('OK: index.html 렌더링 확인, screenshot 01-index.png');

  console.log('--- 2) PDF 업로드 -> 뷰어 ---');
  const pdfPath = path.join(SHOT_DIR, 'sample.pdf');
  await fs.writeFile(pdfPath, await makeSamplePdf());
  await page.setInputFiles('#file-input', pdfPath);
  await page.waitForURL(/viewer\.html\?token=/, { timeout: 20000 });
  await page.waitForSelector('.page-frame.page-ready', { timeout: 20000 });
  await page.screenshot({ path: path.join(SHOT_DIR, '02-viewer.png') });
  console.log('OK: 업로드 -> PDF 변환 -> 뷰어 페이지 렌더링까지 확인, screenshot 02-viewer.png');

  console.log('--- 3) docx 동시편집기 ---');
  const editToken = await seedEditableDoc();
  await page.goto(`${BASE}/editor.html?token=${encodeURIComponent(editToken)}`, { waitUntil: 'domcontentloaded' });
  await page.waitForSelector('.ProseMirror', { timeout: 20000 });
  await page.waitForFunction(
    () => document.querySelector('.ProseMirror')?.textContent?.includes('브라우저 확인용 문단입니다.'),
    { timeout: 20000 },
  );
  await page.screenshot({ path: path.join(SHOT_DIR, '03-editor-loaded.png') });
  console.log('OK: 편집기 로딩 + docx 원본 텍스트 렌더링 확인, screenshot 03-editor-loaded.png');

  await page.click('.ProseMirror');
  await page.keyboard.press('End');
  await page.keyboard.type(' [브라우저에서 타이핑함]');
  await page.waitForFunction(
    () => document.querySelector('.ProseMirror')?.textContent?.includes('[브라우저에서 타이핑함]'),
    { timeout: 5000 },
  );
  await page.screenshot({ path: path.join(SHOT_DIR, '04-editor-typed.png') });
  console.log('OK: 실제 타이핑 입력 반영 확인, screenshot 04-editor-typed.png');

  const connStatus = await page.locator('#conn-status').textContent();
  console.log('연결 상태 배지 텍스트:', connStatus);

  console.log('--- 5) xlsx 동시편집기 ---');
  const gridToken = await seedEditableXlsx();
  await page.goto(`${BASE}/editor-grid.html?token=${encodeURIComponent(gridToken)}`, {
    waitUntil: 'domcontentloaded',
  });
  await page.waitForSelector('#grid-edit-table td[contenteditable]', { timeout: 20000 });
  await page.waitForFunction(
    () =>
      [...document.querySelectorAll('#grid-edit-table td[contenteditable]')].some(
        (td) => td.textContent === '브라우저확인',
      ),
    { timeout: 20000 },
  );
  await page.screenshot({ path: path.join(SHOT_DIR, '05-grid-editor-loaded.png') });
  console.log('OK: xlsx 편집기 로딩 + 원본 셀 값 렌더링 확인, screenshot 05-grid-editor-loaded.png');

  const targetCell = page.locator('#grid-edit-table td[data-row="1"][data-col="1"]');
  await targetCell.click();
  await page.keyboard.press('Control+A');
  await page.keyboard.type('42');
  await targetCell.blur();
  await page.waitForFunction(
    () => document.querySelector('#grid-edit-table td[data-row="1"][data-col="1"]')?.textContent === '42',
    { timeout: 5000 },
  );
  await page.click('#btn-add-row');
  // 이제 시트는 데이터 범위 너머로도 빈 칸을 넉넉히 미리 그려두므로(엑셀처럼) "+행" 클릭 전후로
  // 화면에 그려지는 행 수 자체는 안 바뀐다 — 대신 그 빈 칸에 실제로 입력하면 시트 범위가
  // 자동으로 넓어지는지(commitCell의 자동 확장)를 확인한다.
  const paddedCell = page.locator('#grid-edit-table td[data-row="10"][data-col="0"]');
  await paddedCell.click();
  await page.keyboard.type('패딩영역입력');
  await paddedCell.blur();
  await page.waitForFunction(
    () => document.querySelector('#grid-edit-table td[data-row="10"][data-col="0"]')?.textContent === '패딩영역입력',
    { timeout: 5000 },
  );
  await page.screenshot({ path: path.join(SHOT_DIR, '06-grid-editor-edited.png') });
  console.log('OK: 셀 편집 + 데이터 범위 밖 빈 칸 입력 시 자동 확장 확인, screenshot 06-grid-editor-edited.png');

  const gridConnStatus = await page.locator('#conn-status').textContent();
  console.log('그리드 편집기 연결 상태 배지 텍스트:', gridConnStatus);

  // 마지막 편집(행 추가)의 Yjs 업데이트가 WebSocket으로 서버에 도달할 시간을 준다 —
  // 곧바로 browser.close()하면 네트워크 왕복 전에 연결이 끊겨 서버 저장에 반영되지 않을 수 있다.
  await page.waitForTimeout(500);

  console.log('--- 7) xlsx 실제 업로드 -> UI의 "동시편집 열기" 버튼 클릭 (토큰 직접 서명 우회 없이) ---');
  await page.goto(BASE, { waitUntil: 'domcontentloaded' });
  await page.waitForSelector('#docs-heading');
  const xlsxPath = path.join(SHOT_DIR, 'sample.xlsx');
  await fs.writeFile(xlsxPath, await makeSampleXlsx());
  await page.setInputFiles('#file-input', xlsxPath);
  // 업로드가 끝나면 main.js가 자동으로 뷰어로 이동시킨다 — 목록에서 "동시편집 열기"를 누르는
  // 실제 사용 흐름은 그 뒤 문서 목록으로 돌아온 다음이다.
  await page.waitForURL(/viewer\.html\?token=/, { timeout: 20000 });
  await page.goto(BASE, { waitUntil: 'domcontentloaded' });
  await page.waitForFunction(
    () => document.querySelector('#doc-list')?.textContent?.includes('sample.xlsx'),
    { timeout: 20000 },
  );
  const uploadedItem = page.locator('.doc-item', { hasText: 'sample.xlsx' }).first();
  const quickEditBtn = uploadedItem.getByRole('button', { name: '동시편집 열기' });
  await quickEditBtn.waitFor({ state: 'visible', timeout: 10000 });
  const [gridPage] = await Promise.all([
    context.waitForEvent('page').catch(() => null),
    quickEditBtn.click(),
  ]);
  const editorPage = gridPage ?? page;
  await editorPage.waitForURL(/editor-grid\.html\?token=/, { timeout: 10000 });
  await editorPage.waitForSelector('#grid-edit-table td[contenteditable]', { timeout: 20000 });
  await editorPage.screenshot({ path: path.join(SHOT_DIR, '07-grid-editor-from-ui.png') });
  console.log('OK: 실제 업로드 -> 목록의 "동시편집 열기" 버튼으로 xlsx 편집기 진입 확인, screenshot 07-grid-editor-from-ui.png');
  // main.js는 새 탭이 아니라 location.href로 같은 탭에서 이동한다 — editorPage가 사실상 page와
  // 동일한 참조라 이후 단계(8~9)에서 page를 docx 편집기로 재사용하면 이 URL을 잃는다.
  // 그래서 URL을 저장해 뒀다가 단계 10에서 다시 열어(같은 협업 룸에 재접속) 검증한다.
  const gridEditorUrl = editorPage.url();

  console.log('--- 8) docx 실제 업로드 -> LibreOffice 없이 폴백 변환 -> UI의 "동시편집 열기" 버튼 클릭 ---');
  await page.goto(BASE, { waitUntil: 'domcontentloaded' });
  await page.waitForSelector('#docs-heading');
  const docxPath = path.join(SHOT_DIR, 'sample-real-upload.docx');
  await fs.writeFile(docxPath, makeSampleDocx(['한글 실제 업로드 확인용 문단입니다.']));
  await page.setInputFiles('#file-input', docxPath);
  // docx는 LibreOffice가 없으면 텍스트 전용 폴백 컨버터로 변환된다 — 실제 변환이 끝나
  // 뷰어로 이동할 때까지 기다린다(변환 실패 시 여기서 타임아웃/에러로 드러난다).
  await page.waitForURL(/viewer\.html\?token=/, { timeout: 20000 });
  await page.goto(BASE, { waitUntil: 'domcontentloaded' });
  await page.waitForFunction(
    () => document.querySelector('#doc-list')?.textContent?.includes('sample-real-upload.docx'),
    { timeout: 20000 },
  );
  const uploadedDocxItem = page.locator('.doc-item', { hasText: 'sample-real-upload.docx' }).first();
  const quickEditDocxBtn = uploadedDocxItem.getByRole('button', { name: '동시편집 열기' });
  await quickEditDocxBtn.waitFor({ state: 'visible', timeout: 10000 });
  const [docxEditorPage] = await Promise.all([
    context.waitForEvent('page').catch(() => null),
    quickEditDocxBtn.click(),
  ]);
  const docxEditorTab = docxEditorPage ?? page;
  await docxEditorTab.waitForURL(/editor\.html\?token=/, { timeout: 10000 });
  await docxEditorTab.waitForSelector('.ProseMirror', { timeout: 20000 });
  await docxEditorTab.waitForFunction(
    () => document.querySelector('.ProseMirror')?.textContent?.includes('한글 실제 업로드 확인용 문단입니다.'),
    { timeout: 20000 },
  );
  await docxEditorTab.screenshot({ path: path.join(SHOT_DIR, '08-docx-editor-from-ui.png') });
  console.log('OK: 실제 업로드 -> LibreOffice 없는 폴백 변환 -> "동시편집 열기"로 docx 편집기 진입 확인, screenshot 08-docx-editor-from-ui.png');

  console.log('--- 9) 워드 편집기: 글자색/정렬/확대/검색 도구 ---');
  await docxEditorTab.click('.ProseMirror');
  await docxEditorTab.keyboard.press('Control+A');
  await docxEditorTab.locator('#font-color').fill('#ff0000');
  await docxEditorTab.waitForFunction(
    () => !!document.querySelector('.ProseMirror span[style*="color"]'),
    { timeout: 5000 },
  );
  await docxEditorTab.click('#btn-align-center');
  await docxEditorTab.waitForFunction(
    () => document.querySelector('.ProseMirror p')?.getAttribute('style')?.includes('center'),
    { timeout: 5000 },
  );
  await docxEditorTab.click('#zoom-in');
  await docxEditorTab.click('#zoom-in');
  const docxZoomLevel = await docxEditorTab.locator('#zoom-level').textContent();
  if (docxZoomLevel !== '120%') throw new Error(`워드 편집기 확대 배율이 예상과 다릅니다: ${docxZoomLevel}`);
  await docxEditorTab.click('#search-btn');
  await docxEditorTab.fill('#search-input', '확인용');
  await docxEditorTab.waitForFunction(
    () => !!document.querySelector('.editor-search-hit'),
    { timeout: 5000 },
  );
  const docxSearchCount = await docxEditorTab.locator('#search-count').textContent();
  await docxEditorTab.screenshot({ path: path.join(SHOT_DIR, '09-docx-editor-tools.png') });
  console.log(`OK: 글자색/가운데정렬/확대(${docxZoomLevel})/검색(${docxSearchCount}) 도구 동작 확인, screenshot 09-docx-editor-tools.png`);

  console.log('--- 10) 엑셀 편집기: 셀 서식/수식/확대/검색 도구 ---');
  await page.goto(gridEditorUrl, { waitUntil: 'domcontentloaded' });
  await page.waitForSelector('#grid-edit-table td[contenteditable]', { timeout: 20000 });
  const boldCell = page.locator('#grid-edit-table td[data-row="0"][data-col="0"]');
  await boldCell.click();
  await page.click('#btn-cell-bold');
  await page.locator('#cell-bg-color').fill('#00ff00');
  await page.waitForFunction(
    () => {
      const td = document.querySelector('#grid-edit-table td[data-row="0"][data-col="0"]');
      return td && td.style.fontWeight === '700' && td.style.background !== '';
    },
    { timeout: 5000 },
  );
  const formulaCell = page.locator('#grid-edit-table td[data-row="1"][data-col="1"]');
  await formulaCell.click();
  await page.keyboard.press('Control+A');
  await page.keyboard.type('=1+2');
  // 편집 중에는 원본 수식이, 포커스를 벗어나면(blur) 실시간 계산 결과(3)가 보여야 한다.
  await page.waitForFunction(
    () => document.querySelector('#grid-edit-table td[data-row="1"][data-col="1"]')?.textContent === '=1+2',
    { timeout: 5000 },
  );
  await formulaCell.blur();
  await page.waitForFunction(
    () => document.querySelector('#grid-edit-table td[data-row="1"][data-col="1"]')?.textContent === '3',
    { timeout: 5000 },
  );
  // 다시 클릭하면 계산 결과가 아니라 원본 수식이 편집 가능한 상태로 다시 보여야 한다.
  await formulaCell.click();
  await page.waitForFunction(
    () => document.querySelector('#grid-edit-table td[data-row="1"][data-col="1"]')?.textContent === '=1+2',
    { timeout: 5000 },
  );
  await formulaCell.blur();
  await page.click('#zoom-in');
  const gridZoomLevel = await page.locator('#zoom-level').textContent();
  if (gridZoomLevel !== '110%') throw new Error(`엑셀 편집기 확대 배율이 예상과 다릅니다: ${gridZoomLevel}`);
  await page.click('#search-btn');
  await page.fill('#search-input', '브라우저확인');
  await page.waitForFunction(
    () => !!document.querySelector('#grid-edit-table td.g-hit-current'),
    { timeout: 5000 },
  );
  const gridSearchCount = await page.locator('#search-count').textContent();
  await page.screenshot({ path: path.join(SHOT_DIR, '10-grid-editor-tools.png') });
  console.log(`OK: 셀 굵게/배경색/수식 실시간 계산(편집 중엔 원본, 벗어나면 결과)/확대(${gridZoomLevel})/검색(${gridSearchCount}) 도구 동작 확인, screenshot 10-grid-editor-tools.png`);

  console.log('--- 11) xlsx 읽기전용 뷰어: 데이터 범위 밖에도 패딩된 격자가 보이는지 ---');
  await page.goto(BASE, { waitUntil: 'domcontentloaded' });
  await page.waitForSelector('#docs-heading');
  const viewOnlyXlsxPath = path.join(SHOT_DIR, 'view-only.xlsx');
  await fs.writeFile(viewOnlyXlsxPath, await makeSampleXlsx());
  await page.setInputFiles('#file-input', viewOnlyXlsxPath);
  await page.waitForURL(/viewer\.html\?token=/, { timeout: 20000 });
  await page.waitForSelector('#grid-table td', { timeout: 20000 });
  await page.waitForTimeout(400); // 가상 스크롤 뷰포트가 청크를 받아 렌더링할 시간
  const paddedHeadCount = await page.evaluate(() => document.querySelectorAll('.g-head-cell').length - 1);
  const paddedRowCount = await page.evaluate(() => document.querySelectorAll('#grid-table .g-rownum').length);
  if (paddedHeadCount < 14 || paddedRowCount < 20) {
    throw new Error(`뷰어 그리드가 패딩 없이 데이터 범위(2x2)만 그려지고 있습니다: cols=${paddedHeadCount} rows=${paddedRowCount}`);
  }
  await page.screenshot({ path: path.join(SHOT_DIR, '11-viewer-grid-padded.png') });
  console.log(`OK: 읽기전용 뷰어도 데이터 범위(2x2) 밖까지 패딩된 격자(열 ${paddedHeadCount}개, 행 ${paddedRowCount}개 이상) 렌더링 확인, screenshot 11-viewer-grid-padded.png`);

  console.log('--- console errors ---');
  if (consoleErrors.length === 0) {
    console.log('없음');
  } else {
    for (const e of consoleErrors) console.log('ERROR:', e);
  }

  await browser.close();
  console.log('\nALL CHECKS PASSED');
  if (consoleErrors.length > 0) process.exit(1);
}

main().catch((error) => {
  console.error('BROWSER CHECK FAILED:', error);
  process.exit(1);
});
