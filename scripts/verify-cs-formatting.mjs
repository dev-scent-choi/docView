// 업체별 접속현황 그리드 에디터의 병합/잠금/테두리 렌더링을 실제 브라우저에서 확인한다.
import { chromium } from 'playwright';

const embedUrl = process.argv[2];
if (!embedUrl) {
  console.error('usage: node verify-cs-formatting.mjs <full-editor-url>');
  process.exit(1);
}

const browser = await chromium.launch();
const page = await browser.newPage();
page.on('pageerror', (err) => console.log('PAGEERROR:', err.message));

await page.goto(embedUrl, { waitUntil: 'domcontentloaded' });
await page.waitForSelector('#grid-edit-table td[data-row]', { timeout: 20000 });

// 1) 제목(row0,col1)이 진짜 병합(colspan)인지, 옆칸(row0,col2)이 아예 안 그려지는지
const titleInfo = await page.evaluate(() => {
  const td = document.querySelector('#grid-edit-table td[data-row="0"][data-col="1"]');
  const neighbor = document.querySelector('#grid-edit-table td[data-row="0"][data-col="2"]');
  return { text: td?.textContent, colSpan: td?.colSpan, neighborExists: !!neighbor };
});
console.log('title:', JSON.stringify(titleInfo));
console.log(titleInfo.colSpan === 14 && !titleInfo.neighborExists ? 'OK: 제목 병합 정상(14칸 colspan, 옆칸 미생성)' : 'FAIL: 제목 병합 문제');

// 2) 헤더(row1,col1="유지보수 업체명")와 이름칸(row2,col1)이 잠겨있는지(contenteditable=false)
const lockInfo = await page.evaluate(() => {
  const header = document.querySelector('#grid-edit-table td[data-row="1"][data-col="1"]');
  const nameCell = document.querySelector('#grid-edit-table td[data-row="2"][data-col="1"]');
  const monthCell = document.querySelector('#grid-edit-table td[data-row="2"][data-col="2"]');
  const totalCell = document.querySelector('#grid-edit-table td[data-row="2"][data-col="14"]');
  return {
    headerText: header?.textContent,
    headerEditable: header?.getAttribute('contenteditable'),
    nameText: nameCell?.textContent,
    nameEditable: nameCell?.getAttribute('contenteditable'),
    monthEditable: monthCell?.getAttribute('contenteditable'),
    totalText: totalCell?.textContent,
    totalEditable: totalCell?.getAttribute('contenteditable'),
  };
});
console.log('lock:', JSON.stringify(lockInfo));
const lockOk =
  lockInfo.headerEditable === 'false' &&
  lockInfo.nameEditable === 'false' &&
  lockInfo.monthEditable === 'true' &&
  lockInfo.totalEditable === 'false';
console.log(lockOk ? 'OK: 잠금 상태 정상(헤더/이름/합계 잠김, 월 값만 편집 가능)' : 'FAIL: 잠금 상태 문제');

// 3) 테두리 — 헤더 행 top은 굵게(2px), 데이터 행 안쪽 top은 얇게(1px)
const borderInfo = await page.evaluate(() => {
  const header = document.querySelector('#grid-edit-table td[data-row="1"][data-col="2"]');
  const dataRow2 = document.querySelector('#grid-edit-table td[data-row="3"][data-col="2"]'); // 2번째 데이터 행(안쪽)
  return {
    headerTop: header ? getComputedStyle(header).borderTopWidth : null,
    dataInnerTop: dataRow2 ? getComputedStyle(dataRow2).borderTopWidth : null,
  };
});
console.log('border:', JSON.stringify(borderInfo));
console.log(
  borderInfo.headerTop === '2px' && borderInfo.dataInnerTop === '1px'
    ? 'OK: 테두리 굵기 정상(헤더 상단 굵게, 안쪽 데이터 행 얇게)'
    : 'FAIL: 테두리 굵기 문제',
);

// 4) 실제로 월 값 셀을 타이핑해서 편집되는지, 이름 칸을 클릭해도 편집 모드로 안 들어가는지
const monthCell = page.locator('#grid-edit-table td[data-row="2"][data-col="2"]');
await monthCell.click();
await page.keyboard.press('Control+A');
await page.keyboard.type('9');
await monthCell.blur();
await page.waitForFunction(
  () => document.querySelector('#grid-edit-table td[data-row="2"][data-col="2"]')?.textContent === '9',
  { timeout: 5000 },
);
console.log('OK: 월 값 셀 편집 확인 (9 입력됨)');

const nameClickResult = await page.evaluate(() => {
  const td = document.querySelector('#grid-edit-table td[data-row="2"][data-col="1"]');
  td?.focus();
  return document.activeElement === td;
});
console.log(nameClickResult ? 'FAIL: 잠긴 이름 칸이 포커스를 받음(편집 가능해짐)' : 'OK: 잠긴 이름 칸은 포커스를 받지 않음');

await browser.close();
