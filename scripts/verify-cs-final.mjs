// 열 위치(A열부터), 틀 고정(스크롤 시 제목/헤더/업체명 열 고정), 기본 격자선 색을 확인한다.
import { chromium } from 'playwright';

const embedUrl = process.argv[2];
if (!embedUrl) {
  console.error('usage: node verify-cs-final.mjs <full-editor-url>');
  process.exit(1);
}

const browser = await chromium.launch();
const page = await browser.newPage();
page.on('pageerror', (err) => console.log('PAGEERROR:', err.message));

await page.goto(embedUrl, { waitUntil: 'domcontentloaded' });
await page.waitForSelector('#grid-edit-table td[data-row]', { timeout: 20000 });

// 1) 열 위치: 업체명이 이제 col=0(A열)에 있는지
const colInfo = await page.evaluate(() => {
  const nameCell = document.querySelector('#grid-edit-table td[data-row="2"][data-col="0"]');
  const titleCell = document.querySelector('#grid-edit-table td[data-row="0"][data-col="0"]');
  return { nameText: nameCell?.textContent, titleText: titleCell?.textContent, titleColSpan: titleCell?.colSpan };
});
console.log('col0 info:', JSON.stringify(colInfo));
console.log(
  colInfo.nameText === '(주)전주페이퍼' && colInfo.titleColSpan === 14
    ? 'OK: 업체명이 A열(col=0)부터 시작, 제목 병합도 14칸(A~N)으로 정상'
    : 'FAIL: 열 위치 문제',
);

// 2) 기본 격자선 색상 확인 (연회색 실선)
const borderColor = await page.evaluate(() => {
  const td = document.querySelector('#grid-edit-table td[data-row="10"][data-col="5"]');
  return td ? getComputedStyle(td).borderTopColor : null;
});
console.log('border color:', borderColor);
console.log(
  borderColor === 'rgb(212, 212, 212)' ? 'OK: 기본 격자선이 연회색 실선(#d4d4d4)' : `FAIL: 격자선 색 확인 필요 (${borderColor})`,
);

// 3) 틀 고정: 스크롤 후에도 제목/헤더/업체명 열이 화면에 남아있는지
const scrollBox = page.locator('.grid-editor-scroll');
await scrollBox.evaluate((el) => { el.scrollTop = 2000; el.scrollLeft = 500; });
await page.waitForTimeout(300);

const afterScroll = await page.evaluate(() => {
  const titleCell = document.querySelector('#grid-edit-table td[data-row="0"][data-col="0"]');
  const nameCell = document.querySelector('#grid-edit-table td[data-row="2"][data-col="0"]');
  const box = document.querySelector('.grid-editor-scroll').getBoundingClientRect();
  const titleRect = titleCell?.getBoundingClientRect();
  const nameRect = nameCell?.getBoundingClientRect();
  return {
    titleVisible: titleRect && titleRect.top >= box.top && titleRect.top < box.bottom,
    nameVisible: nameRect && nameRect.left >= box.left && nameRect.left < box.right,
    titlePosition: titleCell ? getComputedStyle(titleCell).position : null,
    namePosition: nameCell ? getComputedStyle(nameCell).position : null,
  };
});
console.log('after scroll:', JSON.stringify(afterScroll));
console.log(
  afterScroll.titlePosition === 'sticky' && afterScroll.namePosition === 'sticky' && afterScroll.titleVisible && afterScroll.nameVisible
    ? 'OK: 스크롤 후에도 제목 행/업체명 열이 sticky로 화면에 남아있음'
    : 'FAIL: 틀 고정이 정상 동작하지 않음',
);

await browser.close();
