// 편집기를 한 번 열었다 닫아서(유휴 자동저장/마지막 클라이언트 퇴장 flush를 트리거) 잠금 상태가
// 그 이후에도 살아남는지 확인한다 — sheetProtection이 ExcelJS 재저장 시 사라지는 버그를 고쳤는지 검증.
import { chromium } from 'playwright';

const embedUrl = process.argv[2];
const browser = await chromium.launch();

// 1) 한 번 열었다 닫는다(disconnect 유발)
const page1 = await browser.newPage();
await page1.goto(embedUrl, { waitUntil: 'domcontentloaded' });
await page1.waitForSelector('#grid-edit-table td[data-row]', { timeout: 20000 });
console.log('편집기 1회 오픈 확인, 닫는 중...');
await page1.close();

// 2) room.ts의 유휴 자동저장(8초)을 넉넉히 기다린다
console.log('자동저장 대기 중(10초)...');
await new Promise((r) => setTimeout(r, 10000));

// 3) 다시 열어서 잠금 상태 확인
const page2 = await browser.newPage();
await page2.goto(embedUrl, { waitUntil: 'domcontentloaded' });
await page2.waitForSelector('#grid-edit-table td[data-row]', { timeout: 20000 });

const lockInfo = await page2.evaluate(() => {
  const get = (r, c) => document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="${c}"]`);
  return {
    header: get(1, 0)?.getAttribute('contenteditable'),
    name: get(2, 0)?.getAttribute('contenteditable'),
    month: get(2, 1)?.getAttribute('contenteditable'),
    total: get(2, 13)?.getAttribute('contenteditable'),
  };
});
console.log('재접속 후 lock:', JSON.stringify(lockInfo));
console.log(
  lockInfo.header === 'false' && lockInfo.name === 'false' && lockInfo.month === 'true' && lockInfo.total === 'false'
    ? 'OK: 한 번 열었다 닫은 뒤(자동저장 거친 뒤)에도 잠금 상태 정상 유지됨'
    : 'FAIL: 잠금 상태가 저장 후 풀림 (버그 재발)',
);

await browser.close();
