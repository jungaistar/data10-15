// 헤드리스 브라우저로 주요 흐름 점검 (선택 — playwright-core 가 있는 PC 에서만)
//   PLAYWRIGHT_CORE=/경로/node_modules/playwright-core node scripts/check-browser.mjs
// file:// 로 index.html 을 열어 샘플 → 8개 화면 × 두 폭 가로 넘침 → 엑셀 여러 개(zip)·한 개 내려받기와 내용 → 인쇄 화면 두 모양
// → 회사 양식 올리기 → 실제 파일 가져오기 순서로 누르고, JS 오류를 잽니다. BREAK=1 이면 일부러 넓은 상자를 넣어 넘침 검사가 잡는지 봅니다.
import { createRequire } from 'node:module';
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { chromium } = require(process.env.PLAYWRIGHT_CORE || 'playwright-core');
const XLSX = require(path.join(root, 'vendor/xlsx.full.min.js'));
const url = 'file://' + path.join(root, 'index.html');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'data10-15-check-'));

const errors = [], problems = [];
const browser = await chromium.launch();
const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1280, height: 900 } });
const page = await ctx.newPage();
page.on('pageerror', e => errors.push('pageerror: ' + e.message));
page.on('console', m => { if (m.type() === 'error') errors.push('console: ' + m.text()); });
page.on('dialog', d => d.accept());
function ok(cond, msg) { if (!cond) { problems.push(msg); console.log('  FAIL ' + msg); } else console.log('  ok  ' + msg); }
async function go(hash) { await page.evaluate(h => { location.hash = h; }, hash); await page.waitForTimeout(200); }
async function overflow() {
  return page.evaluate(() => {
    const W = document.documentElement.clientWidth, bad = [];
    document.querySelectorAll('main *').forEach(el => {
      const r = el.getBoundingClientRect(); if (!r.width || !r.height) return;
      let p = el.parentElement, inScroll = false;
      while (p) { const o = getComputedStyle(p).overflowX; if (o === 'auto' || o === 'scroll') { inScroll = true; break; } p = p.parentElement; }
      if (!inScroll && r.right > W + 1) bad.push(el.tagName + '.' + el.className);
    });
    return bad.slice(0, 3);
  });
}
async function saveDl(p) { const d = await p; const f = path.join(tmp, d.suggestedFilename()); await d.saveAs(f); return { name: d.suggestedFilename(), file: f }; }

await page.goto(url);
await page.evaluate(() => { localStorage.clear(); });
await page.reload();
if (process.env.BREAK) await page.addStyleTag({ content: '.page-head { min-width: 900px; }' });
ok((await page.textContent('main')).includes('데이터 가져오기'), '처음 열면 가져오기 화면');
await page.locator('button', { hasText: '샘플로 시작' }).first().click();
await page.waitForTimeout(400);
ok(await page.locator('#sampleBanner').isVisible(), '샘플 띠 보임');
const g = await page.textContent('main');
ok(/협력사\s*6곳/.test(g) && /확인 필요\s*2행/.test(g) && /주의\s*5행/.test(g), '협력사 6곳 · 확인 필요 2행 · 주의 5행');
ok(g.includes('고른 4곳 문서 만들기'), '「확인 필요」 없는 4곳만 처음에 고름');

for (const id of ['import', 'groups', 'rows', 'docs', 'print', 'tpl', 'history', 'settings']) {
  for (const w of [1280, 390]) {
    await page.setViewportSize({ width: w, height: 900 });
    await go('#/' + id);
    const bad = await overflow();
    ok(!bad.length, id + ' ' + w + 'px 가로 넘침 없음' + (bad.length ? ' — ' + bad.join(', ') : ''));
  }
}
await page.setViewportSize({ width: 1280, height: 900 });

// 검증 화면: 확인 필요만
await go('#/rows');
await page.locator('select[aria-label="판정"]').selectOption('bad');
await page.waitForTimeout(200);
ok((await page.locator('table.list tbody tr').count()) === 2, '확인 필요 2행만 거르기');

// 문서 만들기: 4곳 × 2종 = 8개 → zip
await go('#/docs');
ok((await page.locator('table.list tbody tr').count()) === 8, '만들 문서 8개 미리 보기');
let dl = saveDl(page.waitForEvent('download')); await page.locator('button', { hasText: '내려받기(zip)' }).click();
dl = await dl;
const zbuf = fs.readFileSync(dl.file);
ok(dl.name.endsWith('8개.zip') && zbuf.readUInt32LE(0) === 0x04034b50, 'zip 한 파일로 8개 내려받기');
// 끝의 중앙 디렉터리에서 이름을 읽음(안에 든 xlsx 도 zip 이라 앞에서부터 서명을 찾으면 섞임)
const eocd = zbuf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
const zipNames = []; for (let i = zbuf.readUInt32LE(eocd + 16), k = 0; k < zbuf.readUInt16LE(eocd + 10); k++) { const n = zbuf.readUInt16LE(i + 28); zipNames.push(zbuf.slice(i + 46, i + 46 + n).toString('utf8')); i += 46 + n + zbuf.readUInt16LE(i + 30) + zbuf.readUInt16LE(i + 32); }
ok(zipNames.length === 8 && zipNames.some(n => n.startsWith('합의서_C기공')), 'zip 안 한글 파일 이름 8개 (' + zipNames[0] + ')');
await go('#/history');
ok((await page.locator('table.list tbody tr').count()) === 8, '발행 이력 8건');

// 하나만: C기공 합의서 → 같은 번호(재발행)
await go('#/docs');
for (const n of ['D산업', 'E금속', 'F Hydraulics']) await page.locator('label.opt', { hasText: n }).locator('input').uncheck();
await page.locator('label.opt', { hasText: '인보이스' }).locator('input').uncheck();
await page.waitForTimeout(200);
ok((await page.textContent('main')).includes('재발행(같은 번호)'), '같은 내용은 재발행 표시');
dl = saveDl(page.waitForEvent('download')); await page.locator('button', { hasText: '엑셀 내려받기' }).click(); dl = await dl;
const wb = XLSX.read(fs.readFileSync(dl.file)); const ws = wb.Sheets['합의서'];
ok(dl.name.startsWith('합의서_C기공') && dl.name.endsWith('.xlsx'), '엑셀 한 개 내려받기 (' + dl.name + ')');
ok(ws.B14.v === 'CL-2609-005' && ws.I17.v === 100000 && ws.C9.v === '금 이백일십이만팔천원정(₩2,128,000)', '내려받은 합의서 — 클레임·비율 계산 금액·한글 금액');
ok(/^AG-\d{4}-\d{2}-001$/.test(ws.C3.v) && dl.name.includes(ws.C3.v), '문서 번호 ' + ws.C3.v);
await go('#/history');
ok((await page.locator('table.list tbody tr').count()) === 8, '재발행은 이력이 늘지 않음');

// 인쇄: 4곳 × 2종 다시 고르고 두 모양
await go('#/docs');
for (const n of ['A정밀', 'B유압', 'D산업', 'E금속', 'F Hydraulics']) await page.locator('label.opt', { hasText: n }).locator('input').check();
await page.locator('label.opt', { hasText: '인보이스' }).locator('input').check();
await go('#/print');
ok((await page.locator('article.doc').count()) === 12, '인쇄용 기본 양식 12쪽(6곳 × 2)');
const pt = await page.textContent('main');
ok(pt.includes('USD 14,138.38') && pt.includes('금 일천칠백육십일만원정(₩17,610,000)'), '달러·원화 금액 표기');
await page.evaluate(() => { window.__printed = 0; window.print = () => { window.__printed++; }; });
await page.locator('button', { hasText: '번호 확정 · 인쇄' }).click(); await page.waitForTimeout(500);
ok(await page.evaluate(() => window.__printed) === 1, '인쇄 창 호출');
await go('#/history');
ok((await page.locator('table.list tbody tr').count()) === 12, '인쇄로 새로 고른 A·B 4건 확정(이력 12건)');
await go('#/print');
await page.locator('button', { hasText: '엑셀 양식 모양' }).click(); await page.waitForTimeout(400);
ok((await page.locator('.xl-sheet').count()) === 12, '엑셀 양식 모양 미리 보기 12장');

// 양식: 회사 양식(공유 문자열표) 올리기, 키 없는 엑셀은 거절
await go('#/tpl');
await page.locator('input[type=file][accept=".xlsx,.xlsm"]').nth(0).setInputFiles(path.join(root, 'test/fixtures/합의서_공유문자열.xlsx'));
await page.waitForTimeout(500);
let tt = await page.textContent('main');
ok(tt.includes('올린 양식') && tt.includes('합의서 14행(9칸)'), '회사 양식 올리기 — 반복 줄 14행');
await page.locator('input[type=file][accept=".xlsx,.xlsm"]').nth(1).setInputFiles(path.join(root, 'samples/클레임_합의결과_샘플.xlsx'));
await page.waitForTimeout(400);
ok((await page.textContent('#toast')).includes('하나도 찾지 못했습니다'), '키 없는 엑셀은 인보이스 양식으로 받지 않음');

// 가져오기: 실제 파일로 바꾸면 샘플 띠가 사라짐
await go('#/import');
await page.setInputFiles('input[type=file][accept=".xlsx,.xlsm,.xls,.csv"]', path.join(root, 'samples/클레임_합의결과_샘플.xlsx'));
await page.waitForTimeout(500);
await page.locator('button', { hasText: '개 파일로 바꾸기' }).click();
await page.waitForTimeout(400);
ok(!(await page.locator('#sampleBanner').isVisible()), '실제 파일 반영 뒤 샘플 띠 사라짐');
ok((await page.textContent('#dataChip')).includes('22행'), '가져온 합의 결과 22행');

// 설정: 부가세 10% 별도 → 인쇄 화면 C기공 합계에 부가세
await go('#/settings');
await page.locator('select', { has: page.locator('option[value="add10"]') }).selectOption('add10');
await page.waitForTimeout(200);
await go('#/groups');
ok((await page.textContent('main')).includes('2,340,800'), '부가세 10% 별도면 C기공 총액 2,340,800');

await browser.close();
fs.rmSync(tmp, { recursive: true, force: true });
if (errors.length) { console.log('JS 오류:\n' + errors.join('\n')); }
console.log(errors.length || problems.length ? '\n문제 ' + (errors.length + problems.length) + '건' : '\n모두 통과');
process.exitCode = errors.length || problems.length ? 1 : 0;
