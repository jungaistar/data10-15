// 실행: node test/logic.test.mjs   (의존성 없음 — vendor 의 SheetJS 만 씀)
// 실제 합의 결과 엑셀·회사 양식은 아직 받지 않았습니다(기획서 10장 ①②). 제출 글의 항목대로 만든 가상 샘플과 가상 기본 양식,
// 그리고 기본 양식을 LibreOffice 로 다시 저장한 사본(공유 문자열표를 쓰는 「엑셀다운」 파일)으로 검사합니다.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const L = require('../js/logic.js');
const T = require('../js/xlsx-template.js');
const B = require('../js/builtin-templates.js');
const Sample = require('../js/sample-data.js');
const XLSX = require('../vendor/xlsx.full.min.js');

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + e.message); process.exitCode = 1; }
}
const S = L.cleanSettings({});
const AOA = Sample.build().aoa;
const parsed = L.parseSheet(AOA);
const GS = L.groupBySupplier(parsed.rows, S);
const G = name => GS.find(g => g.supplier === name);
const tpl = kind => T.b64ToU8(B[kind].b64);
function readBack(u8) { return XLSX.read(u8, { type: 'array', cellFormula: true }); }
function sheetXml(u8, path = 'xl/worksheets/sheet1.xml') { return T.openZip(u8).get(path); }
// 큰 묶음: 원화 18행을 한 협력사로 모아 반복 줄이 양식의 빈 줄(8줄)을 넘치게
const BIG = L.groupBySupplier(parsed.rows.filter(r => r.currency === 'KRW').map(r => ({ ...r, supplier: '모두(가상)', supplierCode: 'V-9999' })), S)[0];

console.log('글자 · 숫자 · 날짜');
test('날짜 — ISO·점·한글·두 자리 연도·붙여 쓴 날짜·엑셀 일련번호·Date', () => {
  ['2026-09-03', '2026.9.3', '2026년 9월 3일', '26-09-03', '20260903', 46268].forEach(v => assert.equal(L.toISO(v), '2026-09-03', String(v)));
  assert.equal(L.toISO(new Date(2026, 8, 3)), '2026-09-03');
  assert.equal(L.toISO('2026-02-30'), ''); assert.equal(L.toISO('미정'), '');
});
test('금액 — 쉼표·₩·원·통화 글자·괄호 음수, 못 읽으면 NaN', () => {
  assert.equal(L.toNum('1,200,000'), 1200000); assert.equal(L.toNum('₩1,200,000'), 1200000); assert.equal(L.toNum('1,200,000원'), 1200000);
  assert.equal(L.toNum('USD 1,236.75'), 1236.75); assert.equal(L.toNum('(5,000)'), -5000); assert.equal(L.toNum(''), null);
  assert.ok(Number.isNaN(L.toNum('협의 중')));
});
test('부담 비율 — 0.6 · 「60%」 · 60 → 60, 1 → 100', () => {
  assert.equal(L.toPct(0.6), 60); assert.equal(L.toPct('60%'), 60); assert.equal(L.toPct(60), 60); assert.equal(L.toPct(1), 100);
  assert.equal(L.toPct('12.5 %'), 12.5); assert.equal(L.toPct(0.125), 12.5); assert.equal(L.toPct(null), null);
});

console.log('합의 결과 열 찾기');
test('샘플 — 제목 두 줄 아래 4행이 머리행, 14개 항목을 모두 찾음', () => {
  const d = L.detectColumns(AOA);
  assert.equal(d.headerRow, 3); assert.equal(d.missing.length, 0);
  assert.equal(Object.keys(d.map).length, 14);
  assert.equal(d.map.ratio, 10); assert.equal(d.map.agreedAmt, 11); assert.equal(d.map.supplierCode, 2);
});
test('다른 이름 — 업체명·보상금액·귀책률·클레임 금액 / 합의금액 없이 청구금액+비율만 있어도 됨', () => {
  let d = L.detectColumns([['업체명', '관리번호', '클레임 금액', '귀책률', '보상금액']]);
  assert.deepEqual([d.map.supplier, d.map.claimNo, d.map.claimAmt, d.map.ratio, d.map.agreedAmt], [0, 1, 2, 3, 4]);
  d = L.detectColumns([['거래처', '청구액', '부담비율']]);
  assert.equal(d.missing.length, 0);
  d = L.detectColumns([['거래처', '청구액', '품번']]);
  assert.deepEqual(d.missing, ['agreedAmt']);
});
test('머리행을 못 찾으면 알리고, 열을 고르면(overrides) 다시 맞춤', () => {
  const aoa = [['가', '나', '다'], ['A사', 'X', 100]];
  const d = L.detectColumns(aoa);
  assert.equal(d.headerRow, -1); assert.ok(d.warnings[0].includes('머리행'));
  const p = L.parseSheet(aoa, { overrides: { supplier: 0, agreedAmt: 2 } });
  assert.equal(p.rows.length, 1); assert.equal(p.rows[0].agreedRaw, 100);
});

console.log('행 읽기 · 금액 검증');
test('샘플 22행 읽기 — 맨 아래 「합계」 줄은 건너뛰고, 비율 0.6·「80%」를 퍼센트로', () => {
  assert.equal(parsed.rows.length, Sample.EXPECT.rows); assert.equal(parsed.skipped, 1);
  assert.equal(parsed.rows[4].ratio, 60); assert.equal(parsed.rows[19].ratio, 80); assert.equal(parsed.rows[19].currency, 'USD');
  assert.equal(parsed.rows[3].agreedRaw, null);
});
test('검사 표시 — 심어 둔 문제를 모두 찾음(확인 필요 2 · 주의 5)', () => {
  const ev = L.evaluateRows(parsed.rows, S);
  const count = code => ev.filter(r => r.flags.some(f => f.code === code)).length;
  assert.equal(count('noAmount'), 1); assert.equal(count('over'), 1); assert.equal(count('mismatch'), 2);
  assert.equal(count('computed'), 2); assert.equal(count('dupClaim'), 2); assert.equal(count('noClaimNo'), 1); assert.equal(count('zero'), 1);
  assert.equal(ev.filter(r => r.level === 'bad').length, Sample.EXPECT.bad);
  assert.equal(ev.filter(r => r.level === 'warn').length, Sample.EXPECT.warn);
  assert.equal(ev.filter(r => r.included).length, Sample.EXPECT.included);
});
test('비율로 계산 — 333,333 × 30% : 반올림 100,000 · 버림 99,999 · 100원 단위 버림 99,900', () => {
  const r = [{ supplier: 'X', claimNo: 'c', claimAmt: 333333, ratio: 30, agreedRaw: null }];
  assert.equal(L.evaluateRows(r, {})[0].agreed, 100000);
  assert.equal(L.evaluateRows(r, { roundMode: 'floor' })[0].agreed, 99999);
  assert.equal(L.evaluateRows(r, { roundMode: 'floor', roundUnit: 100 })[0].agreed, 99900);
  assert.equal(L.evaluateRows([{ supplier: 'X', claimNo: 'u', claimAmt: 1236.75, ratio: 50, currency: 'USD' }], {})[0].agreed, 618.38);
});
test('허용 차이 — 합의금액이 계산값과 허용 차이 안이면 「다름」 아님', () => {
  const r = [{ supplier: 'X', claimNo: 'c', claimAmt: 1000005, ratio: 30, agreedRaw: 300000 }];
  assert.ok(!L.evaluateRows(r, { tolerance: 10 })[0].flags.some(f => f.code === 'mismatch'));
  assert.ok(L.evaluateRows(r, { tolerance: 0 })[0].flags.some(f => f.code === 'mismatch'));
});

console.log('협력사별 묶기 · 금액');
test('6개 협력사 합계가 샘플 기대값과 같음(원화 5 · 달러 1), 금액 없는 행은 문서에서 뺌', () => {
  assert.equal(GS.length, Sample.EXPECT.suppliers);
  GS.forEach(g => assert.equal(g.total, Sample.EXPECT.totals[g.supplier], g.supplier));
  assert.equal(G('A정밀(가상)').count, 3); assert.equal(G('A정밀(가상)').excluded.length, 1);
  assert.equal(G('F Hydraulics Co.(가상)').currency, 'USD');
  assert.equal(G('B유압(가상)').ratio, Math.round(17610000 / 29300000 * 1000) / 10);
});
test('협력사 코드가 같으면 이름 표기가 달라도 한 묶음, 통화가 섞이면 확인 필요', () => {
  const rows = [{ supplier: 'A정밀', supplierCode: 'V1', claimNo: '1', agreedRaw: 100 }, { supplier: '(주)A정밀', supplierCode: 'V1', claimNo: '2', agreedRaw: 50, currency: 'USD' }];
  const g = L.groupBySupplier(rows, {});
  assert.equal(g.length, 1); assert.ok(g[0].flags.some(f => f.code === 'currencyMix')); assert.ok(g[0].bad >= 1);
});
test('반올림·부가세 — 음수·부동소수, 10% 별도면 총액 = 합계 + 부가세', () => {
  assert.equal(L.roundMoney(0.1 + 0.2, 'round', 0.01), 0.3); assert.equal(L.roundMoney(-1234.5, 'round', 1), -1235);
  assert.equal(L.roundMoney(1234.5, 'ceil', 100), 1300);
  const g = L.groupBySupplier(parsed.rows, { vat: 'add10' }).find(x => x.supplier === 'C기공(가상)');
  assert.equal(g.vat, 212800); assert.equal(g.total, 2128000 + 212800);
});
test('한글 금액 — 금 일백이십만원정, 영·일십·일억·일백만오백·큰 수', () => {
  assert.equal(L.koreanNumber(0), '영'); assert.equal(L.koreanNumber(10), '일십'); assert.equal(L.koreanNumber(1200000), '일백이십만');
  assert.equal(L.koreanNumber(100000000), '일억'); assert.equal(L.koreanNumber(1000500), '일백만오백');
  assert.equal(L.koreanNumber(27633000), '이천칠백육십삼만삼천'); assert.equal(L.koreanNumber(1234567890123), '일조이천삼백사십오억육천칠백팔십구만일백이십삼');
  assert.equal(L.amountWords(1200000, 'KRW'), '금 일백이십만원정'); assert.equal(L.amountLine(1200000, 'KRW'), '금 일백이십만원정(₩1,200,000)');
  assert.equal(L.amountWords(14138.38, 'USD'), 'USD 14,138.38');
});

console.log('문서 번호 · 발행 이력');
test('번호 형식 — 접두·연·월·일련(자릿수)', () => {
  assert.equal(L.formatDocNo('{접두}-{YYYY}-{MM}-{NNN}', { prefix: 'AG', date: '2026-10-21', seq: 7 }), 'AG-2026-10-007');
  assert.equal(L.formatDocNo('{접두}{YY}{MM}{DD}-{NN}', { prefix: 'INV', date: '2026-01-05', seq: 3 }), 'INV260105-03');
  assert.equal(L.cleanSettings({ docFormat: '번호없음' }).docFormat, '{접두}-{YYYY}-{MM}-{NNN}');
});
test('채번 — 종류마다·달마다 1부터, 같은 내용은 같은 번호(재발행)', () => {
  let h = [];
  const r1 = L.issueDoc(h, 'ag', G('A정밀(가상)'), '2026-10-21', S); h = r1.history;
  const r2 = L.issueDoc(h, 'ag', G('B유압(가상)'), '2026-10-21', S); h = r2.history;
  const r3 = L.issueDoc(h, 'inv', G('A정밀(가상)'), '2026-10-21', S); h = r3.history;
  const r4 = L.issueDoc(h, 'ag', G('C기공(가상)'), '2026-11-02', S); h = r4.history;
  assert.deepEqual([r1.entry.docNo, r2.entry.docNo, r3.entry.docNo, r4.entry.docNo], ['AG-2026-10-001', 'AG-2026-10-002', 'INV-2026-10-001', 'AG-2026-11-001']);
  const again = L.issueDoc(h, 'ag', G('A정밀(가상)'), '2026-10-25', S);
  assert.equal(again.reused, true); assert.equal(again.entry.docNo, 'AG-2026-10-001'); assert.equal(again.entry.times, 2);
  assert.equal(h.find(x => x.docNo === 'AG-2026-10-001').times, 1, '원본 이력은 바꾸지 않음');
});
test('내용이 바뀌면 새 번호, 이전 번호는 「대체됨」 / 취소 표시한 번호는 다시 쓰지 않음', () => {
  let h = L.issueDoc([], 'ag', G('A정밀(가상)'), '2026-10-21', S).history;
  const changed = L.groupBySupplier(parsed.rows.map(r => r.claimNo === 'CL-2609-004' ? { ...r, agreedRaw: 900000 } : r), S).find(g => g.supplier === 'A정밀(가상)');
  const r = L.issueDoc(h, 'ag', changed, '2026-10-21', S);
  assert.equal(r.reused, false); assert.equal(r.entry.docNo, 'AG-2026-10-002'); assert.equal(r.entry.replaces, 'AG-2026-10-001');
  assert.equal(r.history[0].status, 'replaced');
  r.history[1].status = 'void';
  assert.equal(L.issueDoc(r.history, 'ag', changed, '2026-10-21', S).entry.docNo, 'AG-2026-10-003');
});

console.log('자리표시(키)');
test('키 찾기·칸 하나 채우기 — 금액 키 하나면 숫자, 글과 섞이면 쉼표 표기, 모르는 키는 그대로', () => {
  const ctx = L.buildContext(G('B유압(가상)'), 'ag', { issuer: '예시기계(가상)' }, { docNo: 'AG-2026-10-002', issueDate: '2026-10-21' });
  assert.deepEqual(L.placeholdersIn('{{협력사명}} 귀하 {{ 총액 }}'), ['협력사명', '총액']);
  assert.deepEqual(L.fillCell('{{총액}}', ctx), { type: 'n', value: 17610000, kind: 'money' });
  assert.deepEqual(L.fillCell(' {{ 협력사명 }} ', ctx), { type: 's', value: 'B유압(가상)' });
  assert.equal(L.fillCell('합의 금액 {{합의합계}}원 / {{없는키}}', ctx).value, '합의 금액 17,610,000원 / {{없는키}}');
  assert.equal(L.fillText('{{총액_표기}}', ctx), '금 일천칠백육십일만원정(₩17,610,000)');
  assert.equal(L.fillText('{{지급기한}} · {{발행일_한글}} · {{발행처}}', ctx), '2026-11-20 · 2026년 10월 21일 · 예시기계(가상)');
});
test('행 키 — {{행.합의금액}} 은 행 값, 비율은 「60%」 글자', () => {
  const ctx = L.buildContext(G('B유압(가상)'), 'ag', S, {});
  assert.equal(ctx.rows.length, 4);
  assert.deepEqual(L.fillCell('{{행.합의금액}}', ctx, ctx.rows[0]), { type: 'n', value: 7440000, kind: 'money' });
  assert.deepEqual(L.fillCell('{{행.부담비율}}', ctx, ctx.rows[0]), { type: 's', value: '60%' });
  assert.equal(L.fillText('{{행.번호}}. {{행.클레임번호}} ({{통화}})', ctx, ctx.rows[3]), '4. CL-2610-003 (KRW)');
});

console.log('엑셀 양식 채우기');
test('수식 주소 옮기기 — 상대·절대, 범위 늘리기, 따옴표 글·다른 시트·함수 이름은 그대로', () => {
  assert.equal(T.relShift('G14*H14', 2, 0), 'G16*H16');
  assert.equal(T.relShift('$G$14*H14+G$14', 2, 0), '$G$14*H16+G$14');
  assert.equal(T.insertShift('SUM(I14:I21)', 22, 3, 14), 'SUM(I14:I24)');
  assert.equal(T.insertShift('SUM(I14:I14)', 15, 5, 14), 'SUM(I14:I19)');
  assert.equal(T.insertShift('A30+B5', 22, 3, 14), 'A33+B5');
  assert.equal(T.insertShift('"A30"&Sheet2!A30&LOG10(A30)', 22, 3, 14), '"A30"&Sheet2!A30&LOG10(A33)');
  assert.equal(T.insertShift("'합의서'!$A$1:$I$34", 22, 3, 14, { sheet: '합의서' }), "'합의서'!$A$1:$I$37");
});
test('기본 양식 살펴보기 — 합의서 반복 줄 14행(9칸), 인보이스 11행, 모르는 키 없음', () => {
  const a = T.scanTemplate(tpl('ag')), v = T.scanTemplate(tpl('inv'));
  assert.equal(a.sheets[0].repeatRow, 14); assert.equal(a.sheets[0].rowKeys.length, 9); assert.deepEqual(a.unknown, []);
  assert.equal(v.sheets[0].repeatRow, 11); assert.deepEqual(v.unknown, []);
  assert.ok(a.keys.includes('총액_표기') && v.keys.includes('입금계좌'));
});
test('빈 표 줄 안에 들어가면(4건) 줄을 끼우지 않고, 서식(스타일 번호)·병합·숫자 칸을 그대로 둠', () => {
  const ctx = L.buildContext(G('B유압(가상)'), 'ag', { issuer: '예시기계(가상)' }, { docNo: 'AG-2026-10-002', issueDate: '2026-10-21' });
  const src = sheetXml(tpl('ag')), res = T.fillTemplate(tpl('ag'), ctx), out = sheetXml(res.data);
  assert.equal(res.report.sheets[0].inserted, 0);
  const sOf = (x, ref) => (new RegExp('<c r="' + ref + '"[^>]*?s="(\\d+)"').exec(x) || [])[1];
  ['A1', 'C9', 'I14', 'I15', 'I22', 'A33'].forEach(ref => assert.equal(sOf(out, ref), sOf(src, ref), ref + ' 스타일'));
  assert.equal((out.match(/<mergeCell /g) || []).length, (src.match(/<mergeCell /g) || []).length);
  assert.equal(T.openZip(res.data).get('xl/styles.xml'), T.openZip(tpl('ag')).get('xl/styles.xml'), '원화면 styles.xml 그대로');
  const ws = readBack(res.data).Sheets['합의서'];
  assert.equal(ws.B14.v, 'CL-2609-002'); assert.equal(ws.B17.v, 'CL-2610-003'); assert.equal(ws.B18, undefined);
  assert.equal(ws.I14.t, 'n'); assert.equal(ws.I14.v, 7440000); assert.equal(ws.H14.v, '60%');
  assert.equal(ws.C9.v, '금 일천칠백육십일만원정(₩17,610,000)'); assert.equal(ws.C3.v, 'AG-2026-10-002');
  assert.equal(ws.A10.v, '(합의합계 17,610,000 · 부가세 0 · 클레임 4건 · 청구합계 대비 60.1%)');
  assert.ok(!/\{\{/.test(out), '남은 {{키}} 없음');
});
test('넘치면(18건) 줄 10개를 끼우고 합계 줄·SUM 범위·병합·인쇄 영역을 함께 내림', () => {
  const ctx = L.buildContext(BIG, 'ag', S, { docNo: 'AG-2026-10-003', issueDate: '2026-10-21' });
  const res = T.fillTemplate(tpl('ag'), ctx), x = sheetXml(res.data);
  assert.equal(res.report.sheets[0].inserted, 10);
  assert.ok(x.includes('<f>SUM(I14:I31)</f>'), 'SUM 범위 늘림');
  assert.ok(/<mergeCell ref="A32:E32"\/>/.test(x) && /<mergeCell ref="A43:D43"\/>/.test(x), '병합 내림');
  const wbx = T.openZip(res.data).get('xl/workbook.xml');
  assert.ok(wbx.includes("'합의서'!$A$1:$I$44"), '인쇄 영역'); assert.ok(/fullCalcOnLoad="1"/.test(wbx));
  const ws = readBack(res.data).Sheets['합의서'];
  for (let r = 14; r <= 31; r++) assert.equal(ws['A' + r].v, r - 13, 'A' + r);
  assert.equal(ws.A32.v, '합계'); assert.equal(ws.G32.v, BIG.claimSum); assert.equal(ws.A41.v, '2026년 10월 16일');
  assert.ok(ws['!merges'].some(m => m.s.r === 31 && m.e.c === 4));
});
test('공유 문자열표(엑셀·LibreOffice 가 저장한 양식)도 채움 — 남은 키 0, 인쇄 영역(따옴표 없는 한글 시트명)도 내림', () => {
  const u = new Uint8Array(fs.readFileSync(new URL('./fixtures/합의서_공유문자열.xlsx', import.meta.url)));
  assert.ok(T.openZip(u).has('xl/sharedStrings.xml'));
  assert.equal(T.scanTemplate(u).sheets[0].repeatRow, 14);
  const res = T.fillTemplate(u, L.buildContext(BIG, 'ag', S, { docNo: 'AG-X', issueDate: '2026-10-21' }));
  const ws = readBack(res.data).Sheets['합의서'];
  assert.equal(ws.B31.v, BIG.rows[17].claimNo); assert.equal(ws.C3.v, 'AG-X');
  assert.ok(!/\{\{/.test(XLSX.utils.sheet_to_csv(ws)), '남은 {{키}} 없음');
  assert.ok(T.openZip(res.data).get('xl/workbook.xml').includes('합의서!$A$1:$I$44'));
});
test('달러 — 금액 칸에 소수 둘째 자리 스타일(#,##0.00)을 새로 붙임', () => {
  const ctx = L.buildContext(G('F Hydraulics Co.(가상)'), 'inv', S, { docNo: 'INV-1', issueDate: '2026-10-21' });
  const res = T.fillTemplate(tpl('inv'), ctx), z = T.openZip(res.data), st = z.get('xl/styles.xml'), x = z.get('xl/worksheets/sheet1.xml');
  const count = s => +/<cellXfs count="(\d+)"/.exec(s)[1];
  assert.equal(count(st), count(T.openZip(tpl('inv')).get('xl/styles.xml')) + 2, '금액 칸 스타일 2종 복제(표 줄·합계 줄)');
  const s12 = +/<c r="H12" s="(\d+)"/.exec(x)[1];
  const xfs = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(st)[1].match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g);
  assert.ok(/numFmtId="4"/.test(xfs[s12]));
  assert.equal(readBack(res.data).Sheets['인보이스'].H12.v, 618.38);
});
test('클레임이 0건이면 반복 줄의 키만 비움 / 시트 밖 부분: 계산 순서 파일 지우기·그림 위치 내리기', () => {
  const empty = { ...G('A정밀(가상)'), rows: [], count: 0 };
  const r0 = T.fillTemplate(tpl('inv'), L.buildContext(empty, 'inv', S, {}));
  const ws0 = readBack(r0.data).Sheets['인보이스'];
  assert.equal(ws0.B11, undefined); assert.ok(!/\{\{/.test(sheetXml(r0.data)));
  // 기본 양식에 직인 그림(40행 근처 앵커)과 계산 순서 파일을 붙여 넣고 채움
  const z = T.openZip(tpl('ag'));
  z.set('xl/drawings/drawing1.xml', '<xdr:wsDr xmlns:xdr="x"><xdr:twoCellAnchor><xdr:from><xdr:col>7</xdr:col><xdr:row>32</xdr:row></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:row>34</xdr:row></xdr:to></xdr:twoCellAnchor><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>2</xdr:row></xdr:from></xdr:oneCellAnchor></xdr:wsDr>');
  z.set('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
  z.set('xl/calcChain.xml', '<calcChain><c r="I22" i="1"/></calcChain>');
  z.set('[Content_Types].xml', z.get('[Content_Types].xml').replace('</Types>', '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>'));
  const res = T.fillTemplate(z.write(), L.buildContext(BIG, 'ag', S, {}));
  const z2 = T.openZip(res.data), d = z2.get('xl/drawings/drawing1.xml');
  assert.ok(d.includes('<xdr:row>42</xdr:row>') && d.includes('<xdr:row>44</xdr:row>') && d.includes('<xdr:row>2</xdr:row>'), '그림 앵커 ' + d.match(/<xdr:row>\d+/g));
  assert.equal(z2.has('xl/calcChain.xml'), false); assert.ok(!z2.get('[Content_Types].xml').includes('calcChain'));
});
test('공유 수식(t="shared")을 칸마다 풀고, 반복 줄의 수식은 줄마다 옮겨 씀', () => {
  const xml = '<worksheet><sheetData>' +
    '<row r="5"><c r="A5" t="inlineStr"><is><t>{{행.번호}}</t></is></c><c r="B5" t="inlineStr"><is><t>{{행.합의금액}}</t></is></c><c r="C5"><f t="shared" ref="C5:C6" si="0">B5*2</f><v>0</v></c></row>' +
    '<row r="6"><c r="C6"><f t="shared" si="0"/><v>0</v></c></row>' +
    '<row r="7"><c r="A7" t="inlineStr"><is><t>합계</t></is></c><c r="C7"><f>SUM(C5:C6)</f></c></row>' +
    '</sheetData></worksheet>';
  const ctx = L.buildContext(G('B유압(가상)'), 'ag', S, {});
  const r = T.fillSheet(xml, ctx, [], 'S');
  assert.equal(r.repeatRow, 5); assert.equal(r.avail, 1); assert.equal(r.inserted, 2);
  assert.ok(r.xml.includes('<c r="C6"><f>B6*2</f>') && r.xml.includes('<c r="C8"><f>B8*2</f>'), r.xml);
  assert.ok(r.xml.includes('<c r="C9"><f>SUM(C5:C8)</f>'));
  assert.ok(!r.xml.includes('t="shared"'));
});
test('여러 파일 zip — 한글 이름(UTF-8 표시), CRC, 내용이 그대로', () => {
  const a = new TextEncoder().encode('첫 파일'), b = new Uint8Array([1, 2, 3, 250]);
  const z = T.makeZip([{ name: '합의서_A정밀.xlsx', data: a }, { name: 'b.bin', data: b }]);
  const dv = new DataView(z.buffer);
  assert.equal(dv.getUint32(0, true), 0x04034b50); assert.equal(dv.getUint16(6, true), 0x0800);
  const nameLen = dv.getUint16(26, true);
  assert.equal(new TextDecoder().decode(z.slice(30, 30 + nameLen)), '합의서_A정밀.xlsx');
  assert.equal(dv.getUint32(14, true), T.crc32(a));
  assert.equal(new TextDecoder().decode(z.slice(30 + nameLen, 30 + nameLen + a.length)), '첫 파일');
  assert.equal(T.crc32(new TextEncoder().encode('123456789')), 0xCBF43926);
  const end = z.length - 22; assert.equal(dv.getUint32(end, true), 0x06054b50); assert.equal(dv.getUint16(end + 10, true), 2);
});
test('xlsx 가 아니면 알아들을 수 있는 말로 알림', () => {
  assert.throws(() => T.scanTemplate(new TextEncoder().encode('a,b\n1,2')), /엑셀|xlsx/);
});

console.log('저장');
test('저장 모양을 JSON 으로 왕복해도 같고, 망가진 값은 기본값으로', () => {
  const db = L.emptyDb(); db.rows = parsed.rows; db.settings.vat = 'add10'; db.history = L.issueDoc([], 'ag', GS[0], '2026-10-21', S).history;
  db.templates.ag = { name: 'x.xlsx', b64: 'AAAA', at: '2026-10-21' };
  const back = L.restoreDb(JSON.parse(JSON.stringify(db)));
  assert.equal(back.rows.length, 22); assert.equal(back.settings.vat, 'add10'); assert.equal(back.history.length, 1); assert.equal(back.templates.ag.name, 'x.xlsx');
  const bad = L.restoreDb({ rows: [{ x: 1 }, null], settings: { vat: '?', roundUnit: 7, payDays: 'abc' }, history: [{ docNo: 'a', kind: 'zz' }], templates: { ag: { name: 'x' } } });
  assert.equal(bad.rows.length, 0); assert.equal(bad.settings.vat, 'none'); assert.equal(bad.settings.roundUnit, 1); assert.equal(bad.settings.payDays, 30);
  assert.equal(bad.history.length, 0); assert.equal(bad.templates.ag, null);
});
test('검토표·이력 내보내기 표 — 머리와 칸 수가 같음', () => {
  const rv = L.reviewRows(parsed.rows, S); assert.equal(rv.length, 22); rv.forEach(r => assert.equal(r.length, L.REVIEW_HEAD.length));
  assert.equal(rv[3][L.REVIEW_HEAD.indexOf('판정')], '확인 필요');
  const hr = L.historyRows(L.issueDoc([], 'inv', GS[1], '2026-10-21', S).history); assert.equal(hr[0].length, L.HISTORY_HEAD.length); assert.equal(hr[0][0], 'INV-2026-10-001');
});

const { versions } = await import('../scripts/bump-version.mjs');
test('판 맞추기 — index.html ?v= · logic.js · app.js 가 같은 판 (scripts/bump-version.mjs)', () => {
  const v = versions();
  assert.equal(v.logic, L.VERSION, 'logic.js VERSION'); assert.equal(v.app, L.VERSION, 'app.js APP_VERSION');
  assert.ok(v.html.length >= 8 && v.html.every(x => x === L.VERSION), 'index.html 의 모든 js·css 에 ?v=' + L.VERSION + ' — ' + JSON.stringify(v.html));
});

console.log('\n' + passed + '개 통과' + (process.exitCode ? ' — 실패 있음' : ''));
