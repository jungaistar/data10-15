// 가상 합의 결과 엑셀 만들기: node scripts/make-samples.js
// js/sample-data.js 의 22행을 samples/클레임_합의결과_샘플.xlsx 로 씁니다(날짜는 날짜 칸, 부담률은 0% 표시 형식, 금액은 #,##0).
// 쓴 뒤 앱과 같은 길(logic.js)로 다시 읽어 행 수·협력사별 합계가 sample-data.js 의 EXPECT 와 같은지 확인합니다.
import { createRequire } from 'node:module';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const XLSX = require(path.join(root, 'vendor/xlsx.full.min.js'));
const L = require(path.join(root, 'js/logic.js'));
const Sample = require(path.join(root, 'js/sample-data.js'));

const b = Sample.build();
const aoa = b.aoa.map((row, i) => i < 4 ? row : row.map((v, c) => b.dateCols.includes(c) && typeof v === 'string' && v ? new Date(v + 'T00:00:00') : v));
const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
const range = XLSX.utils.decode_range(ws['!ref']);
for (let r = 4; r <= range.e.r; r++) {
  for (const c of [9, 11]) { const a = ws[XLSX.utils.encode_cell({ r, c })]; if (a && a.t === 'n') a.z = '#,##0.##'; }
  const p = ws[XLSX.utils.encode_cell({ r, c: 10 })]; if (p && p.t === 'n') p.z = '0%';
  for (const c of b.dateCols) { const d = ws[XLSX.utils.encode_cell({ r, c })]; if (d && d.t === 'd') d.z = 'yyyy-mm-dd'; }
}
ws['!cols'] = [6, 22, 9, 13, 11, 7, 9, 18, 6, 12, 11, 12, 6, 11, 18].map(w => ({ wch: w }));
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, '합의결과');
const out = path.join(root, 'samples/클레임_합의결과_샘플.xlsx');
fs.writeFileSync(out, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer', bookSST: true }));

// 다시 읽어 확인
const back = XLSX.read(fs.readFileSync(out), { cellDates: true });
const sheet = XLSX.utils.sheet_to_json(back.Sheets[back.SheetNames[0]], { header: 1, raw: true, defval: null });
const parsed = L.parseSheet(sheet);
const groups = L.groupBySupplier(parsed.rows, {});
const bad = [];
if (parsed.rows.length !== Sample.EXPECT.rows) bad.push('행 수 ' + parsed.rows.length);
groups.forEach(g => { if (Sample.EXPECT.totals[g.supplier] !== g.total) bad.push(g.supplier + ' ' + g.total); });
console.log('wrote samples/클레임_합의결과_샘플.xlsx ·', parsed.rows.length, '행 ·', groups.length, '개 협력사', bad.length ? '— 다름: ' + bad.join(', ') : '— 확인 통과');
process.exitCode = bad.length ? 1 : 0;
