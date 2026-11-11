/* 클레임 보상 합의서·인보이스 — 순수 로직 (브라우저·node 공용, 화면 코드 없음)
   1. 글자·숫자·날짜      2. 합의 결과 열 찾기     3. 행 읽기·금액 정하기·검증
   4. 협력사별 묶기       5. 금액(반올림·부가세·한글 금액)
   6. 문서 번호·발행 이력  7. 자리표시(키) 목록과 채우기 값
   8. 설정·저장 모양
   실제 합의 결과 엑셀·양식은 아직 받지 않았습니다(기획서 10장 ①②). 열 이름 별칭은 흔한 표기를 넓게 잡아 두었습니다. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CLLogic = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  /* 판(版) — index.html 의 ?v= · app.js 의 APP_VERSION 과 같아야 합니다(scripts/bump-version.mjs 로 함께 올림).
     브라우저가 옛 logic.js 와 새 app.js 를 섞어 쓰면 app.js 가 이 값을 보고 계산을 멈춥니다. */
  var VERSION = '2026-11-11.6';

  /* ── 1. 글자·숫자·날짜 ───────────────────────────────────── */
  function str(v) { return v == null ? '' : String(v).trim(); }
  function norm(s) { return String(s == null ? '' : s).replace(/[\s()（）\[\]·.,:;_\/\-~'"#*%]/g, '').toLowerCase(); }
  function pad(n, w) { var s = String(n); while (s.length < (w || 2)) s = '0' + s; return s; }
  function isoFromUTC(ms) { var d = new Date(ms); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  function isoOk(y, m, d) {
    if (!(y >= 1990 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return '';
    var t = new Date(Date.UTC(y, m - 1, d));
    return t.getUTCMonth() === m - 1 ? isoFromUTC(t.getTime()) : '';
  }
  /* 엑셀 날짜(일련번호 · Date · 글자) → 'YYYY-MM-DD'. 못 읽으면 '' */
  function toISO(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date || Object.prototype.toString.call(v) === '[object Date]') {
      if (isNaN(v.getTime())) return '';
      var d = new Date(v.getTime() + 12 * 3600 * 1000);
      return isoOk(d.getFullYear(), d.getMonth() + 1, d.getDate());
    }
    if (typeof v === 'number') {
      if (v < 20000 || v > 80000) return '';
      return isoFromUTC(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000);
    }
    var s = String(v).trim(), m;
    if ((m = /^(\d{4})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})/.exec(s))) return isoOk(+m[1], +m[2], +m[3]);
    if ((m = /^(\d{2})[-./](\d{1,2})[-./](\d{1,2})$/.exec(s))) return isoOk(2000 + +m[1], +m[2], +m[3]);
    if ((m = /^(\d{4})(\d{2})(\d{2})$/.exec(s))) return isoOk(+m[1], +m[2], +m[3]);
    if (/^\d{5}(\.\d+)?$/.test(s)) return toISO(Number(s));
    return '';
  }
  function isoToUTC(iso) { var p = String(iso).split('-'); return Date.UTC(+p[0], +p[1] - 1, +p[2]); }
  function addDays(iso, n) { return isoFromUTC(isoToUTC(iso) + n * 86400000); }
  function localToday(now) { var d = now || new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function koDate(iso) { if (!iso) return ''; var p = iso.split('-'); return p[0] + '년 ' + (+p[1]) + '월 ' + (+p[2]) + '일'; }
  function colName(i) { var s = ''; i = Number(i); do { s = String.fromCharCode(65 + (i % 26)) + s; i = Math.floor(i / 26) - 1; } while (i >= 0); return s; }
  /* 금액·수량 글자 → 숫자. 빈칸 null, 못 읽으면 NaN. 「₩1,200,000」「1,200,000원」「USD 1,234.50」「(5,000)」(음수) */
  function toNum(v) {
    if (v == null) return null;
    if (typeof v === 'number') return isFinite(v) ? v : NaN;
    var s = String(v).trim();
    if (!s || s === '-') return null;
    var neg = /^\(.*\)$/.test(s);
    s = s.replace(/^\(|\)$/g, '').replace(/[,\s]/g, '').replace(/^(₩|\\|krw|usd|eur|jpy|cny|\$|€|¥)/i, '').replace(/(원|건|개|ea|pcs|krw|usd|eur)$/i, '');
    if (!/^-?\d+(\.\d+)?$/.test(s)) return NaN;
    return neg ? -Number(s) : Number(s);
  }
  /* 부담 비율 → 퍼센트 숫자. 「30%」 「30」 → 30, 엑셀 백분율 칸(0.3) → 30 */
  function toPct(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return !isFinite(v) ? NaN : (Math.abs(v) <= 1 ? Math.round(v * 1e6) / 1e4 : v);
    var s = String(v).trim().replace(/\s/g, '');
    if (!s || s === '-') return null;
    var pct = /%$/.test(s); s = s.replace(/%$/, '');
    if (!/^-?\d+(\.\d+)?$/.test(s)) return NaN;
    var n = Number(s);
    return pct || Math.abs(n) > 1 ? n : Math.round(n * 1e6) / 1e4;
  }
  function hash(s) { var h = 5381; s = String(s); for (var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return (h >>> 0).toString(36); }

  /* ── 2. 합의 결과 열 찾기 ─────────────────────────────────── */
  var FIELDS = [
    { id: 'supplier', label: '협력사명', req: true, aliases: ['협력사', '협력사명', '협력업체', '협력업체명', '업체', '업체명', '공급사', '공급업체', '거래처', '거래처명', '상호', 'vendor', 'vendorname', 'supplier', 'suppliername', 'partner'] },
    { id: 'supplierCode', label: '협력사 코드', aliases: ['협력사코드', '업체코드', '거래처코드', '공급사코드', '사업자번호', '사업자등록번호', 'vendorcode', 'suppliercode', 'vendorno'] },
    { id: 'claimNo', label: '클레임 번호', aliases: ['클레임번호', '클레임no', '클레임', '클레임관리번호', '관리번호', '접수번호', '보상번호', 'claimno', 'claimnumber', 'claim', 'claimid'] },
    { id: 'date', label: '발생일', aliases: ['발생일', '발생일자', '클레임발생일', '불량발생일', '고장일', '고장일자', '접수일', '접수일자', '발생', 'occurdate', 'failuredate', 'date'] },
    { id: 'model', label: '기종·모델', aliases: ['기종', '모델', '모델명', '장비', '장비명', '차종', 'model', 'machine'] },
    { id: 'partNo', label: '품번', aliases: ['품번', '부품번호', '파트넘버', '자재번호', '품목코드', 'partno', 'partnumber', 'pn', 'itemno'] },
    { id: 'partName', label: '품명', aliases: ['품명', '부품명', '자재명', '품목명', '내용', '불량내용', 'partname', 'description', 'itemname'] },
    { id: 'qty', label: '건수·수량', aliases: ['건수', '수량', '불량수량', '클레임건수', '대수', 'qty', 'quantity', 'count'] },
    { id: 'claimAmt', label: '청구금액', aliases: ['청구금액', '청구액', '클레임금액', '발생금액', '손실금액', '총비용', '수리비', '청구', 'claimamount', 'claimed', 'claimedamount', 'cost'] },
    { id: 'ratio', label: '부담 비율', aliases: ['부담비율', '부담률', '협력사부담비율', '협력사부담률', '귀책비율', '귀책률', '분담비율', '분담률', '구상비율', 'ratio', 'share', 'liability'] },
    { id: 'agreedAmt', label: '합의금액', aliases: ['합의금액', '합의액', '보상금액', '보상액', '확정금액', '최종금액', '구상금액', '협력사부담금액', '부담금액', 'agreedamount', 'settlement', 'settledamount', 'compensation'] },
    { id: 'currency', label: '통화', aliases: ['통화', '화폐', '통화단위', 'currency', 'cur', 'ccy'] },
    { id: 'agreeDate', label: '합의일', aliases: ['합의일', '합의일자', '확정일', '합의완료일', 'agreedate', 'agreeddate', 'settlementdate'] },
    { id: 'note', label: '비고', aliases: ['비고', '메모', '특이사항', '참고', '합의내용', 'remark', 'remarks', 'note', 'notes', 'comment'] }
  ];
  var FIELD_BY_ID = {}; FIELDS.forEach(function (f) { FIELD_BY_ID[f.id] = f; });

  function matchField(text) {
    var n = norm(text);
    if (!n) return null;
    for (var i = 0; i < FIELDS.length; i++) if (FIELDS[i].aliases.indexOf(n) >= 0) return FIELDS[i].id;
    if (/합의|보상|구상|확정|settle|agreed/.test(n) && /금액|액$|amount/.test(n)) return 'agreedAmt';
    if (/합의|확정/.test(n) && /일/.test(n)) return 'agreeDate';
    if (/청구|클레임금|발생금|claim/.test(n) && /금액|액$|amount|비용/.test(n)) return 'claimAmt';
    if (/비율|률$|ratio|share/.test(n)) return 'ratio';
    if (/클레임/.test(n) && /번호|no/.test(n)) return 'claimNo';
    if (/협력|업체|거래처|공급사|vendor|supplier/.test(n)) return /코드|번호|code|no/.test(n) ? 'supplierCode' : 'supplier';
    if (/발생|고장/.test(n)) return 'date';
    if (/품번|부품번호|partno/.test(n)) return 'partNo';
    if (/품명|부품명/.test(n)) return 'partName';
    if (/수량|건수|qty/.test(n)) return 'qty';
    return null;
  }
  function amountMissing(map) { return map.agreedAmt == null && !(map.claimAmt != null && map.ratio != null); }

  /* aoa: 시트를 행 배열로 읽은 것(SheetJS sheet_to_json header:1)
     반환: { headerRow, map:{field:col}, headers:[글자], missing:[field], warnings:[] }
     꼭 필요: 협력사명 + (합의금액 또는 청구금액·부담 비율) */
  function detectColumns(aoa, overrides) {
    aoa = aoa || [];
    var best = -1, bestScore = 0, bestMap = null, bestHeaders = null, bestDup = null, anyRow = 0, anyScore = 0;
    for (var r = 0; r < Math.min(aoa.length, 30); r++) {
      var row = aoa[r] || [], map = {}, headers = [], dup = [];
      for (var c = 0; c < row.length; c++) {
        var t = str(row[c]), fid = matchField(t);
        headers.push(t);
        if (fid) { if (map[fid] == null) map[fid] = c; else dup.push({ label: t, fid: fid }); }
      }
      var score = Object.keys(map).length;
      if (score > anyScore) { anyScore = score; anyRow = r; }
      if (map.supplier != null && (map.agreedAmt != null || map.claimAmt != null) && score > bestScore) { best = r; bestScore = score; bestMap = map; bestHeaders = headers; bestDup = dup; }
    }
    var out = { headerRow: best, map: {}, headers: [], missing: [], warnings: [] };
    if (best < 0 || bestScore < 2) {
      out.headerRow = -1;
      out.warnings.push('머리행(협력사·클레임 번호·합의금액 …)을 찾지 못했습니다. 합의 결과 시트인지 확인하거나 열 연결에서 골라 주세요.');
      out.missing = ['supplier', 'agreedAmt'];
      if (!overrides) return out;
      out.headerRow = anyRow; out.headers = (aoa[anyRow] || []).map(str); out.missing = []; out.warnings = [];
      out.headers.forEach(function (t, c) { var f = matchField(t); if (f && out.map[f] == null) out.map[f] = c; });
    } else {
      out.map = bestMap; out.headers = bestHeaders;
      bestDup.forEach(function (d) { out.warnings.push('「' + d.label + '」 열이 ' + FIELD_BY_ID[d.fid].label + '(으)로 한 번 더 보입니다. 앞의 것(' + colName(out.map[d.fid]) + ')을 씁니다.'); });
    }
    if (overrides) Object.keys(overrides).forEach(function (fid) {
      if (!FIELD_BY_ID[fid]) return;
      var v = overrides[fid];
      if (v === -1 || v === '' || v == null) delete out.map[fid]; else out.map[fid] = Number(v);
    });
    if (out.map.supplier == null) out.missing.push('supplier');
    if (amountMissing(out.map)) out.missing.push('agreedAmt');
    if (out.missing.length) out.warnings.push('꼭 필요한 열이 없습니다: ' + out.missing.map(function (id) { return id === 'agreedAmt' ? '합의금액(또는 청구금액 + 부담 비율)' : FIELD_BY_ID[id].label; }).join(', ') + ' — 열 연결에서 골라 주세요.');
    return out;
  }
  function pickSheet(sheets) {
    var best = null, score = -1;
    (sheets || []).forEach(function (s) {
      var d = detectColumns(s.aoa);
      var sc = d.headerRow < 0 ? 0 : Object.keys(d.map).length * 10 + Math.min(9, (s.aoa.length - d.headerRow - 1) / 10);
      if (sc > score) { score = sc; best = s; }
    });
    return best;
  }

  /* ── 3. 행 읽기·금액 정하기·검증 ─────────────────────────── */
  function isSummaryRow(row) { var t = (row || []).map(str).filter(Boolean).join(' '); return /^(합계|소계|총계|누계|total|sub\s*total)\b/i.test(t.trim()); }
  function parseSheet(aoa, opts) {
    opts = opts || {};
    var det = detectColumns(aoa, opts.overrides);
    var out = { det: det, rows: [], warnings: [], skipped: 0 };
    if (det.headerRow < 0 || det.missing.length) return out;
    var m = det.map;
    function cell(row, f) { return m[f] == null ? null : row[m[f]]; }
    for (var r = det.headerRow + 1; r < aoa.length; r++) {
      var row = aoa[r] || [], line = r + 1;
      if (!row.some(function (v) { return str(v) !== ''; })) continue;
      if (isSummaryRow(row)) { out.skipped++; continue; }
      var sup = str(cell(row, 'supplier'));
      if (!sup) { out.skipped++; out.warnings.push({ row: line, msg: '협력사명이 비어 건너뜁니다.' }); continue; }
      var rec = {
        line: line, supplier: sup, supplierCode: str(cell(row, 'supplierCode')), claimNo: str(cell(row, 'claimNo')),
        date: toISO(cell(row, 'date')), model: str(cell(row, 'model')), partNo: str(cell(row, 'partNo')), partName: str(cell(row, 'partName')),
        qty: null, claimAmt: null, ratio: null, agreedRaw: null, currency: str(cell(row, 'currency')).toUpperCase(),
        agreeDate: toISO(cell(row, 'agreeDate')), note: str(cell(row, 'note'))
      };
      [['qty', 'qty'], ['claimAmt', 'claimAmt'], ['agreedRaw', 'agreedAmt']].forEach(function (p) {
        var raw = cell(row, p[1]), n = toNum(raw);
        if (n != null && isNaN(n)) out.warnings.push({ row: line, msg: FIELD_BY_ID[p[1]].label + ' 「' + str(raw) + '」을 숫자로 읽지 못했습니다.' });
        else rec[p[0]] = n;
      });
      var pr = toPct(cell(row, 'ratio'));
      if (pr != null && isNaN(pr)) out.warnings.push({ row: line, msg: '부담 비율 「' + str(cell(row, 'ratio')) + '」을 읽지 못했습니다.' });
      else rec.ratio = pr;
      if (rec.currency === '원' || rec.currency === '₩' || rec.currency === 'WON') rec.currency = 'KRW';
      if (rec.currency === '$' || rec.currency === 'US$') rec.currency = 'USD';
      rec.id = 'r' + hash([rec.supplier, rec.supplierCode, rec.claimNo, rec.partNo, line].join('|'));
      out.rows.push(rec);
    }
    return out;
  }

  /* 금액 반올림 — mode: round(반올림)·floor(버림)·ceil(올림), unit: 1·10·100·1000(원), 원화가 아니면 0.01 */
  function roundMoney(v, mode, unit) {
    if (v == null || isNaN(v)) return v;
    unit = Number(unit) > 0 ? Number(unit) : 1;
    var q = Math.round(v / unit * 1e6) / 1e6;     // 0.1+0.2 같은 부동소수 찌꺼기 제거
    var r = mode === 'floor' ? Math.floor(q) : mode === 'ceil' ? Math.ceil(q) : (q < 0 ? -Math.round(-q) : Math.round(q));
    return Math.round(r * unit * 100) / 100;
  }
  function unitFor(cur, S) { return cur === 'KRW' ? (S.roundUnit || 1) : 0.01; }

  var FLAG_TEXT = {
    noAmount: '합의금액이 없고 청구금액·부담 비율로도 계산할 수 없어 문서에서 뺍니다.',
    negative: '합의금액이 음수입니다.',
    over: '합의금액이 청구금액보다 큽니다.',
    ratioRange: '부담 비율이 0~100% 밖입니다.',
    mismatch: '부담 비율로 계산한 금액과 합의금액이 다릅니다.',
    zero: '합의금액이 0 입니다(면책·기각이면 비고에 적어 두세요).',
    dupClaim: '같은 협력사에 같은 클레임 번호가 두 번 이상 있습니다.',
    noClaimNo: '클레임 번호가 비어 있습니다.',
    computed: '합의금액 칸이 비어 청구금액 × 부담 비율로 계산했습니다.',
    currencyMix: '한 협력사 안에 통화가 섞여 있습니다. 통화별로 나눠 주세요.'
  };
  var FLAG_LEVEL = { noAmount: 'bad', negative: 'bad', over: 'bad', ratioRange: 'bad', currencyMix: 'bad', mismatch: 'warn', zero: 'warn', dupClaim: 'warn', noClaimNo: 'warn', computed: 'info' };
  function flag(code, extra) { return { code: code, level: FLAG_LEVEL[code], msg: FLAG_TEXT[code] + (extra ? ' ' + extra : '') }; }

  /* 행마다 합의금액을 정하고(합의금액 칸 우선, 없으면 청구금액 × 부담 비율) 검사 표시를 붙입니다 — 원본 행은 바꾸지 않습니다 */
  function evaluateRows(rows, S) {
    S = cleanSettings(S);
    var seen = {};
    var out = (rows || []).map(function (r0) {
      var r = {}; Object.keys(r0).forEach(function (k) { r[k] = r0[k]; });
      r.currency = r.currency || S.currency;
      r.flags = [];
      var unit = unitFor(r.currency, S);
      var calc = r.claimAmt != null && r.ratio != null ? roundMoney(r.claimAmt * r.ratio / 100, S.roundMode, unit) : null;
      if (r.agreedRaw != null) r.agreed = roundMoney(r.agreedRaw, S.roundMode, unit);
      else if (calc != null) { r.agreed = calc; r.flags.push(flag('computed', '(' + fmtMoney(r.claimAmt, r.currency) + ' × ' + fmtPct(r.ratio) + ')')); }
      else r.agreed = null;
      if (r.agreed == null) r.flags.push(flag('noAmount'));
      else {
        if (r.agreed < 0) r.flags.push(flag('negative'));
        if (r.claimAmt != null && r.agreed > r.claimAmt) r.flags.push(flag('over', '(청구 ' + fmtMoney(r.claimAmt, r.currency) + ' < 합의 ' + fmtMoney(r.agreed, r.currency) + ')'));
        if (r.agreed === 0) r.flags.push(flag('zero'));
        var tol = Math.max(unit, S.tolerance);
        if (r.agreedRaw != null && calc != null && Math.abs(calc - r.agreed) > tol) r.flags.push(flag('mismatch', '(계산 ' + fmtMoney(calc, r.currency) + ' · 합의 ' + fmtMoney(r.agreed, r.currency) + ')'));
      }
      if (r.ratio != null && (r.ratio < 0 || r.ratio > 100)) r.flags.push(flag('ratioRange', '(' + fmtPct(r.ratio) + ')'));
      if (!r.claimNo) r.flags.push(flag('noClaimNo'));
      else {
        var k = supplierKey(r) + '|' + norm(r.claimNo);
        if (seen[k]) { r.flags.push(flag('dupClaim', '(' + r.claimNo + ')')); if (!seen[k].dupMarked) { seen[k].flags.push(flag('dupClaim', '(' + r.claimNo + ')')); seen[k].dupMarked = true; } }
        else seen[k] = r;
      }
      r.included = r.agreed != null;
      return r;
    });
    out.forEach(function (r) { delete r.dupMarked; r.level = r.flags.some(function (f) { return f.level === 'bad'; }) ? 'bad' : r.flags.some(function (f) { return f.level === 'warn'; }) ? 'warn' : 'ok'; });
    return out;
  }

  /* ── 4. 협력사별 묶기 ─────────────────────────────────────── */
  function supplierKey(r) { return r.supplierCode ? 'c:' + norm(r.supplierCode) : 'n:' + norm(r.supplier); }
  function groupBySupplier(rows, S) {
    S = cleanSettings(S);
    var ev = evaluateRows(rows, S), by = {}, order = [];
    ev.forEach(function (r) {
      var k = supplierKey(r);
      if (!by[k]) { by[k] = { key: k, id: 'g' + hash(k), supplier: r.supplier, supplierCode: r.supplierCode, rows: [], excluded: [], currencies: [] }; order.push(k); }
      var g = by[k];
      if (!g.supplierCode && r.supplierCode) g.supplierCode = r.supplierCode;
      (r.included ? g.rows : g.excluded).push(r);
      if (r.included && g.currencies.indexOf(r.currency) < 0) g.currencies.push(r.currency);
    });
    return order.map(function (k) { return finishGroup(by[k], S); });
  }
  function finishGroup(g, S) {
    g.currency = g.currencies[0] || S.currency;
    g.flags = [];
    if (g.currencies.length > 1) g.flags.push(flag('currencyMix', '(' + g.currencies.join(', ') + ')'));
    var unit = unitFor(g.currency, S);
    g.count = g.rows.length;
    g.qty = g.rows.reduce(function (a, r) { return a + (r.qty || 0); }, 0);
    g.claimSum = round2(g.rows.reduce(function (a, r) { return a + (r.claimAmt || 0); }, 0));
    g.agreedSum = round2(g.rows.reduce(function (a, r) { return a + r.agreed; }, 0));
    g.vat = S.vat === 'add10' ? roundMoney(g.agreedSum * 0.1, S.roundMode, unit) : 0;
    g.total = round2(g.agreedSum + g.vat);
    g.ratio = g.claimSum > 0 ? Math.round(g.agreedSum / g.claimSum * 1000) / 10 : null;
    var ds = g.rows.map(function (r) { return r.agreeDate; }).filter(Boolean).sort();
    g.agreeDate = ds.length ? ds[ds.length - 1] : '';
    g.bad = g.flags.length + g.rows.concat(g.excluded).filter(function (r) { return r.level === 'bad'; }).length;
    g.warn = g.rows.filter(function (r) { return r.level === 'warn'; }).length;
    g.sig = hash(g.currency + '|' + g.total + '|' + g.rows.map(function (r) { return [r.claimNo, r.partNo, r.agreed].join('/'); }).join(';'));
    return g;
  }
  function round2(v) { return Math.round(v * 100) / 100; }

  /* ── 5. 금액 표기 ─────────────────────────────────────────── */
  function fmtMoney(n, cur) {
    if (n == null || isNaN(n)) return '';
    var dec = cur && cur !== 'KRW' ? 2 : 0;
    var neg = n < 0, v = Math.abs(n).toFixed(dec), p = v.split('.');
    p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return (neg ? '-' : '') + p.join('.');
  }
  function fmtPct(n) { return n == null || isNaN(n) ? '' : (Math.round(n * 10) / 10) + '%'; }
  /* 숫자 → 한글 금액. 1,200,000 → 일백이십만. 고치기 어렵게 「일」을 빼지 않습니다(일십·일백·일천) */
  var KD = ['', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구'], KS = ['', '십', '백', '천'], KB = ['', '만', '억', '조', '경'];
  function koreanNumber(n) {
    n = Math.floor(Math.abs(Number(n) || 0));
    if (n === 0) return '영';
    var s = String(n), out = '', groups = [];
    while (s.length) { groups.unshift(s.slice(-4)); s = s.slice(0, -4); }
    groups.forEach(function (g, gi) {
      var big = KB[groups.length - 1 - gi], part = '';
      g = pad(g, 4);
      for (var i = 0; i < 4; i++) { var d = +g[i]; if (d) part += KD[d] + KS[3 - i]; }
      if (part) out += part + big;
    });
    return out;
  }
  function amountWords(n, cur) {
    if (n == null || isNaN(n)) return '';
    if (!cur || cur === 'KRW') return '금 ' + (n < 0 ? '마이너스 ' : '') + koreanNumber(n) + '원정';
    return cur + ' ' + fmtMoney(n, cur);
  }
  function amountLine(n, cur) {
    if (!cur || cur === 'KRW') return amountWords(n, cur) + '(₩' + fmtMoney(n, cur) + ')';
    return cur + ' ' + fmtMoney(n, cur);
  }

  /* ── 6. 문서 번호·발행 이력 ───────────────────────────────── */
  var DOC_KINDS = { ag: '합의서', inv: '인보이스' };
  /* 형식 낱말: {접두} {YYYY} {YY} {MM} {DD} {N…}(일련, N 개수만큼 0 채움) */
  function formatDocNo(fmt, o) {
    var y = o.date.slice(0, 4), mo = o.date.slice(5, 7), d = o.date.slice(8, 10);
    return String(fmt || '{접두}-{YYYY}-{MM}-{NNN}')
      .replace(/\{접두\}/g, o.prefix).replace(/\{YYYY\}/g, y).replace(/\{YY\}/g, y.slice(2)).replace(/\{MM\}/g, mo).replace(/\{DD\}/g, d)
      .replace(/\{(N+)\}/g, function (_, ns) { return pad(o.seq, ns.length); });
  }
  /* 일련은 문서 종류마다, 발행일의 달마다 1 부터 */
  function nextSeq(history, kind, ym) {
    var mx = 0;
    (history || []).forEach(function (h) { if (h.kind === kind && h.ym === ym && h.seq > mx) mx = h.seq; });
    return mx + 1;
  }
  /* 같은 협력사·같은 내용(행·금액이 같음)을 다시 내면 번호를 그대로 씁니다(재발행).
     내용이 바뀌었으면 새 번호를 주고 이전 번호는 「대체됨」으로 둡니다. 원본 history 는 바꾸지 않고 새 배열을 돌려줍니다. */
  function issueDoc(history, kind, g, issueDate, S, now) {
    S = cleanSettings(S);
    var list = (history || []).map(function (h) { var c = {}; Object.keys(h).forEach(function (k) { c[k] = h[k]; }); return c; });
    var same = list.filter(function (h) { return h.kind === kind && h.supplierKey === g.key && h.sig === g.sig && h.status === 'issued'; })[0];
    if (same) { same.times = (same.times || 1) + 1; same.lastAt = now || ''; return { history: list, entry: same, reused: true }; }
    var ym = issueDate.slice(0, 7), seq = nextSeq(list, kind, ym);
    var e = {
      id: kind + '-' + hash(g.key + '|' + g.sig + '|' + ym + '|' + seq), kind: kind, docNo: formatDocNo(S.docFormat, { prefix: kind === 'ag' ? S.prefixAg : S.prefixInv, date: issueDate, seq: seq }),
      seq: seq, ym: ym, issueDate: issueDate, supplier: g.supplier, supplierCode: g.supplierCode, supplierKey: g.key, sig: g.sig,
      currency: g.currency, count: g.count, total: g.total, status: 'issued', times: 1, at: now || '', lastAt: now || '', replaces: ''
    };
    list.forEach(function (h) { if (h.kind === kind && h.supplierKey === g.key && h.status === 'issued') { h.status = 'replaced'; e.replaces = h.docNo; } });
    list.push(e);
    return { history: list, entry: e, reused: false };
  }
  /* 미리 보기에 쓸 번호(이력에 쓰지 않음) */
  function previewDocNo(history, kind, g, issueDate, S) { return issueDoc(history, kind, g, issueDate, S).entry.docNo; }
  function latestDocNo(history, kind, g) {
    var l = (history || []).filter(function (h) { return h.kind === kind && h.supplierKey === g.key && h.status === 'issued'; });
    return l.length ? l[l.length - 1].docNo : '';
  }

  /* ── 7. 자리표시(키) ─────────────────────────────────────── */
  /* 양식 칸에 {{키}} 를 적으면 그 값으로 바뀝니다. {{행.키}} 가 있는 줄은 클레임 수만큼 반복됩니다.
     type: money(금액 숫자) · int(정수) · pct(비율 글자) · text · date(YYYY-MM-DD 글자) */
  var KEYS = [
    { key: '문서번호', type: 'text', desc: '자동 채번 (예: AG-2026-10-001)' },
    { key: '문서명', type: 'text', desc: '합의서 / 인보이스' },
    { key: '관련문서번호', type: 'text', desc: '인보이스에서는 같은 협력사의 합의서 번호, 합의서에서는 인보이스 번호' },
    { key: '발행일', type: 'date', desc: 'YYYY-MM-DD' },
    { key: '발행일_한글', type: 'text', desc: '2026년 10월 21일' },
    { key: '합의일', type: 'date', desc: '합의 결과의 합의일(가장 늦은 날). 없으면 발행일' },
    { key: '합의일_한글', type: 'text', desc: '2026년 10월 15일' },
    { key: '지급기한', type: 'date', desc: '발행일 + 설정의 지급 기한 일수' },
    { key: '협력사명', type: 'text' },
    { key: '협력사코드', type: 'text' },
    { key: '통화', type: 'text', desc: 'KRW · USD …' },
    { key: '클레임건수', type: 'int', desc: '문서에 들어간 행 수' },
    { key: '수량합계', type: 'int' },
    { key: '청구합계', type: 'money' },
    { key: '합의합계', type: 'money', desc: '부가세 전' },
    { key: '부가세', type: 'money', desc: '설정이 「10% 별도」일 때만, 아니면 0' },
    { key: '총액', type: 'money', desc: '합의합계 + 부가세' },
    { key: '총액_한글', type: 'text', desc: '금 일백이십만원정 (원화가 아니면 USD 1,234.50)' },
    { key: '총액_표기', type: 'text', desc: '금 일백이십만원정(₩1,200,000)' },
    { key: '부담비율', type: 'pct', desc: '합의합계 ÷ 청구합계' },
    { key: '발행처', type: 'text', desc: '설정 — 우리 회사 이름' },
    { key: '발행부서', type: 'text', desc: '설정' },
    { key: '담당자', type: 'text', desc: '설정' },
    { key: '연락처', type: 'text', desc: '설정' },
    { key: '입금계좌', type: 'text', desc: '설정 — 인보이스 입금 계좌' },
    { key: '행.번호', type: 'int', row: true },
    { key: '행.클레임번호', type: 'text', row: true },
    { key: '행.발생일', type: 'date', row: true },
    { key: '행.기종', type: 'text', row: true },
    { key: '행.품번', type: 'text', row: true },
    { key: '행.품명', type: 'text', row: true },
    { key: '행.수량', type: 'int', row: true },
    { key: '행.청구금액', type: 'money', row: true },
    { key: '행.부담비율', type: 'pct', row: true },
    { key: '행.합의금액', type: 'money', row: true },
    { key: '행.비고', type: 'text', row: true }
  ];
  var KEY_BY = {}; KEYS.forEach(function (k) { KEY_BY[k.key] = k; });
  var PH_RE = /\{\{\s*([^{}]+?)\s*\}\}/g;
  function placeholdersIn(text) { var o = [], m; PH_RE.lastIndex = 0; while ((m = PH_RE.exec(String(text || '')))) o.push(m[1]); return o; }
  function singleKey(text) { var m = /^\s*\{\{\s*([^{}]+?)\s*\}\}\s*$/.exec(String(text || '')); return m ? m[1] : null; }
  function isRowKey(k) { return /^행\./.test(k); }

  /* 문서 하나에 들어갈 값. 숫자는 숫자로 두고(엑셀 칸에 숫자로 씀), 글자 칸 안에서는 fmtValue 로 표기합니다 */
  function buildContext(g, kind, S, meta) {
    S = cleanSettings(S); meta = meta || {};
    var issue = meta.issueDate || localToday(), cur = g.currency, agree = g.agreeDate || issue;
    var f = {
      '문서번호': meta.docNo || '', '문서명': DOC_KINDS[kind] || '', '관련문서번호': meta.relatedNo || '',
      '발행일': issue, '발행일_한글': koDate(issue), '합의일': agree, '합의일_한글': koDate(agree), '지급기한': addDays(issue, S.payDays),
      '협력사명': g.supplier, '협력사코드': g.supplierCode || '', '통화': cur,
      '클레임건수': g.count, '수량합계': g.qty, '청구합계': g.claimSum, '합의합계': g.agreedSum, '부가세': g.vat, '총액': g.total,
      '총액_한글': amountWords(g.total, cur), '총액_표기': amountLine(g.total, cur), '부담비율': g.ratio,
      '발행처': S.issuer, '발행부서': S.dept, '담당자': S.person, '연락처': S.contact, '입금계좌': S.bank
    };
    var rows = g.rows.map(function (r, i) {
      return { '행.번호': i + 1, '행.클레임번호': r.claimNo, '행.발생일': r.date, '행.기종': r.model, '행.품번': r.partNo, '행.품명': r.partName,
        '행.수량': r.qty, '행.청구금액': r.claimAmt, '행.부담비율': r.ratio, '행.합의금액': r.agreed, '행.비고': r.note };
    });
    return { fields: f, rows: rows, currency: cur };
  }
  function rawValue(key, ctx, row) {
    var src = isRowKey(key) ? (row || {}) : ctx.fields;
    return Object.prototype.hasOwnProperty.call(src, key) ? src[key] : undefined;
  }
  function fmtValue(key, v, cur) {
    if (v == null || v === '') return '';
    var t = (KEY_BY[key] || {}).type;
    if (t === 'money') return fmtMoney(v, cur);
    if (t === 'pct') return fmtPct(v);
    if (t === 'int') return String(v);
    return String(v);
  }
  /* 칸 하나 채우기 — 칸 전체가 {{키}} 하나이고 값이 숫자(금액·정수)면 숫자로, 아니면 글자로. 모르는 키는 그대로 둡니다 */
  function fillCell(text, ctx, row) {
    var sk = singleKey(text);
    if (sk && KEY_BY[sk]) {
      var t = KEY_BY[sk].type, v = rawValue(sk, ctx, row);
      if ((t === 'money' || t === 'int') && typeof v === 'number') return { type: 'n', value: v, kind: t };
      return { type: 's', value: fmtValue(sk, v, ctx.currency) };
    }
    return { type: 's', value: fillText(text, ctx, row) };
  }
  function fillText(text, ctx, row) {
    return String(text == null ? '' : text).replace(PH_RE, function (all, k) {
      if (!KEY_BY[k]) return all;
      return fmtValue(k, rawValue(k, ctx, row), ctx.currency);
    });
  }
  function unknownKeys(keys) { return keys.filter(function (k) { return !KEY_BY[k]; }); }

  /* ── 8. 설정·저장 모양 ───────────────────────────────────── */
  function defaultSettings() {
    return {
      issuer: '', dept: '품질보증팀 보상파트', person: '', contact: '', bank: '',
      currency: 'KRW', vat: 'none', roundMode: 'round', roundUnit: 1, tolerance: 10, payDays: 30,
      docFormat: '{접두}-{YYYY}-{MM}-{NNN}', prefixAg: 'AG', prefixInv: 'INV',
      agClauses: [
        '을은 위 클레임이 을이 납품한 부품의 품질 문제로 발생하였음을 확인하고, 합의 금액을 갑에게 보상한다.',
        '을은 갑이 발행한 인보이스에 따라 {{지급기한}}까지 합의 금액을 지급한다.',
        '이 합의로 위 클레임에 관한 보상은 끝난 것으로 하며, 같은 원인으로 새로 생기는 클레임은 따로 협의한다.'
      ].join('\n'),
      invNote: '위 금액을 지급 기한까지 아래 계좌로 보내 주십시오.'
    };
  }
  function intOr(v, d, lo, hi) { var n = Number(v); return v === '' || v == null || !isFinite(n) ? d : Math.max(lo, Math.min(hi, Math.round(n))); }
  function cleanSettings(p) {
    var d = defaultSettings(); p = p || {};
    var s = {};
    ['issuer', 'dept', 'person', 'contact', 'bank', 'agClauses', 'invNote', 'docFormat', 'prefixAg', 'prefixInv'].forEach(function (k) { s[k] = p[k] == null ? d[k] : String(p[k]); });
    if (!/\{N+\}/.test(s.docFormat)) s.docFormat = d.docFormat;
    if (!s.prefixAg.trim()) s.prefixAg = d.prefixAg;
    if (!s.prefixInv.trim()) s.prefixInv = d.prefixInv;
    s.currency = /^[A-Z]{3}$/.test(String(p.currency || '').toUpperCase()) ? String(p.currency).toUpperCase() : d.currency;
    s.vat = p.vat === 'add10' ? 'add10' : 'none';
    s.roundMode = ['round', 'floor', 'ceil'].indexOf(p.roundMode) >= 0 ? p.roundMode : 'round';
    s.roundUnit = [1, 10, 100, 1000].indexOf(Number(p.roundUnit)) >= 0 ? Number(p.roundUnit) : 1;
    s.tolerance = intOr(p.tolerance, d.tolerance, 0, 1000000);
    s.payDays = intOr(p.payDays, d.payDays, 0, 365);
    return s;
  }
  function emptyDb() { return { version: 1, rows: [], files: [], settings: defaultSettings(), history: [], templates: { ag: null, inv: null }, issueDate: '', sample: false }; }
  function restoreDb(p) {
    var db = emptyDb();
    if (!p || typeof p !== 'object') return db;
    db.rows = (Array.isArray(p.rows) ? p.rows : []).filter(function (r) { return r && typeof r === 'object' && str(r.supplier); });
    db.files = Array.isArray(p.files) ? p.files.filter(function (f) { return f && f.name; }) : [];
    db.settings = cleanSettings(p.settings);
    db.history = Array.isArray(p.history) ? p.history.filter(function (h) { return h && h.docNo && DOC_KINDS[h.kind]; }) : [];
    if (p.templates && typeof p.templates === 'object') ['ag', 'inv'].forEach(function (k) { var t = p.templates[k]; if (t && t.b64 && t.name) db.templates[k] = t; });
    db.issueDate = toISO(p.issueDate) || '';
    db.sample = !!p.sample;
    return db;
  }

  /* 검토표·이력 내보내기 표 */
  var REVIEW_HEAD = ['행', '협력사명', '협력사 코드', '클레임 번호', '발생일', '기종', '품번', '품명', '수량', '청구금액', '부담 비율', '합의금액', '통화', '판정', '검사 내용'];
  function reviewRows(rows, S) {
    return evaluateRows(rows, S).map(function (r) {
      return [r.line, r.supplier, r.supplierCode, r.claimNo, r.date, r.model, r.partNo, r.partName, r.qty, r.claimAmt, r.ratio == null ? '' : fmtPct(r.ratio), r.agreed, r.currency,
        r.level === 'bad' ? '확인 필요' : r.level === 'warn' ? '주의' : '정상', r.flags.map(function (f) { return f.msg; }).join(' / ')];
    });
  }
  var HISTORY_HEAD = ['문서 번호', '종류', '발행일', '협력사', '협력사 코드', '건수', '총액', '통화', '상태', '대체한 번호', '낸 횟수'];
  function historyRows(history) {
    return (history || []).map(function (h) {
      return [h.docNo, DOC_KINDS[h.kind], h.issueDate, h.supplier, h.supplierCode, h.count, h.total, h.currency, h.status === 'replaced' ? '대체됨' : h.status === 'void' ? '취소' : '발행', h.replaces || '', h.times || 1];
    });
  }

  return {
    VERSION: VERSION,
    str: str, norm: norm, toISO: toISO, addDays: addDays, localToday: localToday, koDate: koDate, colName: colName, toNum: toNum, toPct: toPct, hash: hash,
    FIELDS: FIELDS, matchField: matchField, detectColumns: detectColumns, pickSheet: pickSheet, parseSheet: parseSheet,
    roundMoney: roundMoney, evaluateRows: evaluateRows, supplierKey: supplierKey, groupBySupplier: groupBySupplier, FLAG_TEXT: FLAG_TEXT,
    fmtMoney: fmtMoney, fmtPct: fmtPct, koreanNumber: koreanNumber, amountWords: amountWords, amountLine: amountLine,
    DOC_KINDS: DOC_KINDS, formatDocNo: formatDocNo, nextSeq: nextSeq, issueDoc: issueDoc, previewDocNo: previewDocNo, latestDocNo: latestDocNo,
    KEYS: KEYS, KEY_BY: KEY_BY, placeholdersIn: placeholdersIn, singleKey: singleKey, isRowKey: isRowKey, unknownKeys: unknownKeys,
    buildContext: buildContext, fillCell: fillCell, fillText: fillText, fmtValue: fmtValue,
    defaultSettings: defaultSettings, cleanSettings: cleanSettings, emptyDb: emptyDb, restoreDb: restoreDb,
    REVIEW_HEAD: REVIEW_HEAD, reviewRows: reviewRows, HISTORY_HEAD: HISTORY_HEAD, historyRows: historyRows
  };
});
