/* 엑셀 양식 채우기 — 사용자가 올린 .xlsx 양식의 칸에 적힌 {{키}} 를 값으로 바꿉니다 (브라우저·node 공용)

   왜 이렇게 하나: SheetJS 무료판(Community)은 읽은 파일을 다시 쓸 때 글꼴·테두리·채우기 색을 버립니다.
   그래서 SheetJS 로 「읽고 → 고치고 → 쓰는」 대신, xlsx 파일(zip) 안의 시트 XML 을 직접 고쳐
   양식의 서식(글꼴·테두리·색·병합·열 너비·인쇄 설정·그림·직인 이미지)을 그대로 둡니다.
   zip 을 여닫는 것만 SheetJS 에 들어 있는 CFB 를 씁니다.

   규칙
   - 칸 전체가 {{키}} 하나이고 값이 금액·정수면 숫자 칸으로 씁니다(양식의 표시 형식 #,##0 이 그대로 적용).
   - 「합의 금액: {{총액_표기}}」처럼 글과 섞이면 글자 칸으로 씁니다.
   - {{행.키}} 가 있는 줄(시트마다 한 줄)은 클레임 수만큼 반복합니다. 그 아래 빈 줄(양식에 미리 그려 둔 표 줄)을 먼저 쓰고,
     모자라면 줄을 끼워 넣고 아래 줄·병합·수식·그림 위치·인쇄 영역을 함께 내립니다.
     표 끝에 닿은 합계 수식(=SUM(I12:I21))은 끼운 줄까지 범위를 늘립니다.
   - 공유 수식은 일반 수식으로 풀어 둡니다. 계산값은 엑셀이 열 때 다시 계산하게 합니다(fullCalcOnLoad).
   한계: 도형·글상자 안의 글자, 열 전체·행 전체 참조(3:5), 다른 시트에서 이 시트를 가리키는 수식은 고치지 않습니다. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('../vendor/xlsx.full.min.js'), require('./logic.js'));
  else root.CLTemplate = factory(root.XLSX, root.CLLogic);
})(typeof self !== 'undefined' ? self : this, function (XLSX, L) {
  'use strict';
  var enc = new TextEncoder(), dec = new TextDecoder('utf-8');

  /* ── XML 글자 ─────────────────────────────── */
  function xmlDecode(s) {
    return String(s).replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, function (_, e) {
      if (e[0] === '#') return String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
      return { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" }[e];
    });
  }
  function xmlEncode(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function attr(a, name) { var m = new RegExp('(?:^|\\s)' + name.replace(':', '\\:') + '="([^"]*)"').exec(a || ''); return m ? xmlDecode(m[1]) : null; }
  function dropAttr(a, name) { return (a || '').replace(new RegExp('\\s' + name.replace(':', '\\:') + '="[^"]*"', 'g'), ''); }

  /* ── 칸 주소 ─────────────────────────────── */
  function colIndex(letters) { var n = 0; for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64); return n - 1; }
  function colName(i) { var s = ''; do { s = String.fromCharCode(65 + (i % 26)) + s; i = Math.floor(i / 26) - 1; } while (i >= 0); return s; }
  function splitRef(ref) { var m = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(ref); return m ? { c: colIndex(m[1]), r: +m[2] } : null; }

  /* 수식 안의 칸 주소를 바꿉니다. fn(ref, part) → 새 ref.  ref = {c, r, cAbs, rAbs}, part = 'single'|'start'|'end'
     따옴표 글자와 다른 시트 주소(Sheet2!A1)는 건드리지 않습니다(opts.sheet 와 같은 이름이면 고침). */
  var REF = "(\\$?)([A-Z]{1,3})(\\$?)(\\d{1,7})";
  var TOKEN = new RegExp('("(?:[^"]|"")*")|((?:\'(?:[^\']|\'\')+\'|[A-Za-z_\\u3131-\\uD79D][\\w.\\u3131-\\uD79D]*)!)?' + REF + '(?::' + REF + ')?(?![\\w(])', 'g');
  function mapRefs(formula, fn, opts) {
    opts = opts || {};
    return String(formula).replace(TOKEN, function (all, quoted, sheet, c1a, c1, r1a, r1, c2a, c2, r2a, r2, off, whole) {
      if (quoted) return all;
      var prev = off > 0 ? whole[off - 1] : '';
      if (/[A-Za-z0-9_.$]/.test(prev)) return all;
      if (sheet) {
        var nm = sheet.slice(0, -1).replace(/^'|'$/g, '').replace(/''/g, "'");
        if (!opts.sheet || nm !== opts.sheet) return all;
      }
      function out(ref) { return (ref.cAbs ? '$' : '') + colName(ref.c) + (ref.rAbs ? '$' : '') + ref.r; }
      var a = { c: colIndex(c1), r: +r1, cAbs: !!c1a, rAbs: !!r1a };
      if (c2 == null) return (sheet || '') + out(fn(a, 'single'));
      var b = { c: colIndex(c2), r: +r2, cAbs: !!c2a, rAbs: !!r2a };
      var na = fn(a, 'start', b), nb = fn(b, 'end', a);
      return (sheet || '') + out(na) + ':' + out(nb);
    });
  }
  /* 상대 이동(줄을 복사할 때) — $ 가 없는 쪽만 움직임 */
  function relShift(formula, dRow, dCol, opts) {
    return mapRefs(formula, function (ref) {
      return { c: ref.cAbs ? ref.c : Math.max(0, ref.c + (dCol || 0)), r: ref.rAbs ? ref.r : Math.max(1, ref.r + dRow), cAbs: ref.cAbs, rAbs: ref.rAbs };
    }, opts);
  }
  /* 줄 끼워 넣기 — at 번째 줄부터 ins 줄 아래로. 범위 끝이 바로 위 줄(at-1)이고 시작이 expandFrom 이하면 범위를 늘립니다 */
  function insertShift(formula, at, ins, expandFrom, opts) {
    return mapRefs(formula, function (ref, part, other) {
      var r = ref.r;
      if (r >= at) r += ins;
      else if (part === 'end' && r === at - 1 && expandFrom != null && other.r <= expandFrom) r += ins;
      return { c: ref.c, r: r, cAbs: ref.cAbs, rAbs: ref.rAbs };
    }, opts);
  }
  function shiftSqref(sq, at, ins, expandFrom) { return String(sq).split(/\s+/).map(function (p) { return p ? insertShift(p, at, ins, expandFrom) : p; }).join(' '); }

  /* ── zip(xlsx) 열기 ─────────────────────────── */
  function toU8(c) { return c instanceof Uint8Array ? c : new Uint8Array(c || []); }
  function openZip(data) {
    var cfb;
    try { cfb = XLSX.CFB.read(toU8(data), { type: 'array' }); } catch (e) { throw new Error('엑셀 파일로 열지 못했습니다(' + e.message + ').'); }
    function idx(path) { for (var i = 0; i < cfb.FullPaths.length; i++) if (cfb.FullPaths[i].replace(/^[^/]*\//, '') === path) return i; return -1; }
    var z = {
      has: function (p) { return idx(p) >= 0; },
      get: function (p) { var i = idx(p); return i < 0 ? null : dec.decode(toU8(cfb.FileIndex[i].content)); },
      set: function (p, s) {
        var i = idx(p), b = enc.encode(s);
        if (i < 0) XLSX.CFB.utils.cfb_add(cfb, p, b);
        else { cfb.FileIndex[i].content = b; cfb.FileIndex[i].size = b.length; }
      },
      del: function (p) { var i = idx(p); if (i >= 0) { cfb.FullPaths.splice(i, 1); cfb.FileIndex.splice(i, 1); } },
      write: function () { return toU8(XLSX.CFB.write(cfb, { fileType: 'zip', type: 'array' })); }
    };
    if (!z.has('xl/workbook.xml')) throw new Error('xlsx 형식이 아닙니다. 엑셀에서 「Excel 통합 문서(*.xlsx)」로 저장해 넣어 주세요(.xls · 문서보안(DRM) 파일은 열 수 없습니다).');
    return z;
  }
  function resolve(base, target) {
    var parts = (target[0] === '/' ? target.slice(1) : base + target).split('/'), out = [];
    parts.forEach(function (p) { if (p === '..') out.pop(); else if (p && p !== '.') out.push(p); });
    return out.join('/');
  }
  function rels(z, path) {
    var m = {}, x = z.get(path);
    if (x) x.replace(/<Relationship\b([^>]*?)\/?>/g, function (_, a) { m[attr(a, 'Id')] = { target: attr(a, 'Target'), type: attr(a, 'Type') || '', mode: attr(a, 'TargetMode') }; return _; });
    return m;
  }
  function sheetList(z) {
    var wb = z.get('xl/workbook.xml'), rm = rels(z, 'xl/_rels/workbook.xml.rels'), out = [];
    wb.replace(/<sheet\b([^>]*?)\/?>/g, function (_, a) {
      var rel = rm[attr(a, 'r:id')];
      if (rel && /worksheet$/.test(rel.type)) out.push({ name: attr(a, 'name'), path: resolve('xl/', rel.target), index: out.length });
      return _;
    });
    return out;
  }
  function readSst(z) {
    var x = z.get('xl/sharedStrings.xml'), out = [];
    if (!x) return out;
    x.replace(/<si>([\s\S]*?)<\/si>|<si\/>/g, function (_, inner) { out.push(inner ? textOf(inner) : ''); return _; });
    return out;
  }
  function textOf(inner) {
    var s = '', clean = inner.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    clean.replace(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, function (_, t) { s += xmlDecode(t); return _; });
    return s;
  }

  /* ── 시트 XML 읽기·쓰기 ─────────────────────── */
  function parseSheetXml(xml, sst) {
    var m = /<sheetData\s*\/>|<sheetData>([\s\S]*?)<\/sheetData>/.exec(xml);
    if (!m) throw new Error('시트 내용을 찾지 못했습니다.');
    var rows = [], lastR = 0;
    (m[1] || '').replace(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g, function (_, a, inner) {
      var r = +attr(a, 'r') || lastR + 1; lastR = r;
      var row = { r: r, attrs: dropAttr(a, 'r'), cells: [] }, lastC = -1;
      (inner || '').replace(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, function (__, ca, ci) {
        var ref = attr(ca, 'r'), p = ref ? splitRef(ref) : null, c = p ? p.c : lastC + 1; lastC = c;
        row.cells.push(readCell(c, dropAttr(ca, 'r'), ci || '', sst));
        return __;
      });
      rows.push(row);
      return _;
    });
    return { before: xml.slice(0, m.index), after: xml.slice(m.index + m[0].length), rows: rows };
  }
  function readCell(c, a, inner, sst) {
    var t = attr(a, 't'), cell = { c: c, attrs: a, inner: inner, text: null, f: null, fAttrs: '' }, vm = /<v>([\s\S]*?)<\/v>/.exec(inner), fm = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(inner);
    if (fm) { cell.f = fm[2] != null ? xmlDecode(fm[2]) : ''; cell.fAttrs = fm[1] || ''; }
    if (t === 's' && vm) cell.text = sst[+vm[1]] == null ? '' : sst[+vm[1]];
    else if (t === 'inlineStr') cell.text = textOf(inner);
    else if (t === 'str' && vm && !fm) cell.text = xmlDecode(vm[1]);
    cell.hasValue = !!(fm || (vm && vm[1] !== '') || (cell.text != null && cell.text !== ''));
    return cell;
  }
  function cellXml(r, cell) {
    var a = cell.attrs || '', ref = colName(cell.c) + r;
    return '<c r="' + ref + '"' + a + (cell.inner ? '>' + cell.inner + '</c>' : '/>');
  }
  function sheetXml(p) {
    p.rows.sort(function (a, b) { return a.r - b.r; });
    var body = p.rows.map(function (row) {
      row.cells.sort(function (a, b) { return a.c - b.c; });
      return '<row r="' + row.r + '"' + row.attrs + (row.cells.length ? '>' + row.cells.map(function (c) { return cellXml(row.r, c); }).join('') + '</row>' : '/>');
    }).join('');
    return p.before + (body ? '<sheetData>' + body + '</sheetData>' : '<sheetData/>') + p.after;
  }
  /* 칸에 값을 씁니다 — 스타일(s)은 두고 형식(t)·값만 바꿈 */
  function setValue(cell, res) {
    var a = dropAttr(dropAttr(cell.attrs, 't'), 'cm');
    if (res.type === 'n' && isFinite(res.value)) { cell.attrs = a; cell.inner = '<v>' + res.value + '</v>'; }
    else if (res.value === '' || res.value == null) { cell.attrs = a; cell.inner = ''; }
    else { cell.attrs = a + ' t="inlineStr"'; cell.inner = '<is><t xml:space="preserve">' + xmlEncode(res.value) + '</t></is>'; }
    cell.text = res.type === 'n' ? null : res.value; cell.f = null; cell.hasValue = res.value !== '' && res.value != null;
  }
  function setFormula(cell, f) { cell.attrs = dropAttr(dropAttr(cell.attrs, 't'), 'cm'); cell.inner = '<f>' + xmlEncode(f) + '</f>'; cell.f = f; cell.text = null; cell.hasValue = true; }
  function clone(o) { var c = {}; Object.keys(o).forEach(function (k) { c[k] = o[k]; }); return c; }

  /* 공유 수식(t="shared")을 칸마다의 일반 수식으로 풂 */
  function unshare(rows, sheetName) {
    var masters = {};
    rows.forEach(function (row) { row.cells.forEach(function (c) { if (c.f && attr(c.fAttrs, 't') === 'shared' && attr(c.fAttrs, 'ref')) masters[attr(c.fAttrs, 'si')] = { r: row.r, c: c.c, f: c.f }; }); });
    rows.forEach(function (row) {
      row.cells.forEach(function (c) {
        if (c.f == null || attr(c.fAttrs, 't') !== 'shared') return;
        var m = masters[attr(c.fAttrs, 'si')];
        var f = c.f !== '' ? c.f : m ? relShift(m.f, row.r - m.r, c.c - m.c, { sheet: sheetName }) : null;
        var v = /<v>[\s\S]*?<\/v>/.exec(c.inner);
        if (f == null) { c.inner = v ? v[0] : ''; c.f = null; return; }
        c.inner = '<f>' + xmlEncode(f) + '</f>' + (v ? v[0] : ''); c.f = f; c.fAttrs = '';
      });
    });
  }

  function hasPh(t) { return t != null && /\{\{[^{}]+\}\}/.test(t); }
  function rowKeyCell(c) { return hasPh(c.text) && L.placeholdersIn(c.text).some(L.isRowKey); }

  /* 시트 하나 채우기 → { xml, repeatRow, inserted, at, filled, unknown } */
  function fillSheet(xml, ctx, sst, sheetName, opts) {
    opts = opts || {};
    var p = parseSheetXml(xml, sst), rows = p.rows, byR = {};
    // 원화가 아니면 금액 칸을 소수 둘째 자리 표시 형식으로(양식의 #,##0 이 618.38 을 618 로 보이게 하므로)
    function put(cell, res) {
      setValue(cell, res);
      if (res.kind === 'money' && ctx.currency && ctx.currency !== 'KRW' && opts.decimalStyle) {
        var ns = opts.decimalStyle(attr(cell.attrs, 's') || '0');
        if (ns != null) cell.attrs = dropAttr(cell.attrs, 's') + ' s="' + ns + '"';
      }
    }
    rows.forEach(function (row) { byR[row.r] = row; });
    unshare(rows, sheetName);
    var unknown = [], filled = 0;
    function note(t) { L.unknownKeys(L.placeholdersIn(t)).forEach(function (k) { if (unknown.indexOf(k) < 0) unknown.push(k); }); }

    // 1) 반복 줄 찾기와 그 아래 비어 있는 표 줄 세기(채우기 전 원래 글자로 판단)
    var R = null;
    rows.forEach(function (row) { if (R == null && row.cells.some(rowKeyCell)) R = row.r; });
    var tpl = R != null ? byR[R] : null, avail = 0, phCols = [];
    if (tpl) {
      tpl.cells.forEach(function (c) { if (hasPh(c.text)) phCols.push(c.c); });
      var later = rows.filter(function (row) { return row.r > R; }).sort(function (a, b) { return a.r - b.r; });
      var nextUsed = null;
      later.some(function (row) {
        var used = row.cells.some(function (c) { return hasPh(c.text) || (phCols.indexOf(c.c) >= 0 && c.hasValue); });
        if (used) { nextUsed = row.r; return true; }
        return false;
      });
      avail = nextUsed == null ? Infinity : nextUsed - R - 1;
    }
    // 2) 반복 줄 밖의 칸 채우기
    rows.forEach(function (row) {
      if (row === tpl) return;
      row.cells.forEach(function (c) { if (hasPh(c.text)) { note(c.text); put(c, L.fillCell(c.text, ctx, null)); filled++; } });
    });
    // 3) 반복
    var n = tpl ? ctx.rows.length : 0, ins = 0, at = null, newMerges = [];
    var mergeList = [];
    p.after.replace(/<mergeCell\b[^>]*?ref="([^"]+)"/g, function (_, ref) { mergeList.push(ref); return _; });
    if (tpl) {
      var tplCells = tpl.cells.map(clone);
      tplCells.forEach(function (c) { note(c.text); });
      var need = Math.max(0, n - 1);
      if (need > avail) {
        ins = need - avail; at = R + avail + 1;
        rows.forEach(function (row) { if (row.r >= at) row.r += ins; });
        rows.forEach(function (row) { row.cells.forEach(function (c) { if (c.f) setFormula(c, insertShift(c.f, at, ins, R, { sheet: sheetName })); }); });
        tplCells.forEach(function (c) { if (c.f) c.f = insertShift(c.f, at, ins, R, { sheet: sheetName }); });
        byR = {}; rows.forEach(function (row) { byR[row.r] = row; });
      }
      // 반복 줄의 한 줄짜리 병합(예: 품명 C:D) — 새 줄에도 같은 병합
      var rowMerges = mergeList.map(function (ref) { var q = ref.split(':'); return q.length === 2 ? [splitRef(q[0]), splitRef(q[1])] : null; })
        .filter(function (q) { return q && q[0] && q[1] && q[0].r === R && q[1].r === R; });
      var taken = mergeList.map(function (ref) { var q = ref.split(':').map(splitRef); return q.length === 2 && q[0] && q[1] ? q : null; }).filter(Boolean);
      for (var k = 0; k < Math.max(1, n); k++) {
        var T = R + k, item = n ? ctx.rows[k] : {};
        var row = byR[T];
        if (!row) { row = { r: T, attrs: tpl.attrs, cells: [] }; rows.push(row); byR[T] = row; }
        var byC = {}; row.cells.forEach(function (c) { byC[c.c] = c; });
        tplCells.forEach(function (tc) {
          var cell = byC[tc.c];
          if (!cell) { cell = { c: tc.c, attrs: tc.attrs, inner: '', text: null, f: null }; row.cells.push(cell); byC[tc.c] = cell; }
          else if (k > 0 && attr(cell.attrs, 's') == null && attr(tc.attrs, 's') != null) cell.attrs = dropAttr(cell.attrs, 's') + ' s="' + attr(tc.attrs, 's') + '"';
          if (hasPh(tc.text)) { put(cell, L.fillCell(tc.text, ctx, item)); filled++; }
          else if (k > 0 && tc.f) setFormula(cell, relShift(tc.f, k, 0, { sheet: sheetName }));
          else if (k > 0 && tc.hasValue) { cell.attrs = dropAttr(cell.attrs, 't') + (attr(tc.attrs, 't') ? ' t="' + attr(tc.attrs, 't') + '"' : ''); cell.inner = tc.inner; }
        });
        if (k > 0) rowMerges.forEach(function (q) {
          var a = { c: q[0].c, r: T }, b = { c: q[1].c, r: T };
          var clash = taken.some(function (t) { return !(t[1].r < T || t[0].r > T || t[1].c < a.c || t[0].c > b.c); });
          if (!clash) { var ref = colName(a.c) + T + ':' + colName(b.c) + T; newMerges.push(ref); taken.push([{ c: a.c, r: T }, { c: b.c, r: T }]); }
        });
      }
    }
    p.rows = rows;
    // 4) 시트 바깥 부분: 병합·조건부 서식 범위·선택·크기·쪽 나눔을 내리고, 머리글/바닥글의 {{키}} 채우기
    var after = p.after, before = p.before;
    if (ins) {
      // 병합(mergeCell ref)·조건부 서식·데이터 유효성(sqref)·하이퍼링크(ref)·자동 필터(ref)
      after = after.replace(/(\s(?:sqref|ref)=")([^"]+)(")/g, function (_, a, v, b) { return a + shiftSqref(v, at, ins, R) + b; });
      after = after.replace(/(<brk\b[^>]*?\sid=")(\d+)(")/g, function (_, a, v, b) { return a + (+v >= at ? +v + ins : v) + b; });
      before = before.replace(/(<dimension\b[^>]*?ref=")([^"]+)(")/, function (_, a, v, b) { return a + shiftSqref(v, at, ins, R) + b; });
      before = before.replace(/(\s(?:sqref|activeCell|topLeftCell)=")([^"]+)(")/g, function (_, a, v, b) { return a + shiftSqref(v, at, ins, R) + b; });
    }
    if (newMerges.length) {
      if (/<mergeCells\b/.test(after)) after = after.replace(/<mergeCells\b[^>]*>/, function () { return '<mergeCells>'; }).replace('</mergeCells>', newMerges.map(function (r) { return '<mergeCell ref="' + r + '"/>'; }).join('') + '</mergeCells>');
      else after = after.replace(/^((?:\s*<sheetCalcPr\b[^>]*\/>)?(?:\s*<sheetProtection\b[^>]*\/>)?(?:\s*<protectedRanges>[\s\S]*?<\/protectedRanges>)?(?:\s*<scenarios>[\s\S]*?<\/scenarios>)?(?:\s*<autoFilter\b[\s\S]*?(?:\/>|<\/autoFilter>))?(?:\s*<sortState\b[\s\S]*?(?:\/>|<\/sortState>))?(?:\s*<dataConsolidate\b[\s\S]*?(?:\/>|<\/dataConsolidate>))?(?:\s*<customSheetViews>[\s\S]*?<\/customSheetViews>)?)/, function (m) { return m + '<mergeCells>' + newMerges.map(function (r) { return '<mergeCell ref="' + r + '"/>'; }).join('') + '</mergeCells>'; });
    }
    after = after.replace(/<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>/, function (_, inner) { var c = (inner.match(/<mergeCell\b/g) || []).length; return '<mergeCells count="' + c + '">' + inner + '</mergeCells>'; });
    after = after.replace(/(<(odd|even|first)(Header|Footer)>)([\s\S]*?)(<\/\2\3>)/g, function (_, a, x, y, t, b) { if (!hasPh(xmlDecode(t))) return _; note(xmlDecode(t)); return a + xmlEncode(L.fillText(xmlDecode(t), ctx, null)) + b; });
    p.before = before; p.after = after;
    return { xml: sheetXml(p), repeatRow: R, inserted: ins, at: at, filled: filled, unknown: unknown, avail: avail };
  }

  /* 양식 살펴보기 — 어떤 {{키}} 가 어느 칸에 있는지, 반복 줄이 어디인지 */
  function scanTemplate(data) {
    var z = openZip(data), sst = readSst(z), out = { sheets: [], keys: [], unknown: [], repeat: false };
    sheetList(z).forEach(function (sh) {
      var x = z.get(sh.path); if (!x) return;
      var p = parseSheetXml(x, sst), info = { name: sh.name, cells: [], repeatRow: null, rowKeys: [] };
      p.rows.forEach(function (row) {
        row.cells.forEach(function (c) {
          if (!hasPh(c.text)) return;
          var ks = L.placeholdersIn(c.text);
          info.cells.push({ ref: colName(c.c) + row.r, text: c.text, keys: ks });
          ks.forEach(function (k) {
            if (out.keys.indexOf(k) < 0) out.keys.push(k);
            if (!L.KEY_BY[k] && out.unknown.indexOf(k) < 0) out.unknown.push(k);
            if (L.isRowKey(k)) { if (info.repeatRow == null) info.repeatRow = row.r; if (row.r === info.repeatRow && info.rowKeys.indexOf(k) < 0) info.rowKeys.push(k); }
          });
        });
      });
      if (info.repeatRow != null) out.repeat = true;
      out.sheets.push(info);
    });
    out.count = out.sheets.reduce(function (a, s) { return a + s.cells.length; }, 0);
    return out;
  }

  /* 금액 칸 스타일(s)을 복제해 표시 형식만 #,##0.00(내장 번호 4)으로 바꾼 새 스타일을 만듭니다. 같은 s 는 한 번만 */
  function decimalStyles(z) {
    var x = z.get('xl/styles.xml'), made = {}, extra = [], xfs = [];
    var m = x && /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(x);
    if (m) m[1].replace(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g, function (e) { xfs.push(e); return e; });
    return {
      get: function (s) {
        if (!m) return null;
        if (made[s] != null) return made[s];
        var e = xfs[+s]; if (!e) return null;
        var id = +attr(e.slice(0, e.indexOf('>') + 1), 'numFmtId') || 0;
        if (id !== 3 && id !== 1 && id !== 0 && id < 164) return (made[s] = s);   // 이미 소수 형식(2·4 등)이면 그대로
        var ne = e.replace(/^<xf\b([^>]*?)(\/?)>/, function (_, a, sl) { return '<xf' + dropAttr(dropAttr(a, 'numFmtId'), 'applyNumberFormat') + ' numFmtId="4" applyNumberFormat="1"' + sl + '>'; });
        extra.push(ne); made[s] = String(xfs.length + extra.length - 1);
        return made[s];
      },
      save: function () {
        if (!extra.length) return;
        var nx = x.replace(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/, function (_, inner) { return '<cellXfs count="' + (xfs.length + extra.length) + '">' + inner + extra.join('') + '</cellXfs>'; });
        z.set('xl/styles.xml', nx);
      }
    };
  }

  /* 양식 채우기 → { data: Uint8Array, report } */
  function fillTemplate(data, ctx) {
    var z = openZip(data), sst = readSst(z), sheets = sheetList(z), report = { sheets: [], unknown: [] };
    var wb = z.get('xl/workbook.xml');
    var dec = decimalStyles(z);
    sheets.forEach(function (sh) {
      var x = z.get(sh.path); if (!x) return;
      var res = fillSheet(x, ctx, sst, sh.name, { decimalStyle: dec.get });
      z.set(sh.path, res.xml);
      res.unknown.forEach(function (k) { if (report.unknown.indexOf(k) < 0) report.unknown.push(k); });
      report.sheets.push({ name: sh.name, repeatRow: res.repeatRow, inserted: res.inserted, filled: res.filled });
      if (res.inserted) {
        // 인쇄 영역 같은 이름 정의를 함께 내림
        wb = wb.replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, function (_, a, t, b) {
          return a + xmlEncode(insertShift(xmlDecode(t), res.at, res.inserted, res.repeatRow, { sheet: sh.name })) + b;
        });
        // 그림·직인 이미지 위치(앵커 줄 번호는 0 부터)
        var rp = sh.path.replace(/([^/]+)$/, '_rels/$1.rels'), rm = rels(z, rp);
        Object.keys(rm).forEach(function (id) {
          if (!/drawing$/.test(rm[id].type)) return;
          var dp = resolve(sh.path.replace(/[^/]+$/, ''), rm[id].target), dx = z.get(dp);
          if (dx) z.set(dp, dx.replace(/(<xdr:row>)(\d+)(<\/xdr:row>)/g, function (_, a, v, b) { return a + (+v + 1 >= res.at ? +v + res.inserted : v) + b; }));
        });
      }
    });
    // 계산 순서 파일은 지우고(바뀐 칸과 어긋나면 엑셀이 「복구」를 묻습니다), 열 때 다시 계산
    if (z.has('xl/calcChain.xml')) {
      z.del('xl/calcChain.xml');
      var ct = z.get('[Content_Types].xml'); if (ct) z.set('[Content_Types].xml', ct.replace(/<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, ''));
      var wr = z.get('xl/_rels/workbook.xml.rels'); if (wr) z.set('xl/_rels/workbook.xml.rels', wr.replace(/<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*\/>/, ''));
    }
    if (/<calcPr\b/.test(wb)) wb = wb.replace(/<calcPr\b([^>]*?)(\/?)>/, function (_, a, s) { return '<calcPr' + dropAttr(a, 'fullCalcOnLoad') + ' fullCalcOnLoad="1"' + s + '>'; });
    else if (/<\/definedNames>/.test(wb)) wb = wb.replace('</definedNames>', '</definedNames><calcPr fullCalcOnLoad="1"/>');
    else wb = wb.replace('</sheets>', '</sheets><calcPr fullCalcOnLoad="1"/>');
    z.set('xl/workbook.xml', wb);
    dec.save();
    return { data: z.write(), report: report };
  }

  /* ── 여러 파일 한 번에(zip) — 한글 파일 이름이 깨지지 않게 UTF-8 표시를 켠 무압축 zip ── */
  var CRC = (function () { var t = [], c; for (var n = 0; n < 256; n++) { c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(b) { var c = 0xFFFFFFFF; for (var i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function makeZip(files, now) {
    var d = now || new Date(), time = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1), date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    var parts = [], central = [], off = 0;
    function u16(v) { return [v & 255, (v >>> 8) & 255]; }
    function u32(v) { return [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255]; }
    files.forEach(function (f) {
      var name = enc.encode(f.name), data = toU8(f.data), crc = crc32(data);
      var head = [].concat(u32(0x04034b50), u16(20), u16(0x0800), u16(0), u16(time), u16(date), u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0));
      parts.push(new Uint8Array(head), name, data);
      central.push(new Uint8Array([].concat(u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(0), u16(time), u16(date), u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(off))), name);
      off += head.length + name.length + data.length;
    });
    var cdSize = central.reduce(function (a, b) { return a + b.length; }, 0);
    var end = new Uint8Array([].concat(u32(0x06054b50), u16(0), u16(0), u16(files.length), u16(files.length), u32(cdSize), u32(off), u16(0)));
    var all = parts.concat(central, [end]), len = all.reduce(function (a, b) { return a + b.length; }, 0), out = new Uint8Array(len), p = 0;
    all.forEach(function (b) { out.set(b, p); p += b.length; });
    return out;
  }
  function b64ToU8(b64) {
    if (typeof atob === 'function') { var s = atob(b64), u = new Uint8Array(s.length); for (var i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u; }
    return new Uint8Array(Buffer.from(b64, 'base64'));
  }
  function u8ToB64(u) {
    if (typeof btoa === 'function') { var s = ''; for (var i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }
    return Buffer.from(u).toString('base64');
  }

  return {
    xmlDecode: xmlDecode, xmlEncode: xmlEncode, colIndex: colIndex, mapRefs: mapRefs, relShift: relShift, insertShift: insertShift,
    openZip: openZip, sheetList: sheetList, scanTemplate: scanTemplate, fillTemplate: fillTemplate, fillSheet: fillSheet,
    makeZip: makeZip, crc32: crc32, b64ToU8: b64ToU8, u8ToB64: u8ToB64
  };
});
