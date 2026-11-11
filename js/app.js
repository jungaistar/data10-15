/* 클레임 보상 합의서·인보이스 — 화면
   계산은 js/logic.js(CLLogic), 엑셀 양식 채우기는 js/xlsx-template.js(CLTemplate), 기본 양식은 js/builtin-templates.js,
   저장은 js/store.js. 뼈대(가져오기·열 연결·표·내보내기)는 data10-11·data10-12 에서 가져왔습니다. */
(function () {
  'use strict';
  var L = window.CLLogic, T = window.CLTemplate, B = window.CLBuiltin, S = window.CLStore, Sample = window.CLSample;
  /* 판 맞추기 — 브라우저가 옛 logic.js 와 새 app.js 를 섞어 쓰면(Pages 캐시) 버튼과 결과가 어긋날 수 있어, 섞였으면 계산하지 않고 새로고침을 안내합니다.
     판은 scripts/bump-version.mjs 로 index.html(?v=) · logic.js 와 함께 올립니다. */
  var APP_VERSION = '2026-11-11.6';
  if (!L || L.VERSION !== APP_VERSION) {
    var m0 = document.getElementById('main') || document.body;
    m0.innerHTML = '<section class="card" id="versionMismatch"><h1>새 판이 나왔습니다 — 새로고침(Ctrl+F5)해 주세요</h1><p class="alert warn">이 브라우저가 옛 파일과 새 파일을 섞어 불러왔습니다(js/logic.js ' +
      (L && L.VERSION ? L.VERSION : '옛 판') + ' · 화면 ' + APP_VERSION + '). 이대로 계산하면 버튼과 결과가 어긋날 수 있어 멈췄습니다.</p>' +
      '<p>Windows 는 <b>Ctrl+F5</b>, Mac 은 <b>Cmd+Shift+R</b> 을 한 번 눌러 주세요. 쌓아 둔 자료는 지워지지 않습니다.</p></section>';
    return;
  }
  (function () { var ft = document.querySelector('.site-footer .inner') || document.querySelector('.site-footer'); if (!ft) return; var s = document.createElement('span'); s.className = 'ver'; s.textContent = ' · 판 ' + APP_VERSION; ft.appendChild(s); })();
  var db = S.loadDb();
  var ui = { pending: [], errors: [], sel: null, kinds: { ag: true, inv: true }, fSup: 'all', fLv: 'all', printMode: 'html', tplScan: {} };
  var main = document.getElementById('main');

  /* ── 작은 도구 ───────────────────────────── */
  function h(tag, attrs) {
    var e = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v == null || v === false) return;
      if (k === 'class') e.className = v;
      else if (k.slice(0, 2) === 'on' && typeof v === 'function') e.addEventListener(k.slice(2), v);
      else if (v === true) e.setAttribute(k, '');
      else e.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) add(e, arguments[i]);
    return e;
  }
  function add(e, c) {
    if (c == null || c === false) return;
    if (Array.isArray(c)) { c.forEach(function (x) { add(e, x); }); return; }
    e.appendChild(typeof c === 'object' ? c : document.createTextNode(String(c)));
  }
  function fld(label, control, hint) {
    return h('label', { class: 'field' }, h('span', null, label), control, hint ? h('small', { class: 'hint' }, hint) : null);
  }
  var toastTimer = null;
  function toast(msg, err) {
    var t = document.getElementById('toast');
    t.textContent = msg; t.className = 'toast' + (err ? ' error' : ''); t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.hidden = true; }, 3500);
  }
  function save() {
    var ok = S.saveDb(db);
    document.getElementById('storeNote').textContent = ok ? '이 브라우저에 저장됨 · ' + S.sizeKb() + 'KB' : '저장 안 됨(브라우저 저장소가 막혔거나 가득 참) — JSON 백업을 받아 두세요';
    return ok;
  }
  function download(name, blob) {
    var a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  function copyText(t, msg) {
    function fallback() { var ta = h('textarea'); ta.value = t; document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); toast(msg || '복사했습니다.'); } catch (e) { toast('복사하지 못했습니다. 직접 선택해 복사해 주세요.', true); } ta.remove(); }
    if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(t).then(function () { toast(msg || '복사했습니다.'); }, fallback); else fallback();
  }
  function pageHead(stage, title, lead, right) {
    return h('div', { class: 'page-head no-print' }, h('div', { class: 'titles' }, h('div', { class: 'stage' }, stage), h('h1', null, title), lead ? h('p', { class: 'lead' }, lead) : null), right || null);
  }
  function emptyNotice() {
    return h('section', { class: 'card empty' },
      h('p', null, '아직 가져온 합의 결과가 없습니다.'),
      h('p', null, '협력사와 합의를 마친 클레임 보상 결과 엑셀을 가져오거나, 가상 샘플로 먼저 둘러봐 주세요.'),
      h('div', { class: 'btn-row', style: 'justify-content:center' },
        h('a', { class: 'btn btn-primary', href: '#/import' }, '데이터 가져오기로'),
        h('button', { type: 'button', class: 'btn', onclick: loadSample }, '샘플로 시작')));
  }
  function tile(k, v, sub, cls) { return h('div', { class: 'tile' + (cls ? ' ' + cls : '') }, h('div', { class: 'k' }, k), h('div', { class: 'v' }, v), sub ? h('div', { class: 'note' }, sub) : null); }
  function table(head, rows, opts) {
    opts = opts || {};
    var t = h('table', { class: 'list' + (opts.cls ? ' ' + opts.cls : '') });
    var tr = h('tr');
    head.forEach(function (x, i) { tr.appendChild(h('th', { class: opts.num && opts.num[i] ? 'num' : null, scope: 'col' }, x)); });
    t.appendChild(h('thead', null, tr));
    var tb = h('tbody');
    rows.forEach(function (r) {
      var row = h('tr', { class: r.cls || null });
      (r.cells || r).forEach(function (x, i) { row.appendChild(h('td', { class: ((opts.num && opts.num[i] ? 'num' : '') + (r.tdCls && r.tdCls[i] ? ' ' + r.tdCls[i] : '')) || null }, x)); });
      tb.appendChild(row);
    });
    t.appendChild(tb);
    return h('div', { class: 'table-wrap' }, t);
  }
  var LV = { bad: '확인 필요', warn: '주의', ok: '정상', info: '참고' };
  function lvPill(lv) { return h('span', { class: 'st st-' + lv }, LV[lv]); }
  function money(n, cur) { return n == null ? '' : L.fmtMoney(n, cur); }
  function safeName(s) { return String(s).replace(/[\\/:*?"<>|\s]+/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, ''); }

  /* ── 설정 · 묶음 ─────────────────────────── */
  function settings() { return L.cleanSettings(db.settings); }
  function issueDate() { return db.issueDate || L.localToday(); }
  function groups() { return L.groupBySupplier(db.rows, settings()); }
  function selected(gs) {
    if (!ui.sel) { ui.sel = {}; gs.forEach(function (g) { ui.sel[g.id] = !g.bad && g.count > 0; }); }
    return gs.filter(function (g) { return ui.sel[g.id] && g.count > 0; });
  }
  function kindsOn() { return ['ag', 'inv'].filter(function (k) { return ui.kinds[k]; }); }
  function tplBytes(kind) { var t = db.templates[kind]; return T.b64ToU8(t ? t.b64 : B[kind].b64); }
  function tplName(kind) { var t = db.templates[kind]; return t ? t.name + ' (올린 양식)' : B[kind].name + ' (기본 양식)'; }

  /* 협력사 하나에 대해 고른 문서들의 번호를 정합니다. commit=true 면 이력에 씁니다(재발행이면 같은 번호) */
  function numbersFor(g, kinds, commit) {
    var Sx = settings(), d = issueDate(), hist = db.history, out = {};
    kinds.forEach(function (k) {
      var r = L.issueDoc(hist, k, g, d, Sx, new Date().toISOString().slice(0, 16).replace('T', ' '));
      if (commit) hist = r.history;
      out[k] = { docNo: r.entry.docNo, reused: r.reused };
    });
    if (commit) db.history = hist;
    ['ag', 'inv'].forEach(function (k) { if (!out[k]) { var n = L.latestDocNo(db.history, k, g); if (n) out[k] = { docNo: n, old: true }; } });
    return out;
  }
  function contextFor(g, kind, nums) {
    var other = kind === 'ag' ? 'inv' : 'ag';
    return L.buildContext(g, kind, settings(), { docNo: nums[kind] ? nums[kind].docNo : '', relatedNo: nums[other] ? nums[other].docNo : '', issueDate: issueDate() });
  }

  /* ── 가져오기 ───────────────────────────── */
  function readFile(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () {
        try { resolve(prepare(file.name, XLSX.read(new Uint8Array(fr.result), { type: 'array', cellDates: true }))); }
        catch (e) { reject(new Error(file.name + ': 엑셀로 읽지 못했습니다(' + e.message + '). 문서보안(DRM)이 걸린 파일이면 CSV 로 저장해 넣어 주세요.')); }
      };
      fr.onerror = function () { reject(new Error(file.name + ': 파일을 읽지 못했습니다.')); };
      fr.readAsArrayBuffer(file);
    });
  }
  function prepare(name, wb) {
    var sheets = wb.SheetNames.map(function (n) { return { name: n, aoa: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) }; });
    var pick = L.pickSheet(sheets) || sheets[0];
    var it = { name: name, sheets: sheets, sheet: pick.name, overrides: {} };
    reparse(it);
    return it;
  }
  function reparse(it) {
    var sh = it.sheets.filter(function (s) { return s.name === it.sheet; })[0];
    it.result = L.parseSheet(sh.aoa, { overrides: Object.keys(it.overrides).length ? it.overrides : null });
  }
  function handleFiles(list) {
    var files = Array.prototype.slice.call(list || []).filter(function (f) { return /\.(xlsx|xlsm|xls|csv)$/i.test(f.name); });
    if (!files.length) { toast('엑셀 파일(.xlsx·.xls·.csv)을 골라 주세요.', true); return; }
    ui.errors = [];
    Promise.all(files.map(function (f) { return readFile(f).then(function (it) { ui.pending.push(it); }, function (e) { ui.errors.push(e.message); }); }))
      .then(function () { render(); });
  }
  function commit(items, mode) {
    var all = [];
    items.forEach(function (it) {
      all = all.concat(it.result.rows);
      db.files.push({ name: it.name, sheet: it.sheet, rows: it.result.rows.length, at: new Date().toISOString().slice(0, 16).replace('T', ' ') });
    });
    if (db.sample || mode === 'replace') { db.rows = []; db.sample = false; }
    // 같은 행(id 가 같음)은 새 값으로
    var byId = {}; db.rows.forEach(function (r, i) { byId[r.id] = i; });
    all.forEach(function (r) { if (byId[r.id] != null) db.rows[byId[r.id]] = r; else { byId[r.id] = db.rows.length; db.rows.push(r); } });
    ui.pending = ui.pending.filter(function (p) { return items.indexOf(p) < 0; });
    ui.sel = null;
    save();
    toast('반영했습니다 — 합의 결과 ' + all.length + '행.');
    location.hash = '#/groups';
    render();
  }
  function loadSample() {
    var b = Sample.build();
    var aoa = b.aoa.map(function (row, i) { return i < 4 ? row : row.map(function (v, c) { return b.dateCols.indexOf(c) >= 0 && typeof v === 'string' && v ? new Date(v + 'T00:00:00') : v; }); });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa, { cellDates: true }), '합의결과');
    var it = prepare('가상 샘플 합의 결과', XLSX.read(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }), { type: 'array', cellDates: true }));
    var keepTpl = db.templates;
    db = L.emptyDb();
    db.templates = keepTpl;
    db.rows = it.result.rows;
    db.files = [{ name: it.name, sheet: it.sheet, rows: db.rows.length, at: new Date().toISOString().slice(0, 16).replace('T', ' ') }];
    db.settings.issuer = Sample.issuer;
    db.settings.person = '보상 담당(가상)';
    db.settings.contact = '000-0000-0000';
    db.settings.bank = '가상은행 000-000000-00-000 (예금주 ' + Sample.issuer + ')';
    db.sample = true;
    ui.pending = []; ui.sel = null;
    save();
    toast('가상 샘플 22행(협력사 6곳)을 불러왔습니다.');
    location.hash = '#/groups';
    render();
  }
  function pageImport() {
    var input = h('input', { type: 'file', multiple: true, accept: '.xlsx,.xlsm,.xls,.csv' });
    input.addEventListener('change', function () { handleFiles(input.files); input.value = ''; });
    var drop = h('div', { class: 'drop' },
      h('label', { class: 'btn btn-primary' }, '합의 결과 엑셀 고르기', input),
      h('span', { class: 'note' }, '또는 파일을 여기에 끌어다 놓아 주세요. 여러 개를 함께 넣어도 됩니다.'));
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('over'); });
    drop.addEventListener('drop', function (e) { e.preventDefault(); drop.classList.remove('over'); handleFiles(e.dataTransfer.files); });
    var out = [pageHead('0', '데이터 가져오기', '협력사와 보상 합의를 마친 결과 엑셀을 넣으면 열을 찾아 연결하고, 협력사별로 묶습니다.'),
      h('section', { class: 'card' },
        drop,
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn', onclick: function () { if (!db.rows.length || confirm('지금 데이터를 지우고 가상 샘플을 불러올까요? (올린 양식은 그대로 둡니다)')) loadSample(); } }, '샘플로 시작'),
          h('a', { class: 'btn', href: 'samples/' + encodeURIComponent('클레임_합의결과_샘플.xlsx'), download: '' }, '샘플 합의 결과 내려받기')),
        h('ul', { class: 'note', style: 'margin-top:10px' },
          h('li', null, '꼭 필요한 열: 협력사명 + 합의금액(없으면 청구금액과 부담 비율). 있으면 쓰는 열: 협력사 코드 · 클레임 번호 · 발생일 · 기종 · 품번 · 품명 · 수량 · 통화 · 합의일 · 비고.'),
          h('li', null, '열 이름이 조금 달라도(「업체명」「거래처」「보상금액」「귀책률」 …) 찾고, 못 찾은 열은 아래에서 골라 연결합니다. 제목 줄이 몇 줄 있어도 머리행을 찾습니다.'),
          h('li', null, '맨 아래 「합계」 줄은 건너뜁니다. 같은 파일을 다시 넣으면 같은 행은 새 값으로 바뀝니다.')),
        ui.errors.length ? h('div', { class: 'alert warn' }, h('ul', null, ui.errors.map(function (e) { return h('li', null, e); }))) : null)];
    if (ui.pending.length) {
      var ready = ui.pending.filter(function (it) { return !it.result.det.missing.length && it.result.rows.length; });
      var box = h('section', { class: 'card' }, h('h2', null, '확인 후 반영'));
      ui.pending.forEach(function (it) { box.appendChild(pendingCard(it)); });
      box.appendChild(h('div', { class: 'btn-row', style: 'margin-top:12px' },
        h('button', { type: 'button', class: 'btn btn-primary', disabled: !ready.length, onclick: function () { commit(ready, 'replace'); } }, '확인한 ' + ready.length + '개 파일로 바꾸기'),
        db.rows.length && !db.sample ? h('button', { type: 'button', class: 'btn', disabled: !ready.length, onclick: function () { commit(ready, 'append'); } }, '지금 데이터에 더하기') : null,
        h('button', { type: 'button', class: 'btn', onclick: function () { ui.pending = []; render(); } }, '취소')));
      out.push(box);
    }
    if (db.files.length) out.push(h('section', { class: 'card' }, h('h2', null, '가져온 파일'),
      table(['파일', '시트', '행 수', '가져온 때'], db.files.slice().reverse().map(function (f) { return [f.name, f.sheet, f.rows, f.at]; }), { num: [0, 0, 1, 0] }),
      h('div', { class: 'btn-row', style: 'margin-top:10px' }, h('button', { type: 'button', class: 'btn btn-danger', onclick: function () {
        if (confirm('합의 결과·발행 이력·올린 양식·설정을 모두 지울까요? 되돌릴 수 없습니다(JSON 백업을 먼저 받아 두세요).')) { S.clearDb(); db = L.emptyDb(); ui.sel = null; save(); render(); }
      } }, '모든 데이터 지우기'))));
    return out;
  }
  function pendingCard(it) {
    var r = it.result, det = r.det;
    var sheetSel = h('select', { 'aria-label': '시트' });
    it.sheets.forEach(function (s) { var o = h('option', { value: s.name }, s.name); if (s.name === it.sheet) o.selected = true; sheetSel.appendChild(o); });
    sheetSel.addEventListener('change', function () { it.sheet = sheetSel.value; it.overrides = {}; reparse(it); render(); });
    var hdr = det.headers || [];
    var grid = h('div', { class: 'map-grid' });
    L.FIELDS.forEach(function (f) {
      var sel = h('select', { 'aria-label': f.label + ' 열' });
      sel.appendChild(h('option', { value: '-1' }, '(없음)'));
      hdr.forEach(function (t, i) { var o = h('option', { value: String(i) }, L.colName(i) + ' · ' + (t || '(빈 머리)')); if (det.map[f.id] === i) o.selected = true; sel.appendChild(o); });
      if (det.map[f.id] == null) sel.value = '-1';
      sel.addEventListener('change', function () { it.overrides[f.id] = Number(sel.value); reparse(it); render(); });
      var miss = (f.id === 'supplier' && det.missing.indexOf('supplier') >= 0) || (f.id === 'agreedAmt' && det.missing.indexOf('agreedAmt') >= 0);
      grid.appendChild(h('div', { class: 'map-row' + (miss ? ' missing' : '') }, h('span', { class: 'col' }, f.label, f.req ? h('span', { class: 'req' }, ' 꼭 필요') : f.id === 'agreedAmt' ? h('span', { class: 'req' }, ' 또는 청구금액+부담 비율') : null), sel));
    });
    return h('details', { class: 'block', open: true },
      h('summary', null, it.name + ' — ' + r.rows.length + '행' + (r.skipped ? ' · 건너뜀 ' + r.skipped : '') + (det.missing.length ? ' · 열 연결 필요' : '')),
      h('div', { class: 'form-grid', style: 'margin:8px 0' }, fld('시트', sheetSel, det.headerRow >= 0 ? '머리행: ' + (det.headerRow + 1) + '행' : '')),
      det.warnings.map(function (w) { return h('p', { class: 'alert warn' }, w); }),
      grid,
      r.warnings.length ? h('details', { class: 'warn-list' }, h('summary', null, '행 알림 ' + r.warnings.length + '건'), h('ul', null, r.warnings.slice(0, 80).map(function (w) { return h('li', null, w.row + '행: ' + w.msg); }))) : null);
  }

  /* ── ① 협력사별 합의 ───────────────────────── */
  function totalsByCur(gs, key) {
    var o = {}; gs.forEach(function (g) { o[g.currency] = (o[g.currency] || 0) + g[key]; });
    return Object.keys(o).map(function (c) { return c + ' ' + money(o[c], c); }).join(' · ');
  }
  function pageGroups() {
    if (!db.rows.length) return [pageHead('①', '협력사별 합의', '협력사마다 클레임을 묶어 합계를 냅니다'), emptyNotice()];
    var gs = groups(), Sx = settings(), sel = selected(gs);
    var ev = L.evaluateRows(db.rows, Sx);
    var bad = ev.filter(function (r) { return r.level === 'bad'; }).length, warn = ev.filter(function (r) { return r.level === 'warn'; }).length, inc = ev.filter(function (r) { return r.included; }).length;
    var rows = gs.map(function (g) {
      var cb = h('input', { type: 'checkbox', 'aria-label': g.supplier + ' 문서 만들기' }); cb.checked = !!ui.sel[g.id]; cb.disabled = !g.count;
      cb.addEventListener('change', function () { ui.sel[g.id] = cb.checked; render(); });
      var ag = L.latestDocNo(db.history, 'ag', g), inv = L.latestDocNo(db.history, 'inv', g);
      var st = g.bad ? 'bad' : g.warn ? 'warn' : 'ok';
      return { cells: [cb, h('a', { href: '#/rows', onclick: function () { ui.fSup = g.key; } }, g.supplier), g.supplierCode, g.count + (g.excluded.length ? ' (+뺌 ' + g.excluded.length + ')' : ''), money(g.claimSum, g.currency), money(g.agreedSum, g.currency), money(g.vat, g.currency), h('strong', null, money(g.total, g.currency)), g.currency, g.ratio == null ? '' : L.fmtPct(g.ratio),
        h('span', null, lvPill(st), g.bad ? ' ' + g.bad + '건' : g.warn ? ' ' + g.warn + '건' : ''),
        h('span', { class: 'mono' }, ag || '-', h('br'), inv || '-')], cls: 'lv-' + st };
    });
    return [
      pageHead('①', '협력사별 합의', '합의 결과를 협력사마다 묶었습니다. 검증에서 「확인 필요」가 있는 협력사는 처음에 선택을 꺼 둡니다. 고른 협력사로 ③ 문서 만들기를 합니다.',
        h('div', { class: 'btn-row no-print' }, h('a', { class: 'btn btn-primary', href: '#/docs' }, '고른 ' + sel.length + '곳 문서 만들기'), h('a', { class: 'btn', href: '#/rows' }, '클레임 명세·검증'))),
      h('div', { class: 'tiles wide' },
        tile('협력사', gs.length + '곳', '문서 만들 곳 ' + sel.length + '곳 선택'),
        tile('클레임', inc + '행', '전체 ' + ev.length + '행 · 문서에서 뺀 ' + (ev.length - inc) + '행'),
        tile('합의 합계', totalsByCur(gs, 'agreedSum') || '-', Sx.vat === 'add10' ? '부가세 10% 별도' : '부가세 없음(설정)'),
        tile('확인 필요', bad + '행', '합의 > 청구 · 금액 없음 · 비율 범위 · 통화 섞임', 'lv ' + (bad ? 'bad' : 'ok')),
        tile('주의', warn + '행', '비율 계산과 다름 · 번호 겹침/없음 · 0원', 'lv ' + (warn ? 'warn' : 'ok'))),
      h('section', { class: 'card' }, h('h2', null, '협력사 (' + gs.length + '곳)'),
        h('div', { class: 'btn-row no-print', style: 'margin-bottom:8px' },
          h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { gs.forEach(function (g) { ui.sel[g.id] = g.count > 0; }); render(); } }, '모두 고르기'),
          h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { gs.forEach(function (g) { ui.sel[g.id] = false; }); render(); } }, '모두 풀기'),
          h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { gs.forEach(function (g) { ui.sel[g.id] = !g.bad && g.count > 0; }); render(); } }, '확인 필요 없는 곳만')),
        table(['고름', '협력사', '코드', '건수', '청구합계', '합의합계', '부가세', '총액', '통화', '부담 비율', '검증', '최근 발행(합의서/인보이스)'], rows, { num: [0, 0, 0, 1, 1, 1, 1, 1, 0, 1, 0, 0] })),
      h('section', { class: 'card' }, h('h2', null, '금액을 정하는 규칙'),
        h('ul', { class: 'note' },
          h('li', null, '합의금액 칸이 있으면 그 값을 씁니다. 비어 있으면 청구금액 × 부담 비율로 계산합니다(설정의 반올림 방식·단위, 원화가 아니면 소수 둘째 자리).'),
          h('li', null, '합의금액과 「청구금액 × 부담 비율」이 ' + Sx.tolerance + '원 넘게 다르면 「주의」, 합의금액이 청구금액보다 크면 「확인 필요」입니다.'),
          h('li', null, '부가세는 설정에서 고릅니다(지금: ' + (Sx.vat === 'add10' ? '합의합계의 10% 별도' : '없음') + '). 보상금에 부가세를 붙이는지는 회사 기준을 확인해 주세요(기획서 10장 ③).')))
    ];
  }

  /* ── ② 클레임 명세·검증 ───────────────────── */
  function pageRows() {
    if (!db.rows.length) return [pageHead('②', '클레임 명세·검증', ''), emptyNotice()];
    var Sx = settings(), gs = groups(), ev = L.evaluateRows(db.rows, Sx);
    var supSel = h('select', { 'aria-label': '협력사' }, [['all', '전체 협력사']].concat(gs.map(function (g) { return [g.key, g.supplier]; })).map(function (o) { var op = h('option', { value: o[0] }, o[1]); if (o[0] === ui.fSup) op.selected = true; return op; }));
    supSel.addEventListener('change', function () { ui.fSup = supSel.value; render(); });
    var lvSel = h('select', { 'aria-label': '판정' }, [['all', '전체'], ['issue', '확인 필요 + 주의'], ['bad', '확인 필요'], ['warn', '주의'], ['excluded', '문서에서 뺀 행']].map(function (o) { var op = h('option', { value: o[0] }, o[1]); if (o[0] === ui.fLv) op.selected = true; return op; }));
    lvSel.addEventListener('change', function () { ui.fLv = lvSel.value; render(); });
    var list = ev.filter(function (r) {
      if (ui.fSup !== 'all' && L.supplierKey(r) !== ui.fSup) return false;
      if (ui.fLv === 'issue') return r.level !== 'ok';
      if (ui.fLv === 'excluded') return !r.included;
      return ui.fLv === 'all' || r.level === ui.fLv;
    });
    var rows = list.map(function (r) {
      return { cells: [r.line, r.supplier, r.claimNo || '-', r.date, r.partNo, r.partName, r.qty == null ? '' : r.qty, money(r.claimAmt, r.currency), r.ratio == null ? '' : L.fmtPct(r.ratio), h('strong', null, money(r.agreed, r.currency)), r.currency, lvPill(r.level),
        h('div', null, r.flags.map(function (f) { return h('div', { class: f.level }, f.msg); }), r.note ? h('div', { class: 'sub' }, '비고: ' + r.note) : null)],
        cls: 'lv-' + r.level + (r.included ? '' : ' excluded'), tdCls: { 12: 'flags' } };
    });
    return [
      pageHead('②', '클레임 명세·검증', '합의 결과 한 행 한 행의 금액과 검사 결과입니다. 원본 엑셀을 고친 뒤 다시 가져오면 바로 반영됩니다.',
        h('div', { class: 'btn-row no-print' }, h('button', { type: 'button', class: 'btn', onclick: function () {
          var wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([L.REVIEW_HEAD].concat(L.reviewRows(db.rows, Sx))), '검토표');
          XLSX.writeFile(wb, '클레임_합의_검토표_' + L.localToday() + '.xlsx');
        } }, '검토표 엑셀로'))),
      h('section', { class: 'card filters no-print' }, fld('협력사', supSel), fld('판정', lvSel), h('p', { class: 'note span-all' }, list.length + '행 보임 · 「행」은 원본 엑셀의 줄 번호입니다.')),
      table(['행', '협력사', '클레임 번호', '발생일', '품번', '품명', '수량', '청구금액', '부담 비율', '합의금액', '통화', '판정', '검사 내용'], rows, { num: [1, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, 0] })
    ];
  }

  /* ── ③ 문서 만들기 ─────────────────────────── */
  function makeFiles(gs, kinds) {
    var files = [], warn = [];
    gs.forEach(function (g) {
      var nums = numbersFor(g, kinds, true);
      kinds.forEach(function (k) {
        var res = T.fillTemplate(tplBytes(k), contextFor(g, k, nums));
        res.report.unknown.forEach(function (u) { if (warn.indexOf(u) < 0) warn.push(u); });
        files.push({ name: safeName(L.DOC_KINDS[k] + '_' + g.supplier + '_' + nums[k].docNo) + '.xlsx', data: res.data });
      });
    });
    save();
    return { files: files, unknown: warn };
  }
  function downloadXlsx() {
    var gs = selected(groups()), kinds = kindsOn();
    if (!gs.length || !kinds.length) { toast('협력사와 문서 종류를 하나 이상 골라 주세요.', true); return; }
    try {
      var r = makeFiles(gs, kinds);
      var XT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      if (r.files.length === 1) download(r.files[0].name, new Blob([r.files[0].data], { type: XT }));
      else download('합의서_인보이스_' + issueDate() + '_' + r.files.length + '개.zip', new Blob([T.makeZip(r.files.map(function (f) { return { name: f.name, data: f.data }; }))], { type: 'application/zip' }));
      toast('엑셀 ' + r.files.length + '개를 만들었습니다' + (r.files.length > 1 ? '(zip 한 파일)' : '') + '.' + (r.unknown.length ? ' 양식의 모르는 키: ' + r.unknown.join(', ') : ''), !!r.unknown.length);
      render();
    } catch (e) { toast('만들지 못했습니다: ' + e.message, true); }
  }
  function pageDocs() {
    if (!db.rows.length) return [pageHead('③', '문서 만들기', ''), emptyNotice()];
    var gs = groups(), sel = selected(gs), kinds = kindsOn(), Sx = settings();
    var dateIn = h('input', { type: 'date', value: issueDate() });
    dateIn.addEventListener('change', function () { db.issueDate = dateIn.value; save(); render(); });
    var kindBox = h('div', { class: 'pick-grid' }, ['ag', 'inv'].map(function (k) {
      var cb = h('input', { type: 'checkbox' }); cb.checked = !!ui.kinds[k];
      cb.addEventListener('change', function () { ui.kinds[k] = cb.checked; render(); });
      return h('label', { class: 'opt' }, cb, h('span', null, h('span', { class: 't' }, L.DOC_KINDS[k]), h('span', { class: 's' }, '양식: ' + tplName(k))));
    }));
    var supBox = h('div', { class: 'pick-grid' }, gs.map(function (g) {
      var cb = h('input', { type: 'checkbox' }); cb.checked = !!ui.sel[g.id]; cb.disabled = !g.count;
      cb.addEventListener('change', function () { ui.sel[g.id] = cb.checked; render(); });
      return h('label', { class: 'opt' }, cb, h('span', null, h('span', { class: 't' }, g.supplier, ' ', g.bad ? lvPill('bad') : g.warn ? lvPill('warn') : null), h('span', { class: 's' }, g.count + '건 · ' + g.currency + ' ' + money(g.total, g.currency))));
    }));
    var prev = [];
    sel.forEach(function (g) {
      var nums = numbersFor(g, kinds, false);
      kinds.forEach(function (k) { prev.push([L.DOC_KINDS[k], g.supplier, h('span', { class: 'mono' }, nums[k].docNo), nums[k].reused ? '재발행(같은 번호)' : '새 번호', g.count, money(g.total, g.currency) + ' ' + g.currency, g.currency === 'KRW' ? L.amountWords(g.total, g.currency) : '']); });
    });
    var badSel = sel.filter(function (g) { return g.bad; });
    return [
      pageHead('③', '문서 만들기', '고른 협력사마다 합의서·인보이스를 양식에 채워 엑셀로 내려받거나, 인쇄 화면에서 PDF 로 저장합니다.'),
      h('section', { class: 'card' }, h('h2', null, '1. 발행일과 문서 종류'),
        h('div', { class: 'form-grid' }, fld('발행일', dateIn, '문서 번호의 연·월과 지급 기한(+' + Sx.payDays + '일)의 기준')),
        h('div', { style: 'margin-top:12px' }, kindBox),
        h('p', { class: 'note', style: 'margin-top:8px' }, '양식은 ⑤ 양식 템플릿에서 바꿉니다. 회사 양식에 {{협력사명}} 같은 키만 적어 올리면 서식은 그대로 두고 값만 채웁니다.')),
      h('section', { class: 'card' }, h('h2', null, '2. 협력사 (' + sel.length + '/' + gs.length + '곳)'), supBox,
        badSel.length ? h('p', { class: 'alert warn' }, '「확인 필요」가 있는 협력사를 골랐습니다: ' + badSel.map(function (g) { return g.supplier; }).join(', ') + ' — ② 클레임 명세·검증에서 금액을 확인해 주세요.') : null),
      h('section', { class: 'card' }, h('h2', null, '3. 만들 문서 (' + prev.length + '개)'),
        prev.length ? table(['종류', '협력사', '문서 번호', '번호', '건수', '총액', '한글 금액'], prev, { num: [0, 0, 0, 0, 1, 1, 0] }) : h('p', { class: 'note' }, '협력사와 문서 종류를 골라 주세요.'),
        h('p', { class: 'note', style: 'margin-top:8px' }, '번호는 내려받기·인쇄를 할 때 확정되어 ⑥ 발행 이력에 남습니다. 같은 협력사·같은 내용을 다시 내면 같은 번호(재발행)이고, 내용이 바뀌면 새 번호를 주고 이전 번호는 「대체됨」이 됩니다.'),
        h('div', { class: 'btn-row', style: 'margin-top:10px' },
          h('button', { type: 'button', class: 'btn btn-primary', disabled: !prev.length, onclick: downloadXlsx }, prev.length > 1 ? '엑셀 ' + prev.length + '개 내려받기(zip)' : '엑셀 내려받기'),
          h('a', { class: 'btn', href: '#/print' }, '인쇄 / PDF 화면으로')))
    ];
  }

  /* ── ④ 인쇄 · PDF ───────────────────────────── */
  function paperAg(ctx, g, Sx) {
    var f = ctx.fields, cur = ctx.currency;
    var meta = h('table', null, h('tbody', null,
      h('tr', null, h('th', null, '문서 번호'), h('td', null, f['문서번호']), h('th', null, '작성일'), h('td', null, f['발행일_한글'])),
      h('tr', null, h('th', null, '갑(발주사)'), h('td', null, f['발행처'] || '(설정에서 회사명 입력)'), h('th', null, '을(협력사)'), h('td', null, f['협력사명'])),
      h('tr', null, h('th', null, '담당 부서'), h('td', null, f['발행부서']), h('th', null, '협력사 코드'), h('td', null, f['협력사코드']))));
    var items = h('table', null,
      h('thead', null, h('tr', null, ['번호', '클레임 번호', '발생일', '품번', '품명', '수량', '청구금액', '부담 비율', '합의금액'].map(function (x) { return h('th', null, x); }))),
      h('tbody', null, ctx.rows.map(function (r) {
        return h('tr', null, h('td', { class: 'c' }, r['행.번호']), h('td', null, r['행.클레임번호']), h('td', { class: 'c' }, r['행.발생일']), h('td', null, r['행.품번']), h('td', null, r['행.품명']),
          h('td', { class: 'n' }, r['행.수량'] == null ? '' : r['행.수량']), h('td', { class: 'n' }, money(r['행.청구금액'], cur)), h('td', { class: 'n' }, r['행.부담비율'] == null ? '' : L.fmtPct(r['행.부담비율'])), h('td', { class: 'n' }, money(r['행.합의금액'], cur)));
      }), h('tr', { class: 'sum' }, h('td', { colspan: 5, class: 'c' }, '합계'), h('td', { class: 'n' }, f['수량합계']), h('td', { class: 'n' }, money(f['청구합계'], cur)), h('td', { class: 'n' }, f['부담비율'] == null ? '' : L.fmtPct(f['부담비율'])), h('td', { class: 'n' }, money(f['합의합계'], cur)))));
    var clauses = Sx.agClauses.split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean);
    return h('article', { class: 'doc' },
      h('h1', { class: 'doc-title' }, '클레임 보상 합의서'),
      h('div', { class: 'doc-wrap' }, meta),
      h('p', null, '갑과 을은 아래 클레임의 보상 금액에 대하여 다음과 같이 합의한다.'),
      h('div', { class: 'amount' }, '합의 금액: ' + f['총액_표기'], h('small', null, '(합의합계 ' + money(f['합의합계'], cur) + (f['부가세'] ? ' · 부가세 ' + money(f['부가세'], cur) : '') + ' · 클레임 ' + f['클레임건수'] + '건 · 청구합계 대비 ' + L.fmtPct(f['부담비율']) + ')')),
      h('h2', null, '1. 대상 클레임'), h('div', { class: 'doc-wrap' }, items),
      h('h2', null, '2. 합의 조건'), h('ol', { class: 'clauses' }, clauses.map(function (c) { return h('li', null, L.fillText(c, ctx)); })),
      h('p', { class: 'muted' }, '이 합의를 증명하기 위하여 합의서 2부를 작성하여 갑과 을이 서명(날인)한 뒤 1부씩 보관한다.'),
      h('div', { class: 'date-line' }, f['합의일_한글']),
      h('div', { class: 'signs' },
        h('div', { class: 'sign' }, h('span', { class: 'seal' }, '(인)'), h('div', { class: 'who' }, '갑: ' + (f['발행처'] || '')), h('div', null, [f['발행부서'], f['담당자']].filter(Boolean).join(' '))),
        h('div', { class: 'sign' }, h('span', { class: 'seal' }, '(인)'), h('div', { class: 'who' }, '을: ' + f['협력사명']), h('div', null, '대표자'))));
  }
  function paperInv(ctx, g, Sx) {
    var f = ctx.fields, cur = ctx.currency;
    var meta = h('table', null, h('tbody', null,
      h('tr', null, h('th', null, 'Invoice No.'), h('td', null, f['문서번호']), h('th', null, 'Date'), h('td', null, f['발행일'])),
      h('tr', null, h('th', null, '관련 합의서'), h('td', null, f['관련문서번호'] || '-'), h('th', null, '지급 기한'), h('td', null, f['지급기한'])),
      h('tr', null, h('th', null, '청구처 (Bill To)'), h('td', { colspan: 3 }, f['협력사명'] + (f['협력사코드'] ? ' (' + f['협력사코드'] + ')' : ''))),
      h('tr', null, h('th', null, '발행처 (From)'), h('td', { colspan: 3 }, [f['발행처'] || '(설정에서 회사명 입력)', f['발행부서']].filter(Boolean).join(' '))),
      h('tr', null, h('th', null, '담당 · 연락처'), h('td', { colspan: 3 }, [f['담당자'], f['연락처']].filter(Boolean).join(' ')))));
    var items = h('table', null,
      h('thead', null, h('tr', null, ['No', '클레임 번호', '발생일', '품번', '품명', '수량', '통화', '금액'].map(function (x) { return h('th', null, x); }))),
      h('tbody', null, ctx.rows.map(function (r) {
        return h('tr', null, h('td', { class: 'c' }, r['행.번호']), h('td', null, r['행.클레임번호']), h('td', { class: 'c' }, r['행.발생일']), h('td', null, r['행.품번']), h('td', null, r['행.품명']),
          h('td', { class: 'n' }, r['행.수량'] == null ? '' : r['행.수량']), h('td', { class: 'c' }, cur), h('td', { class: 'n' }, money(r['행.합의금액'], cur)));
      }),
      h('tr', null, h('td', { colspan: 7, class: 'n' }, '공급가액 (합의합계)'), h('td', { class: 'n' }, money(f['합의합계'], cur))),
      h('tr', null, h('td', { colspan: 7, class: 'n' }, '부가세'), h('td', { class: 'n' }, money(f['부가세'], cur))),
      h('tr', { class: 'sum' }, h('td', { colspan: 7, class: 'n' }, '합계 (Total)'), h('td', { class: 'n' }, money(f['총액'], cur)))));
    return h('article', { class: 'doc' },
      h('h1', { class: 'doc-title' }, 'INVOICE'), h('p', { class: 'doc-sub' }, '클레임 보상 청구서'),
      h('div', { class: 'doc-wrap' }, meta), h('div', { class: 'doc-wrap' }, items),
      h('div', { class: 'amount' }, '금액: ' + f['총액_표기']),
      h('table', null, h('tbody', null, h('tr', null, h('th', null, '입금 계좌'), h('td', null, f['입금계좌'] || '(설정에서 입금 계좌 입력)')))),
      h('p', { class: 'muted' }, L.fillText(Sx.invNote, ctx)),
      h('div', { class: 'issuer' }, h('div', { class: 'who', style: 'font-weight:800;font-size:11pt' }, f['발행처'] || ''), h('div', null, '발행 담당: ' + (f['담당자'] || '') + '   (인)')));
  }
  function paperXl(kind, ctx) {
    var res = T.fillTemplate(tplBytes(kind), ctx), wb = XLSX.read(res.data, { type: 'array' });
    return wb.SheetNames.map(function (n) {
      var box = h('div', { class: 'xl-sheet' });
      box.innerHTML = XLSX.utils.sheet_to_html(wb.Sheets[n], { header: '', footer: '' });
      return box;
    });
  }
  function printNow() {
    var gs = selected(groups()), kinds = kindsOn();
    if (!gs.length || !kinds.length) { toast('협력사와 문서 종류를 골라 주세요(③ 문서 만들기).', true); return; }
    gs.forEach(function (g) { numbersFor(g, kinds, true); });
    save(); render();
    setTimeout(function () { window.print(); }, 200);
  }
  function pagePrint() {
    if (!db.rows.length) return [pageHead('④', '인쇄 · PDF', ''), emptyNotice()];
    var gs = selected(groups()), kinds = kindsOn(), Sx = settings();
    var hasCustom = kinds.some(function (k) { return db.templates[k]; });
    var tabs = h('div', { class: 'tabs no-print', role: 'group', 'aria-label': '인쇄 모양' },
      [['html', '인쇄용 기본 양식'], ['xl', '엑셀 양식 모양(미리 보기)']].map(function (x) { return h('button', { type: 'button', 'aria-pressed': ui.printMode === x[0] ? 'true' : 'false', onclick: function () { ui.printMode = x[0]; render(); } }, x[1]); }));
    var docs = h('div', { class: 'docs' });
    var pending = 0;
    gs.forEach(function (g) {
      var nums = numbersFor(g, kinds, false);
      kinds.forEach(function (k) {
        if (!nums[k].reused) pending++;
        var ctx = contextFor(g, k, nums);
        if (ui.printMode === 'xl') add(docs, paperXl(k, ctx));
        else docs.appendChild(k === 'ag' ? paperAg(ctx, g, Sx) : paperInv(ctx, g, Sx));
      });
    });
    return [
      pageHead('④', '인쇄 · PDF', '고른 협력사의 합의서·인보이스를 한 문서씩 한 쪽으로 인쇄합니다. 인쇄 창에서 「PDF로 저장」을 고르면 PDF 가 됩니다.',
        h('div', { class: 'btn-row no-print' }, h('button', { type: 'button', class: 'btn btn-primary', disabled: !gs.length || !kinds.length, onclick: printNow }, '번호 확정 · 인쇄 / PDF'), h('a', { class: 'btn', href: '#/docs' }, '고르기로'))),
      h('section', { class: 'card no-print' },
        h('p', { class: 'note' }, '협력사 ' + gs.length + '곳 × ' + kinds.map(function (k) { return L.DOC_KINDS[k]; }).join('·') + ' = ' + (gs.length * kinds.length) + '쪽. ' + (pending ? '아직 확정하지 않은 번호가 ' + pending + '개 있습니다(인쇄 버튼을 누르면 확정).' : '모든 번호가 확정되어 있습니다.')),
        h('p', { class: 'note' }, ui.printMode === 'xl' ? '엑셀 양식을 채운 결과를 글자와 병합만 살려 보여 줍니다(테두리·색·수식 계산은 빠짐). 양식 그대로의 PDF 는 내려받은 엑셀을 열어 「PDF로 저장」해 주세요.' : '인쇄용 기본 양식은 이 도구가 그리는 합의서·인보이스입니다. 합의 조건 문구·입금 안내는 설정에서 고칩니다.' + (hasCustom ? ' 올린 엑셀 양식 모양은 위 「엑셀 양식 모양」에서 봅니다.' : '')),
        tabs),
      gs.length && kinds.length ? docs : h('section', { class: 'card empty' }, h('p', null, '③ 문서 만들기에서 협력사와 문서 종류를 골라 주세요.'))
    ];
  }

  /* ── ⑤ 양식 템플릿 ─────────────────────────── */
  function scanOf(kind) {
    var key = kind + ':' + (db.templates[kind] ? db.templates[kind].at : 'builtin');
    if (!ui.tplScan[key]) { try { ui.tplScan[key] = T.scanTemplate(tplBytes(kind)); } catch (e) { ui.tplScan[key] = { error: e.message }; } }
    return ui.tplScan[key];
  }
  function uploadTpl(kind, file) {
    var fr = new FileReader();
    fr.onload = function () {
      var u = new Uint8Array(fr.result), sc;
      try { sc = T.scanTemplate(u); } catch (e) { toast(file.name + ': ' + e.message, true); return; }
      if (!sc.count) { toast(file.name + ': {{키}} 를 하나도 찾지 못했습니다. 아래 키 목록을 양식 칸에 적어 주세요.', true); return; }
      db.templates[kind] = { name: file.name, b64: T.u8ToB64(u), at: new Date().toISOString().slice(0, 19).replace('T', ' '), size: u.length };
      if (!save()) toast('양식이 커서 브라우저에 저장하지 못했습니다. 이번 창에서만 씁니다.', true);
      else toast(L.DOC_KINDS[kind] + ' 양식을 바꿨습니다 — 키 ' + sc.count + '칸' + (sc.unknown.length ? ', 모르는 키 ' + sc.unknown.length + '개' : ''));
      render();
    };
    fr.readAsArrayBuffer(file);
  }
  function tplBox(kind) {
    var sc = scanOf(kind), t = db.templates[kind];
    var input = h('input', { type: 'file', accept: '.xlsx,.xlsm' });
    input.addEventListener('change', function () { if (input.files[0]) uploadTpl(kind, input.files[0]); input.value = ''; });
    var cells = sc.error ? [] : sc.sheets.reduce(function (a, s) { return a.concat(s.cells.map(function (c) { return [s.name, c.ref, h('span', { class: 'mono' }, c.text), c.keys.map(function (k) { return L.KEY_BY[k] ? (L.isRowKey(k) ? '반복' : '값') : '모르는 키'; }).join(', ')]; })); }, []);
    var rep = sc.error ? [] : sc.sheets.filter(function (s) { return s.repeatRow != null; });
    return h('div', { class: 'tpl-box' },
      h('h3', null, L.DOC_KINDS[kind], h('span', { class: 'tag' }, t ? '올린 양식' : '기본 양식')),
      h('p', { class: 'note' }, t ? t.name + ' · ' + Math.round(t.size / 1024) + 'KB · ' + t.at : B[kind].name + ' (가상 기본 양식)'),
      sc.error ? h('p', { class: 'alert warn' }, sc.error) : h('ul', { class: 'note' },
        h('li', null, '키가 있는 칸 ' + sc.count + '개 · 시트 ' + sc.sheets.length + '개'),
        h('li', null, rep.length ? '반복 줄: ' + rep.map(function (s) { return s.name + ' ' + s.repeatRow + '행(' + s.rowKeys.length + '칸)'; }).join(', ') + ' — 클레임 수만큼 내려 씁니다.' : '반복 줄이 없습니다. 클레임 목록을 넣으려면 한 줄에 {{행.클레임번호}} 같은 행 키를 적어 주세요.'),
        sc.unknown.length ? h('li', { class: 'alert warn' }, '모르는 키(그대로 남음): ' + sc.unknown.join(', ')) : null),
      h('div', { class: 'btn-row' },
        h('label', { class: 'btn btn-primary btn-sm' }, '회사 양식 올리기', input),
        h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { download(B[kind].name, new Blob([T.b64ToU8(B[kind].b64)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })); } }, '기본 양식 내려받기'),
        t ? h('button', { type: 'button', class: 'btn btn-sm btn-danger', onclick: function () { if (confirm('올린 양식을 빼고 기본 양식으로 돌아갈까요?')) { db.templates[kind] = null; save(); render(); } } }, '기본 양식으로') : null),
      cells.length ? h('details', { class: 'block', style: 'margin-top:10px' }, h('summary', null, '키가 있는 칸 ' + cells.length + '개'), table(['시트', '칸', '적힌 글', '구분'], cells)) : null);
  }
  function pageTpl() {
    return [
      pageHead('⑤', '양식 템플릿', '회사의 합의서·인보이스 엑셀 양식을 그대로 쓰고, 값이 들어갈 칸에만 {{키}} 를 적어 올립니다.'),
      h('div', { class: 'grid-2' }, tplBox('ag'), tplBox('inv')),
      h('section', { class: 'card' }, h('h2', null, '양식 만드는 법'),
        h('ol', { class: 'steps' },
          h('li', null, '지금 쓰는 합의서(또는 인보이스) 엑셀 양식을 복사해 둡니다. 서식·테두리·병합·로고·직인 그림은 그대로 둡니다.'),
          h('li', null, '값이 들어갈 칸에 아래 키를 적습니다. 칸 전체가 {{총액}} 처럼 키 하나면 숫자로 들어가 양식의 표시 형식(#,##0)이 그대로 쓰이고, 「합의 금액: {{총액_표기}}」처럼 글과 섞으면 글자로 들어갑니다.'),
          h('li', null, '클레임 목록 표는 첫 줄에만 {{행.번호}} {{행.클레임번호}} … 를 적습니다. 그 아래 빈 표 줄을 먼저 채우고, 모자라면 줄을 끼워 넣으며 아래의 합계 줄·병합·인쇄 영역·그림 위치를 함께 내립니다. 합계 칸의 =SUM(I14:I21) 같은 수식도 늘어난 줄까지 넓힙니다.'),
          h('li', null, '「회사 양식 올리기」로 넣고, ③ 문서 만들기에서 엑셀을 내려받아 확인합니다.')),
        h('p', { class: 'note' }, '안 되는 것: .xls(옛 형식)·문서보안(DRM) 파일, 도형·글상자 안의 키, 한 시트에 반복 표 두 개. 서식이 그대로인 PDF 는 내려받은 엑셀에서 「PDF로 저장」하고, 브라우저에서 바로 PDF 를 만들 때는 ④ 인쇄용 기본 양식을 씁니다.')),
      h('section', { class: 'card' }, h('h2', null, '키 목록 (' + L.KEYS.length + '개)'),
        h('div', { class: 'key-list' }, L.KEYS.map(function (k) {
          var code = '{{' + k.key + '}}';
          return h('div', { class: 'key-item' }, h('code', null, code), h('span', { class: 'd' }, (k.row ? '반복 줄 · ' : '') + (k.desc || ({ money: '금액(숫자)', int: '정수', pct: '비율', text: '글자', date: '날짜' })[k.type])),
            h('button', { type: 'button', class: 'btn', onclick: function () { copyText(code, code + ' 복사했습니다.'); } }, '복사'));
        })))
    ];
  }

  /* ── ⑥ 발행 이력 ─────────────────────────── */
  function pageHistory() {
    var list = db.history.slice().reverse();
    var rows = list.map(function (e) {
      var st = e.status === 'replaced' ? '대체됨' : e.status === 'void' ? '취소' : '발행';
      return { cells: [h('span', { class: 'mono' }, e.docNo), L.DOC_KINDS[e.kind], e.issueDate, e.supplier, e.count, money(e.total, e.currency) + ' ' + e.currency, st + (e.replaces ? ' (' + e.replaces + ' 대체)' : ''), e.times || 1, e.lastAt || e.at,
        e.status === 'issued' ? h('button', { type: 'button', class: 'btn btn-sm btn-danger', onclick: function () { if (confirm(e.docNo + ' 을(를) 취소로 표시할까요? 번호는 다시 쓰지 않습니다.')) { e.status = 'void'; save(); render(); } } }, '취소 표시') : ''],
        cls: e.status === 'issued' ? null : 'excluded' };
    });
    return [
      pageHead('⑥', '발행 이력', '내려받기·인쇄로 확정한 문서 번호입니다. 일련은 문서 종류마다, 발행일의 달마다 1부터 셉니다.',
        h('div', { class: 'btn-row no-print' },
          h('button', { type: 'button', class: 'btn', disabled: !list.length, onclick: function () { var wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([L.HISTORY_HEAD].concat(L.historyRows(db.history))), '발행 이력'); XLSX.writeFile(wb, '합의서_인보이스_발행이력_' + L.localToday() + '.xlsx'); } }, '엑셀로'),
          h('button', { type: 'button', class: 'btn btn-danger', disabled: !list.length, onclick: function () { if (confirm('발행 이력을 모두 지울까요? 다음 번호가 다시 001 부터 시작합니다.')) { db.history = []; save(); render(); } } }, '이력 비우기'))),
      list.length ? table(['문서 번호', '종류', '발행일', '협력사', '건수', '총액', '상태', '낸 횟수', '마지막', ''], rows, { num: [0, 0, 0, 0, 1, 1, 0, 1, 0, 0] })
        : h('section', { class: 'card empty' }, h('p', null, '아직 확정한 문서가 없습니다. ③ 문서 만들기에서 엑셀을 내려받거나 ④에서 인쇄하면 번호가 남습니다.'))
    ];
  }

  /* ── 설정 ─────────────────────────────────── */
  function pageSettings() {
    var s = settings();
    function inp(k, attrs) { var e = h('input', Object.assign({ value: s[k] }, attrs || {})); e.addEventListener('change', function () { db.settings[k] = e.value; db.settings = L.cleanSettings(db.settings); save(); render(); }); return e; }
    function sel(k, opts) { var e = h('select', null, opts.map(function (o) { var op = h('option', { value: String(o[0]) }, o[1]); if (String(s[k]) === String(o[0])) op.selected = true; return op; })); e.addEventListener('change', function () { db.settings[k] = e.value; db.settings = L.cleanSettings(db.settings); save(); render(); }); return e; }
    function area(k) { var e = h('textarea'); e.value = s[k]; e.addEventListener('change', function () { db.settings[k] = e.value; db.settings = L.cleanSettings(db.settings); save(); render(); }); return e; }
    var ex = L.formatDocNo(s.docFormat, { prefix: s.prefixAg, date: issueDate(), seq: 1 }) + ' · ' + L.formatDocNo(s.docFormat, { prefix: s.prefixInv, date: issueDate(), seq: 12 });
    var backup = h('input', { type: 'file', accept: '.json' });
    backup.addEventListener('change', function () {
      var f = backup.files[0]; if (!f) return;
      f.text().then(function (t) { db = L.restoreDb(JSON.parse(t)); ui.sel = null; save(); toast('백업을 불러왔습니다 — 합의 결과 ' + db.rows.length + '행.'); render(); }).catch(function (e) { toast('불러오지 못했습니다: ' + e.message, true); });
    });
    return [
      pageHead('설정', '설정', '발행처 정보·금액 규칙·문서 번호·문구. 바꾸면 모든 화면에 바로 반영됩니다.'),
      h('section', { class: 'card' }, h('h2', null, '발행처 (우리 회사)'),
        h('div', { class: 'form-grid' },
          fld('회사 이름 {{발행처}}', inp('issuer', { placeholder: '예: ○○(주)' })), fld('부서 {{발행부서}}', inp('dept')), fld('담당자 {{담당자}}', inp('person')),
          fld('연락처 {{연락처}}', inp('contact')), h('div', { class: 'span-all' }, fld('입금 계좌 {{입금계좌}}', inp('bank', { placeholder: '은행 · 계좌번호 · 예금주' }), '인보이스의 입금 안내에 들어갑니다')))),
      h('section', { class: 'card' }, h('h2', null, '금액 규칙'),
        h('div', { class: 'form-grid' },
          fld('기본 통화', inp('currency', { maxlength: 3 }), '합의 결과에 통화 열이 없을 때'),
          fld('부가세', sel('vat', [['none', '없음'], ['add10', '합의합계의 10% 별도']]), '보상금에 붙이는지 회사 기준 확인(기획서 10장 ③)'),
          fld('반올림', sel('roundMode', [['round', '반올림'], ['floor', '버림'], ['ceil', '올림']]), '비율로 계산한 금액·부가세에 적용'),
          fld('원 단위', sel('roundUnit', [[1, '1원'], [10, '10원'], [100, '100원'], [1000, '1,000원']]), '원화만. 다른 통화는 소수 둘째 자리'),
          fld('「비율 계산과 다름」 허용 차이(원)', inp('tolerance', { type: 'number', min: 0, step: 1 })),
          fld('지급 기한 (발행일 + 일)', inp('payDays', { type: 'number', min: 0, max: 365, step: 1 }), '{{지급기한}}'))),
      h('section', { class: 'card' }, h('h2', null, '문서 번호'),
        h('div', { class: 'form-grid' },
          fld('형식', inp('docFormat'), '{접두} {YYYY} {YY} {MM} {DD} {NNN}(일련 자릿수)'), fld('합의서 접두', inp('prefixAg')), fld('인보이스 접두', inp('prefixInv'))),
        h('p', { class: 'note', style: 'margin-top:8px' }, '보기: ' + ex)),
      h('section', { class: 'card' }, h('h2', null, '문구 (인쇄용 기본 양식)'),
        h('div', { class: 'form-grid' },
          h('div', { class: 'span-all' }, fld('합의 조건 — 한 줄에 한 조항, {{키}} 사용 가능', area('agClauses'))),
          h('div', { class: 'span-all' }, fld('인보이스 입금 안내', area('invNote'))))),
      h('section', { class: 'card' }, h('h2', null, '백업 · 불러오기'),
        h('p', { class: 'note' }, '데이터·발행 이력·올린 양식은 이 브라우저에만 있습니다. 다른 PC 로 옮기거나 지우기 전에 JSON 백업을 받아 두세요.'),
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn', onclick: function () { download('클레임보상문서_백업_' + L.localToday() + '.json', new Blob([JSON.stringify(db)], { type: 'application/json' })); } }, 'JSON 백업 내려받기'),
          h('label', { class: 'btn' }, 'JSON 백업 불러오기', backup)))
    ];
  }

  /* ── 라우터 ─────────────────────────────── */
  var PAGES = [
    { id: 'import', no: '0', t: '데이터 가져오기', s: '합의 결과 엑셀', fn: pageImport },
    { id: 'groups', no: '①', t: '협력사별 합의', s: '묶기·합계·고르기', fn: pageGroups },
    { id: 'rows', no: '②', t: '클레임 명세·검증', s: '금액 검사', fn: pageRows },
    { id: 'docs', no: '③', t: '문서 만들기', s: '엑셀 내려받기', fn: pageDocs },
    { id: 'print', no: '④', t: '인쇄 · PDF', s: '합의서·인보이스', fn: pagePrint },
    { id: 'tpl', no: '⑤', t: '양식 템플릿', s: '회사 양식 + {{키}}', fn: pageTpl },
    { id: 'history', no: '⑥', t: '발행 이력', s: '문서 번호', fn: pageHistory },
    { id: 'settings', no: '설', t: '설정', s: '발행처·금액·번호', fn: pageSettings }
  ];
  function route() { var id = (location.hash || '').replace(/^#\//, '').split('?')[0]; return PAGES.filter(function (p) { return p.id === id; })[0] || (db.rows.length ? PAGES[1] : PAGES[0]); }
  function render() {
    var p = route();
    var nav = document.getElementById('nav'); nav.innerHTML = '';
    PAGES.forEach(function (x) {
      nav.appendChild(h('a', { href: '#/' + x.id, 'aria-current': x === p ? 'page' : null }, h('span', { class: 'no' }, x.no), h('span', { class: 't' }, x.t), h('span', { class: 's' }, x.s)));
    });
    var chip = document.getElementById('dataChip');
    if (db.rows.length) {
      var gs = groups();
      chip.textContent = '합의 결과 ' + db.rows.length + '행 · 협력사 ' + gs.length + '곳 · 발행일 ' + issueDate();
    } else chip.textContent = '데이터 없음';
    var b = document.getElementById('sampleBanner');
    b.hidden = !db.sample;
    b.textContent = '지금 보시는 것은 가상 샘플입니다(협력사·클레임 번호·품번·금액·발행처·계좌 모두 지어낸 값). 양식도 가상 기본 양식입니다. 실제 합의 결과를 넣으면 샘플은 지워집니다.';
    main.innerHTML = '';
    add(main, p.fn());
    document.title = p.t + ' — 클레임 보상 문서';
    save();
  }
  window.addEventListener('hashchange', function () { render(); main.focus(); window.scrollTo(0, 0); });
  render();
})();
