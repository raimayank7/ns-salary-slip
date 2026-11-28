/* The whole month's salary from Excel, on salary-slip.niveshsahayak.in (byof-salary, 28 Nov 2026).
   Step 1: month, employer, state for professional tax, the employee sheet (Excel or CSV, any headings) and an optional attendance sheet.
   Step 2: what was read, the columns used (changeable, remembered for the next file with the same headings), problems row by row.
   Step 3: every payslip as one PDF, one PDF each (zip), Print all, the salary register (Excel), a bank transfer sheet, the PF ECR text
   file and the ESI monthly contribution file (.xls), with the PF and ESI challan totals.
   Nothing leaves the browser: the files are made here and saved to the visitor's computer. The slips are drawn with the salary slip
   maker's own printed sheet (slip.js NSSlipOut); pay rules come from salcore.js (NSSalBulk) and payrollcore.js (NSPayroll).
   Screen words: NS.t('slb.*'); the files and the printed slips are in English (like every printed document here).
   Designed & Developed by: Nivesh Sahayak (https://www.niveshsahayak.in/) */
(function () {
  'use strict';
  var $ = NS.$, $$ = NS.$$, esc = NS.esc;
  var SB = window.NSSalBulk, P = window.NSPayroll;
  var SHEETJS = '/assets/vendor/xlsx-0.20.3.full.min.js';
  var LKEY = 'ns-salbulk-layouts', SKEY = 'ns-salbulk-set';
  function L(k, v) { return NS.t('slb.' + k, v); }
  var ST = { book: null, att: null, attName: '', fileName: '', force: null, lines: [], D: null, A: null, sample: false, bizes: [] };

  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function lastMonth() { var d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1); return d.getFullYear() + '-' + pad(d.getMonth() + 1); }
  function get(k, d) { try { var v = JSON.parse(localStorage.getItem(k) || 'null'); return v == null ? d : v; } catch (e) { return d; } }
  function put(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
  function money(n) { return NS.money(n, false); }
  function save(blob, name) { if (NS.saveFile) NS.saveFile(blob, name); else { var a = document.createElement('a'), u = URL.createObjectURL(blob); a.href = u; a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(function () { URL.revokeObjectURL(u); }, 4000); } NS.toast(L('saved', { f: name })); }
  function setOf() { var s = get(SKEY, {}); return Object.assign({ mode: 'full', state: '', tds: false, pffull: false, esi: true }, s); }

  /* ---------- the section ---------- */
  function html() {
    var states = (NS.STATES || []).map(function (x) { return '<option value="' + x[0] + '">' + esc(x[1]) + '</option>'; }).join('');
    return '<h2 id="slb-h">' + esc(L('h')) + '</h2><p>' + esc(L('lead')) + '</p><p class="soft">' + esc(L('private')) + '</p>' +
      '<div class="bk-steps">' +
      '<div class="card"><h3>' + esc(L('s1')) + '</h3>' +
        '<div class="slb-f"><label>' + esc(L('month')) + '<input type="month" id="slb-month" value="' + lastMonth() + '"></label>' +
        '<label>' + esc(L('emp')) + '<select id="slb-biz"><option value="">' + esc(L('emp_typed')) + '</option></select></label>' +
        '<label>' + esc(L('empname')) + '<input id="slb-bname" autocomplete="organization"></label>' +
        '<label>' + esc(L('empaddr')) + '<input id="slb-baddr" autocomplete="street-address"></label>' +
        '<label>' + esc(L('state')) + '<select id="slb-state"><option value="">' + esc(L('state_none')) + '</option>' + states + '</select></label>' +
        '<label>' + esc(L('mode')) + '<select id="slb-mode"><option value="full">' + esc(L('mode_full')) + '</option><option value="earned">' + esc(L('mode_earned')) + '</option></select></label></div>' +
        '<label class="chk"><input type="checkbox" id="slb-esion"> ' + esc(L('esion')) + '</label> <label class="chk"><input type="checkbox" id="slb-tds"> ' + esc(L('tdsest')) + '</label> <label class="chk"><input type="checkbox" id="slb-pffull"> ' + esc(L('pffull')) + '</label>' +
        '<div class="row"><label class="btn primary" for="slb-file">' + esc(L('choose')) + '</label><input type="file" id="slb-file" accept=".xlsx,.xls,.csv,.txt" hidden>' +
        '<label class="btn" for="slb-afile">' + esc(L('chooseatt')) + '</label><input type="file" id="slb-afile" accept=".xlsx,.xls,.csv,.txt" hidden>' +
        '<button type="button" class="btn" data-slb="sample">' + esc(L('try')) + '</button><button type="button" class="btn sm" data-slb="template">' + esc(L('tpl')) + '</button></div>' +
        '<div class="slb-drop" id="slb-drop" tabindex="0">' + esc(L('drop')) + '</div>' +
        '<p class="soft" id="slb-files"></p></div>' +
      '<div class="card" id="slb-s2" hidden><h3>' + esc(L('s2')) + '</h3><div class="bk-out" id="slb-read"></div></div>' +
      '<div class="card" id="slb-s3" hidden><h3>' + esc(L('s3')) + '</h3><div class="bk-out" id="slb-out"></div></div>' +
      '</div>';
  }
  function box() {
    var b = $('[data-salbulk]');
    if (!b) {
      var tool = $('#tool'); if (!tool) return null;
      b = document.createElement('section'); b.className = 'bulk wrap'; b.id = 'bulk'; b.setAttribute('data-salbulk', '');
      tool.parentNode.insertBefore(b, tool.nextSibling);
    }
    b.setAttribute('aria-labelledby', 'slb-h');
    return b;
  }

  /* ---------- reading files ---------- */
  /* every sheet of an Excel file (SheetJS, kept inside NSBiz); a CSV, or Excel when the library cannot load (offline first visit), through NS.xlsx */
  function readBook(file) {
    var name = String(file.name || ''), csv = /\.(csv|txt)$/i.test(name);
    function plain() { return NS.xlsx.read(file).then(function (rows) { return [{ name: name.replace(/\.[^.]+$/, '') || 'Sheet1', rows: rows }]; }); }
    if (csv) return plain();
    return NS.loadJs(SHEETJS).then(function () {
      return file.arrayBuffer().then(function (buf) {
        var X = window.XLSX, wb = X.read(buf, { type: 'array', cellDates: false });
        return wb.SheetNames.map(function (n) { return { name: n, rows: X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }).map(function (r) { return r.map(function (c) { return String(c == null ? '' : c).trim(); }); }).filter(function (r) { return r.some(Boolean); }) }; });
      });
    }, plain);
  }
  function takeFile(file, att) {
    if (!file) return;
    return readBook(file).then(function (book) {
      if (att) { ST.att = book; ST.attName = file.name; } else { ST.book = book; ST.fileName = file.name; ST.force = null; ST.sample = false; ST.att = null; ST.attName = ''; }   /* a new employee file starts without the old attendance */
      run();
    }, function (e) { NS.toast(String(e && e.message) === 'xls' ? L('err_xls') : L('err_read'), 'bad'); });
  }

  /* ---------- working out ---------- */
  function opts() {
    var S = window.NSSlipOut || {};
    return { ym: $('#slb-month').value || lastMonth(), mode: $('#slb-mode').value, state: $('#slb-state').value, tdsEst: $('#slb-tds').checked, esiOn: $('#slb-esion').checked, pfFull: $('#slb-pffull').checked,
      att: ST.A, ptFor: S.ptFor, tdsFor: S.tdsMonth };
  }
  function run() {
    var o = opts();
    put(SKEY, { mode: o.mode, state: o.state, tds: o.tdsEst, pffull: o.pfFull, esi: o.esiOn });
    $('#slb-files').textContent = [ST.fileName ? L('file', { f: ST.fileName }) : '', ST.attName ? L('attfile', { f: ST.attName }) : ''].filter(Boolean).join(' · ');
    if (!ST.book) return;
    ST.D = SB.read(ST.book, { layouts: get(LKEY, {}), force: ST.force });
    /* attendance: the separate file, else an attendance sheet inside the same workbook */
    var empSheet = (ST.D.found.filter(function (x) { return x.kind === 'emp'; })[0] || {}).name;
    var inBook = ST.book.filter(function (sh) { return sh.name !== empSheet && /attend|muster|present|haziri/i.test(sh.name); });
    ST.A = ST.att ? SB.readAtt(ST.att, o.ym, ST.D.emp) : inBook.length ? SB.readAtt(inBook, o.ym, ST.D.emp) : null; o.att = ST.A;
    ST.lines = SB.work(ST.D.emp, o);
    drawRead(); drawOut();
  }
  function msg(c) { return L('e_' + c[0], c[1] || {}); }
  function drawRead() {
    var D = ST.D, lines = ST.lines, f = D.found.filter(function (x) { return x.kind === 'emp'; })[0];
    var bad = lines.filter(function (l) { return !l.ok; }), ff = lines.filter(function (l) { return l.ok && l.ferr.length; }), ww = lines.filter(function (l) { return l.warn.length; });
    var h = '';
    if (!f) {
      h = '<p class="warn">' + esc(L('nosheet')) + '</p>';
      $('#slb-s2').hidden = false; $('#slb-read').innerHTML = h; $('#slb-s3').hidden = true; return;
    }
    h += '<p>' + esc(L('read', { n: lines.length, s: f.name, r: f.headRow + 1 })) + (ST.sample ? ' <b>' + esc(L('samplenote')) + '</b>' : '') + '</p>';
    var skipped = D.found.filter(function (x) { return x.kind !== 'emp' && x.rows; });
    if (D.totals && D.totals.length) h += '<p class="soft">' + esc(L('totals', { l: D.totals.join(', ') })) + '</p>';
    if (skipped.length) h += '<p class="soft">' + esc(L('otherSheets', { l: skipped.map(function (x) { return x.name; }).join(', ') })) + '</p>';
    if (D.probs.length) h += '<ul class="warn">' + D.probs.map(function (p) { return '<li>' + esc(L('nocol', { c: p.msg.replace(/^No column for |: choose it with Columns\.$/g, '') })) + '</li>'; }).join('') + '</ul>';
    if (ST.A) h += '<p>' + esc(L('attread', { n: ST.A.matched })) + (ST.A.unknown.length ? ' <span class="warn">' + esc(L('attunknown', { l: ST.A.unknown.slice(0, 8).join(', ') })) + '</span>' : '') + (ST.A.bad.length ? ' <span class="warn">' + esc(L('attbad', { l: ST.A.bad.slice(0, 5).join('; ') })) + '</span>' : '') + '</p>';
    /* the columns used */
    var heads = f.heads || [], opt = function (sel) { return '<option value="">' + esc(L('col_none')) + '</option>' + heads.map(function (x, i) { return x ? '<option value="' + i + '"' + (sel === i ? ' selected' : '') + '>' + esc(x) + '</option>' : ''; }).join(''); };
    h += '<details class="slb-cols"' + (D.probs.length ? ' open' : '') + '><summary>' + esc(L('cols', { n: Object.keys(f.map).length })) + (f.remembered ? ' · ' + esc(L('remembered')) : '') + '</summary><div class="slb-colgrid">' +
      SB.EMP.cols.map(function (c) { return '<label>' + esc(c.h) + (c.need ? ' *' : '') + '<select data-col="' + c.k + '">' + opt(f.map[c.k]) + '</select></label>'; }).join('') + '</div>' +
      '<p class="soft">' + esc(L('colhelp')) + '</p><button type="button" class="btn sm" data-slb="remember">' + esc(L('rememberbtn')) + '</button></details>';
    /* problems first */
    var row = function (l) { return L('rowof', { r: l.row, n: l.name || '?' }); };
    if (bad.length) h += '<h4 class="warn">' + esc(L('errs', { n: bad.length })) + '</h4><ul class="warn">' + bad.map(function (l) { return '<li>' + esc(row(l)) + ': ' + esc(l.errs.map(msg).join('; ')) + '</li>'; }).join('') + '</ul>';
    if (ff.length) h += '<h4>' + esc(L('ferrs', { n: ff.length })) + '</h4><ul class="slb-ferr">' + ff.map(function (l) { return '<li>' + esc(row(l)) + ': ' + esc(l.ferr.map(msg).join('; ')) + '</li>'; }).join('') + '</ul>';
    if (ww.length) h += '<details><summary>' + esc(L('warns', { n: ww.length })) + '</summary><ul>' + ww.map(function (l) { return '<li>' + esc(row(l)) + ': ' + esc(l.warn.map(msg).join('; ')) + '</li>'; }).join('') + '</ul></details>';
    if (!bad.length && !ff.length) h += '<p class="ok">' + esc(L('allok')) + '</p>';
    /* the pay of everyone (first 300 rows on the screen; the files have all) */
    var show = lines.slice(0, 300);
    h += '<div class="tbl"><table class="bk-t"><thead><tr><th>' + esc(L('t_row')) + '</th><th>' + esc(L('t_name')) + '</th><th class="n">' + esc(L('t_paid')) + '</th><th class="n">' + esc(L('t_gross')) + '</th><th class="n">' + esc(L('t_pf')) + '</th><th class="n">' + esc(L('t_esi')) + '</th><th class="n">' + esc(L('t_ded')) + '</th><th class="n">' + esc(L('t_net')) + '</th></tr></thead><tbody>' +
      show.map(function (l) { var d = function (n) { var x = l.ded.filter(function (y) { return y.desc === n; })[0]; return x ? money(x.amt) : ''; };
        return '<tr class="' + (l.ok ? 'ok' : 'bad') + '"><td>' + l.row + '</td><td>' + esc(l.name) + (l.code ? ' <span class="soft">' + esc(l.code) + '</span>' : '') + '</td><td class="n">' + l.paid + ' / ' + l.days + '</td><td class="n">' + money(l.gross) + '</td><td class="n">' + d('Provident fund (PF)') + '</td><td class="n">' + d('ESI') + '</td><td class="n">' + money(l.tded) + '</td><td class="n"><b>' + money(l.net) + '</b></td></tr>'; }).join('') +
      '</tbody></table></div>' + (lines.length > show.length ? '<p class="soft">' + esc(L('more', { n: lines.length - show.length })) + '</p>' : '');
    $('#slb-s2').hidden = false; $('#slb-read').innerHTML = h;
  }
  function drawOut() {
    var ok = ST.lines.filter(function (l) { return l.ok; });
    if (!ST.D || !ok.length) { $('#slb-s3').hidden = !ST.D; $('#slb-out').innerHTML = ST.D ? '<p class="warn">' + esc(L('nothing')) + '</p>' : ''; return; }
    var pf = SB.pfLines(ST.lines), es = SB.esiLines(ST.lines), pc = P.pfChallan(pf), ec = P.esiChallan(es);
    var net = ok.reduce(function (s, l) { return s + l.net; }, 0), gross = ok.reduce(function (s, l) { return s + l.gross; }, 0);
    var h = '<p>' + esc(L('ready', { n: ok.length, g: NS.money(gross), t: NS.money(net) })) + (ok.length < ST.lines.length ? ' <span class="warn">' + esc(L('leftout', { n: ST.lines.length - ok.length })) + '</span>' : '') + '</p>' +
      '<div class="row"><button type="button" class="btn primary" data-slb="pdf">' + esc(L('b_pdf')) + '</button><button type="button" class="btn" data-slb="zip">' + esc(L('b_zip')) + '</button><button type="button" class="btn" data-slb="print">' + esc(L('b_print')) + '</button></div>' +
      (ok.length > 300 ? '<p class="soft">' + esc(L('bigpdf')) + '</p>' : '') +
      '<div class="row"><button type="button" class="btn" data-slb="reg">' + esc(L('b_reg')) + '</button><button type="button" class="btn" data-slb="bank">' + esc(L('b_bank')) + '</button></div>' +
      '<div class="row"><button type="button" class="btn" data-slb="ecr"' + (pf.length ? '' : ' disabled') + '>' + esc(L('b_ecr', { n: pf.length })) + '</button><button type="button" class="btn" data-slb="esi"' + (es.length ? '' : ' disabled') + '>' + esc(L('b_esi', { n: es.length })) + '</button></div>' +
      '<div class="slb-ch"><div><b>' + esc(L('ch_pf')) + '</b><ul class="rows"><li><span>' + esc(L('ch_members')) + '</span><b>' + pc.members + '</b></li><li><span>A/c 01 (' + esc(L('ch_ee')) + ' + ' + esc(L('ch_er')) + ')</span><b>' + NS.money(pc.ac01.total) + '</b></li><li><span>A/c 10 (EPS)</span><b>' + NS.money(pc.ac10) + '</b></li><li><span>A/c 02 (' + esc(L('ch_admin')) + ')</span><b>' + NS.money(pc.ac02) + '</b></li><li><span>A/c 21 (EDLI)</span><b>' + NS.money(pc.ac21) + '</b></li><li><span>' + esc(L('ch_total')) + '</span><b>' + NS.money(pc.total) + '</b></li></ul></div>' +
      '<div><b>' + esc(L('ch_esi')) + '</b><ul class="rows"><li><span>' + esc(L('ch_members')) + '</span><b>' + ec.members + '</b></li><li><span>' + esc(L('ch_ee')) + '</span><b>' + NS.money(ec.ee) + '</b></li><li><span>' + esc(L('ch_er')) + '</span><b>' + NS.money(ec.er) + '</b></li><li><span>' + esc(L('ch_total')) + '</span><b>' + NS.money(ec.total) + '</b></li></ul></div></div>' +
      '<p class="soft">' + esc(L('upload')) + '</p><p class="soft">' + esc(L('check')) + '</p>';
    $('#slb-s3').hidden = false; $('#slb-out').innerHTML = h;
  }

  /* ---------- the payslips ---------- */
  function seller() {
    var id = $('#slb-biz').value, b = ST.bizes.filter(function (x) { return x.id === id; })[0];
    var s = b ? Object.assign({}, b) : {};
    var n = $('#slb-bname').value.trim(), a = $('#slb-baddr').value.trim();
    if (n) s.name = n; if (a) s.addr = a;
    return s;
  }
  function slipHtmls() {
    var o = opts(), S = window.NSSlipOut, sel = seller(), n = 0, ok = ST.lines.filter(function (l) { return l.ok; });
    if (!S) return [];
    return ok.map(function (l) { n++; return { l: l, h: S.html(SB.slipDoc(l, sel, o.ym, 'SAL/' + o.ym.replace('-', '') + '/' + (l.code || pad(n)))) }; });
  }
  function fname(l, ym) { return ('Salary slip ' + ym + ' ' + (l.code ? l.code + ' ' : '') + l.name).replace(/[\\\/:*?"<>|]+/g, '-').replace(/\s+/g, ' ').trim().slice(0, 90); }
  function needName() { if (seller().name) return true; NS.toast(L('needemp'), 'bad'); $('#slb-bname').focus(); return false; }
  function stage(list) {
    var hold = document.createElement('div'); hold.style.cssText = 'position:fixed;left:-12000px;top:0;width:794px'; hold.setAttribute('aria-hidden', 'true');
    hold.innerHTML = list.map(function (x) { return x.h; }).join(''); document.body.appendChild(hold);
    return hold;
  }
  /* a small bar with the count and Cancel (same look as NS.eachSeq's) */
  function bar(label, n) {
    var b = document.createElement('div'), stop = { yes: false };
    b.className = 'ns-prog'; b.setAttribute('role', 'status');
    b.innerHTML = '<span class="ns-prog-t"></span><span class="ns-prog-b"><i></i></span><button type="button" class="btn sm">' + esc(NS.t('prog.cancel')) + '</button>';
    b.querySelector('button').addEventListener('click', function () { stop.yes = true; this.disabled = true; });
    document.body.appendChild(b);
    stop.step = function (i) { if (stop.yes) throw new Error('cancel'); b.querySelector('.ns-prog-t').textContent = NS.t('prog.of', { l: label, a: i, n: n }); b.querySelector('i').style.width = Math.round(i / n * 100) + '%'; return new Promise(function (r) { setTimeout(r, 0); }); };
    stop.close = function () { b.remove(); };
    return stop;
  }
  function onePdf() {
    if (!needName()) return;
    var list = slipHtmls(), ym = opts().ym, hold = stage(list), b = bar(L('making'), list.length);
    NS.makePdf($$('.sh', hold), 'a4', { step: b.step }).then(function (blob) { save(blob, 'Salary slips ' + ym + '.pdf'); }, function (e) { if (!/cancel/.test(e && e.message)) NS.toast(L('pdffail'), 'bad'); })
      .then(function () { hold.remove(); b.close(); });
  }
  function zipPdfs() {
    if (!needName()) return;
    var list = slipHtmls(), ym = opts().ym, files = {}, used = {};
    NS.eachSeq(list, function (x) {
      var hold = stage([x]);
      return NS.makePdf($$('.sh', hold), 'a4').then(function (blob) { return blob.arrayBuffer(); }).then(function (buf) {
        var nm = fname(x.l, ym), k = nm, i = 2; while (used[k]) k = nm + ' (' + (i++) + ')'; used[k] = 1;
        files[k + '.pdf'] = new Uint8Array(buf); hold.remove();
      }, function (e) { hold.remove(); throw e; });
    }, { label: L('making') }).then(function (r) {
      if (!Object.keys(files).length) return;
      save(NS.xlsx.zip(files), 'Salary slips ' + ym + (r.cancelled ? ' (part)' : '') + '.zip');
    }, function () { NS.toast(L('pdffail'), 'bad'); });
  }
  function printAll() {
    if (!needName()) return;
    var sh = $('#sheet'), tool = $('#tool'); if (!sh) return;
    var keep = sh.innerHTML, tab = tool.getAttribute('data-tab'), old = document.title;
    sh.innerHTML = slipHtmls().map(function (x) { return x.h; }).join('');
    tool.setAttribute('data-tab', 'preview'); document.title = 'Salary slips ' + opts().ym;
    var done = function () { window.removeEventListener('afterprint', done); sh.innerHTML = keep; tool.setAttribute('data-tab', tab); document.title = old; window.dispatchEvent(new Event('resize')); };
    window.addEventListener('afterprint', done);
    setTimeout(function () { window.print(); }, 80);
  }

  /* ---------- the files ---------- */
  function xlsxOut(rows, name, sheet) { save(NS.xlsx.write([[sheet, rows]]), name); }
  function ecr() { var t = P.pfFile(SB.pfLines(ST.lines)); save(new Blob([t], { type: 'text/plain' }), 'ECR ' + opts().ym + '.txt'); }
  function esiFile() {
    var rows = P.esiRows(SB.esiLines(ST.lines)), fn = 'ESI contribution ' + opts().ym + '.xls';
    NS.loadJs(SHEETJS).then(function () {
      var X = window.XLSX, ws = X.utils.aoa_to_sheet(rows), r = X.utils.decode_range(ws['!ref']);
      for (var R = r.s.r; R <= r.e.r; R++) for (var C = r.s.c; C <= r.e.c; C++) { var c = ws[X.utils.encode_cell({ r: R, c: C })]; if (c) { c.t = 's'; c.v = String(c.v); c.z = '@'; } }
      var wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, 'Sheet1');
      save(new Blob([X.write(wb, { bookType: 'biff8', type: 'array' })], { type: 'application/vnd.ms-excel' }), fn);
    }, function () { NS.toast(L('err_lib'), 'bad'); });
  }

  /* ---------- wiring ---------- */
  function wire(b) {
    var s = setOf();
    $('#slb-mode').value = s.mode; $('#slb-state').value = s.state; $('#slb-tds').checked = !!s.tds; $('#slb-pffull').checked = !!s.pffull; $('#slb-esion').checked = s.esi !== false;
    $('#slb-file').addEventListener('change', function () { takeFile(this.files[0], false); this.value = ''; });
    $('#slb-afile').addEventListener('change', function () { takeFile(this.files[0], true); this.value = ''; });
    ['slb-month', 'slb-mode', 'slb-state', 'slb-tds', 'slb-pffull', 'slb-esion'].forEach(function (id) { $('#' + id).addEventListener('change', run); });
    $('#slb-biz').addEventListener('change', function () { var x = ST.bizes.filter(function (y) { return y.id === this.value; }, this)[0]; $('#slb-bname').value = x ? x.name || '' : ''; $('#slb-baddr').value = x ? [x.addr, x.city].filter(Boolean).join(', ') : ''; if (x && x.state && !$('#slb-state').value) { $('#slb-state').value = x.state; run(); } });
    var drop = $('#slb-drop');
    ['dragenter', 'dragover'].forEach(function (e) { drop.addEventListener(e, function (ev) { ev.preventDefault(); drop.classList.add('on'); }); });
    ['dragleave', 'drop'].forEach(function (e) { drop.addEventListener(e, function () { drop.classList.remove('on'); }); });
    drop.addEventListener('drop', function (ev) { ev.preventDefault(); var f = ev.dataTransfer && ev.dataTransfer.files; if (!f || !f.length) return; takeFile(f[0], !!(ST.book && f.length === 1 && /attend|muster|present/i.test(f[0].name))); if (f[1]) takeFile(f[1], true); });
    drop.addEventListener('keydown', function (ev) { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); $('#slb-file').click(); } });
    if (window.NSStore) NSStore.list('biz').then(function (l) {
      ST.bizes = (l || []).filter(function (x) { return x && x.name; });
      var sel = $('#slb-biz'); ST.bizes.forEach(function (x) { var o = document.createElement('option'); o.value = x.id; o.textContent = x.name; sel.appendChild(o); });
      if (ST.bizes.length === 1) { sel.value = ST.bizes[0].id; sel.dispatchEvent(new Event('change')); }
    }).catch(function () {});
    if (b._slbWired) return; b._slbWired = true;   /* the section is redrawn when the language changes; its own listeners stay once */
    b.addEventListener('change', function (e) {
      var c = e.target.closest('[data-col]'); if (!c || !ST.D) return;
      var f = ST.D.found.filter(function (x) { return x.kind === 'emp'; })[0], map = {};
      $$('[data-col]', b).forEach(function (s) { if (s.value !== '') map[s.getAttribute('data-col')] = +s.value; });
      ST.force = {}; ST.force[f.name] = { kind: 'emp', row: f.headRow, map: map, sig: f.sig, heads: f.heads }; run();
    });
    b.addEventListener('click', function (e) {
      var x = e.target.closest('[data-slb]'); if (!x) return;
      var a = x.getAttribute('data-slb');
      if (a === 'sample') { ST.book = SB.sampleRows(); ST.att = null; ST.attName = ''; ST.fileName = L('samplefile'); ST.force = null; ST.sample = true; if (!$('#slb-state').value) $('#slb-state').value = '27'; if (!$('#slb-bname').value) $('#slb-bname').value = L('samplebiz'); run(); $('#slb-s2').scrollIntoView({ block: 'start' }); }
      else if (a === 'template') save(NS.xlsx.write(SB.sampleBook(opts().ym)), 'Salary sheet - sample.xlsx');
      else if (a === 'remember') { var f = ST.D && ST.D.found.filter(function (y) { return y.kind === 'emp'; })[0]; if (!f) return; var all = get(LKEY, {}); all[f.sig] = { kind: 'emp', map: (ST.force && ST.force[f.name] ? ST.force[f.name].map : f.map) }; put(LKEY, all); NS.toast(L('rememberok')); }
      else if (a === 'pdf') onePdf();
      else if (a === 'zip') zipPdfs();
      else if (a === 'print') printAll();
      else if (a === 'reg') xlsxOut(SB.register(ST.lines.filter(function (l) { return l.ok; })), 'Salary register ' + opts().ym + '.xlsx', 'Salary register');
      else if (a === 'bank') xlsxOut(SB.bankRows(ST.lines.filter(function (l) { return l.ok; })), 'Salary bank transfer ' + opts().ym + '.xlsx', 'Bank transfer');
      else if (a === 'ecr') ecr();
      else if (a === 'esi') esiFile();
    });
  }
  function draw() {
    var b = box(); if (!b || !SB || !P) return;
    var wasBook = ST.book, keep = {};
    ['slb-month', 'slb-bname', 'slb-baddr'].forEach(function (id) { var x = $('#' + id); if (x) keep[id] = x.value; });
    b.innerHTML = html(); wire(b);
    Object.keys(keep).forEach(function (id) { if (keep[id]) $('#' + id).value = keep[id]; });
    if (wasBook) run();
  }
  function init() { draw(); document.addEventListener('ns-lang', draw); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
