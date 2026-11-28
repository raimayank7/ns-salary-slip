/* NSSalBulk: a whole month's salary from one Excel or CSV sheet of employees (byof-salary, 28 Nov 2026). Pure: no screen, no storage
   (node scripts/test_salbulk.js). salbulk.js is the screen on salary-slip.niveshsahayak.in.
   Reading: NSTpl.readAs with the two field sets below (any heading row in the first 15 rows, headings matched by name and the usual
   other words; a remembered layout is used again). Pay: PF and ESI lines come from NSPayroll (payrollcore.js, the same rules as the
   App's PF ECR and ESI file screens); professional tax and the TDS estimate come from the salary slip maker (slip.js ptFor / tdsMonth,
   handed in as opts.ptFor / opts.tdsFor), so all three tools give the same figures.
   Rules used (checked 28 Nov 2026, see payrollcore.js for sources): EPF 12% of basic + DA up to the wage ceiling of Rs 25,000 (raised
   from Rs 15,000 on 17 Sep 2026, S.O. 5109(E)); EPS 8.33% within the ceiling; ESI 0.75% employee / 3.25% employer when gross pay is
   Rs 21,000 a month or less, rounded up to the next rupee; no employee ESI when the average daily wage is Rs 176 or less.
   Designed & Developed by: Nivesh Sahayak (https://www.niveshsahayak.in/) */
(function (G) {
  'use strict';
  /* the employee sheet: k field, h heading in our sample sheet, need must be there, w what to type, ex example, a other headings */
  var EMP = { k: 'emp', name: 'Employees', words: /employee|staff|salary|payroll|wage|register|pay/, cols: [
    { k: 'code', h: 'Employee ID', w: 'Your own code for the employee', ex: 'E001', a: ['emp id', 'emp code', 'employee code', 'code', 'id', 'staff id', 'emp no', 'employee no', 'employee number', 'emp number', 'staff code', 'pay code', 'sr no emp', 'empid', 'empcode', 'payroll no'] },
    { k: 'name', h: 'Name', need: 1, w: 'Full name as it should print', ex: 'Ramesh Kumar', a: ['employee name', 'emp name', 'staff name', 'name of employee', 'employee', 'full name', 'member name'] },
    { k: 'desig', h: 'Designation', w: '', ex: 'Accountant', a: ['post', 'role', 'job title', 'designation name'] },
    { k: 'dept', h: 'Department', w: '', ex: 'Accounts', a: ['dept', 'division', 'section'] },
    { k: 'doj', h: 'Date of joining', w: 'dd-mm-yyyy', ex: '01-04-2024', a: ['doj', 'joining date', 'date of join', 'joined on'] },
    { k: 'uan', h: 'UAN', w: '12 digits (PF members)', ex: '100123456789', a: ['uan no', 'uan number', 'pf uan', 'universal account number'] },
    { k: 'esino', h: 'ESI No', w: '10-digit IP number (ESI members)', ex: '3112345678', a: ['esi number', 'esic no', 'esic number', 'ip number', 'ip no', 'insurance number', 'esi ip'] },
    { k: 'pan', h: 'PAN', w: '', ex: 'ABCPK1234F', a: ['pan no', 'pan number'] },
    { k: 'bank', h: 'Bank', w: 'Bank name', ex: 'State Bank of India', a: ['bank name'] },
    { k: 'acc', h: 'Account No', w: '', ex: '30123456789', a: ['bank account', 'account number', 'a/c no', 'ac no', 'bank a/c', 'bank account no'] },
    { k: 'ifsc', h: 'IFSC', w: '', ex: 'SBIN0001234', a: ['ifsc code', 'bank ifsc'] },
    { k: 'days', h: 'Days in month', w: 'Empty = the calendar days of the month', ex: '30', a: ['days', 'month days', 'total days', 'days of month', 'working days in month'] },
    { k: 'paid', h: 'Paid days', w: 'Days worked + paid leave + holidays (empty = all days)', ex: '28', a: ['days paid', 'present days', 'days present', 'payable days', 'days worked', 'attendance', 'present', 'pay days', 'no of days'] },
    { k: 'lop', h: 'Loss of pay days', w: 'Only if you have no Paid days column', ex: '2', a: ['lop', 'lop days', 'absent days', 'leave without pay', 'lwp', 'unpaid days'] },
    { k: 'basic', h: 'Basic', need: 1, w: 'Monthly basic', ex: '15000', a: ['basic salary', 'basic pay', 'basic wages'] },
    { k: 'da', h: 'DA', w: 'Dearness allowance', ex: '2000', a: ['dearness allowance', 'vda', 'variable da'] },
    { k: 'hra', h: 'HRA', w: '', ex: '6000', a: ['house rent allowance', 'house rent'] },
    { k: 'conv', h: 'Conveyance', w: '', ex: '1600', a: ['conveyance allowance', 'transport allowance', 'travel allowance', 'ta'] },
    { k: 'special', h: 'Special allowance', w: '', ex: '3000', a: ['special', 'spl allowance', 'special pay'] },
    { k: 'other', h: 'Other allowances', w: 'Any other fixed allowance', ex: '500', a: ['other allowance', 'allowances', 'misc allowance', 'medical allowance', 'other earnings'] },
    { k: 'extra', h: 'Overtime / bonus', w: 'This month only, not cut for days', ex: '1200', a: ['overtime', 'ot', 'ot amount', 'bonus', 'incentive', 'arrears', 'extra pay'] },
    { k: 'pf', h: 'PF member', w: 'Y or N (empty = Y when there is a UAN)', ex: 'Y', a: ['pf', 'pf applicable', 'epf', 'pf yes/no', 'pf y/n'] },
    { k: 'esi', h: 'ESI member', w: 'Y or N (empty = Y when gross is within the ESI limit); C = still covered this contribution period although pay went above the limit', ex: 'Y', a: ['esi', 'esi applicable', 'esic', 'esi y/n'] },
    { k: 'pt', h: 'Professional tax', w: 'Empty = worked out for the state chosen', ex: '200', a: ['prof tax', 'p tax', 'ptax', 'pt'] },
    { k: 'tds', h: 'TDS', w: 'Empty = 0 (or the estimate, if ticked)', ex: '0', a: ['income tax', 'tds amount', 'it', 'tax deducted'] },
    { k: 'adv', h: 'Advance', w: 'Advance recovered this month', ex: '1000', a: ['advance recovery', 'salary advance', 'advance deduction'] },
    { k: 'loan', h: 'Loan EMI', w: '', ex: '0', a: ['loan', 'loan recovery', 'emi', 'loan deduction'] },
    { k: 'oded', h: 'Other deductions', w: '', ex: '0', a: ['other deduction', 'deductions', 'misc deduction', 'fine', 'canteen'] },
    { k: 'regime', h: 'Tax regime', w: 'New or Old (only for the TDS estimate)', ex: 'New', a: ['regime', 'income tax regime'] }] };
  /* optional attendance sheet with one line per employee (a day-by-day sheet with columns 1..31 is read by NSPayroll.parseSheet) */
  var ATT = { k: 'att', name: 'Attendance', words: /attend|present|muster/, cols: [
    { k: 'code', h: 'Employee ID', w: '', ex: 'E001', a: EMP.cols[0].a },
    { k: 'name', h: 'Name', w: '', ex: 'Ramesh Kumar', a: EMP.cols[1].a },
    { k: 'paid', h: 'Paid days', need: 1, w: '', ex: '28', a: EMP.cols[12].a },
    { k: 'days', h: 'Days in month', w: '', ex: '30', a: EMP.cols[11].a }] };
  var SETS = [EMP, ATT];

  function T(s) { return String(s == null ? '' : s).trim(); }
  function key(s) { return T(s).toLowerCase().replace(/\s+/g, ' '); }
  function num(s) { var x = T(s).replace(/,/g, '').replace(/^(rs\.?|₹)\s*/i, ''); var n = parseFloat(x); return isFinite(n) ? n : 0; }
  function has(s) { return T(s) !== ''; }
  function rup(v) { return Math.round(v + 1e-9); }
  function yes(s) { return /^(y|yes|1|true|haan|ha)$/i.test(T(s)); }
  function no(s) { return /^(n|no|0|false|nahi|na)$/i.test(T(s)); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function dim(ym) { return new Date(+ym.slice(0, 4), +ym.slice(5, 7), 0).getDate(); }

  /* the blank sample sheet with example rows (made-up people), as [sheet name, rows, widths] for NS.xlsx.write */
  var SAMPLE = [
    ['E001', 'Ramesh Kumar', 'Accountant', 'Accounts', '01-04-2023', '100123456781', '3112345671', 'ABCPK1234F', 'State Bank of India', '30123456781', 'SBIN0001234', '', '30', '', '15000', '2000', '6000', '1600', '1500', '0', '0', 'Y', 'Y', '', '', '1000', '0', '0', 'New'],
    ['E002', 'Sunita Devi', 'Store keeper', 'Stores', '15-06-2024', '100123456782', '3112345672', '', 'Punjab National Bank', '40123456782', 'PUNB0123400', '', '27', '', '11000', '1000', '4000', '800', '0', '0', '900', 'Y', 'Y', '', '', '0', '0', '0', 'New'],
    ['E003', 'Arjun Mehta', 'Sales manager', 'Sales', '01-01-2022', '100123456783', '', 'AHMPM4321K', 'HDFC Bank', '50100123456783', 'HDFC0000123', '', '30', '', '32000', '0', '12800', '1600', '8000', '2000', '5000', 'Y', 'N', '', '2500', '0', '3000', '0', 'New'],
    ['E004', 'Pooja Sharma', 'Office assistant', 'Admin', '10-08-2025', '', '3112345674', '', 'State Bank of India', '30123456784', 'SBIN0001234', '', '29', '', '9500', '0', '3000', '800', '0', '0', '0', 'N', 'Y', '', '', '500', '0', '0', 'New'],
    ['E005', 'Mohd. Salim', 'Driver', 'Transport', '05-02-2021', '', '3112345675', '', 'Bank of Baroda', '60123456785', 'BARB0LUDHIA', '', '26', '', '12000', '0', '3500', '0', '0', '0', '1500', 'Y', 'Y', '', '', '0', '0', '0', 'New'],
    ['E006', 'Kiran Joshi', 'Helper', 'Stores', '01-11-2026', '100123456786', '3112345676', '', 'State Bank of India', '30123456786', 'SBIN0001234', '', '20', '', '8000', '500', '2000', '0', '0', '0', '0', 'Y', 'Y', '', '', '0', '0', '0', 'New']
  ];
  function sampleBook(ym) {
    var head = EMP.cols.map(function (c) { return c.h; });
    var how = [['Column', 'Must fill', 'What to type', 'Example']].concat(EMP.cols.map(function (c) { return [c.h, c.need ? 'Yes' : '', c.w, c.ex]; }));
    how.push(['', '', 'One row per employee. Keep the heading row. Amounts without commas or Rs. Delete the example rows (made-up people) and type your own. Then drop this file on salary-slip.niveshsahayak.in.', '']);
    return [['Employees', [head].concat(SAMPLE), head.map(function (h) { return Math.max(11, Math.min(22, h.length + 3)); })], ['How to fill', how, [22, 10, 70, 18], 'wrap']];
  }
  function sampleRows() { return [{ name: 'Employees', rows: [EMP.cols.map(function (c) { return c.h; })].concat(SAMPLE) }]; }

  /* rows -> employee records. opts: { layouts, force } as in NSTpl.readAs */
  function read(book, opts) {
    var D = (G.NSTpl || require('./tplcore.js')).readAs(book, [EMP], opts);
    /* total lines at the foot of a register are not employees */
    D.totals = []; D.emp = D.emp.filter(function (e) { if (/^(grand\s*|sub\s*-?\s*)?totals?(\s|:|$)/i.test(T(e.name)) || (!T(e.name) && /^(grand\s*)?totals?(\s|:|$)/i.test(T(e.code)))) { D.totals.push(e._r); return false; } return true; });
    return D;
  }
  /* the optional attendance file: day-by-day (columns headed 1..31) through NSPayroll.parseSheet, else one line per employee */
  function readAtt(book, ym, emps, P) {
    P = P || G.NSPayroll;
    var out = { by: {}, matched: 0, unknown: [], bad: [], kind: '' };
    var list = emps.map(function (e, i) { return { id: 'r' + i, code: e.code, name: e.name }; });
    (book || []).some(function (sh) {
      var r = P.parseSheet(sh.rows || [], ym, list);
      if (!r.error) {
        out.kind = 'days';
        Object.keys(r.month).forEach(function (id) { var s = P.summary(r.month[id], ym, P.defaultCfg(), {}); out.by[id] = { paid: s.paid, days: s.days, unmarked: s.unmarked }; });
        out.matched = r.matched; out.unknown = r.unknown; out.bad = r.bad; return true;
      }
      return false;
    });
    if (out.kind) return out;
    var D = (G.NSTpl || require('./tplcore.js')).readAs(book, [ATT], {});
    var byCode = {}, byName = {}; list.forEach(function (e) { if (e.code) byCode[key(e.code)] = e; byName[key(e.name)] = e; });
    D.att.forEach(function (a) {
      var e = (a.code && byCode[key(a.code)]) || byName[key(a.name)];
      if (!e) { out.unknown.push(a.name || a.code); return; }
      out.by[e.id] = { paid: num(a.paid), days: has(a.days) ? num(a.days) : 0 }; out.matched++;
    });
    out.kind = D.att.length ? 'summary' : '';
    return out;
  }

  /* the month's pay for every employee.
     opts: { ym 'YYYY-MM', mode 'full' (amounts are full monthly pay, cut for days not paid) | 'earned' (already this month's),
             state (2-digit GST state code for professional tax), tdsEst (estimate TDS when the column is empty), esiOn (false: the business is not under ESI), pfFull (PF on the whole
             basic + DA above the ceiling), att (readAtt result), pf / esi (NSPayroll settings), ptFor(state, gross, month), tdsFor(gross, regime, pf, pt) }
     Each line: { i, row, e (the sheet row), name, code, days, paid, earn [{desc, amt}], ded [{desc, amt}], gross, tded, net, erPf, erEsi,
       pf (NSPayroll.pfLine or null), esi (NSPayroll.esiLine or null), errs [], warn [], ferr [] (keeps it out of the PF or ESI file only) } */
  function work(rows, opts, P) {
    P = P || G.NSPayroll; opts = opts || {};
    var ym = opts.ym, mdays = dim(ym), mon = +ym.slice(5, 7), seen = {}, att = opts.att || { by: {} };
    var pfc = Object.assign({}, P.PF, opts.pf || {}), esc = Object.assign({}, P.ESI, opts.esi || {});
    return rows.map(function (e, i) {
      var L = { i: i, row: e._r, sheet: e._s, e: e, name: T(e.name), code: T(e.code), errs: [], warn: [], ferr: [] };
      if (!L.name) L.errs.push(['noname']);
      if (L.code) { var ck = key(L.code); if (seen[ck] != null) L.errs.push(['dupcode', { r: seen[ck] }]); else seen[ck] = e._r; }
      var days = has(e.days) ? num(e.days) : mdays, paid;
      var a = att.by['r' + i];
      if (a) { paid = a.paid; if (a.days) days = a.days; if (a.unmarked) L.warn.push(['unmarked', { n: a.unmarked }]); }
      else if (has(e.paid)) paid = num(e.paid);
      else if (has(e.lop)) paid = days - num(e.lop);
      else paid = days;
      if (days > mdays) L.errs.push(['days', { n: days, m: mdays }]);
      if (!(days > 0)) L.errs.push(['nodays']);
      if (paid > days) L.errs.push(['paid', { p: paid, d: days }]);
      if (paid < 0) L.errs.push(['negpaid']);
      L.days = days; L.paid = paid;
      var f = opts.mode === 'earned' || !(days > 0) ? 1 : Math.max(0, Math.min(1, paid / days));
      var cut = function (k) { return rup(num(e[k]) * f); };
      var basic = num(e.basic);
      if (!(basic > 0)) L.errs.push(['nobasic']);
      L.earn = [['Basic salary', cut('basic')], ['Dearness allowance (DA)', cut('da')], ['House rent allowance (HRA)', cut('hra')], ['Conveyance allowance', cut('conv')],
        ['Special allowance', cut('special')], ['Other allowances', cut('other')], ['Overtime / bonus', rup(num(e.extra))]]
        .filter(function (x, j) { return j === 0 || x[1]; }).map(function (x) { return { desc: x[0], amt: x[1] }; });
      L.gross = L.earn.reduce(function (s, x) { return s + x.amt; }, 0);
      var bda = cut('basic') + cut('da');
      /* PF: a member when the column says Y, or (no column / empty) when a UAN is given */
      var uan = T(e.uan).replace(/\s/g, ''), isPf = yes(e.pf) || (!no(e.pf) && !!uan);
      L.pf = null; var pfAmt = 0;
      if (isPf) {
        L.pf = P.pfLine({ uan: uan, name: L.name, gross: L.gross, basic: bda, days: days, paid: paid, full: !!opts.pfFull }, pfc);
        pfAmt = L.pf.epf;
        if (L.pf.errs.indexOf('uan') >= 0) L.ferr.push([uan ? 'baduan' : 'nouan']);
        if (bda > pfc.ceil && !opts.pfFull) L.warn.push(['pfceil', { c: pfc.ceil }]);
      }
      /* ESI: unless the column says N; applies when gross is within the limit */
      var ip = T(e.esino).replace(/\s/g, ''), esiAmt = 0; L.esi = null;
      if (opts.esiOn !== false && !no(e.esi)) {
        L.esi = P.esiLine({ ip: ip, name: L.name, gross: L.gross, paid: paid, days: days, covered: /^c/i.test(T(e.esi)) }, esc);
        if (L.esi.applies) { esiAmt = L.esi.ee; if (L.esi.errs.indexOf('ip') >= 0) L.ferr.push([ip ? 'badip' : 'noip']); }
        else if (L.gross > esc.ceil && ip) L.warn.push(['esiceil', { c: esc.ceil }]);
        if (L.esi.exempt) L.warn.push(['esilow', { d: esc.daily }]);
        if (L.esi.over) L.warn.push(['esicov', { c: esc.ceil }]);
        if (!L.esi.applies && !(ip && L.gross === 0)) L.esi = null;   /* in the ESI file: members within the limit, and members with no wages (days 0) */
      }
      /* professional tax: the column, else the state's table */
      var pt = 0;
      if (has(e.pt)) pt = rup(num(e.pt));
      else if (opts.state && opts.ptFor) { var v = opts.ptFor(opts.state, L.gross, mon); if (v == null) L.warn.push(['nopt']); else pt = v; }
      var tds = has(e.tds) ? rup(num(e.tds)) : opts.tdsEst && opts.tdsFor ? opts.tdsFor(L.gross, /old/i.test(T(e.regime)) ? 'old' : 'new', pfAmt, pt) : 0;
      L.ded = [['Provident fund (PF)', pfAmt], ['ESI', esiAmt], ['Professional tax', pt], ['TDS (income tax)', tds], ['Salary advance', rup(num(e.adv))], ['Loan EMI', rup(num(e.loan))], ['Other deductions', rup(num(e.oded))]]
        .filter(function (x) { return x[1]; }).map(function (x) { return { desc: x[0], amt: x[1] }; });
      L.tded = L.ded.reduce(function (s, x) { return s + x.amt; }, 0);
      L.net = L.gross - L.tded;
      if (L.net < 0) L.errs.push(['neg', { n: L.net }]);
      L.erPf = L.pf ? L.pf.diff + L.pf.eps : 0; L.erEsi = L.esi && L.esi.applies ? L.esi.er : 0;
      var pan = T(e.pan).toUpperCase(); if (pan && !/^[A-Z]{5}\d{4}[A-Z]$/.test(pan)) L.warn.push(['pan']);
      L.ok = !L.errs.length;
      return L;
    });
  }

  /* a salary slip document for slip.js's printed sheet */
  function slipDoc(L, seller, ym, no) {
    var e = L.e;
    return { type: 'salary', no: no, date: ym + '-' + pad(dim(ym)), month: ym, seller: seller || {}, buyer: { name: L.name, addr: '', phone: '', gstin: '', pan: T(e.pan).toUpperCase() },
      emp: { code: L.code, desig: T(e.desig), dept: T(e.dept), doj: (G.NSTpl || require('./tplcore.js')).date(e.doj), pan: T(e.pan).toUpperCase(), uan: T(e.uan), esi: T(e.esino), bankName: T(e.bank), bankAcc: T(e.acc), bankIfsc: T(e.ifsc).toUpperCase() },
      days: String(L.days), paid: String(L.paid), earn: L.earn.map(function (x) { return { desc: x.desc, amt: String(x.amt) }; }), ded: L.ded.map(function (x) { return { desc: x.desc, amt: String(x.amt) }; }), acct: false };
  }
  /* the salary register: one row per employee (numbers stay numbers), a total row */
  var REG = ['Employee ID', 'Name', 'Designation', 'UAN', 'ESI No', 'PAN', 'Days in month', 'Paid days', 'Basic', 'DA', 'HRA', 'Conveyance', 'Special allowance', 'Other allowances', 'Overtime / bonus', 'Gross pay',
    'PF (employee)', 'ESI (employee)', 'Professional tax', 'TDS', 'Advance', 'Loan EMI', 'Other deductions', 'Total deductions', 'Net pay', 'Employer PF + EPS', 'Employer ESI', 'Bank', 'Account No', 'IFSC'];
  function amtOf(list, d) { var x = list.filter(function (l) { return l.desc === d; })[0]; return x ? x.amt : 0; }
  function register(lines) {
    var rows = [REG], tot = REG.map(function () { return 0; });
    lines.forEach(function (L) {
      var e = L.e, r = [L.code, L.name, T(e.desig), T(e.uan), T(e.esino), T(e.pan).toUpperCase(), L.days, L.paid,
        amtOf(L.earn, 'Basic salary'), amtOf(L.earn, 'Dearness allowance (DA)'), amtOf(L.earn, 'House rent allowance (HRA)'), amtOf(L.earn, 'Conveyance allowance'), amtOf(L.earn, 'Special allowance'), amtOf(L.earn, 'Other allowances'), amtOf(L.earn, 'Overtime / bonus'), L.gross,
        amtOf(L.ded, 'Provident fund (PF)'), amtOf(L.ded, 'ESI'), amtOf(L.ded, 'Professional tax'), amtOf(L.ded, 'TDS (income tax)'), amtOf(L.ded, 'Salary advance'), amtOf(L.ded, 'Loan EMI'), amtOf(L.ded, 'Other deductions'), L.tded, L.net, L.erPf, L.erEsi,
        T(e.bank), T(e.acc), T(e.ifsc).toUpperCase()];
      r.forEach(function (v, j) { if (j >= 8 && j <= 26 && typeof v === 'number') tot[j] += v; });
      rows.push(r);
    });
    var t = tot.map(function (v, j) { return j === 1 ? 'Total (' + lines.length + ')' : j >= 8 && j <= 26 ? v : ''; });
    rows.push(t);
    return rows;
  }
  /* bank upload sheet (for net pay transfers): name, account, IFSC, amount */
  function bankRows(lines) { return [['Name', 'Account No', 'IFSC', 'Amount', 'Narration']].concat(lines.filter(function (L) { return L.net > 0; }).map(function (L) { return [L.name, T(L.e.acc), T(L.e.ifsc).toUpperCase(), L.net, 'Salary']; })); }
  function pfLines(lines) { return lines.filter(function (L) { return L.ok && L.pf && L.pf.ok; }).map(function (L) { return L.pf; }); }
  function esiLines(lines) { return lines.filter(function (L) { return L.ok && L.esi && L.esi.ok; }).map(function (L) { return L.esi; }); }

  var API = { EMP: EMP, ATT: ATT, SETS: SETS, sampleBook: sampleBook, sampleRows: sampleRows, read: read, readAtt: readAtt, work: work, slipDoc: slipDoc, register: register, bankRows: bankRows, pfLines: pfLines, esiLines: esiLines, dim: dim, num: num };
  if (typeof module !== 'undefined' && module.exports) module.exports = API; else G.NSSalBulk = API;
})(typeof window !== "undefined" ? window : globalThis);
