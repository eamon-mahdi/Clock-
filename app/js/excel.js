/* Import templates, import parsing (with name matching) and the roster export,
 * all built on RotaXlsx. The template layouts here are the agreed import formats. */
(function (root) {
  'use strict';
  const R = root.RotaRules || (typeof require !== 'undefined' ? require('./rules.js') : null);
  const X = root.RotaXlsx || (typeof require !== 'undefined' ? require('./xlsx.js') : null);
  const S = root.RotaStore || (typeof require !== 'undefined' ? require('./store.js') : null);

  // ---------- colours (match the department's Excel) ----------
  const FILL = { GUH: 'F28B82', RGH: '6FB1EA', NHH: '93C47D', LEAVE: 'FFE36E', ONCALL: 'D9D2E9', LAB: 'F4CCCC', HEAD: 'D9E1EA', WEEKEND: '262626', GONE: '595959', OFF: 'FFFFFF', PENDING: 'FFF2B3', SICK: 'F4C7C3' };
  const head = (v) => ({ v, style: { bold: true, fill: FILL.HEAD, border: true, center: true, wrap: true } });
  const note = (v) => ({ v, style: { color: '595959', wrap: true } });

  function fillFor(cell) {
    switch (cell.kind) {
      case 'site': return FILL[R.siteOf(cell.code)];
      case 'half': return FILL[R.siteOf(cell.baseCode)] || FILL.LEAVE;
      case 'lab': return FILL.LAB;
      case 'leave': return FILL.LEAVE;
      case 'sick': return FILL.SICK;
      case 'oncall': return FILL.ONCALL;
      case 'gone': return FILL.GONE;
      case 'off': return FILL.OFF;
      default: return null;
    }
  }

  // ---------- templates ----------
  const ONCALL_HELP = [
    'How to fill in the on-call import',
    '',
    'Use either sheet. "On-call (list)" has one row per shift. "On-call (grid)" has dates down the side and doctors across the top, like the roster.',
    'Dates: dd/mm/yyyy. Doctor: the name as it appears in the trust rota. The rota coordinator matches any new spelling once and it is remembered.',
    'Shift: one of the codes on the "Shift codes" sheet. Leave a cell blank when a doctor has no on-call that day.',
    'Importing replaces earlier imported on-calls for the dates covered by the file, so you can re-import a corrected week safely.',
    'On-call shifts override the weekly plan for that day. Leave authorised in the rota app is not changed by an import; any clash is highlighted.'
  ];
  const CODE_MEANINGS = {
    'LONG DAY': 'Long day on call', NIGHTS: 'Night shift', TWILIGHT: 'Twilight shift', 'EARLY TWILIGHT': 'Early twilight shift',
    'WARD COVER': 'Weekend / ward cover', 'W/E SD': 'Weekend short day', 'FLEXIBLE W/C': 'Flexible ward cover',
    'OFF PRE ON-CALL': 'Rest day before on-call', 'POST NIGHTS': 'Rest day after nights',
    GUH: 'Grange (inpatient)', 'GUH LAB': 'Lab (does not count as ward or day-unit cover)', RGH: 'RGH (counts as RGH day unit)',
    'RGH - DAY UNIT': 'RGH day unit', 'RGH MDCU': 'RGH MDCU (counts as RGH day unit)', NHH: 'NHH (counts as NHH day unit)',
    'NHH - DAY UNIT': 'NHH day unit', OFF: 'Rostered day off', TEACHING: 'Teaching', 'SpR TRAINING DAY': 'SpR training day'
  };

  function oncallTemplate(todayIso) {
    const mon = R.addDays(R.mondayOf(todayIso), 7);
    const d = (n) => ({ v: R.addDays(mon, n), date: true });
    const list = [
      [head('Date'), head('Doctor'), head('Shift'), head('Notes (optional)')],
      [d(0), 'E. Ward', 'NIGHTS', ''], [d(1), 'E. Ward', 'NIGHTS', ''], [d(2), 'E. Ward', 'NIGHTS', ''], [d(3), 'E. Ward', 'NIGHTS', ''], [d(4), 'E. Ward', 'POST NIGHTS', ''],
      [d(1), 'S. Malik', 'TWILIGHT', ''], [d(4), 'T. Hughes', 'OFF PRE ON-CALL', ''], [d(5), 'T. Hughes', 'LONG DAY', ''], [d(6), 'T. Hughes', 'LONG DAY', ''],
      [d(5), 'L. Osei', 'WARD COVER', 'Example row: delete the examples before importing']
    ];
    const names = ['E. Ward', 'P. Nair', 'T. Hughes', 'S. Malik', 'B. Carter', 'L. Osei'];
    const grid = [[head('Date'), head('Day')].concat(names.map(head))];
    for (let i = 0; i < 7; i++) {
      const iso = R.addDays(mon, i);
      const row = [{ v: iso, date: true }, R.DOW_LONG[i]].concat(names.map(() => ''));
      if (i < 4) row[2] = 'NIGHTS';
      if (i === 4) row[2] = 'POST NIGHTS';
      if (i === 1) row[5] = 'TWILIGHT';
      if (i >= 5) row[4] = 'LONG DAY';
      grid.push(row);
    }
    const codes = [[head('Shift code'), head('Meaning')]].concat(R.ONCALL_CODES.map((c) => [c, CODE_MEANINGS[c] || '']));
    return X.write([
      { name: 'On-call (list)', rows: list, widths: [14, 22, 20, 44], freeze: { row: 1 } },
      { name: 'On-call (grid)', rows: grid, widths: [14, 12].concat(names.map(() => 16)), freeze: { row: 1, col: 2 } },
      { name: 'Shift codes', rows: codes, widths: [20, 40] },
      { name: 'How to fill this in', rows: ONCALL_HELP.map((t, i) => [i === 0 ? { v: t, style: { bold: true, size: 12 } } : note(t)]), widths: [120] }
    ]);
  }

  function planTemplate(state) {
    const rows = [[head('Doctor'), head('Week (A or B)'), head('Mon'), head('Tue'), head('Wed'), head('Thu'), head('Fri'), head('Applies from (optional)')]];
    for (const d of state.doctors) {
      if (d.end && d.end < R.toIso(new Date())) continue;
      rows.push([d.name, 'A'].concat((d.plan && d.plan.A) || ['', '', '', '', '']).concat(['']));
      if (d.plan && d.plan.B) rows.push([d.name, 'B'].concat(d.plan.B).concat(['']));
    }
    const codes = [[head('Site code'), head('Meaning')]].concat(R.SITE_CODES.map((c) => [c, CODE_MEANINGS[c] || '']));
    const help = [
      'How to fill in the weekly plan',
      '',
      'One row per doctor with their usual place on each weekday. Use the codes on the "Site codes" sheet.',
      'Alternating weeks: add a second row for the same doctor with Week = B. Week A starts on the "Applies from" date (or the doctor\'s start date).',
      'On-calls imported from the trust rota and authorised leave sit on top of this plan, so do not enter them here.'
    ];
    return X.write([
      { name: 'Weekly plan', rows, widths: [22, 14, 18, 18, 18, 18, 18, 22], freeze: { row: 1, col: 1 } },
      { name: 'Site codes', rows: codes, widths: [20, 50] },
      { name: 'How to fill this in', rows: help.map((t, i) => [i === 0 ? { v: t, style: { bold: true, size: 12 } } : note(t)]), widths: [120] }
    ]);
  }

  function staffTemplate(state) {
    const rows = [[head('Name'), head('Grade'), head('Group (Resident or Middle grade)'), head('Email'), head('In post from'), head('In post to'), head('Annual leave days'), head('Study leave days'), head('Other spellings in the on-call sheet (comma separated)')]];
    for (const d of state.doctors) {
      rows.push([d.name, d.grade, d.group === 'mg' ? 'Middle grade' : 'Resident', d.email || '', d.start ? { v: d.start, date: true } : '', d.end ? { v: d.end, date: true } : '', d.alEntitlement, d.slEntitlement, (d.aliases || []).join(', ')]);
    }
    return X.write([{ name: 'Staff', rows, widths: [20, 18, 18, 30, 14, 14, 12, 12, 40], freeze: { row: 1, col: 1 } }]);
  }

  // ---------- import parsing ----------
  function cellText(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'object' && v.date) return v.date;
    return String(v).trim();
  }
  function parseDate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'object' && v.date) return v.date;
    if (typeof v === 'number' && v > 20000 && v < 80000) return X.serialToIso(v);
    const s = String(v).trim();
    let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    m = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/.exec(s);
    if (m) {
      const y = m[3].length === 2 ? '20' + m[3] : m[3];
      const iso = `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
      return isNaN(R.toDate(iso)) ? null : iso;
    }
    return null;
  }
  const tokens = (s) => String(s).toLowerCase().replace(/\(.*?\)/g, ' ').replace(/\bdr\b\.?/g, ' ').split(/[^a-z']+/).filter(Boolean);

  // Match a name from a file to a doctor: exact, known alias, or surname + initial.
  function matchDoctor(state, raw) {
    const name = String(raw || '').trim();
    if (!name) return null;
    const low = name.toLowerCase();
    let hit = state.doctors.find((d) => d.name.toLowerCase() === low || (d.aliases || []).some((a) => a.toLowerCase() === low));
    if (hit) return hit;
    const t = tokens(name);
    const cands = state.doctors.filter((d) => {
      const dt = tokens(d.name);
      return dt.length && dt.every((x) => (x.length === 1 ? t.some((y) => y[0] === x) : t.includes(x)));
    });
    return cands.length === 1 ? cands[0] : null;
  }

  function findHeaderRow(rows) {
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
      if ((rows[i] || []).some((v) => /^date$/i.test(cellText(v)) || /^doctor$|^name$/i.test(cellText(v)))) return i;
    }
    return 0;
  }

  // Returns {entries, unknownNames:{raw:count}, badCodes:[...], badDates:n, from, to, format}
  function parseOncall(state, sheets) {
    const sheet = sheets.find((s) => /on-?call/i.test(s.name) && s.rows.some((r) => r && r.some((v) => v !== null && v !== undefined && v !== ''))) || sheets[0];
    const rows = sheet.rows;
    const h = findHeaderRow(rows);
    const header = (rows[h] || []).map((v) => cellText(v).toLowerCase());
    const iDate = header.findIndex((x) => x.startsWith('date'));
    const iDoc = header.findIndex((x) => x === 'doctor' || x === 'name');
    const iShift = header.findIndex((x) => x.startsWith('shift') || x === 'code');
    const raw = [];
    let format;
    if (iDoc >= 0 && iShift >= 0) {
      format = 'list';
      for (let r = h + 1; r < rows.length; r++) {
        const row = rows[r] || [];
        if (!row.some((v) => cellText(v))) continue;
        raw.push({ row: r + 1, date: row[iDate], name: cellText(row[iDoc]), code: cellText(row[iShift]) });
      }
    } else {
      format = 'grid';
      const skip = /^(date|day|consultant|notes?|column\d*|leave restriction)$/;
      const cols = header.map((x, i) => ({ i, name: cellText((rows[h] || [])[i]) })).filter((c) => c.name && !skip.test(c.name.toLowerCase()) && c.i !== iDate);
      for (let r = h + 1; r < rows.length; r++) {
        const row = rows[r] || [];
        for (const c of cols) {
          const code = cellText(row[c.i]);
          if (code) raw.push({ row: r + 1, date: row[iDate < 0 ? 0 : iDate], name: c.name.replace(/\s*\n\s*/g, ' '), code });
        }
      }
    }
    const entries = [], unknownNames = {}, badCodes = [], skipped = [];
    let badDates = 0;
    for (const x of raw) {
      const date = parseDate(x.date);
      if (!date) { badDates++; continue; }
      const code = R.canonicalCode(x.code);
      const kind = R.kindOf(code);
      if (kind === 'other' || kind === 'none') { badCodes.push({ row: x.row, code: x.code }); continue; }
      if (kind !== 'oncall' && kind !== 'off' && kind !== 'leave') { skipped.push({ row: x.row, code: x.code }); continue; }
      const doc = matchDoctor(state, x.name);
      if (!doc) { unknownNames[x.name] = (unknownNames[x.name] || 0) + 1; entries.push({ name: x.name, date, code }); continue; }
      entries.push({ doctorId: doc.id, name: x.name, date, code });
    }
    const dates = entries.map((e) => e.date).sort();
    return { format, sheet: sheet.name, entries, unknownNames, badCodes, skipped, badDates, from: dates[0] || null, to: dates[dates.length - 1] || null };
  }

  function parsePlans(state, sheets) {
    const sheet = sheets.find((s) => /plan/i.test(s.name)) || sheets[0];
    const rows = sheet.rows;
    const h = findHeaderRow(rows);
    const header = (rows[h] || []).map((v) => cellText(v).toLowerCase());
    const idx = (re) => header.findIndex((x) => re.test(x));
    const iDoc = idx(/^(doctor|name)/), iWeek = idx(/^week/), iFrom = idx(/^applies|^from/);
    const iDays = ['mon', 'tue', 'wed', 'thu', 'fri'].map((d) => idx(new RegExp('^' + d)));
    const byDoc = {}, unknownNames = {}, badCodes = [];
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const name = cellText(row[iDoc]);
      if (!name) continue;
      const codes = iDays.map((i) => (i >= 0 ? R.canonicalCode(cellText(row[i])) : ''));
      codes.forEach((c) => { if (c && !['site', 'lab', 'off', 'leave'].includes(R.kindOf(c))) badCodes.push({ row: r + 1, code: c }); });
      const doc = matchDoctor(state, name);
      const key = doc ? doc.id : 'name:' + name;
      if (!doc) unknownNames[name] = (unknownNames[name] || 0) + 1;
      byDoc[key] = byDoc[key] || { doctorId: doc ? doc.id : null, name, A: null, B: null, from: null };
      const week = /b/i.test(cellText(row[iWeek])) ? 'B' : 'A';
      byDoc[key][week] = codes;
      const from = parseDate(row[iFrom]);
      if (from) byDoc[key].from = from;
    }
    const plans = Object.values(byDoc).map((p) => Object.assign(p, { A: p.A || ['', '', '', '', ''] }));
    return { plans, unknownNames, badCodes };
  }

  function parseStaff(state, sheets) {
    const rows = sheets[0].rows;
    const h = findHeaderRow(rows);
    const header = (rows[h] || []).map((v) => cellText(v).toLowerCase());
    const idx = (re) => header.findIndex((x) => re.test(x));
    const c = { name: idx(/^name/), grade: idx(/^grade/), group: idx(/^group/), email: idx(/^email/), from: idx(/from/), to: idx(/post to|^to/), al: idx(/annual/), sl: idx(/study/), alias: idx(/spelling|alias/) };
    const staff = [], problems = [];
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const name = cellText(row[c.name]);
      if (!name) continue;
      const groupTxt = cellText(row[c.group]).toLowerCase();
      const grade = cellText(row[c.grade]);
      const group = /middle|registrar|spr|sas|mg/.test(groupTxt + ' ' + grade.toLowerCase()) && !/resident/.test(groupTxt) ? 'mg' : 'res';
      const num = (v) => { const n = Number(cellText(v)); return isFinite(n) ? n : 0; };
      staff.push({
        name, grade, group, email: cellText(row[c.email]), start: parseDate(row[c.from]) || '', end: parseDate(row[c.to]) || '',
        alEntitlement: num(row[c.al]), slEntitlement: num(row[c.sl]),
        aliases: cellText(row[c.alias]).split(',').map((s) => s.trim()).filter(Boolean)
      });
      if (!grade) problems.push(`Row ${r + 1}: no grade for ${name}`);
    }
    return { staff, problems };
  }

  // ---------- roster export ----------
  function rosterWorkbook(state, from, to) {
    const docs = state.doctors.filter((d) => R.dateRange(from, to).some((iso) => R.inPost(d, iso) || (d.start && d.start <= to)))
      .sort((a, b) => (a.group === b.group ? 0 : a.group === 'res' ? -1 : 1));
    const header = [head('Date'), head('Day'), head('Consultant'), head('Leave restriction')]
      .concat(docs.map((d) => head(`${d.grade}\n${d.name}`)))
      .concat([head('Residents (GUH · RGH · NHH)'), head('Middle grade (GUH ward · RGH DU · NHH DU)')]);
    const rows = [header];
    for (const iso of R.dateRange(from, to)) {
      const wk = R.isWeekend(iso);
      const base = { border: true, center: true, wrap: true };
      const dark = wk ? { fill: FILL.WEEKEND, color: 'FFFFFF' } : {};
      const restr = S.restrictionOn(state, iso);
      const row = [
        { v: iso, date: true, style: Object.assign({}, base, dark, restr && !wk ? { fill: restr.level === 'NO AL' ? 'C00000' : '7030A0', color: 'FFFFFF', bold: true } : {}) },
        { v: R.DOW_LONG[R.dow(iso)], style: Object.assign({}, base, dark) },
        { v: wk ? '' : (state.settings.consultantOfWeek[R.mondayOf(iso)] || ''), style: Object.assign({}, base, dark) },
        { v: restr ? restr.level : '', style: Object.assign({}, base, dark) }
      ];
      for (const d of docs) {
        const cell = R.resolveCell(state, d, iso);
        let v = cell.code;
        if (cell.pending && cell.kind !== 'leave') v = (v ? v + '\n' : '') + `(${R.LEAVE_TYPES[cell.pending.type].label} requested)`;
        const fill = fillFor(cell) || (wk ? FILL.WEEKEND : null);
        row.push({ v, style: Object.assign({}, base, fill ? { fill } : {}, cell.kind === 'gone' || (wk && !fill) ? { color: 'FFFFFF' } : {}) });
      }
      const cover = R.coverWithRisk(state, iso);
      const txt = (group) => cover ? cover.rules.filter((r) => r.rule.group === group).map((r) => `${r.rule.short} ${({ ok: 'OK', gap: 'GAP', exception: 'resident only', risk: 'AT RISK', na: 'n/a' })[r.status]}`).join(' · ') : '';
      const bad = (group) => cover && cover.rules.some((r) => r.rule.group === group && (r.status === 'gap' || r.status === 'risk'));
      row.push({ v: txt('res'), style: Object.assign({}, base, dark, bad('res') ? { color: 'C00000', bold: true } : {}) });
      row.push({ v: txt('mg'), style: Object.assign({}, base, dark, bad('mg') ? { color: 'C00000', bold: true } : {}) });
      rows.push(row);
    }
    const log = [[head('Submitted'), head('Doctor'), head('Type'), head('From'), head('To'), head('Part'), head('Working days'), head('Status'), head('Decided by'), head('Decided at'), head('Cover'), head('Reason / note')]];
    for (const r of state.requests.slice().sort((a, b) => (a.submittedAt < b.submittedAt ? 1 : -1))) {
      const d = R.doctorById(state, r.doctorId);
      log.push([r.submittedAt.slice(0, 16).replace('T', ' '), d ? d.name : r.doctorId, R.LEAVE_TYPES[r.type].label, { v: r.start, date: true }, { v: r.end, date: true }, r.part, r.days, r.status, r.decidedBy || '', (r.decidedAt || '').slice(0, 16).replace('T', ' '), (r.cover || []).map((c) => c.text).join('; '), r.reason || r.note || '']);
    }
    const audit = [[head('When'), head('Who'), head('What'), head('Details')]].concat(state.audit.map((a) => [a.at.slice(0, 16).replace('T', ' '), a.by, a.action, a.detail]));
    return X.write([
      { name: 'Roster', rows, widths: [12, 11, 14, 11].concat(docs.map(() => 15)).concat([30, 40]), freeze: { row: 1, col: 2 }, heights: { 0: 42 } },
      { name: 'Leave log', rows: log, widths: [17, 16, 22, 12, 12, 7, 8, 11, 14, 17, 40, 40], freeze: { row: 1 } },
      { name: 'Change log', rows: audit, widths: [17, 18, 26, 90], freeze: { row: 1 } },
      { name: 'About', rows: [['Exported from the Haematology Rota prototype. This copy is read-only: make changes in the app and export again.'], [`Dates ${R.fmtUk(from)} to ${R.fmtUk(to)}.`]], widths: [110] }
    ]);
  }

  const api = { oncallTemplate, planTemplate, staffTemplate, parseOncall, parsePlans, parseStaff, matchDoctor, parseDate, rosterWorkbook, FILL };
  root.RotaExcel = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
