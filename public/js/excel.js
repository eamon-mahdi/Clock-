/* Import templates, import parsing (with name matching) and the roster export,
 * all built on RotaXlsx. The template layouts here are the agreed import formats. */
(function (root) {
  'use strict';
  const R = root.RotaRules || (typeof require !== 'undefined' ? require('./rules.js') : null);
  const X = root.RotaXlsx || (typeof require !== 'undefined' ? require('./xlsx.js') : null);
  const S = root.RotaStore || (typeof require !== 'undefined' ? require('./store.js') : null);

  // ---------- colours (match the department's Excel) ----------
  const FILL = { GUH: 'F28B82', RGH: '6FB1EA', NHH: '93C47D', LEAVE: 'FFE36E', ONCALL: 'FFE36E', LAB: 'F4CCCC', HEAD: 'D9E1EA', WEEKEND: '262626', GONE: '595959', OFF: 'FFFFFF', PENDING: 'FFF2B3', SICK: 'FFE36E' };
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
      case 'other': return FILL.ONCALL;
      case 'edt': return FILL[R.siteOf(cell.baseCode)] || FILL.LAB;
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
    const rows = [[head('Name'), head('Grade'), head('Group (Resident or Senior doctor)'), head('Email'), head('In post from'), head('In post to'), head('Annual leave days'), head('Study leave days'), head('Other spellings in the on-call sheet (comma separated)')]];
    for (const d of state.doctors) {
      rows.push([d.name, d.grade, d.group === 'mg' ? 'Senior doctor' : 'Resident', d.email || '', d.start ? { v: d.start, date: true } : '', d.end ? { v: d.end, date: true } : '', d.alEntitlement, d.slEntitlement, (d.aliases || []).join(', ')]);
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
      const group = /senior|middle|registrar|spr|sas|mg|specialty|specialist|\bst[3-9]/.test(groupTxt + ' ' + grade.toLowerCase()) && !/resident/.test(groupTxt) ? 'mg' : 'res';
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


  // ---------- master workbook (doctors + consultant of the week + leave restrictions) ----------
  const MASTER_COLS = ['Name', 'Grade', 'Category (Resident or Senior)', 'Starts', 'Finishes', 'Finish date confirmed (Yes/No)', 'Takes over the column of (name)', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'NHS email', 'Annual leave days', 'Carried over / extra days', 'Study leave days', 'Other spellings (comma separated)'];
  const MASTER_HELP = [
    'How to fill in the rota master workbook',
    '',
    'This one file sets up the rota. Fill in the tabs you need and import it under Set up → Import → Master workbook. You can import it again at any time: doctors are matched by name and updated, new ones are added. Nobody is deleted by an import.',
    '',
    'Doctors tab: one row per doctor (or per post, for placeholders).',
    '  • Grade: e.g. F1, F2, IMT1, IMT2, IMT3, JCF, Registrar ST3–ST8, Specialty doctor, Specialist grade, Senior clinical fellow, or anything else.',
    '  • Category: Resident (F1, F2, IMT, JCF) or Senior (registrars, specialty doctors, specialist grade, senior clinical fellows). Left blank, it is worked out from the grade.',
    '  • Starts / Finishes: dd/mm/yyyy. Left blank for a Foundation, IMT or registrar grade, the next rotation dates are filled in (see the Rotation dates tab).',
    '  • Finish date confirmed: No if the finish date may change (e.g. a registrar who may stay 6 or 12 months).',
    '  • Takes over the column of: the name of the doctor they replace, so the rota continues in the same column at the changeover.',
    '  • Name not known yet: write "New F1" (or New + grade). It shows as TBC until you type the real name over it.',
    '  • Mon–Fri: usual location each weekday, using the codes on the Codes tab (GUH, RGH - DAY UNIT, NHH - DAY UNIT, GUH LAB, OFF …).',
    '  • Annual leave days: left blank, the Wales default for the grade and time in post is used. Carried over / extra: accrued or carried-over days.',
    '',
    'Consultant of the week tab: the Monday of each week and the consultant. Add as many weeks as you like.',
    'Leave restrictions tab: dates when annual leave is not normally approved (NO AL) or should be avoided (AVOID).',
    'Rotation dates tab: reference only. The changeover dates worked out from the programme rules for the next few years.',
    '',
    'A CSV version of the Doctors tab is also available (one sheet only: CSV files cannot have tabs).'
  ];
  const yesNo = (b) => (b ? 'Yes' : 'No');
  function doctorRows(state) {
    const rows = [MASTER_COLS.map(head)];
    const cols = R.columns(state, '1900-01-01', '2999-12-31');
    const order = [];
    for (const g of ['res', 'mg']) for (const c of cols) if (c.docs[0].group === g) order.push(...c.docs);
    for (const d of order) {
      const prev = R.predecessorOf(state, d);
      const plan = (d.plan && d.plan.A) || ['', '', '', '', ''];
      rows.push([d.name, d.grade, d.group === 'mg' ? 'Senior' : 'Resident', d.start ? { v: d.start, date: true } : '', d.end ? { v: d.end, date: true } : '',
        d.end ? yesNo(!d.endProvisional) : '', prev ? prev.name : ''].concat(plan.map((x) => x || '')).concat([d.email || '', d.alEntitlement === undefined ? '' : d.alEntitlement, d.alExtra || '', d.slEntitlement === undefined ? '' : d.slEntitlement, (d.aliases || []).join(', ')]));
    }
    return rows;
  }
  function consultantRows(state, todayIso) {
    const map = (state.settings && state.settings.consultantOfWeek) || {};
    const rows = [[head('Week commencing (Monday)'), head('Consultant')]];
    const keys = Object.keys(map).sort();
    let w = R.mondayOf((state.settings && state.settings.rotation && state.settings.rotation.start) || todayIso);
    if (keys.length && keys[0] < w) w = keys[0];
    let last = R.addDays(R.mondayOf(todayIso), 7 * 26);
    if (keys.length && keys[keys.length - 1] > last) last = keys[keys.length - 1];
    for (; w <= last; w = R.addDays(w, 7)) rows.push([{ v: w, date: true }, map[w] || '']);
    return rows;
  }
  function restrictionRows(state) {
    return [[head('From'), head('To'), head('Level (NO AL or AVOID)'), head('Note')]].concat(((state.settings && state.settings.restricted) || []).slice().sort((a, b) => (a.start < b.start ? -1 : 1))
      .map((r) => [{ v: r.start, date: true }, { v: r.end || r.start, date: true }, r.level, r.note || '']));
  }
  function rotationRows(todayIso, years) {
    const rows = [[head('Programme'), head('Rotation starts'), head('Day before the next start'), head('Rule')]];
    const from = R.addMonths(todayIso, -6);
    for (const [k, p] of Object.entries(R.PROGRAMMES)) {
      const starts = R.programmeStarts(k, from, (years || 5) * p.months.length + 1);
      for (let i = 0; i < starts.length - 1; i++) rows.push([p.label, { v: starts[i], date: true }, { v: R.addDays(starts[i + 1], -1), date: true }, i === 0 ? p.note : '']);
    }
    return rows;
  }
  function codesRows() {
    const rows = [[head('Location code (Mon–Fri columns)'), head('Meaning')]].concat(R.SITE_CODES.map((c) => [c, CODE_MEANINGS[c] || '']));
    rows.push(['', ''], [head('Grades'), head('Category')]);
    ['F1', 'F2', 'IMT1', 'IMT2', 'IMT3', 'JCF'].forEach((g) => rows.push([g, 'Resident']));
    ['Registrar ST3', 'Registrar ST4', 'Registrar ST5', 'Registrar ST6', 'Registrar ST7', 'Registrar ST8', 'Specialty doctor', 'Specialist grade', 'Senior clinical fellow'].forEach((g) => rows.push([g, 'Senior']));
    return rows;
  }
  const helpSheet = (lines) => ({ name: 'How to use', rows: lines.map((t, i) => [i === 0 ? { v: t, style: { bold: true, size: 12 } } : note(t)]), widths: [130] });
  const MASTER_WIDTHS = [20, 20, 14, 13, 13, 13, 22, 15, 15, 15, 15, 15, 30, 10, 10, 10, 30];
  function masterTemplate(state, todayIso) {
    return X.write([
      helpSheet(MASTER_HELP),
      { name: 'Doctors', rows: doctorRows(state), widths: MASTER_WIDTHS, freeze: { row: 1, col: 1 }, heights: { 0: 42 } },
      { name: 'Consultant of the week', rows: consultantRows(state, todayIso), widths: [24, 30], freeze: { row: 1 } },
      { name: 'Leave restrictions', rows: restrictionRows(state), widths: [13, 13, 20, 50], freeze: { row: 1 } },
      { name: 'Rotation dates', rows: rotationRows(todayIso, 5), widths: [28, 18, 22, 70], freeze: { row: 1 } },
      { name: 'Codes', rows: codesRows(), widths: [30, 50] }
    ]);
  }

  // CSV (one sheet only). Dates are written dd/mm/yyyy.
  const csvCell = (v) => {
    let t = v === null || v === undefined ? '' : typeof v === 'object' ? (v.date ? R.fmtUk(v.v) : String(v.v === undefined ? '' : v.v)) : String(v);
    if (/^[=+\-@]/.test(t)) t = "'" + t; // keep spreadsheet formulas from running
    return /[",\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t;
  };
  const toCsv = (rows) => '﻿' + rows.map((r) => r.map(csvCell).join(',')).join('\r\n') + '\r\n';
  function parseCsv(text) {
    const t = String(text).replace(/^﻿/, '');
    const rows = []; let row = [], cur = '', q = false;
    for (let i = 0; i < t.length; i++) {
      const ch = t[i];
      if (q) { if (ch === '"') { if (t[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; continue; }
      if (ch === '"') q = true;
      else if (ch === ',' || ch === ';' && !t.includes(',')) { row.push(cur); cur = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && t[i + 1] === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; }
      else cur += ch;
    }
    if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
    return rows.map((r) => r.map((c) => c.replace(/^'(?=[=+\-@])/, '')));
  }
  const doctorsCsvTemplate = (state) => toCsv(doctorRows(state));

  // Reads a master workbook (or a CSV of one of its tabs). Returns what would change; nothing is applied here.
  function parseMaster(state, sheets) {
    const out = { doctors: [], consultants: {}, restrictions: [], problems: [], found: [] };
    const headerOf = (sh) => { const h = findHeaderRowAny(sh.rows); return { h, cols: (sh.rows[h] || []).map((v) => cellText(v).toLowerCase()) }; };
    for (const sh of sheets) {
      const { h, cols } = headerOf(sh);
      const idx = (re) => cols.findIndex((x) => re.test(x));
      if (idx(/^name$|^name\b/) >= 0 && idx(/^grade/) >= 0) {
        out.found.push(sh.name || 'Doctors');
        const c = { name: idx(/^name/), grade: idx(/^grade/), cat: idx(/^categ|^group/), start: idx(/^start/), end: idx(/^finish(es)?$|^finishes|^end/), conf: idx(/confirm/), rep: idx(/takes over|replac/),
          email: idx(/email/), al: idx(/^annual/), extra: idx(/carried|extra/), sl: idx(/^study/), alias: idx(/spelling|alias/), days: ['mon', 'tue', 'wed', 'thu', 'fri'].map((d) => idx(new RegExp('^' + d))) };
        for (let r = h + 1; r < sh.rows.length; r++) {
          const row = sh.rows[r] || [];
          const name = cellText(row[c.name]);
          if (!name) continue;
          const grade = cellText(row[c.grade]);
          if (!grade) { out.problems.push(`Doctors row ${r + 1}: no grade for ${name}; skipped.`); continue; }
          const catTxt = cellText(row[c.cat]).toLowerCase();
          const prog = R.programmeOf(grade);
          const group = /senior|middle|\bmg\b/.test(catTxt) ? 'mg' : /resident|junior/.test(catTxt) ? 'res'
            : (prog === 'registrar' || /specialty|specialist|senior|registrar|\bsas\b|\bst[3-9]/i.test(grade)) ? 'mg' : 'res';
          let start = parseDate(row[c.start]) || '', end = parseDate(row[c.end]) || '';
          if ((cellText(row[c.start]) && !start) || (cellText(row[c.end]) && !end)) out.problems.push(`Doctors row ${r + 1}: could not read a date for ${name}. Use dd/mm/yyyy.`);
          let suggested = false;
          if (!start && prog !== 'other') { const sug = R.suggestDates(grade, R.toIso(new Date())); start = sug.start; end = end || sug.end; suggested = true; }
          const plan = c.days.map((i) => (i >= 0 ? R.canonicalCode(cellText(row[i])) : ''));
          plan.forEach((code, k) => { if (code && !['site', 'lab', 'off', 'leave'].includes(R.kindOf(code))) out.problems.push(`Doctors row ${r + 1}: "${cellText(row[c.days[k]])}" is not a location code (${name}, ${R.DOW_SHORT[k]}).`); });
          const num = (v) => { const t = cellText(v); if (t === '') return undefined; const n = Number(t); return isFinite(n) ? n : undefined; };
          const conf = cellText(row[c.conf]).toLowerCase();
          out.doctors.push({ row: r + 1, name, grade, group, start, end, endProvisional: conf ? /^n/.test(conf) : prog !== 'foundation' && !!end && suggested,
            replacesName: cellText(row[c.rep]), plan, email: cellText(row[c.email]), alEntitlement: num(row[c.al]), alExtra: num(row[c.extra]), slEntitlement: num(row[c.sl]),
            aliases: cellText(row[c.alias]).split(',').map((x) => x.trim()).filter(Boolean), placeholder: /^new\b|\btbc\b/i.test(name), suggested });
        }
      } else if (idx(/week/) >= 0 && idx(/consultant/) >= 0) {
        out.found.push(sh.name || 'Consultant of the week');
        const iw = idx(/week/), ic = idx(/consultant/);
        for (let r = h + 1; r < sh.rows.length; r++) {
          const row = sh.rows[r] || [];
          const d = parseDate(row[iw]);
          const who = cellText(row[ic]);
          if (!d) { if (cellText(row[iw])) out.problems.push(`Consultant row ${r + 1}: could not read the date.`); continue; }
          if (who) out.consultants[R.mondayOf(d)] = who;
        }
      } else if (idx(/^level/) >= 0 && idx(/^from/) >= 0) {
        out.found.push(sh.name || 'Leave restrictions');
        const ia = idx(/^from/), ib = idx(/^to/), il = idx(/^level/), inote = idx(/^note/);
        for (let r = h + 1; r < sh.rows.length; r++) {
          const row = sh.rows[r] || [];
          const a = parseDate(row[ia]);
          if (!a) continue;
          const b = parseDate(row[ib]) || a;
          const level = /avoid/i.test(cellText(row[il])) ? 'AVOID' : 'NO AL';
          out.restrictions.push({ start: a, end: b < a ? a : b, level, note: cellText(row[inote]) });
        }
      }
    }
    if (!out.found.length) out.problems.push('No Doctors, Consultant of the week or Leave restrictions tab was found. Use the master template.');
    return out;
  }
  function findHeaderRowAny(rows) {
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
      if ((rows[i] || []).some((v) => /^(name|grade|week|from|date|doctor)/i.test(cellText(v)))) return i;
    }
    return 0;
  }

  // Full backup: the master tabs (re-importable) plus everything else, read-only.
  function backupWorkbook(state, todayIso) {
    const docName = (id) => (R.doctorById(state, id) || { name: id }).name;
    const oncall = [[head('Date'), head('Doctor'), head('Shift')]];
    Object.entries(state.oncall || {}).forEach(([id, m]) => Object.entries(m).forEach(([d, code]) => oncall.push([{ v: d, date: true }, docName(id), code])));
    oncall.splice(1, oncall.length, ...oncall.slice(1).sort((a, b) => (a[0].v < b[0].v ? -1 : a[0].v > b[0].v ? 1 : 0)));
    const changes = [[head('Date'), head('Doctor'), head('Changed to'), head('Reason'), head('By')]];
    Object.entries(state.overrides || {}).forEach(([id, m]) => Object.entries(m).forEach(([d, o]) => changes.push([{ v: d, date: true }, docName(id), o.code, o.reason || '', o.by || ''])));
    const leave = [[head('Doctor'), head('Type'), head('From'), head('To'), head('Part'), head('Working days'), head('Status'), head('Days not approved'), head('Decided by'), head('Sent')]].concat(state.requests.map((r) => [docName(r.doctorId), r.type === 'SICK' ? 'Unavailable' : (R.LEAVE_TYPES[r.type] || R.LEAVE_TYPES.OTHER).label, { v: r.start, date: true }, { v: r.end, date: true }, r.part, r.days, r.status, (r.excludedDates || []).map(R.fmtUk).join(', '), r.decidedBy || '', (r.submittedAt || '').slice(0, 10)]));
    const edt = [[head('Date'), head('Session'), head('Doctor'), head('Type'), head('Booked by')]].concat((state.edt || []).map((e) => [{ v: e.date, date: true }, e.session, docName(e.doctorId), e.type, e.bookedBy || '']));
    const audit = [[head('When'), head('Who'), head('What'), head('Details')]].concat((state.audit || []).map((a) => [a.at.slice(0, 16).replace('T', ' '), a.by, a.action, a.detail]));
    return X.write([
      helpSheet([`Haematology rota backup, ${R.fmtLong(todayIso)}`, '', 'The Doctors, Consultant of the week and Leave restrictions tabs can be imported again (Set up → Import → Master workbook). The On-call (list) tab can be imported as an on-call file.', 'For an exact restore of everything, use Set up → Backups on the website.']),
      { name: 'Doctors', rows: doctorRows(state), widths: MASTER_WIDTHS, freeze: { row: 1, col: 1 }, heights: { 0: 42 } },
      { name: 'Consultant of the week', rows: consultantRows(state, todayIso), widths: [24, 30], freeze: { row: 1 } },
      { name: 'Leave restrictions', rows: restrictionRows(state), widths: [13, 13, 20, 50], freeze: { row: 1 } },
      { name: 'Leave', rows: leave, widths: [18, 22, 12, 12, 7, 8, 11, 30, 14, 12], freeze: { row: 1 } },
      { name: 'On-call (list)', rows: oncall, widths: [12, 20, 20], freeze: { row: 1 } },
      { name: 'Rota changes', rows: changes, widths: [12, 20, 20, 40, 14], freeze: { row: 1 } },
      { name: 'EDT SDT', rows: edt, widths: [12, 8, 20, 8, 14], freeze: { row: 1 } },
      { name: 'Change log', rows: audit, widths: [17, 18, 26, 90], freeze: { row: 1 } }
    ]);
  }

  // ---------- roster export ----------
  function rosterWorkbook(state, from, to) {
    const docs = state.doctors.filter((d) => R.dateRange(from, to).some((iso) => R.inPost(d, iso) || (d.start && d.start <= to)))
      .sort((a, b) => (a.group === b.group ? 0 : a.group === 'res' ? -1 : 1));
    const header = [head('Date'), head('Day'), head('Consultant'), head('Leave restriction')]
      .concat(docs.map((d) => head(`${d.grade}\n${d.name}`)))
      .concat([head('Residents (GUH · RGH · NHH)'), head('Senior doctors (GUH ward · RGH DU · NHH DU)')]);
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
        if (cell.pending && cell.kind !== 'leave') v = (v ? v + '\n' : '') + `(${(R.LEAVE_TYPES[cell.pending.type] || R.LEAVE_TYPES.OTHER).label} requested)`;
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
      log.push([(r.submittedAt || '').slice(0, 16).replace('T', ' '), d ? d.name : r.doctorId, r.type === 'SICK' || r.type === 'UNAVAIL' ? 'Unavailable' : (R.LEAVE_TYPES[r.type] || R.LEAVE_TYPES.OTHER).label, { v: r.start, date: true }, { v: r.end, date: true }, r.part, r.days, r.status, r.decidedBy || '', (r.decidedAt || '').slice(0, 16).replace('T', ' '), (r.cover || []).map((c) => c.text).join('; '), r.type === 'SICK' ? '' : (r.message || r.note || '')]);
    }
    const audit = [[head('When'), head('Who'), head('What'), head('Details')]].concat(state.audit.map((a) => [a.at.slice(0, 16).replace('T', ' '), a.by, a.action, a.detail]));
    return X.write([
      { name: 'Roster', rows, widths: [12, 11, 14, 11].concat(docs.map(() => 15)).concat([30, 40]), freeze: { row: 1, col: 2 }, heights: { 0: 42 } },
      { name: 'Leave log', rows: log, widths: [17, 16, 22, 12, 12, 7, 8, 11, 14, 17, 40, 40], freeze: { row: 1 } },
      { name: 'Change log', rows: audit, widths: [17, 18, 26, 90], freeze: { row: 1 } },
      { name: 'About', rows: [['Exported from the Haematology Rota website. This copy is read-only: make changes in the app and export again.'], [`Dates ${R.fmtUk(from)} to ${R.fmtUk(to)}.`]], widths: [110] }
    ]);
  }

  const api = { oncallTemplate, planTemplate, staffTemplate, parseOncall, parsePlans, parseStaff, matchDoctor, parseDate, rosterWorkbook, FILL,
    masterTemplate, doctorsCsvTemplate, parseMaster, parseCsv, toCsv, backupWorkbook };
  root.RotaExcel = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
