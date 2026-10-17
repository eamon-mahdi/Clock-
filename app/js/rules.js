/* Haematology rota rules: dates, codes, how a roster cell is worked out, cover checks (per morning
 * and afternoon session), leave day counts, balances, cover suggestions, the staffing check used for
 * leave requests, EDT/SDT rules and Wales default entitlements.
 * Pure functions over the app state, so they run in the browser, in the Cloudflare server and in Node tests. */
(function (root) {
  'use strict';

  // ---------- dates (ISO yyyy-mm-dd, UTC) ----------
  const toDate = (iso) => { const [y, m, d] = iso.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
  const toIso = (dt) => dt.toISOString().slice(0, 10);
  const addDays = (iso, n) => { const d = toDate(iso); d.setUTCDate(d.getUTCDate() + n); return toIso(d); };
  const dow = (iso) => (toDate(iso).getUTCDay() + 6) % 7; // 0 = Monday … 6 = Sunday
  const isWeekend = (iso) => dow(iso) >= 5;
  const mondayOf = (iso) => addDays(iso, -dow(iso));
  const daysBetween = (a, b) => Math.round((toDate(b) - toDate(a)) / 86400000);
  function dateRange(a, b) { const out = []; for (let d = a; d <= b; d = addDays(d, 1)) out.push(d); return out; }
  const isIso = (s) => typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(toDate(s));
  const DOW_SHORT = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const DOW_LONG = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const fmtShort = (iso) => { const d = toDate(iso); return `${DOW_SHORT[dow(iso)]} ${d.getUTCDate()} ${MONTHS[d.getUTCMonth()]}`; };
  const fmtLong = (iso) => { const d = toDate(iso); return `${DOW_LONG[dow(iso)]} ${d.getUTCDate()} ${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}`; };
  const fmtUk = (iso) => { const [y, m, d] = iso.split('-'); return `${d}/${m}/${y}`; };
  const halfRound = (n) => Math.round(n * 2) / 2;

  // England & Wales bank holidays (editable in Settings)
  const BANK_HOLIDAYS = [
    '2026-01-01', '2026-04-03', '2026-04-06', '2026-05-04', '2026-05-25', '2026-08-31', '2026-12-25', '2026-12-28',
    '2027-01-01', '2027-03-26', '2027-03-29', '2027-05-03', '2027-05-31', '2027-08-30', '2027-12-27', '2027-12-28'
  ];

  // ---------- codes ----------
  const SITE_CODES = ['GUH', 'GUH LAB', 'RGH', 'RGH - DAY UNIT', 'RGH MDCU', 'NHH', 'NHH - DAY UNIT', 'OFF', 'TEACHING', 'INDUCTION', 'SpR TRAINING DAY'];
  const ONCALL_CODES = ['LONG DAY', 'NIGHTS', 'TWILIGHT', 'EARLY TWILIGHT', 'WARD COVER', 'W/E SD', 'FLEXIBLE W/C', 'OFF PRE ON-CALL', 'POST NIGHTS'];
  // The leave form offers AL, SL, SICK and OTHER (as the old Microsoft Form did); the rest are kept for imports.
  const LEAVE_TYPES = {
    AL: { label: 'Annual leave', code: 'ANNUAL LEAVE', balance: 'al' },
    SL: { label: 'Study leave', code: 'STUDY LEAVE', balance: 'sl' },
    SICK: { label: 'Sick leave', code: 'UNAVAILABLE' },
    OTHER: { label: 'Other leave', code: 'OTHER LEAVE' },
    UNAVAIL: { label: 'Unavailable', code: 'UNAVAILABLE' },
    PL: { label: 'Professional / exam leave', code: 'PROFESSIONAL LEAVE' },
    TEACH: { label: 'Teaching', code: 'TEACHING' },
    COMP: { label: 'Compassionate leave', code: 'COMPASSIONATE LEAVE' },
    TOIL: { label: 'Day off in lieu', code: 'DAY OFF IN LIEU' }
  };
  const FORM_LEAVE_TYPES = ['AL', 'SL', 'SICK', 'OTHER'];
  const LEAVE_CODES = Object.values(LEAVE_TYPES).map((t) => t.code);
  const RGH_CODES = ['RGH', 'RGH - DAY UNIT', 'RGH MDCU'];
  const NHH_CODES = ['NHH', 'NHH - DAY UNIT'];

  const norm = (code) => String(code || '').trim().replace(/\s+/g, ' ').toUpperCase();
  function siteOf(code) {
    const c = norm(code);
    if (c === 'GUH') return 'GUH';
    if (RGH_CODES.includes(c)) return 'RGH';
    if (NHH_CODES.includes(c)) return 'NHH';
    return null;
  }
  function kindOf(code) {
    const c = norm(code);
    if (!c) return 'none';
    if (c === 'NOT IN POST') return 'gone';
    if (c === 'BANK HOLIDAY') return 'bh';
    if (c === 'UNAVAILABLE' || c === 'SICK') return 'sick';
    if (siteOf(c)) return 'site';
    if (c.includes('LAB')) return 'lab';
    if (c === 'OFF') return 'off';
    if (ONCALL_CODES.includes(c)) return 'oncall';
    if (LEAVE_CODES.includes(c) || c === 'SPR TRAINING DAY' || c === 'INDUCTION') return 'leave';
    return 'other';
  }
  function canonicalCode(raw) {
    let c = norm(raw).replace(/^(RGH|NHH|GUH)\s*-\s*/, '$1 - ').replace(/\s+-\s*|\s*-\s+/g, ' - ');
    const aliases = {
      'RGH DAY UNIT': 'RGH - DAY UNIT', 'RGH DU': 'RGH - DAY UNIT', 'NHH DAY UNIT': 'NHH - DAY UNIT', 'NHH DU': 'NHH - DAY UNIT',
      'RGH - MDCU': 'RGH MDCU', 'GUH - LAB': 'GUH LAB', 'NIGHT': 'NIGHTS', 'LD': 'LONG DAY', 'N': 'NIGHTS',
      'WE SD': 'W/E SD', 'OFF PRE-ON-CALL': 'OFF PRE ON-CALL', 'OFF PRE ON CALL': 'OFF PRE ON-CALL', 'OFF - PRE ON-CALL': 'OFF PRE ON-CALL',
      'POST NIGHT': 'POST NIGHTS', 'POST-NIGHTS': 'POST NIGHTS', 'SPR TRAINING': 'SpR TRAINING DAY', 'SPR TRAINING DAY': 'SpR TRAINING DAY',
      'AL': 'ANNUAL LEAVE', 'SL': 'STUDY LEAVE', 'SICK LEAVE': 'UNAVAILABLE', 'SICK': 'UNAVAILABLE'
    };
    if (aliases[c]) c = aliases[c];
    return c === 'SPR TRAINING DAY' ? 'SpR TRAINING DAY' : c;
  }

  // ---------- settings ----------
  const DEFAULT_RULES = {
    minResGUH: 1, minResRGH: 1, minResNHH: 1, minSenGUH: 1, minSenRGH: 0, minSenNHH: 0,
    edtOthersAtGUH: 2, noticeWeeks: 6, reminderDays: 3
  };
  const ruleNum = (state, key) => {
    const r = (state.settings && state.settings.rules) || {};
    return typeof r[key] === 'number' ? r[key] : DEFAULT_RULES[key];
  };

  // ---------- cover rules ----------
  const ALL_WEEKDAYS = [0, 1, 2, 3, 4];
  const TUE_FRI = [1, 2, 3, 4];
  const COVER_RULES = [
    { id: 'res-GUH', group: 'res', short: 'GUH', label: 'Resident at GUH', codes: ['GUH'], days: ALL_WEEKDAYS, min: 'minResGUH' },
    { id: 'res-RGH', group: 'res', short: 'RGH', label: 'Resident at RGH', codes: RGH_CODES, days: ALL_WEEKDAYS, min: 'minResRGH' },
    { id: 'res-NHH', group: 'res', short: 'NHH', label: 'Resident at NHH', codes: NHH_CODES, days: TUE_FRI, min: 'minResNHH' },
    { id: 'mg-GUH', group: 'mg', short: 'GUH', label: 'Senior doctor on GUH ward', codes: ['GUH'], days: ALL_WEEKDAYS, min: 'minSenGUH' },
    { id: 'mg-RGH', group: 'mg', short: 'RGH DU', label: 'Senior doctor, RGH day unit', codes: RGH_CODES, days: ALL_WEEKDAYS, fallbackGroup: 'res', min: 'minSenRGH' },
    { id: 'mg-NHH', group: 'mg', short: 'NHH DU', label: 'Senior doctor, NHH day unit', codes: NHH_CODES, days: TUE_FRI, fallbackGroup: 'res', min: 'minSenNHH' }
  ];
  const STATUS_RANK = { na: 0, ok: 0, risk: 1, exception: 2, gap: 3 };
  const worse = (a, b) => (STATUS_RANK[b] > STATUS_RANK[a] ? b : a);

  // ---------- state helpers ----------
  // England and Wales bank holidays for any year, worked out from the rules (with weekend substitute days).
  const bhCache = {};
  function bankHolidaysFor(year) {
    if (bhCache[year]) return bhCache[year];
    const pad = (n) => String(n).padStart(2, '0');
    const iso = (m, d) => `${year}-${pad(m)}-${pad(d)}`;
    // Easter Sunday (anonymous Gregorian algorithm)
    const a = year % 19, b = Math.floor(year / 100), c = year % 100, d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3);
    const h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451);
    const em = Math.floor((h + l - 7 * m + 114) / 31), ed = ((h + l - 7 * m + 114) % 31) + 1;
    const easter = iso(em, ed);
    const firstMon = (mo) => { let x = iso(mo, 1); while (dow(x) !== 0) x = addDays(x, 1); return x; };
    const lastMon = (mo) => { let x = addDays(iso(mo + 1, 1), -1); while (dow(x) !== 0) x = addDays(x, -1); return x; };
    const out = [addDays(easter, -2), addDays(easter, 1), firstMon(5), lastMon(5), lastMon(8)];
    // New Year, Christmas and Boxing Day move to the next free weekday when they fall at a weekend.
    const fixed = [iso(1, 1), iso(12, 25), iso(12, 26)];
    fixed.filter((x) => !isWeekend(x)).forEach((x) => out.push(x));
    for (const x of fixed.filter(isWeekend)) { let y = addDays(x, 1); while (isWeekend(y) || out.includes(y)) y = addDays(y, 1); out.push(y); }
    bhCache[year] = out.sort();
    return bhCache[year];
  }
  // Settings can add one-off extras (e.g. a coronation) or remove a date that the government has moved.
  function isBankHoliday(state, iso) {
    const st = (state && state.settings) || {};
    if ((st.bankHolidayRemoved || []).includes(iso)) return false;
    if ((st.bankHolidayExtra || []).includes(iso)) return true;
    return bankHolidaysFor(Number(iso.slice(0, 4))).includes(iso);
  }
  const bankHolidaysBetween = (state, from, to) => dateRange(from, to).filter((x) => isBankHoliday(state, x));
  const isWorkday = (state, iso) => !isWeekend(iso) && !isBankHoliday(state, iso);
  const inPost = (doc, iso) => (!doc.start || iso >= doc.start) && (!doc.end || iso <= doc.end);
  const doctorById = (state, id) => state.doctors.find((d) => d.id === id);
  const covers = (req, iso) => iso >= req.start && iso <= req.end && !(req.excludedDates || []).includes(iso);
  const groupLabel = (g) => (g === 'mg' ? 'Senior doctor' : 'Resident');

  function planCode(doc, iso) {
    if (!doc.plan) return '';
    const i = dow(iso);
    if (i > 4) return '';
    let week = doc.plan.A || [];
    if (doc.plan.B && doc.plan.B.some(Boolean)) {
      const anchor = mondayOf(doc.plan.anchor || doc.start || iso);
      const n = Math.floor(daysBetween(anchor, mondayOf(iso)) / 7);
      if (Math.abs(n) % 2 === 1) week = doc.plan.B;
    }
    return week[i] || '';
  }
  const leaveLabelFor = (req) => (req.type === 'OTHER' && req.otherText ? 'OTHER LEAVE' : (LEAVE_TYPES[req.type] || LEAVE_TYPES.OTHER).code);
  const edtOn = (state, doctorId, iso) => (state.edt || []).find((e) => e.doctorId === doctorId && e.date === iso) || null;
  const coverCodeOf = (code) => (siteOf(code) ? norm(code) : '');

  // Works out what one doctor is doing on one date, including who counts as cover each session.
  function resolveCell(state, doc, iso) {
    if (!inPost(doc, iso)) return { code: 'NOT IN POST', kind: 'gone', source: 'gone', am: '', pm: '' };
    const oncall = (state.oncall[doc.id] || {})[iso];
    const reqs = state.requests.filter((r) => r.doctorId === doc.id && covers(r, iso) &&
      !(r.type !== 'SICK' && r.type !== 'UNAVAIL' && !isWorkday(state, iso) && !oncall));
    const authorised = reqs.find((r) => r.status === 'authorised');
    const pending = reqs.find((r) => r.status === 'pending' || r.status === 'hold') || null;
    const override = (state.overrides[doc.id] || {})[iso];

    let base;
    if (override) base = { code: override.code, kind: kindOf(override.code), source: 'override', note: override.reason || '' };
    else if (oncall) base = { code: oncall, kind: kindOf(oncall) === 'other' ? 'oncall' : kindOf(oncall), source: 'oncall', locked: true };
    else if (isWeekend(iso)) base = { code: '', kind: 'none', source: 'weekend' };
    else if (isBankHoliday(state, iso)) base = { code: 'BANK HOLIDAY', kind: 'bh', source: 'bh' };
    else { const p = planCode(doc, iso); base = { code: p, kind: p ? kindOf(p) : 'none', source: p ? 'plan' : 'none' }; }
    const cc = coverCodeOf(base.code);

    if (authorised) {
      const label = leaveLabelFor(authorised);
      if (authorised.part === 'AM' || authorised.part === 'PM') {
        const text = authorised.part === 'PM' ? `${base.code || 'OFF'} AM; PM ${label}` : `AM ${label}; ${base.code || 'OFF'} PM`;
        return Object.assign({}, base, { code: text, kind: 'half', baseCode: base.code, leave: authorised, pending, halfOff: authorised.part, am: authorised.part === 'AM' ? '' : cc, pm: authorised.part === 'PM' ? '' : cc });
      }
      return {
        code: label, kind: authorised.type === 'SICK' || authorised.type === 'UNAVAIL' ? 'sick' : 'leave', source: 'leave', leave: authorised,
        conflict: oncall ? `Clashes with imported on-call: ${oncall}` : null, am: '', pm: ''
      };
    }
    const edt = edtOn(state, doc.id, iso);
    if (edt && (base.kind === 'site' || base.kind === 'lab')) {
      const text = edt.session === 'AM' ? `${edt.type} AM · ${base.code} PM` : `${base.code} AM · ${edt.type} PM`;
      return Object.assign({}, base, { code: text, kind: 'edt', baseCode: base.code, edt, pending, am: edt.session === 'AM' ? '' : cc, pm: edt.session === 'PM' ? '' : cc });
    }
    return Object.assign(base, { pending, am: cc, pm: cc });
  }
  const coverCode = (cell) => norm(cell.kind === 'half' || cell.kind === 'edt' ? cell.baseCode : cell.code);

  // Cover for one date. opts: absent [ids] (all day), absentSession {id: 'AM'|'PM'}, moves {id: code} (all day),
  // pendingAbsent (treat full-day pending requests as approved), skipPendingIds [request ids].
  function coverFor(state, iso, opts) {
    opts = opts || {};
    if (!isWorkday(state, iso)) return null;
    const absent = new Set(opts.absent || []);
    const absentSession = opts.absentSession || {};
    const moves = opts.moves || {};
    const sessions = { AM: [], PM: [] };
    for (const doc of state.doctors) {
      if (!inPost(doc, iso) || absent.has(doc.id)) continue;
      const cell = resolveCell(state, doc, iso);
      if (opts.pendingAbsent && cell.pending && cell.pending.part === 'FULL' && cell.kind !== 'leave' && cell.kind !== 'sick') continue;
      for (const s of ['AM', 'PM']) {
        if (absentSession[doc.id] === s) continue;
        const code = moves[doc.id] ? norm(moves[doc.id]) : (s === 'AM' ? cell.am : cell.pm);
        if (code) sessions[s].push({ doc, code });
      }
    }
    const d = dow(iso);
    const rules = COVER_RULES.map((rule) => {
      const min = ruleNum(state, rule.min);
      if (!rule.days.includes(d) || min <= 0) return { rule, status: 'na', names: [], min };
      let status = 'ok';
      const names = new Set();
      for (const s of ['AM', 'PM']) {
        const holders = sessions[s].filter((p) => p.doc.group === rule.group && rule.codes.includes(p.code));
        holders.forEach((h) => names.add(h.doc.name));
        let st = 'ok';
        if (holders.length < min) {
          const fb = rule.fallbackGroup ? sessions[s].filter((p) => p.doc.group === rule.fallbackGroup && rule.codes.includes(p.code)) : [];
          st = fb.length ? 'exception' : 'gap';
          fb.forEach((h) => names.add(h.doc.name));
        }
        status = worse(status, st);
      }
      return { rule, status, names: Array.from(names), min };
    });
    return { iso, rules };
  }
  function coverWithRisk(state, iso) {
    const actual = coverFor(state, iso);
    if (!actual) return null;
    const ifPending = coverFor(state, iso, { pendingAbsent: true });
    actual.rules.forEach((r, i) => { if (r.status === 'ok' && STATUS_RANK[ifPending.rules[i].status] > 0) r.status = 'risk'; });
    return actual;
  }
  const worstStatus = (cover) => (cover ? cover.rules.reduce((w, r) => worse(w, r.status), 'ok') : 'na');

  // ---------- leave ----------
  function leaveDays(state, start, end, part, excluded) {
    if (!isIso(start) || !isIso(end) || end < start) return 0;
    const ex = excluded || [];
    const n = dateRange(start, end).filter((d) => isWorkday(state, d) && !ex.includes(d)).length;
    return (part === 'AM' || part === 'PM') && start === end ? n * 0.5 : n;
  }
  function balance(state, doc) {
    const sum = (type, statuses) => state.requests.filter((r) => r.doctorId === doc.id && r.type === type && statuses.includes(r.status))
      .reduce((a, r) => a + (r.days || 0), 0);
    const mk = (type, ent, extra) => {
      ent = (Number(ent) || 0) + (Number(extra) || 0);
      const used = sum(type, ['authorised']); const pending = sum(type, ['pending', 'hold']);
      return { entitlement: ent || 0, used, pending, left: (ent || 0) - used };
    };
    return { al: mk('AL', doc.alEntitlement, doc.alExtra), sl: mk('SL', doc.slEntitlement, doc.slExtra) };
  }
  const restrictedIn = (state, start, end) => ((state.settings && state.settings.restricted) || []).filter((r) => r.start <= end && (r.end || r.start) >= start)
    .sort((a, b) => (a.start < b.start ? -1 : 1));

  // Suggest who could cover a doctor on each working day (moves that leave no new gap).
  function suggestCover(state, doctorId, start, end, part) {
    const doc = doctorById(state, doctorId);
    if (!doc || !isIso(start) || !isIso(end) || end < start || part === 'AM' || part === 'PM') return [];
    const out = [];
    for (const iso of dateRange(start, end)) {
      if (!isWorkday(state, iso) || !inPost(doc, iso)) continue;
      const before = coverFor(state, iso);
      const without = coverFor(state, iso, { absent: [doctorId] });
      const broken = without.rules.filter((r, i) => STATUS_RANK[r.status] > STATUS_RANK[before.rules[i].status]);
      const cell = resolveCell(state, doc, iso);
      if (!broken.length) { out.push({ date: iso, needed: false, from: cell.code, broken: [], options: [] }); continue; }
      const options = [];
      for (const b of broken) {
        const target = b.rule.codes.includes(coverCode(cell)) ? coverCode(cell) : b.rule.codes[0];
        for (const cand of state.doctors) {
          if (cand.id === doctorId || cand.group !== b.rule.group || !inPost(cand, iso)) continue;
          const av = availability(state, cand, iso, null);
          if (!av.ok) continue;
          const cc = resolveCell(state, cand, iso);
          let type;
          if (cc.kind === 'off' || (cc.kind === 'none' && cc.source !== 'weekend')) type = 'dayoff';
          else if ((cc.kind === 'site' || cc.kind === 'lab') && !b.rule.codes.includes(coverCode(cc))) type = 'move';
          else continue;
          const after = coverFor(state, iso, { absent: [doctorId], moves: { [cand.id]: target } });
          if (after.rules.some((r, i) => r.status === 'gap' && before.rules[i].status !== 'gap')) continue;
          const exceptions = after.rules.filter((r, i) => r.status === 'exception' && before.rules[i].status !== 'exception').map((r) => r.rule.label);
          const text = type === 'dayoff' ? `${cand.name} is off that day and could work ${target} instead` : `${cand.name} could move from ${cc.code} to ${target}`;
          options.push({ doctorId: cand.id, name: cand.name, from: cc.code || 'OFF', to: target, type, exceptions, score: exceptions.length * 10 + (type === 'dayoff' ? 5 : 0), text });
        }
      }
      options.sort((a, b) => a.score - b.score || a.name.localeCompare(b.name));
      out.push({ date: iso, needed: true, from: cell.code, broken: broken.map((r) => r.rule.label), options: options.slice(0, 5) });
    }
    return out;
  }

  // Is a named doctor free to cover on a date? ignoreRequestId: the request being checked.
  function availability(state, doc, iso, ignoreRequestId) {
    if (!doc) return { ok: false, why: 'not on the staff list' };
    if (!inPost(doc, iso)) return { ok: false, why: 'not in post' };
    const cell = resolveCell(state, doc, iso);
    if (cell.kind === 'leave') return { ok: false, why: 'on leave' };
    if (cell.kind === 'sick') return { ok: false, why: 'unavailable' };
    if (cell.kind === 'oncall') return { ok: false, why: `on call (${cell.code.toLowerCase()})` };
    if (cell.kind === 'half') return { ok: false, why: 'on leave for half the day' };
    if (cell.pending && cell.pending.id !== ignoreRequestId && cell.pending.part === 'FULL') return { ok: false, why: 'has asked for leave that day' };
    if (edtOn(state, doc.id, iso)) return { ok: false, why: 'has EDT/SDT booked' };
    const already = state.requests.find((r) => r.id !== ignoreRequestId && ['pending', 'hold', 'authorised'].includes(r.status) && covers(r, iso) &&
      (r.cover || []).some((c) => c.date === iso && c.doctorId === doc.id));
    if (already) { const who = doctorById(state, already.doctorId); return { ok: false, why: `already covering ${who ? who.name : 'someone'}` }; }
    return { ok: true, why: cell.code ? `free (usually ${cell.code})` : 'free' };
  }

  // Which named cover applies on a date: day-unit cover where the requester is at RGH/NHH, GUH cover at GUH.
  function coverRoleFor(state, doc, iso) {
    const cell = resolveCell(state, doc, iso);
    const site = siteOf(coverCode(cell));
    if (site === 'GUH') return { role: 'guh', code: coverCode(cell), cell };
    if (site === 'RGH' || site === 'NHH') return { role: 'dayunit', code: coverCode(cell), cell };
    return { role: null, code: '', cell };
  }

  // The staffing check for a leave request (shown to the requester and flagged to authorisers).
  // input: {doctorId, start, end, part, dayUnitCoverId, guhCoverId, requestId}
  function staffingCheck(state, input) {
    const doc = doctorById(state, input.doctorId);
    const out = { days: [], bad: 0, warn: 0, cover: [], alternatives: [] };
    if (!doc || !isIso(input.start) || !isIso(input.end) || input.end < input.start) return out;
    const half = (input.part === 'AM' || input.part === 'PM') && input.start === input.end ? input.part : null;
    for (const iso of dateRange(input.start, input.end)) {
      if (!isWorkday(state, iso) || !inPost(doc, iso)) continue;
      const role = coverRoleFor(state, doc, iso);
      const day = { date: iso, at: role.cell.code || 'nothing rostered', status: 'ok', text: '', problems: [], options: [], coverName: '' };
      if (!role.role) {
        day.text = `Not needed at a haematology site that day (${role.cell.code ? role.cell.code.toLowerCase() : 'off'}).`;
        out.days.push(day);
        continue;
      }
      const coverId = role.role === 'guh' ? input.guhCoverId : input.dayUnitCoverId;
      const coverDoc = coverId ? doctorById(state, coverId) : null;
      const av = coverDoc ? availability(state, coverDoc, iso, input.requestId) : null;
      const moves = coverDoc && av.ok ? { [coverDoc.id]: role.code } : {};
      if (coverDoc && av.ok) out.cover.push({ date: iso, doctorId: coverDoc.id, to: role.code, text: `${fmtShort(iso)}: ${coverDoc.name} covers ${role.code}` });
      const absentOpts = half ? { absentSession: { [doc.id]: half } } : { absent: [doc.id] };
      const before = coverFor(state, iso, { skip: true });
      const after = coverFor(state, iso, Object.assign({ moves }, absentOpts));
      const competing = state.requests.filter((r) => r.id !== input.requestId && r.doctorId !== doc.id && ['pending', 'hold'].includes(r.status) && covers(r, iso) && r.part === 'FULL');
      const withAll = coverFor(state, iso, Object.assign({ moves }, half ? { absentSession: { [doc.id]: half } } : {}, { absent: (half ? [] : [doc.id]).concat(competing.map((r) => r.doctorId)) }));
      const newGaps = after.rules.filter((r, i) => r.status === 'gap' && before.rules[i].status !== 'gap');
      const oldGaps = after.rules.filter((r, i) => r.status === 'gap' && before.rules[i].status === 'gap' && before.rules[i].names.includes(doc.name)); // already short, and this doctor was one of too few
      const exceptions = after.rules.filter((r, i) => r.status === 'exception' && before.rules[i].status !== 'exception');
      const competeWorse = withAll.rules.filter((r, i) => STATUS_RANK[r.status] > STATUS_RANK[after.rules[i].status]);
      if (coverDoc && !av.ok) day.problems.push(`${coverDoc.name} can't cover: ${av.why}.`);
      if (coverDoc && av.ok) day.coverName = coverDoc.name;
      if (!coverDoc) day.problems.push(role.role === 'guh' ? 'No GUH cover named.' : 'No day-unit cover named.');
      if (newGaps.length || oldGaps.length) {
        day.status = 'bad';
        const labels = newGaps.concat(oldGaps).map((r) => r.rule.label.toLowerCase());
        day.text = `Not enough doctors: ${labels.join(', ')}${oldGaps.length && !newGaps.length ? ' (already short before this request)' : ''}.`;
        const groups = new Set(newGaps.concat(oldGaps).map((r) => r.rule.group));
        const others = state.doctors.filter((d) => d.id !== doc.id && groups.has(d.group) && inPost(d, iso)).map((d) => {
          const c = resolveCell(state, d, iso);
          return `${d.name}: ${c.code ? c.code.toLowerCase() : 'off'}`;
        });
        day.whoElse = others.join(' · ');
        day.options = suggestCover(state, doc.id, iso, iso, 'FULL').flatMap((s) => s.options).slice(0, 3).map((o) => o.text);
      } else {
        day.text = 'Enough doctors.' + (day.coverName ? ` ${day.coverName} covers ${role.code}.` : '');
        if (exceptions.length) { day.status = 'warn'; day.problems.push(`Only a resident would cover: ${exceptions.map((r) => r.rule.label.toLowerCase()).join(', ')}.`); }
        if (competeWorse.length) { day.status = 'warn'; day.problems.push(`If ${competing.map((r) => (doctorById(state, r.doctorId) || {}).name).join(' and ')}'s request${competing.length > 1 ? 's are' : ' is'} also approved: ${competeWorse.map((r) => r.rule.label.toLowerCase()).join(', ')} would be short.`); }
        if (coverDoc && !av.ok) day.status = 'warn';
      }
      if (competing.length) day.competing = competing.map((r) => ({ id: r.id, name: (doctorById(state, r.doctorId) || {}).name, submittedAt: r.submittedAt }));
      if (day.status === 'bad') out.bad++;
      if (day.status === 'warn') out.warn++;
      out.days.push(day);
    }
    if (out.bad) out.alternatives = nearestDates(state, doc, input.end, 3);
    return out;
  }
  function nearestDates(state, doc, from, n) {
    const found = [];
    for (let i = 1; i <= 40 && found.length < n; i++) {
      const iso = addDays(from, i);
      if (!isWorkday(state, iso) || !inPost(doc, iso)) continue;
      if (restrictedIn(state, iso, iso).some((r) => r.level === 'NO AL')) continue;
      const before = coverFor(state, iso);
      const after = coverFor(state, iso, { absent: [doc.id] });
      if (!after.rules.some((r, k) => r.status === 'gap' && before.rules[k].status !== 'gap')) found.push(iso);
    }
    return found;
  }

  // ---------- EDT / SDT ----------
  const DEFAULT_EDT = [
    { match: /^F1\b/i, type: 'SDT', hours: 1 },
    { match: /^F2\b/i, type: 'SDT', hours: 3 },
    { match: /senior clinical fellow|\bSCF\b/i, type: 'EDT', hours: 4 },
    { match: /specialist/i, type: 'SPA', hours: 4 },
    { match: /specialty doctor|specialty dr|\bSAS\b|associate specialist/i, type: 'SPA', hours: 4 },
    { match: /\bIMT|\bCT\d|core/i, type: 'EDT', hours: 2 },
    { match: /\bJCF\b|clinical fellow/i, type: 'EDT', hours: 2 },
    { match: /\bST\d|registrar|\bStR\b|\bSpR\b/i, type: 'EDT', hours: 4 }
  ];
  function edtDefault(doc) {
    if (doc.edtHours !== undefined && doc.edtHours !== null && doc.edtHours !== '') {
      const t = DEFAULT_EDT.find((r) => r.match.test(doc.grade || ''));
      return { type: doc.edtType || (t ? t.type : 'EDT'), hours: Number(doc.edtHours) };
    }
    const t = DEFAULT_EDT.find((r) => r.match.test(doc.grade || ''));
    if (t) return { type: t.type, hours: t.hours };
    return doc.group === 'mg' ? { type: 'EDT', hours: 4 } : { type: 'EDT', hours: 2 };
  }
  function edtEntitlement(state, doc, todayIso) {
    const def = edtDefault(doc);
    const fte = typeof doc.fte === 'number' ? doc.fte : 1;
    const rot = (state.settings && state.settings.rotation) || {};
    const a = doc.start && doc.start > (rot.start || '') ? doc.start : (rot.start || doc.start);
    const b = doc.end || rot.end || addDays(a, 182);
    const weeks = Math.max(0, (daysBetween(a, b) + 1) / 7);
    const done = todayIso < a ? 0 : Math.max(0, (daysBetween(a, todayIso < b ? todayIso : b) + 1) / 7);
    const mine = (state.edt || []).filter((e) => e.doctorId === doc.id && e.date >= a && e.date <= b);
    return {
      type: def.type, hours: def.hours, fte, from: a, to: b, weeks: Math.round(weeks * 10) / 10,
      due: Math.round(def.hours * fte * weeks), dueByToday: Math.round(def.hours * fte * done * 10) / 10,
      booked: mine.length * 4, taken: mine.filter((e) => e.date <= todayIso).length * 4
    };
  }
  // Check whether an EDT/SDT/SPA session can be booked.
  function edtCheck(state, doctorId, iso, session) {
    const doc = doctorById(state, doctorId);
    const checks = [];
    const add = (ok, rule, detail) => checks.push({ ok, rule, detail });
    if (!doc || !isIso(iso) || (session !== 'AM' && session !== 'PM')) return { ok: false, checks: [{ ok: false, rule: 'Choose a doctor, date and session', detail: '' }] };
    const sesWord = session === 'AM' ? 'morning' : 'afternoon';
    if (!isWorkday(state, iso)) { add(false, 'Weekday only', 'EDT/SDT is booked on working weekdays.'); return { ok: false, checks }; }
    if (!inPost(doc, iso)) { add(false, 'In post', `${doc.name} is not in post that day.`); return { ok: false, checks }; }
    if (edtOn(state, doctorId, iso)) { add(false, 'One session a day', 'There is already a session booked that day.'); return { ok: false, checks }; }
    const cell = resolveCell(state, doc, iso);
    if (!(cell.kind === 'site' || cell.kind === 'lab')) { add(false, 'Working at a site that day', `${doc.name} is ${cell.code ? cell.code.toLowerCase() : 'not rostered'} that day.`); return { ok: false, checks }; }
    const before = coverFor(state, iso);
    const after = coverFor(state, iso, { absentSession: { [doc.id]: session } });
    const newGaps = after.rules.filter((r, i) => STATUS_RANK[r.status] > STATUS_RANK[before.rules[i].status]);
    if (doc.group === 'res') {
      const atGuh = coverCode(cell) === 'GUH';
      add(atGuh, 'Residents take EDT/SDT at GUH only', atGuh ? `${doc.name} is at GUH that day.` : `${doc.name} is rostered at ${cell.code}. Move them to GUH for the day first (change the roster cell), then book.`);
      const need = ruleNum(state, 'edtOthersAtGUH');
      const others = state.doctors.filter((d) => d.id !== doc.id && d.group === 'res' && inPost(d, iso)).filter((d) => {
        const c = resolveCell(state, d, iso);
        return (session === 'AM' ? c.am : c.pm) === 'GUH';
      });
      add(others.length >= need, `At least ${need} other residents at GUH that ${sesWord}`, others.length ? `${others.map((d) => `${d.name} (${d.grade})`).join(', ')}${others.length < need ? `: ${need - others.length} more needed` : ''}.` : 'Nobody else is at GUH.');
    } else {
      add(true, 'Senior doctors can take EDT at any site', `${doc.name} is at ${cell.code}.`);
    }
    add(!newGaps.length, 'Site cover is kept', newGaps.length ? `That ${sesWord} would leave: ${newGaps.map((r) => r.rule.label.toLowerCase()).join(', ')}.` : 'All posts stay covered.');
    return { ok: checks.every((c) => c.ok), checks, type: edtDefault(doc).type };
  }

  // ---------- Wales default annual leave (editable) ----------
  function alPerYear(doc) {
    const g = doc.grade || '';
    if (/specialist/i.test(g)) return 34;
    if (/specialty doctor|specialty dr|SAS/i.test(g)) return 29;
    if (/senior clinical fellow|SCF/i.test(g)) return 28;
    if (/\bST[3-9]|registrar|StR|SpR/i.test(g)) return 33;
    return 28;
  }
  function alProRata(doc) {
    const per = alPerYear(doc);
    const fte = typeof doc.fte === 'number' ? doc.fte : 1;
    if (!doc.start || !doc.end) return { perYear: per, days: halfRound(per * fte), basis: `${per} days a year${fte < 1 ? ' × ' + fte : ''}` };
    const n = daysBetween(doc.start, doc.end) + 1;
    return { perYear: per, days: halfRound(per * n / 365 * fte), basis: `${per} × ${n} days in post ÷ 365${fte < 1 ? ' × ' + fte : ''}` };
  }

  // Changeovers in a date range: doctors starting or leaving, plus manual changeover dates.
  function changeoversIn(state, from, to) {
    const map = {};
    const add = (iso, text) => { if (iso >= from && iso <= to) (map[iso] = map[iso] || []).push(text); };
    const replaced = new Set();
    for (const d of state.doctors) {
      if (!d.start) continue;
      const prev = predecessorOf(state, d);
      if (prev) { replaced.add(prev.id); add(d.start, `${prev.name} → ${d.name} (${d.grade})`); } else add(d.start, `${d.name} (${d.grade}) starts`);
    }
    for (const d of state.doctors) if (d.end && !replaced.has(d.id)) add(addDays(d.end, 1), `${d.name} (${d.grade}) leaves${d.endProvisional ? ' (date not confirmed)' : ''}`);
    for (const c of (state.settings && state.settings.changeovers) || []) add(c.date, c.label);
    return map;
  }


  // ---------- rota columns (posts): a new doctor can continue the column of the doctor they replace ----------
  const slotOf = (doc) => doc.slot || doc.id;
  function predecessorOf(state, doc) {
    if (!doc.slot || doc.slot === doc.id) return null;
    const same = state.doctors.filter((d) => d.id !== doc.id && slotOf(d) === doc.slot && (d.start || '') < (doc.start || '9999'));
    same.sort((a, b) => ((a.start || '') < (b.start || '') ? 1 : -1));
    return same[0] || null;
  }
  // Columns for a date range, in the order doctors were added. Each has the doctors who hold it, earliest first.
  function columns(state, from, to, filter) {
    const map = new Map();
    for (const d of state.doctors) {
      if (filter && !filter(d)) continue;
      const k = slotOf(d);
      if (!map.has(k)) map.set(k, []);
      map.get(k).push(d);
    }
    const out = [];
    for (const [slot, docs] of map) {
      docs.sort((a, b) => ((a.start || '') < (b.start || '') ? -1 : 1));
      const visible = docs.filter((d) => (!d.start || d.start <= to) && (!d.end || d.end >= addDays(from, -28)));
      if (visible.length) out.push({ slot, docs: visible, group: visible[visible.length - 1].group });
    }
    return out;
  }
  // Who holds a column on a date (or the nearest holder, shown as "not in post").
  function holderOn(col, iso) {
    return col.docs.find((d) => inPost(d, iso)) || col.docs.filter((d) => !d.start || d.start <= iso).pop() || col.docs[0];
  }

  // ---------- rotation dates (suggestions; every date can be changed) ----------
  function programmeOf(grade) {
    const g = grade || '';
    if (/^F[12]\b|foundation/i.test(g)) return 'foundation';
    if (/\bIMT|\bCT\d|core/i.test(g)) return 'imt';
    if (/\bST\d|registrar|\bStR\b|\bSpR\b/i.test(g)) return 'registrar';
    return 'other';
  }
  const PROGRAMMES = {
    foundation: { label: 'Foundation (F1/F2)', months: [3, 7, 11], wednesday: true, length: 4, note: 'First Wednesday of August, December and April; 4-month posts.' },
    imt: { label: 'IMT', months: [1, 7], wednesday: true, length: 6, note: 'First Wednesday of August and February; usually 6 months.' },
    registrar: { label: 'Registrars (training grade)', months: [1, 7], wednesday: false, length: 6, note: '1 August or 1 February; 6 months, sometimes extended to 12 (not always certain).' }
  };
  function firstWednesday(y, m) { let d = toIso(new Date(Date.UTC(y, m, 1))); while (dow(d) !== 2) d = addDays(d, 1); return d; }
  function programmeStarts(prog, fromIso, count) {
    const p = PROGRAMMES[prog];
    if (!p) return [];
    const out = [];
    let y = Number(fromIso.slice(0, 4)) - 1;
    while (out.length < count && y < 2200) {
      for (const m of p.months) {
        const d = p.wednesday ? firstWednesday(y, m) : toIso(new Date(Date.UTC(y, m, 1)));
        if (d >= fromIso && out.length < count) out.push(d);
      }
      y++;
    }
    return out;
  }
  // Suggested start and end for a new doctor of this grade, starting on or after fromIso.
  function suggestDates(grade, fromIso) {
    const prog = programmeOf(grade);
    if (prog === 'other') return { programme: prog, start: '', end: '', provisional: true, note: 'No fixed rotation: enter the dates.' };
    const starts = programmeStarts(prog, fromIso, 3);
    const p = PROGRAMMES[prog];
    const end = starts[1] ? addDays(starts[1], -1) : '';
    return { programme: prog, start: starts[0], end, provisional: prog !== 'foundation', note: p.note, nextStarts: starts };
  }
  // Same day n months later (clamped to the month's last day).
  function addMonths(iso, n) {
    const d = toDate(iso), day = d.getUTCDate();
    d.setUTCDate(1); d.setUTCMonth(d.getUTCMonth() + n);
    const last = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0)).getUTCDate();
    d.setUTCDate(Math.min(day, last));
    return toIso(d);
  }
  const shortNotice = (state, start, todayIso) => daysBetween(todayIso, start) < 7 * ruleNum(state, 'noticeWeeks');
  // The furthest date the rota needs to show (no fixed end: grows as doctors and changeovers are added).
  function rotaHorizon(state, todayIso) {
    let last = addDays(todayIso, 56);
    for (const d of state.doctors) { if (d.end && d.end > last) last = d.end; if (d.start && d.start > last) last = d.start; }
    for (const c of (state.settings && state.settings.changeovers) || []) if (c.date > last) last = c.date;
    return addDays(last, 14);
  }

  const api = {
    toDate, toIso, addDays, dow, isWeekend, isWorkday, mondayOf, daysBetween, dateRange, isIso, fmtShort, fmtLong, fmtUk, DOW_SHORT, DOW_LONG, halfRound,
    BANK_HOLIDAYS, SITE_CODES, ONCALL_CODES, LEAVE_TYPES, FORM_LEAVE_TYPES, LEAVE_CODES, COVER_RULES, STATUS_RANK, DEFAULT_RULES, ruleNum, groupLabel,
    norm, siteOf, kindOf, canonicalCode, isBankHoliday, bankHolidaysFor, bankHolidaysBetween, inPost, doctorById, planCode, resolveCell, coverFor, coverWithRisk, coverCode,
    worstStatus, leaveDays, balance, restrictedIn, suggestCover, availability, coverRoleFor, staffingCheck, nearestDates,
    edtDefault, edtEntitlement, edtCheck, edtOn, alPerYear, alProRata, changeoversIn,
    slotOf, predecessorOf, columns, holderOn, programmeOf, PROGRAMMES, programmeStarts, firstWednesday, suggestDates, shortNotice, rotaHorizon, addMonths
  };
  root.RotaRules = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
