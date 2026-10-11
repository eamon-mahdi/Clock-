/* Haematology rota rules: dates, codes, how a roster cell is worked out, cover checks,
 * leave day counts, balances and cover suggestions. Pure functions over the app state,
 * so they run in the browser and in Node tests. */
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
  const DOW_SHORT = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const DOW_LONG = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const fmtShort = (iso) => { const d = toDate(iso); return `${DOW_SHORT[dow(iso)]} ${d.getUTCDate()} ${MONTHS[d.getUTCMonth()]}`; };
  const fmtLong = (iso) => { const d = toDate(iso); return `${DOW_LONG[dow(iso)]} ${d.getUTCDate()} ${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}`; };
  const fmtUk = (iso) => { const [y, m, d] = iso.split('-'); return `${d}/${m}/${y}`; };

  // England & Wales bank holidays (editable in Settings)
  const BANK_HOLIDAYS = [
    '2026-01-01', '2026-04-03', '2026-04-06', '2026-05-04', '2026-05-25', '2026-08-31', '2026-12-25', '2026-12-28',
    '2027-01-01', '2027-03-26', '2027-03-29', '2027-05-03', '2027-05-31', '2027-08-30', '2027-12-27', '2027-12-28'
  ];

  // ---------- codes ----------
  const SITE_CODES = ['GUH', 'GUH LAB', 'RGH', 'RGH - DAY UNIT', 'RGH MDCU', 'NHH', 'NHH - DAY UNIT', 'OFF', 'TEACHING', 'SpR TRAINING DAY'];
  const ONCALL_CODES = ['LONG DAY', 'NIGHTS', 'TWILIGHT', 'EARLY TWILIGHT', 'WARD COVER', 'W/E SD', 'FLEXIBLE W/C', 'OFF PRE ON-CALL', 'POST NIGHTS'];
  const LEAVE_TYPES = {
    AL: { label: 'Annual leave', code: 'ANNUAL LEAVE', balance: 'al' },
    SL: { label: 'Study leave', code: 'STUDY LEAVE', balance: 'sl' },
    PL: { label: 'Professional / exam leave', code: 'PROFESSIONAL LEAVE' },
    TEACH: { label: 'Teaching (not from study leave)', code: 'TEACHING' },
    COMP: { label: 'Compassionate leave', code: 'COMPASSIONATE LEAVE' },
    TOIL: { label: 'Day off in lieu', code: 'DAY OFF IN LIEU' },
    OTHER: { label: 'Other leave', code: 'OTHER LEAVE' },
    SICK: { label: 'Sickness', code: 'SICK' }
  };
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
    if (c === 'SICK') return 'sick';
    if (siteOf(c)) return 'site';
    if (c.includes('LAB')) return 'lab';
    if (c === 'OFF') return 'off';
    if (ONCALL_CODES.includes(c)) return 'oncall';
    if (LEAVE_CODES.includes(c) || c === 'SPR TRAINING DAY') return 'leave';
    return 'other';
  }
  // Canonical spelling for codes typed in imports ("rgh-day unit" -> "RGH - DAY UNIT")
  function canonicalCode(raw) {
    let c = norm(raw).replace(/^(RGH|NHH|GUH)\s*-\s*/, '$1 - ').replace(/\s+-\s*|\s*-\s+/g, ' - ');
    const aliases = {
      'RGH DAY UNIT': 'RGH - DAY UNIT', 'RGH DU': 'RGH - DAY UNIT', 'NHH DAY UNIT': 'NHH - DAY UNIT', 'NHH DU': 'NHH - DAY UNIT',
      'RGH - MDCU': 'RGH MDCU', 'GUH - LAB': 'GUH LAB', 'NIGHT': 'NIGHTS', 'LD': 'LONG DAY', 'N': 'NIGHTS',
      'W/E SD': 'W/E SD', 'WE SD': 'W/E SD', 'OFF PRE-ON-CALL': 'OFF PRE ON-CALL', 'OFF PRE ON CALL': 'OFF PRE ON-CALL', 'OFF - PRE ON-CALL': 'OFF PRE ON-CALL',
      'POST NIGHT': 'POST NIGHTS', 'POST-NIGHTS': 'POST NIGHTS', 'SPR TRAINING': 'SpR TRAINING DAY', 'SPR TRAINING DAY': 'SpR TRAINING DAY',
      'AL': 'ANNUAL LEAVE', 'SL': 'STUDY LEAVE', 'SICK LEAVE': 'SICK'
    };
    if (aliases[c]) c = aliases[c];
    return c === 'SPR TRAINING DAY' ? 'SpR TRAINING DAY' : c;
  }

  // ---------- cover rules ----------
  const ALL_WEEKDAYS = [0, 1, 2, 3, 4];
  const TUE_FRI = [1, 2, 3, 4];
  const COVER_RULES = [
    { id: 'res-GUH', group: 'res', short: 'GUH', label: 'Resident at GUH', codes: ['GUH'], days: ALL_WEEKDAYS },
    { id: 'res-RGH', group: 'res', short: 'RGH', label: 'Resident at RGH', codes: RGH_CODES, days: ALL_WEEKDAYS },
    { id: 'res-NHH', group: 'res', short: 'NHH', label: 'Resident at NHH', codes: NHH_CODES, days: TUE_FRI },
    { id: 'mg-GUH', group: 'mg', short: 'GUH', label: 'Middle grade on GUH ward', codes: ['GUH'], days: ALL_WEEKDAYS },
    { id: 'mg-RGH', group: 'mg', short: 'RGH DU', label: 'RGH day unit', codes: RGH_CODES, days: ALL_WEEKDAYS, fallbackGroup: 'res' },
    { id: 'mg-NHH', group: 'mg', short: 'NHH DU', label: 'NHH day unit', codes: NHH_CODES, days: TUE_FRI, fallbackGroup: 'res' }
  ];
  const STATUS_RANK = { na: 0, ok: 0, risk: 1, exception: 2, gap: 3 };

  // ---------- state helpers ----------
  const isBankHoliday = (state, iso) => (state.settings.bankHolidays || BANK_HOLIDAYS).includes(iso);
  const inPost = (doc, iso) => (!doc.start || iso >= doc.start) && (!doc.end || iso <= doc.end);
  const doctorById = (state, id) => state.doctors.find((d) => d.id === id);
  const covers = (req, iso) => iso >= req.start && iso <= req.end;

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

  function leaveLabelFor(req) { return (LEAVE_TYPES[req.type] || LEAVE_TYPES.OTHER).code; }

  // Works out what one doctor is doing on one date.
  function resolveCell(state, doc, iso) {
    if (!inPost(doc, iso)) return { code: 'NOT IN POST', kind: 'gone', source: 'gone' };
    const reqs = state.requests.filter((r) => r.doctorId === doc.id && covers(r, iso) && !(r.type !== 'SICK' && (isWeekend(iso) || isBankHoliday(state, iso)) && !(state.oncall[doc.id] || {})[iso]));
    const authorised = reqs.find((r) => r.status === 'authorised');
    const pending = reqs.find((r) => r.status === 'pending') || null;
    const oncall = (state.oncall[doc.id] || {})[iso];
    const override = (state.overrides[doc.id] || {})[iso];

    let base;
    if (override) base = { code: override.code, kind: kindOf(override.code), source: 'override', note: override.reason || '' };
    else if (oncall) base = { code: oncall, kind: kindOf(oncall) === 'other' ? 'oncall' : kindOf(oncall), source: 'oncall', locked: true };
    else if (isWeekend(iso)) base = { code: '', kind: 'none', source: 'weekend' };
    else if (isBankHoliday(state, iso)) base = { code: 'BANK HOLIDAY', kind: 'bh', source: 'bh' };
    else { const p = planCode(doc, iso); base = { code: p, kind: p ? kindOf(p) : 'none', source: p ? 'plan' : 'none' }; }

    if (authorised) {
      const label = leaveLabelFor(authorised);
      if (authorised.part === 'AM' || authorised.part === 'PM') {
        const text = authorised.part === 'PM' ? `${base.code || 'OFF'} AM; PM ${label}` : `AM ${label}; ${base.code || 'OFF'} PM`;
        return Object.assign({}, base, { code: text, kind: 'half', baseCode: base.code, leave: authorised, pending });
      }
      return {
        code: label, kind: authorised.type === 'SICK' ? 'sick' : 'leave', source: 'leave', leave: authorised,
        conflict: oncall ? `Clashes with imported on-call: ${oncall}` : null
      };
    }
    return Object.assign(base, { pending });
  }

  // Code that counts for cover (half days count as present at their base site)
  const coverCode = (cell) => norm(cell.kind === 'half' ? cell.baseCode : cell.code);

  function coverFor(state, iso, opts) {
    opts = opts || {};
    if (isWeekend(iso) || isBankHoliday(state, iso)) return null;
    const absent = new Set(opts.absent || []);
    const moves = opts.moves || {};
    const present = [];
    for (const doc of state.doctors) {
      if (!inPost(doc, iso) || absent.has(doc.id)) continue;
      const cell = resolveCell(state, doc, iso);
      if (opts.pendingAbsent && cell.pending && cell.pending.part === 'FULL' && cell.kind !== 'leave' && cell.kind !== 'sick') continue;
      const code = moves[doc.id] ? norm(moves[doc.id]) : coverCode(cell);
      present.push({ doc, code });
    }
    const d = dow(iso);
    const rules = COVER_RULES.map((rule) => {
      if (!rule.days.includes(d)) return { rule, status: 'na', names: [] };
      const holders = present.filter((p) => p.doc.group === rule.group && rule.codes.includes(p.code));
      if (holders.length) return { rule, status: 'ok', names: holders.map((h) => h.doc.name) };
      if (rule.fallbackGroup) {
        const fb = present.filter((p) => p.doc.group === rule.fallbackGroup && rule.codes.includes(p.code));
        if (fb.length) return { rule, status: 'exception', names: fb.map((h) => h.doc.name) };
      }
      return { rule, status: 'gap', names: [] };
    });
    return { iso, rules };
  }

  // Cover with "at risk" where a pending request would open a gap if authorised.
  function coverWithRisk(state, iso) {
    const actual = coverFor(state, iso);
    if (!actual) return null;
    const ifPending = coverFor(state, iso, { pendingAbsent: true });
    actual.rules.forEach((r, i) => {
      if (r.status === 'ok' && STATUS_RANK[ifPending.rules[i].status] > STATUS_RANK.ok) r.status = 'risk';
    });
    return actual;
  }

  function worstStatus(cover) {
    if (!cover) return 'na';
    return cover.rules.reduce((w, r) => (STATUS_RANK[r.status] > STATUS_RANK[w] ? r.status : w), 'ok');
  }

  // ---------- leave ----------
  function leaveDays(state, start, end, part) {
    if (!start || !end || end < start) return 0;
    const n = dateRange(start, end).filter((d) => !isWeekend(d) && !isBankHoliday(state, d)).length;
    return (part === 'AM' || part === 'PM') && start === end ? n * 0.5 : n;
  }

  function balance(state, doc) {
    const sum = (type, status) => state.requests.filter((r) => r.doctorId === doc.id && r.type === type && r.status === status)
      .reduce((a, r) => a + (r.days || 0), 0);
    const mk = (type, ent) => {
      const used = sum(type, 'authorised'); const pending = sum(type, 'pending');
      return { entitlement: ent || 0, used, pending, left: (ent || 0) - used };
    };
    return { al: mk('AL', doc.alEntitlement), sl: mk('SL', doc.slEntitlement) };
  }

  const restrictedIn = (state, start, end) => (state.settings.restricted || []).filter((r) => r.start <= end && (r.end || r.start) >= start)
    .sort((a, b) => (a.start < b.start ? -1 : 1));

  // Suggest who could cover each working day of a leave request.
  function suggestCover(state, doctorId, start, end, part) {
    const doc = doctorById(state, doctorId);
    if (!doc || !start || !end || end < start || part === 'AM' || part === 'PM') return [];
    const out = [];
    for (const iso of dateRange(start, end)) {
      if (isWeekend(iso) || isBankHoliday(state, iso) || !inPost(doc, iso)) continue;
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
          const cc = resolveCell(state, cand, iso);
          if (cc.pending && cc.pending.part === 'FULL') continue;
          let type;
          if (cc.kind === 'off' || (cc.kind === 'none' && cc.source !== 'weekend')) type = 'dayoff';
          else if ((cc.kind === 'site' || cc.kind === 'lab') && !b.rule.codes.includes(coverCode(cc))) type = 'move';
          else continue;
          const after = coverFor(state, iso, { absent: [doctorId], moves: { [cand.id]: target } });
          const newGaps = after.rules.filter((r, i) => r.status === 'gap' && before.rules[i].status !== 'gap');
          if (newGaps.length) continue;
          const exceptions = after.rules.filter((r, i) => r.status === 'exception' && before.rules[i].status !== 'exception').map((r) => r.rule.label);
          const score = exceptions.length * 10 + (type === 'dayoff' ? 5 : 0);
          const text = type === 'dayoff'
            ? `${cand.name} is off that day and works ${target} instead (swaps their day off)`
            : `${cand.name} moves from ${cc.code} to ${target}`;
          options.push({ doctorId: cand.id, name: cand.name, from: cc.code || 'OFF', to: target, type, exceptions, score, text });
        }
      }
      options.sort((a, b) => a.score - b.score || a.name.localeCompare(b.name));
      out.push({ date: iso, needed: true, from: cell.code, broken: broken.map((r) => r.rule.label), options: options.slice(0, 5) });
    }
    return out;
  }

  const api = {
    toDate, toIso, addDays, dow, isWeekend, mondayOf, daysBetween, dateRange, fmtShort, fmtLong, fmtUk, DOW_SHORT, DOW_LONG,
    BANK_HOLIDAYS, SITE_CODES, ONCALL_CODES, LEAVE_TYPES, LEAVE_CODES, COVER_RULES, STATUS_RANK,
    norm, siteOf, kindOf, canonicalCode, isBankHoliday, inPost, doctorById, planCode, resolveCell, coverFor, coverWithRisk,
    worstStatus, leaveDays, balance, restrictedIn, suggestCover, coverCode
  };
  root.RotaRules = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
