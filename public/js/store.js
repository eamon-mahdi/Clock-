/* App state, demo data, the public actions (send a leave request, check or withdraw it),
 * the authoriser actions, what viewers are allowed to see, and the store used by the screens.
 * The same file runs in the browser and in the Cloudflare server. */
(function (root) {
  'use strict';
  const R = root.RotaRules || (typeof require !== 'undefined' ? require('./rules.js') : null);
  const LOCAL_KEY = 'haem-rota-demo-v2';
  const uid = (p) => p + '-' + Math.random().toString(36).slice(2, 10);
  const nowIso = () => new Date().toISOString();
  const clone = (x) => JSON.parse(JSON.stringify(x));

  // ---------- default text for the Rules tab (from the placement email, updated) ----------
  const DEFAULT_RULES_TEXT = [
    { id: 'placement', title: 'Placement and rota access', items: [
      'The current placement started on Wednesday 5 August 2026.',
      'Changeovers: Foundation doctors on the first Wednesday of August, December and April; IMTs on the first Wednesday of August and February; training-grade registrars on 1 August or 1 February (usually 6 months, sometimes extended to 12). Other doctors start and finish on agreed dates. Changeovers show as blue bands, and a new doctor continues in the same column as the doctor they replace.',
      'This website is the live rota. Viewing needs no password. It replaces the shared Doctor Rota.xlsx.',
      'On-calls come from the trust medical rota and are imported each week. Check the rota shows your on-calls correctly; if anything is wrong, tell the rota authorisers and the medical rota team.',
      'The rota can change at short notice. Check it before each week and before each shift.'] },
    { id: 'moving', title: 'Moving between sites', items: [
      'Resident doctors may need to move between sites because of sickness or service needs. Moves are made only by the rota authorisers and are logged.',
      'Senior doctors may be asked to cover a day unit at RGH or NHH; this is entered on the rota by the authorisers.',
      'If a change affects you, the authorisers will tell you. If you have site preferences, travel concerns or other constraints, tell them as early as possible.',
      'Do not swap daytime sites between yourselves without telling the authorisers, or the cover checks will be wrong.'] },
    { id: 'annual', title: 'Annual leave', items: [
      'Request all annual leave with the leave form on this website at least 6 weeks before the first day, in line with trust policy.',
      'Requests with less than 6 weeks\' notice are considered only in extenuating circumstances. Tick the box on the form and explain why; the request may still be declined.',
      'Name the specific doctor covering you (day unit and/or GUH), not a grade. Check with them first.',
      'Leave is not confirmed until an authoriser approves it. Do not book travel until it is approved.',
      'Authorisers decide day by day. If there are not enough doctors on some days, they may approve the other days and decline those, or suggest other dates. You will be told exactly which days are approved.',
      'Your allowance is set from your contract: pro-rata for time in post (resident doctors) or the leave year (senior doctors), and for less-than-full-time working. Carried-over or accrued days can be added by the authorisers. See your remaining days under "Check my request".',
      'Dates marked NO AL are not normally approved; please avoid dates marked AVOID.',
      'Leave does not cancel on-call shifts. Arrange any on-call swap with the medical rota team before you request leave.'] },
    { id: 'study', title: 'Study leave', items: [
      'Request study leave on the form here and also through Intrepid (or CODI), following the HEIW study leave process.',
      'The same notice period and staffing checks apply as for annual leave.',
      'Study leave comes from your study leave allowance, not your annual leave.'] },
    { id: 'sick', title: 'Sickness and other absence', items: [
      'If you are unwell, report it on the form as early as possible on the day, and follow the trust sickness reporting policy (including telephoning your clinical area).',
      'Sickness shows on the rota only as "Unavailable". No reasons or medical details are recorded here, and it never comes off annual leave.',
      'For other absence (for example compassionate, parental or exam leave) choose "Other" on the form and say what it is.'] },
    { id: 'edt', title: 'EDT and SDT', items: [
      'Taken as a morning or an afternoon session and shown on the rota.',
      'Resident doctors (F1, F2, IMT, JCF) take it at GUH only, and only when enough other resident doctors are working at GUH in that session (see the number above).',
      'Senior doctors (registrars, specialty doctors, specialist grade, senior clinical fellows) can take it at any site, as long as GUH keeps its senior doctor.',
      'Sessions are booked by the authorisers. Ask them as early as possible.'] },
    { id: 'induction', title: 'Induction and first day', items: [
      'There are enough staff at GUH on the first day of each rotation.',
      'If you know of a formal induction at your site, tell the rota authorisers so it can go on the rota.'] },
    { id: 'where', title: 'Where to go', items: [
      'Grange University Hospital: Ward B4.',
      'Nevill Hall Hospital: Ward 2/4 (Raglan Ward).',
      'Royal Gwent Hospital: Medical Day Case Unit (MDCU), 3rd floor, opposite the canteen.'] },
    { id: 'questions', title: 'Questions', items: [
      'Contact the rota authorisers directly. Please avoid "Reply all" on rota emails.',
      'Only the rota authorisers can change the rota. Every change is logged with who made it and when.'] }
  ];
  const OLD_RULES_TEXT_JSON = JSON.stringify([
    { id: 'placement', title: 'Placement and rota access', items: [
      'The current placement started on Wednesday 5 August 2026. Foundation doctors change every 4 months and IMTs every 6 months; other changes are shown as blue bands on the roster.',
      'This website is the live rota. Viewing needs no password. It replaces the shared Doctor Rota.xlsx.',
      'Check the rota shows your on-call commitments correctly. If anything is wrong or has changed, tell the rota authorisers.'] },
    { id: 'moving', title: 'Moving between sites', items: [
      'Resident doctors may need to move between sites because of sickness or service needs.',
      'Check the live rota regularly. If a change affects you, the authorisers will tell you.',
      'If you have site preferences, travel concerns or other constraints, tell the rota authorisers as early as possible.'] },
    { id: 'leave', title: 'Requesting leave', items: [
      'All annual, study, sick and other leave is requested with the leave form on this website. It replaces the Microsoft Form and PDF.',
      'Ask well ahead (see the notice period above).',
      'Choose your own name, and name the specific doctor covering you (day unit and/or GUH), not a grade.',
      'Study leave must also be requested through Intrepid (or CODI).',
      'Leave is not confirmed until a rota authoriser approves it. Check progress any time under "Check my request".',
      'If there would not be enough doctors on some days, the authorisers see a warning. They may approve the other days, offer another date, or arrange cover.',
      'Sickness: report it on the form on the day. It shows on the rota only as "Unavailable" and never comes off annual leave.'] },
    { id: 'edt', title: 'EDT and SDT', items: [
      'Taken as a morning or an afternoon session and shown on the rota.',
      'Residents (F1, F2, IMT, JCF) take it at GUH only, and only when enough other residents are working at GUH in that session (see the number above).',
      'Senior doctors (registrars, specialty doctors, specialist grade, senior clinical fellows) can take it at any site, as long as that site keeps its senior cover.'] },
    { id: 'induction', title: 'Induction and first day', items: [
      'There are enough staff at GUH on the first day.',
      'If you know of a formal induction at your site, tell the rota authorisers so it can go on the rota.'] },
    { id: 'where', title: 'Where to go', items: [
      'Grange University Hospital: Ward B4.',
      'Nevill Hall Hospital: Ward 2/4 (Raglan Ward).',
      'Royal Gwent Hospital: Medical Day Case Unit (MDCU), 3rd floor, opposite the canteen.'] },
    { id: 'questions', title: 'Questions', items: [
      'Contact the rota authorisers directly. Please avoid "Reply all" on rota emails.',
      'Only the rota authorisers can change the rota. Every change is logged with who made it and when.'] }
  ]);

  // Bring older saved rotas up to date (run on every load).
  function upgrade(state) {
    if (!state || !state.settings) return state;
    const st = state.settings;
    if ((st.schema || 0) < 3) {
      st.rules = Object.assign({}, R.DEFAULT_RULES, st.rules || {}, { minSenRGH: 0, minSenNHH: 0 });
      if (!st.rulesText || JSON.stringify(st.rulesText) === OLD_RULES_TEXT_JSON) st.rulesText = clone(DEFAULT_RULES_TEXT);
      if (st.rotation) delete st.rotation.end;
      st.consultantOfWeek = st.consultantOfWeek || {};
      st.schema = 3;
    }
    return state;
  }

  function baseSettings() {
    return {
      bankHolidays: R.BANK_HOLIDAYS.slice(), restricted: [], consultantOfWeek: {}, rules: Object.assign({}, R.DEFAULT_RULES),
      rulesText: clone(DEFAULT_RULES_TEXT), changeovers: [], rotation: { start: '2026-08-05' }, schema: 3,
      authorisers: ['Authoriser 1', 'Authoriser 2', 'Authoriser 3'], authoriserEmails: [], demo: false
    };
  }
  function emptyState() {
    return { version: 2, settings: baseSettings(), doctors: [], oncall: {}, overrides: {}, requests: [], edt: [], audit: [], notifications: [] };
  }

  // ---------- demo data (all names are fictional) ----------
  function demoState(todayIso) {
    const today = todayIso || R.toIso(new Date());
    const mon = R.mondayOf(today);
    const plan = (a, b) => ({ A: a.split('|'), B: b ? b.split('|') : null, anchor: mon });
    const rotStart = '2026-08-05', rotEnd = '2026-12-01', imtEnd = '2027-02-02';
    const doctors = [
      { id: 'r1', name: 'E. Ward', grade: 'F1', group: 'res', start: rotStart, end: rotEnd, plan: plan('GUH|GUH|GUH|GUH|GUH') },
      { id: 'r2', name: 'P. Nair', grade: 'F2', group: 'res', start: rotStart, end: rotEnd, plan: plan('RGH|NHH|NHH|NHH|NHH') },
      { id: 'r3', name: 'J. Reid', grade: 'F2', group: 'res', start: rotStart, end: R.addDays(mon, -10), plan: plan('RGH|RGH|RGH|RGH|RGH'), note: 'Left the rotation early' },
      { id: 'r8', name: 'K. Ahmed', grade: 'F2', group: 'res', slot: 'r3', start: R.addDays(mon, 7), end: rotEnd, plan: plan('RGH|RGH|RGH|RGH|RGH') },
      { id: 'r4', name: 'T. Hughes', grade: 'F2 (on calls NHH)', group: 'res', start: rotStart, end: rotEnd, plan: plan('GUH|GUH|RGH|GUH|GUH') },
      { id: 'r5', name: 'S. Malik', grade: 'IMT2', group: 'res', start: rotStart, end: imtEnd, plan: plan('GUH|RGH|RGH|RGH|RGH') },
      { id: 'r6', name: 'B. Carter', grade: 'IMT3', group: 'res', start: rotStart, end: imtEnd, fte: 0.8, plan: plan('GUH|GUH|GUH|RGH|GUH') },
      { id: 'r7', name: 'L. Osei', grade: 'JCF', group: 'res', start: rotStart, end: '2027-08-03', plan: plan('RGH|RGH|GUH|RGH|RGH') },
      { id: 'r9', name: 'New F1', grade: 'F1', group: 'res', slot: 'r1', placeholder: true, start: '2026-12-02', end: '2027-04-06', email: '', plan: plan('GUH|GUH|GUH|GUH|GUH') },
      { id: 'r10', name: 'New F2', grade: 'F2', group: 'res', slot: 'r2', placeholder: true, start: '2026-12-02', end: '2027-04-06', email: '', plan: plan('RGH|NHH|NHH|NHH|NHH') },
      { id: 'm1', name: 'Omar', grade: 'Registrar ST6', group: 'mg', start: '2024-01-01', end: '', plan: plan('GUH|GUH|GUH|GUH|GUH') },
      { id: 'm2', name: 'Hannah', grade: 'Registrar ST4', group: 'mg', start: '2024-01-01', end: '', plan: plan('RGH - DAY UNIT|RGH - DAY UNIT|RGH - DAY UNIT|RGH - DAY UNIT|RGH - DAY UNIT') },
      { id: 'm3', name: 'Daniel', grade: 'Specialty doctor', group: 'mg', start: '2024-01-01', end: '', plan: plan('GUH LAB|NHH - DAY UNIT|NHH - DAY UNIT|NHH - DAY UNIT|NHH - DAY UNIT') },
      { id: 'm4', name: 'Zara', grade: 'Registrar ST5', group: 'mg', start: '2024-01-01', end: '', plan: plan('RGH MDCU|RGH|RGH|RGH|RGH') },
      { id: 'm5', name: 'Meera', grade: 'Specialty doctor', group: 'mg', start: '2024-01-01', end: '', plan: plan('GUH LAB|GUH LAB|GUH|GUH|OFF', 'GUH LAB|GUH|OFF|GUH|GUH') },
      { id: 'm6', name: 'Yusuf', grade: 'Senior clinical fellow', group: 'mg', start: '2024-01-01', end: '', plan: plan('RGH MDCU|RGH|RGH|OFF|RGH') }
    ].map((d) => {
      const doc = Object.assign({ email: d.placeholder ? '' : d.name.toLowerCase().replace(/[^a-z]+/g, '.').replace(/^\.|\.$/g, '') + '@example.nhs.uk', aliases: [], fte: 1, slEntitlement: d.group === 'mg' ? 10 : 15 }, d);
      doc.alEntitlement = R.alProRata(doc).days;
      return doc;
    });
    const oncall = {};
    const residents = doctors.filter((d) => d.group === 'res' && !['r3', 'r8', 'r9', 'r10'].includes(d.id));
    for (let w = -1; w < 8; w++) {
      const wk = R.addDays(mon, w * 7);
      residents.forEach((d, i) => {
        const set = (offset, code) => { const iso = R.addDays(wk, offset); if (R.inPost(d, iso)) (oncall[d.id] = oncall[d.id] || {})[iso] = code; };
        const k = (w + i * 2 + 12) % 6;
        if (k === 0) { [0, 1, 2, 3].forEach((o) => set(o, 'NIGHTS')); set(4, 'POST NIGHTS'); }
        if (k === 3) { set(4, 'OFF PRE ON-CALL'); set(5, 'LONG DAY'); set(6, 'LONG DAY'); }
        if (k === 1) { set(1, 'TWILIGHT'); set(2, 'TWILIGHT'); }
      });
    }
    const consultants = ['Dr Fraser', 'Dr Lin', 'Dr Okoro', 'Dr Grant', 'Dr Shah'];
    const consultantOfWeek = {};
    for (let w = -10; w < 12; w++) consultantOfWeek[R.addDays(mon, w * 7)] = consultants[(w + 20) % consultants.length];
    const settings = Object.assign(baseSettings(), {
      demo: true, consultantOfWeek, authorisers: ['J. Hale', 'S. Price', 'R. Dunn'],
      restricted: [{ id: uid('x'), start: R.addDays(mon, 14), end: R.addDays(mon, 15), level: 'NO AL', note: 'Short-staffed: two doctors on induction' }],
      changeovers: [{ id: uid('c'), date: '2027-02-03', label: 'IMT 6-month changeover (names to follow)' }]
    });
    const state = { version: 2, settings, doctors, oncall, overrides: {}, requests: [], edt: [], audit: [], notifications: [] };
    const add = (doctorId, type, start, end, part, status, note, extra) => {
      const doc = R.doctorById(state, doctorId);
      const r = Object.assign({ id: uid('lr'), doctorId, email: doc.email, type, start, end, part, note: note || '', status, cover: [], excludedDates: [],
        submittedAt: R.addDays(today, -6) + 'T09:12:00Z', days: R.leaveDays(state, start, end, part) }, extra || {});
      if (status !== 'pending') Object.assign(r, { decidedBy: 'J. Hale', decidedAt: R.addDays(today, -5) + 'T14:30:00Z' });
      state.requests.push(r);
      return r;
    };
    add('m2', 'AL', R.addDays(mon, 9), R.addDays(mon, 11), 'FULL', 'authorised', 'Half term');
    add('r6', 'AL', R.addDays(mon, 17), R.addDays(mon, 17), 'PM', 'authorised', 'Appointment');
    add('r1', 'AL', R.addDays(mon, 43), R.addDays(mon, 46), 'FULL', 'authorised', 'Holiday');
    add('m6', 'SICK', R.addDays(mon, -3), R.addDays(mon, -3), 'FULL', 'authorised', '', { decidedBy: 'Self-reported' });
    const reqs = [
      ['r2', 'AL', R.addDays(mon, 42), R.addDays(mon, 46), 'Family wedding abroad', { dayUnitCoverId: 'r8' }],
      ['m5', 'AL', R.addDays(mon, 45), R.addDays(mon, 45), "Sister's wedding", { guhCoverId: 'm1' }],
      ['r7', 'SL', R.addDays(mon, 52), R.addDays(mon, 52), 'Regional teaching course', { intrepid: true, dayUnitCoverId: 'r5' }]
    ];
    for (const [id, type, s, e, note, extra] of reqs) {
      const r = add(id, type, s, e, 'FULL', 'pending', note, extra);
      const chk = R.staffingCheck(state, { doctorId: id, start: s, end: e, part: 'FULL', dayUnitCoverId: r.dayUnitCoverId, guhCoverId: r.guhCoverId, requestId: r.id });
      r.cover = chk.cover; r.check = { bad: chk.bad, warn: chk.warn };
    }
    const edtDays = [['r5', 9, 'PM'], ['m1', 3, 'PM'], ['m2', 1, 'PM'], ['r1', 10, 'AM']];
    for (const [id, off, ses] of edtDays) {
      const iso = R.addDays(mon, off);
      const chk = R.edtCheck(state, id, iso, ses);
      if (chk.ok) state.edt.push({ id: uid('e'), doctorId: id, date: iso, session: ses, type: chk.type, bookedBy: 'J. Hale' });
    }
    state.audit.push({ at: nowIso(), by: 'System', action: 'Demo data loaded', detail: 'Fictional doctors, weekly plans, on-calls, EDT and leave requests' });
    state.notifications.push({ id: uid('n'), to: 'auth', at: nowIso(), text: '3 leave requests are waiting for a decision.', read: false });
    return state;
  }

  // ---------- helpers shared by server and browser ----------
  const docName = (state, id) => (R.doctorById(state, id) || { name: '?' }).name;
  const rangeText = (a, b) => (a === b ? R.fmtShort(a) : `${R.fmtShort(a)} – ${R.fmtShort(b)}`);
  function log(state, by, action, detail) {
    state.audit.unshift({ at: nowIso(), by, action, detail });
    if (state.audit.length > 3000) state.audit.length = 3000;
  }
  function notify(state, to, text) {
    state.notifications.unshift({ id: uid('n'), to, at: nowIso(), text, read: false });
    if (state.notifications.length > 600) state.notifications.length = 600;
  }
  const emailMatches = (doc, email) => {
    const a = String(email || '').trim().toLowerCase();
    return !!a && (!doc.email || doc.email.trim().toLowerCase() === a);
  };
  const authorEmails = (state) => ((state.settings && state.settings.authoriserEmails) || []).filter(Boolean);

  // A doctor (no password) sends a leave request. Returns {request, check, emails}.
  function submitLeave(state, p, opts) {
    const today = (opts && opts.today) || R.toIso(new Date());
    const doc = R.doctorById(state, p.doctorId);
    if (!doc || doc.placeholder) throw new Error('Choose your name from the list.');
    if (!emailMatches(doc, p.email)) throw new Error(doc.email ? 'That NHS email does not match the staff list for your name. Check it, or ask an authoriser to update it.' : 'Enter your NHS email address.');
    if (!R.FORM_LEAVE_TYPES.includes(p.type)) throw new Error('Choose a leave type.');
    if (!R.isIso(p.start) || !R.isIso(p.end) || p.end < p.start) throw new Error('Choose a start date and an end date on or after it.');
    if (R.daysBetween(p.start, p.end) > 92) throw new Error('A single request can cover up to 3 months. Split longer leave into separate requests.');
    if (p.type === 'SL' && !p.intrepid) throw new Error('Study leave must also be requested through Intrepid (or CODI). Tick the box to confirm.');
    if (p.type === 'OTHER' && !String(p.otherText || '').trim()) throw new Error('Describe the type of leave under "Other".');
    const part = p.start === p.end && (p.part === 'AM' || p.part === 'PM') ? p.part : 'FULL';
    const short = p.type === 'AL' && R.shortNotice(state, p.start, today);
    const why = String(p.extenuatingReason || '').trim();
    if (short && (!p.extenuating || why.length < 10)) throw new Error(`Annual leave needs at least ${R.ruleNum(state, 'noticeWeeks')} weeks' notice (trust policy). If there are extenuating circumstances, tick the box and explain them.`);
    const overlap = state.requests.find((r) => r.doctorId === doc.id && ['pending', 'hold', 'authorised'].includes(r.status) && r.start <= p.end && r.end >= p.start);
    if (overlap) throw new Error(`You already have ${overlap.status === 'authorised' ? 'approved' : 'a pending request for'} leave on ${rangeText(overlap.start, overlap.end)}.`);
    const cleanId = (id) => (id && R.doctorById(state, id) && id !== doc.id ? id : '');
    const r = {
      id: uid('lr'), doctorId: doc.id, email: String(p.email).trim(), type: p.type, otherText: String(p.otherText || '').slice(0, 200),
      start: p.start, end: p.end, part, note: String(p.note || '').slice(0, 1000), intrepid: !!p.intrepid, shortNotice: short, extenuatingReason: short ? why.slice(0, 600) : '',
      dayUnitCoverId: cleanId(p.dayUnitCoverId), guhCoverId: cleanId(p.guhCoverId), excludedDates: [],
      submittedAt: nowIso(), days: R.leaveDays(state, p.start, p.end, part), status: p.type === 'SICK' ? 'authorised' : 'pending', cover: []
    };
    const check = R.staffingCheck(state, { doctorId: doc.id, start: r.start, end: r.end, part, dayUnitCoverId: r.dayUnitCoverId, guhCoverId: r.guhCoverId, requestId: r.id });
    r.cover = check.cover;
    r.check = { bad: check.bad, warn: check.warn };
    const label = R.LEAVE_TYPES[r.type].label;
    const when = rangeText(r.start, r.end);
    const emails = [];
    if (r.type === 'SICK') {
      Object.assign(r, { decidedBy: 'Self-reported', decidedAt: nowIso() });
      notify(state, 'auth', `${doc.name} reported sick: ${when}.${check.bad ? ' Not enough doctors on ' + check.bad + ' day(s).' : ''}`);
      log(state, doc.name, 'Sickness reported', `${R.fmtUk(r.start)} to ${R.fmtUk(r.end)}`);
      authorEmails(state).forEach((to) => emails.push({ to, subject: `Sickness reported: ${doc.name}, ${when}`, text: `${doc.name} (${doc.grade}) has reported sick for ${when}.${check.bad ? `\n\nNot enough doctors on ${check.bad} day(s). Please arrange cover.` : ''}\n\nOpen the rota to review cover.` }));
    } else {
      notify(state, 'auth', `New ${label.toLowerCase()} request from ${doc.name}: ${when}.${check.bad ? ' Not enough doctors on ' + check.bad + ' day(s).' : ''}`);
      log(state, doc.name, 'Leave requested', `${label}, ${R.fmtUk(r.start)} to ${R.fmtUk(r.end)} (${r.days} day${r.days === 1 ? '' : 's'})${short ? '; less than the notice period (extenuating circumstances given)' : ''}${check.bad ? `; flagged: not enough doctors on ${check.bad} day(s)` : ''}`);
      authorEmails(state).forEach((to) => emails.push({ to, subject: `${check.bad ? 'Needs a decision (not enough doctors): ' : 'Leave request: '}${doc.name}, ${when}`, text: `${doc.name} (${doc.grade}) asked for ${label.toLowerCase()}: ${when}, ${r.days} working day(s).${short ? `\n\nLess than ${R.ruleNum(state, 'noticeWeeks')} weeks' notice. Extenuating circumstances given: "${why}"` : ''}${check.bad ? `\n\nNot enough doctors on ${check.bad} day(s).` : ''}${check.warn ? `\nWarnings on ${check.warn} day(s).` : ''}\n\nSign in as an authoriser to decide.` }));
      emails.push({ to: r.email, subject: `We've got your leave request: ${when}`, text: `Your ${label.toLowerCase()} request for ${when} has been sent to the rota authorisers. You'll get an email when it's decided. You can check progress under "Check my request".` });
      const covers = new Set(r.cover.map((c) => c.doctorId));
      covers.forEach((id) => { const cd = R.doctorById(state, id); if (cd && cd.email) emails.push({ to: cd.email, subject: `You've been named as cover: ${doc.name}, ${when}`, text: `${doc.name} has named you as cover while on leave (${when}), if the request is approved:\n${r.cover.filter((c) => c.doctorId === id).map((c) => c.text).join('\n')}` }); });
    }
    notify(state, 'd:' + doc.id, r.type === 'SICK' ? `Sickness recorded for ${when}.` : `Request sent: ${label.toLowerCase()} ${when}.`);
    state.requests.push(r);
    return { request: r, check, emails };
  }
  function withdrawLeave(state, id, email) {
    const r = state.requests.find((x) => x.id === id);
    if (!r) throw new Error('Request not found.');
    const doc = R.doctorById(state, r.doctorId);
    if (!emailMatches(doc || {}, email) || String(r.email || '').toLowerCase() !== String(email || '').trim().toLowerCase()) throw new Error('That email does not match the request.');
    if (!['pending', 'hold'].includes(r.status)) throw new Error('Only requests still waiting can be withdrawn.');
    r.status = 'cancelled';
    log(state, doc ? doc.name : '?', 'Leave request withdrawn', `${R.fmtUk(r.start)} to ${R.fmtUk(r.end)}`);
    notify(state, 'auth', `${doc ? doc.name : 'A doctor'} withdrew their request for ${rangeText(r.start, r.end)}.`);
    return r;
  }
  function myRequests(state, doctorId, email) {
    const doc = R.doctorById(state, doctorId);
    if (!doc || !emailMatches(doc, email)) throw new Error('Name and NHS email do not match the staff list.');
    const mail = String(email).trim().toLowerCase();
    const bal = R.balance(state, doc);
    return {
      balance: { al: bal.al, sl: bal.sl, alBasis: R.alProRata(doc).basis, start: doc.start || '', end: doc.end || '' },
      requests: state.requests.filter((r) => r.doctorId === doc.id && (!r.email || r.email.toLowerCase() === mail || doc.email)).sort((a, b) => (a.submittedAt < b.submittedAt ? 1 : -1)),
      notifications: state.notifications.filter((n) => n.to === 'd:' + doc.id).slice(0, 30)
    };
  }

  // What anyone with the link can see (no emails, notes, reasons, sickness or audit).
  function publicView(state) {
    const s = clone(state);
    s.doctors.forEach((d) => { delete d.email; delete d.alEntitlement; delete d.slEntitlement; delete d.alExtra; delete d.slExtra; delete d.note; });
    s.requests = s.requests.filter((r) => ['pending', 'hold', 'authorised'].includes(r.status)).map((r) => ({
      id: r.id, doctorId: r.doctorId, type: r.type === 'SICK' ? 'UNAVAIL' : (['AL', 'SL'].includes(r.type) ? r.type : 'OTHER'),
      start: r.start, end: r.end, part: r.part, status: r.status, excludedDates: r.excludedDates || [], days: r.days, cover: (r.cover || []).map((c) => ({ date: c.date, doctorId: c.doctorId, to: c.to }))
    }));
    Object.values(s.overrides || {}).forEach((m) => Object.values(m).forEach((o) => { delete o.reason; delete o.by; }));
    s.edt = (s.edt || []).map((e) => ({ id: e.id, doctorId: e.doctorId, date: e.date, session: e.session, type: e.type }));
    s.audit = []; s.notifications = [];
    s.settings.authoriserEmails = [];
    s.public = true;
    return s;
  }

  // ---------- authoriser actions (pure functions on state; return emails to send) ----------
  const ops = {
    setCell(state, by, doctorId, iso, code, reason) {
      const d = R.doctorById(state, doctorId);
      state.overrides[doctorId] = state.overrides[doctorId] || {};
      const before = R.resolveCell(state, d, iso).code || '(blank)';
      if (code) state.overrides[doctorId][iso] = { code: R.canonicalCode(code), reason: reason || '', by, at: nowIso() };
      else delete state.overrides[doctorId][iso];
      log(state, by, 'Roster changed', `${d.name}, ${R.fmtShort(iso)}: ${before} → ${R.resolveCell(state, d, iso).code || '(blank)'}${reason ? ` (${reason})` : ''}`);
      return [];
    },
    // decision: 'approve' | 'approveCovered' | 'approveDays' (opts.days = dates to approve) | 'hold' | 'decline'
    decide(state, by, id, decision, message, opts) {
      opts = opts || {};
      const r = state.requests.find((x) => x.id === id);
      if (!r || !['pending', 'hold'].includes(r.status)) throw new Error('This request has already been decided.');
      const doc = R.doctorById(state, r.doctorId);
      const check = R.staffingCheck(state, { doctorId: r.doctorId, start: r.start, end: r.end, part: r.part, dayUnitCoverId: r.dayUnitCoverId, guhCoverId: r.guhCoverId, requestId: r.id });
      const badDays = check.days.filter((d) => d.status === 'bad').map((d) => d.date);
      const msg = String(message || '').trim();
      const workdays = R.dateRange(r.start, r.end).filter((d) => R.isWorkday(state, d) && R.inPost(doc, d));
      let approveSet = null;
      if (decision === 'approveDays') {
        approveSet = new Set((opts.days || []).filter((d) => workdays.includes(d)));
        if (!approveSet.size) throw new Error('Tick at least one day to approve, or decline the request.');
        if (approveSet.size === workdays.length) decision = 'approve';
      }
      const approvingBad = decision === 'approve' ? badDays : approveSet ? badDays.filter((d) => approveSet.has(d)) : [];
      if (approvingBad.length && !msg) throw new Error(`Not enough doctors on ${approvingBad.map(R.fmtShort).join(', ')}. Give a reason (e.g. "locum booked") to approve ${approvingBad.length === 1 ? 'that day' : 'those days'}, or untick ${approvingBad.length === 1 ? 'it' : 'them'}.`);
      if ((decision === 'decline' || decision === 'hold') && !msg) throw new Error(decision === 'decline' ? 'Give a reason for declining. It is sent to the doctor.' : 'Write the question for the doctor.');
      const label = R.LEAVE_TYPES[r.type].label;
      const when = rangeText(r.start, r.end);
      let subject, text, status;
      if (decision === 'hold') {
        r.status = 'hold'; r.question = msg;
        subject = `Question about your leave request: ${when}`; text = `${by} has a question about your ${label.toLowerCase()} request (${when}):\n\n"${msg}"\n\nPlease reply to the rota authorisers.`;
        status = 'Question sent';
      } else if (decision === 'decline') {
        r.status = 'declined';
        subject = `Leave not approved: ${when}`; text = `Your ${label.toLowerCase()} request (${when}) was not approved by ${by}.\n\nReason: ${msg}`;
        status = 'Leave declined';
      } else {
        if (approveSet) r.excludedDates = workdays.filter((d) => !approveSet.has(d));
        if (decision === 'approveCovered') {
          r.excludedDates = badDays;
          if (!R.dateRange(r.start, r.end).some((d) => R.isWorkday(state, d) && !badDays.includes(d))) throw new Error('No days have enough doctors. Decline, hold or approve all with a reason instead.');
        }
        r.status = 'authorised';
        r.days = R.leaveDays(state, r.start, r.end, r.part, r.excludedDates);
        r.cover = check.cover.filter((c) => !(r.excludedDates || []).includes(c.date));
        for (const c of r.cover) {
          state.overrides[c.doctorId] = state.overrides[c.doctorId] || {};
          state.overrides[c.doctorId][c.date] = { code: c.to, reason: `Covering ${doc.name}'s leave`, by, at: nowIso() };
          notify(state, 'd:' + c.doctorId, `You are covering ${c.to} on ${R.fmtShort(c.date)} for ${doc.name}.`);
        }
        const partial = (r.excludedDates || []).length;
        subject = partial ? `Leave partly approved: ${when}` : `Leave approved: ${when}`;
        text = partial
          ? `Your ${label.toLowerCase()} request (${when}) was partly approved by ${by}.\n\nApproved: ${workdays.filter((d) => !r.excludedDates.includes(d)).map((d) => R.fmtShort(d)).join(', ')}.\nNot approved: ${r.excludedDates.map((d) => R.fmtShort(d)).join(', ')}.${check.alternatives.length ? `\nDays with cover nearby: ${check.alternatives.map((d) => R.fmtShort(d)).join(', ')}.` : ''}${msg ? `\n\n"${msg}"` : ''}`
          : `Your ${label.toLowerCase()} request (${when}) was approved by ${by}.${msg ? `\n\n"${msg}"` : ''}`;
        status = partial ? 'Leave partly approved' : 'Leave approved';
        if (approvingBad.length) r.overrideReason = msg;
      }
      r.decidedBy = by; r.decidedAt = nowIso(); r.message = msg;
      notify(state, 'd:' + r.doctorId, `${status}: ${label.toLowerCase()} ${when}.${msg ? ' "' + msg + '"' : ''}`);
      log(state, by, status, `${doc.name}: ${label}, ${R.fmtUk(r.start)} to ${R.fmtUk(r.end)}${(r.excludedDates || []).length ? `; not approved ${r.excludedDates.map(R.fmtUk).join(', ')}` : ''}${msg ? `; "${msg}"` : ''}`);
      const emails = [];
      const to = r.email || doc.email;
      if (to) emails.push({ to, subject, text });
      if (r.status === 'authorised') new Set(r.cover.map((c) => c.doctorId)).forEach((cid) => { const cd = R.doctorById(state, cid); if (cd && cd.email) emails.push({ to: cd.email, subject: `You're covering for ${doc.name}`, text: r.cover.filter((c) => c.doctorId === cid).map((c) => c.text).join('\n') }); });
      return emails;
    },
    bookEdt(state, by, doctorId, iso, session, reason) {
      const chk = R.edtCheck(state, doctorId, iso, session);
      if (!chk.ok && !String(reason || '').trim()) throw new Error('This session breaks a rule. Give a reason to book it anyway.');
      const doc = R.doctorById(state, doctorId);
      state.edt.push({ id: uid('e'), doctorId, date: iso, session, type: chk.type || R.edtDefault(doc).type, bookedBy: by, reason: reason || '' });
      notify(state, 'd:' + doctorId, `${chk.type || 'EDT'} booked: ${R.fmtShort(iso)} ${session === 'AM' ? 'morning' : 'afternoon'}.`);
      log(state, by, 'EDT/SDT booked', `${doc.name}: ${R.fmtUk(iso)} ${session}${reason ? ` (override: ${reason})` : ''}`);
      return doc.email ? [{ to: doc.email, subject: `${chk.type || 'EDT'} booked: ${R.fmtShort(iso)} ${session}`, text: `${by} booked your ${chk.type || 'EDT'} for ${R.fmtLong(iso)} (${session === 'AM' ? 'morning' : 'afternoon'}).` }] : [];
    },
    removeEdt(state, by, id) {
      const e = state.edt.find((x) => x.id === id);
      state.edt = state.edt.filter((x) => x.id !== id);
      if (e) log(state, by, 'EDT/SDT removed', `${docName(state, e.doctorId)}: ${R.fmtUk(e.date)} ${e.session}`);
      return [];
    },
    saveDoctor(state, by, d) {
      const existing = R.doctorById(state, d.id);
      if (existing) {
        const changes = Object.keys(d).filter((k) => k !== 'id' && JSON.stringify(existing[k]) !== JSON.stringify(d[k]));
        Object.assign(existing, d);
        if (changes.length) log(state, by, 'Doctor updated', `${existing.name}: ${changes.join(', ')}`);
      } else {
        const doc = Object.assign({ id: uid('d'), aliases: [], fte: 1, plan: { A: ['', '', '', '', ''], B: null } }, d);
        if (doc.alEntitlement === undefined || doc.alEntitlement === '') doc.alEntitlement = R.alProRata(doc).days;
        state.doctors.push(doc);
        log(state, by, 'Doctor added', `${doc.name} (${doc.grade})`);
      }
      return [];
    },
    // Put several doctors at one location for a block of dates (only the chosen weekdays, working days, while in post).
    bulkAssign(state, by, ids, from, to, days, code, reason) {
      if (!ids || !ids.length) throw new Error('Tick at least one doctor.');
      if (!R.isIso(from) || !R.isIso(to) || to < from) throw new Error('Choose a start date and an end date on or after it.');
      if (R.daysBetween(from, to) > 400) throw new Error('Choose up to about a year at a time.');
      if (!days || !days.length) throw new Error('Tick at least one weekday.');
      let n = 0;
      for (const id of ids) {
        const d = R.doctorById(state, id);
        if (!d) continue;
        for (const iso of R.dateRange(from, to)) {
          if (!R.isWorkday(state, iso) || !days.includes(R.dow(iso)) || !R.inPost(d, iso)) continue;
          state.overrides[id] = state.overrides[id] || {};
          if (code) state.overrides[id][iso] = { code: R.canonicalCode(code), reason: reason || 'Group change', by, at: nowIso() };
          else delete state.overrides[id][iso];
          n++;
        }
      }
      if (!n) throw new Error('None of those doctors are in post on the chosen days.');
      log(state, by, 'Group change', `${ids.map((id) => docName(state, id)).join(', ')}: ${code || '(usual location)'} on ${n} day${n === 1 ? '' : 's'}, ${R.fmtUk(from)} to ${R.fmtUk(to)}${reason ? ` (${reason})` : ''}`);
      return [];
    },
    // Apply a parsed master workbook: update doctors matched by name, add new ones, set consultants, add restrictions.
    applyMaster(state, by, parsed, fileName) {
      let added = 0, updated = 0;
      const byName = (n) => state.doctors.find((d) => d.name.toLowerCase() === String(n || '').trim().toLowerCase()) || null;
      const missing = [];
      for (const x of parsed.doctors || []) {
        const fields = { name: x.name, grade: x.grade, group: x.group, start: x.start || '', end: x.end || '', endProvisional: !!x.endProvisional, plan: { A: x.plan, B: null }, placeholder: !!x.placeholder };
        if (x.email) fields.email = x.email;
        if (x.alEntitlement !== undefined) fields.alEntitlement = x.alEntitlement;
        if (x.alExtra !== undefined) fields.alExtra = x.alExtra;
        if (x.slEntitlement !== undefined) fields.slEntitlement = x.slEntitlement;
        if (x.aliases && x.aliases.length) fields.aliases = x.aliases;
        if (x.replacesName) { const prev = byName(x.replacesName); if (prev && prev.name.toLowerCase() !== x.name.toLowerCase()) fields.slot = R.slotOf(prev); else if (!prev) missing.push(`${x.name}: "${x.replacesName}" not found, so a new column was used`); }
        const existing = byName(x.name);
        if (existing) {
          if (!x.plan.some(Boolean)) delete fields.plan; // no locations in the file: keep the doctor's usual week
          Object.assign(existing, fields);
          updated++;
        } else {
          const doc = Object.assign({ id: uid('d'), aliases: [], fte: 1, email: '', slEntitlement: x.group === 'mg' ? 10 : 15 }, fields);
          if (doc.alEntitlement === undefined) doc.alEntitlement = R.alProRata(doc).days;
          state.doctors.push(doc); added++;
        }
      }
      const weeks = Object.entries(parsed.consultants || {});
      weeks.forEach(([w, n]) => { state.settings.consultantOfWeek[w] = n; });
      let restr = 0;
      for (const r of parsed.restrictions || []) {
        if (state.settings.restricted.some((x) => x.start === r.start && (x.end || x.start) === r.end && x.level === r.level)) continue;
        state.settings.restricted.push(Object.assign({ id: uid('x') }, r)); restr++;
      }
      let shifts = 0;
      const oc = parsed.oncall;
      if (oc && oc.from && oc.to) {
        const entries = (oc.entries || []).map((e) => Object.assign({}, e, { doctorId: e.doctorId || (parsed.mapping || {})[e.name] || null })).filter((e) => e.doctorId);
        for (const id of Object.keys(state.oncall)) for (const iso of Object.keys(state.oncall[id])) if (iso >= oc.from && iso <= oc.to) delete state.oncall[id][iso];
        for (const e of entries) (state.oncall[e.doctorId] = state.oncall[e.doctorId] || {})[e.date] = e.code;
        for (const [alias, id] of Object.entries(parsed.mapping || {})) { const d = R.doctorById(state, id); if (id && d && !(d.aliases || []).includes(alias)) (d.aliases = d.aliases || []).push(alias); }
        shifts = entries.length;
      }
      // Rota tab: only the cells that were changed. Applied last so it wins over the On-calls tab.
      let cells = 0;
      for (const c of (parsed.rota && parsed.rota.changes) || []) {
        if (!R.doctorById(state, c.doctorId)) continue;
        state.overrides[c.doctorId] = state.overrides[c.doctorId] || {};
        state.oncall[c.doctorId] = state.oncall[c.doctorId] || {};
        if (c.kind === 'oncall') { delete state.overrides[c.doctorId][c.date]; state.oncall[c.doctorId][c.date] = c.code; }
        else { delete state.oncall[c.doctorId][c.date]; state.overrides[c.doctorId][c.date] = { code: c.code, reason: 'Rota tab of the data file', by, at: nowIso() }; }
        cells++;
      }
      log(state, by, 'Rota data file imported', `${fileName || 'file'}: ${parsed.rota ? `${cells} rota cells changed, ` : ''}${added} doctors added, ${updated} updated, ${weeks.length} consultant weeks, ${restr} leave restrictions${oc ? `, ${shifts} on-call shifts (${R.fmtUk(oc.from)} to ${R.fmtUk(oc.to)})` : ''}${missing.length ? '. ' + missing.join('; ') : ''}`);
      return [];
    },
    // Put back an earlier copy of the whole rota (from a backup).
    restoreState(state, by, copy, label) {
      if (!copy || copy.version !== 2 || !Array.isArray(copy.doctors)) throw new Error('That is not a rota backup.');
      const audit = state.audit;
      Object.keys(state).forEach((k) => delete state[k]);
      Object.assign(state, clone(copy));
      upgrade(state);
      state.audit = audit;
      log(state, by, 'Rota restored from a backup', label || '');
      return [];
    },
    copyDoctor(state, by, fromId, fields) {
      const src = R.doctorById(state, fromId);
      if (!src) throw new Error('Choose a doctor to copy.');
      if (!String(fields.name || '').trim()) throw new Error('Type the new doctor\'s name.');
      const doc = { id: uid('d'), name: fields.name.trim(), grade: fields.grade || src.grade, group: src.group, email: fields.email || '', start: fields.start || '', end: fields.end || '',
        fte: src.fte || 1, aliases: [], plan: clone(src.plan || { A: ['', '', '', '', ''], B: null }), edtHours: src.edtHours, slEntitlement: src.slEntitlement };
      doc.alEntitlement = R.alProRata(doc).days;
      state.doctors.push(doc);
      log(state, by, 'Doctor added (copied)', `${doc.name} copied from ${src.name}: grade, group and weekly plan`);
      return [];
    },
    removeDoctor(state, by, id) {
      const d = R.doctorById(state, id);
      state.doctors = state.doctors.filter((x) => x.id !== id);
      log(state, by, 'Doctor removed', d ? d.name : id);
      return [];
    },
    setPlanDay(state, by, doctorId, week, i, code) {
      const d = R.doctorById(state, doctorId);
      d.plan = d.plan || { A: ['', '', '', '', ''], B: null };
      if (week === 'B' && !d.plan.B) d.plan.B = d.plan.A.slice();
      d.plan[week][i] = code;
      log(state, by, 'Weekly plan changed', `${d.name}, week ${week} ${R.DOW_SHORT[i]}: ${code || '(blank)'}`);
      return [];
    },
    toggleWeekB(state, by, doctorId, on) {
      const d = R.doctorById(state, doctorId);
      d.plan.B = on ? d.plan.A.slice() : null;
      if (on) d.plan.anchor = d.plan.anchor || d.start;
      log(state, by, 'Weekly plan changed', `${d.name}: alternate week B ${on ? 'on' : 'off'}`);
      return [];
    },
    applyImport(state, by, kind, payload) {
      if (kind === 'oncall') {
        const { entries, from, to, fileName } = payload;
        for (const id of Object.keys(state.oncall)) for (const iso of Object.keys(state.oncall[id])) if (iso >= from && iso <= to) delete state.oncall[id][iso];
        for (const e of entries) (state.oncall[e.doctorId] = state.oncall[e.doctorId] || {})[e.date] = e.code;
        log(state, by, 'On-call rota imported', `${fileName || 'file'}: ${entries.length} shifts, ${R.fmtUk(from)} to ${R.fmtUk(to)}`);
      } else if (kind === 'plan') {
        for (const p of payload.plans) { const d = R.doctorById(state, p.doctorId); d.plan = { A: p.A, B: p.B, anchor: p.from || (d.plan && d.plan.anchor) || d.start }; }
        log(state, by, 'Weekly plans imported', `${payload.fileName || 'file'}: ${payload.plans.length} doctors`);
      } else if (kind === 'staff') {
        let added = 0, updated = 0;
        for (const s of payload.staff) {
          const existing = state.doctors.find((d) => d.name.toLowerCase() === s.name.toLowerCase());
          if (existing) { Object.assign(existing, s, { id: existing.id, plan: existing.plan }); updated++; } else {
            const doc = Object.assign({ id: uid('d'), fte: 1, plan: { A: ['', '', '', '', ''], B: null } }, s);
            if (!doc.alEntitlement) doc.alEntitlement = R.alProRata(doc).days;
            state.doctors.push(doc); added++;
          }
        }
        log(state, by, 'Staff list imported', `${payload.fileName || 'file'}: ${added} added, ${updated} updated`);
      }
      for (const [alias, doctorId] of Object.entries(payload.newAliases || {})) {
        const d = R.doctorById(state, doctorId);
        if (d && !(d.aliases || []).includes(alias)) (d.aliases = d.aliases || []).push(alias);
      }
      return [];
    },
    setSetting(state, by, path, value, what) {
      const keys = path.split('.');
      let o = state.settings;
      for (let i = 0; i < keys.length - 1; i++) o = o[keys[i]] = o[keys[i]] || {};
      o[keys[keys.length - 1]] = value;
      log(state, by, 'Settings changed', what || `${path}: ${JSON.stringify(value).slice(0, 120)}`);
      return [];
    },
    addRestricted(state, by, start, end, level, note) {
      if (!R.isIso(start) || (end && end < start)) throw new Error('Choose a start date and an end date on or after it.');
      state.settings.restricted.push({ id: uid('x'), start, end: end || start, level, note: note || '' });
      log(state, by, 'Leave restriction added', `${level}: ${R.fmtUk(start)} to ${R.fmtUk(end || start)}`);
      return [];
    },
    removeRestricted(state, by, id) {
      state.settings.restricted = state.settings.restricted.filter((r) => r.id !== id);
      log(state, by, 'Leave restriction removed', id);
      return [];
    },
    addChangeover(state, by, date, label) {
      if (!R.isIso(date) || !String(label || '').trim()) throw new Error('Give a date and a description for the changeover.');
      state.settings.changeovers.push({ id: uid('c'), date, label: label.trim() });
      log(state, by, 'Changeover added', `${R.fmtUk(date)}: ${label}`);
      return [];
    },
    removeChangeover(state, by, id) {
      state.settings.changeovers = state.settings.changeovers.filter((c) => c.id !== id);
      log(state, by, 'Changeover removed', id);
      return [];
    },
    startFresh(state, by) {
      const keep = { authorisers: state.settings.authorisers, authoriserEmails: state.settings.authoriserEmails, rulesText: state.settings.rulesText, rules: state.settings.rules };
      const fresh = emptyState();
      Object.assign(fresh.settings, keep);
      Object.keys(state).forEach((k) => delete state[k]);
      Object.assign(state, fresh);
      log(state, by, 'Started fresh', 'Demo data removed. Rules, authorisers and alert emails kept.');
      return [];
    }
  };

  // ---------- the store used by the screens ----------
  function createStore(opts) {
    opts = opts || {};
    const listeners = [];
    const emit = () => listeners.forEach((fn) => fn());
    const store = { mode: 'local', state: null, version: 0, me: null, hasPassword: true, emailOn: false, error: null };
    const storage = opts.storage !== undefined ? opts.storage : (() => {
      try { const ls = root.localStorage; return ls ? { get: (k) => { try { return ls.getItem(k); } catch (e) { return null; } }, set: (k, v) => { try { ls.setItem(k, v); } catch (e) { /* blocked */ } } } : null; } catch (e) { return null; }
    })();

    async function api(method, path, body) {
      const res = await fetch(path, { method, headers: body ? { 'Content-Type': 'application/json' } : {}, body: body ? JSON.stringify(body) : undefined, credentials: 'same-origin', cache: 'no-store' });
      let data = null;
      try { data = await res.json(); } catch (e) { /* not json */ }
      if (!res.ok) { const err = new Error((data && data.error) || `The server replied ${res.status}.`); err.status = res.status; err.data = data; throw err; }
      return data;
    }
    function loadLocal() {
      let s = null;
      if (storage) { try { const raw = storage.get(LOCAL_KEY); s = raw ? JSON.parse(raw) : null; } catch (e) { s = null; } }
      store.state = s && s.version === 2 ? upgrade(s) : demoState(opts.today);
      const who = storage && storage.get(LOCAL_KEY + '-me');
      store.me = who || null;
    }
    const saveLocal = () => { if (storage) storage.set(LOCAL_KEY, JSON.stringify(store.state)); };

    Object.assign(store, {
      subscribe(fn) { listeners.push(fn); },
      isAuth() { return !!store.me; },
      async init() {
        if (opts.forceLocal) { loadLocal(); emit(); return; }
        try {
          const ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
          const t = ctrl ? setTimeout(() => ctrl.abort(), 6000) : null;
          const res = await fetch('/api/state', { cache: 'no-store', credentials: 'same-origin', signal: ctrl ? ctrl.signal : undefined });
          if (t) clearTimeout(t);
          const data = await res.json();
          if (!res.ok || !data || data.mode !== 'server') throw new Error('no server');
          store.mode = 'remote';
          applyServer(data);
        } catch (e) { store.mode = 'local'; loadLocal(); }
        emit();
      },
      async refresh(force) {
        if (store.mode !== 'remote') return;
        const before = store.version + '|' + store.me;
        try { applyServer(await api('GET', '/api/state')); store.error = null; } catch (e) { store.error = e.message; }
        if (force || before !== store.version + '|' + store.me) emit();
      },
      async login(name, password) {
        if (store.mode === 'local') { store.me = name; if (storage) storage.set(LOCAL_KEY + '-me', name); emit(); return; }
        await api('POST', '/api/login', { name, password });
        await store.refresh(true);
      },
      async logout() {
        if (store.mode === 'local') { store.me = null; if (storage) storage.set(LOCAL_KEY + '-me', ''); emit(); return; }
        await api('POST', '/api/logout');
        store.me = null;
        await store.refresh(true);
      },
      async setupPassword(password, setupCode, names) {
        await api('POST', '/api/setup', { password, setupCode, names });
        await store.refresh(true);
      },
      async changePassword(current, next) {
        if (store.mode === 'local') throw new Error('Passwords only apply on the hosted version.');
        await api('POST', '/api/password', { current, next });
        store.me = null;
        await store.refresh(true);
      },
      async submitLeave(payload) {
        if (store.mode === 'local') { const out = submitLeave(store.state, payload); saveLocal(); emit(); return { request: out.request, check: out.check, emailed: false }; }
        const out = await api('POST', '/api/requests', payload);
        await store.refresh();
        return out;
      },
      async withdraw(id, email) {
        if (store.mode === 'local') { withdrawLeave(store.state, id, email); saveLocal(); emit(); return; }
        await api('POST', `/api/requests/${encodeURIComponent(id)}/withdraw`, { email });
        await store.refresh();
      },
      async listBackups() {
        if (store.mode === 'local') return { backups: [], local: true };
        return api('GET', '/api/backups');
      },
      async getBackup(id) { return api('GET', `/api/backups/${encodeURIComponent(id)}`); },
      async getCsvBackup(id) {
        const res = await fetch(`/api/csv-backups/${encodeURIComponent(id)}`, { credentials: 'same-origin', cache: 'no-store' });
        if (!res.ok) { let m = `The server replied ${res.status}.`; try { m = (await res.json()).error || m; } catch (e) { /* not json */ } throw new Error(m); }
        return new Uint8Array(await res.arrayBuffer());
      },
      async myRequests(doctorId, email) {
        if (store.mode === 'local') return myRequests(store.state, doctorId, email);
        return api('POST', '/api/my-requests', { doctorId, email });
      },
      // Authoriser change: run fn on the state, then save. On a clash with another authoriser, reload and redo once.
      async act(fn) {
        if (!store.me) throw new Error('Sign in as an authoriser to make changes.');
        if (store.mode === 'local') { const emails = fn(store.state, store.me) || []; saveLocal(); emit(); return { emails, sent: 0 }; }
        for (let attempt = 0; attempt < 2; attempt++) {
          const draft = clone(store.state);
          const emails = fn(draft, store.me) || [];
          try {
            const out = await api('PUT', '/api/state', { baseVersion: store.version, state: draft, emails });
            store.state = draft; store.version = out.version; emit();
            return { emails, sent: out.sent || 0 };
          } catch (e) {
            if (e.status === 409 && attempt === 0) { applyServer(e.data); continue; }
            if (e.status === 401) { store.me = null; emit(); }
            throw e;
          }
        }
        return { emails: [], sent: 0 };
      },
      resetLocalDemo() { store.state = demoState(opts.today); saveLocal(); emit(); }
    });
    function applyServer(data) {
      store.state = upgrade(data.state); store.version = data.version; store.me = data.me || null;
      store.hasPassword = data.hasPassword !== false; store.emailOn = !!data.emailOn; store.setupCodeRequired = !!data.setupCodeRequired;
    }
    return store;
  }

  function restrictionOn(state, iso) {
    return ((state.settings && state.settings.restricted) || []).find((r) => iso >= r.start && iso <= (r.end || r.start)) || null;
  }

  const api = { createStore, demoState, emptyState, upgrade, submitLeave, withdrawLeave, myRequests, publicView, ops, restrictionOn, DEFAULT_RULES_TEXT, LOCAL_KEY };
  root.RotaStore = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
