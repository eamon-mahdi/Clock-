/* App state, demo data, saving, and every change the app can make.
 * Each change checks the signed-in role and writes an audit entry.
 * In this prototype the state is saved in the browser; the real version keeps it on a server
 * that enforces the same permissions. */
(function (root) {
  'use strict';
  const R = root.RotaRules || (typeof require !== 'undefined' ? require('./rules.js') : null);
  const STORAGE_KEY = 'haem-rota-prototype-v1';
  const USER_KEY = 'haem-rota-prototype-user';

  const uid = (p) => p + '-' + Math.random().toString(36).slice(2, 9);
  const nowIso = () => new Date().toISOString();

  // ---------- demo data (all names are fictional) ----------
  function demoState(todayIso) {
    const today = todayIso || R.toIso(new Date());
    const mon = R.mondayOf(today);
    const plan = (a, b) => ({ A: a.split('|'), B: b ? b.split('|') : null, anchor: mon });
    const doctors = [
      { id: 'r1', name: 'E. Ward', grade: 'F1', group: 'res', start: R.addDays(mon, -68), end: R.addDays(mon, 50), alEntitlement: 9, slEntitlement: 10, plan: plan('GUH|GUH|GUH|GUH|GUH') },
      { id: 'r2', name: 'P. Nair', grade: 'F2', group: 'res', start: R.addDays(mon, -68), end: R.addDays(mon, 50), alEntitlement: 9, slEntitlement: 10, plan: plan('RGH|NHH|NHH|NHH|NHH') },
      { id: 'r3', name: 'J. Reid', grade: 'F2', group: 'res', start: R.addDays(mon, -68), end: R.addDays(mon, -10), alEntitlement: 9, slEntitlement: 10, plan: plan('RGH|RGH|RGH|RGH|RGH'), note: 'Left the rotation early' },
      { id: 'r4', name: 'T. Hughes', grade: 'F2 (on calls NHH)', group: 'res', start: R.addDays(mon, -68), end: R.addDays(mon, 50), alEntitlement: 9, slEntitlement: 10, plan: plan('GUH|GUH|RGH|GUH|GUH') },
      { id: 'r5', name: 'S. Malik', grade: 'IMT2', group: 'res', start: R.addDays(mon, -68), end: R.addDays(mon, 113), alEntitlement: 14, slEntitlement: 15, plan: plan('GUH|RGH|RGH|RGH|RGH') },
      { id: 'r6', name: 'B. Carter', grade: 'IMT3', group: 'res', start: R.addDays(mon, -68), end: R.addDays(mon, 113), alEntitlement: 14, slEntitlement: 15, plan: plan('GUH|GUH|GUH|RGH|GUH') },
      { id: 'r7', name: 'L. Osei', grade: 'JCF', group: 'res', start: R.addDays(mon, -68), end: R.addDays(mon, 295), alEntitlement: 27, slEntitlement: 30, plan: plan('RGH|RGH|RGH|RGH|RGH') },
      { id: 'm1', name: 'Omar', grade: 'Middle grade', group: 'mg', start: '2024-01-01', end: '', alEntitlement: 32, slEntitlement: 10, plan: plan('GUH|GUH|GUH|GUH|GUH') },
      { id: 'm2', name: 'Hannah', grade: 'Middle grade', group: 'mg', start: '2024-01-01', end: '', alEntitlement: 32, slEntitlement: 10, plan: plan('RGH - DAY UNIT|RGH - DAY UNIT|RGH - DAY UNIT|RGH - DAY UNIT|RGH - DAY UNIT') },
      { id: 'm3', name: 'Daniel', grade: 'Middle grade', group: 'mg', start: '2024-01-01', end: '', alEntitlement: 32, slEntitlement: 10, plan: plan('GUH LAB|NHH - DAY UNIT|NHH - DAY UNIT|NHH - DAY UNIT|NHH - DAY UNIT') },
      { id: 'm4', name: 'Zara', grade: 'Middle grade', group: 'mg', start: '2024-01-01', end: '', alEntitlement: 32, slEntitlement: 10, plan: plan('RGH MDCU|RGH|RGH|RGH|RGH') },
      { id: 'm5', name: 'Meera', grade: 'Middle grade', group: 'mg', start: '2024-01-01', end: '', alEntitlement: 27, slEntitlement: 10, plan: plan('GUH LAB|GUH LAB|GUH|GUH|OFF', 'GUH LAB|GUH|OFF|GUH|GUH') },
      { id: 'm6', name: 'Yusuf', grade: 'Middle grade', group: 'mg', start: '2024-01-01', end: '', alEntitlement: 32, slEntitlement: 10, plan: plan('RGH MDCU|RGH|RGH|OFF|RGH') }
    ].map((d) => Object.assign({ email: d.name.toLowerCase().replace(/[^a-z]+/g, '.').replace(/^\.|\.$/g, '') + '@example.nhs.uk', aliases: [] }, d));

    // On-calls as if imported from the trust medical rota (residents only)
    const oncall = {};
    const residents = doctors.filter((d) => d.group === 'res' && d.id !== 'r3');
    for (let w = -1; w < 8; w++) {
      const wk = R.addDays(mon, w * 7);
      residents.forEach((d, i) => {
        const set = (offset, code) => { const iso = R.addDays(wk, offset); if (R.inPost(d, iso)) (oncall[d.id] = oncall[d.id] || {})[iso] = code; };
        const k = (w + i * 2 + 12) % 6;
        if (k === 0) { [0, 1, 2, 3].forEach((o) => set(o, 'NIGHTS')); set(4, 'POST NIGHTS'); }
        if (k === 3) { set(4, 'OFF PRE ON-CALL'); set(5, 'LONG DAY'); set(6, 'LONG DAY'); }
        if (k === 1) { set(1, 'TWILIGHT'); set(2, 'TWILIGHT'); }
        if (k === 4 && i % 2 === 0) { set(5, 'NIGHTS'); set(6, 'NIGHTS'); }
      });
    }

    const users = [
      { id: 'c1', name: 'J. Hale', role: 'coordinator' },
      { id: 'c2', name: 'S. Price', role: 'coordinator' },
      { id: 'c3', name: 'R. Dunn', role: 'coordinator' },
      { id: 'v1', name: 'Ward / switchboard', role: 'viewer' }
    ].concat(doctors.map((d) => ({ id: 'u-' + d.id, name: d.name, role: 'doctor', doctorId: d.id })));

    const consultants = ['Dr Fraser', 'Dr Lin', 'Dr Okoro', 'Dr Grant', 'Dr Shah'];
    const consultantOfWeek = {};
    for (let w = -2; w < 12; w++) consultantOfWeek[R.addDays(mon, w * 7)] = consultants[(w + 10) % consultants.length];

    const state = {
      version: 1,
      createdFor: today,
      settings: {
        bankHolidays: R.BANK_HOLIDAYS.slice(),
        restricted: [
          { id: uid('x'), start: R.addDays(mon, 7), end: R.addDays(mon, 8), level: 'NO AL', note: 'Short-staffed: two doctors on induction' },
          { id: uid('x'), start: R.addDays(mon, 21), end: R.addDays(mon, 25), level: 'AVOID', note: 'Leave discouraged: rotation changeover prep' }
        ],
        consultantOfWeek
      },
      doctors, users, oncall, overrides: {}, requests: [], audit: [], notifications: []
    };

    const add = (doctorId, type, start, end, part, status, note, extra) => {
      const r = Object.assign({
        id: uid('lr'), doctorId, type, start, end, part, note: note || '', status, cover: [],
        submittedAt: R.addDays(today, -6) + 'T09:12:00Z', days: R.leaveDays(state, start, end, part)
      }, extra || {});
      if (status !== 'pending') Object.assign(r, { decidedBy: 'J. Hale', decidedAt: R.addDays(today, -5) + 'T14:30:00Z' });
      state.requests.push(r);
      return r;
    };
    add('m2', 'AL', R.addDays(mon, 9), R.addDays(mon, 11), 'FULL', 'authorised', 'Half term');
    add('r6', 'AL', R.addDays(mon, 17), R.addDays(mon, 17), 'PM', 'authorised', 'Appointment');
    add('r1', 'AL', R.addDays(mon, 43), R.addDays(mon, 46), 'FULL', 'authorised', 'Holiday');
    add('m4', 'AL', R.addDays(mon, 42), R.addDays(mon, 46), 'FULL', 'authorised', 'Holiday');
    add('m6', 'SICK', R.addDays(mon, -3), R.addDays(mon, -3), 'FULL', 'authorised', '', { decidedBy: 'Self-reported' });
    const meera = add('m5', 'AL', R.addDays(mon, 45), R.addDays(mon, 45), 'FULL', 'pending', "Sister's wedding");
    const sugg = R.suggestCover(state, 'm5', meera.start, meera.end, 'FULL');
    meera.cover = sugg.filter((s) => s.needed && s.options.length).map((s) => ({ date: s.date, doctorId: s.options[0].doctorId, to: s.options[0].to, text: s.options[0].text }));
    add('r7', 'SL', R.addDays(mon, 52), R.addDays(mon, 52), 'FULL', 'pending', 'Regional teaching course');
    add('r4', 'TOIL', R.addDays(mon, 8), R.addDays(mon, 8), 'FULL', 'pending', 'In lieu of study day worked on a weekend');

    state.audit.push({ at: nowIso(), by: 'System', action: 'Demo data loaded', detail: 'Fictional doctors, weekly plans, imported on-calls and leave requests' });
    state.notifications.push({ id: uid('n'), to: 'coord', at: nowIso(), text: '3 leave requests are waiting for authorisation.', read: false });
    return state;
  }

  // ---------- store ----------
  function createStore(opts) {
    opts = opts || {};
    const storage = opts.storage === undefined ? safeStorage() : opts.storage;
    let state = load() || demoState(opts.today);
    let userId = (storage && storage.get(USER_KEY)) || 'c1';
    const listeners = [];

    function safeStorage() {
      try {
        const ls = root.localStorage;
        if (!ls) return null;
        return { get: (k) => { try { return ls.getItem(k); } catch (e) { return null; } }, set: (k, v) => { try { ls.setItem(k, v); } catch (e) { /* storage blocked */ } } };
      } catch (e) { return null; }
    }
    function load() {
      if (!storage) return null;
      try { const raw = storage.get(STORAGE_KEY); return raw ? JSON.parse(raw) : null; } catch (e) { return null; }
    }
    function save() { if (storage) storage.set(STORAGE_KEY, JSON.stringify(state)); }
    function emit() { save(); listeners.forEach((fn) => fn()); }

    const user = () => state.users.find((u) => u.id === userId) || state.users[0];
    const isCoord = () => user().role === 'coordinator';
    function requireCoord(what) { if (!isCoord()) throw new Error(`Only rota coordinators can ${what}.`); }
    function log(action, detail) { state.audit.unshift({ at: nowIso(), by: user().name, action, detail }); }
    function notify(to, text) { state.notifications.unshift({ id: uid('n'), to, at: nowIso(), text, read: false }); }
    const docName = (id) => (R.doctorById(state, id) || { name: id }).name;

    const api = {
      get state() { return state; },
      user, isCoord,
      subscribe(fn) { listeners.push(fn); },
      setUser(id) { userId = id; if (storage) storage.set(USER_KEY, id); listeners.forEach((fn) => fn()); },
      resetDemo() { state = demoState(opts.today); emit(); },
      replaceState(s) { state = s; emit(); },

      setOverride(doctorId, iso, code, reason) {
        requireCoord('edit the roster');
        state.overrides[doctorId] = state.overrides[doctorId] || {};
        const before = R.resolveCell(state, R.doctorById(state, doctorId), iso).code || '(blank)';
        if (code) state.overrides[doctorId][iso] = { code: R.canonicalCode(code), reason: reason || '', by: user().name, at: nowIso() };
        else delete state.overrides[doctorId][iso];
        const after = R.resolveCell(state, R.doctorById(state, doctorId), iso).code || '(blank)';
        log('Roster cell changed', `${docName(doctorId)}, ${R.fmtShort(iso)}: ${before} → ${after}${reason ? ` (${reason})` : ''}`);
        emit();
      },

      submitRequest(req) {
        const u = user();
        if (u.role === 'viewer') throw new Error('Viewers cannot request leave.');
        if (u.role === 'doctor' && req.doctorId !== u.doctorId) throw new Error('You can only request leave for yourself.');
        if (!req.start || !req.end || req.end < req.start) throw new Error('Choose a start date and an end date on or after it.');
        const r = {
          id: uid('lr'), doctorId: req.doctorId, type: req.type, start: req.start, end: req.end, part: req.part || 'FULL',
          note: req.note || '', cover: req.cover || [], submittedAt: nowIso(), submittedBy: u.name,
          days: R.leaveDays(state, req.start, req.end, req.part || 'FULL'), status: 'pending'
        };
        const label = R.LEAVE_TYPES[r.type].label;
        if (r.type === 'SICK') {
          Object.assign(r, { status: 'authorised', decidedBy: 'Self-reported', decidedAt: nowIso() });
          notify('coord', `${docName(r.doctorId)} reported sick: ${R.fmtShort(r.start)}${r.end !== r.start ? ' – ' + R.fmtShort(r.end) : ''}.`);
          log('Sickness reported', `${docName(r.doctorId)}: ${R.fmtUk(r.start)} to ${R.fmtUk(r.end)}`);
        } else {
          notify('coord', `New ${label.toLowerCase()} request from ${docName(r.doctorId)}: ${R.fmtShort(r.start)}${r.end !== r.start ? ' – ' + R.fmtShort(r.end) : ''}.`);
          log('Leave requested', `${docName(r.doctorId)}: ${label}, ${R.fmtUk(r.start)} to ${R.fmtUk(r.end)} (${r.days} day${r.days === 1 ? '' : 's'})`);
        }
        state.requests.push(r);
        emit();
        return r;
      },

      setRequestCover(id, date, choice) {
        requireCoord('change the cover for a request');
        const r = state.requests.find((x) => x.id === id);
        if (!r || r.status !== 'pending') throw new Error('This request has already been decided.');
        r.cover = (r.cover || []).filter((c) => c.date !== date);
        if (choice) r.cover.push(Object.assign({ date }, choice));
        log('Cover proposal changed', `${docName(r.doctorId)}'s leave, ${R.fmtShort(date)}: ${choice ? choice.text : 'no cover move'}`);
        emit();
      },

      cancelRequest(id) {
        const r = state.requests.find((x) => x.id === id);
        if (!r) return;
        const u = user();
        if (!(isCoord() || (u.doctorId === r.doctorId && r.status === 'pending'))) throw new Error('You can only withdraw your own pending requests.');
        r.status = 'cancelled';
        log('Leave request withdrawn', `${docName(r.doctorId)}: ${R.fmtUk(r.start)} to ${R.fmtUk(r.end)}`);
        emit();
      },

      decideRequest(id, decision, reason) {
        requireCoord('authorise or decline leave');
        const r = state.requests.find((x) => x.id === id);
        if (!r || r.status !== 'pending') throw new Error('This request has already been decided.');
        if (decision === 'declined' && !String(reason || '').trim()) throw new Error('Give a reason for declining. It is sent to the doctor.');
        r.status = decision; r.decidedBy = user().name; r.decidedAt = nowIso(); r.reason = reason || '';
        const label = R.LEAVE_TYPES[r.type].label;
        if (decision === 'authorised') {
          for (const c of r.cover || []) {
            state.overrides[c.doctorId] = state.overrides[c.doctorId] || {};
            state.overrides[c.doctorId][c.date] = { code: c.to, reason: `Covering ${docName(r.doctorId)}'s leave`, by: user().name, at: nowIso() };
            notify('u-' + c.doctorId, `You are covering ${c.to} on ${R.fmtShort(c.date)} for ${docName(r.doctorId)}.`);
          }
        }
        notify('u-' + r.doctorId, `Your ${label.toLowerCase()} (${R.fmtShort(r.start)}${r.end !== r.start ? ' – ' + R.fmtShort(r.end) : ''}) was ${decision} by ${user().name}${reason ? ': ' + reason : '.'}`);
        log(decision === 'authorised' ? 'Leave authorised' : 'Leave declined', `${docName(r.doctorId)}: ${label}, ${R.fmtUk(r.start)} to ${R.fmtUk(r.end)}${(r.cover || []).length && decision === 'authorised' ? `; cover: ${r.cover.map((c) => `${docName(c.doctorId)} → ${c.to} (${R.fmtUk(c.date)})`).join(', ')}` : ''}${reason ? `; reason: ${reason}` : ''}`);
        emit();
      },

      saveDoctor(d) {
        requireCoord('change the staff list');
        const existing = R.doctorById(state, d.id);
        if (existing) {
          const changes = Object.keys(d).filter((k) => k !== 'plan' && JSON.stringify(existing[k]) !== JSON.stringify(d[k]));
          Object.assign(existing, d);
          const u = state.users.find((x) => x.doctorId === existing.id);
          if (u) u.name = existing.name;
          if (changes.length) log('Doctor updated', `${existing.name}: ${changes.join(', ')}`);
        } else {
          const doc = Object.assign({ id: uid('d'), aliases: [], plan: { A: ['', '', '', '', ''], B: null } }, d);
          state.doctors.push(doc);
          state.users.push({ id: 'u-' + doc.id, name: doc.name, role: 'doctor', doctorId: doc.id });
          log('Doctor added', `${doc.name} (${doc.grade})`);
        }
        emit();
      },

      removeDoctor(id) {
        requireCoord('change the staff list');
        const d = R.doctorById(state, id);
        state.doctors = state.doctors.filter((x) => x.id !== id);
        state.users = state.users.filter((u) => u.doctorId !== id);
        log('Doctor removed', d ? d.name : id);
        emit();
      },

      setPlanDay(doctorId, week, dayIndex, code) {
        requireCoord('change weekly plans');
        const d = R.doctorById(state, doctorId);
        d.plan = d.plan || { A: ['', '', '', '', ''], B: null };
        if (week === 'B' && !d.plan.B) d.plan.B = d.plan.A.slice();
        d.plan[week][dayIndex] = code;
        log('Weekly plan changed', `${d.name}, week ${week} ${R.DOW_SHORT[dayIndex]}: ${code || '(blank)'}`);
        emit();
      },

      toggleWeekB(doctorId, on) {
        requireCoord('change weekly plans');
        const d = R.doctorById(state, doctorId);
        d.plan.B = on ? d.plan.A.slice() : null;
        if (on) d.plan.anchor = d.plan.anchor || d.start;
        log('Weekly plan changed', `${d.name}: alternate week B ${on ? 'turned on' : 'turned off'}`);
        emit();
      },

      applyImport(kind, payload) {
        requireCoord('import files');
        if (kind === 'oncall') {
          const { entries, from, to, fileName } = payload; // entries: [{doctorId, date, code}]
          for (const d of state.doctors) {
            const m = state.oncall[d.id];
            if (!m) continue;
            for (const iso of Object.keys(m)) if (iso >= from && iso <= to) delete m[iso];
          }
          for (const e of entries) (state.oncall[e.doctorId] = state.oncall[e.doctorId] || {})[e.date] = e.code;
          log('On-call rota imported', `${fileName || 'file'}: ${entries.length} shifts, ${R.fmtUk(from)} to ${R.fmtUk(to)} (replaced earlier imports for those dates)`);
        } else if (kind === 'plan') {
          for (const p of payload.plans) {
            const d = R.doctorById(state, p.doctorId);
            d.plan = { A: p.A, B: p.B, anchor: p.from || d.plan && d.plan.anchor || d.start };
          }
          log('Weekly plans imported', `${payload.fileName || 'file'}: ${payload.plans.length} doctors`);
        } else if (kind === 'staff') {
          let added = 0, updated = 0;
          for (const s of payload.staff) {
            const existing = state.doctors.find((d) => d.name.toLowerCase() === s.name.toLowerCase());
            if (existing) { Object.assign(existing, s, { id: existing.id, plan: existing.plan }); updated++; } else {
              const doc = Object.assign({ id: uid('d'), plan: { A: ['', '', '', '', ''], B: null } }, s);
              state.doctors.push(doc);
              state.users.push({ id: 'u-' + doc.id, name: doc.name, role: 'doctor', doctorId: doc.id });
              added++;
            }
          }
          log('Staff list imported', `${payload.fileName || 'file'}: ${added} added, ${updated} updated`);
        }
        for (const [alias, doctorId] of Object.entries(payload.newAliases || {})) {
          const d = R.doctorById(state, doctorId);
          if (d && !d.aliases.includes(alias)) d.aliases.push(alias);
        }
        emit();
      },

      addRestricted(start, end, level, note) {
        requireCoord('set leave restrictions');
        if (!start || (end && end < start)) throw new Error('Choose a start date and an end date on or after it.');
        state.settings.restricted.push({ id: uid('x'), start, end: end || start, level, note: note || '' });
        log('Leave restriction added', `${level}: ${R.fmtUk(start)} to ${R.fmtUk(end || start)}${note ? ' (' + note + ')' : ''}`);
        emit();
      },
      removeRestricted(id) {
        requireCoord('set leave restrictions');
        const x = state.settings.restricted.find((r) => r.id === id);
        state.settings.restricted = state.settings.restricted.filter((r) => r.id !== id);
        if (x) log('Leave restriction removed', `${x.level}: ${R.fmtUk(x.start)} to ${R.fmtUk(x.end)}`);
        emit();
      },
      setConsultant(weekIso, name) {
        requireCoord('set the consultant of the week');
        state.settings.consultantOfWeek[weekIso] = name;
        log('Consultant of the week set', `Week of ${R.fmtUk(weekIso)}: ${name || '(none)'}`);
        emit();
      },
      markNotificationsRead() {
        const u = user();
        state.notifications.forEach((n) => { if (n.to === u.id || (n.to === 'coord' && u.role === 'coordinator')) n.read = true; });
        emit();
      }
    };
    return api;
  }

  // Restriction ranges -> per-date lookup
  function restrictionOn(state, iso) {
    return (state.settings.restricted || []).find((r) => iso >= r.start && iso <= (r.end || r.start)) || null;
  }

  const api = { createStore, demoState, restrictionOn, STORAGE_KEY };
  root.RotaStore = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
