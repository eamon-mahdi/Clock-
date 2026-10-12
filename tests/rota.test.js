// Run with: npm test
const test = require('node:test');
const assert = require('node:assert/strict');
global.RotaRules = require('../app/js/rules.js');
global.RotaXlsx = require('../app/js/xlsx.js');
global.RotaStore = require('../app/js/store.js');
const R = global.RotaRules;
const X = global.RotaXlsx;
const S = global.RotaStore;
const E = require('../app/js/excel.js');

const TODAY = '2026-10-11'; // a Sunday; demo week starts Mon 12 Oct
const fresh = () => ({ state: S.demoState(TODAY) });

function blankState(doctors) {
  return {
    version: 2, settings: { bankHolidays: R.BANK_HOLIDAYS, restricted: [], consultantOfWeek: {}, rules: Object.assign({}, R.DEFAULT_RULES), rotation: {}, changeovers: [] },
    doctors, oncall: {}, overrides: {}, requests: [], edt: [], audit: [], notifications: []
  };
}
const doc = (id, group, A, extra) => Object.assign({ id, name: id, grade: group === 'mg' ? 'MG' : 'F1', group, start: '2026-01-01', end: '', plan: { A: A.split('|'), B: null }, aliases: [] }, extra || {});

test('dates: day of week and ranges', () => {
  assert.equal(R.dow('2026-10-12'), 0); // Monday
  assert.equal(R.dow('2026-10-18'), 6); // Sunday
  assert.equal(R.mondayOf('2026-10-15'), '2026-10-12');
  assert.deepEqual(R.dateRange('2026-12-30', '2027-01-02'), ['2026-12-30', '2026-12-31', '2027-01-01', '2027-01-02']);
});

test('leave days skip weekends and bank holidays; half days count 0.5', () => {
  const st = blankState([]);
  assert.equal(R.leaveDays(st, '2026-12-21', '2026-12-29', 'FULL'), 5); // 25 and 28 Dec are bank holidays
  assert.equal(R.leaveDays(st, '2026-11-26', '2026-11-26', 'PM'), 0.5);
  assert.equal(R.leaveDays(st, '2026-11-28', '2026-11-29', 'FULL'), 0);
});

test('day-unit cover: all RGH codes count, LAB does not, resident-only is an exception', () => {
  const st = blankState([
    doc('res1', 'res', 'GUH|GUH|GUH|GUH|GUH'), doc('res2', 'res', 'RGH|RGH|RGH|RGH|RGH'), doc('res3', 'res', 'NHH|NHH|NHH|NHH|NHH'),
    doc('mg1', 'mg', 'GUH|GUH|GUH|GUH|GUH'), doc('mg2', 'mg', 'RGH MDCU|RGH MDCU|RGH MDCU|RGH MDCU|RGH MDCU'), doc('mg3', 'mg', 'GUH LAB|GUH LAB|GUH LAB|GUH LAB|GUH LAB')
  ]);
  st.settings.rules.minSenRGH = 1;
  const tue = R.coverFor(st, '2026-10-13');
  const by = Object.fromEntries(tue.rules.map((r) => [r.rule.id, r.status]));
  assert.equal(by['mg-RGH'], 'ok'); // RGH MDCU counts for the RGH day unit
  assert.equal(by['mg-NHH'], 'na'); // a senior at the day units is not compulsory by default
  st.settings.rules.minSenNHH = 1;
  assert.equal(R.coverFor(st, '2026-10-13').rules.find((r) => r.rule.id === 'mg-NHH').status, 'exception'); // if required: only a resident at NHH
  assert.equal(by['mg-GUH'], 'ok');
  const mon = R.coverFor(st, '2026-10-12');
  assert.equal(mon.rules.find((r) => r.rule.id === 'mg-NHH').status, 'na'); // NHH day unit closed Mondays
  assert.equal(mon.rules.find((r) => r.rule.id === 'res-NHH').status, 'na');
  st.doctors[3].plan.A = 'GUH LAB|GUH LAB|GUH LAB|GUH LAB|GUH LAB'.split('|');
  assert.equal(R.coverFor(st, '2026-10-13').rules.find((r) => r.rule.id === 'mg-GUH').status, 'gap'); // lab is not ward cover
  assert.equal(R.coverFor(st, '2026-10-17'), null); // weekend
  assert.equal(R.coverFor(st, '2026-12-25'), null); // bank holiday
});

test('cell precedence: not in post > authorised leave > override > on-call > plan', () => {
  const d = doc('a', 'res', 'GUH|GUH|GUH|GUH|GUH', { end: '2026-10-30' });
  const st = blankState([d]);
  assert.equal(R.resolveCell(st, d, '2026-10-13').code, 'GUH');
  st.oncall.a = { '2026-10-13': 'NIGHTS' };
  assert.equal(R.resolveCell(st, d, '2026-10-13').code, 'NIGHTS');
  st.overrides.a = { '2026-10-13': { code: 'RGH' } };
  assert.equal(R.resolveCell(st, d, '2026-10-13').code, 'RGH');
  st.requests.push({ doctorId: 'a', type: 'AL', start: '2026-10-13', end: '2026-10-13', part: 'FULL', status: 'pending' });
  const pend = R.resolveCell(st, d, '2026-10-13');
  assert.equal(pend.code, 'RGH'); // pending leave does not reach the roster
  assert.ok(pend.pending);
  st.requests[0].status = 'authorised';
  assert.equal(R.resolveCell(st, d, '2026-10-13').code, 'ANNUAL LEAVE');
  assert.equal(R.resolveCell(st, d, '2026-11-02').code, 'NOT IN POST');
});

test('pending leave shows the day as at risk', () => {
  const st = blankState([doc('mg1', 'mg', 'GUH|GUH|GUH|GUH|GUH'), doc('mg2', 'mg', 'RGH|RGH|RGH|RGH|RGH')]);
  st.requests.push({ doctorId: 'mg1', type: 'AL', start: '2026-10-14', end: '2026-10-14', part: 'FULL', status: 'pending' });
  assert.equal(R.coverWithRisk(st, '2026-10-14').rules.find((r) => r.rule.id === 'mg-GUH').status, 'risk');
});


test('half-day leave keeps the doctor counted at their site for the other session', () => {
  const d = doc('a', 'res', 'RGH|RGH|RGH|RGH|RGH');
  const st = blankState([d]);
  st.requests.push({ doctorId: 'a', type: 'AL', start: '2026-10-15', end: '2026-10-15', part: 'PM', status: 'authorised' });
  const c = R.resolveCell(st, d, '2026-10-15');
  assert.equal(c.code, 'RGH AM; PM ANNUAL LEAVE');
  assert.equal(c.am, 'RGH');
  assert.equal(c.pm, '');
  // the afternoon has no resident at RGH, so the day is short
  assert.equal(R.coverFor(st, '2026-10-15').rules.find((r) => r.rule.id === 'res-RGH').status, 'gap');
});

test('cover suggestions never leave a gap and prefer clean moves', () => {
  const st = blankState([
    doc('mg1', 'mg', 'GUH|GUH|GUH|GUH|GUH'), doc('mg2', 'mg', 'RGH|RGH|RGH|RGH|RGH'), doc('mg3', 'mg', 'RGH MDCU|RGH MDCU|RGH MDCU|RGH MDCU|RGH MDCU'),
    doc('mg4', 'mg', 'OFF|OFF|OFF|OFF|OFF'), doc('mg5', 'mg', 'NHH|NHH|NHH|NHH|NHH'), doc('res1', 'res', 'NHH|NHH|NHH|NHH|NHH')
  ]);
  const s = R.suggestCover(st, 'mg1', '2026-10-14', '2026-10-14', 'FULL');
  assert.equal(s.length, 1);
  assert.ok(s[0].needed);
  assert.deepEqual(s[0].broken, ['Senior doctor on GUH ward']);
  assert.ok(['mg2', 'mg3'].includes(s[0].options[0].doctorId), 'a spare RGH senior moves first');
  assert.equal(s[0].options[0].to, 'GUH');
  assert.ok(s[0].options.some((o) => o.doctorId === 'mg4' && o.type === 'dayoff'));
});

test('staffing check: flags days without enough doctors and uses named cover', () => {
  const st = blankState([doc('mg1', 'mg', 'GUH|GUH|GUH|GUH|GUH'), doc('mg2', 'mg', 'OFF|OFF|OFF|OFF|OFF'), doc('r1', 'res', 'GUH|GUH|GUH|GUH|GUH')]);
  const noCover = R.staffingCheck(st, { doctorId: 'mg1', start: '2026-10-14', end: '2026-10-15', part: 'FULL' });
  assert.equal(noCover.bad, 2);
  assert.match(noCover.days[0].text, /Not enough doctors: senior doctor on guh ward/);
  assert.ok(noCover.alternatives.length === 0, 'no senior can be spared on any nearby day either');
  const withCover = R.staffingCheck(st, { doctorId: 'mg1', start: '2026-10-14', end: '2026-10-15', part: 'FULL', guhCoverId: 'mg2' });
  assert.equal(withCover.bad, 0);
  assert.equal(withCover.cover.length, 2);
  assert.equal(withCover.cover[0].to, 'GUH');
});

test('staffing check: competing requests for the same day are warned about', () => {
  const st = blankState([doc('mg1', 'mg', 'GUH|GUH|GUH|GUH|GUH'), doc('mg2', 'mg', 'GUH|GUH|GUH|GUH|GUH')]);
  st.requests.push({ id: 'x', doctorId: 'mg2', type: 'AL', start: '2026-10-14', end: '2026-10-14', part: 'FULL', status: 'pending' });
  const chk = R.staffingCheck(st, { doctorId: 'mg1', start: '2026-10-14', end: '2026-10-14', part: 'FULL' });
  assert.equal(chk.bad, 0);
  assert.equal(chk.warn, 1);
  assert.equal(chk.days[0].competing[0].name, 'mg2');
});

test('leave form: validation, email match, study leave needs Intrepid, sickness applies at once', () => {
  const st = S.demoState(TODAY);
  const base = { doctorId: 'r5', email: 's.malik@example.nhs.uk', type: 'AL', start: '2026-11-24', end: '2026-11-25', part: 'FULL' };
  assert.throws(() => S.submitLeave(st, Object.assign({}, base, { email: 'someone@else.uk' }), { today: TODAY }), /does not match/);
  assert.throws(() => S.submitLeave(st, Object.assign({}, base, { type: 'SL' }), { today: TODAY }), /Intrepid/);
  assert.throws(() => S.submitLeave(st, Object.assign({}, base, { type: 'OTHER' }), { today: TODAY }), /Describe/);
  assert.throws(() => S.submitLeave(st, Object.assign({}, base, { end: '2026-11-20' }), { today: TODAY }), /start date/);
  const out = S.submitLeave(st, Object.assign({}, base, { dayUnitCoverId: 'r7' }), { today: TODAY });
  assert.equal(out.request.status, 'pending');
  assert.equal(out.request.days, 2);
  assert.throws(() => S.submitLeave(st, base, { today: TODAY }), /already have/);
  // pending leave is not on the rota
  assert.notEqual(R.resolveCell(st, R.doctorById(st, 'r5'), '2026-11-24').kind, 'leave');
  const before = R.balance(st, R.doctorById(st, 'r5')).al.left;
  const sick = S.submitLeave(st, { doctorId: 'r5', email: 'S.Malik@example.nhs.uk ', type: 'SICK', start: '2026-10-14', end: '2026-10-14' }, { today: TODAY });
  assert.equal(sick.request.status, 'authorised');
  assert.equal(st.notifications.find((n) => n.to === 'auth').text.includes('reported sick'), true);
  assert.equal(R.resolveCell(st, R.doctorById(st, 'r5'), '2026-10-14').code, 'UNAVAILABLE');
  assert.equal(R.balance(st, R.doctorById(st, 'r5')).al.left, before);
  const mine = S.myRequests(st, 'r5', 's.malik@example.nhs.uk');
  assert.equal(mine.requests.length, 2);
  S.withdrawLeave(st, out.request.id, 's.malik@example.nhs.uk');
  assert.equal(out.request.status, 'cancelled');
});

test('authoriser decisions: decline needs a reason; part-approval leaves out short days; cover goes on the rota', () => {
  const st = blankState([doc('mg1', 'mg', 'GUH|GUH|GUH|GUH|GUH', { email: 'a@x.uk' }), doc('mg2', 'mg', 'GUH|GUH|OFF|GUH|GUH', { email: 'b@x.uk' }), doc('mg3', 'mg', 'RGH|RGH|RGH|RGH|RGH')]);
  st.settings.rules.minSenRGH = 0;
  const r = S.submitLeave(st, { doctorId: 'mg1', email: 'a@x.uk', type: 'AL', start: '2026-10-13', end: '2026-10-14', part: 'FULL', extenuating: true, extenuatingReason: 'Family emergency abroad' }, { today: TODAY }).request;
  // Tue 13: mg2 at GUH (fine). Wed 14: mg2 off, nobody else at GUH (short)
  assert.throws(() => S.ops.decide(st, 'J. Hale', r.id, 'decline', ''), /reason/);
  assert.throws(() => S.ops.decide(st, 'J. Hale', r.id, 'approve', ''), /Give a reason/);
  const emails = S.ops.decide(st, 'J. Hale', r.id, 'approveCovered', '');
  assert.equal(r.status, 'authorised');
  assert.deepEqual(r.excludedDates, ['2026-10-14']);
  assert.equal(r.days, 1);
  assert.equal(R.resolveCell(st, st.doctors[0], '2026-10-13').code, 'ANNUAL LEAVE');
  assert.equal(R.resolveCell(st, st.doctors[0], '2026-10-14').code, 'GUH');
  assert.match(emails[0].subject, /partly approved/);
  assert.equal(st.audit[0].by, 'J. Hale');
  // named cover is written to the rota on approval
  const r2 = S.submitLeave(st, { doctorId: 'mg1', email: 'a@x.uk', type: 'SL', intrepid: true, start: '2026-10-21', end: '2026-10-21', guhCoverId: 'mg3' }, { today: TODAY }).request;
  S.ops.decide(st, 'S. Price', r2.id, 'approve', '');
  assert.equal(R.resolveCell(st, st.doctors[2], '2026-10-21').code, 'GUH');
});

test('EDT/SDT: residents at GUH only with 2 other residents; seniors anywhere if cover kept', () => {
  const st = blankState([
    doc('r1', 'res', 'GUH|GUH|GUH|GUH|GUH'), doc('r2', 'res', 'GUH|GUH|GUH|GUH|GUH'), doc('r3', 'res', 'GUH|RGH|GUH|GUH|GUH'), doc('r4', 'res', 'RGH|RGH|RGH|RGH|RGH'),
    doc('m1', 'mg', 'GUH|GUH|GUH|GUH|GUH'), doc('m2', 'mg', 'RGH|RGH|RGH|RGH|RGH'), doc('m3', 'mg', 'GUH|GUH|GUH|GUH|GUH')
  ]);
  assert.equal(R.edtCheck(st, 'r1', '2026-10-12', 'AM').ok, true, 'Mon: r2 and r3 also at GUH');
  const tue = R.edtCheck(st, 'r1', '2026-10-13', 'AM');
  assert.equal(tue.ok, false, 'Tue: only r2 at GUH');
  assert.match(tue.checks.find((c) => !c.ok).rule, /2 other residents/);
  assert.equal(R.edtCheck(st, 'r4', '2026-10-12', 'PM').ok, false, 'resident not at GUH');
  assert.equal(R.edtCheck(st, 'm2', '2026-10-12', 'PM').ok, true, 'senior cover at RGH is not compulsory');
  st.doctors[6].plan.A = 'RGH|RGH|RGH|RGH|RGH'.split('|');
  assert.equal(R.edtCheck(st, 'm1', '2026-10-12', 'PM').ok, false, 'GUH would lose its only senior');
  st.doctors[6].plan.A = 'GUH|GUH|GUH|GUH|GUH'.split('|');
  assert.equal(R.edtCheck(st, 'm1', '2026-10-12', 'PM').ok, true, 'another senior stays at GUH');
  S.ops.bookEdt(st, 'J. Hale', 'r1', '2026-10-12', 'PM', '');
  const cell = R.resolveCell(st, st.doctors[0], '2026-10-12');
  assert.equal(cell.kind, 'edt');
  assert.equal(cell.code, 'GUH AM · SDT PM');
  assert.equal(cell.pm, '');
  assert.throws(() => S.ops.bookEdt(st, 'J. Hale', 'r4', '2026-10-12', 'PM', ''), /reason/);
});

test('entitlements: Wales annual leave pro-rata and EDT by grade', () => {
  assert.equal(R.alProRata({ grade: 'F1', start: '2026-08-05', end: '2026-12-01' }).days, 9);
  assert.equal(R.alProRata({ grade: 'Registrar ST5' }).days, 33);
  assert.equal(R.alProRata({ grade: 'IMT2', fte: 0.8 }).days, 22.5);
  assert.deepEqual(R.edtDefault({ grade: 'F2' }), { type: 'SDT', hours: 3 });
  assert.deepEqual(R.edtDefault({ grade: 'Specialty doctor', group: 'mg' }), { type: 'SPA', hours: 4 });
  assert.deepEqual(R.edtDefault({ grade: 'IMT3', edtHours: 3 }), { type: 'EDT', hours: 3 });
});

test('rules numbers change the checks', () => {
  const st = blankState([doc('r1', 'res', 'GUH|GUH|GUH|GUH|GUH')]);
  assert.equal(R.coverFor(st, '2026-10-14').rules.find((r) => r.rule.id === 'res-GUH').status, 'ok');
  st.settings.rules.minResGUH = 2;
  assert.equal(R.coverFor(st, '2026-10-14').rules.find((r) => r.rule.id === 'res-GUH').status, 'gap');
  st.settings.rules.minResGUH = 0;
  assert.equal(R.coverFor(st, '2026-10-14').rules.find((r) => r.rule.id === 'res-GUH').status, 'na');
});

test('public view hides emails, notes, reasons and sickness', () => {
  const st = S.demoState(TODAY);
  S.submitLeave(st, { doctorId: 'r5', email: 's.malik@example.nhs.uk', type: 'SICK', start: '2026-10-14', end: '2026-10-14', note: 'private' }, { today: TODAY });
  st.settings.authoriserEmails = ['boss@example.nhs.uk'];
  const pub = S.publicView(st);
  const text = JSON.stringify(pub);
  assert.ok(!text.includes('@example.nhs.uk'));
  assert.ok(!text.includes('private'));
  assert.ok(!text.includes('SICK'));
  assert.equal(pub.audit.length, 0);
  assert.equal(R.resolveCell(pub, R.doctorById(pub, 'r5'), '2026-10-14').code, 'UNAVAILABLE');
  assert.ok(pub.requests.every((r) => !('note' in r) && !('email' in r)));
});

test('copy a doctor and start fresh', () => {
  const st = S.demoState(TODAY);
  S.ops.copyDoctor(st, 'J. Hale', 'r1', { name: 'A. New', start: '2026-12-02', end: '2027-04-06' });
  const nd = st.doctors.find((d) => d.name === 'A. New');
  assert.deepEqual(nd.plan.A, R.doctorById(st, 'r1').plan.A);
  assert.equal(nd.grade, 'F1');
  assert.ok(R.changeoversIn(st, '2026-12-01', '2026-12-03')['2026-12-02'].some((t) => /A. New/.test(t)));
  st.settings.authoriserEmails = ['x@y.uk'];
  S.ops.startFresh(st, 'J. Hale');
  assert.equal(st.doctors.length, 0);
  assert.deepEqual(st.settings.authoriserEmails, ['x@y.uk']);
  assert.equal(st.settings.demo, false);
});

test('name matching: exact, alias, and surname + initial', () => {
  const st = fresh().state;
  assert.equal(E.matchDoctor(st, 'S. Malik').id, 'r5');
  assert.equal(E.matchDoctor(st, 'MALIK S (IMT)').id, 'r5');
  assert.equal(E.matchDoctor(st, 'Dr Tom Hughes').id, 'r4');
  assert.equal(E.matchDoctor(st, 'Somebody Else'), null);
});

test('on-call template round-trips through the importer (list and grid)', async () => {
  const st = fresh().state;
  const sheets = await X.read(E.oncallTemplate(TODAY));
  const list = E.parseOncall(st, sheets);
  assert.equal(list.format, 'list');
  assert.equal(list.entries.length, 10);
  assert.ok(list.entries.every((e) => e.doctorId));
  assert.equal(list.badCodes.length, 0);
  const grid = E.parseOncall(st, [sheets[1]]);
  assert.equal(grid.format, 'grid');
  assert.equal(grid.entries.filter((e) => e.code === 'NIGHTS').length, 4);
  S.ops.applyImport(st, 'J. Hale', 'oncall', { entries: list.entries, from: list.from, to: list.to, fileName: 'test.xlsx' });
  assert.equal(st.oncall.r1[list.from], 'NIGHTS');
});

test('on-call import flags unknown names and codes', async () => {
  const st = fresh().state;
  const bytes = X.write([{ name: 'On-call (list)', rows: [['Date', 'Doctor', 'Shift'], ['19/10/2026', 'New Person', 'NIGHTS'], ['20/10/2026', 'E. Ward', 'BANANA'], ['bad', 'E. Ward', 'NIGHTS']] }]);
  const p = E.parseOncall(st, await X.read(bytes));
  assert.deepEqual(Object.keys(p.unknownNames), ['New Person']);
  assert.equal(p.badCodes.length, 1);
  assert.equal(p.badDates, 1);
});

test('weekly plan and staff templates round-trip', async () => {
  const st = fresh().state;
  const plans = E.parsePlans(st, await X.read(E.planTemplate(st)));
  const meera = plans.plans.find((p) => p.doctorId === 'm5');
  assert.deepEqual(meera.A, ['GUH LAB', 'GUH LAB', 'GUH', 'GUH', 'OFF']);
  assert.ok(meera.B);
  const staff = E.parseStaff(st, await X.read(E.staffTemplate(st)));
  assert.equal(staff.staff.length, st.doctors.length);
  assert.equal(staff.staff.find((s) => s.name === 'Omar').group, 'mg');
  assert.equal(staff.staff.find((s) => s.name === 'Meera').group, 'mg');
  assert.equal(staff.staff.find((s) => s.name === 'L. Osei').group, 'res');
});

test('roster export opens and has the expected sheets (also from the public view)', async () => {
  const st = fresh().state;
  for (const s of [st, S.publicView(st)]) {
    const wb = await X.read(E.rosterWorkbook(s, '2026-10-12', '2026-10-25'));
    assert.deepEqual(wb.map((x) => x.name), ['Roster', 'Leave log', 'Change log', 'About']);
    assert.equal(wb[0].rows.length, 15);
    assert.deepEqual(wb[0].rows[1][0], { date: '2026-10-12' });
  }
});

test('annual leave needs 6 weeks notice unless there are extenuating circumstances', () => {
  const st = S.demoState(TODAY);
  const base = { doctorId: 'r5', email: 's.malik@example.nhs.uk', type: 'AL', start: '2026-11-02', end: '2026-11-03' };
  assert.throws(() => S.submitLeave(st, base, { today: TODAY }), /6 weeks/);
  assert.throws(() => S.submitLeave(st, Object.assign({ extenuating: true, extenuatingReason: 'short' }, base), { today: TODAY }), /explain/);
  const r = S.submitLeave(st, Object.assign({ extenuating: true, extenuatingReason: 'Close family member unwell abroad' }, base), { today: TODAY }).request;
  assert.equal(r.shortNotice, true);
  assert.equal(S.publicView(st).requests.find((x) => x.id === r.id).extenuatingReason, undefined);
  // sickness and study leave are not held to the annual leave notice period
  assert.equal(S.submitLeave(st, { doctorId: 'r7', email: 'l.osei@example.nhs.uk', type: 'SL', intrepid: true, start: '2026-10-20', end: '2026-10-20' }, { today: TODAY }).request.status, 'pending');
  assert.throws(() => S.submitLeave(st, { doctorId: 'r9', email: '', type: 'SICK', start: '2026-10-14', end: '2026-10-14' }, { today: TODAY }), /Choose your name/, 'placeholders cannot request leave');
});

test('authorisers can approve individual days of a request', () => {
  const st = blankState([doc('mg1', 'mg', 'GUH|GUH|GUH|GUH|GUH', { email: 'a@x.uk' }), doc('mg2', 'mg', 'GUH|GUH|OFF|GUH|GUH')]);
  const r = S.submitLeave(st, { doctorId: 'mg1', email: 'a@x.uk', type: 'AL', start: '2026-11-23', end: '2026-11-27' }, { today: TODAY }).request;
  // Wed 25 has no other senior at GUH
  assert.throws(() => S.ops.decide(st, 'J. Hale', r.id, 'approveDays', '', { days: [] }), /Tick at least one/);
  assert.throws(() => S.ops.decide(st, 'J. Hale', r.id, 'approveDays', '', { days: ['2026-11-24', '2026-11-25'] }), /Not enough doctors on Wed 25 Nov/);
  const emails = S.ops.decide(st, 'J. Hale', r.id, 'approveDays', '', { days: ['2026-11-23', '2026-11-24', '2026-11-26'] });
  assert.deepEqual(r.excludedDates, ['2026-11-25', '2026-11-27']);
  assert.equal(r.days, 3);
  assert.match(emails[0].text, /Approved: Mon 23 Nov, Tue 24 Nov, Thu 26 Nov\.\nNot approved: Wed 25 Nov, Fri 27 Nov\./);
  assert.equal(R.resolveCell(st, st.doctors[0], '2026-11-27').code, 'GUH');
  const mine = S.myRequests(st, 'mg1', 'a@x.uk');
  assert.equal(mine.balance.al.used, 3);
});

test('a new doctor continues the column of the doctor they replace', () => {
  const st = S.demoState(TODAY);
  const cols = R.columns(st, '2026-11-30', '2026-12-13', (d) => d.group === 'res');
  const col = cols.find((c) => c.slot === 'r1');
  assert.deepEqual(col.docs.map((d) => d.name), ['E. Ward', 'New F1']);
  assert.equal(R.holderOn(col, '2026-12-01').name, 'E. Ward');
  assert.equal(R.holderOn(col, '2026-12-02').name, 'New F1');
  const ch = R.changeoversIn(st, '2026-12-02', '2026-12-02')['2026-12-02'];
  assert.ok(ch.includes('E. Ward → New F1 (F1)'), ch.join(' | '));
  const k = R.doctorById(st, 'r8');
  assert.ok(R.changeoversIn(st, '2026-10-01', '2026-11-01')[k.start].includes('J. Reid → K. Ahmed (F2)'));
});

test('rotation dates follow the programme and can be overridden', () => {
  assert.deepEqual(R.programmeStarts('foundation', '2026-08-01', 4), ['2026-08-05', '2026-12-02', '2027-04-07', '2027-08-04']);
  assert.deepEqual(R.programmeStarts('imt', '2026-08-06', 2), ['2027-02-03', '2027-08-04']);
  assert.deepEqual(R.programmeStarts('registrar', '2026-08-02', 2), ['2027-02-01', '2027-08-01']);
  const f = R.suggestDates('F2', '2026-10-12');
  assert.equal(f.start, '2026-12-02'); assert.equal(f.end, '2027-04-06'); assert.equal(f.provisional, false);
  assert.equal(R.suggestDates('Registrar ST5', '2026-10-12').provisional, true);
  assert.equal(R.suggestDates('Specialty doctor', '2026-10-12').start, '');
  assert.ok(R.rotaHorizon(S.demoState(TODAY), TODAY) >= '2027-08-03', 'whole rotation runs to the last known date');
});

test('older saved rotas are upgraded', () => {
  const st = S.demoState(TODAY);
  st.settings.schema = 2; st.settings.rules.minSenRGH = 1; st.settings.rotation.end = '2026-12-01';
  S.upgrade(st);
  assert.equal(st.settings.rules.minSenRGH, 0);
  assert.equal(st.settings.rotation.end, undefined);
  assert.equal(st.settings.schema, 3);
});

test('a group of doctors can be put at a location for a block of dates', () => {
  const st = blankState([doc('a', 'res', 'GUH|GUH|GUH|GUH|GUH'), doc('b', 'mg', 'RGH|RGH|RGH|RGH|RGH')]);
  assert.throws(() => S.ops.bulkAssign(st, 'J. Hale', [], '2026-10-12', '2026-10-16', [0], 'NHH', ''), /Tick at least one doctor/);
  S.ops.bulkAssign(st, 'J. Hale', ['a', 'b'], '2026-10-12', '2026-10-25', [1, 2, 3, 4], 'NHH - DAY UNIT', 'NHH cover');
  assert.equal(R.resolveCell(st, st.doctors[0], '2026-10-12').code, 'GUH'); // Monday not ticked
  assert.equal(R.resolveCell(st, st.doctors[1], '2026-10-13').code, 'NHH - DAY UNIT');
  assert.equal(R.resolveCell(st, st.doctors[1], '2026-10-17').code, ''); // weekend untouched
  assert.match(st.audit[0].detail, /on 16 days/); // 2 doctors x 8 weekdays
});
