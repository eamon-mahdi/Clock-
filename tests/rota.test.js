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
  const tue = R.coverFor(st, '2026-10-13');
  const by = Object.fromEntries(tue.rules.map((r) => [r.rule.id, r.status]));
  assert.equal(by['mg-RGH'], 'ok'); // RGH MDCU counts for the RGH day unit
  assert.equal(by['mg-NHH'], 'exception'); // only a resident at NHH
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
  assert.throws(() => S.submitLeave(st, Object.assign({}, base, { email: 'someone@else.uk' })), /does not match/);
  assert.throws(() => S.submitLeave(st, Object.assign({}, base, { type: 'SL' })), /Intrepid/);
  assert.throws(() => S.submitLeave(st, Object.assign({}, base, { type: 'OTHER' })), /Describe/);
  assert.throws(() => S.submitLeave(st, Object.assign({}, base, { end: '2026-11-20' })), /start date/);
  const out = S.submitLeave(st, Object.assign({}, base, { dayUnitCoverId: 'r7', remainingStated: '20 days' }));
  assert.equal(out.request.status, 'pending');
  assert.equal(out.request.days, 2);
  assert.equal(out.request.remainingStated, '20 days');
  assert.throws(() => S.submitLeave(st, base), /already have/);
  // pending leave is not on the rota
  assert.notEqual(R.resolveCell(st, R.doctorById(st, 'r5'), '2026-11-24').kind, 'leave');
  const before = R.balance(st, R.doctorById(st, 'r5')).al.left;
  const sick = S.submitLeave(st, { doctorId: 'r5', email: 'S.Malik@example.nhs.uk ', type: 'SICK', start: '2026-10-14', end: '2026-10-14' });
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
  const r = S.submitLeave(st, { doctorId: 'mg1', email: 'a@x.uk', type: 'AL', start: '2026-10-13', end: '2026-10-14', part: 'FULL' }).request;
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
  const r2 = S.submitLeave(st, { doctorId: 'mg1', email: 'a@x.uk', type: 'AL', start: '2026-10-21', end: '2026-10-21', guhCoverId: 'mg3' }).request;
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
  assert.equal(R.edtCheck(st, 'm2', '2026-10-12', 'PM').ok, false, 'only senior at RGH');
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
  S.submitLeave(st, { doctorId: 'r5', email: 's.malik@example.nhs.uk', type: 'SICK', start: '2026-10-14', end: '2026-10-14', note: 'private' });
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
