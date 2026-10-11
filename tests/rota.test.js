// Run with: node --test tests/
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
const fresh = () => S.createStore({ storage: null, today: TODAY });

function blankState(doctors) {
  return {
    settings: { bankHolidays: R.BANK_HOLIDAYS, restricted: [], consultantOfWeek: {} },
    doctors, users: [], oncall: {}, overrides: {}, requests: [], audit: [], notifications: []
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

test('half-day leave keeps the doctor counted at their site', () => {
  const d = doc('a', 'res', 'RGH|RGH|RGH|RGH|RGH');
  const st = blankState([d]);
  st.requests.push({ doctorId: 'a', type: 'AL', start: '2026-10-15', end: '2026-10-15', part: 'PM', status: 'authorised' });
  const c = R.resolveCell(st, d, '2026-10-15');
  assert.equal(c.code, 'RGH AM; PM ANNUAL LEAVE');
  assert.equal(R.coverFor(st, '2026-10-15').rules.find((r) => r.rule.id === 'res-RGH').status, 'ok');
});

test('pending leave shows the day as at risk', () => {
  const st = blankState([doc('mg1', 'mg', 'GUH|GUH|GUH|GUH|GUH'), doc('mg2', 'mg', 'RGH|RGH|RGH|RGH|RGH')]);
  st.requests.push({ doctorId: 'mg1', type: 'AL', start: '2026-10-14', end: '2026-10-14', part: 'FULL', status: 'pending' });
  assert.equal(R.coverWithRisk(st, '2026-10-14').rules.find((r) => r.rule.id === 'mg-GUH').status, 'risk');
});

test('cover suggestions never leave a gap and prefer clean moves', () => {
  const st = blankState([
    doc('mg1', 'mg', 'GUH|GUH|GUH|GUH|GUH'), doc('mg2', 'mg', 'RGH|RGH|RGH|RGH|RGH'), doc('mg3', 'mg', 'RGH MDCU|RGH MDCU|RGH MDCU|RGH MDCU|RGH MDCU'),
    doc('mg4', 'mg', 'OFF|OFF|OFF|OFF|OFF'), doc('mg5', 'mg', 'NHH|NHH|NHH|NHH|NHH'), doc('res1', 'res', 'NHH|NHH|NHH|NHH|NHH')
  ]);
  const s = R.suggestCover(st, 'mg1', '2026-10-14', '2026-10-14', 'FULL');
  assert.equal(s.length, 1);
  assert.ok(s[0].needed);
  assert.deepEqual(s[0].broken, ['Middle grade on GUH ward']);
  const first = s[0].options[0];
  assert.ok(['mg2', 'mg3'].includes(first.doctorId), 'a spare RGH middle grade moves first');
  assert.equal(first.to, 'GUH');
  assert.ok(s[0].options.every((o) => o.doctorId !== 'mg5' || o.exceptions.length), 'moving NHH leaves the NHH day unit to a resident (exception)');
  assert.ok(s[0].options.some((o) => o.doctorId === 'mg4' && o.type === 'dayoff'));
});

test('store: doctors cannot edit the roster; leave needs a coordinator', () => {
  const store = fresh();
  store.setUser('u-r1');
  assert.throws(() => store.setOverride('r1', '2026-10-13', 'RGH'), /Only rota coordinators/);
  const r = store.submitRequest({ doctorId: 'r1', type: 'AL', start: '2026-10-27', end: '2026-10-28', part: 'FULL' });
  assert.equal(r.status, 'pending');
  assert.throws(() => store.decideRequest(r.id, 'authorised'), /Only rota coordinators/);
  assert.throws(() => store.submitRequest({ doctorId: 'r2', type: 'AL', start: '2026-10-27', end: '2026-10-27' }), /only request leave for yourself/);
  store.setUser('c2');
  assert.throws(() => store.decideRequest(r.id, 'declined', ''), /Give a reason/);
  store.decideRequest(r.id, 'authorised');
  assert.equal(R.resolveCell(store.state, R.doctorById(store.state, 'r1'), '2026-10-27').code, 'ANNUAL LEAVE');
  assert.equal(store.state.audit[0].action, 'Leave authorised');
  assert.equal(store.state.audit[0].by, 'S. Price');
});

test('store: sickness applies at once, notifies coordinators, and does not use annual leave', () => {
  const store = fresh();
  store.setUser('u-r5');
  const before = R.balance(store.state, R.doctorById(store.state, 'r5')).al.left;
  const r = store.submitRequest({ doctorId: 'r5', type: 'SICK', start: '2026-10-14', end: '2026-10-14' });
  assert.equal(r.status, 'authorised');
  assert.equal(store.state.notifications[0].to, 'coord');
  assert.equal(R.balance(store.state, R.doctorById(store.state, 'r5')).al.left, before);
});

test('authorising applies the chosen cover move', () => {
  const store = fresh();
  const meera = store.state.requests.find((r) => r.doctorId === 'm5' && r.status === 'pending');
  assert.ok(meera.cover.length >= 0);
  store.setUser('c1');
  store.decideRequest(meera.id, 'authorised');
  for (const c of meera.cover) {
    assert.equal(R.resolveCell(store.state, R.doctorById(store.state, c.doctorId), c.date).code, c.to);
  }
});

test('name matching: exact, alias, and surname + initial', () => {
  const st = fresh().state;
  assert.equal(E.matchDoctor(st, 'S. Malik').id, 'r5');
  assert.equal(E.matchDoctor(st, 'MALIK S (IMT)').id, 'r5');
  assert.equal(E.matchDoctor(st, 'Dr Tom Hughes').id, 'r4');
  assert.equal(E.matchDoctor(st, 'Somebody Else'), null);
});

test('on-call template round-trips through the importer (list and grid)', async () => {
  const store = fresh();
  const sheets = await X.read(E.oncallTemplate(TODAY));
  const list = E.parseOncall(store.state, sheets);
  assert.equal(list.format, 'list');
  assert.equal(list.entries.length, 10);
  assert.ok(list.entries.every((e) => e.doctorId));
  assert.equal(list.badCodes.length, 0);
  const grid = E.parseOncall(store.state, [sheets[1]]);
  assert.equal(grid.format, 'grid');
  assert.equal(grid.entries.filter((e) => e.code === 'NIGHTS').length, 4);
  store.applyImport('oncall', { entries: list.entries, from: list.from, to: list.to, fileName: 'test.xlsx' });
  assert.equal(store.state.oncall.r1[list.from], 'NIGHTS');
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
  assert.equal(staff.staff.find((s) => s.name === 'L. Osei').group, 'res');
});

test('roster export opens and has the expected sheets', async () => {
  const st = fresh().state;
  const wb = await X.read(E.rosterWorkbook(st, '2026-10-12', '2026-10-25'));
  assert.deepEqual(wb.map((s) => s.name), ['Roster', 'Leave log', 'Change log', 'About']);
  assert.equal(wb[0].rows.length, 15);
  assert.deepEqual(wb[0].rows[1][0], { date: '2026-10-12' });
});
