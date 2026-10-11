/* Screens and interactions for the Haematology Rota prototype. */
(function () {
  'use strict';
  const R = window.RotaRules, S = window.RotaStore, E = window.RotaExcel, X = window.RotaXlsx;
  const store = S.createStore();
  const TODAY = R.toIso(new Date());
  const nextWeekday = (iso) => { let d = iso; while (R.isWeekend(d)) d = R.addDays(d, 1); return d; };

  const ui = {
    view: 'roster', rosterStart: R.mondayOf(TODAY), weeks: 2, group: 'both', hi: 'all', q: '',
    day: nextWeekday(TODAY), editing: null, reqSel: null, declining: false, setupTab: 'people',
    importPreview: null, confirm: null, logQ: '', form: null, lastSubmitted: null
  };
  try { const v = localStorage.getItem('haem-rota-view'); if (v) ui.view = v; } catch (e) { /* no storage */ }

  // ---------- helpers ----------
  const esc = (s) => String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const $ = (sel) => document.querySelector(sel);
  const st = () => store.state;
  const me = () => store.user();
  const role = () => me().role;
  const docName = (id) => (R.doctorById(st(), id) || { name: '?' }).name;
  const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;
  const weeksDays = (n) => `${plural(Math.floor(n / 7), 'week')} ${plural(n % 7, 'day')}`;
  const rangeText = (a, b) => (a === b ? R.fmtShort(a) : `${R.fmtShort(a)} – ${R.fmtShort(b)}`);
  const statusWord = { ok: 'covered', gap: 'gap', exception: 'resident only (exception)', risk: 'at risk (leave requested)', na: 'not needed' };
  const statusMark = { ok: '✓', gap: '✕', exception: 'res', risk: '!', na: '–' };

  function toast(msg, isError) {
    const el = document.createElement('div');
    el.className = 'toast' + (isError ? ' error' : '');
    el.textContent = msg;
    $('#toasts').appendChild(el);
    setTimeout(() => el.remove(), isError ? 6000 : 3500);
  }
  function attempt(fn) {
    try { fn(); return true; } catch (e) { toast(e.message || String(e), true); return false; }
  }

  async function saveFile(filename, bytes) {
    let dl = null;
    if (window.claude && typeof window.claude.use === 'function') {
      try { dl = await window.claude.use('downloads'); } catch (e) { dl = null; }
      if (!dl) { toast('Downloads are not available in this view.', true); return; }
      try { await dl.save({ filename, data: new Blob([bytes]) }); toast(`Saved ${filename}`); } catch (e) {
        if (e && e.code !== 'declined') toast(`Could not save the file (${e.code || e.message}).`, true);
      }
      return;
    }
    const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = document.createElement('a');
    a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 3000);
    toast(`Downloaded ${filename}`);
  }

  // ---------- navigation ----------
  function navItems() {
    const pending = st().requests.filter((r) => r.status === 'pending').length;
    const unread = st().notifications.filter((n) => !n.read && n.to === me().id).length;
    const items = [['roster', 'Roster'], ['day', 'Day by site']];
    if (role() === 'coordinator') items.push(['requests', 'Requests', pending], ['request', 'Request leave'], ['setup', 'Set up'], ['log', 'Change log']);
    if (role() === 'doctor') items.push(['request', 'Request leave'], ['myleave', 'My leave', unread]);
    items.push(['how', 'How it works']);
    return items;
  }
  function renderNav() {
    const items = navItems();
    if (!items.some((i) => i[0] === ui.view)) ui.view = 'roster';
    $('#nav').innerHTML = items.map(([id, label, n]) => `<button type="button" data-act="nav" data-view="${id}"${ui.view === id ? ' aria-current="page"' : ''}>${esc(label)}${n ? `<span class="badge">${n}</span>` : ''}</button>`).join('');
    const users = st().users;
    const opt = (u) => `<option value="${esc(u.id)}"${u.id === me().id ? ' selected' : ''}>${esc(u.name)}</option>`;
    $('#whoami').innerHTML =
      `<optgroup label="Rota coordinators">${users.filter((u) => u.role === 'coordinator').map(opt).join('')}</optgroup>` +
      `<optgroup label="Doctors">${users.filter((u) => u.role === 'doctor').map(opt).join('')}</optgroup>` +
      `<optgroup label="Read-only">${users.filter((u) => u.role === 'viewer').map(opt).join('')}</optgroup>`;
  }

  // ---------- roster ----------
  function cellCategory(cell) {
    if (cell.kind === 'site') return R.siteOf(cell.code);
    if (cell.kind === 'half') return 'leave';
    if (cell.kind === 'leave' || cell.kind === 'sick') return 'leave';
    if (cell.kind === 'oncall') return 'oncall';
    if (cell.kind === 'off') return 'off';
    return cell.kind;
  }
  function describeCell(doc, iso, cell) {
    const parts = [`${doc.name}, ${R.fmtLong(iso)}: ${cell.code || 'nothing rostered'}`];
    if (cell.source === 'plan') parts.push('From the weekly plan.');
    if (cell.source === 'oncall') parts.push('Imported from the trust on-call rota.');
    if (cell.source === 'override') parts.push(`Changed by the rota coordinator${cell.note ? ': ' + cell.note : '.'}`);
    if (cell.leave) parts.push(cell.leave.type === 'SICK' ? 'Sickness (self-reported).' : `Authorised by ${cell.leave.decidedBy}.`);
    if (cell.pending) parts.push(`${R.LEAVE_TYPES[cell.pending.type].label} requested, waiting for a rota coordinator. Not on the roster until authorised.`);
    if (cell.conflict) parts.push(cell.conflict);
    return parts.join(' ');
  }
  function cellHtml(doc, iso, canEdit) {
    const cell = R.resolveCell(st(), doc, iso);
    const site = cell.kind === 'half' ? (R.siteOf(cell.baseCode) || 'none') : R.siteOf(cell.code);
    let label = esc(cell.code);
    if (cell.kind === 'oncall') label = '⇣ ' + label;
    if (cell.kind === 'gone') label = '<span class="small">not in post</span>';
    if (cell.pending && cell.kind !== 'leave') label = `<span>${label}</span><span style="font-weight:500;font-size:0.66rem">${esc({ AL: 'AL', SL: 'SL', PL: 'PL', TOIL: 'TOIL' }[cell.pending.type] || 'leave')} requested</span>`;
    const cls = `cell k-${cell.kind}${site ? ' s-' + site : ''}${cell.pending && cell.kind !== 'leave' ? ' pending' : ''}`;
    const marks = (cell.source === 'override' ? '<span class="mark override" aria-hidden="true"></span>' : '') + (cell.conflict ? '<span class="mark conflict" aria-hidden="true"></span>' : '');
    const title = esc(describeCell(doc, iso, cell));
    if (canEdit && cell.kind !== 'gone') return `<button type="button" class="${cls}" data-act="edit-cell" data-doc="${doc.id}" data-iso="${iso}" title="${title}" aria-label="${title}">${label}${marks}</button>`;
    return `<span class="${cls}" title="${title}">${label}${marks}</span>`;
  }
  const coverChips = (cover, group) => (cover ? cover.rules.filter((r) => r.rule.group === group)
    .map((r) => `<span class="cv ${r.status}" title="${esc(r.rule.label)}: ${esc(statusWord[r.status])}${r.names.length ? ' (' + esc(r.names.join(', ')) + ')' : ''}">${esc(r.rule.short)} ${statusMark[r.status]}</span>`).join('') : '');

  function renderRoster() {
    const from = ui.rosterStart, to = R.addDays(from, ui.weeks * 7 - 1);
    const dates = R.dateRange(from, to);
    const q = ui.q.trim().toLowerCase();
    const visible = (d) => (!q || d.name.toLowerCase().includes(q) || d.grade.toLowerCase().includes(q)) &&
      dates.some((iso) => R.inPost(d, iso) || (d.end && d.end >= R.addDays(from, -28)));
    const res = st().doctors.filter((d) => d.group === 'res' && visible(d));
    const mg = st().doctors.filter((d) => d.group === 'mg' && visible(d));
    const showRes = ui.group !== 'mg', showMg = ui.group !== 'res';
    const canEdit = role() === 'coordinator';
    const docTh = (d) => `<th class="doc" scope="col"><span class="grade">${esc(d.grade)}</span>${esc(d.name)}${d.end && d.end < to ? `<span class="dates">${d.end < from ? 'left ' : 'to '}${esc(R.fmtUk(d.end).slice(0, 5))}</span>` : ''}</th>`;
    const chipBtn = (act, val, cur, label) => `<button type="button" class="chip" data-act="${act}" data-val="${val}" aria-pressed="${cur === val}">${label}</button>`;
    const td = (d, iso) => {
      const cell = R.resolveCell(st(), d, iso);
      const dim = ui.hi !== 'all' && cell.kind !== 'none' && cellCategory(cell) !== ui.hi;
      return `<td class="${dim ? 'dim' : ''}">${cellHtml(d, iso, canEdit)}</td>`;
    };
    const rows = dates.map((iso) => {
      const wk = R.isWeekend(iso);
      const restr = S.restrictionOn(st(), iso);
      const cover = R.coverWithRisk(st(), iso);
      const bh = R.isBankHoliday(st(), iso);
      return `<tr class="${wk ? 'weekend' : ''}${iso === TODAY ? ' today' : ''}">
        <th class="col-date" scope="row"><span class="tabular">${esc(R.fmtShort(iso))}</span>${restr ? `<span class="restr ${restr.level.replace(' ', '-')}" title="${esc(restr.note || '')}">${esc(restr.level)}</span>` : ''}${bh ? '<span class="restr AVOID" title="Bank holiday">BH</span>' : ''}</th>
        <td class="cons">${wk ? '' : esc(st().settings.consultantOfWeek[R.mondayOf(iso)] || '')}</td>
        ${showRes ? res.map((d) => td(d, iso)).join('') + '<td class="gap-col"></td>' : ''}
        ${showMg ? mg.map((d) => td(d, iso)).join('') + '<td class="gap-col"></td>' : ''}
        ${showRes ? `<td class="cover-cell">${coverChips(cover, 'res')}</td>` : ''}
        ${showMg ? `<td class="cover-cell">${coverChips(cover, 'mg')}</td>` : ''}
      </tr>`;
    }).join('');
    const gaps = dates.map((iso) => R.coverWithRisk(st(), iso)).filter(Boolean)
      .reduce((a, c) => { c.rules.forEach((r) => { if (r.status === 'gap') a.gap++; if (r.status === 'exception') a.exc++; if (r.status === 'risk') a.risk++; }); return a; }, { gap: 0, exc: 0, risk: 0 });
    return `
    <div class="page-head">
      <div><h1>Roster</h1><p class="muted">${esc(R.fmtLong(from))} to ${esc(R.fmtLong(to))}. ${canEdit ? 'Click any cell to change it. Every change is logged.' : 'Read-only. Only the rota coordinators can change the roster.'}</p></div>
      <div class="row">
        <button type="button" class="btn icon" data-act="roster-move" data-val="-7" aria-label="Previous week">‹</button>
        <button type="button" class="btn" data-act="roster-today">This week</button>
        <button type="button" class="btn icon" data-act="roster-move" data-val="7" aria-label="Next week">›</button>
        <label class="sr-only" for="weeks">Weeks shown</label>
        <select id="weeks" data-act="weeks"><option value="1"${ui.weeks === 1 ? ' selected' : ''}>1 week</option><option value="2"${ui.weeks === 2 ? ' selected' : ''}>2 weeks</option><option value="4"${ui.weeks === 4 ? ' selected' : ''}>4 weeks</option></select>
        <button type="button" class="btn primary" data-act="export">Download Excel</button>
      </div>
    </div>
    <div class="row">
      <span class="cv gap">${plural(gaps.gap, 'gap')}</span>
      <span class="cv exception">${gaps.exc} day-unit exception${gaps.exc === 1 ? '' : 's'}</span>
      <span class="cv risk">${gaps.risk} at risk from leave requests</span>
    </div>
    <div class="toolbar">
      <div class="group"><span class="label">Show</span>${chipBtn('group', 'both', ui.group, 'Everyone')}${chipBtn('group', 'res', ui.group, 'Residents')}${chipBtn('group', 'mg', ui.group, 'Middle grade')}</div>
      <div class="group"><span class="label">Highlight</span>${[['all', 'All'], ['GUH', 'GUH'], ['RGH', 'RGH'], ['NHH', 'NHH'], ['leave', 'Leave'], ['oncall', 'On call'], ['off', 'Off'], ['lab', 'Lab']].map(([v, l]) => chipBtn('hi', v, ui.hi, l)).join('')}</div>
      <label class="group" for="q"><span class="label">Find</span><input type="search" id="q" data-live="q" value="${esc(ui.q)}" placeholder="Name or grade"></label>
    </div>
    <div class="table-wrap">
      <table class="roster">
        <thead>
          <tr class="groups"><th class="col-date"></th><th></th>
            ${showRes ? `<th colspan="${res.length}">Resident doctors · 1 at GUH and RGH Mon–Fri, 1 at NHH Tue–Fri</th><th class="gap-col"></th>` : ''}
            ${showMg ? `<th colspan="${mg.length}">Middle grade · GUH ward (mandatory), RGH day unit Mon–Fri, NHH day unit Tue–Fri</th><th class="gap-col"></th>` : ''}
            <th colspan="${(showRes ? 1 : 0) + (showMg ? 1 : 0)}">Cover check</th></tr>
          <tr class="names"><th class="col-date" scope="col">Date</th><th scope="col">Consultant</th>
            ${showRes ? res.map(docTh).join('') + '<th class="gap-col"></th>' : ''}
            ${showMg ? mg.map(docTh).join('') + '<th class="gap-col"></th>' : ''}
            ${showRes ? '<th scope="col">Residents</th>' : ''}${showMg ? '<th scope="col">Middle grade</th>' : ''}</tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    <div class="legend">
      <span><span class="sw k-site s-GUH"></span>GUH (inpatient)</span>
      <span><span class="sw k-site s-RGH"></span>RGH incl. day unit and MDCU</span>
      <span><span class="sw k-site s-NHH"></span>NHH incl. day unit</span>
      <span><span class="sw k-lab"></span>Lab (separate, not site cover)</span>
      <span><span class="sw k-leave"></span>Leave</span>
      <span><span class="sw k-oncall"></span>⇣ On call (imported)</span>
      <span><span class="sw cell pending" style="min-height:0;display:inline-block"></span>Leave requested, not authorised</span>
      <span><span class="sw k-gone"></span>Not in post</span>
      <span><span class="cv ok">✓</span>covered <span class="cv exception">res</span>resident only <span class="cv risk">!</span>at risk <span class="cv gap">✕</span>gap</span>
    </div>`;
  }

  function renderEditSheet() {
    if (!ui.editing) return '';
    const doc = R.doctorById(st(), ui.editing.doc);
    const iso = ui.editing.iso;
    const cell = R.resolveCell(st(), doc, iso);
    const isLeave = cell.kind === 'leave' || cell.kind === 'sick' || cell.kind === 'half';
    const opts = (list) => list.map((c) => `<option value="${esc(c)}"${c === cell.code ? ' selected' : ''}>${esc(c)}</option>`).join('');
    return `<section class="sheet" role="dialog" aria-labelledby="sheet-title">
      <div><h3 id="sheet-title">Change ${esc(doc.name)}, ${esc(R.fmtLong(iso))}</h3><p class="muted small">Now: <strong>${esc(cell.code || 'nothing rostered')}</strong> (${esc({ plan: 'weekly plan', oncall: 'imported on-call', override: 'changed by coordinator', leave: 'authorised leave', weekend: 'weekend', bh: 'bank holiday', none: 'blank' }[cell.source] || cell.source)})</p></div>
      ${cell.source === 'oncall' ? '<p class="notice warn small">This shift came from the trust on-call rota. Changing it here overrides the import for this day only and is logged.</p>' : ''}
      ${isLeave ? '<p class="notice info small">This is authorised leave or sickness. To change it, open the request under Requests or ask the doctor to withdraw it.</p>' : `
      <label class="field" for="edit-code">New value
        <select id="edit-code">
          <optgroup label="Sites">${opts(R.SITE_CODES)}</optgroup>
          <optgroup label="On call">${opts(R.ONCALL_CODES)}</optgroup>
        </select></label>
      <label class="field" for="edit-reason">Reason <span class="hint">Shown in the change log</span><input type="text" id="edit-reason" placeholder="e.g. covering RGH day unit"></label>`}
      <div class="row">
        ${isLeave ? '' : '<button type="button" class="btn primary" data-act="edit-save">Save change</button>'}
        ${cell.source === 'override' ? '<button type="button" class="btn" data-act="edit-clear">Undo change</button>' : ''}
        <button type="button" class="btn ghost" data-act="edit-close">Close</button>
      </div>
    </section>`;
  }

  // ---------- day by site ----------
  function renderDay() {
    const iso = ui.day;
    const cover = R.coverWithRisk(st(), iso);
    const wk = R.isWeekend(iso) || R.isBankHoliday(st(), iso);
    const groupOk = (d) => ui.group === 'both' || d.group === ui.group;
    const people = st().doctors.filter((d) => groupOk(d) && (!ui.q || d.name.toLowerCase().includes(ui.q.toLowerCase())))
      .map((d) => ({ d, cell: R.resolveCell(st(), d, iso) }));
    const mon = R.mondayOf(iso);
    const days = [0, 1, 2, 3, 4].map((i) => R.addDays(mon, i));
    const dayBtns = days.map((d) => {
      const w = R.worstStatus(R.coverWithRisk(st(), d));
      return `<button type="button" class="daybtn" data-act="pick-day" data-val="${d}" aria-pressed="${d === iso}"><span class="d">${R.DOW_SHORT[R.dow(d)]}</span><span class="n tabular">${R.toDate(d).getUTCDate()}</span><span class="f ${w}">${w === 'na' ? '–' : statusMark[w] === 'res' ? '!' : statusMark[w]}</span></button>`;
    }).join('');
    const personRow = (p, where) => `<div class="person"><span class="gbadge ${p.d.group}">${esc(p.d.group === 'mg' ? 'MG' : p.d.grade.split(' ')[0])}</span><span style="flex:1;min-width:0"><strong>${esc(p.d.name)}</strong><br><span class="muted small">${esc(where)}${p.cell.pending ? ` · ${esc(R.LEAVE_TYPES[p.cell.pending.type].label.toLowerCase())} requested` : ''}</span></span></div>`;
    const sites = [['GUH', 'Inpatient'], ['RGH', 'Day unit Mon–Fri'], ['NHH', 'Day unit Tue–Fri']].map(([s, sub]) => {
      const reqs = cover ? cover.rules.filter((r) => r.rule.codes.includes(s === 'GUH' ? 'GUH' : s === 'RGH' ? 'RGH' : 'NHH')) : [];
      const here = people.filter((p) => R.siteOf(R.coverCode(p.cell)) === s).sort((a, b) => (a.d.group === b.d.group ? 0 : a.d.group === 'mg' ? -1 : 1));
      return `<article class="card site-card">
        <div class="site-head s-${s}"><h2>${s}</h2><span class="small">${sub}</span></div>
        <div class="body">
          <div class="row">${reqs.map((r) => `<span class="cv ${r.status}">${esc(r.rule.label)}: ${esc(statusWord[r.status])}</span>`).join('')}</div>
          ${here.map((p) => personRow(p, p.cell.kind === 'half' ? p.cell.code : p.cell.code)).join('') || '<p class="muted small">Nobody here from this group.</p>'}
        </div></article>`;
    }).join('');
    const bucket = (title, swatch, test) => {
      const list = people.filter((p) => test(p.cell));
      return `<article class="card"><h3><span class="sw ${swatch}" style="display:inline-block;width:14px;height:14px;border-radius:3px;vertical-align:-2px;margin-right:6px"></span>${title} <span class="muted">· ${list.length}</span></h3>
        ${list.map((p) => `<div class="small"><strong>${esc(p.d.name)}</strong> <span class="muted">${esc(p.d.grade)} · ${esc(p.cell.kind === 'gone' ? 'not in post' : p.cell.code)}</span></div>`).join('') || '<p class="muted small">None</p>'}</article>`;
    };
    return `
    <div class="page-head">
      <div><h1>${esc(R.fmtLong(iso))}</h1><p class="muted">Who is where, and whether each post is covered.</p></div>
      <div class="row">
        <button type="button" class="btn icon" data-act="day-move" data-val="-1" aria-label="Previous day">‹</button>
        <label class="sr-only" for="day-pick">Date</label><input type="date" id="day-pick" data-act="day-pick" value="${iso}">
        <button type="button" class="btn icon" data-act="day-move" data-val="1" aria-label="Next day">›</button>
      </div>
    </div>
    <div class="days" role="group" aria-label="Days this week">${dayBtns}</div>
    <div class="toolbar">
      <div class="group"><span class="label">Show</span>${[['both', 'Everyone'], ['res', 'Residents'], ['mg', 'Middle grade']].map(([v, l]) => `<button type="button" class="chip" data-act="group" data-val="${v}" aria-pressed="${ui.group === v}">${l}</button>`).join('')}</div>
      <label class="group" for="q"><span class="label">Find</span><input type="search" id="q" data-live="q" value="${esc(ui.q)}" placeholder="Name"></label>
    </div>
    ${wk ? `<p class="notice info">${R.isBankHoliday(st(), iso) ? 'Bank holiday.' : 'Weekend.'} Haematology sites are closed; on-call shifts are shown below.</p>` : `<div class="grid-3">${sites}</div>`}
    <div class="grid-3">
      ${bucket('On call', 'k-oncall', (c) => c.kind === 'oncall')}
      ${bucket('On leave', 'k-leave', (c) => c.kind === 'leave' || c.kind === 'half')}
      ${bucket('Sick', 'k-sick', (c) => c.kind === 'sick')}
      ${bucket('Lab', 'k-lab', (c) => c.kind === 'lab')}
      ${bucket('Off', 'k-off', (c) => c.kind === 'off' || (c.kind === 'none' && !wk))}
      ${bucket('Not in post', 'k-gone', (c) => c.kind === 'gone')}
    </div>
    <p class="muted small">Sickness shows only as "Sick", never the reason. It is reported by the doctor, the coordinators are notified, and it is not taken from annual leave.</p>`;
  }

  // ---------- leave request ----------
  function defaultForm() {
    const start = nextWeekday(R.addDays(TODAY, 49));
    return { doctorId: me().doctorId || (st().doctors.find((d) => R.inPost(d, TODAY)) || {}).id, type: 'AL', start, end: start, part: 'FULL', note: '', cover: {} };
  }
  function requestChecks(f) {
    const s = st();
    const doc = R.doctorById(s, f.doctorId);
    const out = { doc, days: R.leaveDays(s, f.start, f.end, f.part), problems: [], warnings: [], info: [] };
    if (!doc) return out;
    if (!f.start || !f.end || f.end < f.start) { out.problems.push('Choose a start date and an end date on or after it.'); return out; }
    const bal = R.balance(s, doc);
    out.bal = bal;
    if (!R.inPost(doc, f.start) || !R.inPost(doc, f.end)) out.warnings.push(`Some of these dates are outside ${doc.name}'s time in post (${doc.start ? R.fmtUk(doc.start) : '?'} to ${doc.end ? R.fmtUk(doc.end) : 'no end date'}).`);
    out.notice = R.daysBetween(TODAY, f.start);
    if (f.type !== 'SICK' && out.notice < 42) out.warnings.push(`Short notice: ${weeksDays(out.notice)} (policy is 6 weeks). The coordinator can still authorise it.`);
    const pool = f.type === 'AL' ? bal.al : f.type === 'SL' ? bal.sl : null;
    if (pool) {
      out.poolLeft = pool.left - pool.pending;
      if (out.days > out.poolLeft) out.warnings.push(`This uses ${out.days} days but only ${out.poolLeft} are left after other pending requests.`);
    }
    if (f.type !== 'SICK') for (const r of R.restrictedIn(s, f.start, f.end)) {
      (r.level === 'NO AL' ? out.warnings : out.info).push(`${r.level === 'NO AL' ? 'Leave restricted' : 'Leave discouraged'}: ${rangeText(r.start, r.end)}${r.note ? ` (${r.note})` : ''}.`);
    }
    const oc = s.oncall[doc.id] || {};
    const clashes = R.dateRange(f.start, f.end).filter((d) => oc[d]);
    if (clashes.length) out.warnings.push(`On-call shifts in these dates from the trust rota: ${clashes.map((d) => `${R.fmtShort(d)} ${oc[d]}`).join(', ')}. Arrange a swap with the trust rota too.`);
    const overlap = s.requests.filter((r) => r.doctorId === doc.id && (r.status === 'pending' || r.status === 'authorised') && r.start <= f.end && r.end >= f.start);
    if (overlap.length) out.problems.push(`You already have ${overlap[0].status} ${R.LEAVE_TYPES[overlap[0].type].label.toLowerCase()} on ${rangeText(overlap[0].start, overlap[0].end)}.`);
    out.suggestions = f.type === 'SICK' ? [] : R.suggestCover(s, doc.id, f.start, f.end, f.part);
    return out;
  }
  function renderRequest() {
    if (role() === 'viewer') return '<p class="notice info">Read-only viewers cannot request leave.</p>';
    if (!ui.form) ui.form = defaultForm();
    const f = ui.form;
    const c = requestChecks(f);
    const sick = f.type === 'SICK';
    const single = f.start === f.end;
    const typeBtns = Object.entries(R.LEAVE_TYPES).map(([k, t]) => `<button type="button" class="chip" data-act="form-type" data-val="${k}" aria-pressed="${f.type === k}">${esc(t.label)}</button>`).join('');
    const docSelect = role() === 'coordinator'
      ? `<label class="field" for="f-doc">Request for<select id="f-doc" data-form="doctorId">${st().doctors.filter((d) => R.inPost(d, TODAY) || (d.start && d.start > TODAY)).map((d) => `<option value="${d.id}"${d.id === f.doctorId ? ' selected' : ''}>${esc(d.name)} (${esc(d.grade)})</option>`).join('')}</select></label>`
      : `<p class="muted">Requesting as <strong>${esc(me().name)}</strong>. Your name, email and the date are filled in for you.</p>`;
    const coverHtml = !c.suggestions || !c.suggestions.length ? '' : c.suggestions.map((s) => {
      if (!s.needed) return `<div class="small"><strong>${esc(R.fmtShort(s.date))}</strong> <span class="muted">· you're at ${esc(s.from || 'nothing rostered')}. Every post stays covered without you.</span></div>`;
      const choice = f.cover[s.date] === undefined ? (s.options.length ? 0 : -1) : f.cover[s.date];
      return `<fieldset class="stack" style="border:0;padding:0;margin:0">
        <legend class="small"><strong>${esc(R.fmtShort(s.date))}</strong> · you're at ${esc(s.from)}. Without you: <span class="cv gap">${esc(s.broken.join(', '))}</span></legend>
        ${s.options.map((o, i) => `<label class="opt"><input type="radio" name="cover-${s.date}" data-act="cover-pick" data-date="${s.date}" value="${i}"${choice === i ? ' checked' : ''}><span><strong>${esc(o.text)}</strong>${o.exceptions.length ? `<br><span class="small" style="color:var(--warn-fg)">Leaves ${esc(o.exceptions.join(', '))} with a resident only. The coordinator must agree.</span>` : ''}</span></label>`).join('')}
        <label class="opt"><input type="radio" name="cover-${s.date}" data-act="cover-pick" data-date="${s.date}" value="-1"${choice === -1 ? ' checked' : ''}><span>${s.options.length ? 'Leave it to the rota coordinator' : 'No one can move without leaving another gap. The rota coordinator will arrange cover.'}</span></label>
      </fieldset>`;
    }).join('');
    const pool = f.type === 'AL' ? 'Annual leave' : f.type === 'SL' ? 'Study leave' : null;
    const poolObj = c.bal ? (f.type === 'AL' ? c.bal.al : c.bal.sl) : null;
    return `
    <div class="page-head"><div><h1>${sick ? 'Report sickness' : 'Request leave'}</h1><p class="muted">${sick ? 'Sickness goes on the roster straight away and the rota coordinators are told. It never counts against annual leave.' : 'Leave only goes on the roster when a rota coordinator authorises it. Until then it shows as "requested".'}</p></div></div>
    ${ui.lastSubmitted ? `<p class="notice ok">${esc(ui.lastSubmitted)}</p>` : ''}
    <div class="split">
      <form class="card main" id="leave-form" novalidate>
        ${docSelect}
        <fieldset style="border:0;padding:0;margin:0" class="stack"><legend class="label" style="margin-bottom:6px">Type</legend><div class="row">${typeBtns}</div></fieldset>
        <div class="row">
          <label class="field" for="f-start">First day<input type="date" id="f-start" data-form="start" value="${esc(f.start)}"></label>
          <label class="field" for="f-end">Last day<input type="date" id="f-end" data-form="end" value="${esc(f.end)}"></label>
          ${single ? `<label class="field" for="f-part">Part of day<select id="f-part" data-form="part"><option value="FULL"${f.part === 'FULL' ? ' selected' : ''}>Full day</option><option value="AM"${f.part === 'AM' ? ' selected' : ''}>Morning only</option><option value="PM"${f.part === 'PM' ? ' selected' : ''}>Afternoon only</option></select></label>` : ''}
        </div>
        ${sick ? '' : `<div class="stack"><h2>Cover</h2><p class="muted small">Only moves that keep GUH, both day units and resident cover in place are offered.</p>${coverHtml || '<p class="muted small">No working days in these dates.</p>'}</div>`}
        <label class="field" for="f-note">Note for the coordinator <span class="hint">Optional</span><textarea id="f-note" rows="2" data-form="note">${esc(f.note)}</textarea></label>
        <div class="row"><button type="submit" class="btn primary"${c.problems.length ? ' disabled' : ''}>${sick ? 'Report sickness' : 'Send to rota coordinators'}</button></div>
      </form>
      <aside class="card side" aria-label="Worked out for you">
        <h2>Worked out for you</h2>
        <dl class="kv">
          <dt>Working days</dt><dd class="tabular">${c.days}</dd>
          ${pool && poolObj ? `<dt>${pool} left</dt><dd class="tabular">${poolObj.left - poolObj.pending} → ${poolObj.left - poolObj.pending - c.days} of ${poolObj.entitlement}</dd>` : ''}
          ${!sick && c.notice !== undefined ? `<dt>Notice</dt><dd class="tabular">${weeksDays(c.notice)}</dd>` : ''}
        </dl>
        ${c.problems.map((p) => `<p class="notice bad small">${esc(p)}</p>`).join('')}
        ${c.warnings.map((p) => `<p class="notice warn small">${esc(p)}</p>`).join('')}
        ${c.info.map((p) => `<p class="notice info small">${esc(p)}</p>`).join('')}
        ${!c.problems.length && !c.warnings.length ? '<p class="notice ok small">No problems found.</p>' : ''}
        <p class="muted small">Weekends and bank holidays are not counted. Allowances are set per person by the rota coordinators.</p>
      </aside>
    </div>`;
  }
  function submitRequest() {
    const f = ui.form;
    const c = requestChecks(f);
    if (c.problems.length) { toast(c.problems[0], true); return; }
    const cover = [];
    for (const s of c.suggestions || []) {
      if (!s.needed) continue;
      const choice = f.cover[s.date] === undefined ? (s.options.length ? 0 : -1) : f.cover[s.date];
      if (choice >= 0 && s.options[choice]) cover.push({ date: s.date, doctorId: s.options[choice].doctorId, to: s.options[choice].to, text: `${R.fmtShort(s.date)}: ${s.options[choice].text}` });
    }
    attempt(() => {
      const r = store.submitRequest({ doctorId: f.doctorId, type: f.type, start: f.start, end: f.end, part: f.start === f.end ? f.part : 'FULL', note: f.note, cover });
      ui.lastSubmitted = r.type === 'SICK'
        ? `Sickness recorded for ${rangeText(r.start, r.end)}. The rota coordinators have been told.`
        : `Sent: ${R.LEAVE_TYPES[r.type].label.toLowerCase()} for ${rangeText(r.start, r.end)} (${r.days} day${r.days === 1 ? '' : 's'}). You'll be notified when a coordinator decides.`;
      ui.form = defaultForm();
    });
  }

  // ---------- my leave ----------
  function renderMyLeave() {
    const doc = R.doctorById(st(), me().doctorId);
    if (!doc) return '<p class="notice info">This view is for doctors.</p>';
    const bal = R.balance(st(), doc);
    const mine = st().requests.filter((r) => r.doctorId === doc.id).sort((a, b) => (a.start < b.start ? 1 : -1));
    const notes = st().notifications.filter((n) => n.to === me().id);
    const balCard = (title, b) => `<article class="card"><span class="label">${title}</span><p style="font-size:1.6rem;font-weight:700" class="tabular">${b.left - b.pending} <span class="muted small" style="font-weight:400">of ${b.entitlement} days left</span></p><p class="small muted tabular">${b.used} taken or booked · ${b.pending} waiting for authorisation</p></article>`;
    return `
    <div class="page-head"><div><h1>My leave</h1><p class="muted">${esc(doc.name)} · ${esc(doc.grade)} · in post ${doc.start ? esc(R.fmtUk(doc.start)) : '?'} to ${doc.end ? esc(R.fmtUk(doc.end)) : 'no end date'}</p></div>
      <button type="button" class="btn primary" data-act="nav" data-view="request">Request leave or report sickness</button></div>
    <div class="grid-3">${balCard('Annual leave', bal.al)}${balCard('Study leave', bal.sl)}</div>
    <div class="grid-2">
      <section class="card"><h2>My requests</h2>
        <div class="scroll-x"><table class="plain"><thead><tr><th>Dates</th><th>Type</th><th>Days</th><th>Status</th><th></th></tr></thead><tbody>
        ${mine.map((r) => `<tr><td class="tabular">${esc(rangeText(r.start, r.end))}${r.part !== 'FULL' ? ` (${r.part})` : ''}</td><td>${esc(R.LEAVE_TYPES[r.type].label)}</td><td class="tabular">${r.days}</td>
          <td><span class="pill ${r.status}">${esc(r.status)}</span>${r.decidedBy && r.status !== 'pending' ? `<br><span class="small muted">${esc(r.decidedBy)}${r.reason ? ': ' + esc(r.reason) : ''}</span>` : ''}</td>
          <td>${r.status === 'pending' ? `<button type="button" class="btn ghost" data-act="cancel-req" data-id="${r.id}">Withdraw</button>` : ''}</td></tr>`).join('') || '<tr><td colspan="5" class="muted">No requests yet.</td></tr>'}
        </tbody></table></div></section>
      <section class="card"><h2>Notifications</h2>${notes.map((n) => `<div class="small"><span class="muted tabular">${esc(n.at.slice(0, 16).replace('T', ' '))}</span><br>${esc(n.text)}</div>`).join('') || '<p class="muted small">Nothing yet. You will see decisions about your leave here, and in the live version by email.</p>'}</section>
    </div>`;
  }

  // ---------- requests (coordinator) ----------
  function renderRequests() {
    const pending = st().requests.filter((r) => r.status === 'pending').sort((a, b) => (a.start < b.start ? -1 : 1));
    if (!pending.some((r) => r.id === ui.reqSel)) ui.reqSel = pending[0] ? pending[0].id : null;
    const sel = pending.find((r) => r.id === ui.reqSel);
    const decided = st().requests.filter((r) => r.status !== 'pending').sort((a, b) => ((a.decidedAt || '') < (b.decidedAt || '') ? 1 : -1)).slice(0, 12);
    const notes = st().notifications.filter((n) => n.to === 'coord').slice(0, 8);
    const flagFor = (r) => {
      const risky = R.dateRange(r.start, r.end).some((d) => { const c = R.coverWithRisk(st(), d); return c && c.rules.some((x) => x.status === 'risk'); });
      const restricted = R.restrictedIn(st(), r.start, r.end).some((x) => x.level === 'NO AL');
      return restricted ? '<span class="cv gap">Restricted date</span>' : risky ? '<span class="cv risk">Needs cover</span>' : '<span class="cv ok">Cover OK</span>';
    };
    let detail = '<section class="card main"><p class="muted">No requests are waiting. New requests appear here and the coordinators are notified.</p></section>';
    if (sel) {
      const doc = R.doctorById(st(), sel.doctorId);
      const bal = R.balance(st(), doc);
      const pool = sel.type === 'AL' ? bal.al : sel.type === 'SL' ? bal.sl : null;
      const notice = R.daysBetween(sel.submittedAt.slice(0, 10), sel.start);
      const restr = R.restrictedIn(st(), sel.start, sel.end);
      const oc = st().oncall[doc.id] || {};
      const clashes = R.dateRange(sel.start, sel.end).filter((d) => oc[d]);
      const sugg = R.suggestCover(st(), doc.id, sel.start, sel.end, sel.part);
      const coverRows = R.dateRange(sel.start, sel.end).filter((d) => !R.isWeekend(d) && !R.isBankHoliday(st(), d)).map((d) => {
        const chosen = (sel.cover || []).find((c) => c.date === d);
        const before = R.coverFor(st(), d);
        const after = R.coverFor(st(), d, { absent: [doc.id], moves: chosen ? { [chosen.doctorId]: chosen.to } : {} });
        const s = sugg.find((x) => x.date === d);
        const options = s && s.needed ? s.options : [];
        const select = options.length ? `<label class="sr-only" for="cov-${d}">Cover for ${esc(R.fmtShort(d))}</label><select id="cov-${d}" data-act="set-cover" data-date="${d}"><option value="-1">No cover move</option>${options.map((o, i) => `<option value="${i}"${chosen && chosen.doctorId === o.doctorId ? ' selected' : ''}>${esc(o.text)}${o.exceptions.length ? ' (resident-only exception)' : ''}</option>`).join('')}</select>` : `<span class="muted small">${s && !s.needed ? 'No cover needed' : 'No one can move without another gap'}</span>`;
        return `<tr><td class="tabular">${esc(R.fmtShort(d))}</td><td>${esc((s && s.from) || R.resolveCell(st(), doc, d).code)}</td><td>${select}</td>
          <td class="cover-cell">${coverChips(before, 'res')}${coverChips(before, 'mg')}</td><td class="cover-cell">${coverChips(after, 'res')}${coverChips(after, 'mg')}</td></tr>`;
      }).join('');
      detail = `<section class="card main">
        <div class="page-head"><div><span class="label">${esc(R.LEAVE_TYPES[sel.type].label)}</span><h2 style="font-size:1.4rem">${esc(doc.name)} · ${esc(doc.grade)}</h2>
          <p class="muted">${esc(rangeText(sel.start, sel.end))}${sel.part !== 'FULL' ? ` (${sel.part === 'AM' ? 'morning' : 'afternoon'})` : ''} · ${plural(sel.days, 'working day')} · submitted ${esc(sel.submittedAt.slice(0, 10).split('-').reverse().join('/'))}${sel.note ? ` · "${esc(sel.note)}"` : ''}</p></div></div>
        <dl class="kv">
          ${pool ? `<dt>${sel.type === 'AL' ? 'Annual' : 'Study'} leave</dt><dd class="tabular">${pool.left} of ${pool.entitlement} left → ${pool.left - sel.days} if authorised</dd>` : ''}
          <dt>Notice</dt><dd class="tabular" style="color:${notice < 42 ? 'var(--warn-fg)' : 'var(--ok-fg)'}">${weeksDays(notice)}${notice < 42 ? ' (short notice)' : ''}</dd>
          <dt>Restricted dates</dt><dd>${restr.length ? restr.map((r) => `<span class="restr ${r.level.replace(' ', '-')}">${esc(r.level)}</span> ${esc(rangeText(r.start, r.end))}`).join('; ') : 'None'}</dd>
          <dt>On-call clashes</dt><dd>${clashes.length ? esc(clashes.map((d) => `${R.fmtShort(d)} ${oc[d]}`).join(', ')) : 'None'}</dd>
        </dl>
        <h3>Cover, day by day</h3>
        <div class="scroll-x"><table class="plain"><thead><tr><th>Day</th><th>Usually at</th><th>Cover move</th><th>Cover now</th><th>If authorised</th></tr></thead><tbody>${coverRows || '<tr><td colspan="5" class="muted">No working days.</td></tr>'}</tbody></table></div>
        ${ui.declining ? `<div class="stack"><label class="field" for="decline-reason">Reason for declining <span class="hint">Required. Emailed to ${esc(doc.name)}.</span><textarea id="decline-reason" rows="3"></textarea></label>
          <div class="row"><button type="button" class="btn danger solid" data-act="decline-send">Send decline</button><button type="button" class="btn ghost" data-act="decline-cancel">Cancel</button></div></div>`
          : `<div class="row"><button type="button" class="btn good" data-act="authorise">Authorise and update roster</button><button type="button" class="btn danger" data-act="decline">Decline…</button></div>`}
      </section>`;
    }
    return `
    <div class="page-head"><div><h1>Requests</h1><p class="muted">Nothing reaches the roster until one of the three rota coordinators authorises it.</p></div></div>
    <div class="split">
      <section class="card side" style="padding:0;gap:0" aria-label="Waiting">
        <h2 style="padding:14px">Waiting · ${pending.length}</h2>
        <div class="list">${pending.map((r) => `<button type="button" class="list-item" data-act="pick-req" data-id="${r.id}" aria-current="${r.id === ui.reqSel}"><span style="flex:1;min-width:0"><strong>${esc(docName(r.doctorId))}</strong><br><span class="small muted">${esc(R.LEAVE_TYPES[r.type].label)} · ${esc(rangeText(r.start, r.end))}</span></span>${flagFor(r)}</button>`).join('') || '<p class="muted small" style="padding:14px">Nothing waiting.</p>'}</div>
      </section>
      ${detail}
    </div>
    <div class="grid-2">
      <section class="card"><h2>Recently decided</h2><div class="scroll-x"><table class="plain"><thead><tr><th>Doctor</th><th>Leave</th><th>Status</th><th>By</th></tr></thead><tbody>
        ${decided.map((r) => `<tr><td>${esc(docName(r.doctorId))}</td><td class="small">${esc(R.LEAVE_TYPES[r.type].label)}<br>${esc(rangeText(r.start, r.end))}</td><td><span class="pill ${r.status}">${esc(r.status)}</span></td><td class="small">${esc(r.decidedBy || '')}${r.reason ? '<br>' + esc(r.reason) : ''}</td></tr>`).join('')}
      </tbody></table></div></section>
      <section class="card"><h2>Notifications for coordinators</h2>${notes.map((n) => `<div class="small"><span class="muted tabular">${esc(n.at.slice(0, 16).replace('T', ' '))}</span><br>${esc(n.text)}</div>`).join('') || '<p class="muted small">None.</p>'}</section>
    </div>`;
  }

  // ---------- set up ----------
  function renderSetup() {
    const tabs = [['people', 'People'], ['plan', 'Weekly plan'], ['import', 'Import'], ['restrict', 'Leave restrictions'], ['access', 'Access']];
    let body = '';
    if (ui.setupTab === 'people') body = renderPeople();
    if (ui.setupTab === 'plan') body = renderPlans();
    if (ui.setupTab === 'import') body = renderImport();
    if (ui.setupTab === 'restrict') body = renderRestrictions();
    if (ui.setupTab === 'access') body = renderAccess();
    return `<div class="page-head"><div><h1>Set up</h1><p class="muted">People, their usual weekly pattern, imported on-calls and leave rules. The roster is built from these.</p></div></div>
      <div class="tabs" role="tablist">${tabs.map(([id, l]) => `<button type="button" role="tab" data-act="setup-tab" data-val="${id}" aria-selected="${ui.setupTab === id}">${l}</button>`).join('')}</div>${body}`;
  }
  function renderPeople() {
    const inp = (d, field, type, val, w) => `<input type="${type}" id="p-${d.id}-${field}" aria-label="${field} for ${esc(d.name)}" data-act="doc-field" data-doc="${d.id}" data-field="${field}" value="${esc(val)}"${w ? ` style="width:${w}"` : ''}>`;
    const rows = st().doctors.map((d) => `<tr>
      <td>${inp(d, 'name', 'text', d.name, '130px')}</td>
      <td>${inp(d, 'grade', 'text', d.grade, '130px')}</td>
      <td><select aria-label="Group for ${esc(d.name)}" id="p-${d.id}-group" data-act="doc-field" data-doc="${d.id}" data-field="group"><option value="res"${d.group === 'res' ? ' selected' : ''}>Resident</option><option value="mg"${d.group === 'mg' ? ' selected' : ''}>Middle grade</option></select></td>
      <td>${inp(d, 'start', 'date', d.start || '')}</td><td>${inp(d, 'end', 'date', d.end || '')}</td>
      <td>${inp(d, 'alEntitlement', 'number', d.alEntitlement)}</td><td>${inp(d, 'slEntitlement', 'number', d.slEntitlement)}</td>
      <td>${inp(d, 'aliases', 'text', (d.aliases || []).join(', '), '150px')}</td>
      <td><button type="button" class="btn ghost" data-act="remove-doc" data-doc="${d.id}">${ui.confirm === 'remove:' + d.id ? 'Confirm remove' : 'Remove'}</button></td></tr>`).join('');
    return `<section class="card"><p class="muted small">Annual leave days are for this post: the NHS default pro-rata to time in post, changed here by hand when needed. Every change is logged. Doctors who leave early: set the "To" date and they show as "Not in post" from the next day.</p>
      <div class="scroll-x"><table class="plain"><thead><tr><th>Name</th><th>Grade</th><th>Group</th><th>In post from</th><th>To</th><th>AL days</th><th>SL days</th><th>Other spellings in on-call sheet</th><th></th></tr></thead><tbody>${rows}
      <tr><td><input type="text" id="new-name" placeholder="New doctor" style="width:130px"></td><td><input type="text" id="new-grade" placeholder="e.g. F1" style="width:130px"></td>
      <td><select id="new-group" aria-label="Group for new doctor"><option value="res">Resident</option><option value="mg">Middle grade</option></select></td>
      <td><input type="date" id="new-start" aria-label="In post from"></td><td><input type="date" id="new-end" aria-label="In post to"></td>
      <td><input type="number" id="new-al" value="9" aria-label="AL days"></td><td><input type="number" id="new-sl" value="10" aria-label="SL days"></td><td></td>
      <td><button type="button" class="btn primary" data-act="add-doc">Add</button></td></tr></tbody></table></div>
      <p class="muted small">Suggested rotation lengths: Foundation 4 months, IMT 6 months.</p></section>`;
  }
  function renderPlans() {
    const opts = (cur) => ['<option value="">(blank)</option>'].concat(R.SITE_CODES.map((c) => `<option value="${esc(c)}"${c === cur ? ' selected' : ''}>${esc(c)}</option>`)).join('');
    const rowFor = (d, week) => `<tr><td>${week === 'A' ? `<strong>${esc(d.name)}</strong><br><span class="small muted">${esc(d.grade)}</span>` : '<span class="small muted">week B</span>'}</td>
      ${[0, 1, 2, 3, 4].map((i) => `<td><select aria-label="${esc(d.name)} week ${week} ${R.DOW_LONG[i]}" id="plan-${d.id}-${week}-${i}" data-act="plan-day" data-doc="${d.id}" data-week="${week}" data-i="${i}">${opts((d.plan[week] || [])[i] || '')}</select></td>`).join('')}
      <td>${week === 'A' ? `<label class="small"><input type="checkbox" data-act="toggle-b" data-doc="${d.id}"${d.plan.B ? ' checked' : ''}> Alternate weeks</label>` : ''}</td></tr>`;
    const docs = st().doctors.filter((d) => !d.end || d.end >= TODAY);
    return `<section class="card"><p class="muted small">Each doctor's usual place on each weekday. Imported on-calls and authorised leave sit on top of this. You can also fill it from the weekly plan sheet under Import.</p>
      <div class="scroll-x"><table class="plain"><thead><tr><th>Doctor</th>${R.DOW_LONG.slice(0, 5).map((d) => `<th>${d}</th>`).join('')}<th></th></tr></thead>
      <tbody>${docs.map((d) => rowFor(d, 'A') + (d.plan.B ? rowFor(d, 'B') : '')).join('')}</tbody></table></div></section>`;
  }
  const IMPORT_INFO = {
    oncall: ['On-call rota (from the trust)', 'Each week, put the trust on-calls into this template (list or grid layout) and import it. Earlier imports for the same dates are replaced.'],
    plan: ['Weekly plan', 'Everyone\'s usual site for each weekday. Alternating weeks are supported with a "B" row.'],
    staff: ['Staff list', 'Names, grades, dates in post and leave allowances. Existing doctors are updated by name.']
  };
  function renderImport() {
    const p = ui.importPreview;
    const card = (kind) => {
      const [title, text] = IMPORT_INFO[kind];
      let preview = '';
      if (p && p.kind === kind) {
        const unknown = Object.keys(p.parsed.unknownNames || {});
        const mapRows = unknown.map((name, i) => `<div class="row small"><span>"${esc(name)}" (${p.parsed.unknownNames[name]} row${p.parsed.unknownNames[name] === 1 ? '' : 's'}) is</span><label class="sr-only" for="map-${i}">Match ${esc(name)}</label><select id="map-${i}" data-act="map-name" data-name="${esc(name)}"><option value="">Not one of ours: skip these rows</option>${st().doctors.map((d) => `<option value="${d.id}"${p.mapping[name] === d.id ? ' selected' : ''}>${esc(d.name)} (${esc(d.grade)})</option>`).join('')}</select></div>`).join('');
        let summary = '';
        if (kind === 'oncall') {
          const usable = p.parsed.entries.filter((e) => e.doctorId || p.mapping[e.name]).length;
          const clashes = p.parsed.entries.filter((e) => { const id = e.doctorId || p.mapping[e.name]; if (!id) return false; const c = R.resolveCell(st(), R.doctorById(st(), id), e.date); return c.kind === 'leave'; }).length;
          summary = `<p><strong>${usable}</strong> on-call shifts from ${p.parsed.from ? esc(R.fmtUk(p.parsed.from)) : '?'} to ${p.parsed.to ? esc(R.fmtUk(p.parsed.to)) : '?'} (${p.parsed.format} layout, sheet "${esc(p.parsed.sheet)}"). Imported on-calls already in those dates will be replaced.</p>
            ${clashes ? `<p class="notice warn small">${plural(clashes, 'shift')} clash with authorised leave. The leave stays; the clash is marked on the roster.</p>` : ''}
            ${p.parsed.skipped.length ? `<p class="small muted">${plural(p.parsed.skipped.length, 'cell')} with site codes were ignored (sites come from the weekly plan).</p>` : ''}`;
          p.count = usable;
        } else if (kind === 'plan') {
          const usable = p.parsed.plans.filter((x) => x.doctorId || p.mapping[x.name]).length;
          summary = `<p>Weekly plans for <strong>${usable}</strong> doctors. Their current plans will be replaced.</p>`;
          p.count = usable;
        } else {
          const existing = p.parsed.staff.filter((s) => st().doctors.some((d) => d.name.toLowerCase() === s.name.toLowerCase())).length;
          summary = `<p><strong>${p.parsed.staff.length}</strong> people: ${existing} updated, ${p.parsed.staff.length - existing} new.</p>${(p.parsed.problems || []).map((x) => `<p class="small" style="color:var(--warn-fg)">${esc(x)}</p>`).join('')}`;
          p.count = p.parsed.staff.length;
        }
        const bad = (p.parsed.badCodes || []).slice(0, 8).map((b) => `Row ${b.row}: "${esc(b.code)}"`).join(', ');
        preview = `<div class="notice info stack"><strong>${esc(p.fileName)}</strong>${summary}
          ${unknown.length ? `<p class="small"><strong>New names to match once.</strong> The spelling is remembered for next time.</p>${mapRows}` : ''}
          ${bad ? `<p class="small" style="color:var(--bad-text)">Codes not recognised (these rows are skipped): ${bad}${p.parsed.badCodes.length > 8 ? ' …' : ''}</p>` : ''}
          ${p.parsed.badDates ? `<p class="small" style="color:var(--bad-text)">${plural(p.parsed.badDates, 'row')} skipped because the date could not be read. Use dd/mm/yyyy.</p>` : ''}
          <div class="row"><button type="button" class="btn primary" data-act="confirm-import"${p.count ? '' : ' disabled'}>Import</button><button type="button" class="btn ghost" data-act="cancel-import">Cancel</button></div></div>`;
      }
      return `<article class="card"><h2>${title}</h2><p class="small muted">${text}</p>
        <div class="drop"><div class="row"><button type="button" class="btn" data-act="template" data-kind="${kind}">Download template</button></div>
        <label class="field" for="file-${kind}">Choose a filled-in .xlsx file<input type="file" id="file-${kind}" accept=".xlsx" data-act="import-file" data-kind="${kind}"></label></div>${preview}</article>`;
    };
    return `<div class="grid-3">${card('oncall')}${card('plan')}${card('staff')}</div>
      <p class="muted small">On-call sheet layouts accepted: a list (Date, Doctor, Shift) or a grid like the roster (dates down the side, doctors across the top). Names are matched by exact name, a remembered spelling, or surname and initial, so "MALIK S" and "Dr Sara Malik" both match "S. Malik".</p>`;
  }
  function renderRestrictions() {
    const list = (st().settings.restricted || []).slice().sort((a, b) => (a.start < b.start ? -1 : 1));
    const mon = R.mondayOf(TODAY);
    const weeks = Array.from({ length: 10 }, (_, i) => R.addDays(mon, i * 7));
    return `<div class="grid-2">
      <section class="card"><h2>Leave restrictions</h2><p class="small muted">"NO AL" warns the doctor and flags the request in red. "AVOID" is a softer warning. Coordinators can still authorise.</p>
        <div class="scroll-x"><table class="plain"><thead><tr><th>Dates</th><th>Level</th><th>Note</th><th></th></tr></thead><tbody>
        ${list.map((r) => `<tr><td class="tabular">${esc(rangeText(r.start, r.end))}</td><td><span class="restr ${r.level.replace(' ', '-')}">${esc(r.level)}</span></td><td class="small">${esc(r.note)}</td><td><button type="button" class="btn ghost" data-act="remove-restr" data-id="${r.id}">Remove</button></td></tr>`).join('') || '<tr><td colspan="4" class="muted">None.</td></tr>'}
        </tbody></table></div>
        <div class="row" style="align-items:flex-end">
          <label class="field" for="rs-start">From<input type="date" id="rs-start"></label><label class="field" for="rs-end">To<input type="date" id="rs-end"></label>
          <label class="field" for="rs-level">Level<select id="rs-level"><option>NO AL</option><option>AVOID</option></select></label>
          <label class="field" for="rs-note">Note<input type="text" id="rs-note" placeholder="Why"></label>
          <button type="button" class="btn primary" data-act="add-restr">Add</button></div></section>
      <section class="card"><h2>Consultant of the week</h2><p class="small muted">Shown in the roster's Consultant column and the Excel copy.</p>
        ${weeks.map((w) => `<label class="row small" for="cons-${w}"><span class="tabular" style="width:110px">w/c ${esc(R.fmtUk(w))}</span><input type="text" id="cons-${w}" data-act="cons" data-week="${w}" value="${esc(st().settings.consultantOfWeek[w] || '')}"></label>`).join('')}</section></div>`;
  }
  function renderAccess() {
    const coords = st().users.filter((u) => u.role === 'coordinator');
    return `<div class="grid-2"><section class="card"><h2>Who can do what</h2>
      <table class="plain"><thead><tr><th>Role</th><th>Can</th></tr></thead><tbody>
      <tr><td><strong>Rota coordinator</strong> (3)</td><td>Edit the roster and weekly plans, import files, authorise or decline leave, set allowances and restrictions</td></tr>
      <tr><td><strong>Doctor</strong></td><td>See the roster, request leave, report sickness, withdraw their own pending requests</td></tr>
      <tr><td><strong>Read-only viewer</strong></td><td>See the roster and day view (no leave reasons)</td></tr></tbody></table>
      <p class="small muted">Coordinators now: ${coords.map((u) => esc(u.name)).join(', ')}. In the live version people sign in with their NHS email and access is given by the coordinators; the server checks every change, not just the screen.</p></section>
      <section class="card"><h2>Prototype data</h2><p class="small muted">Everything here is fictional and saved only in this browser. Reset to start again from the demo data.</p>
        <div class="row"><button type="button" class="btn danger" data-act="reset">${ui.confirm === 'reset' ? 'Confirm: reset all demo data' : 'Reset demo data'}</button></div></section></div>`;
  }

  // ---------- change log ----------
  function renderLog() {
    const q = ui.logQ.toLowerCase();
    const rows = st().audit.filter((a) => !q || (a.by + a.action + a.detail).toLowerCase().includes(q)).slice(0, 300);
    return `<div class="page-head"><div><h1>Change log</h1><p class="muted">Every change: who, when and what. Nothing is overwritten without a record.</p></div>
      <label class="field" for="log-q">Search<input type="search" id="log-q" data-live="logQ" value="${esc(ui.logQ)}" placeholder="Name, action or detail"></label></div>
      <section class="card"><div class="scroll-x"><table class="plain"><thead><tr><th>When</th><th>Who</th><th>What</th><th>Details</th></tr></thead><tbody>
      ${rows.map((a) => `<tr><td class="tabular small">${esc(a.at.slice(0, 16).replace('T', ' '))}</td><td>${esc(a.by)}</td><td>${esc(a.action)}</td><td class="small">${esc(a.detail)}</td></tr>`).join('') || '<tr><td colspan="4" class="muted">No matching entries.</td></tr>'}
      </tbody></table></div></section>`;
  }

  // ---------- how it works ----------
  function renderHow() {
    return `<div class="page-head"><div><h1>How it works</h1><p class="muted">One place for the roster, leave and cover, replacing the Excel sheet, the Microsoft Form and the Power Automate emails.</p></div></div>
    <ol class="steps">
      <li><h3>Set up once per rotation</h3><p class="small">Coordinators add the doctors with their dates in post and leave allowance, and each doctor's usual site for each weekday (the weekly plan).</p></li>
      <li><h3>Import on-calls weekly</h3><p class="small">Put the trust on-call rota into the import template and upload it. On-calls replace the weekly plan on those days and are marked ⇣.</p></li>
      <li><h3>The roster builds itself</h3><p class="small">Weekly plan + on-calls + authorised leave. The cover check flags gaps at GUH, RGH and NHH and in both day units, every day.</p></li>
      <li><h3>Doctors request leave</h3><p class="small">Days, balance, notice and restricted dates are worked out for them. The app suggests who could cover without leaving a gap.</p></li>
      <li><h3>A coordinator authorises</h3><p class="small">Only then does the leave appear on the roster, with the cover move applied. Declines need a reason. Both are logged and the doctor is told.</p></li>
      <li><h3>Excel when you need it</h3><p class="small">Download the roster in the familiar layout, with a leave log and change log, at any time.</p></li>
    </ol>
    <div class="grid-2">
      <section class="card prose"><h2>Cover rules used</h2><ul>
        <li><strong>Residents</strong> (F1, F2, IMT, JCF): 1 at GUH and 1 at RGH Monday to Friday; 1 at NHH Tuesday to Friday.</li>
        <li><strong>Middle grade</strong>: 1 on GUH ward Monday to Friday (mandatory).</li>
        <li><strong>RGH day unit</strong> Monday to Friday and <strong>NHH day unit</strong> Tuesday to Friday: a middle grade at that hospital. RGH, RGH day unit and RGH MDCU all count (the same for NHH). A resident only is an amber exception.</li>
        <li><strong>Lab</strong> is separate and never counts as cover.</li>
        <li>Sickness is self-reported, coordinators are notified, and it never comes off annual leave.</li></ul></section>
      <section class="card prose"><h2>Prototype or real?</h2><ul>
        <li>This prototype runs in your browser with fictional people. Use "Viewing as" at the top to act as a coordinator, a doctor or a ward viewer.</li>
        <li>In the live version everyone signs in with their NHS email, data is kept on a secure server, and the server enforces who can change what.</li>
        <li>Emails replace the in-app notifications, and the Excel copy can also be emailed weekly.</li>
        <li>Hosting needs approval from your health board's information governance team before real names go in.</li></ul>
        <p class="small muted">Try it: view as S. Malik and request a day of annual leave, then switch to J. Hale, open Requests and authorise it. Then look at the Roster and the Change log.</p></section>
    </div>`;
  }

  // ---------- render ----------
  function render() {
    const active = document.activeElement;
    const focusId = active && active.id;
    const sel = active && typeof active.selectionStart === 'number' ? [active.selectionStart, active.selectionEnd] : null;
    renderNav();
    const views = { roster: renderRoster, day: renderDay, request: renderRequest, myleave: renderMyLeave, requests: renderRequests, setup: renderSetup, log: renderLog, how: renderHow };
    $('#view').innerHTML = (views[ui.view] || renderRoster)();
    $('#sheet-root').innerHTML = ui.view === 'roster' ? renderEditSheet() : '';
    if (focusId) {
      const el = document.getElementById(focusId);
      if (el) { el.focus({ preventScroll: true }); if (sel && typeof el.setSelectionRange === 'function') { try { el.setSelectionRange(sel[0], sel[1]); } catch (e) { /* not a text field */ } } }
    }
  }
  store.subscribe(render);

  function go(view) {
    ui.view = view; ui.editing = null; ui.confirm = null; ui.declining = false;
    if (view !== 'request') ui.lastSubmitted = null;
    if (view === 'myleave') { store.markNotificationsRead(); }
    try { localStorage.setItem('haem-rota-view', view); } catch (e) { /* no storage */ }
    render();
    $('#view').focus({ preventScroll: false });
  }

  // ---------- events ----------
  document.addEventListener('click', (ev) => {
    const t = ev.target.closest('[data-act]');
    if (!t || t.tagName === 'SELECT' || (t.tagName === 'INPUT' && t.type !== 'button')) return;
    const act = t.dataset.act, val = t.dataset.val;
    if (act !== 'remove-doc' && act !== 'reset') ui.confirm = null;
    switch (act) {
      case 'nav': go(t.dataset.view); return;
      case 'roster-move': ui.rosterStart = R.addDays(ui.rosterStart, Number(val)); break;
      case 'roster-today': ui.rosterStart = R.mondayOf(TODAY); break;
      case 'group': ui.group = val; break;
      case 'hi': ui.hi = val; break;
      case 'export': {
        const to = R.addDays(ui.rosterStart, ui.weeks * 7 - 1);
        saveFile(`haematology-rota-${ui.rosterStart}-to-${to}.xlsx`, E.rosterWorkbook(st(), ui.rosterStart, to));
        return;
      }
      case 'edit-cell': ui.editing = { doc: t.dataset.doc, iso: t.dataset.iso }; break;
      case 'edit-close': ui.editing = null; break;
      case 'edit-save': {
        const code = $('#edit-code').value, reason = $('#edit-reason').value.trim();
        if (attempt(() => store.setOverride(ui.editing.doc, ui.editing.iso, code, reason))) { ui.editing = null; toast('Roster updated and logged.'); }
        break;
      }
      case 'edit-clear': if (attempt(() => store.setOverride(ui.editing.doc, ui.editing.iso, null, 'Undo'))) { ui.editing = null; toast('Change undone.'); } break;
      case 'pick-day': ui.day = val; break;
      case 'day-move': ui.day = R.addDays(ui.day, Number(val)); break;
      case 'form-type': ui.form.type = val; break;
      case 'pick-req': ui.reqSel = t.dataset.id; ui.declining = false; break;
      case 'authorise': if (attempt(() => store.decideRequest(ui.reqSel, 'authorised'))) toast('Authorised. The roster is updated and the doctor notified.'); break;
      case 'decline': ui.declining = true; break;
      case 'decline-cancel': ui.declining = false; break;
      case 'decline-send': {
        const reason = $('#decline-reason').value;
        if (attempt(() => store.decideRequest(ui.reqSel, 'declined', reason))) { ui.declining = false; toast('Declined. The doctor has been sent the reason.'); }
        break;
      }
      case 'cancel-req': if (attempt(() => store.cancelRequest(t.dataset.id))) toast('Request withdrawn.'); break;
      case 'setup-tab': ui.setupTab = val; ui.importPreview = null; break;
      case 'add-doc': {
        const name = $('#new-name').value.trim();
        if (!name) { toast('Type the new doctor\'s name first.', true); break; }
        attempt(() => store.saveDoctor({ name, grade: $('#new-grade').value.trim() || '?', group: $('#new-group').value, start: $('#new-start').value, end: $('#new-end').value, alEntitlement: Number($('#new-al').value) || 0, slEntitlement: Number($('#new-sl').value) || 0, email: '' }));
        break;
      }
      case 'remove-doc':
        if (ui.confirm === 'remove:' + t.dataset.doc) { ui.confirm = null; attempt(() => store.removeDoctor(t.dataset.doc)); } else ui.confirm = 'remove:' + t.dataset.doc;
        break;
      case 'template': {
        const kind = t.dataset.kind;
        const bytes = kind === 'oncall' ? E.oncallTemplate(TODAY) : kind === 'plan' ? E.planTemplate(st()) : E.staffTemplate(st());
        saveFile({ oncall: 'oncall-import-template.xlsx', plan: 'weekly-plan-template.xlsx', staff: 'staff-list-template.xlsx' }[kind], bytes);
        return;
      }
      case 'cancel-import': ui.importPreview = null; break;
      case 'confirm-import': confirmImport(); break;
      case 'add-restr': attempt(() => store.addRestricted($('#rs-start').value, $('#rs-end').value, $('#rs-level').value, $('#rs-note').value.trim())); break;
      case 'remove-restr': attempt(() => store.removeRestricted(t.dataset.id)); break;
      case 'reset':
        if (ui.confirm === 'reset') { ui.confirm = null; store.resetDemo(); ui.form = null; toast('Demo data reset.'); return; }
        ui.confirm = 'reset'; break;
      default: return;
    }
    render();
  });

  document.addEventListener('change', async (ev) => {
    const t = ev.target;
    if (t.id === 'whoami') { store.setUser(t.value); ui.form = null; ui.editing = null; ui.lastSubmitted = null; render(); return; }
    const act = t.dataset.act;
    if (t.dataset.form) { ui.form[t.dataset.form] = t.value; if (t.dataset.form === 'start' && (!ui.form.end || ui.form.end < t.value)) ui.form.end = t.value; ui.form.cover = {}; render(); return; }
    switch (act) {
      case 'weeks': ui.weeks = Number(t.value); break;
      case 'day-pick': if (t.value) ui.day = t.value; break;
      case 'cover-pick': ui.form.cover[t.dataset.date] = Number(t.value); break;
      case 'set-cover': {
        const r = st().requests.find((x) => x.id === ui.reqSel);
        const s = R.suggestCover(st(), r.doctorId, r.start, r.end, r.part).find((x) => x.date === t.dataset.date);
        const i = Number(t.value);
        const o = i >= 0 && s ? s.options[i] : null;
        attempt(() => store.setRequestCover(r.id, t.dataset.date, o ? { doctorId: o.doctorId, to: o.to, text: `${R.fmtShort(s.date)}: ${o.text}` } : null));
        return;
      }
      case 'doc-field': {
        const d = R.doctorById(st(), t.dataset.doc);
        const f = t.dataset.field;
        let v = t.value;
        if (f === 'alEntitlement' || f === 'slEntitlement') v = Number(v) || 0;
        if (f === 'aliases') v = v.split(',').map((x) => x.trim()).filter(Boolean);
        if (f === 'name' && !String(v).trim()) { toast('A name is required.', true); break; }
        attempt(() => store.saveDoctor({ id: d.id, [f]: v }));
        return;
      }
      case 'plan-day': attempt(() => store.setPlanDay(t.dataset.doc, t.dataset.week, Number(t.dataset.i), t.value)); return;
      case 'toggle-b': attempt(() => store.toggleWeekB(t.dataset.doc, t.checked)); return;
      case 'cons': attempt(() => store.setConsultant(t.dataset.week, t.value.trim())); return;
      case 'map-name': ui.importPreview.mapping[t.dataset.name] = t.value; break;
      case 'import-file': {
        const file = t.files && t.files[0];
        if (!file) return;
        if (!/\.xlsx$/i.test(file.name)) { toast('Save the file as an Excel Workbook (.xlsx) and choose it again.', true); return; }
        try {
          const sheets = await X.read(await file.arrayBuffer());
          const kind = t.dataset.kind;
          const parsed = kind === 'oncall' ? E.parseOncall(st(), sheets) : kind === 'plan' ? E.parsePlans(st(), sheets) : E.parseStaff(st(), sheets);
          ui.importPreview = { kind, fileName: file.name, parsed, mapping: {} };
        } catch (e) { toast(`Could not read ${file.name}: ${e.message}`, true); }
        break;
      }
      default: return;
    }
    render();
  });

  document.addEventListener('input', (ev) => {
    const t = ev.target;
    if (t.dataset.live) { ui[t.dataset.live] = t.value; render(); return; }
    if (t.dataset.form === 'note') ui.form.note = t.value; // no re-render while typing a note
  });

  document.addEventListener('submit', (ev) => {
    if (ev.target.id === 'leave-form') { ev.preventDefault(); submitRequest(); render(); }
  });

  document.addEventListener('keydown', (ev) => { if (ev.key === 'Escape' && ui.editing) { ui.editing = null; render(); } });

  function confirmImport() {
    const p = ui.importPreview;
    const resolve = (x) => x.doctorId || p.mapping[x.name] || null;
    const newAliases = {};
    for (const [name, id] of Object.entries(p.mapping)) if (id) newAliases[name] = id;
    const ok = attempt(() => {
      if (p.kind === 'oncall') {
        const entries = p.parsed.entries.map((e) => Object.assign({}, e, { doctorId: resolve(e) })).filter((e) => e.doctorId);
        store.applyImport('oncall', { entries, from: p.parsed.from, to: p.parsed.to, fileName: p.fileName, newAliases });
      } else if (p.kind === 'plan') {
        const plans = p.parsed.plans.map((x) => Object.assign({}, x, { doctorId: resolve(x) })).filter((x) => x.doctorId);
        store.applyImport('plan', { plans, fileName: p.fileName, newAliases });
      } else {
        store.applyImport('staff', { staff: p.parsed.staff, fileName: p.fileName, newAliases });
      }
    });
    if (ok) { toast(`Imported ${p.fileName}.`); ui.importPreview = null; }
  }

  render();
})();
