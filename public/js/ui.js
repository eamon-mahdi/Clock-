/* Screens and interactions for the Haematology Rota website. */
(function () {
  'use strict';
  const R = window.RotaRules, S = window.RotaStore, E = window.RotaExcel, X = window.RotaXlsx;
  const store = S.createStore();
  const TODAY = R.toIso(new Date());
  const narrow = () => window.matchMedia('(max-width: 760px)').matches;

  const ui = {
    view: 'rota', start: TODAY, end: R.addDays(R.addMonths(TODAY, 2), -1), scrollToIso: null, days: [], layout: narrow() ? 'list' : 'grid', group: 'both', hi: 'all', q: '', hideWeekends: false,
    editing: null, form: null, submitted: null, mine: { doctorId: '', email: '', result: null }, reqSel: null, reqFilter: 'open',
    edt: { doctorId: '', date: '', session: 'AM', reason: '' }, setupTab: 'master', nd: null, consExtra: 0, bulk: null, importPreview: null, confirm: null, logQ: '',
    rulesDraft: null, dialog: null, busy: false, scrollToToday: true, copyFrom: ''
  };
  try { const v = sessionStorage.getItem('haem-rota-view'); if (v) ui.view = v; } catch (e) { /* storage blocked */ }
  try { const m = JSON.parse(localStorage.getItem('haem-rota-mine') || 'null'); if (m) Object.assign(ui.mine, { doctorId: m.doctorId || '', email: m.email || '' }); } catch (e) { /* storage blocked */ }

  // ---------- helpers ----------
  const esc = (s) => String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const $ = (sel) => document.querySelector(sel);
  const st = () => store.state;
  const isAuth = () => !!store.me;
  const docName = (id) => (R.doctorById(st(), id) || { name: '?' }).name;
  const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;
  const rangeText = (a, b) => (a === b ? R.fmtShort(a) : `${R.fmtShort(a)} – ${R.fmtShort(b)}`);
  const typeLabel = (r) => (r.type === 'SICK' || r.type === 'UNAVAIL' ? 'Unavailable (sickness)' : r.type === 'OTHER' && r.otherText ? `Other: ${r.otherText}` : (R.LEAVE_TYPES[r.type] || R.LEAVE_TYPES.OTHER).label);
  const statusWord = { ok: 'covered', gap: 'not covered', exception: 'resident only (exception)', risk: 'at risk (leave requested)', na: 'not needed' };
  const statusMark = { ok: '✓', gap: '✕', exception: 'res', risk: '!', na: '–' };
  const reqStatus = { pending: 'Waiting', hold: 'Question asked', authorised: 'Approved', declined: 'Not approved', cancelled: 'Withdrawn' };
  const shortGrade = (d) => (d.group === 'mg' ? d.grade.replace(/^Registrar /, '').replace('Specialty doctor', 'SD').replace('Senior clinical fellow', 'SCF').replace('Specialist grade', 'SG') : d.grade.split(' ')[0]);
  const docsSorted = () => st().doctors.slice().sort((a, b) => (a.group === b.group ? a.name.localeCompare(b.name) : a.group === 'res' ? -1 : 1));
  const GRADES = { res: ['F1', 'F2', 'IMT1', 'IMT2', 'IMT3', 'JCF'], mg: ['Registrar ST3', 'Registrar ST4', 'Registrar ST5', 'Registrar ST6', 'Registrar ST7', 'Registrar ST8', 'Specialty doctor', 'Specialist grade', 'Senior clinical fellow'] };
  const GRADE_LABEL = { JCF: 'JCF (junior clinical fellow)' };
  const groupForGrade = (g) => (GRADES.res.includes(g) ? 'res' : GRADES.mg.includes(g) ? 'mg' : null);
  const gradeOptions = (sel) => ['res', 'mg'].map((g) => `<optgroup label="${g === 'res' ? 'Resident doctors' : 'Senior doctors'}">${GRADES[g].map((x) => `<option value="${esc(x)}"${x === sel ? ' selected' : ''}>${esc(GRADE_LABEL[x] || x)}</option>`).join('')}</optgroup>`).join('');
  const docOptions = (sel, filter, blank) => (blank !== undefined ? `<option value="">${esc(blank)}</option>` : '') +
    ['res', 'mg'].map((g) => `<optgroup label="${g === 'res' ? 'Resident doctors' : 'Senior doctors'}">${docsSorted().filter((d) => d.group === g && !d.placeholder && (!filter || filter(d))).map((d) => `<option value="${d.id}"${d.id === sel ? ' selected' : ''}>${esc(d.name)} (${esc(d.grade)})</option>`).join('')}</optgroup>`).join('');

  function toast(msg, isError) {
    const el = document.createElement('div');
    el.className = 'toast' + (isError ? ' error' : '');
    el.textContent = msg;
    $('#toasts').appendChild(el);
    setTimeout(() => el.remove(), isError ? 7000 : 4000);
  }
  // Run an authoriser change (on the server when hosted). Returns true when saved.
  async function change(fn, okMsg) {
    if (ui.busy) return false;
    ui.busy = true;
    try {
      const out = await store.act(fn);
      if (okMsg) toast(okMsg + (out && out.sent ? ` ${plural(out.sent, 'email')} sent.` : ''));
      return true;
    } catch (e) { toast(e.message || String(e), true); render(); return false; } finally { ui.busy = false; }
  }
  const op = (name, ...args) => (s, by) => S.ops[name](s, by, ...args);

  function saveFile(filename, bytes) {
    const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = document.createElement('a');
    a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 3000);
    toast(`Downloaded ${filename}`);
  }

  // ---------- header ----------
  function navItems() {
    const items = [['rota', 'Rota'], ['request', 'Request leave'], ['mine', 'Check my request'], ['rules', 'Rules']];
    if (isAuth()) {
      const waiting = st().requests.filter((r) => r.status === 'pending' || r.status === 'hold').length;
      items.push(['requests', 'Requests', waiting], ['edt', 'EDT / SDT'], ['setup', 'Set up'], ['log', 'Change log']);
    }
    return items;
  }
  function renderHead() {
    const items = navItems();
    if (!items.some((i) => i[0] === ui.view)) ui.view = 'rota';
    $('#nav').innerHTML = items.map(([id, label, n]) => `<button type="button" data-act="nav" data-view="${id}"${ui.view === id ? ' aria-current="page"' : ''}>${esc(label)}${n ? `<span class="badge">${n}</span>` : ''}</button>`).join('');
    const theme = document.documentElement.getAttribute('data-theme') || 'dark';
    const sw = `<div class="theme-switch" role="group" aria-label="Colour theme">${[['dark', 'Dark'], ['light', 'Light'], ['blue', 'Blue']].map(([v, l]) => `<button type="button" data-act="theme" data-val="${v}" aria-pressed="${theme === v}">${l}</button>`).join('')}</div>`;
    $('#who').innerHTML = sw + (isAuth()
      ? `<span>Authoriser: <strong>${esc(store.me)}</strong></span><button type="button" class="btn small-btn" data-act="sign-out">Sign out</button>`
      : `<button type="button" class="btn small-btn" data-act="open-signin">Authoriser sign in</button>`);
    const bar = [];
    if (store.mode === 'local') bar.push('<strong>Demo mode.</strong> This copy is not connected to the server, so changes stay in this browser. Authoriser sign-in needs no password here.');
    else if (!store.hasPassword) bar.push('<strong>First-time setup:</strong> no authoriser password has been set yet. <button type="button" class="linkish" data-act="open-setup">Set it now</button> before sharing the link.');
    if (st().settings.demo) bar.push(`Showing <strong>example data with fictional doctors</strong>.${isAuth() ? ' To start for real: Set up → Settings → Start fresh.' : ''}`);
    if (store.error) bar.push(`<span style="color:var(--bad-text)">Could not reach the server: ${esc(store.error)}</span>`);
    $('#bar').innerHTML = bar.map((b) => `<div>${b}</div>`).join('');
    $('#bar').hidden = !bar.length;
  }

  // ---------- rota: cells ----------
  function cellSite(cell) {
    if (cell.kind === 'half' || cell.kind === 'edt') return R.siteOf(cell.baseCode) || (R.kindOf(cell.baseCode) === 'lab' ? 'LAB' : 'none');
    return R.siteOf(cell.code);
  }
  function cellCategory(cell) {
    if (cell.kind === 'site' || cell.kind === 'half' || cell.kind === 'edt') return cellSite(cell);
    if (cell.kind === 'leave' || cell.kind === 'sick') return 'leave';
    return cell.kind;
  }
  function describeCell(doc, iso, cell) {
    const parts = [`${doc.name}, ${R.fmtLong(iso)}: ${cell.code || 'nothing rostered'}.`];
    if (cell.source === 'plan') parts.push('From the weekly plan.');
    if (cell.source === 'oncall') parts.push('From the trust on-call rota.');
    if (cell.source === 'override') parts.push(`Changed by an authoriser${cell.note ? ': ' + cell.note : '.'}`);
    if (cell.edt) parts.push(`${cell.edt.type} booked for the ${cell.edt.session === 'AM' ? 'morning' : 'afternoon'}.`);
    if (cell.pending) parts.push('Leave requested, waiting for an authoriser. Not on the rota until approved.');
    if (cell.conflict) parts.push(cell.conflict);
    return parts.join(' ');
  }
  function cellHtml(doc, iso, cell, canEdit) {
    const site = cellSite(cell);
    let label;
    if (cell.kind === 'gone') label = '<span class="small">not in post</span>';
    else if (cell.kind === 'edt') label = `<span>${esc(cell.baseCode)}</span><span class="emb">${esc(cell.edt.type)} ${cell.edt.session}</span>`;
    else if (cell.kind === 'half') label = `<span>${esc(cell.baseCode || 'OFF')}</span><span class="emb leave">${cell.halfOff} leave</span>`;
    else if (cell.kind === 'oncall') label = '⇣ ' + esc(cell.code);
    else if (cell.kind === 'sick') label = 'UNAVAILABLE';
    else label = esc(cell.code);
    if (cell.pending && cell.kind !== 'leave' && cell.kind !== 'sick') label += '<span class="req">leave req.</span>';
    const cls = `cell k-${cell.kind}${site ? ' s-' + site : ''}${cell.pending && cell.kind !== 'leave' && cell.kind !== 'sick' ? ' pending' : ''}`;
    const marks = (cell.source === 'override' ? '<span class="mark override" aria-hidden="true"></span>' : '') + (cell.conflict ? '<span class="mark conflict" aria-hidden="true"></span>' : '');
    const title = esc(describeCell(doc, iso, cell));
    if (canEdit && cell.kind !== 'gone') return `<button type="button" class="${cls}" data-act="edit-cell" data-doc="${doc.id}" data-iso="${iso}" title="${title}" aria-label="${title}">${label}${marks}</button>`;
    return `<span class="${cls}" title="${title}">${label}${marks}</span>`;
  }
  const coverChips = (cover, group) => (cover ? cover.rules.filter((r) => r.rule.group === group)
    .map((r) => `<span class="cv ${r.status}" title="${esc(r.rule.label)}: ${esc(statusWord[r.status])}${r.names.length ? ' (' + esc(r.names.join(', ')) + ')' : ''}">${esc(r.rule.short)} ${statusMark[r.status]}</span>`).join('') : '');

  // The rota opens on today and shows 4 months; "Show 2 more months" / "Show 2 earlier months" extend it.
  function rotaRange() { return { from: ui.start, to: ui.end }; }
  const resetRange = (iso) => { ui.start = iso; ui.end = R.addDays(R.addMonths(iso, 2), -1); };
  const moreButton = (where) => where === 'top'
    ? `<div class="more-row"><button type="button" class="btn" data-act="rota-earlier">Show 2 earlier months</button></div>`
    : `<div class="more-row"><button type="button" class="btn primary" data-act="rota-more">Show 2 more months</button><span class="small muted">Showing to ${esc(R.fmtLong(ui.end))}</span></div>`;
  // Columns (posts) for the range: a new doctor continues the column of the doctor they replace.
  function rotaColumns(from, to) {
    const q = ui.q.trim().toLowerCase();
    return R.columns(st(), from, to, (d) => ui.group === 'both' || d.group === ui.group)
      .filter((c) => !q || c.docs.some((d) => d.name.toLowerCase().includes(q) || d.grade.toLowerCase().includes(q)));
  }

  function renderRota() {
    const { from, to } = rotaRange();
    const dates = R.dateRange(from, to).filter((iso) => (ui.days.length ? ui.days.includes(R.dow(iso)) : !(ui.hideWeekends && R.isWeekend(iso))));
    const cols = rotaColumns(from, to);
    const docs = cols.flatMap((c) => c.docs);
    const changeovers = R.changeoversIn(st(), from, to);
    const counts = { gap: 0, exception: 0, risk: 0 };
    const covers = {};
    for (const iso of dates) {
      covers[iso] = R.coverWithRisk(st(), iso);
      if (covers[iso]) covers[iso].rules.forEach((r) => { if (counts[r.status] !== undefined) counts[r.status]++; });
    }
    const chip = (act, val, cur, label) => `<button type="button" class="chip" data-act="${act}" data-val="${val}" aria-pressed="${cur === val}">${label}</button>`;
    const head = `
    <div class="page-head">
      <div><h1>Rota</h1><p class="muted">${esc(R.fmtLong(from))} to ${esc(R.fmtLong(to))}. Scroll down for "Show 2 more months". ${isAuth() ? 'Click a cell to change it or book EDT/SDT. Every change is logged.' : 'Read-only. Leave shows on the rota only once an authoriser approves it.'}</p></div>
      <div class="row">
        <button type="button" class="btn" data-act="rota-today"${ui.start === TODAY && ui.end === R.addDays(R.addMonths(TODAY, 2), -1) ? ' aria-pressed="true"' : ''}>Today</button>
        <label class="field inline" for="rota-goto"><span class="sr-only">Go to date</span><input type="date" id="rota-goto" data-act="rota-goto" value="${esc(ui.start)}" title="Show 2 months from this date"></label>
        <button type="button" class="btn" data-act="export">Download Excel</button>
      </div>
    </div>
    <details class="filters"${ui.filtersOpen === undefined ? (narrow() ? '' : ' open') : ui.filtersOpen ? ' open' : ''}><summary>View and filters${ui.layout === 'list' ? ' · list' : ''}${ui.group !== 'both' ? ' · ' + (ui.group === 'res' ? 'residents' : 'senior doctors') : ''}${ui.hi !== 'all' ? ' · ' + ui.hi : ''}${ui.days.length ? ' · ' + ui.days.map((k) => R.DOW_SHORT[k]).join(', ') + ' only' : ''}${ui.q ? ' · "' + esc(ui.q) + '"' : ''}</summary>
    <div class="toolbar">
      <div class="group"><span class="label">View</span>${chip('layout', 'grid', ui.layout, 'Grid')}${chip('layout', 'list', ui.layout, 'List')}${chip('layout', 'site', ui.layout, 'By site')}</div>
      <div class="group"><span class="label">Show</span>${chip('group', 'both', ui.group, 'Everyone')}${chip('group', 'res', ui.group, 'Residents')}${chip('group', 'mg', ui.group, 'Senior doctors')}</div>
      <div class="group"><span class="label">Highlight</span>${[['all', 'All'], ['GUH', 'GUH'], ['RGH', 'RGH'], ['NHH', 'NHH'], ['leave', 'Leave'], ['oncall', 'On call'], ['edt', 'EDT/SDT']].map(([v, l]) => chip('hi', v, ui.hi, l)).join('')}</div>
      <label class="group" for="q"><span class="label">Find</span><input type="search" id="q" data-live="q" value="${esc(ui.q)}" placeholder="Name or grade"></label>
      <div class="group"><span class="label">Days</span><button type="button" class="chip" data-act="day-all" aria-pressed="${!ui.days.length}">All</button>${R.DOW_SHORT.slice(0, 5).map((d, k) => `<button type="button" class="chip" data-act="day-toggle" data-val="${k}" aria-pressed="${ui.days.includes(k)}">${d}</button>`).join('')}</div>
      ${ui.days.length ? '' : `<label class="group small"><input type="checkbox" data-act="hide-weekends"${ui.hideWeekends ? ' checked' : ''}> Hide weekends</label>`}
    </div></details>
    <div class="row small">
      <span class="cv ${counts.gap ? 'gap' : 'ok'}">${counts.gap ? plural(counts.gap, 'post') + ' not covered' : 'No gaps'}</span>
      ${counts.exception ? `<span class="cv exception">${plural(counts.exception, 'day unit')} with a resident only</span>` : ''}
      ${counts.risk ? `<span class="cv risk">${counts.risk} at risk if pending leave is approved</span>` : ''}
    </div>`;
    const body = ui.layout === 'list' ? renderRotaList(dates, docs, covers, changeovers) : ui.layout === 'site' ? renderBySite(dates, docs, changeovers) : renderRotaGrid(dates, cols, covers, changeovers);
    return head + body + legend();
  }

  function renderRotaGrid(dates, cols, covers, changeovers) {
    const res = cols.filter((c) => c.group === 'res'), mg = cols.filter((c) => c.group === 'mg');
    const showRes = ui.group !== 'mg' && res.length, showMg = ui.group !== 'res' && mg.length;
    const canEdit = isAuth();
    const first = dates[0];
    const endTxt = (d) => (d.end ? `to ${R.fmtUk(d.end).slice(0, 5)}${d.endProvisional ? '?' : ''}` : '');
    const colTh = (c) => {
      const d = R.holderOn(c, first);
      const later = c.docs.filter((x) => x.start && x.start > first);
      return `<th class="doc" scope="col"><span class="grade">${esc(d.grade)}</span>${esc(d.name)}${d.placeholder ? ' <span class="tbc">TBC</span>' : ''}<span class="dates">${esc(endTxt(d))}</span>${later.map((x) => `<span class="next">then ${esc(x.name)}</span>`).join('')}</th>`;
    };
    const allDocs = cols.flatMap((c) => c.docs);
    const hiOn = ui.hi !== 'all';
    const td = (c, iso) => {
      const d = R.holderOn(c, iso);
      const cell = R.resolveCell(st(), d, iso);
      const cat = ui.hi === 'edt' ? (cell.kind === 'edt' ? 'edt' : 'x') : ui.hi === 'oncall' ? (cell.kind === 'oncall' || cell.kind === 'other' ? 'oncall' : 'x') : cellCategory(cell);
      const dim = hiOn && cell.kind !== 'none' && cat !== ui.hi;
      return `<td class="${dim ? 'dim' : ''}">${cellHtml(d, iso, cell, canEdit)}</td>`;
    };
    const bandCells = (list, iso) => list.map((c) => {
      const d = c.docs.find((x) => x.start === iso && x !== c.docs[0]);
      return d ? `<td class="band-cell" title="${esc(d.name)} (${esc(d.grade)}) starts">→ ${esc(d.name)}<span>${esc(d.grade)}</span></td>` : '<td></td>';
    }).join('');
    const rows = dates.map((iso) => {
      const wk = R.isWeekend(iso);
      const restr = S.restrictionOn(st(), iso);
      const bh = R.isBankHoliday(st(), iso);
      const band = changeovers[iso] ? `<tr class="band"><th class="col-date" scope="row" title="${esc(changeovers[iso].join(' · '))}">Changeover<br>${esc(R.fmtShort(iso))}</th><td class="band-note">${esc(changeovers[iso].filter((t) => !t.includes('→')).join(' · '))}</td>
        ${showRes ? bandCells(res, iso) + '<td class="gap-col"></td>' : ''}${showMg ? bandCells(mg, iso) + '<td class="gap-col"></td>' : ''}<td colspan="${(showRes ? 1 : 0) + (showMg ? 1 : 0)}"></td></tr>` : '';
      return band + `<tr class="${wk ? 'weekend' : ''}${iso === TODAY ? ' today' : ''}" data-date="${iso}"${iso === TODAY ? ' id="today-row"' : ''}>
        <th class="col-date" scope="row"><span class="tabular">${esc(R.fmtShort(iso))}</span>${restr ? `<span class="restr ${restr.level.replace(' ', '-')}" title="${esc(restr.note || '')}">${esc(restr.level)}</span>` : ''}${bh ? '<span class="restr BH" title="Bank holiday">BH</span>' : ''}${wk || bh ? '' : `<br>${availChip(availableOn(iso, allDocs))}`}</th>
        <td class="cons">${wk ? '' : esc(st().settings.consultantOfWeek[R.mondayOf(iso)] || '')}</td>
        ${showRes ? res.map((c) => td(c, iso)).join('') + '<td class="gap-col"></td>' : ''}
        ${showMg ? mg.map((c) => td(c, iso)).join('') + '<td class="gap-col"></td>' : ''}
        ${showRes ? `<td class="cover-cell">${coverChips(covers[iso], 'res')}</td>` : ''}
        ${showMg ? `<td class="cover-cell">${coverChips(covers[iso], 'mg')}</td>` : ''}
      </tr>`;
    }).join('');
    if (!cols.length) return '<p class="notice info">Nobody matches these filters.</p>';
    return `<div class="table-wrap" id="rota-scroll">${moreButton('top')}
      <table class="roster${hiOn ? ' hi-on' : ''}">
        <thead>
          <tr class="groups"><th class="col-date"></th><th></th>
            ${showRes ? `<th colspan="${res.length}">Resident doctors</th><th class="gap-col"></th>` : ''}
            ${showMg ? `<th colspan="${mg.length}">Senior doctors</th><th class="gap-col"></th>` : ''}
            <th colspan="${(showRes ? 1 : 0) + (showMg ? 1 : 0)}">Cover check</th></tr>
          <tr class="names"><th class="col-date" scope="col">Date</th><th scope="col">Consultant</th>
            ${showRes ? res.map(colTh).join('') + '<th class="gap-col"></th>' : ''}
            ${showMg ? mg.map(colTh).join('') + '<th class="gap-col"></th>' : ''}
            ${showRes ? '<th scope="col">Residents</th>' : ''}${showMg ? '<th scope="col">Seniors</th>' : ''}</tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>${moreButton('bottom')}
    </div>`;
  }

  // How many doctors are working at a site (GUH/RGH/NHH or the lab) on a date, at least part of the day.
  function availableOn(iso, docs) {
    const n = { res: 0, mg: 0 };
    if (!R.isWorkday(st(), iso)) return null;
    for (const d of docs) {
      if (!R.inPost(d, iso)) continue;
      const c = R.resolveCell(st(), d, iso);
      if (c.am || c.pm || c.kind === 'lab' || (c.kind === 'edt' && R.kindOf(c.baseCode) === 'lab')) n[d.group === 'mg' ? 'mg' : 'res']++;
    }
    return n;
  }
  const availChip = (n) => (n ? `<span class="avail" title="Doctors working at a site that day (residents / senior doctors)">${ui.group !== 'mg' ? `<b>${n.res}</b> res` : ''}${ui.group === 'both' ? ' · ' : ''}${ui.group !== 'res' ? `<b>${n.mg}</b> sen` : ''}</span>` : '');

  // Who is where on a date, split by site and by grade group.
  function whoWhere(iso, docs) {
    const site = () => ({ res: [], mg: [], lab: [] });
    const out = { GUH: site(), RGH: site(), NHH: site(), ELSE: { oncall: [], leave: [], unavail: [], edt: [], other: [] }, REQ: [] };
    for (const d of docs) {
      if (!R.inPost(d, iso)) continue;
      const cell = R.resolveCell(st(), d, iso);
      const nm = `${d.name} (${shortGrade(d)})`;
      const extra = (code) => (/DAY UNIT/.test(code) ? ', day unit' : /MDCU/.test(code) ? ', MDCU' : '');
      const put = (code, note) => {
        const s = R.siteOf(code);
        if (s) out[s][d.group === 'mg' ? 'mg' : 'res'].push(nm + extra(R.norm(code)) + (note || ''));
        else if (R.kindOf(code) === 'lab') out.GUH.lab.push(nm + (note || ''));
      };
      if (cell.pending && cell.kind !== 'leave' && cell.kind !== 'sick') out.REQ.push(nm);
      if (cell.kind === 'site' || cell.kind === 'lab') put(cell.code);
      else if (cell.kind === 'edt') { put(cell.baseCode, cell.edt.session === 'AM' ? ', PM only' : ', AM only'); out.ELSE.edt.push(`${nm} ${cell.edt.type} ${cell.edt.session}`); }
      else if (cell.kind === 'half') { put(cell.baseCode, cell.halfOff === 'AM' ? ', PM only' : ', AM only'); out.ELSE.leave.push(`${nm} ${cell.halfOff}`); }
      else if (cell.kind === 'oncall') out.ELSE.oncall.push(`${nm}: ${cell.code.toLowerCase()}`);
      else if (cell.kind === 'leave') out.ELSE.leave.push(nm);
      else if (cell.kind === 'sick') out.ELSE.unavail.push(nm);
      else if (cell.kind === 'other') out.ELSE.other.push(`${nm}: ${cell.code.toLowerCase()}`);
    }
    return out;
  }
  const SITE_NOTE = { GUH: 'Ward B4', RGH: 'Day unit / MDCU', NHH: 'Day unit (Tue–Fri)' };
  const siteDim = (key) => ui.hi !== 'all' && (['GUH', 'RGH', 'NHH'].includes(ui.hi) ? ui.hi !== key : key !== 'ELSE');
  function dayParts(iso, w) {
    const wk = !R.isWorkday(st(), iso);
    const cons = wk ? '' : st().settings.consultantOfWeek[R.mondayOf(iso)] || '';
    const parts = [];
    if (!wk) for (const key of ['GUH', 'RGH', 'NHH']) {
      const x = w[key];
      const notNeeded = key === 'NHH' && R.dow(iso) === 0;
      const lines = [];
      if (ui.group !== 'mg') lines.push(['Resident doctors', x.res, notNeeded ? 'not needed on Mondays' : 'nobody']);
      if (ui.group !== 'res') lines.push(['Senior doctors', x.mg, 'nobody']);
      if (key === 'GUH' && x.lab.length && ui.group !== 'res') lines.push(['GUH lab', x.lab, '']);
      if (key === 'GUH' && cons) lines.push(['Consultant', [cons], '']);
      parts.push({ key, label: key, cls: 's-' + key, lines });
    }
    const e = w.ELSE;
    const el = [['On call', e.oncall], ['Leave', e.leave], ['Unavailable', e.unavail], ['EDT/SDT', e.edt], ['Other', e.other]].filter(([, l]) => l.length).map(([t, l]) => [t, l, '']);
    if (el.length) parts.push({ key: 'ELSE', label: 'Elsewhere', cls: 'k-oncall', lines: el });
    return parts;
  }
  function dayHead(iso, covers, docs) {
    const restr = S.restrictionOn(st(), iso);
    const n = docs ? availableOn(iso, docs) : null;
    return `<div class="day-card-head"><h3>${esc(R.fmtLong(iso))}${iso === TODAY ? ' <span class="pill authorised">Today</span>' : ''}${restr ? ` <span class="restr ${restr.level.replace(' ', '-')}">${esc(restr.level)}</span>` : ''}${R.isBankHoliday(st(), iso) ? ' <span class="restr BH">Bank holiday</span>' : ''} ${availChip(n)}</h3>
      ${covers ? `<div>${coverChips(covers[iso], 'res')} ${coverChips(covers[iso], 'mg')}</div>` : ''}</div>`;
  }
  const bandCard = (iso, changeovers) => (changeovers[iso] ? `<div class="band-card"><strong>Changeover ${esc(R.fmtShort(iso))}:</strong> ${esc(changeovers[iso].join(' · '))}</div>` : '');
  // List: one line per grade group, title in bold.
  function renderRotaList(dates, docs, covers, changeovers) {
    const days = dates.map((iso) => {
      const w = whoWhere(iso, docs);
      const parts = dayParts(iso, w);
      const body = parts.map((p) => `<div class="wl-block${siteDim(p.key) ? ' dim' : ''}"><span class="wl-site ${p.cls}">${esc(p.label)}</span><div>${p.lines.map(([t, l, empty]) => `<p class="wl"><b>${esc(t)}:</b> ${l.length ? l.map(esc).join(', ') : `<span class="muted">${esc(empty)}</span>`}</p>`).join('')}</div></div>`).join('')
        + (w.REQ.length ? `<p class="wl small"><span class="wl-site pending">Leave requested</span> ${w.REQ.map(esc).join(', ')} <span class="muted">(not approved yet)</span></p>` : '');
      return bandCard(iso, changeovers) + `<article class="day-card${!R.isWorkday(st(), iso) ? ' wk' : ''}${iso === TODAY ? ' today' : ''}" data-date="${iso}"${iso === TODAY ? ' id="today-row"' : ''}>${dayHead(iso, covers, docs)}${body || '<p class="muted small">Nothing rostered.</p>'}</article>`;
    }).join('');
    return `${moreButton('top')}<div class="day-list" id="rota-scroll">${days}</div>${moreButton('bottom')}`;
  }
  // By site: each site as a column of bullet lists.
  function renderBySite(dates, docs, changeovers) {
    const days = dates.map((iso) => {
      const parts = dayParts(iso, whoWhere(iso, docs));
      const cols = parts.map((p) => `<section class="site-col${siteDim(p.key) ? ' dim' : ''}"><h4 class="${p.cls}">${esc(p.label)}${SITE_NOTE[p.key] ? `<span>${esc(SITE_NOTE[p.key])}</span>` : ''}</h4>
        ${p.lines.map(([t, l, empty]) => `<div class="site-sec"><b>${esc(t)}</b>${l.length ? `<ul>${l.map((n) => `<li>${esc(n)}</li>`).join('')}</ul>` : `<p class="muted small">${esc(empty || 'nobody')}</p>`}</div>`).join('')}</section>`).join('');
      return bandCard(iso, changeovers) + `<article class="day-card bysite${!R.isWorkday(st(), iso) ? ' wk' : ''}${iso === TODAY ? ' today' : ''}" data-date="${iso}"${iso === TODAY ? ' id="today-row"' : ''}>${dayHead(iso, null, docs)}<div class="site-cols">${cols || '<p class="muted small">Weekend or bank holiday: nothing rostered at the sites.</p>'}</div></article>`;
    }).join('');
    return `${moreButton('top')}<div class="day-list one-col" id="rota-scroll">${days}</div>${moreButton('bottom')}`;
  }
  function legend() {
    return `<div class="legend">
      <span><span class="sw k-site s-GUH"></span>GUH</span><span><span class="sw k-site s-RGH"></span>RGH (incl. day unit, MDCU)</span><span><span class="sw k-site s-NHH"></span>NHH (incl. day unit)</span>
      <span><span class="sw k-lab"></span>GUH lab (not ward cover)</span><span><span class="sw k-oncall"></span>Leave, unavailable, ⇣ on call and anything not at GUH, RGH or NHH</span><span><span class="emb">EDT PM</span> EDT/SDT session</span>
      <span><span class="sw pending"></span>Leave requested, not approved</span><span><span class="sw k-gone"></span>Not in post</span>
      <span><span class="cv ok">✓</span>covered <span class="cv exception">res</span>resident only <span class="cv risk">!</span>at risk <span class="cv gap">✕</span>not covered</span>
    </div>`;
  }

  function renderEditSheet() {
    if (!ui.editing) return '';
    const doc = R.doctorById(st(), ui.editing.doc);
    if (!doc) return '';
    const iso = ui.editing.iso;
    const cell = R.resolveCell(st(), doc, iso);
    const isLeave = cell.kind === 'leave' || cell.kind === 'sick' || cell.kind === 'half';
    const cur = cell.kind === 'edt' ? cell.baseCode : cell.code;
    const opts = (list) => list.map((c) => `<option value="${esc(c)}"${c === cur ? ' selected' : ''}>${esc(c)}</option>`).join('');
    const edtBlock = () => {
      if (cell.edt) return `<p class="notice info small">${esc(cell.edt.type)} booked for the ${cell.edt.session === 'AM' ? 'morning' : 'afternoon'}.</p><div class="row"><button type="button" class="btn danger" data-act="edt-remove" data-id="${cell.edt.id}">Remove ${esc(cell.edt.type)}</button></div>`;
      if (!(cell.kind === 'site' || cell.kind === 'lab') || !R.isWorkday(st(), iso)) return '';
      const sess = ['AM', 'PM'].map((s) => ({ s, chk: R.edtCheck(st(), doc.id, iso, s) }));
      return `<div class="stack"><h3>Book ${esc(R.edtDefault(doc).type)}</h3>${sess.map(({ s, chk }) => `<div class="edt-opt ${chk.ok ? 'ok' : 'bad'}"><div><strong>${s === 'AM' ? 'Morning' : 'Afternoon'}</strong> ${chk.ok ? '✓ allowed' : '✕ breaks a rule'}<br><span class="small">${chk.checks.filter((c) => !c.ok).map((c) => esc(c.rule + ': ' + c.detail)).join(' ') || 'All rules met.'}</span></div>
        <button type="button" class="btn${chk.ok ? ' primary' : ''}" data-act="edt-book-cell" data-session="${s}">${chk.ok ? 'Book' : 'Book anyway…'}</button></div>`).join('')}</div>`;
    };
    return `<section class="sheet" role="dialog" aria-labelledby="sheet-title"><div class="grabber" aria-hidden="true"></div>
      <div class="row sheet-head" style="justify-content:space-between"><h3 id="sheet-title">${esc(doc.name)}, ${esc(R.fmtLong(iso))}</h3><button type="button" class="btn ghost icon" data-act="edit-close" aria-label="Close">✕</button></div>
      <p class="muted small">Now: <strong>${esc(cell.kind === 'sick' ? 'UNAVAILABLE' : cell.code || 'nothing rostered')}</strong> (${esc({ plan: 'weekly plan', oncall: 'imported on-call', override: 'changed by an authoriser', leave: 'approved leave', weekend: 'weekend', bh: 'bank holiday', none: 'blank' }[cell.source] || cell.source)})</p>
      ${cell.source === 'oncall' ? '<p class="notice warn small">This shift came from the trust on-call rota. Changing it here replaces it for this day only, and is logged.</p>' : ''}
      ${isLeave ? '<p class="notice info small">This is approved leave or sickness. To change it, open the request under Requests.</p>' : `
      <label class="field" for="edit-code">Change to
        <select id="edit-code"><option value="">(blank)</option>
          <optgroup label="Sites">${opts(R.SITE_CODES)}</optgroup>
          <optgroup label="On call">${opts(R.ONCALL_CODES)}</optgroup>
        </select></label>
      <div class="row small"><span class="muted">Quick:</span>${['GUH', 'RGH - DAY UNIT', 'NHH - DAY UNIT', 'OFF'].map((c) => `<button type="button" class="chip" data-act="quick-code" data-val="${c}">${c === 'GUH' ? 'GUH ward' : c === 'OFF' ? 'Off' : c.replace(' - DAY UNIT', ' day unit')}</button>`).join('')}</div>
      <label class="field" for="edit-reason">Reason <span class="hint">Shown in the change log</span><input type="text" id="edit-reason" placeholder="e.g. moved to cover RGH day unit"></label>
      <div class="row"><button type="button" class="btn primary" data-act="edit-save">Save change</button>
        ${cell.source === 'override' ? '<button type="button" class="btn" data-act="edit-clear">Undo change</button>' : ''}</div>
      ${edtBlock()}`}
    </section>`;
  }

  // ---------- leave request (the old form's 11 questions) ----------
  function defaultForm() {
    const start = R.addDays(TODAY, 7 * R.ruleNum(st(), 'noticeWeeks') + 1);
    return { doctorId: ui.mine.doctorId || '', email: ui.mine.email || '', type: 'AL', otherText: '', start, end: start, part: 'FULL', dayUnitCoverId: '', guhCoverId: '', note: '', intrepid: false, extenuating: false, extenuatingReason: '' };
  }
  function renderRequest() {
    if (ui.submitted) return renderSubmitted();
    const f = ui.form = ui.form || defaultForm();
    const s = st();
    const doc = R.doctorById(s, f.doctorId);
    const sick = f.type === 'SICK';
    const okDates = R.isIso(f.start) && R.isIso(f.end) && f.end >= f.start;
    const single = okDates && f.start === f.end;
    const part = single ? f.part : 'FULL';
    const days = okDates ? R.leaveDays(s, f.start, f.end, part) : 0;
    const check = doc && okDates ? R.staffingCheck(s, { doctorId: doc.id, start: f.start, end: f.end, part, dayUnitCoverId: f.dayUnitCoverId, guhCoverId: f.guhCoverId }) : null;
    const notes = [];
    const need = R.ruleNum(s, 'noticeWeeks');
    const short = okDates && f.type === 'AL' && R.shortNotice(s, f.start, TODAY);
    if (okDates && !sick) {
      if (!short && f.type !== 'AL' && R.daysBetween(TODAY, f.start) < need * 7) notes.push(['warn', `This is less than ${need} weeks' notice. You can still ask.`]);
      R.restrictedIn(s, f.start, f.end).forEach((r) => notes.push([r.level === 'NO AL' ? 'bad' : 'warn', `${rangeText(r.start, r.end)} is marked ${r.level}${r.note ? ': ' + r.note : ''}. ${r.level === 'NO AL' ? 'Annual leave is not normally approved then.' : 'Please avoid these dates if you can.'}`]));
    }
    if (doc && okDates && !sick) {
      const clash = s.requests.find((r) => r.doctorId === doc.id && ['pending', 'hold', 'authorised'].includes(r.status) && r.start <= f.end && r.end >= f.start);
      if (clash) notes.push(['bad', `You already have leave ${clash.status === 'authorised' ? 'approved' : 'requested'} for ${rangeText(clash.start, clash.end)}.`]);
    }
    const needs = check ? { guh: check.days.some((d) => /GUH/.test(d.at) && R.siteOf(d.at) === 'GUH'), du: check.days.some((d) => ['RGH', 'NHH'].includes(R.siteOf(d.at.toUpperCase()))) } : { guh: true, du: true };
    const q = (n, label, inner, hint) => `<div class="q"><div class="qn">${n}</div><div class="qb"><div class="ql">${label}${hint ? `<span class="hint">${hint}</span>` : ''}</div>${inner}</div></div>`;
    const typeOpt = (v, l) => `<label class="opt"><input type="radio" name="ltype" data-form="type" value="${v}"${f.type === v ? ' checked' : ''}> <span>${l}</span></label>`;
    const coverSel = (field, id) => `<select id="${id}" data-form="${field}"><option value="">${sick ? 'Not known' : 'Nobody needed / not arranged'}</option>${docOptions(f[field], (d) => d.id !== f.doctorId && (!okDates || R.inPost(d, f.start) || R.inPost(d, f.end)))}</select>`;
    const dayRow = (d) => `<tr><td class="tabular">${esc(R.fmtShort(d.date))}</td><td>${esc(d.at)}</td><td><span class="cv ${d.status === 'ok' ? 'ok' : d.status === 'warn' ? 'exception' : 'gap'}">${d.status === 'ok' ? '✓' : d.status === 'warn' ? '!' : '✕'}</span> ${esc(d.text)}${d.problems.length ? `<br><span class="small muted">${esc(d.problems.join(' '))}</span>` : ''}</td></tr>`;
    return `
    <div class="page-head"><div><h1>${sick ? 'Report sickness' : 'Request leave'}</h1><p class="muted">No password needed. ${sick ? 'Sickness is recorded straight away and the authorisers are told. It shows on the rota only as "Unavailable".' : `Leave is not on the rota until an authoriser approves it. See the decision under "Check my request"${st().settings.emailDoctors ? ' (you\'ll also get an email)' : ''}.`}</p></div></div>
    <form id="leave-form" class="split" novalidate>
      <div class="main card qform">
        ${q(1, 'Date of request', `<p class="tabular"><strong>${esc(R.fmtLong(TODAY))}</strong></p>`)}
        ${q(2, 'Requestor', `<select id="f-doc" data-form="doctorId" required><option value="">Choose your name</option>${docOptions(f.doctorId, (d) => !d.end || d.end >= TODAY)}</select>`, 'Your own name, from the staff list')}
        ${q(3, 'NHS email', `<input type="email" id="f-email" data-form="email" value="${esc(f.email)}" placeholder="name@wales.nhs.uk" autocomplete="email" required>`, 'Must match the email the authorisers hold for you; it is how we know it is you. Your remaining leave is worked out for you: see "Check my request".')}
        ${q(4, 'Leave type', `<div class="opts">${typeOpt('AL', 'Annual leave')}${typeOpt('SL', 'Study leave')}${typeOpt('SICK', 'Sick leave')}${typeOpt('OTHER', 'Other')}</div>
          ${f.type === 'OTHER' ? '<label class="field" for="f-other">Type of leave<input type="text" id="f-other" data-form="otherText" value="' + esc(f.otherText) + '" placeholder="e.g. parental, compassionate, exam"></label>' : ''}
          ${short ? `<div class="notice bad small stack"><strong>This is less than ${need} weeks' notice.</strong> Trust policy is that annual leave is requested at least ${need} weeks ahead. It can only be considered in extenuating circumstances.
            <label class="opt"><input type="checkbox" id="f-ext" data-form="extenuating"${f.extenuating ? ' checked' : ''}> <span>There are extenuating circumstances</span></label>
            ${f.extenuating ? `<label class="field" for="f-ext-why">Explain the circumstances <span class="hint">Seen only by the authorisers</span><textarea id="f-ext-why" data-form="extenuatingReason" rows="2">${esc(f.extenuatingReason)}</textarea></label>` : ''}</div>` : ''}
          ${f.type === 'SL' ? `<label class="opt"><input type="checkbox" id="f-intrepid" data-form="intrepid"${f.intrepid ? ' checked' : ''}> <span>I have also requested this study leave through <strong>Intrepid</strong> (or CODI).</span></label>` : ''}`)}
        ${q(5, 'Start', `<input type="date" id="f-start" data-form="start" value="${esc(f.start)}" required>`)}
        ${q(6, 'End', `<input type="date" id="f-end" data-form="end" value="${esc(f.end)}" min="${esc(f.start)}" required>
          ${single ? `<div class="opts" role="radiogroup" aria-label="Part of day">${[['FULL', 'Whole day'], ['AM', 'Morning only'], ['PM', 'Afternoon only']].map(([v, l]) => `<label class="opt"><input type="radio" name="lpart" data-form="part" value="${v}"${f.part === v ? ' checked' : ''}> <span>${l}</span></label>`).join('')}</div>` : ''}`)}
        ${q(7, 'Total days', `<p><strong class="tabular">${days}</strong> working day${days === 1 ? '' : 's'} <span class="muted small">(weekends and bank holidays not counted)</span></p>`)}
        ${q(8, 'Doctor covering (day unit)', coverSel('dayUnitCoverId', 'f-du'), needs.du ? 'A named doctor, not a grade. Needed if you are at RGH or NHH on any of these days.' : 'You are not at a day unit on these days.')}
        ${q(9, 'Doctor covering (GUH)', coverSel('guhCoverId', 'f-guh'), needs.guh ? 'A named doctor, not a grade. Needed if you are at GUH on any of these days.' : 'You are not at GUH on these days.')}
        ${q(10, 'Additional notes', `<textarea id="f-note" data-form="note" rows="3" placeholder="${sick ? 'Optional. Do not include medical details.' : 'Anything the authorisers should know'}">${esc(f.note)}</textarea>`)}
        <div class="row"><button type="submit" class="btn primary"${ui.busy ? ' disabled' : ''}>${sick ? 'Report sickness' : 'Send request'}</button><button type="button" class="btn ghost" data-act="form-reset">Clear form</button></div>
      </div>
      <aside class="side stack">
        <section class="card"><h2>Will there be enough doctors?</h2>
          ${!doc ? '<p class="muted small">Choose your name and dates to see the check.</p>' : !okDates ? '<p class="muted small">Choose valid dates.</p>' : !check.days.length ? '<p class="muted small">No working days in this range.</p>' : `
          <p class="small">${check.bad ? `<strong style="color:var(--bad-text)">Not enough doctors on ${plural(check.bad, 'day')}.</strong> You can still send it: the authorisers will see this and may approve the other days, suggest another date or arrange cover.` : check.warn ? `<strong>Enough doctors, with ${plural(check.warn, 'warning')}.</strong>` : '<strong style="color:var(--ok-fg)">Enough doctors on every day.</strong>'}</p>
          <div class="scroll-x"><table class="plain small"><thead><tr><th>Day</th><th>You are</th><th>Check</th></tr></thead><tbody>${check.days.map(dayRow).join('')}</tbody></table></div>
          ${check.alternatives.length ? `<p class="small">Nearby days with enough doctors: ${check.alternatives.map((d) => esc(R.fmtShort(d))).join(', ')}.</p>` : ''}`}
        </section>
        ${notes.map(([cls, t]) => `<p class="notice ${cls} small">${esc(t)}</p>`).join('')}
      </aside>
    </form>`;
  }
  function renderSubmitted() {
    const { request: r, check, emailed } = ui.submitted;
    const sick = r.type === 'SICK' || r.type === 'UNAVAIL';
    return `<div class="page-head"><div><h1>${sick ? 'Sickness recorded' : 'Request sent'}</h1></div></div>
      <section class="card" style="max-width:720px">
        <p class="notice ok">${sick ? `Recorded for ${esc(rangeText(r.start, r.end))}. The authorisers have been told. It shows on the rota as "Unavailable" and is not taken from annual leave.` : `Your request for ${esc(rangeText(r.start, r.end))} (${plural(r.days, 'working day')}) has been sent to the rota authorisers.`}</p>
        ${check && check.bad && !sick ? `<p class="notice warn small">There may not be enough doctors on ${plural(check.bad, 'day')}. The authorisers will look at this; they may approve the other days or suggest another date.</p>` : ''}
        <p class="small muted">${emailed ? 'A confirmation email is on its way. ' : ''}${sick ? '' : 'See the decision under "Check my request". It is not on the rota until approved.'}</p>
        <div class="row"><button type="button" class="btn primary" data-act="nav" data-view="mine">Check my requests</button><button type="button" class="btn" data-act="form-new">Send another</button></div>
      </section>`;
  }
  async function submitForm() {
    const f = ui.form;
    if (!f.doctorId) return toast('Choose your name (question 2).', true);
    if (!f.email.trim()) return toast('Enter your NHS email (question 3).', true);
    const single = f.start === f.end;
    const payload = Object.assign({}, f, { part: single ? f.part : 'FULL', email: f.email.trim() });
    ui.busy = true; render();
    try {
      const out = await store.submitLeave(payload);
      ui.mine = { doctorId: f.doctorId, email: f.email.trim(), result: null };
      try { localStorage.setItem('haem-rota-mine', JSON.stringify({ doctorId: ui.mine.doctorId, email: ui.mine.email })); } catch (e) { /* storage blocked */ }
      ui.submitted = { request: Object.assign({ days: out.request.days }, out.request, { type: f.type }), check: out.check, emailed: out.emailed };
      ui.form = null;
    } catch (e) { toast(e.message, true); } finally { ui.busy = false; render(); window.scrollTo(0, 0); }
  }

  // ---------- check my request ----------
  function renderMine() {
    const m = ui.mine;
    const res = m.result;
    const card = (r) => `<article class="card req-card">
      <div class="row" style="justify-content:space-between"><h3>${esc(typeLabel(r))}: ${esc(rangeText(r.start, r.end))}${r.part !== 'FULL' ? ` (${r.part === 'AM' ? 'morning' : 'afternoon'})` : ''}</h3><span class="pill ${r.status}">${esc(reqStatus[r.status] || r.status)}</span></div>
      <p class="small muted">${plural(r.days, 'working day')} · sent ${esc(R.fmtShort((r.submittedAt || '').slice(0, 10)))}${r.decidedBy ? ` · ${r.status === 'hold' ? 'question from' : 'decided by'} ${esc(r.decidedBy)}` : ''}</p>
      ${(r.excludedDates || []).length ? `<p class="notice warn small">Not approved on ${esc(r.excludedDates.map(R.fmtShort).join(', '))}.</p>` : ''}
      ${r.message ? `<p class="notice ${r.status === 'declined' ? 'bad' : 'info'} small">"${esc(r.message)}"</p>` : ''}
      ${(r.cover || []).length ? `<p class="small">Cover: ${esc(r.cover.map((c) => c.text).join('; '))}</p>` : ''}
      ${['pending', 'hold'].includes(r.status) ? `<div class="row"><button type="button" class="btn danger" data-act="withdraw" data-id="${r.id}">${ui.confirm === 'wd:' + r.id ? 'Confirm: withdraw' : 'Withdraw request'}</button></div>` : ''}
    </article>`;
    return `<div class="page-head"><div><h1>Check my request</h1><p class="muted">See your leave requests and decisions. Enter the same name and NHS email you used on the form.</p></div></div>
      <form id="mine-form" class="toolbar">
        <label class="field" for="m-doc">Name<select id="m-doc" data-mine="doctorId"><option value="">Choose your name</option>${docOptions(m.doctorId)}</select></label>
        <label class="field" for="m-email">NHS email<input type="email" id="m-email" data-mine="email" value="${esc(m.email)}" autocomplete="email"></label>
        <button type="submit" class="btn primary">Show my requests</button>
      </form>
      ${res && res.balance ? `<section class="card bal-card"><h2>Your leave</h2><div class="bal-grid">
        ${[['Annual leave', res.balance.al], ['Study leave', res.balance.sl]].map(([t, b]) => `<div class="bal"><span class="bal-n">${b.left}</span><span class="bal-t">${t} days left</span><span class="small muted">${b.entitlement} allowance · ${b.used} approved · ${b.pending} waiting</span></div>`).join('')}</div>
        <p class="small muted">Annual leave allowance: ${esc(res.balance.alBasis)}, including any carried-over days added by the authorisers. Ask them if this looks wrong.</p></section>` : ''}
      ${res ? `<div class="grid-2"><div class="stack">${res.requests.map(card).join('') || '<p class="muted">No requests yet.</p>'}</div>
        <section class="card"><h2>Messages</h2>${res.notifications.map((n) => `<div class="small"><span class="muted tabular">${esc(n.at.slice(0, 10))}</span> ${esc(n.text)}</div>`).join('') || '<p class="muted small">None yet.</p>'}</section></div>` : ''}`;
  }
  async function loadMine() {
    if (!ui.mine.doctorId || !ui.mine.email.trim()) return toast('Choose your name and enter your NHS email.', true);
    try {
      ui.mine.result = await store.myRequests(ui.mine.doctorId, ui.mine.email.trim());
      try { localStorage.setItem('haem-rota-mine', JSON.stringify({ doctorId: ui.mine.doctorId, email: ui.mine.email.trim() })); } catch (e) { /* storage blocked */ }
    } catch (e) { ui.mine.result = null; toast(e.message, true); }
    render();
  }

  // ---------- rules tab ----------
  const RULE_FIELDS = [
    ['minResGUH', 'Residents at GUH, Mon–Fri', 'minimum'], ['minResRGH', 'Residents at RGH, Mon–Fri', 'minimum'], ['minResNHH', 'Residents at NHH, Tue–Fri', 'minimum'],
    ['minSenGUH', 'Senior doctors on GUH ward, Mon–Fri', 'minimum'], ['minSenRGH', 'Senior doctor, RGH day unit, Mon–Fri (0 = not compulsory)', 'minimum'], ['minSenNHH', 'Senior doctor, NHH day unit, Tue–Fri (0 = not compulsory)', 'minimum'],
    ['edtOthersAtGUH', 'Other residents at GUH for a resident\'s EDT/SDT', 'per session'], ['noticeWeeks', 'Ask for leave ahead', 'weeks']
  ];
  function checkedRulesText() {
    const n = (k) => R.ruleNum(st(), k);
    const atLeast = (k, who) => (n(k) > 0 ? `at least ${n(k) === 1 ? 'one' : n(k)} ${who}${n(k) === 1 ? '' : 's'}` : `no minimum for ${who}s`);
    return [
      { title: 'Minimum cover: resident doctors (previously "SHO")', items: [
        `Grange University Hospital (GUH): ${atLeast('minResGUH', 'resident doctor')} every weekday (Monday to Friday).`,
        `Royal Gwent Hospital (RGH): ${atLeast('minResRGH', 'resident doctor')} every weekday (Monday to Friday), for day unit or clinic cover.`,
        `Nevill Hall Hospital (NHH): ${atLeast('minResNHH', 'resident doctor')}, Tuesday to Friday, for day unit or clinic cover.`,
        'Residents are F1, F2, IMT and junior clinical fellows (JCF).'] },
      { title: 'Minimum cover: senior doctors', items: [
        'Senior doctors are registrars, specialty doctors, specialist grade doctors and senior clinical fellows, all counted together.',
        `GUH ward: ${atLeast('minSenGUH', 'senior doctor')} every weekday. This is compulsory.`,
        n('minSenRGH') || n('minSenNHH') ? `RGH day unit (Monday to Friday): ${atLeast('minSenRGH', 'senior doctor')}; NHH day unit (Tuesday to Friday): ${atLeast('minSenNHH', 'senior doctor')}. A resident covering alone is flagged.` : 'A senior doctor at the RGH or NHH day unit is not compulsory. Senior doctors can be put on a day unit when needed (the authorisers change the rota).',
        'Day-unit cover counts anyone at RGH, RGH - DAY UNIT or RGH MDCU (and the same for NHH). The lab is separate and never counts as cover.'] },
      { title: 'Leave and EDT/SDT checks', items: [
        `Annual leave must be requested at least ${plural(n('noticeWeeks'), 'week')} ahead (trust policy). Shorter notice is only considered in extenuating circumstances, which must be explained on the form.`,
        `A resident's EDT/SDT needs at least ${n('edtOthersAtGUH')} other residents at GUH that session.`,
        'Each morning and afternoon is checked separately, so half-day leave and EDT/SDT are counted properly.'] }
    ];
  }
  function renderRules() {
    const s = st().settings;
    const d = ui.rulesDraft;
    const tag = (t) => `<span class="tag ${t === 'Checked by the app' ? 'blue' : ''}">${t}</span>`;
    const section = (sec, t, idx) => `<section class="card rule-sec"><div><h2>${esc(sec.title)}</h2>${tag(t)}</div><ul>${sec.items.map((i) => `<li>${esc(i)}</li>`).join('')}</ul></section>`;
    const last = st().audit.find((a) => /Rules/.test(a.action));
    if (d) {
      return `<div class="page-head"><div><h1>Edit rules</h1><p class="muted">Numbers change the app's checks straight away. Text is what everyone reads. Each save is logged.</p></div>
        <div class="row"><button type="button" class="btn good" data-act="rules-save">Save changes</button><button type="button" class="btn ghost" data-act="rules-cancel">Cancel</button></div></div>
        <section class="card"><h2>Numbers the app checks</h2><div class="grid-3">${RULE_FIELDS.map(([k, l, u]) => `<label class="field box" for="rn-${k}">${esc(l)}<span class="row"><input type="number" min="0" max="9" id="rn-${k}" data-rule="${k}" value="${esc(d.rules[k])}"><span class="muted" style="font-weight:400">${u}</span></span></label>`).join('')}</div></section>
        ${d.text.map((sec, i) => `<section class="card"><div class="row"><label class="field" style="flex:1" for="rt-${i}">Heading<input type="text" id="rt-${i}" data-rtitle="${i}" value="${esc(sec.title)}"></label><button type="button" class="btn ghost" data-act="rules-del" data-i="${i}">Remove section</button></div>
          <label class="field" for="ri-${i}">Points <span class="hint">One per line</span><textarea id="ri-${i}" data-ritems="${i}" rows="${Math.max(3, sec.items.length + 1)}">${esc(sec.items.join('\n'))}</textarea></label></section>`).join('')}
        <div class="row"><button type="button" class="btn" data-act="rules-add">Add a section</button></div>`;
    }
    const text = s.rulesText || [];
    return `<div class="page-head"><div><h1>Rules</h1><p class="muted">How the haematology rota works. Everyone can read this. Only authorisers can change it.${last ? ` Last changed by ${esc(last.by)}, ${esc(last.at.slice(0, 10))}.` : ''}</p></div>
      ${isAuth() ? '<button type="button" class="btn" data-act="rules-edit">Edit rules</button>' : ''}</div>
      ${text.slice(0, 1).map((x) => section(x, 'From the placement email')).join('')}
      ${checkedRulesText().map((x) => section(x, 'Checked by the app')).join('')}
      ${text.slice(1).map((x) => section(x, 'From the placement email')).join('')}`;
  }

  // ---------- authoriser: requests ----------
  function renderRequests() {
    const s = st();
    const order = { pending: 0, hold: 1, authorised: 2, declined: 3, cancelled: 4 };
    let list = s.requests.slice();
    if (ui.reqFilter === 'open') list = list.filter((r) => r.status === 'pending' || r.status === 'hold');
    list.sort((a, b) => (order[a.status] - order[b.status]) || (ui.reqFilter === 'open' ? (a.start < b.start ? -1 : 1) : (a.submittedAt < b.submittedAt ? 1 : -1)));
    list = list.slice(0, 200);
    if (!ui.reqSel || !list.some((r) => r.id === ui.reqSel)) ui.reqSel = list[0] ? list[0].id : null;
    const unread = s.notifications.filter((n) => n.to === 'auth' && !n.read);
    const item = (r) => {
      const chk = r.status === 'pending' || r.status === 'hold' ? R.staffingCheck(s, { doctorId: r.doctorId, start: r.start, end: r.end, part: r.part, dayUnitCoverId: r.dayUnitCoverId, guhCoverId: r.guhCoverId, requestId: r.id }) : null;
      return `<button type="button" class="list-item" data-act="pick-req" data-id="${r.id}"${r.id === ui.reqSel ? ' aria-current="true"' : ''}>
        <span style="flex:1;min-width:0"><strong>${esc(docName(r.doctorId))}</strong> <span class="muted small">${esc(typeLabel(r))}</span><br><span class="small tabular">${esc(rangeText(r.start, r.end))}</span></span>
        ${chk && chk.bad ? `<span class="cv gap">✕ ${chk.bad}</span>` : chk && chk.warn ? `<span class="cv exception">! ${chk.warn}</span>` : ''}
        <span class="pill ${r.status}">${esc(reqStatus[r.status])}</span></button>`;
    };
    return `<div class="page-head"><div><h1>Requests</h1><p class="muted">Approve, part-approve, ask a question or decline. The doctor sees the decision under "Check my request"${st().settings.emailDoctors ? ' and by email' : ''}.</p></div>
      <div class="seg">${[['open', 'Waiting'], ['all', 'All']].map(([v, l]) => `<button type="button" class="chip" data-act="req-filter" data-val="${v}" aria-pressed="${ui.reqFilter === v}">${l}</button>`).join('')}</div></div>
      ${unread.length ? `<section class="notice info stack"><div class="row" style="justify-content:space-between"><strong>${plural(unread.length, 'new alert')}</strong><button type="button" class="btn small-btn" data-act="mark-read">Mark as read</button></div>${unread.slice(0, 8).map((n) => `<div class="small"><span class="muted tabular">${esc(n.at.slice(0, 16).replace('T', ' '))}</span> ${esc(n.text)}</div>`).join('')}</section>` : ''}
      <div class="split"><section class="side card" style="padding:0;gap:0;overflow:hidden"><div class="list">${list.map(item).join('') || '<p class="muted" style="padding:14px">Nothing waiting.</p>'}</div></section>
      <section class="main">${ui.reqSel ? renderRequestDetail(s.requests.find((r) => r.id === ui.reqSel)) : ''}</section></div>`;
  }
  function renderRequestDetail(r) {
    const s = st();
    const doc = R.doctorById(s, r.doctorId) || { name: '?', grade: '' };
    const open = r.status === 'pending' || r.status === 'hold';
    const chk = R.staffingCheck(s, { doctorId: r.doctorId, start: r.start, end: r.end, part: r.part, dayUnitCoverId: r.dayUnitCoverId, guhCoverId: r.guhCoverId, requestId: r.id });
    const bal = doc.id ? R.balance(s, doc) : null;
    const restr = R.restrictedIn(s, r.start, r.end);
    const notice = R.daysBetween((r.submittedAt || TODAY).slice(0, 10), r.start) / 7;
    const sick = r.type === 'SICK';
    const balTxt = (b, name) => `${name}: ${b.entitlement} days, ${b.used} approved, ${b.pending} waiting, <strong>${b.left} left</strong>`;
    const dayRow = (d) => `<tr class="${d.status}">${open ? `<td><input type="checkbox" class="day-tick" value="${d.date}"${d.status !== 'bad' ? ' checked' : ''} aria-label="Approve ${esc(R.fmtShort(d.date))}"></td>` : `<td>${(r.excludedDates || []).includes(d.date) ? '✕' : r.status === 'authorised' ? '✓' : ''}</td>`}<td class="tabular">${esc(R.fmtShort(d.date))}</td><td>${esc(d.at)}</td>
      <td><span class="cv ${d.status === 'ok' ? 'ok' : d.status === 'warn' ? 'exception' : 'gap'}">${d.status === 'ok' ? '✓' : d.status === 'warn' ? '!' : '✕'}</span> ${esc(d.text)}
        ${d.problems.length ? `<br><span class="small">${esc(d.problems.join(' '))}</span>` : ''}
        ${d.whoElse ? `<br><span class="small muted">Others that day: ${esc(d.whoElse)}</span>` : ''}
        ${d.options && d.options.length ? `<br><span class="small"><strong>Options:</strong> ${esc(d.options.join('; '))}</span>` : ''}
        ${d.competing ? `<br><span class="small muted">Also asked for this day: ${esc(d.competing.map((c) => `${c.name} (sent ${(c.submittedAt || '').slice(0, 10)})`).join(', '))}</span>` : ''}</td></tr>`;
    return `<article class="card">
      <div class="row" style="justify-content:space-between"><div><h2>${esc(doc.name)} <span class="muted small">${esc(doc.grade)}</span></h2><p class="small muted">${esc(typeLabel(r))} · ${esc(rangeText(r.start, r.end))}${r.part !== 'FULL' ? ` (${r.part})` : ''} · ${plural(r.days, 'working day')}</p></div><span class="pill ${r.status}">${esc(reqStatus[r.status])}</span></div>
      ${open && chk.bad ? `<p class="notice bad"><strong>Not enough doctors on ${plural(chk.bad, 'day')}.</strong> Those days are unticked below. You can approve the other days, arrange cover first (change the rota, then come back), ask the doctor a question, or tick them and give a reason (e.g. locum booked).${chk.alternatives.length ? ` Nearby days with enough doctors: ${esc(chk.alternatives.map(R.fmtShort).join(', '))}.` : ''}</p>` : ''}
      ${open && !chk.bad && chk.warn ? `<p class="notice warn">Enough doctors, with ${plural(chk.warn, 'warning')}. See below.</p>` : ''}
      <dl class="kv">
        <dt>1 Date of request</dt><dd>${esc((r.submittedAt || '').slice(0, 10) ? R.fmtLong(r.submittedAt.slice(0, 10)) : '')}${!sick && notice < R.ruleNum(s, 'noticeWeeks') ? ` <span class="cv exception">${Math.max(0, Math.floor(notice))} weeks' notice</span>` : ''}</dd>
        <dt>2 Requestor</dt><dd>${esc(doc.name)} (${esc(doc.grade)})</dd>
        <dt>3 NHS email</dt><dd>${esc(r.email || doc.email || '')}</dd>
        <dt>4 Leave type</dt><dd>${esc(typeLabel(r))}${r.type === 'SL' ? (r.intrepid ? ' · Intrepid/CODI confirmed' : ' · <span style="color:var(--bad-text)">Intrepid not confirmed</span>') : ''}</dd>
        <dt>5–6 Dates</dt><dd>${esc(R.fmtLong(r.start))}${r.end !== r.start ? ' to ' + esc(R.fmtLong(r.end)) : ''}${r.part !== 'FULL' ? ` (${r.part === 'AM' ? 'morning' : 'afternoon'})` : ''}</dd>
        <dt>7 Total days</dt><dd>${r.days}</dd>
        <dt>8 Cover (day unit)</dt><dd>${esc(r.dayUnitCoverId ? docName(r.dayUnitCoverId) : 'None named')}</dd>
        <dt>9 Cover (GUH)</dt><dd>${esc(r.guhCoverId ? docName(r.guhCoverId) : 'None named')}</dd>
        <dt>Leave left</dt><dd>${bal && !sick ? `<span class="small" style="font-weight:400">${balTxt(bal.al, 'Annual')}; ${balTxt(bal.sl, 'Study')}</span>` : '—'}</dd>
        ${r.shortNotice ? `<dt>Short notice</dt><dd><span class="cv gap">Less than ${R.ruleNum(s, 'noticeWeeks')} weeks</span> <span style="font-weight:400">Extenuating circumstances: "${esc(r.extenuatingReason || '')}"</span></dd>` : ''}
        <dt>10 Notes</dt><dd style="font-weight:400">${esc(r.note || '—')}</dd>
      </dl>
      ${restr.length ? restr.map((x) => `<p class="notice ${x.level === 'NO AL' ? 'bad' : 'warn'} small">${esc(rangeText(x.start, x.end))} is marked ${esc(x.level)}${x.note ? ': ' + esc(x.note) : ''}.</p>`).join('') : ''}
      ${chk.days.length ? `<h3>Staffing check, day by day</h3><div class="scroll-x"><table class="plain small check"><thead><tr><th>${open ? 'Approve' : ''}</th><th>Day</th><th>Usually</th><th>Result</th></tr></thead><tbody>${chk.days.map(dayRow).join('')}</tbody></table></div>${open ? '<p class="small muted">Untick any day you don\'t want to approve. The doctor is told exactly which days are approved.</p>' : ''}` : ''}
      ${open ? `<div class="decide stack"><label class="field" for="dec-msg">Message to the doctor <span class="hint">Needed to decline, ask a question, or approve days without enough doctors</span><textarea id="dec-msg" rows="2"></textarea></label>
        <div class="row">
          <button type="button" class="btn good" data-act="decide" data-val="${chk.days.length ? 'approveDays' : 'approve'}"${ui.busy ? ' disabled' : ''}>${chk.days.length ? 'Approve ticked days' : 'Approve'}</button>
          <button type="button" class="btn" data-act="decide" data-val="hold">Ask a question</button>
          <button type="button" class="btn danger" data-act="decide" data-val="decline">Decline</button>
        </div>
        ${r.question ? `<p class="small muted">Question already asked: "${esc(r.question)}"</p>` : ''}</div>` : `<p class="small muted">${esc(reqStatus[r.status])}${r.decidedBy ? ` by ${esc(r.decidedBy)}` : ''}${r.decidedAt ? `, ${esc(r.decidedAt.slice(0, 16).replace('T', ' '))}` : ''}.${r.message ? ` "${esc(r.message)}"` : ''}${(r.excludedDates || []).length ? ` Not approved: ${esc(r.excludedDates.map(R.fmtShort).join(', '))}.` : ''}</p>
        ${(r.cover || []).length ? `<p class="small">Cover put on the rota: ${esc(r.cover.map((c) => c.text).join('; '))}</p>` : ''}
        ${r.status === 'authorised' ? `<div class="row"><button type="button" class="btn danger" data-act="revoke" data-id="${r.id}">${ui.confirm === 'revoke:' + r.id ? 'Confirm: cancel this leave' : 'Cancel this leave'}</button></div>` : ''}`}
    </article>`;
  }

  // ---------- authoriser: EDT / SDT ----------
  function renderEdt() {
    const s = st();
    const f = ui.edt;
    const docs = docsSorted().filter((d) => R.inPost(d, TODAY) || (d.start && d.start > TODAY));
    const doc = R.doctorById(s, f.doctorId);
    const chk = doc && R.isIso(f.date) ? R.edtCheck(s, doc.id, f.date, f.session) : null;
    const rows = docs.map((d) => {
      const e = R.edtEntitlement(s, d, TODAY);
      const behind = e.booked < e.dueByToday;
      return `<tr><td><strong>${esc(d.name)}</strong> <span class="muted small">${esc(d.grade)}</span></td><td>${esc(e.type)}</td><td class="tabular">${e.hours}h/wk${e.fte < 1 ? ` × ${e.fte}` : ''}</td>
        <td class="tabular small">${esc(R.fmtUk(e.from))} – ${esc(R.fmtUk(e.to))}</td><td class="tabular">${e.due}h</td><td class="tabular">${e.dueByToday}h</td><td class="tabular${behind ? ' warn-text' : ''}">${e.booked}h</td>
        <td><input type="number" min="0" max="10" step="0.5" aria-label="Hours a week for ${esc(d.name)}" data-act="doc-field" data-doc="${d.id}" data-field="edtHours" value="${esc(d.edtHours === undefined || d.edtHours === null ? '' : d.edtHours)}" placeholder="${R.edtDefault(Object.assign({}, d, { edtHours: undefined })).hours}"></td></tr>`;
    }).join('');
    const upcoming = (s.edt || []).filter((e) => e.date >= R.addDays(TODAY, -7)).sort((a, b) => (a.date < b.date ? -1 : 1)).slice(0, 80);
    return `<div class="page-head"><div><h1>EDT / SDT</h1><p class="muted">Booked as a morning or afternoon session and shown on the rota. Residents at GUH only, with enough other residents there; senior doctors at any site if senior cover is kept.</p></div></div>
      <div class="split"><section class="main card"><h2>Book a session</h2>
        <div class="row" style="align-items:flex-end">
          <label class="field" for="edt-doc">Doctor<select id="edt-doc" data-edt="doctorId"><option value="">Choose</option>${docOptions(f.doctorId)}</select></label>
          <label class="field" for="edt-date">Date<input type="date" id="edt-date" data-edt="date" value="${esc(f.date)}"></label>
          <div class="field">Session<div class="seg">${['AM', 'PM'].map((x) => `<button type="button" class="chip" data-act="edt-session" data-val="${x}" aria-pressed="${f.session === x}">${x === 'AM' ? 'Morning' : 'Afternoon'}</button>`).join('')}</div></div>
        </div>
        ${chk ? `<ul class="checks">${chk.checks.map((c) => `<li class="${c.ok ? 'ok' : 'bad'}"><span aria-hidden="true">${c.ok ? '✓' : '✕'}</span> <strong>${esc(c.rule)}</strong><br><span class="small">${esc(c.detail)}</span></li>`).join('')}</ul>
          ${chk.ok ? `<div class="row"><button type="button" class="btn primary" data-act="edt-book">Book ${esc(chk.type)} ${f.session === 'AM' ? 'morning' : 'afternoon'}</button></div>` : `<div class="row" style="align-items:flex-end"><label class="field" for="edt-reason" style="flex:1">Book anyway: reason<input type="text" id="edt-reason" placeholder="Why the rule can be broken this time"></label><button type="button" class="btn" data-act="edt-book">Book anyway</button></div>`}`
          : '<p class="muted small">Choose a doctor and date to check the rules.</p>'}
      </section>
      <section class="side card"><h2>Booked</h2>${upcoming.map((e) => `<div class="row small" style="justify-content:space-between"><span><span class="tabular">${esc(R.fmtShort(e.date))} ${e.session}</span> · <strong>${esc(docName(e.doctorId))}</strong> ${esc(e.type)}${e.reason ? ' <span class="cv exception">override</span>' : ''}</span><button type="button" class="btn ghost small-btn" data-act="edt-remove" data-id="${e.id}">Remove</button></div>`).join('') || '<p class="muted small">None booked.</p>'}</section></div>
      <section class="card"><h2>Entitlement by grade (Wales defaults, editable)</h2>
        <p class="small muted">Defaults: F1 SDT 1h/week, F2 SDT 3h, IMT and JCF EDT 2h, registrars and senior clinical fellows EDT 4h, specialty and specialist doctors SPA 4h. Pro-rata for less than full time. Each session counts as 4 hours. Type a number to override for one doctor. Please check these against current Welsh contracts and local agreements.</p>
        <div class="scroll-x"><table class="plain"><thead><tr><th>Doctor</th><th>Type</th><th>Default</th><th>Period</th><th>Due in period</th><th>Due by today</th><th>Booked</th><th>Override h/wk</th></tr></thead><tbody>${rows}</tbody></table></div></section>`;
  }

  // ---------- authoriser: set up ----------
  function renderSetup() {
    const tabs = [['master', 'Master sheet'], ['people', 'People & leave'], ['import', 'Import'], ['rotation', 'Rotations'], ['consultant', 'Consultant of the week'], ['leave', 'Leave restrictions'], ['settings', 'Settings']];
    const body = { master: renderMaster, people: renderPeople, import: renderImport, rotation: renderRotation, consultant: renderConsultant, leave: renderRestrictions, settings: renderSettings }[ui.setupTab] || renderMaster;
    return `<div class="page-head"><div><h1>Set up</h1><p class="muted">Doctors, their usual locations, imported on-calls, rotations and settings. The rota is built from these.</p></div></div>
      <div class="tabs" role="tablist">${tabs.map(([id, l]) => `<button type="button" role="tab" data-act="setup-tab" data-val="${id}" aria-selected="${ui.setupTab === id}">${l}</button>`).join('')}</div>${body()}`;
  }
  const LOCATIONS = R.SITE_CODES;
  const locOptions = (cur, blank) => `<option value="">${blank || '(blank)'}</option>` + LOCATIONS.map((c) => `<option value="${esc(c)}"${c === cur ? ' selected' : ''}>${esc(c)}</option>`).join('');
  function gradeSelect(attrs, cur) {
    const known = groupForGrade(cur);
    return `<select ${attrs}>${gradeOptions(cur)}<option value="__other"${!known ? ' selected' : ''}>${known ? 'Other…' : `Other: ${esc(cur)}`}</option></select>`;
  }
  function newDocForm() {
    const nd = ui.nd = ui.nd || { grade: 'F1', other: '', group: 'res', name: '', email: '', start: '', end: '', provisional: false, replaces: '', pattern: '', placeholder: false, suggested: false };
    if (!nd.suggested) {
      const prev = nd.replaces ? R.doctorById(st(), nd.replaces) : null;
      const from = prev && prev.end ? R.addDays(prev.end, 1) : TODAY;
      const sug = R.suggestDates(nd.grade === '__other' ? nd.other : nd.grade, from);
      nd.start = sug.start || (prev && prev.end ? from : ''); nd.end = sug.end; nd.provisional = sug.provisional; nd.note = sug.note; nd.suggested = true;
    }
    const isOther = nd.grade === '__other';
    const ends = docsSorted().filter((d) => d.end && d.end >= R.addDays(TODAY, -60));
    return `<section class="card"><h2>Add a new doctor</h2>
      <div class="grid-3">
        <label class="field" for="nd-grade">Grade<select id="nd-grade" data-nd="grade">${gradeOptions(isOther ? '' : nd.grade)}<option value="__other"${isOther ? ' selected' : ''}>Other…</option></select></label>
        ${isOther ? `<label class="field" for="nd-other">Grade (other)<input type="text" id="nd-other" data-nd="other" value="${esc(nd.other)}" placeholder="e.g. Trust grade doctor"></label>
        <label class="field" for="nd-group">Category<select id="nd-group" data-nd="group"><option value="res"${nd.group === 'res' ? ' selected' : ''}>Resident doctor</option><option value="mg"${nd.group === 'mg' ? ' selected' : ''}>Senior doctor</option></select></label>` : ''}
        <label class="field" for="nd-replaces">Takes over the column of <span class="hint">So the rota continues in the same column</span><select id="nd-replaces" data-nd="replaces"><option value="">Nobody: new column</option>${ends.map((d) => `<option value="${d.id}"${d.id === nd.replaces ? ' selected' : ''}>${esc(d.name)} (${esc(d.grade)}), leaves ${esc(R.fmtUk(d.end))}</option>`).join('')}</select></label>
        <label class="field" for="nd-name">Name<input type="text" id="nd-name" data-nd="name" value="${esc(nd.name)}" placeholder="${nd.placeholder ? 'Name not known yet' : 'e.g. A. Jones'}"${nd.placeholder ? ' disabled' : ''}></label>
        <label class="field" for="nd-email">NHS email<input type="email" id="nd-email" data-nd="email" value="${esc(nd.email)}"${nd.placeholder ? ' disabled' : ''}></label>
        <label class="field" for="nd-start">Starts<input type="date" id="nd-start" data-nd="start" value="${esc(nd.start)}"></label>
        <label class="field" for="nd-end">Finishes<input type="date" id="nd-end" data-nd="end" value="${esc(nd.end)}"></label>
        <label class="field" for="nd-pattern">Usual locations: copy from<select id="nd-pattern" data-nd="pattern"><option value="">Set them after adding</option>${docOptions(nd.pattern || nd.replaces)}</select></label>
      </div>
      <div class="row small"><label class="opt"><input type="checkbox" data-nd="placeholder"${nd.placeholder ? ' checked' : ''}> <span>Name not known yet: add as "New ${esc(isOther ? nd.other || 'doctor' : nd.grade)}"</span></label>
        <label class="opt"><input type="checkbox" data-nd="provisional"${nd.provisional ? ' checked' : ''}> <span>Finish date not confirmed</span></label></div>
      ${nd.note ? `<p class="small muted">Suggested dates: ${esc(nd.note)} You can change them.</p>` : ''}
      <div class="row"><button type="button" class="btn primary" data-act="add-doc">Add doctor</button></div></section>`;
  }
  function renderMaster() {
    const cols = R.columns(st(), '1900-01-01', '2999-12-31');
    const row = (d, i) => {
      const plan = d.plan || { A: ['', '', '', '', ''] };
      return `<tr class="${d.end && d.end < TODAY ? 'past' : ''}${i ? ' follows' : ''}">
        <td>${i ? '<span class="follow" aria-hidden="true">↳</span>' : ''}<input type="text" aria-label="Name" data-act="doc-field" data-doc="${d.id}" data-field="name" value="${esc(d.name)}" style="width:120px">${d.placeholder ? ' <span class="tbc">TBC</span>' : ''}</td>
        <td>${gradeSelect(`aria-label="Grade for ${esc(d.name)}" data-act="doc-grade" data-doc="${d.id}"`, d.grade)}</td>
        <td><select aria-label="Category for ${esc(d.name)}" data-act="doc-field" data-doc="${d.id}" data-field="group"><option value="res"${d.group === 'res' ? ' selected' : ''}>Resident</option><option value="mg"${d.group === 'mg' ? ' selected' : ''}>Senior</option></select></td>
        <td><input type="date" aria-label="Starts" data-act="doc-field" data-doc="${d.id}" data-field="start" value="${esc(d.start || '')}"></td>
        <td><input type="date" aria-label="Finishes" data-act="doc-field" data-doc="${d.id}" data-field="end" value="${esc(d.end || '')}"><br><label class="small"><input type="checkbox" data-act="doc-field" data-doc="${d.id}" data-field="endProvisional"${d.endProvisional ? ' checked' : ''}> not confirmed</label></td>
        ${[0, 1, 2, 3, 4].map((k) => `<td><select aria-label="${esc(d.name)} ${R.DOW_LONG[k]}" data-act="plan-day" data-doc="${d.id}" data-week="A" data-i="${k}">${locOptions((plan.A || [])[k] || '')}</select></td>`).join('')}
        <td><button type="button" class="btn ghost small-btn" data-act="remove-doc" data-doc="${d.id}">${ui.confirm === 'remove:' + d.id ? 'Confirm' : 'Remove'}</button></td></tr>`;
    };
    const section = (g, title) => cols.filter((c) => c.docs[0].group === g).map((c) => c.docs.map((d, i) => row(d, i)).join('')).join('') ? `<tr class="sec"><th colspan="11">${title}</th></tr>` + cols.filter((c) => c.docs[0].group === g).map((c) => c.docs.map((d, i) => row(d, i)).join('')).join('') : '';
    const b = ui.bulk = ui.bulk || { ids: [], from: TODAY, to: R.addDays(TODAY, 4), days: [0, 1, 2, 3, 4], code: 'GUH', reason: '' };
    return `<section class="card"><p class="muted small">Everyone on the rota, in column order. A doctor who takes over another's column is shown underneath them (↳) and continues in the same column on the rota. The Monday–Friday locations are each doctor's usual week; on-calls, leave and one-off changes sit on top. Changes save straight away.</p>
      <div class="scroll-x"><table class="plain master"><thead><tr><th>Name</th><th>Grade</th><th>Category</th><th>Starts</th><th>Finishes</th>${R.DOW_SHORT.slice(0, 5).map((d) => `<th>${d}</th>`).join('')}<th></th></tr></thead>
      <tbody>${section('res', 'Resident doctors')}${section('mg', 'Senior doctors')}</tbody></table></div></section>
      <div class="grid-2">${newDocForm()}
      <section class="card"><h2>Put a group of doctors somewhere</h2><p class="small muted">For a block of dates (a week, a month or a whole rotation). This changes the rota for those days only; the usual week above stays the same.</p>
        <div class="pick-list">${docsSorted().filter((d) => !d.end || d.end >= TODAY).map((d) => `<label class="opt small"><input type="checkbox" data-bulk-id="${d.id}"${b.ids.includes(d.id) ? ' checked' : ''}> <span>${esc(d.name)} <span class="muted">${esc(d.grade)}</span></span></label>`).join('')}</div>
        <div class="row" style="align-items:flex-end">
          <label class="field" for="bk-from">From<input type="date" id="bk-from" data-bulk="from" value="${esc(b.from)}"></label>
          <label class="field" for="bk-to">To<input type="date" id="bk-to" data-bulk="to" value="${esc(b.to)}"></label>
          <label class="field" for="bk-code">Location<select id="bk-code" data-bulk="code">${locOptions(b.code)}</select></label></div>
        <div class="row small">${R.DOW_SHORT.slice(0, 5).map((d, k) => `<label class="opt"><input type="checkbox" data-bulk-day="${k}"${b.days.includes(k) ? ' checked' : ''}> <span>${d}</span></label>`).join('')}</div>
        <label class="field" for="bk-reason">Reason <span class="hint">Shown in the change log</span><input type="text" id="bk-reason" data-bulk="reason" value="${esc(b.reason)}" placeholder="e.g. covering NHH day unit for the rotation"></label>
        <div class="row"><button type="button" class="btn primary" data-act="bulk-apply">Apply to the rota</button></div></section></div>`;
  }
  function renderPeople() {
    const inp = (d, field, type, val, w, extra) => `<input type="${type}" aria-label="${field} for ${esc(d.name)}" data-act="doc-field" data-doc="${d.id}" data-field="${field}" value="${esc(val)}"${w ? ` style="width:${w}"` : ''}${extra || ''}>`;
    const rows = docsSorted().filter((d) => !d.placeholder).map((d) => {
      const pr = R.alProRata(d);
      const bal = R.balance(st(), d);
      return `<tr${d.end && d.end < TODAY ? ' class="past"' : ''}>
      <td><strong>${esc(d.name)}</strong><br><span class="small muted">${esc(d.grade)}</span></td>
      <td>${inp(d, 'email', 'email', d.email || '', '200px')}</td>
      <td>${inp(d, 'fte', 'number', d.fte === undefined ? 1 : d.fte, '64px', ' step="0.1" min="0.1" max="1"')}</td>
      <td>${inp(d, 'alEntitlement', 'number', d.alEntitlement, '70px', ' step="0.5"')}<br><span class="small muted" title="${esc(pr.basis)}">Wales default ${pr.days}${Number(d.alEntitlement) !== pr.days ? ` <button type="button" class="linkish" data-act="use-al" data-doc="${d.id}">use</button>` : ''}</span></td>
      <td>${inp(d, 'alExtra', 'number', d.alExtra || 0, '64px', ' step="0.5"')}</td>
      <td><strong>${bal.al.left}</strong><span class="small muted"> of ${bal.al.entitlement}</span></td>
      <td>${inp(d, 'slEntitlement', 'number', d.slEntitlement, '64px', ' step="0.5"')}<br><span class="small muted">${bal.sl.left} left</span></td>
      <td>${inp(d, 'aliases', 'text', (d.aliases || []).join(', '), '130px')}</td></tr>`;
    }).join('');
    return `<section class="card"><p class="muted small">Annual leave allowance is filled in from the Wales default for the grade, dates in post and FTE (28 days a year for foundation, core and fellows; 33 for registrars ST3+; 29 for specialty doctors, 34 after 2 years in grade; 34 for specialist grade) and can be changed. Use "Carried over / extra" for accrued or carried-over days. Doctors see their remaining days under "Check my request".</p>
      <div class="scroll-x"><table class="plain people"><thead><tr><th>Doctor</th><th>NHS email</th><th>FTE</th><th>Annual leave days</th><th>Carried over / extra</th><th>AL left</th><th>Study leave days</th><th>Other spellings (on-call sheet)</th></tr></thead><tbody>${rows}</tbody></table></div></section>`;
  }
  const IMPORT_INFO = {
    oncall: ['On-call rota (from the trust)', 'Each week, put the trust on-calls into this template (list or grid layout) and import it. Earlier imports for the same dates are replaced.'],
    plan: ['Weekly plan', 'Everyone\'s usual site for each weekday. Alternating weeks are supported with a "B" row.'],
    staff: ['Staff list', 'Names, grades, emails, dates in post and leave allowances. Existing doctors are updated by name.']
  };
  function renderImport() {
    const p = ui.importPreview;
    const card = (kind) => {
      const [title, text] = IMPORT_INFO[kind];
      let preview = '';
      if (p && p.kind === kind) {
        const unknown = Object.keys(p.parsed.unknownNames || {});
        const mapRows = unknown.map((name, i) => `<div class="row small"><span>"${esc(name)}" (${plural(p.parsed.unknownNames[name], 'row')}) is</span><label class="sr-only" for="map-${i}">Match ${esc(name)}</label><select id="map-${i}" data-act="map-name" data-name="${esc(name)}"><option value="">Not one of ours: skip</option>${docOptions(p.mapping[name])}</select></div>`).join('');
        let summary = '';
        if (kind === 'oncall') {
          const usable = p.parsed.entries.filter((e) => e.doctorId || p.mapping[e.name]).length;
          summary = `<p><strong>${usable}</strong> on-call shifts from ${p.parsed.from ? esc(R.fmtUk(p.parsed.from)) : '?'} to ${p.parsed.to ? esc(R.fmtUk(p.parsed.to)) : '?'} (${esc(p.parsed.format)} layout, sheet "${esc(p.parsed.sheet)}"). Imported on-calls already in those dates will be replaced.</p>
            ${p.parsed.skipped.length ? `<p class="small muted">${plural(p.parsed.skipped.length, 'cell')} with site codes were ignored (sites come from the weekly plan).</p>` : ''}`;
          p.count = usable;
        } else if (kind === 'plan') {
          p.count = p.parsed.plans.filter((x) => x.doctorId || p.mapping[x.name]).length;
          summary = `<p>Weekly plans for <strong>${p.count}</strong> doctors. Their current plans will be replaced.</p>`;
        } else {
          const existing = p.parsed.staff.filter((x) => st().doctors.some((d) => d.name.toLowerCase() === x.name.toLowerCase())).length;
          summary = `<p><strong>${p.parsed.staff.length}</strong> people: ${existing} updated, ${p.parsed.staff.length - existing} new.</p>${(p.parsed.problems || []).map((x) => `<p class="small warn-text">${esc(x)}</p>`).join('')}`;
          p.count = p.parsed.staff.length;
        }
        const bad = (p.parsed.badCodes || []).slice(0, 8).map((b) => `Row ${b.row}: "${esc(b.code)}"`).join(', ');
        preview = `<div class="notice info stack"><strong>${esc(p.fileName)}</strong>${summary}
          ${unknown.length ? `<p class="small"><strong>New names to match once.</strong> The spelling is remembered for next time.</p>${mapRows}` : ''}
          ${bad ? `<p class="small" style="color:var(--bad-text)">Codes not recognised (skipped): ${bad}${p.parsed.badCodes.length > 8 ? ' …' : ''}</p>` : ''}
          ${p.parsed.badDates ? `<p class="small" style="color:var(--bad-text)">${plural(p.parsed.badDates, 'row')} skipped because the date could not be read. Use dd/mm/yyyy.</p>` : ''}
          <div class="row"><button type="button" class="btn primary" data-act="confirm-import"${p.count ? '' : ' disabled'}>Import</button><button type="button" class="btn ghost" data-act="cancel-import">Cancel</button></div></div>`;
      }
      return `<article class="card"><h2>${title}</h2><p class="small muted">${text}</p>
        <div class="drop"><div class="row"><button type="button" class="btn" data-act="template" data-kind="${kind}">Download template</button></div>
        <label class="field" for="file-${kind}">Choose a filled-in .xlsx file<input type="file" id="file-${kind}" accept=".xlsx" data-act="import-file" data-kind="${kind}"></label></div>${preview}</article>`;
    };
    return `<div class="grid-3">${card('oncall')}${card('plan')}${card('staff')}</div>
      <p class="muted small">On-call layouts accepted: a list (Date, Doctor, Shift) or a grid like the rota (dates down the side, doctors across the top). Names are matched by exact name, a remembered spelling, or surname and initial.</p>`;
  }
  function renderRotation() {
    const s = st().settings;
    const list = (s.changeovers || []).slice().sort((a, b) => (a.date < b.date ? -1 : 1));
    const auto = Object.entries(R.changeoversIn(st(), R.addDays(TODAY, -30), R.addDays(TODAY, 400))).sort();
    const progs = Object.entries(R.PROGRAMMES).map(([k, p]) => `<tr><td><strong>${esc(p.label)}</strong><br><span class="small muted">${esc(p.note)}</span></td><td class="tabular">${R.programmeStarts(k, TODAY, 3).map((d) => esc(R.fmtLong(d))).join('<br>')}</td></tr>`).join('');
    const open = st().doctors.filter((d) => d.end && d.end >= R.addDays(TODAY, -7) && !d.placeholder && !st().doctors.some((x) => x.slot === R.slotOf(d) && x.start && x.start > d.end));
    return `<div class="grid-2">
      <section class="card"><h2>Rotation</h2>
        <div class="row" style="align-items:flex-end"><label class="field" for="rot-start">Current rotation started<input type="date" id="rot-start" value="${esc(s.rotation.start || '')}"></label><button type="button" class="btn primary" data-act="save-rotation">Save</button></div>
        <p class="small muted">The rota opens on today and shows 2 months, with "Show 2 more months" (and "Show 2 earlier months") to see further. Doctors are rostered until ${esc(R.fmtLong(R.rotaHorizon(st(), TODAY)))} at the moment.</p>
        <h3>Next start dates</h3><div class="scroll-x"><table class="plain small"><tbody>${progs}</tbody></table></div>
        <p class="small muted">These are suggestions. Each doctor's dates can be changed on the Master sheet, and finish dates can be marked "not confirmed".</p></section>
      <section class="card"><h2>Who follows whom</h2>
        ${open.length ? `<p class="small muted">These doctors finish with nobody named to take over their column. Add a placeholder now ("New F1") and fill in the name later.</p>
        ${open.map((d) => `<div class="row small" style="justify-content:space-between"><span><strong>${esc(d.name)}</strong> (${esc(d.grade)}) finishes ${esc(R.fmtUk(d.end))}${d.endProvisional ? ' (not confirmed)' : ''}</span><button type="button" class="btn small-btn" data-act="add-placeholder" data-doc="${d.id}">Add "New ${esc(d.grade.split(' ')[0])}"</button></div>`).join('')}` : '<p class="small muted">Everyone who is finishing has somebody (or a placeholder) following them.</p>'}
        <h3>Changeover bands</h3>
        <p class="small muted">Shown across the rota. Doctors starting or leaving are added automatically; add other dates here.</p>
        ${list.map((c) => `<div class="row small" style="justify-content:space-between"><span><span class="tabular">${esc(R.fmtUk(c.date))}</span> · ${esc(c.label)}</span><button type="button" class="btn ghost small-btn" data-act="remove-changeover" data-id="${c.id}">Remove</button></div>`).join('') || ''}
        <div class="row" style="align-items:flex-end"><label class="field" for="co-date">Date<input type="date" id="co-date"></label><label class="field" for="co-label" style="flex:1">Description<input type="text" id="co-label" placeholder="e.g. IMT changeover"></label><button type="button" class="btn primary" data-act="add-changeover">Add</button></div>
        <details><summary class="small">Automatic changeovers (${auto.length})</summary>${auto.map(([d, t]) => `<div class="small"><span class="tabular">${esc(R.fmtUk(d))}</span> · ${esc(t.join(' · '))}</div>`).join('')}</details></section></div>`;
  }
  function renderConsultant() {
    const map = st().settings.consultantOfWeek || {};
    const keys = Object.keys(map).sort();
    const startMon = R.mondayOf(st().settings.rotation.start || TODAY);
    const first = keys.length && keys[0] < startMon ? keys[0] : startMon;
    let last = R.addDays(R.mondayOf(TODAY), 7 * 8);
    if (keys.length && keys[keys.length - 1] > last) last = keys[keys.length - 1];
    last = R.addDays(last, 7 * 4 * ui.consExtra);
    const months = new Map();
    for (let w = R.mondayOf(first); w <= last; w = R.addDays(w, 7)) {
      const k = w.slice(0, 7);
      if (!months.has(k)) months.set(k, []);
      months.get(k).push(w);
    }
    const MONTH = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const thisMonth = TODAY.slice(0, 7), nextMonth = R.addDays(TODAY.slice(0, 8) + '28', 7).slice(0, 7);
    const blocks = [...months].map(([k, weeks]) => {
      const missing = weeks.filter((w) => !map[w]).length;
      return `<details class="month"${k === thisMonth || k === nextMonth ? ' open' : ''}><summary><strong>${MONTH[Number(k.slice(5)) - 1]} ${k.slice(0, 4)}</strong> <span class="small muted">${plural(weeks.length, 'week')}${missing ? ` · ${missing} not set` : ''}</span></summary>
        ${weeks.map((w) => `<label class="row small cons-row" for="cons-${w}"><span class="tabular" style="width:130px">w/c ${esc(R.fmtShort(w))}</span><input type="text" id="cons-${w}" data-act="cons" data-week="${w}" value="${esc(map[w] || '')}" placeholder="Consultant"></label>`).join('')}</details>`;
    }).join('');
    return `<section class="card"><h2>Consultant of the week</h2><p class="small muted">Shown in the rota's Consultant column, the list and by-site views, and the Excel copy. Months are folded away so the list stays short; open one to edit it. Changes save when you leave the box.</p>
      ${blocks}<div class="row"><button type="button" class="btn" data-act="cons-more">Add 4 more weeks</button></div></section>`;
  }
  function renderRestrictions() {
    const list = (st().settings.restricted || []).slice().sort((a, b) => (a.start < b.start ? -1 : 1));
    return `<section class="card"><h2>Leave restrictions</h2><p class="small muted">These replace the red and purple date colours in the old Excel. "NO AL" (red) warns doctors and flags the request; "AVOID" (purple) is a softer warning. Authorisers can still approve.</p>
        <div class="scroll-x"><table class="plain"><thead><tr><th>Dates</th><th>Level</th><th>Note</th><th></th></tr></thead><tbody>
        ${list.map((r) => `<tr><td class="tabular">${esc(rangeText(r.start, r.end))}</td><td><span class="restr ${r.level.replace(' ', '-')}">${esc(r.level)}</span></td><td class="small">${esc(r.note)}</td><td><button type="button" class="btn ghost small-btn" data-act="remove-restr" data-id="${r.id}">Remove</button></td></tr>`).join('') || '<tr><td colspan="4" class="muted">None.</td></tr>'}
        </tbody></table></div>
        <div class="row" style="align-items:flex-end">
          <label class="field" for="rs-start">From<input type="date" id="rs-start"></label><label class="field" for="rs-end">To<input type="date" id="rs-end"></label>
          <label class="field" for="rs-level">Level<select id="rs-level"><option>NO AL</option><option>AVOID</option></select></label>
          <label class="field" for="rs-note">Note<input type="text" id="rs-note" placeholder="Why"></label>
          <button type="button" class="btn primary" data-act="add-restr">Add</button></div></section>`;
  }
  function renderSettings() {
    const s = st().settings;
    return `<div class="grid-2">
      <section class="card"><h2>Authorisers</h2><p class="small muted">The names that can sign in (they share one password). Shown in the change log next to every change.</p>
        <label class="field" for="set-auth">Names <span class="hint">One per line</span><textarea id="set-auth" rows="4">${esc((s.authorisers || []).join('\n'))}</textarea></label>
        <label class="opt small"><input type="checkbox" id="set-email-docs"${s.emailDoctors ? ' checked' : ''}> <span>Also email doctors (decisions and cover). Off: doctors check "Check my request" instead, and only the alert emails below are sent.</span></label>
        <label class="field" for="set-mail">Alert emails <span class="hint">Who is emailed about new requests and sickness. One per line.</span><textarea id="set-mail" rows="3">${esc((s.authoriserEmails || []).join('\n'))}</textarea></label>
        <div class="row"><button type="button" class="btn primary" data-act="save-authorisers">Save</button><span class="small muted">Emails: ${store.emailOn ? '<strong style="color:var(--ok-fg)">switched on</strong>' : store.mode === 'local' ? 'not available in demo mode' : 'not switched on yet (in-app messages only). See the README to add an email key.'}</span></div></section>
      ${store.mode === 'remote' ? `<section class="card"><h2>Change the authoriser password</h2><p class="small muted">Everyone is signed out afterwards. Share the new password in person or by phone, never by email.</p>
        <label class="field" for="pw-cur">Current password<input type="password" id="pw-cur" autocomplete="current-password"></label>
        <label class="field" for="pw-new">New password <span class="hint">At least 12 characters</span><input type="password" id="pw-new" autocomplete="new-password"></label>
        <div class="row"><button type="button" class="btn primary" data-act="change-pw">Change password</button></div></section>` : ''}
      <section class="card"><h2>Bank holidays</h2><p class="small muted">One date per line (yyyy-mm-dd). Bank holidays are not working days.</p>
        <textarea id="set-bh" rows="6">${esc((s.bankHolidays || []).join('\n'))}</textarea><div class="row"><button type="button" class="btn" data-act="save-bh">Save</button></div></section>
      <section class="card"><h2>${s.demo ? 'Start for real' : 'Data'}</h2>
        <p class="small muted">${s.demo ? 'The rota is showing example data with fictional doctors. "Start fresh" removes all doctors, plans, on-calls, requests and EDT, and keeps the rules, authorisers and alert emails. Then add your doctors under People or import the staff list.' : 'Download a full Excel copy from the Rota page at any time.'}</p>
        <div class="row">${s.demo || store.mode === 'local' ? `<button type="button" class="btn danger" data-act="start-fresh">${ui.confirm === 'fresh' ? 'Confirm: remove all example data' : 'Start fresh'}</button>` : ''}
        ${store.mode === 'local' ? '<button type="button" class="btn" data-act="reset-demo">Reload example data</button>' : ''}</div></section></div>`;
  }

  // ---------- change log ----------
  function renderLog() {
    const q = ui.logQ.toLowerCase();
    const rows = st().audit.filter((a) => !q || (a.by + a.action + a.detail).toLowerCase().includes(q)).slice(0, 400);
    return `<div class="page-head"><div><h1>Change log</h1><p class="muted">Every change: who, when and what.</p></div>
      <label class="field" for="log-q">Search<input type="search" id="log-q" data-live="logQ" value="${esc(ui.logQ)}" placeholder="Name, action or detail"></label></div>
      <section class="card"><div class="scroll-x"><table class="plain"><thead><tr><th>When</th><th>Who</th><th>What</th><th>Details</th></tr></thead><tbody>
      ${rows.map((a) => `<tr><td class="tabular small">${esc(a.at.slice(0, 16).replace('T', ' '))}</td><td>${esc(a.by)}</td><td>${esc(a.action)}</td><td class="small">${esc(a.detail)}</td></tr>`).join('') || '<tr><td colspan="4" class="muted">No matching entries.</td></tr>'}
      </tbody></table></div></section>`;
  }

  // ---------- dialogs: sign in and first-time setup ----------
  function renderDialog() {
    if (!ui.dialog) return '';
    if (ui.dialog === 'setup') {
      return `<div class="scrim" data-act="close-dialog"></div><section class="dialog" role="dialog" aria-modal="true" aria-labelledby="dlg-title"><div class="grabber" aria-hidden="true"></div><form id="setup-form" class="stack">
        <h2 id="dlg-title" class="sheet-head">First-time setup</h2>
        <p class="small">Choose the one password the rota authorisers will share. Viewers and doctors sending leave requests never need it.</p>
        <p class="small muted">Make it long and unguessable: for example, 3 or 4 unrelated words mixed with a haematology term only you would think of. Don't use the department name, ward names or anything on this site. It is stored only as a scrambled hash; nobody (including the developer) can read it back.</p>
        ${store.setupCodeRequired ? '<label class="field" for="su-code">Setup code<input type="password" id="su-code" autocomplete="off"></label>' : ''}
        <label class="field" for="su-names">Authoriser names <span class="hint">One per line, e.g. initial and surname</span><textarea id="su-names" rows="3">${esc((st().settings.authorisers || []).join('\n'))}</textarea></label>
        <label class="field" for="su-pw">Password <span class="hint">At least 12 characters</span><input type="password" id="su-pw" autocomplete="new-password"></label>
        <label class="field" for="su-pw2">Type it again<input type="password" id="su-pw2" autocomplete="new-password"></label>
        <div class="row"><button type="submit" class="btn primary">Set password</button><button type="button" class="btn ghost" data-act="close-dialog">Cancel</button></div></form></section>`;
    }
    const names = st().settings.authorisers || [];
    return `<div class="scrim" data-act="close-dialog"></div><section class="dialog" role="dialog" aria-modal="true" aria-labelledby="dlg-title"><div class="grabber" aria-hidden="true"></div><form id="signin-form" class="stack">
      <h2 id="dlg-title" class="sheet-head">Authoriser sign in</h2>
      <p class="small muted">Only the rota authorisers need to sign in. Everyone else can view the rota and send leave requests without a password.</p>
      <label class="field" for="si-name">Your name<select id="si-name">${names.map((n) => `<option>${esc(n)}</option>`).join('')}</select></label>
      ${store.mode === 'local' ? '<p class="notice info small">Demo mode: no password needed.</p>' : '<label class="field" for="si-pw">Password<input type="password" id="si-pw" autocomplete="current-password"></label>'}
      <div class="row"><button type="submit" class="btn primary">Sign in</button><button type="button" class="btn ghost" data-act="close-dialog">Cancel</button></div>
      ${store.mode === 'remote' && !store.hasPassword ? '<p class="small"><button type="button" class="linkish" data-act="open-setup">No password set yet? First-time setup</button></p>' : ''}</form></section>`;
  }


  // ---------- motion: springs, closing along the way things came in, drag to dismiss (phones) ----------
  const reduceMotion = () => window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const isPhone = () => window.matchMedia('(max-width: 760px)').matches;
  // A critically damped spring by default (damping 1, response in seconds). Starts from the current value and velocity.
  function spring(from, to, velocity, opts, onFrame, onDone) {
    const response = (opts && opts.response) || 0.35, damping = (opts && opts.damping) || 1;
    const k = Math.pow(2 * Math.PI / response, 2), c = 4 * Math.PI * damping / response;
    let x = from, v = velocity || 0, last = performance.now(), raf = 0, stopped = false;
    const step = (now) => {
      if (stopped) return;
      let dt = Math.min(0.032, (now - last) / 1000); last = now;
      while (dt > 0) { const h = Math.min(dt, 1 / 240); v += (-k * (x - to) - c * v) * h; x += v * h; dt -= h; }
      if (Math.abs(x - to) < 0.5 && Math.abs(v) < 20) { onFrame(to); if (onDone) onDone(); return; }
      onFrame(x); raf = requestAnimationFrame(step);
    };
    raf = requestAnimationFrame(step);
    return { stop() { stopped = true; cancelAnimationFrame(raf); }, value: () => x, velocity: () => v };
  }
  const project = (v, rate) => (v / 1000) * (rate || 0.998) / (1 - (rate || 0.998)); // where a flick would come to rest
  const rubberband = (over, dim) => (over * dim * 0.55) / (dim + 0.55 * Math.abs(over));
  let overlayAnim = null;

  function overlayEl() { return document.querySelector('#sheet-root .dialog') || document.querySelector('#sheet-root .sheet'); }
  // Close a sheet or dialog: phones slide it down, wider screens reverse the way it arrived.
  function closeOverlay(apply, velocity) {
    const el = overlayEl();
    const finish = () => { overlayAnim = null; apply(); render(); };
    if (!el || reduceMotion()) return finish();
    if (overlayAnim) overlayAnim.stop();
    const scrim = document.querySelector('#sheet-root .scrim');
    el.classList.remove('enter');
    if (isPhone()) {
      const h = el.getBoundingClientRect().height;
      const from = currentY(el);
      overlayAnim = spring(from, h, velocity || 0, { response: 0.3 }, (y) => { el.style.transform = `translateY(${y}px)`; if (scrim) scrim.style.opacity = String(Math.max(0, 1 - y / h)); }, finish);
    } else {
      el.style.transition = 'opacity 180ms ease-in, transform 180ms ease-in, filter 180ms ease-in';
      el.style.opacity = '0'; el.style.transform = 'scale(0.96)'; el.style.filter = 'blur(4px)';
      if (scrim) { scrim.style.transition = 'opacity 180ms ease-in'; scrim.style.opacity = '0'; }
      setTimeout(finish, 180);
    }
  }
  function currentY(el) { const m = new DOMMatrixReadOnly(getComputedStyle(el).transform); return m.m42 || 0; }

  // Drag the sheet down by its grabber or title (phones). Tracks 1:1, rubber-bands upwards, and uses the
  // flick's velocity to decide whether to close; otherwise springs back.
  document.addEventListener('pointerdown', (ev) => {
    if (!isPhone() || ev.button > 0) return;
    const handle = ev.target.closest('#sheet-root .grabber, #sheet-root .sheet-head');
    if (!handle || ev.target.closest('button, input, select, textarea, a')) return;
    const el = overlayEl();
    if (!el) return;
    if (overlayAnim) { overlayAnim.stop(); overlayAnim = null; }
    el.classList.remove('enter');
    const startY = ev.clientY, startOffset = currentY(el), h = el.getBoundingClientRect().height;
    const scrim = document.querySelector('#sheet-root .scrim');
    let dragging = false, hist = [];
    const move = (e) => {
      const dy = e.clientY - startY;
      if (!dragging) { if (Math.abs(dy) < 8) return; dragging = true; try { handle.setPointerCapture(ev.pointerId); } catch (x) { /* already released */ } }
      let y = startOffset + dy;
      if (y < 0) y = rubberband(y, h);
      el.style.transform = `translateY(${y}px)`;
      if (scrim) scrim.style.opacity = String(Math.max(0, 1 - y / h));
      hist.push([e.timeStamp, y]); if (hist.length > 6) hist.shift();
    };
    const up = () => {
      document.removeEventListener('pointermove', move); document.removeEventListener('pointerup', up); document.removeEventListener('pointercancel', up);
      if (!dragging) return;
      const y = currentY(el);
      const a = hist[0], b = hist[hist.length - 1];
      const v = a && b && b[0] > a[0] ? ((b[1] - a[1]) / (b[0] - a[0])) * 1000 : 0; // px per second
      const closing = y + project(v) > h * 0.45 && v > -50;
      if (closing) {
        const apply = ui.dialog ? () => { ui.dialog = null; } : () => { ui.editing = null; };
        closeOverlay(apply, v);
      } else {
        overlayAnim = spring(y, 0, v, { response: 0.3, damping: Math.abs(v) > 300 ? 0.8 : 1 }, (yy) => { el.style.transform = `translateY(${yy}px)`; if (scrim) scrim.style.opacity = String(Math.max(0, 1 - yy / h)); }, () => { overlayAnim = null; el.style.transform = ''; });
      }
    };
    document.addEventListener('pointermove', move); document.addEventListener('pointerup', up); document.addEventListener('pointercancel', up);
  });

  // ---------- render ----------
  function render() {
    if (!st()) return;
    const active = document.activeElement;
    const focusId = active && active.id;
    const sel = active && typeof active.selectionStart === 'number' ? [active.selectionStart, active.selectionEnd] : null;
    const scroller = $('#rota-scroll');
    const keep = scroller ? [scroller.scrollTop, scroller.scrollLeft] : null;
    renderHead();
    const views = { rota: renderRota, request: renderRequest, mine: renderMine, rules: renderRules, requests: renderRequests, edt: renderEdt, setup: renderSetup, log: renderLog };
    $('#view').innerHTML = (views[ui.view] || renderRota)();
    $('#sheet-root').innerHTML = (ui.view === 'rota' && isAuth() ? renderEditSheet() : '') + renderDialog();
    const overlayKey = (ui.dialog || '') + '|' + (ui.editing && ui.view === 'rota' ? ui.editing.doc + ui.editing.iso : '');
    if (overlayKey !== ui.overlayKey) { document.querySelectorAll('#sheet-root .sheet, #sheet-root .dialog, #sheet-root .scrim').forEach((el) => el.classList.add('enter')); ui.overlayKey = overlayKey; }
    if (ui.view === 'rota') {
      const sc = $('#rota-scroll');
      if (ui.scrollToIso) {
        // Showing earlier months: keep the date that was at the top in view.
        const row = document.querySelector(`#view [data-date="${ui.scrollToIso}"]`);
        ui.scrollToIso = null; ui.scrollToToday = false;
        if (row && sc && ui.layout === 'grid') { const th = sc.querySelector('thead'); sc.scrollTop = Math.max(0, row.offsetTop - (th ? th.offsetHeight : 70) - 2); if (keep) sc.scrollLeft = keep[1]; }
        else if (row) row.scrollIntoView({ block: 'start' });
      } else if (ui.scrollToToday) {
        ui.scrollToToday = false;
        const row = $('#today-row');
        if (row && sc && ui.layout === 'grid') { const th = sc.querySelector('thead'); sc.scrollTop = Math.max(0, row.offsetTop - (th ? th.offsetHeight : 70) - 2); }
      } else if (sc && keep) { sc.scrollTop = keep[0]; sc.scrollLeft = keep[1]; }
    }
    if (focusId) {
      const el = document.getElementById(focusId);
      if (el && el !== document.activeElement) { el.focus({ preventScroll: true }); if (sel && typeof el.setSelectionRange === 'function') { try { el.setSelectionRange(sel[0], sel[1]); } catch (e) { /* not a text field */ } } }
    }
    if (ui.dialog) { const first = document.querySelector('.dialog input, .dialog select, .dialog textarea'); if (first && !document.querySelector('.dialog').contains(document.activeElement)) first.focus(); }
  }
  store.subscribe(render);

  function go(view) {
    ui.view = view; ui.editing = null; ui.confirm = null; ui.rulesDraft = null;
    if (view !== 'request') ui.submitted = null;
    if (view === 'rota') ui.scrollToToday = true;
    try { sessionStorage.setItem('haem-rota-view', view); } catch (e) { /* storage blocked */ }
    render();
    window.scrollTo(0, 0);
    $('#view').focus({ preventScroll: true });
    if (view === 'mine' && ui.mine.doctorId && ui.mine.email && !ui.mine.result) loadMine();
  }

  // ---------- events ----------
  document.addEventListener('click', async (ev) => {
    const t = ev.target.closest('[data-act]');
    if (!t || t.tagName === 'SELECT' || (t.tagName === 'INPUT' && t.type !== 'button')) return;
    const act = t.dataset.act, val = t.dataset.val;
    const confirmStep = (key) => { if (ui.confirm === key) { ui.confirm = null; return true; } ui.confirm = key; render(); return false; };
    if (!['remove-doc', 'start-fresh', 'withdraw', 'revoke'].includes(act)) ui.confirm = null;
    switch (act) {
      case 'nav': go(t.dataset.view); return;
      case 'open-signin': ui.dialog = 'signin'; break;
      case 'open-setup': ui.dialog = 'setup'; break;
      case 'close-dialog': closeOverlay(() => { ui.dialog = null; }); return;
      case 'sign-out': await store.logout(); toast('Signed out.'); go('rota'); return;
      case 'day-all': ui.days = []; break;
      case 'day-toggle': { const k = Number(val); ui.days = ui.days.includes(k) ? ui.days.filter((x) => x !== k) : ui.days.concat(k).sort(); break; }
      case 'rota-more': ui.end = R.addDays(R.addMonths(R.addDays(ui.end, 1), 2), -1); break;
      case 'rota-earlier': ui.scrollToIso = ui.start; ui.start = R.addMonths(ui.start, -2); break;
      case 'rota-today': resetRange(TODAY); ui.scrollToToday = true; break;
      case 'layout': ui.layout = val; ui.scrollToToday = true; break;
      case 'group': ui.group = val; break;
      case 'hi': ui.hi = val; break;
      case 'export': { const { from, to } = rotaRange(); saveFile(`haematology-rota-${from}-to-${to}.xlsx`, E.rosterWorkbook(st(), from, to)); return; }
      case 'edit-cell': ui.editing = { doc: t.dataset.doc, iso: t.dataset.iso }; break;
      case 'edit-close': closeOverlay(() => { ui.editing = null; }); return;
      case 'edit-save': {
        const code = $('#edit-code').value, reason = $('#edit-reason').value.trim();
        const { doc, iso } = ui.editing;
        if (await change(op('setCell', doc, iso, code, reason), 'Rota updated and logged.')) ui.editing = null;
        break;
      }
      case 'edit-clear': { const { doc, iso } = ui.editing; if (await change(op('setCell', doc, iso, null, 'Undo'), 'Change undone.')) ui.editing = null; break; }
      case 'edt-book-cell': {
        const { doc, iso } = ui.editing;
        const chk = R.edtCheck(st(), doc, iso, t.dataset.session);
        let reason = '';
        if (!chk.ok) { reason = window.prompt('This breaks a rule. Why can it be booked anyway?') || ''; if (!reason.trim()) return; }
        await change(op('bookEdt', doc, iso, t.dataset.session, reason), `${chk.type || 'EDT'} booked.`);
        break;
      }
      case 'edt-remove': await change(op('removeEdt', t.dataset.id), 'Session removed.'); break;
      case 'form-reset': ui.form = defaultForm(); break;
      case 'form-new': ui.submitted = null; ui.form = defaultForm(); break;
      case 'withdraw': if (!confirmStep('wd:' + t.dataset.id)) return; try { await store.withdraw(t.dataset.id, ui.mine.email.trim()); toast('Request withdrawn.'); await loadMine(); } catch (e) { toast(e.message, true); } return;
      case 'pick-req': ui.reqSel = t.dataset.id; break;
      case 'req-filter': ui.reqFilter = val; ui.reqSel = null; break;
      case 'mark-read': await change((s) => { s.notifications.forEach((n) => { if (n.to === 'auth') n.read = true; }); return []; }); break;
      case 'decide': {
        const msg = $('#dec-msg').value;
        const days = [...document.querySelectorAll('.day-tick')].filter((x) => x.checked).map((x) => x.value);
        const words = { approve: 'Approved', approveDays: 'Approved for the ticked days', hold: 'Question sent', decline: 'Declined' };
        await change(op('decide', ui.reqSel, val, msg, { days }), `${words[val]}. The doctor has been told.`);
        break;
      }
      case 'revoke': {
        if (!confirmStep('revoke:' + t.dataset.id)) return;
        const id = t.dataset.id;
        await change((s, by) => {
          const r = s.requests.find((x) => x.id === id);
          r.status = 'cancelled';
          (r.cover || []).forEach((c) => { const o = (s.overrides[c.doctorId] || {})[c.date]; if (o && /^Covering/.test(o.reason || '')) delete s.overrides[c.doctorId][c.date]; });
          s.audit.unshift({ at: new Date().toISOString(), by, action: 'Approved leave cancelled', detail: `${docName(r.doctorId)}: ${R.fmtUk(r.start)} to ${R.fmtUk(r.end)}` });
          s.notifications.unshift({ id: 'n-' + Date.now(), to: 'd:' + r.doctorId, at: new Date().toISOString(), text: `Your leave ${rangeText(r.start, r.end)} was cancelled by ${by}.`, read: false });
          const d = R.doctorById(s, r.doctorId);
          return r.email || (d && d.email) ? [{ to: r.email || d.email, subject: `Leave cancelled: ${rangeText(r.start, r.end)}`, text: `Your approved leave (${rangeText(r.start, r.end)}) was cancelled by ${by}. Please contact the rota authorisers.` }] : [];
        }, 'Leave cancelled. The doctor has been told.');
        break;
      }
      case 'edt-session': ui.edt.session = val; break;
      case 'edt-book': {
        const f = ui.edt;
        const reason = $('#edt-reason') ? $('#edt-reason').value.trim() : '';
        await change(op('bookEdt', f.doctorId, f.date, f.session, reason), 'Session booked and shown on the rota.');
        break;
      }
      case 'rules-edit': ui.rulesDraft = { rules: Object.assign({}, R.DEFAULT_RULES, st().settings.rules), text: JSON.parse(JSON.stringify(st().settings.rulesText || [])) }; break;
      case 'rules-cancel': ui.rulesDraft = null; break;
      case 'rules-add': ui.rulesDraft.text.push({ id: 's' + Date.now(), title: 'New section', items: [] }); break;
      case 'rules-del': ui.rulesDraft.text.splice(Number(t.dataset.i), 1); break;
      case 'rules-save': {
        const d = ui.rulesDraft;
        const rules = {};
        for (const [k] of RULE_FIELDS) { const n = Number(d.rules[k]); if (!Number.isInteger(n) || n < 0 || n > 9) { toast('Each number must be a whole number from 0 to 9.', true); return; } rules[k] = n; }
        rules.reminderDays = R.ruleNum(st(), 'reminderDays');
        const text = d.text.map((x) => ({ id: x.id, title: String(x.title).trim(), items: x.items.map((i) => String(i).trim()).filter(Boolean) })).filter((x) => x.title);
        const ok = await change((s, by) => {
          const changed = Object.keys(rules).filter((k) => s.settings.rules[k] !== rules[k]).map((k) => `${k} ${s.settings.rules[k]} → ${rules[k]}`);
          s.settings.rules = rules; s.settings.rulesText = text;
          s.audit.unshift({ at: new Date().toISOString(), by, action: 'Rules changed', detail: changed.length ? changed.join(', ') : 'Text updated' });
          return [];
        }, 'Rules saved.');
        if (ok) ui.rulesDraft = null;
        break;
      }
      case 'setup-tab': ui.setupTab = val; ui.importPreview = null; break;
      case 'add-doc': {
        const nd = ui.nd;
        const grade = nd.grade === '__other' ? nd.other.trim() : nd.grade;
        if (!grade) { toast('Type the grade under "Other".', true); return; }
        const group = nd.grade === '__other' ? nd.group : groupForGrade(grade);
        const name = nd.placeholder ? `New ${grade}` : nd.name.trim();
        if (!name) { toast('Type the doctor\'s name, or tick "Name not known yet".', true); return; }
        if (nd.start && nd.end && nd.end < nd.start) { toast('The finish date is before the start date.', true); return; }
        const prev = nd.replaces ? R.doctorById(st(), nd.replaces) : null;
        const src = R.doctorById(st(), nd.pattern || nd.replaces);
        const fields = { name, grade, group, email: nd.placeholder ? '' : nd.email.trim(), start: nd.start, end: nd.end, endProvisional: !!nd.provisional, placeholder: !!nd.placeholder,
          slot: prev ? R.slotOf(prev) : undefined, plan: src && src.plan ? JSON.parse(JSON.stringify(src.plan)) : { A: ['', '', '', '', ''], B: null } };
        if (await change(op('saveDoctor', fields), `${name} added${prev ? ` in ${prev.name}'s column` : ''}.`)) ui.nd = null;
        break;
      }
      case 'add-placeholder': {
        const prev = R.doctorById(st(), t.dataset.doc);
        const grade = prev.grade.split(' (')[0];
        const sug = R.suggestDates(grade, R.addDays(prev.end, 1));
        const start = sug.start && sug.start <= R.addDays(prev.end, 7) ? sug.start : R.addDays(prev.end, 1);
        await change(op('saveDoctor', { name: `New ${grade}`, grade, group: prev.group, email: '', start, end: sug.end || '', endProvisional: true, placeholder: true, slot: R.slotOf(prev), plan: JSON.parse(JSON.stringify(prev.plan || { A: ['', '', '', '', ''], B: null })) }), `Placeholder added after ${prev.name}. Fill in the name on the Master sheet when known.`);
        break;
      }
      case 'bulk-apply': {
        const b = ui.bulk;
        await change(op('bulkAssign', b.ids, b.from, b.to, b.days, b.code, b.reason), 'Rota updated for those dates.');
        break;
      }
      case 'cons-more': ui.consExtra++; break;
      case 'theme': {
        document.documentElement.setAttribute('data-theme', val);
        try { localStorage.setItem('haem-rota-theme', val); } catch (e) { /* storage blocked */ }
        renderHead();
        return;
      }
      case 'quick-code': { const sel = $('#edit-code'); if (sel) sel.value = val; return; }
      case 'use-al': { const d = R.doctorById(st(), t.dataset.doc); await change(op('saveDoctor', { id: d.id, alEntitlement: R.alProRata(d).days })); break; }
      case 'remove-doc': if (!confirmStep('remove:' + t.dataset.doc)) return; await change(op('removeDoctor', t.dataset.doc), 'Doctor removed.'); break;
      case 'template': {
        const kind = t.dataset.kind;
        const bytes = kind === 'oncall' ? E.oncallTemplate(TODAY) : kind === 'plan' ? E.planTemplate(st()) : E.staffTemplate(st());
        saveFile({ oncall: 'oncall-import-template.xlsx', plan: 'weekly-plan-template.xlsx', staff: 'staff-list-template.xlsx' }[kind], bytes);
        return;
      }
      case 'cancel-import': ui.importPreview = null; break;
      case 'confirm-import': await confirmImport(); break;
      case 'save-rotation': await change(op('setSetting', 'rotation', { start: $('#rot-start').value, end: $('#rot-end').value }, `Rotation dates: ${$('#rot-start').value} to ${$('#rot-end').value}`), 'Rotation dates saved.'); break;
      case 'add-changeover': await change(op('addChangeover', $('#co-date').value, $('#co-label').value), 'Changeover added.'); break;
      case 'remove-changeover': await change(op('removeChangeover', t.dataset.id)); break;
      case 'add-restr': await change(op('addRestricted', $('#rs-start').value, $('#rs-end').value, $('#rs-level').value, $('#rs-note').value.trim()), 'Restriction added.'); break;
      case 'remove-restr': await change(op('removeRestricted', t.dataset.id)); break;
      case 'save-authorisers': {
        const names = $('#set-auth').value.split('\n').map((x) => x.trim()).filter(Boolean);
        const mails = $('#set-mail').value.split('\n').map((x) => x.trim()).filter(Boolean);
        const emailDocs = $('#set-email-docs').checked;
        if (!names.length) { toast('Keep at least one authoriser name.', true); return; }
        if (!names.includes(store.me)) { toast(`Keep your own name (${store.me}) in the list, or you will be locked out of signing in by name.`, true); return; }
        if (mails.some((m) => !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m))) { toast('One of the alert emails does not look right.', true); return; }
        await change((s, by) => { S.ops.setSetting(s, by, 'authorisers', names, `Authorisers: ${names.join(', ')}`); S.ops.setSetting(s, by, 'authoriserEmails', mails, `Alert emails: ${mails.length}`); S.ops.setSetting(s, by, 'emailDoctors', emailDocs, `Email doctors: ${emailDocs ? 'on' : 'off'}`); return []; }, 'Saved.');
        break;
      }
      case 'save-bh': {
        const list = $('#set-bh').value.split(/\s+/).map((x) => x.trim()).filter(Boolean);
        if (list.some((x) => !R.isIso(x))) { toast('Write each date as yyyy-mm-dd.', true); return; }
        await change(op('setSetting', 'bankHolidays', list.sort(), `Bank holidays: ${list.length}`), 'Bank holidays saved.');
        break;
      }
      case 'change-pw': {
        const cur = $('#pw-cur').value, next = $('#pw-new').value;
        try { await store.changePassword(cur, next); toast('Password changed. Please sign in again.'); go('rota'); } catch (e) { toast(e.message, true); }
        return;
      }
      case 'start-fresh': if (!confirmStep('fresh')) return; await change(op('startFresh'), 'Example data removed. Add your doctors under People.'); ui.setupTab = 'people'; break;
      case 'reset-demo': store.resetLocalDemo(); toast('Example data reloaded.'); return;
      default: return;
    }
    render();
  });

  document.addEventListener('change', async (ev) => {
    const t = ev.target;
    if (t.dataset.form && ui.form) {
      const k = t.dataset.form;
      ui.form[k] = t.type === 'checkbox' ? t.checked : t.value;
      if (t.tagName === 'TEXTAREA' || t.type === 'text' || t.type === 'email') return; // typed text doesn't change the checks; re-rendering here would swallow the next click
      if (k === 'start' && (!ui.form.end || ui.form.end < t.value)) ui.form.end = t.value;
      if (k === 'doctorId') { if (ui.form.dayUnitCoverId === t.value) ui.form.dayUnitCoverId = ''; if (ui.form.guhCoverId === t.value) ui.form.guhCoverId = ''; }
      render(); return;
    }
    if (t.dataset.mine) { ui.mine[t.dataset.mine] = t.value; ui.mine.result = null; return; }
    if (t.dataset.edt) { ui.edt[t.dataset.edt] = t.value; render(); return; }
    if (t.dataset.nd && ui.nd) {
      const k = t.dataset.nd;
      ui.nd[k] = t.type === 'checkbox' ? t.checked : t.value;
      if (k === 'grade' || k === 'replaces' || (k === 'other' && ui.nd.grade === '__other')) {
        ui.nd.suggested = false;
        if (k === 'replaces' && t.value) { const prev = R.doctorById(st(), t.value); if (prev) { const g = prev.grade.split(' (')[0]; if (groupForGrade(g)) ui.nd.grade = g; else { ui.nd.grade = '__other'; ui.nd.other = g; ui.nd.group = prev.group; } } }
      }
      if (t.type === 'text' || t.type === 'email') return;
      render(); return;
    }
    if (t.dataset.bulk && ui.bulk) { ui.bulk[t.dataset.bulk] = t.value; return; }
    if (t.dataset.bulkId && ui.bulk) { const id = t.dataset.bulkId; ui.bulk.ids = ui.bulk.ids.filter((x) => x !== id).concat(t.checked ? [id] : []); return; }
    if (t.dataset.bulkDay && ui.bulk) { const k = Number(t.dataset.bulkDay); ui.bulk.days = ui.bulk.days.filter((x) => x !== k).concat(t.checked ? [k] : []); return; }
    if (t.dataset.rule && ui.rulesDraft) { ui.rulesDraft.rules[t.dataset.rule] = t.value; return; }
    const act = t.dataset.act;
    switch (act) {
      case 'rota-goto': if (R.isIso(t.value)) { resetRange(t.value); ui.scrollToToday = false; } break;
      case 'hide-weekends': ui.hideWeekends = t.checked; break;
      case 'doc-field': {
        const d = R.doctorById(st(), t.dataset.doc);
        const f = t.dataset.field;
        let v = t.value;
        if (['alEntitlement', 'slEntitlement', 'alExtra', 'fte'].includes(f)) v = Number(v) || 0;
        if (t.type === 'checkbox') v = t.checked;
        if (f === 'edtHours') v = t.value === '' ? undefined : Number(t.value);
        if (f === 'aliases') v = v.split(',').map((x) => x.trim()).filter(Boolean);
        if (f === 'name' && !String(v).trim()) { toast('A name is required.', true); render(); return; }
        await change(op('saveDoctor', { id: d.id, [f]: v }));
        return;
      }
      case 'doc-grade': {
        const d = R.doctorById(st(), t.dataset.doc);
        let grade = t.value;
        if (grade === '__other') { grade = (window.prompt('Grade for ' + d.name, groupForGrade(d.grade) ? '' : d.grade) || '').trim(); if (!grade) { render(); return; } }
        await change(op('saveDoctor', { id: d.id, grade, group: groupForGrade(grade) || d.group }));
        return;
      }
      case 'plan-day': await change(op('setPlanDay', t.dataset.doc, t.dataset.week, Number(t.dataset.i), t.value)); return;
      case 'toggle-b': await change(op('toggleWeekB', t.dataset.doc, t.checked)); return;
      case 'cons': { const w = t.dataset.week, v = t.value.trim(); await change((s, by) => S.ops.setSetting(s, by, 'consultantOfWeek.' + w, v, `Consultant w/c ${R.fmtUk(w)}: ${v || '(blank)'}`)); return; }
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
    if (t.dataset.form && ui.form && (t.tagName === 'TEXTAREA' || t.type === 'text' || t.type === 'email')) ui.form[t.dataset.form] = t.value; // keep typing without re-render
    if (t.dataset.mine) ui.mine[t.dataset.mine] = t.value;
    if (t.dataset.nd && ui.nd && (t.type === 'text' || t.type === 'email')) ui.nd[t.dataset.nd] = t.value;
    if (t.dataset.bulk === 'reason' && ui.bulk) ui.bulk.reason = t.value;
    if (ui.rulesDraft) {
      if (t.dataset.rtitle !== undefined) ui.rulesDraft.text[Number(t.dataset.rtitle)].title = t.value;
      if (t.dataset.ritems !== undefined) ui.rulesDraft.text[Number(t.dataset.ritems)].items = t.value.split('\n');
      if (t.dataset.rule) ui.rulesDraft.rules[t.dataset.rule] = t.value;
    }
  });

  document.addEventListener('submit', async (ev) => {
    ev.preventDefault();
    const id = ev.target.id;
    if (id === 'leave-form') return submitForm();
    if (id === 'mine-form') return loadMine();
    if (id === 'signin-form') {
      try { await store.login($('#si-name').value, $('#si-pw') ? $('#si-pw').value : ''); ui.dialog = null; toast(`Signed in as ${store.me}.`); render(); } catch (e) { toast(e.message, true); }
      return;
    }
    if (id === 'setup-form') {
      const pw = $('#su-pw').value, pw2 = $('#su-pw2').value;
      if (pw !== pw2) return toast('The two passwords are different.', true);
      const names = $('#su-names').value.split('\n').map((x) => x.trim()).filter(Boolean);
      if (!names.length) return toast('Add at least one authoriser name.', true);
      try { await store.setupPassword(pw, $('#su-code') ? $('#su-code').value : '', names); ui.dialog = 'signin'; toast('Password set. Now sign in.'); render(); } catch (e) { toast(e.message, true); }
    }
  });

  document.addEventListener('toggle', (ev) => { if (ev.target.classList && ev.target.classList.contains('filters')) ui.filtersOpen = ev.target.open; }, true);

  document.addEventListener('keydown', (ev) => {
    if (ev.key !== 'Escape') return;
    if (ui.dialog) closeOverlay(() => { ui.dialog = null; }); else if (ui.editing) closeOverlay(() => { ui.editing = null; });
  });

  async function confirmImport() {
    const p = ui.importPreview;
    const resolve = (x) => x.doctorId || p.mapping[x.name] || null;
    const newAliases = {};
    for (const [name, id] of Object.entries(p.mapping)) if (id) newAliases[name] = id;
    let payload;
    if (p.kind === 'oncall') payload = { entries: p.parsed.entries.map((e) => Object.assign({}, e, { doctorId: resolve(e) })).filter((e) => e.doctorId), from: p.parsed.from, to: p.parsed.to, fileName: p.fileName, newAliases };
    else if (p.kind === 'plan') payload = { plans: p.parsed.plans.map((x) => Object.assign({}, x, { doctorId: resolve(x) })).filter((x) => x.doctorId), fileName: p.fileName, newAliases };
    else payload = { staff: p.parsed.staff, fileName: p.fileName, newAliases };
    if (await change(op('applyImport', p.kind, payload), `Imported ${p.fileName}.`)) ui.importPreview = null;
  }

  // Keep the rota up to date when several people use it (only when nothing is being edited).
  setInterval(() => {
    if (store.mode !== 'remote' || document.hidden || ui.busy || ui.editing || ui.dialog || ui.rulesDraft) return;
    const a = document.activeElement;
    if (a && /INPUT|TEXTAREA|SELECT/.test(a.tagName)) return;
    store.refresh();
  }, 60000);
  document.addEventListener('visibilitychange', () => { if (!document.hidden && store.mode === 'remote' && !ui.busy) store.refresh(); });

  $('#view').innerHTML = '<p class="muted" style="padding:20px">Loading the rota…</p>';
  store.init().then(() => {
    if (store.mode === 'remote' && !store.hasPassword) ui.dialog = null;
    render();
  });
})();
