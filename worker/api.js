/* The rota server: a small JSON API on Cloudflare (Workers or Pages Functions) with a D1 database.
 * Anyone with the link can read the public roster and send a leave request.
 * Authorisers sign in with the shared password (stored only as a salted PBKDF2 hash) and can change everything. */
import R from '../app/js/rules.js';
import S from '../app/js/store.js';
import E from '../app/js/excel.js';

const SESSION_COOKIE = 'hr_session';
const SESSION_HOURS = 12;
const PBKDF2_ITERATIONS = 100000; // the most Cloudflare Workers allows
const MAX_LOGIN_FAILS = 5;
const LOCK_MINUTES = 15;
const MAX_REQUESTS_PER_HOUR = 20;
const MAX_STATE_BYTES = 1800000; // D1 allows about 2 MB per value
const MAX_EMAILS_PER_CALL = 25;

const enc = new TextEncoder();
const hex = (buf) => [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');
const randomHex = (n) => hex(crypto.getRandomValues(new Uint8Array(n)));
const sha256 = async (s) => hex(await crypto.subtle.digest('SHA-256', enc.encode(s)));
const todayIso = () => new Date().toISOString().slice(0, 10);

async function pbkdf2(password, saltHex) {
  const key = await crypto.subtle.importKey('raw', enc.encode(password), 'PBKDF2', false, ['deriveBits']);
  const salt = new Uint8Array(saltHex.match(/../g).map((h) => parseInt(h, 16)));
  return hex(await crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt, iterations: PBKDF2_ITERATIONS }, key, 256));
}
function sameText(a, b) { // constant-time compare for equal-length hex strings
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function json(data, status, extraHeaders) {
  return new Response(JSON.stringify(data), {
    status: status || 200,
    headers: Object.assign({ 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' }, extraHeaders || {})
  });
}
const fail = (status, error, extra) => json(Object.assign({ error }, extra || {}), status);

// ---------- database ----------
let schemaReady = false;
async function ensureSchema(db) {
  if (schemaReady) return;
  await db.batch([
    db.prepare('CREATE TABLE IF NOT EXISTS kv (id TEXT PRIMARY KEY, json TEXT NOT NULL, version INTEGER NOT NULL, updated_at TEXT NOT NULL)'),
    db.prepare('CREATE TABLE IF NOT EXISTS auth (id TEXT PRIMARY KEY, salt TEXT NOT NULL, hash TEXT NOT NULL, set_at TEXT NOT NULL)'),
    db.prepare('CREATE TABLE IF NOT EXISTS sessions (token_hash TEXT PRIMARY KEY, name TEXT NOT NULL, expires_at INTEGER NOT NULL)'),
    db.prepare('CREATE TABLE IF NOT EXISTS hits (key TEXT NOT NULL, at INTEGER NOT NULL)'),
    db.prepare('CREATE INDEX IF NOT EXISTS hits_key ON hits (key, at)'),
    db.prepare('CREATE TABLE IF NOT EXISTS backups (id INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT NOT NULL, kind TEXT NOT NULL, by TEXT, note TEXT, version INTEGER, bytes INTEGER, data BLOB NOT NULL)'),
    db.prepare('CREATE INDEX IF NOT EXISTS backups_kind ON backups (kind, id)'),
    db.prepare('CREATE TABLE IF NOT EXISTS csv_backups (id INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT NOT NULL, day TEXT NOT NULL, bytes INTEGER, data BLOB NOT NULL)')
  ]);
  schemaReady = true;
}
async function loadState(db) {
  const row = await db.prepare("SELECT json, version FROM kv WHERE id = 'state'").first();
  if (row) return { state: S.upgrade(JSON.parse(row.json)), version: row.version };
  const state = S.demoState(todayIso());
  await db.prepare("INSERT OR IGNORE INTO kv (id, json, version, updated_at) VALUES ('state', ?, 1, ?)").bind(JSON.stringify(state), new Date().toISOString()).run();
  const again = await db.prepare("SELECT json, version FROM kv WHERE id = 'state'").first();
  return { state: JSON.parse(again.json), version: again.version };
}
function trim(state) {
  if (state.audit && state.audit.length > 2000) state.audit.length = 2000;
  if (state.notifications && state.notifications.length > 500) state.notifications.length = 500;
}
// ---------- backups: a compressed copy after every change, plus one each day ----------
const KEEP_CHANGE_BACKUPS = 500;
const KEEP_DAILY_BACKUPS = 730;
const KEEP_CSV_BACKUPS = 730;
async function gzip(text) {
  const cs = new CompressionStream('gzip');
  const w = cs.writable.getWriter(); w.write(typeof text === 'string' ? enc.encode(text) : text); w.close();
  return new Uint8Array(await new Response(cs.readable).arrayBuffer());
}
async function gunzip(bytes) {
  const ds = new DecompressionStream('gzip');
  const w = ds.writable.getWriter(); w.write(bytes); w.close();
  return new Response(ds.readable).text();
}
async function snapshot(db, kind, body, version, by, note) {
  const data = await gzip(body);
  await db.prepare('INSERT INTO backups (at, kind, by, note, version, bytes, data) VALUES (?, ?, ?, ?, ?, ?, ?)')
    .bind(new Date().toISOString(), kind, String(by || '').slice(0, 80), String(note || '').slice(0, 300), version, body.length, data).run();
  const keep = kind === 'daily' ? KEEP_DAILY_BACKUPS : KEEP_CHANGE_BACKUPS;
  await db.prepare('DELETE FROM backups WHERE kind = ? AND id NOT IN (SELECT id FROM backups WHERE kind = ? ORDER BY id DESC LIMIT ?)').bind(kind, kind, keep).run();
}
export async function dailyBackup(env) {
  if (!env.DB) return;
  await ensureSchema(env.DB);
  const day = todayIso();
  const done = await env.DB.prepare("SELECT id FROM backups WHERE kind = 'daily' AND at >= ? LIMIT 1").bind(day).first();
  const row = await env.DB.prepare("SELECT json, version FROM kv WHERE id = 'state'").first();
  if (!row) return;
  if (!done) await snapshot(env.DB, 'daily', row.json, row.version, 'Daily backup', day);
  // A CSV copy too (one CSV per table, zipped), so there is a nightly copy that opens in Excel.
  const csvDone = await env.DB.prepare('SELECT id FROM csv_backups WHERE day = ? LIMIT 1').bind(day).first();
  if (!csvDone) {
    const zip = E.csvBackupZip(S.upgrade(JSON.parse(row.json)), day);
    await env.DB.prepare('INSERT INTO csv_backups (at, day, bytes, data) VALUES (?, ?, ?, ?)').bind(new Date().toISOString(), day, zip.length, await gzip(zip)).run();
    await env.DB.prepare('DELETE FROM csv_backups WHERE id NOT IN (SELECT id FROM csv_backups ORDER BY id DESC LIMIT ?)').bind(KEEP_CSV_BACKUPS).run();
  }
}

// Save only if nobody else saved since we read it. Returns the new version, or null on a clash.
async function saveState(db, state, baseVersion) {
  trim(state);
  const body = JSON.stringify(state);
  if (body.length > MAX_STATE_BYTES) throw Object.assign(new Error('The rota has grown too large to save. Remove old on-call data or ask for help.'), { status: 413 });
  const res = await db.prepare("UPDATE kv SET json = ?, version = version + 1, updated_at = ? WHERE id = 'state' AND version = ?").bind(body, new Date().toISOString(), baseVersion).run();
  if (!(res.meta && res.meta.changes === 1)) return null;
  const last = (state.audit && state.audit[0]) || {};
  try { await snapshot(db, 'change', body, baseVersion + 1, last.by, [last.action, last.detail].filter(Boolean).join(': ')); } catch (e) { console.error('backup failed', e); }
  return baseVersion + 1;
}
// Read, change, save; retry if someone else saved at the same moment.
async function mutate(db, fn) {
  for (let i = 0; i < 4; i++) {
    const { state, version } = await loadState(db);
    const out = fn(state);
    const v = await saveState(db, state, version);
    if (v) return { out, state, version: v };
  }
  throw Object.assign(new Error('The rota is busy. Please try again.'), { status: 503 });
}

// Simple rate limits stored in D1 (per IP).
async function countHits(db, key, sinceMs) {
  const row = await db.prepare('SELECT COUNT(*) AS n FROM hits WHERE key = ? AND at > ?').bind(key, Date.now() - sinceMs).first();
  return row ? row.n : 0;
}
const addHit = (db, key) => db.prepare('INSERT INTO hits (key, at) VALUES (?, ?)').bind(key, Date.now()).run();

// ---------- sessions ----------
function cookieValue(request, name) {
  const raw = request.headers.get('Cookie') || '';
  const m = raw.split(/;\s*/).find((c) => c.startsWith(name + '='));
  return m ? decodeURIComponent(m.slice(name.length + 1)) : '';
}
async function currentUser(db, request) {
  const token = cookieValue(request, SESSION_COOKIE);
  if (!token || token.length !== 64) return null;
  const row = await db.prepare('SELECT name, expires_at FROM sessions WHERE token_hash = ?').bind(await sha256(token)).first();
  if (!row || row.expires_at < Date.now()) return null;
  return row.name;
}
function sessionCookie(token, maxAgeSeconds, secure) {
  return `${SESSION_COOKIE}=${token}; Path=/; HttpOnly; SameSite=Strict; Max-Age=${maxAgeSeconds}${secure ? '; Secure' : ''}`;
}

// ---------- email (optional, via Resend) ----------
function allowedRecipients(state) {
  const set = new Set();
  const add = (e) => { if (e && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) set.add(String(e).trim().toLowerCase()); };
  // Doctors are emailed only if the authorisers switch it on (Set up → Settings); otherwise only alert emails go out.
  if (state.settings && state.settings.emailDoctors) {
    (state.doctors || []).forEach((d) => add(d.email));
    (state.requests || []).forEach((r) => add(r.email));
  }
  ((state.settings && state.settings.authoriserEmails) || []).forEach(add);
  return set;
}
async function sendEmails(env, state, emails) {
  if (!env.RESEND_API_KEY || !env.MAIL_FROM || !emails || !emails.length) return 0;
  const ok = allowedRecipients(state);
  const list = emails.filter((m) => m && m.to && ok.has(String(m.to).trim().toLowerCase())).slice(0, MAX_EMAILS_PER_CALL);
  const link = env.SITE_URL ? `\n\nOpen the rota: ${env.SITE_URL}` : '';
  const results = await Promise.allSettled(list.map((m) => fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { Authorization: `Bearer ${env.RESEND_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ from: env.MAIL_FROM, to: [m.to], subject: String(m.subject).slice(0, 200), text: String(m.text).slice(0, 5000) + link + '\n\n(Haematology rota. Please do not reply all.)' })
  })));
  return results.filter((r) => r.status === 'fulfilled' && r.value.ok).length;
}

// ---------- shape checks for a state sent by an authoriser ----------
function looksLikeState(s) {
  return s && typeof s === 'object' && s.version === 2 && s.settings && typeof s.settings === 'object' &&
    Array.isArray(s.doctors) && Array.isArray(s.requests) && Array.isArray(s.edt) && Array.isArray(s.audit) && Array.isArray(s.notifications) &&
    s.oncall && typeof s.oncall === 'object' && s.overrides && typeof s.overrides === 'object' && !s.public;
}

async function readJson(request) {
  const text = await request.text();
  if (text.length > 2500000) throw Object.assign(new Error('Too much data.'), { status: 413 });
  try { return text ? JSON.parse(text) : {}; } catch (e) { throw Object.assign(new Error('That was not valid JSON.'), { status: 400 }); }
}

// ---------- the API ----------
export async function handleApi(request, env) {
  const url = new URL(request.url);
  const path = url.pathname.replace(/\/+$/, '');
  const method = request.method;
  if (!env.DB) return fail(503, 'The database is not connected yet. In Cloudflare, add a D1 database binding called DB to this project, then redeploy.', { mode: 'server-unconfigured' });
  const db = env.DB;
  const secure = url.protocol === 'https:';
  const ip = request.headers.get('CF-Connecting-IP') || 'local';

  // Block cross-site form posts: changes must come from this site.
  if (method !== 'GET' && method !== 'HEAD') {
    const origin = request.headers.get('Origin');
    if (origin && origin !== url.origin) return fail(403, 'Requests must come from the rota website.');
  }

  try {
    await ensureSchema(db);
    const me = await currentUser(db, request);

    if (path === '/api/state' && method === 'GET') {
      const { state, version } = await loadState(db);
      const auth = await db.prepare("SELECT id FROM auth WHERE id = 'authoriser'").first();
      return json({
        mode: 'server', version, me, hasPassword: !!auth, emailOn: !!(env.RESEND_API_KEY && env.MAIL_FROM),
        setupCodeRequired: !!env.SETUP_CODE, state: me ? state : S.publicView(state)
      });
    }

    if (path === '/api/me' && method === 'GET') return json({ me });

    // First-time setup: choose the authoriser password (only while none is set).
    if (path === '/api/setup' && method === 'POST') {
      const body = await readJson(request);
      const existing = await db.prepare("SELECT id FROM auth WHERE id = 'authoriser'").first();
      if (existing) return fail(409, 'The authoriser password has already been set. Sign in instead.');
      if (env.SETUP_CODE && !sameText(await sha256(String(body.setupCode || '')), await sha256(env.SETUP_CODE))) return fail(403, 'The setup code is not right.');
      const pw = String(body.password || '');
      if (pw.length < 12) return fail(400, 'Use at least 12 characters. A short phrase of 3 or 4 unrelated words is good.');
      const salt = randomHex(16);
      await db.prepare("INSERT INTO auth (id, salt, hash, set_at) VALUES ('authoriser', ?, ?, ?)").bind(salt, await pbkdf2(pw, salt), new Date().toISOString()).run();
      const names = Array.isArray(body.names) ? body.names.map((n) => String(n).trim().slice(0, 60)).filter(Boolean).slice(0, 10) : [];
      await mutate(db, (state) => {
        if (names.length) state.settings.authorisers = names;
        state.audit.unshift({ at: new Date().toISOString(), by: 'Setup', action: 'Authoriser password set', detail: names.length ? `Authorisers: ${names.join(', ')}` : '' });
      });
      return json({ ok: true });
    }

    if (path === '/api/login' && method === 'POST') {
      const body = await readJson(request);
      const key = 'login:' + ip;
      if ((await countHits(db, key, LOCK_MINUTES * 60000)) >= MAX_LOGIN_FAILS) return fail(429, `Too many wrong passwords. Try again in ${LOCK_MINUTES} minutes.`);
      const row = await db.prepare("SELECT salt, hash FROM auth WHERE id = 'authoriser'").first();
      if (!row) return fail(409, 'No authoriser password has been set yet. Use first-time setup.');
      const { state } = await loadState(db);
      const name = String(body.name || '').trim();
      if (!(state.settings.authorisers || []).includes(name)) return fail(400, 'Choose your name from the list of authorisers.');
      if (!sameText(await pbkdf2(String(body.password || ''), row.salt), row.hash)) {
        await addHit(db, key);
        return fail(401, 'That password is not right.');
      }
      const token = randomHex(32);
      await db.batch([
        db.prepare('DELETE FROM sessions WHERE expires_at < ?').bind(Date.now()),
        db.prepare('DELETE FROM hits WHERE at < ?').bind(Date.now() - 86400000),
        db.prepare('INSERT INTO sessions (token_hash, name, expires_at) VALUES (?, ?, ?)').bind(await sha256(token), name, Date.now() + SESSION_HOURS * 3600000)
      ]);
      return json({ ok: true, me: name }, 200, { 'Set-Cookie': sessionCookie(token, SESSION_HOURS * 3600, secure) });
    }

    if (path === '/api/logout' && method === 'POST') {
      const token = cookieValue(request, SESSION_COOKIE);
      if (token) await db.prepare('DELETE FROM sessions WHERE token_hash = ?').bind(await sha256(token)).run();
      return json({ ok: true }, 200, { 'Set-Cookie': sessionCookie('', 0, secure) });
    }

    if (path === '/api/password' && method === 'POST') {
      if (!me) return fail(401, 'Sign in as an authoriser first.');
      const body = await readJson(request);
      const row = await db.prepare("SELECT salt, hash FROM auth WHERE id = 'authoriser'").first();
      if (!sameText(await pbkdf2(String(body.current || ''), row.salt), row.hash)) return fail(401, 'The current password is not right.');
      const pw = String(body.next || '');
      if (pw.length < 12) return fail(400, 'Use at least 12 characters.');
      const salt = randomHex(16);
      await db.batch([
        db.prepare("UPDATE auth SET salt = ?, hash = ?, set_at = ? WHERE id = 'authoriser'").bind(salt, await pbkdf2(pw, salt), new Date().toISOString()),
        db.prepare('DELETE FROM sessions')
      ]);
      await mutate(db, (state) => { state.audit.unshift({ at: new Date().toISOString(), by: me, action: 'Authoriser password changed', detail: 'Everyone was signed out' }); });
      return json({ ok: true }, 200, { 'Set-Cookie': sessionCookie('', 0, secure) });
    }

    // A doctor sends a leave request or reports sickness (no password).
    if (path === '/api/requests' && method === 'POST') {
      const key = 'req:' + ip;
      if ((await countHits(db, key, 3600000)) >= MAX_REQUESTS_PER_HOUR) return fail(429, 'Too many requests from this connection. Try again later.');
      const body = await readJson(request);
      await addHit(db, key);
      const { out, state } = await mutate(db, (s) => S.submitLeave(s, body));
      const sent = await sendEmails(env, state, out.emails);
      return json({ ok: true, request: S.publicView({ ...state, requests: [out.request] }).requests[0], check: out.check, emailed: sent > 0 });
    }

    if (path === '/api/my-requests' && method === 'POST') {
      const key = 'look:' + ip;
      if ((await countHits(db, key, 600000)) >= 30) return fail(429, 'Too many look-ups. Try again in a few minutes.');
      await addHit(db, key);
      const body = await readJson(request);
      const { state } = await loadState(db);
      return json(S.myRequests(state, body.doctorId, body.email));
    }

    const m = path.match(/^\/api\/requests\/([\w-]+)\/withdraw$/);
    if (m && method === 'POST') {
      const body = await readJson(request);
      const { state } = await mutate(db, (s) => S.withdrawLeave(s, m[1], body.email));
      await sendEmails(env, state, (state.settings.authoriserEmails || []).map((to) => ({ to, subject: 'Leave request withdrawn', text: 'A doctor withdrew a leave request. Open the rota for details.' })));
      return json({ ok: true });
    }

    // Backups (authorisers only): list, and fetch one to download or restore.
    if (path === '/api/backups' && method === 'GET') {
      if (!me) return fail(401, 'Sign in as an authoriser first.');
      const changes = await db.prepare("SELECT id, at, kind, by, note, version, bytes FROM backups WHERE kind != 'daily' ORDER BY id DESC LIMIT 200").all();
      const daily = await db.prepare("SELECT id, at, kind, by, note, version, bytes FROM backups WHERE kind = 'daily' ORDER BY id DESC LIMIT 400").all();
      const csv = await db.prepare('SELECT id, at, day, bytes FROM csv_backups ORDER BY id DESC LIMIT 400').all();
      return json({ changes: changes.results || [], daily: daily.results || [], csv: csv.results || [] });
    }
    const mc = path.match(/^\/api\/csv-backups\/(\d+)$/);
    if (mc && method === 'GET') {
      if (!me) return fail(401, 'Sign in as an authoriser first.');
      const row = await db.prepare('SELECT day, data FROM csv_backups WHERE id = ?').bind(Number(mc[1])).first();
      if (!row) return fail(404, 'That backup no longer exists.');
      const ds = new DecompressionStream('gzip');
      const w = ds.writable.getWriter(); w.write(new Uint8Array(row.data)); w.close();
      return new Response(ds.readable, { headers: { 'Content-Type': 'application/zip', 'Content-Disposition': `attachment; filename="haematology-rota-csv-${row.day}.zip"`, 'Cache-Control': 'no-store' } });
    }
    const mb = path.match(/^\/api\/backups\/(\d+)$/);
    if (mb && method === 'GET') {
      if (!me) return fail(401, 'Sign in as an authoriser first.');
      const row = await db.prepare('SELECT id, at, kind, version, data FROM backups WHERE id = ?').bind(Number(mb[1])).first();
      if (!row) return fail(404, 'That backup no longer exists.');
      const text = await gunzip(new Uint8Array(row.data));
      return json({ id: row.id, at: row.at, kind: row.kind, version: row.version, state: JSON.parse(text) });
    }

    // An authoriser saves the whole rota (after changing it in the browser).
    if (path === '/api/state' && method === 'PUT') {
      if (!me) return fail(401, 'Your sign-in has expired. Sign in again.');
      const body = await readJson(request);
      if (!looksLikeState(body.state)) return fail(400, 'That does not look like a rota.');
      const v = await saveState(db, body.state, Number(body.baseVersion));
      if (!v) {
        const cur = await loadState(db);
        return json({ error: 'Someone else changed the rota at the same time.', mode: 'server', me, version: cur.version, state: cur.state, hasPassword: true, emailOn: !!(env.RESEND_API_KEY && env.MAIL_FROM) }, 409);
      }
      const sent = await sendEmails(env, body.state, Array.isArray(body.emails) ? body.emails : []);
      return json({ ok: true, version: v, sent });
    }

    return fail(404, 'Not found.');
  } catch (e) {
    const status = e.status || (e.constructor === Error && !/D1|SQL|database/i.test(e.message) ? 400 : 500);
    if (status === 500) console.error(e);
    return fail(status, status === 500 ? 'Something went wrong on the server. Try again.' : e.message);
  }
}
