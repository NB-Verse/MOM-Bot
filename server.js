'use strict';
/**
 * MOM Bot — invite-only AI meeting notetaker.
 *
 * - Sends a bot (via Recall.ai) into Microsoft Teams / Zoom / Google Meet / Webex meetings.
 * - Receives verified webhooks, transcribes the recording after the call.
 * - Uses an LLM (Anthropic Claude or OpenAI) to write the Minutes of Meeting (MOM).
 * - Access is invite-only: an admin issues single-use invite links.
 *
 * Zero dependencies. Requires Node.js 20+.  Start with:  node server.js
 */

const http = require('node:http');
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');

const ROOT = __dirname;

/* ------------------------------------------------------------------ */
/* Configuration                                                       */
/* ------------------------------------------------------------------ */

(function loadDotEnv() {
  const file = path.join(ROOT, '.env');
  if (!fs.existsSync(file)) return;
  for (const raw of fs.readFileSync(file, 'utf8').split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line.startsWith('#')) continue;
    const eq = line.indexOf('=');
    if (eq < 1) continue;
    const key = line.slice(0, eq).trim();
    let val = line.slice(eq + 1).trim();
    if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) val = val.slice(1, -1);
    if (process.env[key] === undefined) process.env[key] = val;
  }
})();

const env = (k, d = '') => String(process.env[k] ?? d).trim();
const envList = (k) => env(k).split(',').map((s) => s.trim().toLowerCase()).filter(Boolean);
const envInt = (k, d) => { const n = parseInt(env(k, String(d)), 10); return Number.isFinite(n) ? n : d; };

const CFG = {
  port: envInt('PORT', 3000),
  publicBaseUrl: env('PUBLIC_BASE_URL').replace(/\/+$/, ''),
  dataFile: path.resolve(ROOT, env('DATA_FILE', 'data.json')),
  timezone: env('TIMEZONE', 'Asia/Kolkata'),
  adminEmail: env('ADMIN_EMAIL').toLowerCase(),
  adminPassword: env('ADMIN_PASSWORD'),
  adminName: env('ADMIN_NAME', 'Administrator'),
  allowedEmailDomains: envList('ALLOWED_EMAIL_DOMAINS'),
  allowedMeetingHosts: envList('ALLOWED_MEETING_HOSTS'),
  sessionDays: envInt('SESSION_DAYS', 7),
  inviteDaysDefault: envInt('INVITE_DAYS', 7),

  recallRegion: env('RECALL_REGION', 'us-west-2'),
  recallApiKey: env('RECALL_API_KEY'),
  recallVerificationSecret: env('RECALL_VERIFICATION_SECRET'),
  recallSvixSecret: env('RECALL_SVIX_WEBHOOK_SECRET'),
  recallApiBase: env('RECALL_API_BASE'), // testing override only

  // {name} is replaced with the team member's name. Each person can also set their own bot name in the app.
  botName: env('BOT_NAME', '{name} (Notetaker)'),
  botJoinMessage: env('BOT_JOIN_MESSAGE',
    "Hi everyone, this is {name}'s AI notetaker. I'm recording and transcribing this meeting only to prepare the minutes (MOM). If you'd prefer not to be recorded, please let {name} know."),
  calendarLookaheadHours: envInt('CALENDAR_LOOKAHEAD_HOURS', 24),
  icsPollMinutes: Math.max(2, envInt('ICS_POLL_MINUTES', 5)),
  googleClientId: env('GOOGLE_CLIENT_ID'),
  googleClientSecret: env('GOOGLE_CLIENT_SECRET'),
  msClientId: env('MICROSOFT_CLIENT_ID'),
  msClientSecret: env('MICROSOFT_CLIENT_SECRET'),
  msTenant: env('MICROSOFT_TENANT', 'common'),
  // test-only overrides
  googleAuthUrl: env('GOOGLE_AUTH_URL', 'https://accounts.google.com/o/oauth2/v2/auth'),
  googleTokenUrl: env('GOOGLE_TOKEN_URL', 'https://oauth2.googleapis.com/token'),
  googleUserinfoUrl: env('GOOGLE_USERINFO_URL', 'https://www.googleapis.com/oauth2/v3/userinfo'),
  msLoginBase: env('MICROSOFT_LOGIN_BASE', 'https://login.microsoftonline.com').replace(/\/+$/, ''),
  msUserinfoUrl: env('MICROSOFT_USERINFO_URL', 'https://graph.microsoft.com/oidc/userinfo'),
  transcriptLanguage: env('TRANSCRIPT_LANGUAGE', 'auto'),
  momLanguage: env('MOM_LANGUAGE', 'English'),

  llmProvider: env('LLM_PROVIDER', 'anthropic').toLowerCase(),
  anthropicKey: env('ANTHROPIC_API_KEY'),
  anthropicModel: env('ANTHROPIC_MODEL', 'claude-sonnet-5-5'),
  anthropicBase: env('ANTHROPIC_BASE_URL', 'https://api.anthropic.com').replace(/\/+$/, ''),
  openaiKey: env('OPENAI_API_KEY'),
  openaiModel: env('OPENAI_MODEL'),
  openaiBase: env('OPENAI_BASE_URL', 'https://api.openai.com').replace(/\/+$/, ''),
  momMaxTokens: envInt('MOM_MAX_TOKENS', 12000),
};

const RECALL_REGIONS = ['us-west-2', 'us-east-1', 'eu-central-1', 'ap-northeast-1'];
const RECALL_BASE = CFG.recallApiBase.replace(/\/+$/, '') || `https://${CFG.recallRegion}.recall.ai`;

/* ------------------------------------------------------------------ */
/* Utilities                                                           */
/* ------------------------------------------------------------------ */

const nowIso = () => new Date().toISOString();
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const newId = (bytes = 12) => crypto.randomBytes(bytes).toString('base64url');
const sha256 = (s) => crypto.createHash('sha256').update(String(s)).digest('hex');

class HttpError extends Error {
  constructor(status, message) { super(message); this.status = status; }
}

function log(level, msg, extra) {
  const line = `[${nowIso()}] ${level.toUpperCase().padEnd(5)} ${msg}`;
  const out = level === 'error' || level === 'warn' ? console.error : console.log;
  if (extra === undefined) out(line);
  else out(line, typeof extra === 'string' ? extra : JSON.stringify(extra));
}

function str(v, max = 2000) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/\u0000/g, '').trim().slice(0, max);
}

function fmtDate(iso) {
  if (!iso) return '';
  try {
    return new Intl.DateTimeFormat('en-IN', {
      timeZone: CFG.timezone, day: '2-digit', month: 'short', year: 'numeric',
      hour: '2-digit', minute: '2-digit', hour12: true,
    }).format(new Date(iso));
  } catch { return iso; }
}

function fmtClock(sec) {
  sec = Math.max(0, Math.floor(Number(sec) || 0));
  const h = Math.floor(sec / 3600), m = Math.floor((sec % 3600) / 60), s = sec % 60;
  const pad = (n) => String(n).padStart(2, '0');
  return h ? `${h}:${pad(m)}:${pad(s)}` : `${pad(m)}:${pad(s)}`;
}

/* ------------------------------------------------------------------ */
/* Storage: a single JSON file, written atomically                     */
/* ------------------------------------------------------------------ */

let db;

function emptyDb() {
  return {
    version: 1,
    meta: { sessionSecret: crypto.randomBytes(48).toString('hex'), createdAt: nowIso() },
    users: [], invites: [], meetings: [], processedEvents: [],
  };
}

function loadDb() {
  const tryRead = (file) => JSON.parse(fs.readFileSync(file, 'utf8'));
  if (fs.existsSync(CFG.dataFile)) {
    try { db = tryRead(CFG.dataFile); }
    catch (e) {
      log('error', `Data file is corrupt (${e.message}); trying backup`);
      if (fs.existsSync(CFG.dataFile + '.bak')) db = tryRead(CFG.dataFile + '.bak');
      else throw new Error(`Cannot read ${CFG.dataFile} and no backup exists. Fix or remove it.`);
    }
  } else {
    fs.mkdirSync(path.dirname(CFG.dataFile), { recursive: true });
    db = emptyDb();
  }
  const base = emptyDb();
  for (const k of Object.keys(base)) if (db[k] === undefined) db[k] = base[k];
  if (!db.meta.sessionSecret) db.meta.sessionSecret = base.meta.sessionSecret;
}

let saveChain = Promise.resolve();
let savePending = false;
function save() {
  if (savePending) return saveChain;
  savePending = true;
  saveChain = saveChain.then(async () => {
    savePending = false;
    const tmp = CFG.dataFile + '.tmp';
    await fsp.writeFile(tmp, JSON.stringify(db));
    try { await fsp.copyFile(CFG.dataFile, CFG.dataFile + '.bak'); } catch { /* first write */ }
    await fsp.rename(tmp, CFG.dataFile);
  }).catch((e) => log('error', 'Saving data failed', e.message));
  return saveChain;
}

/* Per-key async lock so events for one meeting are processed in order. */
const locks = new Map();
function withLock(key, fn) {
  const prev = locks.get(key) || Promise.resolve();
  const run = prev.catch(() => {}).then(fn);
  const tail = run.catch(() => {});
  locks.set(key, tail);
  tail.then(() => { if (locks.get(key) === tail) locks.delete(key); });
  return run;
}

/* ------------------------------------------------------------------ */
/* Auth: passwords, sessions, invites                                  */
/* ------------------------------------------------------------------ */

const scryptAsync = (pw, salt) => new Promise((resolve, reject) =>
  crypto.scrypt(pw, salt, 64, { N: 16384, r: 8, p: 1 }, (e, k) => (e ? reject(e) : resolve(k))));

async function hashPassword(pw) {
  const salt = crypto.randomBytes(16).toString('hex');
  return `scrypt$${salt}$${(await scryptAsync(pw, salt)).toString('hex')}`;
}

const DUMMY_HASH = 'scrypt$00000000000000000000000000000000$' + '0'.repeat(128);
async function verifyPassword(pw, stored) {
  const [alg, salt, hex] = String(stored || DUMMY_HASH).split('$');
  if (alg !== 'scrypt' || !salt || !hex) return false;
  const key = await scryptAsync(String(pw), salt);
  const expected = Buffer.from(hex, 'hex');
  return expected.length === key.length && crypto.timingSafeEqual(expected, key);
}

function validatePassword(pw) {
  if (typeof pw !== 'string' || pw.length < 10) throw new HttpError(400, 'Password must be at least 10 characters.');
  if (pw.length > 200) throw new HttpError(400, 'Password is too long.');
}

function validateEmail(email) {
  const e = str(email, 200).toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) throw new HttpError(400, 'Enter a valid email address.');
  if (CFG.allowedEmailDomains.length) {
    const domain = e.split('@')[1];
    if (!CFG.allowedEmailDomains.some((d) => domain === d || domain.endsWith('.' + d))) {
      throw new HttpError(403, `Only ${CFG.allowedEmailDomains.join(', ')} email addresses can be used.`);
    }
  }
  return e;
}

function parseCookies(header) {
  const out = {};
  for (const part of String(header || '').split(';')) {
    const i = part.indexOf('=');
    if (i > 0) out[part.slice(0, i).trim()] = decodeURIComponent(part.slice(i + 1).trim());
  }
  return out;
}

function hmac(data) { return crypto.createHmac('sha256', db.meta.sessionSecret).update(data).digest('base64url'); }

function makeSession(user) {
  const payload = Buffer.from(JSON.stringify({
    u: user.id, v: user.sessionVersion || 0, e: Date.now() + CFG.sessionDays * 864e5,
  })).toString('base64url');
  return `${payload}.${hmac(payload)}`;
}

function userFromRequest(req) {
  const token = parseCookies(req.headers.cookie).sid;
  if (!token) return null;
  const [payload, sig] = token.split('.');
  if (!payload || !sig) return null;
  const a = Buffer.from(sig), b = Buffer.from(hmac(payload));
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
  let data;
  try { data = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8')); } catch { return null; }
  if (!data || !data.e || data.e < Date.now()) return null;
  const user = db.users.find((u) => u.id === data.u);
  if (!user || user.disabled || (user.sessionVersion || 0) !== data.v) return null;
  return user;
}

function isHttps(req) {
  return CFG.publicBaseUrl.startsWith('https://') || req.headers['x-forwarded-proto'] === 'https';
}

function setSessionCookie(req, res, user) {
  const parts = [`sid=${makeSession(user)}`, 'Path=/', 'HttpOnly', 'SameSite=Lax', `Max-Age=${CFG.sessionDays * 86400}`];
  if (isHttps(req)) parts.push('Secure');
  res.setHeader('Set-Cookie', parts.join('; '));
}

function clearSessionCookie(req, res) {
  res.setHeader('Set-Cookie', `sid=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0${isHttps(req) ? '; Secure' : ''}`);
}

const publicUser = (u) => u && ({
  id: u.id, email: u.email, name: u.name, role: u.role, disabled: !!u.disabled, createdAt: u.createdAt, lastLoginAt: u.lastLoginAt || null,
  botName: botNameFor(u), customBotName: u.botName || '',
});

/* Simple in-memory rate limiter */
const buckets = new Map();
function rateLimit(key, max, windowMs) {
  const t = Date.now();
  const b = buckets.get(key);
  if (!b || b.reset < t) { buckets.set(key, { count: 1, reset: t + windowMs }); return; }
  b.count += 1;
  if (b.count > max) {
    const mins = Math.ceil((b.reset - t) / 60000);
    throw new HttpError(429, `Too many attempts. Please try again in ${mins} minute${mins === 1 ? '' : 's'}.`);
  }
}
setInterval(() => { const t = Date.now(); for (const [k, b] of buckets) if (b.reset < t) buckets.delete(k); }, 10 * 60000).unref();

async function bootstrapAdmin() {
  if (db.users.length) return;
  if (!CFG.adminEmail || !CFG.adminPassword) {
    log('warn', 'No users exist yet. Set ADMIN_EMAIL and ADMIN_PASSWORD and restart to create the first admin.');
    return;
  }
  if (CFG.adminPassword.length < 10) {
    log('error', 'ADMIN_PASSWORD must be at least 10 characters. Admin not created.');
    return;
  }
  db.users.push({
    id: newId(), email: CFG.adminEmail, name: CFG.adminName, role: 'admin',
    passwordHash: await hashPassword(CFG.adminPassword), sessionVersion: 0, createdAt: nowIso(),
  });
  await save();
  log('info', `Created first admin account: ${CFG.adminEmail}`);
}

/* ------------------------------------------------------------------ */
/* Meeting links                                                       */
/* ------------------------------------------------------------------ */

function parseMeetingUrl(input) {
  let raw = str(input, 3000);
  if (!raw) throw new HttpError(400, 'Paste the meeting link.');
  if (!/^https?:\/\//i.test(raw)) raw = 'https://' + raw;
  let u;
  try { u = new URL(raw); } catch { throw new HttpError(400, 'That does not look like a valid meeting link.'); }

  // Unwrap Outlook Safe Links and Google redirect links.
  for (let i = 0; i < 2; i++) {
    const h = u.hostname.toLowerCase();
    const inner = h.endsWith('safelinks.protection.outlook.com') ? u.searchParams.get('url')
      : ((h === 'www.google.com' || h === 'google.com') && u.pathname === '/url') ? u.searchParams.get('q') : null;
    if (!inner) break;
    try { u = new URL(inner); } catch { throw new HttpError(400, 'Could not read the meeting link inside the redirect link.'); }
  }

  if (u.protocol === 'http:') u.protocol = 'https:';
  if (u.protocol !== 'https:') throw new HttpError(400, 'Meeting links must start with https://');
  const host = u.hostname.toLowerCase();
  const is = (d) => host === d || host.endsWith('.' + d);
  let platform = null;
  if (is('teams.microsoft.com') || is('teams.live.com') || is('teams.cloud.microsoft')) platform = 'Microsoft Teams';
  else if (is('zoom.us') || is('zoomgov.com')) platform = 'Zoom';
  else if (host === 'meet.google.com') platform = 'Google Meet';
  else if (is('webex.com')) platform = 'Webex';
  if (!platform) throw new HttpError(400, 'Only Microsoft Teams, Zoom, Google Meet and Webex meeting links are supported.');
  if (CFG.allowedMeetingHosts.length && !CFG.allowedMeetingHosts.some((h) => is(h))) {
    throw new HttpError(403, 'Meetings on this platform are not allowed by your administrator.');
  }
  if (platform === 'Google Meet' && !/^\/[a-z]{3}-[a-z]{4}-[a-z]{3}/i.test(u.pathname)) {
    throw new HttpError(400, 'That Google Meet link looks incomplete (expected meet.google.com/abc-defg-hij).');
  }
  return { url: u.toString(), platform };
}

/* ------------------------------------------------------------------ */
/* Recall.ai client                                                    */
/* ------------------------------------------------------------------ */

function recallConfigured() { return !!CFG.recallApiKey; }

const recall = (method, apiPath, body) => recallReq(method, `/api/v1${apiPath}`, body);
const recall2 = (method, apiPath, body) => recallReq(method, `/api/v2${apiPath}`, body);

async function recallReq(method, apiPath, body) {
  if (!recallConfigured()) throw new Error('Recall.ai is not configured (RECALL_API_KEY is missing).');
  const url = `${RECALL_BASE}${apiPath}`;
  const maxAttempts = 6;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    let res;
    try {
      res = await fetch(url, {
        method,
        headers: {
          Authorization: CFG.recallApiKey,
          accept: 'application/json',
          ...(body ? { 'content-type': 'application/json' } : {}),
        },
        body: body ? JSON.stringify(body) : undefined,
        signal: AbortSignal.timeout(30000),
      });
    } catch (e) {
      if (attempt === maxAttempts) throw new Error(`Recall.ai could not be reached: ${e.message}`);
      await sleep(2000 * attempt);
      continue;
    }
    let wait = null;
    if (res.status === 429) wait = parseInt(res.headers.get('retry-after') || '5', 10) || 5;
    else if (res.status === 502 || res.status === 503 || res.status === 504) wait = 10;
    else if (res.status === 507) wait = 30; // ad-hoc bot pool drained
    if (wait !== null && attempt < maxAttempts) {
      log('warn', `Recall.ai ${method} ${apiPath} returned ${res.status}; retrying in ~${wait}s`);
      await sleep(1000 * (wait + Math.ceil(Math.random() * 5)));
      continue;
    }
    const text = await res.text();
    let data = null;
    try { data = text ? JSON.parse(text) : null; } catch { data = text; }
    if (!res.ok) {
      const detail = typeof data === 'string' ? data : JSON.stringify(data);
      const err = new Error(`Recall.ai ${method} ${apiPath} failed (${res.status}): ${String(detail).slice(0, 400)}`);
      err.status = res.status;
      err.data = data;
      throw err;
    }
    return data;
  }
  throw new Error(`Recall.ai ${method} ${apiPath} kept failing after ${maxAttempts} attempts.`);
}

function verifyRecallRequest(headers, rawBody) {
  const id = headers['webhook-id'] || headers['svix-id'];
  const ts = headers['webhook-timestamp'] || headers['svix-timestamp'];
  const sigHeader = headers['webhook-signature'] || headers['svix-signature'];
  if (!id || !ts || !sigHeader) return 'missing signature headers';
  const tsNum = parseInt(ts, 10);
  if (!tsNum || Math.abs(Date.now() / 1000 - tsNum) > 600) return 'timestamp outside the allowed window';
  const secrets = [CFG.recallVerificationSecret, CFG.recallSvixSecret].filter((s) => s && s.startsWith('whsec_'));
  if (!secrets.length) return 'no verification secret configured (RECALL_VERIFICATION_SECRET)';
  const toSign = `${id}.${ts}.${rawBody}`;
  const passed = String(sigHeader).split(' ')
    .map((p) => p.split(','))
    .filter(([v, s]) => v === 'v1' && s)
    .map(([, s]) => Buffer.from(s, 'base64'));
  for (const secret of secrets) {
    const key = Buffer.from(secret.slice('whsec_'.length), 'base64');
    const expected = crypto.createHmac('sha256', key).update(toSign).digest();
    if (passed.some((b) => b.length === expected.length && crypto.timingSafeEqual(b, expected))) return null;
  }
  return 'signature mismatch';
}

/* ------------------------------------------------------------------ */
/* Meeting state machine                                               */
/* ------------------------------------------------------------------ */

const STAGE_LABEL = {
  dispatching: 'Sending the bot…',
  scheduled: 'Scheduled',
  joining: 'Joining the meeting',
  waiting_room: 'Waiting in the lobby — please admit the bot',
  in_call: 'In the meeting (not recording yet)',
  recording: 'Recording and taking notes',
  processing: 'Meeting ended — processing recording',
  transcribing: 'Transcribing the meeting',
  generating: 'Writing the minutes',
  completed: 'MOM ready',
  failed: 'Failed',
  cancelled: 'Cancelled',
};
const EARLY_STAGES = new Set(['dispatching', 'scheduled', 'joining', 'waiting_room', 'in_call', 'recording', 'failed']);
const FINAL_STAGES = new Set(['completed', 'failed', 'cancelled']);

const SUB_CODE_TEXT = {
  timeout_exceeded_waiting_room: 'Nobody admitted the bot from the waiting room / lobby in time.',
  call_ended_by_platform_waiting_room_timeout: 'Microsoft Teams removed the bot after 30 minutes in the lobby.',
  timeout_exceeded_noone_joined: 'Nobody else joined the meeting, so the bot left.',
  timeout_exceeded_everyone_left: 'Everyone left the meeting, so the bot left too.',
  timeout_exceeded_only_bots_detected_using_participant_names: 'Only bots were left in the meeting, so the bot left.',
  timeout_exceeded_only_bots_detected_using_participant_events: 'Only bots were left in the meeting, so the bot left.',
  timeout_exceeded_silence_detected: 'The meeting was silent for too long, so the bot left.',
  timeout_exceeded_in_call_not_recording: 'The bot was never allowed to record, so it left.',
  timeout_exceeded_recording_permission_denied: 'The host denied recording permission.',
  recording_permission_denied: 'The host denied recording permission.',
  bot_kicked_from_call: 'Someone removed the bot from the meeting.',
  bot_kicked_from_waiting_room: 'Someone declined the bot in the waiting room.',
  bot_received_leave_call: 'The bot was asked to leave the meeting.',
  call_ended_by_host: 'The host ended the meeting.',
  call_ended_by_platform_max_length: 'The meeting platform ended the call (maximum length).',
  meeting_not_found: 'The meeting could not be found — check the link.',
  meeting_not_started: 'The meeting has not started yet.',
  meeting_link_expired: 'The meeting link has expired.',
  meeting_link_invalid: 'The meeting link is not valid.',
  meeting_password_incorrect: 'The meeting passcode is wrong or missing — use the full invite link.',
  meeting_full: 'The meeting is full.',
  meeting_requires_registration: 'The meeting requires registration, which the bot cannot complete.',
  google_meet_sign_in_required: 'This Google Meet only admits signed-in users from the organisation.',
  google_meet_login_required: 'This Google Meet only admits signed-in users from the organisation.',
  microsoft_teams_call_dropped: 'Microsoft Teams dropped the bot from the call.',
  zoom_web_disallowed: 'This Zoom meeting does not allow joining from the web.',
};

function subCodeText(code) {
  if (!code) return '';
  return SUB_CODE_TEXT[code] || `Reason reported by the meeting bot: ${code.replace(/_/g, ' ')}.`;
}

function touch(m) { m.updatedAt = nowIso(); }

function addHistory(m, code, subCode, at, source = 'event') {
  m.statusHistory = m.statusHistory || [];
  const when = at || nowIso();
  if (m.statusHistory.some((h) => h.code === code && h.at === when)) return false;
  m.statusHistory.push({ code, subCode: subCode || null, at: when, source });
  m.statusHistory.sort((a, b) => String(a.at).localeCompare(String(b.at)));
  if (m.statusHistory.length > 200) m.statusHistory = m.statusHistory.slice(-200);
  return true;
}

function setStage(m, stage, { force = false } = {}) {
  if (!force && !EARLY_STAGES.has(m.stage)) return;
  m.stage = stage;
  touch(m);
}

function applyBotStatus(m, code, subCode, at) {
  addHistory(m, code, subCode, at);
  switch (code) {
    case 'joining_call': setStage(m, 'joining'); break;
    case 'in_waiting_room': setStage(m, 'waiting_room'); break;
    case 'in_call_not_recording':
    case 'recording_permission_allowed': setStage(m, 'in_call'); break;
    case 'recording_permission_denied':
      setStage(m, 'in_call');
      addWarning(m, subCodeText(subCode) || 'Recording permission was denied by the host.');
      break;
    case 'in_call_recording': m.everRecorded = true; m.error = null; setStage(m, 'recording'); break;
    case 'call_ended':
      if (subCode && !m.everRecorded) m.endReason = subCodeText(subCode);
      setStage(m, m.everRecorded ? 'processing' : m.stage);
      break;
    case 'done': m.botDone = true; break;
    case 'fatal':
      m.fatal = true;
      m.endReason = subCodeText(subCode) || 'The meeting bot hit an error.';
      if (!m.everRecorded) { m.error = m.endReason; setStage(m, 'failed'); }
      else addWarning(m, m.endReason + ' Notes are prepared from what was recorded.');
      break;
    default: break;
  }
  touch(m);
}

function addWarning(m, text) {
  if (!text) return;
  m.warnings = m.warnings || [];
  if (!m.warnings.includes(text)) m.warnings.push(text);
}

function ensureRec(m, recId) {
  m.recordings = m.recordings || {};
  if (!m.recordings[recId]) m.recordings[recId] = { id: recId, createdAt: nowIso(), recordingStatus: null, transcriptId: null, transcriptStatus: null, segments: null, error: null };
  return m.recordings[recId];
}

function botNameFor(user) {
  const name = (user && user.name) || 'Team';
  return (str(user && user.botName, 100) || CFG.botName.replace(/\{name\}/g, name)).slice(0, 100);
}

function botConfig(m) {
  const owner = db.users.find((u) => u.id === m.ownerId);
  const personName = (owner && owner.name) || m.ownerName || 'the organiser';
  m.botName = botNameFor(owner || { name: m.ownerName });
  return {
    bot_name: m.botName,
    metadata: { meeting_id: m.id, app: 'mom-bot' },
    chat: { on_bot_join: { send_to: 'everyone', message: CFG.botJoinMessage.replace(/\{name\}/g, personName).slice(0, 1000), pin: true } },
    automatic_leave: {
      waiting_room_timeout: 1200,
      noone_joined_timeout: 1200,
      everyone_left_timeout: { timeout: 60, activate_after: 0 },
      bot_detection: {
        using_participant_names: {
          matches: ['notetaker', 'otter.ai', 'fireflies', 'read.ai', 'fathom', 'tl;dv'].filter((x) => !m.botName.toLowerCase().includes(x)),
          timeout: 600,
          activate_after: 1200,
        },
      },
    },
  };
}

async function dispatchBot(m) {
  return withLock(m.id, async () => {
    if (m.botId) return;
    try {
      const body = { meeting_url: m.meetingUrl, ...botConfig(m) };
      if (m.joinAt && new Date(m.joinAt).getTime() - Date.now() > 60000) body.join_at = new Date(m.joinAt).toISOString();
      const bot = await recall('POST', '/bot/', body);
      if (!bot || !bot.id) throw new Error('Recall.ai did not return a bot id.');
      m.botId = bot.id;
      m.error = null;
      addHistory(m, body.join_at ? 'bot_scheduled' : 'bot_sent', null, nowIso(), 'app');
      setStage(m, body.join_at ? 'scheduled' : 'joining', { force: true });
      log('info', `Bot ${bot.id} created for meeting ${m.id}${body.join_at ? ' (scheduled)' : ''}`);
    } catch (e) {
      log('error', `Could not create bot for meeting ${m.id}`, e.message);
      m.error = friendlyRecallError(e);
      setStage(m, 'failed', { force: true });
    }
    await save();
  });
}

function friendlyRecallError(e) {
  if (e.status === 401 || e.status === 403) return 'Recall.ai rejected the API key. Check RECALL_API_KEY and RECALL_REGION (the key only works in its own region).';
  if (e.status === 400) return `The meeting bot service rejected the request: ${e.message.replace(/^Recall\.ai [A-Z]+ \S+ failed \(400\): /, '')}`;
  return e.message;
}

async function requestTranscript(m, recId) {
  const rec = ensureRec(m, recId);
  if (rec.transcriptStatus === 'requested' || rec.transcriptStatus === 'done') return;
  try {
    const t = await recall('POST', `/recording/${recId}/create_transcript/`, {
      provider: { recallai_async: { language_code: CFG.transcriptLanguage || 'auto' } },
      diarization: { use_separate_streams_when_available: true },
    });
    rec.transcriptId = (t && t.id) || rec.transcriptId;
    rec.transcriptStatus = 'requested';
    rec.error = null;
    setStage(m, 'transcribing', { force: !['generating', 'completed'].includes(m.stage) });
    log('info', `Transcription requested for recording ${recId} (meeting ${m.id})`);
  } catch (e) {
    // A transcript may already exist (e.g. a retried webhook) — look it up instead of failing.
    if (e.status === 400 || e.status === 409) {
      const found = await findExistingTranscript(m, recId).catch(() => null);
      if (found) return;
    }
    rec.transcriptStatus = 'failed';
    rec.error = e.message;
    log('error', `Transcription request failed for recording ${recId}`, e.message);
  }
  touch(m);
}

async function findExistingTranscript(m, recId) {
  const r = await recall('GET', `/recording/${recId}/`);
  const t = r && r.media_shortcuts && r.media_shortcuts.transcript;
  if (!t || !t.id) return false;
  const rec = ensureRec(m, recId);
  rec.transcriptId = t.id;
  const code = t.status && t.status.code;
  if (code === 'done') await fetchTranscript(m, recId, t.id);
  else if (code === 'failed') { rec.transcriptStatus = 'failed'; rec.error = (t.status && t.status.sub_code) || 'Transcription failed.'; }
  else rec.transcriptStatus = 'requested';
  return true;
}

async function fetchTranscript(m, recId, transcriptId) {
  const t = await recall('GET', `/transcript/${transcriptId}/`);
  const resolvedRecId = recId || (t && t.recording && t.recording.id) || `transcript-${transcriptId}`;
  const rec = ensureRec(m, resolvedRecId);
  rec.transcriptId = transcriptId;
  const url = t && t.data && t.data.download_url;
  if (!url) throw new Error('Transcript is marked done but has no download link yet.');
  const res = await fetch(url, { signal: AbortSignal.timeout(60000) });
  if (!res.ok) throw new Error(`Downloading the transcript failed (${res.status}).`);
  const json = await res.json();
  rec.segments = normalizeTranscript(json);
  rec.transcriptStatus = 'done';
  rec.error = null;
  touch(m);
  log('info', `Transcript downloaded for meeting ${m.id}: ${rec.segments.length} segments`);
}

function normalizeTranscript(json) {
  const arr = Array.isArray(json) ? json : (json && Array.isArray(json.segments) ? json.segments : []);
  const out = [];
  for (const seg of arr) {
    const words = Array.isArray(seg && seg.words) ? seg.words : [];
    const text = words.map((w) => str(w && w.text, 500)).filter(Boolean).join(' ')
      .replace(/\s+([,.!?;:%])/g, '$1').replace(/\s+/g, ' ').trim();
    if (!text) continue;
    const p = seg.participant || {};
    const speaker = str(p.name, 120) || (p.id !== undefined && p.id !== null ? `Speaker ${p.id}` : 'Unknown speaker');
    const start = Number(words[0] && words[0].start_timestamp && words[0].start_timestamp.relative) || 0;
    const last = words[words.length - 1] || {};
    const end = Number((last.end_timestamp && last.end_timestamp.relative) || (last.start_timestamp && last.start_timestamp.relative)) || start;
    const prev = out[out.length - 1];
    if (prev && prev.speaker === speaker && start - prev.end < 20 && prev.text.length < 4000) {
      prev.text += ' ' + text;
      prev.end = Math.max(prev.end, end);
    } else {
      out.push({ speaker, start, end, text });
    }
  }
  return out;
}

function combinedSegments(m) {
  if (Array.isArray(m.manualSegments)) return m.manualSegments;
  const recs = Object.values(m.recordings || {}).sort((a, b) => String(a.createdAt).localeCompare(String(b.createdAt)));
  const all = [];
  let offset = 0;
  for (const r of recs) {
    if (!Array.isArray(r.segments) || !r.segments.length) continue;
    for (const s of r.segments) all.push({ ...s, start: s.start + offset, end: s.end + offset });
    offset = all[all.length - 1].end + 1;
  }
  return all;
}

/** Called whenever something changes after the call: decides whether the MOM can be written. */
async function maybeFinalize(m) {
  if (FINAL_STAGES.has(m.stage) && m.stage !== 'failed') return;
  if (m.stage === 'generating') return;
  if (!m.botDone) return;
  const recs = Object.values(m.recordings || {});
  if (!recs.length) {
    if (m.everRecorded) { setStage(m, 'processing', { force: true }); return; } // recording.done still on its way
    m.error = m.endReason || 'The bot did not record anything (it may not have been admitted to the meeting).';
    setStage(m, 'failed', { force: true });
    return;
  }
  const pending = recs.filter((r) => r.transcriptStatus !== 'done' && r.transcriptStatus !== 'failed');
  if (pending.length) return;
  const segs = combinedSegments(m);
  if (!segs.length) {
    const failed = recs.filter((r) => r.transcriptStatus === 'failed');
    m.error = failed.length === recs.length
      ? `Transcription failed: ${failed[0].error || 'unknown error'}`
      : 'No speech was captured in the recording.';
    setStage(m, 'failed', { force: true });
    return;
  }
  // Mark now (prevents duplicates) and run after the current locked step finishes,
  // so webhooks and HTTP requests are never blocked by the AI call.
  setStage(m, 'generating', { force: true });
  setImmediate(() => withLock(m.id, () => generateMom(m)).catch((e) => log('error', 'generateMom crashed', e.message)));
}

async function handleRecallEvent(evt) {
  const type = str(evt && evt.event, 100);
  const d = (evt && evt.data) || {};
  if (type === 'calendar.sync_events' || type === 'calendar.update') return handleCalendarWebhook(type, d);
  const botId = d.bot && d.bot.id;
  const metaId = d.bot && d.bot.metadata && d.bot.metadata.meeting_id;
  const m = db.meetings.find((x) => (botId && x.botId === botId) || (metaId && x.id === metaId));
  if (!m) { log('info', `Ignoring ${type} for unknown bot ${botId || '?'}`); return; }

  await withLock(m.id, async () => {
    if (!m.botId && botId) m.botId = botId;
    const info = d.data || {};
    try {
      if (type.startsWith('bot.')) {
        const code = info.code || type.slice(4);
        applyBotStatus(m, code, info.sub_code, info.updated_at);
        if (code === 'done') await afterBotDone(m);
      } else if (type === 'recording.done') {
        const rec = ensureRec(m, d.recording.id);
        rec.recordingStatus = 'done';
        m.everRecorded = true;
        await requestTranscript(m, d.recording.id);
        await maybeFinalize(m);
      } else if (type === 'recording.failed') {
        const rec = ensureRec(m, d.recording.id);
        rec.recordingStatus = 'failed';
        rec.transcriptStatus = 'failed';
        rec.error = info.sub_code || 'Recording failed.';
        await maybeFinalize(m);
      } else if (type === 'transcript.done') {
        await fetchTranscript(m, d.recording && d.recording.id, d.transcript.id);
        await maybeFinalize(m);
      } else if (type === 'transcript.failed') {
        const recId = (d.recording && d.recording.id) || `transcript-${d.transcript && d.transcript.id}`;
        const rec = ensureRec(m, recId);
        rec.transcriptStatus = 'failed';
        rec.error = info.sub_code || 'Transcription failed.';
        await maybeFinalize(m);
      }
    } catch (e) {
      log('error', `Processing ${type} for meeting ${m.id} failed`, e.message);
      addWarning(m, `A processing step failed (${type}). Use "Sync with bot" to retry.`);
    }
    touch(m);
    await save();
  });
}

async function afterBotDone(m) {
  m.botDone = true;
  // If recording webhooks are not subscribed (or delayed), discover recordings with a single lookup.
  if (!Object.keys(m.recordings || {}).length) {
    try { await syncFromRecall(m, { finalize: false }); } catch (e) { log('warn', `Lookup after bot.done failed for ${m.id}`, e.message); }
  }
  await maybeFinalize(m);
}

/** Manual / recovery sync: one GET of the bot, then act on what it says. */
async function syncFromRecall(m, { finalize = true } = {}) {
  if (!m.botId) return;
  const bot = await recall('GET', `/bot/${m.botId}/`);
  for (const sc of (bot && bot.status_changes) || []) {
    if (!sc || !sc.code) continue;
    const at = sc.created_at || sc.updated_at;
    if (!(m.statusHistory || []).some((h) => h.code === sc.code && h.at === at)) applyBotStatus(m, sc.code, sc.sub_code, at);
  }
  for (const r of (bot && bot.recordings) || []) {
    if (!r || !r.id) continue;
    const rec = ensureRec(m, r.id);
    if (r.created_at) rec.createdAt = r.created_at;
    rec.recordingStatus = (r.status && r.status.code) || rec.recordingStatus;
    const t = r.media_shortcuts && r.media_shortcuts.transcript;
    if (t && t.id) {
      rec.transcriptId = t.id;
      const code = t.status && t.status.code;
      if (code === 'done' && rec.transcriptStatus !== 'done') await fetchTranscript(m, r.id, t.id);
      else if (code === 'failed') { rec.transcriptStatus = 'failed'; rec.error = (t.status && t.status.sub_code) || 'Transcription failed.'; }
      else if (code && rec.transcriptStatus !== 'done') rec.transcriptStatus = 'requested';
    } else if (rec.recordingStatus === 'done' && rec.transcriptStatus !== 'requested' && rec.transcriptStatus !== 'done') {
      rec.transcriptStatus = null;
      await requestTranscript(m, r.id);
    }
    if (rec.recordingStatus === 'done' || rec.recordingStatus === 'processing') m.everRecorded = true;
  }
  const last = ((bot && bot.status_changes) || []).slice(-1)[0];
  if (last && (last.code === 'done' || last.code === 'fatal')) m.botDone = true;
  touch(m);
  if (finalize) await maybeFinalize(m);
}

/* ------------------------------------------------------------------ */
/* AI: Minutes of Meeting                                              */
/* ------------------------------------------------------------------ */

function llmConfigured() {
  return CFG.llmProvider === 'openai' ? !!(CFG.openaiKey && CFG.openaiModel) : !!CFG.anthropicKey;
}

async function llm(system, user, { json = true, maxTokens = CFG.momMaxTokens } = {}) {
  if (!llmConfigured()) {
    throw new Error(CFG.llmProvider === 'openai'
      ? 'AI is not configured: set OPENAI_API_KEY and OPENAI_MODEL.'
      : 'AI is not configured: set ANTHROPIC_API_KEY.');
  }
  let lastErr;
  for (let attempt = 1; attempt <= 4; attempt++) {
    try {
      let res;
      if (CFG.llmProvider === 'openai') {
        const body = { model: CFG.openaiModel, messages: [{ role: 'system', content: system }, { role: 'user', content: user }] };
        if (json) body.response_format = { type: 'json_object' };
        res = await fetch(`${CFG.openaiBase}/v1/chat/completions`, {
          method: 'POST',
          headers: { authorization: `Bearer ${CFG.openaiKey}`, 'content-type': 'application/json' },
          body: JSON.stringify(body),
          signal: AbortSignal.timeout(300000),
        });
      } else {
        res = await fetch(`${CFG.anthropicBase}/v1/messages`, {
          method: 'POST',
          headers: { 'x-api-key': CFG.anthropicKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
          body: JSON.stringify({ model: CFG.anthropicModel, max_tokens: maxTokens, system, messages: [{ role: 'user', content: user }] }),
          signal: AbortSignal.timeout(300000),
        });
      }
      const text = await res.text();
      if (res.status === 429 || res.status === 529 || res.status >= 500) {
        lastErr = new Error(`AI provider is busy (HTTP ${res.status}).`);
        const ra = parseInt(res.headers.get('retry-after') || '0', 10) || 0;
        if (attempt < 4) await sleep(Math.max(ra * 1000, 2 ** attempt * 4000));
        continue;
      }
      if (!res.ok) {
        const err = new Error(`AI provider error (HTTP ${res.status}): ${text.slice(0, 400)}`);
        err.fatal = true;
        throw err;
      }
      const data = JSON.parse(text);
      const out = CFG.llmProvider === 'openai'
        ? data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content
        : (data.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
      if (!out) throw new Error('The AI returned an empty response.');
      const truncated = data.stop_reason === 'max_tokens' || (data.choices && data.choices[0] && data.choices[0].finish_reason === 'length');
      if (truncated) log('warn', 'AI response hit the token limit and may be truncated');
      return out;
    } catch (e) {
      if (e.fatal) throw e;
      lastErr = e;
      if (attempt < 4) await sleep(2 ** attempt * 3000);
    }
  }
  throw lastErr || new Error('AI request failed.');
}

function parseJsonLoose(text) {
  if (!text) return null;
  let t = String(text).trim().replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
  try { return JSON.parse(t); } catch { /* fall through */ }
  const a = t.indexOf('{'), b = t.lastIndexOf('}');
  if (a >= 0 && b > a) { try { return JSON.parse(t.slice(a, b + 1)); } catch { /* ignore */ } }
  return null;
}

function transcriptToText(segs) {
  return segs.map((s) => `[${fmtClock(s.start)}] ${s.speaker}: ${s.text}`).join('\n');
}

function splitByLines(text, max) {
  const chunks = [];
  let cur = '';
  for (const line of text.split('\n')) {
    if (cur.length + line.length + 1 > max && cur) { chunks.push(cur); cur = ''; }
    if (line.length > max) { for (let i = 0; i < line.length; i += max) chunks.push(line.slice(i, i + max)); continue; }
    cur += (cur ? '\n' : '') + line;
  }
  if (cur) chunks.push(cur);
  return chunks;
}

const MOM_SYSTEM = () => `You are an experienced executive assistant who writes accurate, concise, professional Minutes of Meeting (MOM) for corporate team meetings.
Rules:
- Work strictly from the transcript or notes provided. Never invent facts, names, numbers, owners or dates.
- If something is unclear or not stated, use null (or "Not specified" inside text).
- The transcript is machine-generated: expect recognition errors, filler words, and mixed languages (for example English with Hindi or Marathi). Interpret sensibly.
- Write the minutes in ${CFG.momLanguage}.
- An action item needs a clear task. The owner is the person who agreed to it or was explicitly assigned it. Keep due dates exactly as said (e.g. "by Friday").
- Decisions are things the group agreed on, not mere suggestions.
- Be concise and businesslike; group related discussion into topics.
- Respond with ONE JSON object only. No markdown fences, no commentary.`;

const NOTES_SYSTEM = `You condense one part of a long meeting transcript into detailed, factual notes for a minute-taker.
Keep: who said what on each topic, every decision, every action item with owner and deadline, numbers, dates, risks and open questions.
Do not invent anything. Plain-text bullet points only.`;

const MOM_SCHEMA = `{
  "title": "short meeting title",
  "executive_summary": "3-6 sentence summary of purpose and outcomes",
  "attendees": ["names of people who spoke or were mentioned as present"],
  "agenda_items": [
    { "topic": "topic name", "discussion": ["key point", "..."], "outcome": "conclusion or null" }
  ],
  "decisions": [ { "decision": "what was decided", "owner": "name or null" } ],
  "action_items": [ { "task": "what needs to be done", "owner": "name or null", "due": "deadline as stated or null", "priority": "High | Medium | Low" } ],
  "risks_and_issues": ["..."],
  "open_questions": ["..."],
  "next_meeting": "date/time or plan for next meeting, or null"
}`;

function asArr(v) { return Array.isArray(v) ? v : (v ? [v] : []); }
function asStrOrNull(v) { const s = str(v, 2000); return s && !/^(null|none|n\/a|not specified)$/i.test(s) ? s : null; }

function sanitizeMom(raw) {
  const o = raw && typeof raw === 'object' ? raw : {};
  return {
    title: str(o.title, 200) || null,
    executive_summary: str(o.executive_summary, 5000),
    attendees: asArr(o.attendees).map((x) => str(typeof x === 'object' ? x.name : x, 120)).filter(Boolean),
    agenda_items: asArr(o.agenda_items).map((a) => ({
      topic: str(a && a.topic, 300) || 'Discussion',
      discussion: asArr(a && a.discussion).map((x) => str(x, 2000)).filter(Boolean),
      outcome: asStrOrNull(a && a.outcome),
    })),
    decisions: asArr(o.decisions).map((d) => (typeof d === 'string' ? { decision: str(d, 2000), owner: null }
      : { decision: str(d && d.decision, 2000), owner: asStrOrNull(d && d.owner) })).filter((d) => d.decision),
    action_items: asArr(o.action_items).map((a) => (typeof a === 'string' ? { task: str(a, 2000), owner: null, due: null, priority: 'Medium' } : {
      task: str(a && a.task, 2000), owner: asStrOrNull(a && a.owner), due: asStrOrNull(a && a.due),
      priority: ['High', 'Medium', 'Low'].find((p) => p.toLowerCase() === str(a && a.priority, 20).toLowerCase()) || 'Medium',
    })).filter((a) => a.task),
    risks_and_issues: asArr(o.risks_and_issues).map((x) => str(x, 2000)).filter(Boolean),
    open_questions: asArr(o.open_questions).map((x) => str(x, 2000)).filter(Boolean),
    next_meeting: asStrOrNull(o.next_meeting),
  };
}

function meetingDurationMin(m, segs) {
  const start = (m.statusHistory || []).find((h) => h.code === 'in_call_recording');
  const end = (m.statusHistory || []).slice().reverse().find((h) => h.code === 'call_ended');
  if (start && end) {
    const mins = Math.round((new Date(end.at) - new Date(start.at)) / 60000);
    if (mins > 0 && mins < 24 * 60) return mins;
  }
  if (segs.length) return Math.max(1, Math.round((segs[segs.length - 1].end - segs[0].start) / 60));
  return null;
}

function meetingStartIso(m) {
  const rec = (m.statusHistory || []).find((h) => h.code === 'in_call_recording' || h.code === 'in_call_not_recording');
  return (rec && rec.at) || m.joinAt || m.createdAt;
}

async function generateMom(m) {
  const segs = combinedSegments(m);
  if (!segs.length) throw new Error('There is no transcript to write minutes from.');
  setStage(m, 'generating', { force: true });
  m.momError = null;
  touch(m);
  await save();

  try {
    const transcript = transcriptToText(segs);
    const speakers = [...new Set(segs.map((s) => s.speaker))];
    let source = transcript;
    let chunked = false;
    const CHUNK = 90000;
    if (transcript.length > CHUNK) {
      chunked = true;
      const parts = splitByLines(transcript, CHUNK);
      const notes = [];
      for (let i = 0; i < parts.length; i++) {
        notes.push(await llm(NOTES_SYSTEM, `Meeting: ${m.title}\nThis is part ${i + 1} of ${parts.length} of the transcript.\n\n${parts[i]}`, { json: false, maxTokens: 4000 }));
      }
      source = notes.map((n, i) => `--- Notes from part ${i + 1} of ${parts.length} ---\n${n}`).join('\n\n');
    }

    const prompt = [
      'Prepare the Minutes of Meeting for this meeting.',
      '',
      `Meeting title given by organiser: ${m.title}`,
      `Platform: ${m.platform}`,
      `Date/time: ${fmtDate(meetingStartIso(m))} (${CFG.timezone})`,
      m.agenda ? `Planned agenda (from organiser):\n${m.agenda}` : 'Planned agenda: not provided',
      `Speakers detected in the transcript: ${speakers.join(', ') || 'unknown'}`,
      `Note: the silent notetaker bot joined as "${m.botName || 'Notetaker'}". It never speaks, so ignore it as an attendee; anyone with that name who actually spoke is a real attendee.`,
      '',
      `Return JSON exactly in this shape:\n${MOM_SCHEMA}`,
      '',
      chunked ? 'MEETING NOTES (condensed from a long transcript, in order):' : 'TRANSCRIPT ([time] Speaker: words):',
      source,
    ].join('\n');

    let raw = await llm(MOM_SYSTEM(), prompt);
    let parsed = parseJsonLoose(raw);
    if (!parsed) {
      raw = await llm(MOM_SYSTEM(), `${prompt}\n\nIMPORTANT: reply with ONLY the JSON object, and keep it concise so it is complete.`);
      parsed = parseJsonLoose(raw);
    }
    if (!parsed) throw new Error('The AI response could not be read as structured minutes. Try "Regenerate MOM".');

    const mom = sanitizeMom(parsed);
    m.mom = mom;
    m.momMarkdown = momToMarkdown(m, mom, segs);
    m.momEdited = false;
    m.momGeneratedAt = nowIso();
    m.error = null;
    setStage(m, 'completed', { force: true });
    m.completedAt = nowIso();
    log('info', `MOM generated for meeting ${m.id}`);
  } catch (e) {
    log('error', `MOM generation failed for ${m.id}`, e.message);
    m.momError = e.message;
    m.error = `Writing the minutes failed: ${e.message}`;
    setStage(m, 'failed', { force: true });
  }
  touch(m);
  await save();
}

function mdCell(s) { return String(s || '—').replace(/\|/g, '\\|').replace(/\n+/g, ' '); }

function momToMarkdown(m, mom, segs) {
  const L = [];
  const dur = meetingDurationMin(m, segs);
  L.push(`# Minutes of Meeting — ${m.title || mom.title || 'Team meeting'}`);
  L.push('');
  L.push(`**Date:** ${fmtDate(meetingStartIso(m))}  `);
  L.push(`**Platform:** ${m.platform}${dur ? `  \n**Duration:** about ${dur} minute${dur === 1 ? '' : 's'}` : ''}  `);
  L.push(`**Organiser:** ${m.ownerName || '—'}  `);
  L.push(`**Attendees:** ${mom.attendees.length ? mom.attendees.join(', ') : 'Not captured'}`);
  L.push('');
  if (m.agenda) { L.push('## Agenda'); L.push(''); m.agenda.split(/\r?\n/).filter((x) => x.trim()).forEach((x) => L.push(`- ${x.replace(/^[-*\d.)\s]+/, '')}`)); L.push(''); }
  L.push('## Summary');
  L.push('');
  L.push(mom.executive_summary || 'No summary available.');
  L.push('');
  if (mom.agenda_items.length) {
    L.push('## Discussion');
    L.push('');
    mom.agenda_items.forEach((a, i) => {
      L.push(`### ${i + 1}. ${a.topic}`);
      a.discussion.forEach((d) => L.push(`- ${d}`));
      if (a.outcome) L.push(`- **Outcome:** ${a.outcome}`);
      L.push('');
    });
  }
  L.push('## Decisions');
  L.push('');
  if (mom.decisions.length) mom.decisions.forEach((d, i) => L.push(`${i + 1}. ${d.decision}${d.owner ? ` _(Owner: ${d.owner})_` : ''}`));
  else L.push('No formal decisions were recorded.');
  L.push('');
  L.push('## Action items');
  L.push('');
  if (mom.action_items.length) {
    L.push('| # | Action | Owner | Due | Priority |');
    L.push('|---|---|---|---|---|');
    mom.action_items.forEach((a, i) => L.push(`| ${i + 1} | ${mdCell(a.task)} | ${mdCell(a.owner || 'Unassigned')} | ${mdCell(a.due || 'Not set')} | ${mdCell(a.priority)} |`));
  } else L.push('No action items were recorded.');
  L.push('');
  if (mom.risks_and_issues.length) { L.push('## Risks & issues'); L.push(''); mom.risks_and_issues.forEach((x) => L.push(`- ${x}`)); L.push(''); }
  if (mom.open_questions.length) { L.push('## Open questions'); L.push(''); mom.open_questions.forEach((x) => L.push(`- ${x}`)); L.push(''); }
  L.push('## Next meeting');
  L.push('');
  L.push(mom.next_meeting || 'Not scheduled during the meeting.');
  L.push('');
  L.push('---');
  L.push(`_Prepared by ${m.botName || 'the AI notetaker'} from an automatic transcript. Please review before circulating._`);
  return L.join('\n');
}

/* Manual transcript import (for meetings the bot could not attend) */
function parseManualTranscript(text) {
  const segs = [];
  const raw = String(text || '').replace(/\r/g, '');
  const lines = raw.split('\n');
  let t = 0;
  let pendingTime = null;
  const tsRe = /^(\d{1,2}:)?\d{1,2}:\d{2}(\.\d+)?\s*-->\s*(\d{1,2}:)?\d{1,2}:\d{2}/;
  const toSec = (s) => s.split(':').map(Number).reduce((acc, n) => acc * 60 + n, 0);
  for (let line of lines) {
    line = line.trim();
    if (!line || line === 'WEBVTT' || /^NOTE\b/.test(line) || /^\d+$/.test(line) || /^[0-9a-f-]{20,}(\/\d+-\d+)?$/i.test(line)) continue;
    if (tsRe.test(line)) { pendingTime = toSec(line.split('-->')[0].trim()); continue; }
    let speaker = null, body = line;
    const v = line.match(/^<v\s+([^>]+)>(.*?)(<\/v>)?$/i);
    if (v) { speaker = v[1].trim(); body = v[2]; }
    else {
      const c = line.match(/^(?:\[?(\d{1,2}:\d{2}(?::\d{2})?)\]?\s+)?([A-Za-zÀ-ɏऀ-ॿ][\w .'À-ɏऀ-ॿ-]{0,60}?)\s*:\s+(.+)$/);
      if (c) { if (c[1]) pendingTime = toSec(c[1]); speaker = c[2].trim(); body = c[3]; }
    }
    body = body.replace(/<[^>]+>/g, '').trim();
    if (!body) continue;
    const start = pendingTime !== null ? pendingTime : t;
    pendingTime = null;
    const prev = segs[segs.length - 1];
    if (prev && (!speaker || speaker === prev.speaker)) { prev.text += ' ' + body; prev.end = Math.max(prev.end, start); }
    else segs.push({ speaker: speaker || 'Unknown speaker', start, end: start, text: body });
    t = start + 5;
  }
  return segs;
}

/* ------------------------------------------------------------------ */
/* Calendar: iCal (ICS) feeds                                          */
/* ------------------------------------------------------------------ */

const WINDOWS_TZ = {
  'india standard time': 'Asia/Kolkata', 'sri lanka standard time': 'Asia/Colombo', 'nepal standard time': 'Asia/Kathmandu',
  'bangladesh standard time': 'Asia/Dhaka', 'pakistan standard time': 'Asia/Karachi', 'arabian standard time': 'Asia/Dubai',
  'arab standard time': 'Asia/Riyadh', 'singapore standard time': 'Asia/Singapore', 'malay peninsula standard time': 'Asia/Kuala_Lumpur',
  'china standard time': 'Asia/Shanghai', 'taipei standard time': 'Asia/Taipei', 'tokyo standard time': 'Asia/Tokyo',
  'korea standard time': 'Asia/Seoul', 'se asia standard time': 'Asia/Bangkok', 'w. australia standard time': 'Australia/Perth',
  'aus eastern standard time': 'Australia/Sydney', 'new zealand standard time': 'Pacific/Auckland', 'utc': 'UTC',
  'coordinated universal time': 'UTC', 'gmt standard time': 'Europe/London', 'greenwich standard time': 'Atlantic/Reykjavik',
  'w. europe standard time': 'Europe/Berlin', 'romance standard time': 'Europe/Paris', 'central europe standard time': 'Europe/Budapest',
  'central european standard time': 'Europe/Warsaw', 'gtb standard time': 'Europe/Bucharest', 'fle standard time': 'Europe/Kiev',
  'turkey standard time': 'Europe/Istanbul', 'russian standard time': 'Europe/Moscow', 'israel standard time': 'Asia/Jerusalem',
  'egypt standard time': 'Africa/Cairo', 'south africa standard time': 'Africa/Johannesburg', 'e. africa standard time': 'Africa/Nairobi',
  'eastern standard time': 'America/New_York', 'central standard time': 'America/Chicago', 'mountain standard time': 'America/Denver',
  'us mountain standard time': 'America/Phoenix', 'pacific standard time': 'America/Los_Angeles', 'alaskan standard time': 'America/Anchorage',
  'hawaiian standard time': 'Pacific/Honolulu', 'atlantic standard time': 'America/Halifax', 'e. south america standard time': 'America/Sao_Paulo',
};

const validTz = (tz) => { try { new Intl.DateTimeFormat('en-US', { timeZone: tz }); return true; } catch { return false; } };

function resolveTz(tzid) {
  if (!tzid) return CFG.timezone;
  const t = String(tzid).replace(/^"|"$/g, '').trim();
  if (validTz(t)) return t;
  if (WINDOWS_TZ[t.toLowerCase()]) return WINDOWS_TZ[t.toLowerCase()];
  const tail = t.match(/[A-Za-z]+\/[A-Za-z_]+(?:\/[A-Za-z_]+)?$/);
  if (tail && validTz(tail[0])) return tail[0];
  if (/india|kolkata|calcutta|mumbai|new delhi/i.test(t)) return 'Asia/Kolkata';
  return CFG.timezone;
}

const tzFmtCache = new Map();
function tzOffsetMs(ts, tz) {
  let f = tzFmtCache.get(tz);
  if (!f) {
    f = new Intl.DateTimeFormat('en-US', { timeZone: tz, hourCycle: 'h23', year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric', second: 'numeric' });
    tzFmtCache.set(tz, f);
  }
  const p = {};
  for (const x of f.formatToParts(new Date(ts))) p[x.type] = Number(x.value);
  return Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second) - Math.floor(ts / 1000) * 1000;
}

/** Wall-clock time in a time zone -> UTC milliseconds (DST-safe). */
function zonedToUtc(y, mo, d, h, mi, s, tz) {
  const guess = Date.UTC(y, mo, d, h, mi, s);
  if (tz === 'UTC') return guess;
  let utc = guess - tzOffsetMs(guess, tz);
  utc = guess - tzOffsetMs(utc, tz);
  return utc;
}

function wallParts(ts, tz) {
  const o = tz === 'UTC' ? 0 : tzOffsetMs(ts, tz);
  const d = new Date(ts + o);
  return { y: d.getUTCFullYear(), mo: d.getUTCMonth(), d: d.getUTCDate(), h: d.getUTCHours(), mi: d.getUTCMinutes(), s: d.getUTCSeconds() };
}

function icsUnescape(v) { return String(v || '').replace(/\\n/gi, '\n').replace(/\\([,;\\])/g, '$1'); }

function parseIcsLine(line) {
  let i = 0, inQ = false;
  for (; i < line.length; i++) {
    const c = line[i];
    if (c === '"') inQ = !inQ;
    else if (c === ':' && !inQ) break;
  }
  const head = line.slice(0, i), value = line.slice(i + 1);
  const parts = head.match(/(?:[^;"]+|"[^"]*")+/g) || [head];
  const params = {};
  for (const p of parts.slice(1)) {
    const eq = p.indexOf('=');
    if (eq > 0) params[p.slice(0, eq).toUpperCase()] = p.slice(eq + 1).replace(/^"|"$/g, '');
  }
  return { name: parts[0].toUpperCase(), params, value };
}

/** Parse an ICS date value -> { ts, tz, allDay } */
function parseIcsDate(value, params = {}) {
  const v = String(value || '').trim();
  const m = v.match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?$/);
  if (!m) return null;
  const [, y, mo, d, h, mi, s, z] = m;
  if (h === undefined || params.VALUE === 'DATE') {
    const tz = CFG.timezone;
    return { ts: zonedToUtc(+y, +mo - 1, +d, 0, 0, 0, tz), tz, allDay: true };
  }
  const tz = z ? 'UTC' : resolveTz(params.TZID);
  return { ts: zonedToUtc(+y, +mo - 1, +d, +h, +mi, +(s || 0), tz), tz, allDay: false };
}

function parseDuration(v) {
  const m = String(v || '').match(/^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/);
  if (!m) return null;
  const ms = ((+m[2] || 0) * 7 * 86400 + (+m[3] || 0) * 86400 + (+m[4] || 0) * 3600 + (+m[5] || 0) * 60 + (+m[6] || 0)) * 1000;
  return m[1] === '-' ? -ms : ms;
}

function parseIcs(text) {
  const lines = String(text || '').replace(/\r\n?/g, '\n').replace(/\n[ \t]/g, '').split('\n');
  const events = [];
  const stack = [];
  let cur = null;
  for (const raw of lines) {
    if (!raw.trim()) continue;
    const { name, params, value } = parseIcsLine(raw);
    if (name === 'BEGIN') {
      stack.push(value.toUpperCase());
      if (value.toUpperCase() === 'VEVENT' && stack.length <= 2) cur = { props: {} };
      continue;
    }
    if (name === 'END') {
      const top = stack.pop();
      if (top === 'VEVENT' && cur) { events.push(cur); cur = null; }
      continue;
    }
    if (cur && stack[stack.length - 1] === 'VEVENT') (cur.props[name] = cur.props[name] || []).push({ params, value });
  }
  return events;
}

const DAY_CODES = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];

/** Expand an RRULE into occurrence start times (UTC ms) between from..to. */
function expandRrule(rruleStr, dt, from, to, exdates) {
  const r = {};
  for (const part of String(rruleStr).split(';')) { const [k, v] = part.split('='); if (k && v) r[k.toUpperCase()] = v.toUpperCase(); }
  const freq = r.FREQ;
  const interval = Math.max(1, parseInt(r.INTERVAL, 10) || 1);
  const count = parseInt(r.COUNT, 10) || 0;
  let until = Infinity;
  if (r.UNTIL) { const u = parseIcsDate(r.UNTIL, { TZID: dt.tz }); if (u) until = u.allDay ? u.ts + 86399000 : u.ts; }
  const base = wallParts(dt.ts, dt.tz);
  const mk = (y, mo, d) => zonedToUtc(y, mo, d, base.h, base.mi, base.s, dt.tz);
  const dim = (y, mo) => new Date(Date.UTC(y, mo + 1, 0)).getUTCDate();
  const byday = r.BYDAY ? r.BYDAY.split(',').map((x) => { const mm = x.match(/^([+-]?\d+)?(SU|MO|TU|WE|TH|FR|SA)$/); return mm ? { n: mm[1] ? +mm[1] : 0, wd: DAY_CODES.indexOf(mm[2]) } : null; }).filter(Boolean) : null;
  const bymonthday = r.BYMONTHDAY ? r.BYMONTHDAY.split(',').map(Number).filter(Boolean) : null;
  const out = [];
  let n = 0, guard = 0;
  const emit = (ts) => {
    if (ts < dt.ts) return true;
    if (ts > until || ts > to) return false;
    n += 1;
    if (count && n > count) return false;
    if (ts >= from && !exdates.has(ts)) out.push(ts);
    return true;
  };
  const baseDate = Date.UTC(base.y, base.mo, base.d);
  if (freq === 'DAILY') {
    for (let k = 0; guard++ < 20000; k += interval) {
      const dd = new Date(baseDate + k * 864e5);
      if (!emit(mk(dd.getUTCFullYear(), dd.getUTCMonth(), dd.getUTCDate()))) break;
    }
  } else if (freq === 'WEEKLY') {
    const baseDow = new Date(baseDate).getUTCDay();
    const weekStart = baseDate - ((baseDow + 6) % 7) * 864e5; // Monday
    const days = (byday && byday.length ? byday.map((b) => b.wd) : [baseDow]).map((wd) => (wd + 6) % 7).sort((a, b) => a - b);
    outer: for (let w = 0; guard++ < 20000; w += interval) {
      for (const di of days) {
        const dd = new Date(weekStart + (w * 7 + di) * 864e5);
        if (!emit(mk(dd.getUTCFullYear(), dd.getUTCMonth(), dd.getUTCDate()))) break outer;
      }
    }
  } else if (freq === 'MONTHLY' || freq === 'YEARLY') {
    const step = freq === 'YEARLY' ? 12 * interval : interval;
    outer2: for (let k = 0; guard++ < 5000; k += step) {
      const y = base.y + Math.floor((base.mo + k) / 12), mo = (base.mo + k) % 12, last = dim(y, mo);
      let days = [];
      if (byday && byday.length && freq === 'MONTHLY') {
        for (const b of byday) {
          const all = [];
          for (let d = 1; d <= last; d++) if (new Date(Date.UTC(y, mo, d)).getUTCDay() === b.wd) all.push(d);
          if (b.n > 0 && all[b.n - 1]) days.push(all[b.n - 1]);
          else if (b.n < 0 && all[all.length + b.n]) days.push(all[all.length + b.n]);
          else if (!b.n) days.push(...all);
        }
      } else if (bymonthday) {
        days = bymonthday.map((d) => (d < 0 ? last + d + 1 : d)).filter((d) => d >= 1 && d <= last);
      } else if (base.d <= last) days = [base.d];
      for (const d of [...new Set(days)].sort((a, b) => a - b)) if (!emit(mk(y, mo, d))) break outer2;
    }
  } else if (dt.ts >= from && dt.ts <= to) out.push(dt.ts);
  return out;
}

const MEETING_PATTERNS = [
  [/teams\.(microsoft|live)\.com\/(l\/meetup-join|meet\/)/i, 5], [/zoom(gov)?\.(us|com)\/(j|w|my|s)\//i, 5],
  [/meet\.google\.com\/[a-z]{3}-[a-z]{4}-[a-z]{3}/i, 5], [/webex\.com\/(meet|join|.*j\.php|wbxmjs)/i, 4],
];

function findMeetingLink(texts) {
  let best = null;
  for (const t of texts) {
    for (const raw of String(t || '').match(/https?:\/\/[^\s<>"'\])]+/gi) || []) {
      const cleaned = raw.replace(/[.,;:]+$/, '');
      let parsed;
      try { parsed = parseMeetingUrl(cleaned); } catch { continue; }
      const score = (MEETING_PATTERNS.find(([re]) => re.test(parsed.url)) || [null, 1])[1];
      if (!best || score > best.score) best = { ...parsed, score };
    }
  }
  return best;
}

/** Turn an ICS feed into upcoming meetings (with links) between from..to. */
function icsToMeetings(text, from, to, userEmail) {
  const raw = parseIcs(text);
  const first = (e, k) => (e.props[k] && e.props[k][0]) || null;
  const val = (e, k) => { const p = first(e, k); return p ? icsUnescape(p.value) : ''; };
  const overrides = new Map();
  const masters = [];
  for (const e of raw) {
    const uid = val(e, 'UID') || crypto.createHash('sha1').update(JSON.stringify(e.props)).digest('hex');
    e.uid = uid;
    const rid = first(e, 'RECURRENCE-ID');
    if (rid) { const d = parseIcsDate(rid.value, rid.params); if (d) overrides.set(`${uid}|${d.ts}`, e); }
    else masters.push(e);
  }
  const me = String(userEmail || '').toLowerCase();
  const describe = (e, startTs, origTs) => {
    const ds = first(e, 'DTSTART');
    const dt = ds && parseIcsDate(ds.value, ds.params);
    if (!dt || dt.allDay) return null;
    const title = val(e, 'SUMMARY') || 'Meeting';
    const status = val(e, 'STATUS').toUpperCase();
    if (status === 'CANCELLED' || /^(canceled|cancelled)\s*:/i.test(title)) return null;
    const declined = me && (e.props.ATTENDEE || []).some((a) => String(a.value).toLowerCase().replace(/^mailto:/, '') === me && String(a.params.PARTSTAT || '').toUpperCase() === 'DECLINED');
    if (declined) return null;
    const de = first(e, 'DTEND');
    const dur = de ? (parseIcsDate(de.value, de.params) || {}).ts - dt.ts : parseDuration(val(e, 'DURATION'));
    const link = findMeetingLink(['X-MICROSOFT-SKYPETEAMSMEETINGURL', 'X-GOOGLE-CONFERENCE', 'URL', 'LOCATION', 'DESCRIPTION', 'X-ALT-DESC']
      .flatMap((k) => (e.props[k] || []).map((p) => icsUnescape(p.value))));
    if (!link) return null;
    return {
      key: origTs === null ? `ics:${e.uid}` : `ics:${e.uid}|${origTs}`, title: title.slice(0, 200), start: new Date(startTs).toISOString(),
      end: new Date(startTs + (dur > 0 && dur < 864e5 ? dur : 3600000)).toISOString(), meetingUrl: link.url, platform: link.platform,
    };
  };
  const out = [];
  for (const e of masters) {
    const ds = first(e, 'DTSTART');
    const dt = ds && parseIcsDate(ds.value, ds.params);
    if (!dt) continue;
    const rr = first(e, 'RRULE');
    if (!rr) {
      if (dt.ts >= from && dt.ts <= to) { const x = describe(e, dt.ts, null); if (x) out.push(x); }
      continue;
    }
    const ex = new Set();
    for (const p of e.props.EXDATE || []) for (const v of p.value.split(',')) { const d = parseIcsDate(v, { ...p.params, TZID: p.params.TZID || (dt.tz !== 'UTC' ? dt.tz : undefined) }); if (d) ex.add(d.ts); }
    for (const ts of expandRrule(rr.value, dt, from - 864e5, to, ex)) {
      const ov = overrides.get(`${e.uid}|${ts}`);
      if (ov) continue; // handled below
      if (ts >= from) { const x = describe(e, ts, ts); if (x) out.push(x); }
    }
  }
  for (const [k, e] of overrides) {
    const ds = first(e, 'DTSTART');
    const dt = ds && parseIcsDate(ds.value, ds.params);
    if (!dt || dt.ts < from || dt.ts > to) continue;
    const x = describe(e, dt.ts, Number(k.split('|').pop()));
    if (x) out.push(x);
  }
  return out.sort((a, b) => a.start.localeCompare(b.start));
}

async function fetchIcs(url) {
  const res = await fetch(url, { headers: { accept: 'text/calendar, */*' }, redirect: 'follow', signal: AbortSignal.timeout(20000) });
  if (!res.ok) throw new Error(`Your calendar link returned an error (HTTP ${res.status}). Check that it is still valid.`);
  const text = await res.text();
  if (text.length > 15 * 1024 * 1024) throw new Error('The calendar feed is too large.');
  if (!/BEGIN:VCALENDAR/i.test(text)) throw new Error('That link does not return a calendar (.ics) file.');
  return text;
}

function normalizeIcsUrl(input) {
  let s = str(input, 2000);
  if (!s) throw new HttpError(400, 'Paste your calendar link.');
  s = s.replace(/^webcals?:\/\//i, 'https://');
  let u;
  try { u = new URL(s); } catch { throw new HttpError(400, 'That is not a valid link.'); }
  if (process.env.MOMBOT_TEST_ALLOW_LOCAL_ICS === '1') return u.toString(); // automated tests only
  if (u.protocol !== 'https:') throw new HttpError(400, 'The calendar link must start with https://');
  if (/^(localhost|127\.|10\.|172\.(1[6-9]|2\d|3[01])\.|192\.168\.|169\.254\.|0\.|\[)/.test(u.hostname) || u.hostname.endsWith('.internal')) throw new HttpError(400, 'That calendar address is not allowed.');
  return u.toString();
}

/* ------------------------------------------------------------------ */
/* Calendar: Google / Outlook via Recall.ai Calendar V2                */
/* ------------------------------------------------------------------ */

const oauthConfigured = (p) => (p === 'google' ? !!(CFG.googleClientId && CFG.googleClientSecret) : !!(CFG.msClientId && CFG.msClientSecret));
const oauthRedirectUri = () => `${CFG.publicBaseUrl}/oauth/callback`;

function signState(obj) { const p = Buffer.from(JSON.stringify(obj)).toString('base64url'); return `${p}.${hmac('oauth:' + p)}`; }
function readState(s) {
  const [p, sig] = String(s || '').split('.');
  if (!p || !sig) return null;
  const a = Buffer.from(sig), b = Buffer.from(hmac('oauth:' + p));
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
  try { const o = JSON.parse(Buffer.from(p, 'base64url').toString('utf8')); return o.e > Date.now() ? o : null; } catch { return null; }
}

function oauthStartUrl(user, provider) {
  const state = signState({ u: user.id, p: provider, e: Date.now() + 15 * 60000, n: newId(6) });
  if (provider === 'google') {
    const u = new URL(CFG.googleAuthUrl);
    u.search = new URLSearchParams({
      client_id: CFG.googleClientId, redirect_uri: oauthRedirectUri(), response_type: 'code', access_type: 'offline', prompt: 'consent',
      scope: 'https://www.googleapis.com/auth/calendar.events.readonly https://www.googleapis.com/auth/userinfo.email', state,
    }).toString();
    return u.toString();
  }
  const u = new URL(`${CFG.msLoginBase}/${encodeURIComponent(CFG.msTenant)}/oauth2/v2.0/authorize`);
  u.search = new URLSearchParams({
    client_id: CFG.msClientId, redirect_uri: oauthRedirectUri(), response_type: 'code', prompt: 'consent',
    scope: 'offline_access openid email https://graph.microsoft.com/Calendars.Read', state,
  }).toString();
  return u.toString();
}

async function oauthExchange(provider, code) {
  const google = provider === 'google';
  const tokenUrl = google ? CFG.googleTokenUrl : `${CFG.msLoginBase}/${encodeURIComponent(CFG.msTenant)}/oauth2/v2.0/token`;
  const res = await fetch(tokenUrl, {
    method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' }, signal: AbortSignal.timeout(20000),
    body: new URLSearchParams({
      client_id: google ? CFG.googleClientId : CFG.msClientId, client_secret: google ? CFG.googleClientSecret : CFG.msClientSecret,
      redirect_uri: oauthRedirectUri(), grant_type: 'authorization_code', code,
    }),
  });
  const tok = await res.json().catch(() => ({}));
  if (!res.ok || !tok.refresh_token) throw new Error(`Calendar sign-in failed: ${tok.error_description || tok.error || 'no refresh token returned'}`);
  const ui = await fetch(google ? CFG.googleUserinfoUrl : CFG.msUserinfoUrl, { headers: { authorization: `Bearer ${tok.access_token}` }, signal: AbortSignal.timeout(20000) });
  const info = await ui.json().catch(() => ({}));
  return { refreshToken: tok.refresh_token, email: str(info.email || info.preferred_username, 200).toLowerCase() || null };
}

function recallEventToMeeting(e) {
  const raw = e.raw || {};
  const title = str(raw.summary || raw.subject, 200) || 'Calendar meeting';
  const cancelled = raw.status === 'cancelled' || raw.isCancelled === true;
  const declined = (Array.isArray(raw.attendees) && raw.attendees.some((a) => a && a.self && a.responseStatus === 'declined'))
    || (raw.responseStatus && raw.responseStatus.response === 'declined');
  let link = null;
  if (e.meeting_url) { try { link = parseMeetingUrl(e.meeting_url); } catch { link = null; } }
  return {
    key: `rc:${e.id}`, recallEventId: e.id, title, start: e.start_time, end: e.end_time,
    meetingUrl: link ? link.url : null, platform: link ? link.platform : null, skip: !!(e.is_deleted || cancelled || declined || raw.isAllDay),
    deleted: !!e.is_deleted,
  };
}

async function listRecallEvents(calendarId, query) {
  const out = [];
  let next = null, pages = 0;
  do {
    let p;
    if (next && /^https?:/i.test(next)) { const u = new URL(next); p = u.pathname.replace(/^\/api\/v2/, '') + u.search; }
    else {
      const qs = new URLSearchParams({ calendar_id: calendarId, ...query });
      if (next) qs.set('next', next);
      p = `/calendar-events/?${qs}`;
    }
    const page = await recall2('GET', p);
    out.push(...((page && page.results) || []));
    next = page && page.next;
  } while (next && ++pages < 50);
  return out;
}

/* ------------------------------------------------------------------ */
/* Calendar: scheduling (shared by both sources)                       */
/* ------------------------------------------------------------------ */

const calNow = () => Date.now();
const userCal = (u) => u.calendar || null;

function wantsEvent(user, ev) {
  const cal = userCal(user);
  if (!cal || ev.skip || !ev.meetingUrl) return false;
  const ov = cal.overrides && cal.overrides[ev.key];
  if (ov === true || ov === false) return ov;
  return cal.autoJoin !== 'manual';
}

function findCalMeeting(user, key) {
  return db.meetings.find((m) => m.calendarKeys && m.calendarKeys[user.id] === key && m.stage !== 'cancelled');
}

async function cancelBotRemote(m) {
  if (!m.botId && !m.recallEventId) return;
  try {
    if (m.calendarSource === 'recall' && m.recallEventId) await recall2('DELETE', `/calendar-events/${m.recallEventId}/bot/`);
    else if (m.botId) await recall('DELETE', `/bot/${m.botId}/`);
  } catch (e) { if (e.status !== 404 && e.status !== 400) throw e; }
}

/** The user no longer wants the bot for this meeting. */
async function unwantMeeting(user, m, { remoteGone = false } = {}) {
  if (m.stage !== 'scheduled' && m.stage !== 'dispatching') { if (m.calendarKeys) delete m.calendarKeys[user.id]; return; }
  const others = [m.ownerId, ...(m.sharedWith || [])].filter((id) => id !== user.id && m.calendarKeys && m.calendarKeys[id]);
  if (m.calendarKeys) delete m.calendarKeys[user.id];
  m.sharedWith = (m.sharedWith || []).filter((id) => id !== user.id);
  if (others.length) {
    if (m.ownerId === user.id) {
      const nu = db.users.find((u) => u.id === others[0]);
      if (nu) { m.ownerId = nu.id; m.ownerName = nu.name; m.ownerEmail = nu.email; m.sharedWith = m.sharedWith.filter((id) => id !== nu.id); }
    }
    touch(m);
    return;
  }
  if (!remoteGone) await cancelBotRemote(m);
  setStage(m, 'cancelled', { force: true });
  addHistory(m, 'cancelled', null, nowIso(), 'calendar');
}

async function scheduleRecallCalendarBot(m) {
  const key = `${m.meetingUrl}-${m.eventStart}`;
  const ev = await recall2('POST', `/calendar-events/${m.recallEventId}/bot/`, { deduplication_key: key, bot_config: botConfig(m) });
  const bot = ((ev && ev.bots) || []).find((b) => b.deduplication_key === key) || ((ev && ev.bots) || [])[0];
  if (bot && bot.bot_id) m.botId = bot.bot_id;
  m.error = null;
  addHistory(m, 'bot_scheduled', null, nowIso(), 'calendar');
  setStage(m, 'scheduled', { force: true });
}

/** Make sure the bot is (or is not) booked for one calendar meeting. */
async function applyCalendarEvent(user, ev, source) {
  const existing = findCalMeeting(user, ev.key);
  const startMs = new Date(ev.start).getTime();
  const want = wantsEvent(user, ev);

  if (existing) {
    if (!want || ev.skip) return unwantMeeting(user, existing, { remoteGone: ev.deleted });
    if (existing.stage !== 'scheduled') return; // already joined / finished
    const moved = existing.eventStart !== new Date(startMs).toISOString();
    const relinked = existing.meetingUrl !== ev.meetingUrl;
    if (!moved && !relinked) { if (existing.title !== ev.title && !existing.titleEdited) { existing.title = ev.title; touch(existing); } return; }
    existing.eventStart = new Date(startMs).toISOString();
    existing.joinAt = new Date(startMs - 60000).toISOString();
    existing.title = ev.title;
    if (source === 'recall') { existing.meetingUrl = ev.meetingUrl; existing.platform = ev.platform; await scheduleRecallCalendarBot(existing); }
    else if (relinked || startMs - calNow() < 2 * 60000) {
      await cancelBotRemote(existing);
      existing.botId = null; existing.meetingUrl = ev.meetingUrl; existing.platform = ev.platform;
      existing.stage = 'dispatching';
      await dispatchBot(existing);
    } else {
      try { await recall('PATCH', `/bot/${existing.botId}/`, { join_at: existing.joinAt }); }
      catch (e) { log('warn', `Rescheduling bot ${existing.botId} failed`, e.message); }
    }
    addHistory(existing, 'rescheduled', null, nowIso(), 'calendar');
    touch(existing);
    return;
  }

  if (!want) return;
  if (startMs < calNow() - 10 * 60000 || startMs > calNow() + CFG.calendarLookaheadHours * 3600000) return;

  // Same meeting already booked by a teammate (or by a pasted link)? Share it instead of sending a second bot.
  const isoStart = new Date(startMs).toISOString();
  const shared = db.meetings.find((m) => m.meetingUrl === ev.meetingUrl && ['dispatching', 'scheduled', 'joining', 'waiting_room', 'in_call', 'recording'].includes(m.stage)
    && (!m.eventStart || Math.abs(new Date(m.eventStart) - startMs) < 30 * 60000));
  if (shared) {
    shared.calendarKeys = shared.calendarKeys || {};
    shared.calendarKeys[user.id] = ev.key;
    if (shared.ownerId !== user.id && !(shared.sharedWith || []).includes(user.id)) (shared.sharedWith = shared.sharedWith || []).push(user.id);
    touch(shared);
    return;
  }

  const m = {
    id: newId(), ownerId: user.id, ownerName: user.name, ownerEmail: user.email,
    title: ev.title, platform: ev.platform, meetingUrl: ev.meetingUrl, agenda: '',
    source: 'calendar', calendarSource: source, recallEventId: ev.recallEventId || null, calendarKeys: { [user.id]: ev.key },
    eventStart: isoStart, joinAt: new Date(startMs - 60000).toISOString(), sharedWith: [],
    stage: 'dispatching', statusHistory: [], recordings: {}, warnings: [], createdAt: nowIso(), updatedAt: nowIso(),
  };
  db.meetings.push(m);
  if (source === 'recall') {
    try { await scheduleRecallCalendarBot(m); }
    catch (e) { m.error = friendlyRecallError(e); setStage(m, 'failed', { force: true }); }
  } else {
    await dispatchBot(m);
  }
  log('info', `Calendar: booked "${m.title}" at ${isoStart} for ${user.email}`);
}

async function syncIcs(user) {
  const cal = userCal(user);
  if (!cal || cal.type !== 'ics') return;
  const now = calNow();
  try {
    const text = await fetchIcs(cal.icsUrl);
    const horizon = Math.max(CFG.calendarLookaheadHours, 7 * 24) * 3600000;
    const all = icsToMeetings(text, now - 15 * 60000, now + horizon, user.email);
    cal.upcoming = all.slice(0, 60);
    cal.lastSyncAt = nowIso();
    cal.lastError = null;
    const windowEnd = now + CFG.calendarLookaheadHours * 3600000;
    const seen = new Set();
    for (const ev of all) {
      if (new Date(ev.start).getTime() > windowEnd) continue;
      seen.add(ev.key);
      try { await applyCalendarEvent(user, ev, 'ics'); } catch (e) { log('warn', `Calendar booking failed for "${ev.title}"`, e.message); }
    }
    // Meetings that disappeared from the calendar (cancelled / declined) -> cancel the bot.
    for (const m of db.meetings) {
      const key = m.calendarKeys && m.calendarKeys[user.id];
      if (!key || !key.startsWith('ics:') || seen.has(key) || m.stage !== 'scheduled') continue;
      const st = new Date(m.eventStart).getTime();
      if (st > now && st <= windowEnd) { try { await unwantMeeting(user, m); } catch (e) { log('warn', 'Cancelling removed event failed', e.message); } }
    }
  } catch (e) {
    cal.lastError = e.message;
    cal.lastSyncAt = nowIso();
    log('warn', `Calendar sync failed for ${user.email}`, e.message);
  }
  await save();
}

async function syncRecallCalendar(user, { since } = {}) {
  const cal = userCal(user);
  if (!cal || cal.type !== 'recall' || !cal.recallCalendarId) return;
  const now = calNow();
  try {
    const query = since ? { updated_at__gte: since }
      : { start_time__gte: new Date(now - 15 * 60000).toISOString(), start_time__lte: new Date(now + Math.max(CFG.calendarLookaheadHours, 168) * 3600000).toISOString() };
    const events = (await listRecallEvents(cal.recallCalendarId, query)).map(recallEventToMeeting);
    if (!since) cal.upcoming = events.filter((e) => !e.skip && e.meetingUrl).sort((a, b) => a.start.localeCompare(b.start)).slice(0, 60);
    else {
      const map = new Map((cal.upcoming || []).map((e) => [e.key, e]));
      for (const e of events) { if (e.skip || !e.meetingUrl) map.delete(e.key); else map.set(e.key, e); }
      cal.upcoming = [...map.values()].filter((e) => new Date(e.end || e.start).getTime() > now).sort((a, b) => a.start.localeCompare(b.start)).slice(0, 60);
    }
    for (const ev of events) {
      try { await applyCalendarEvent(user, ev, 'recall'); } catch (e) { log('warn', `Calendar booking failed for "${ev.title}"`, e.message); }
    }
    cal.lastSyncAt = nowIso();
    cal.lastError = null;
  } catch (e) {
    cal.lastError = e.message;
    log('warn', `Recall calendar sync failed for ${user.email}`, e.message);
  }
  await save();
}

const syncCalendar = (user, opts) => withLock('calendar', () => (userCal(user) && userCal(user).type === 'ics' ? syncIcs(user) : syncRecallCalendar(user, opts)));

async function handleCalendarWebhook(type, d) {
  const user = db.users.find((u) => u.calendar && u.calendar.recallCalendarId && u.calendar.recallCalendarId === d.calendar_id);
  if (!user) { log('info', `Ignoring ${type} for unknown calendar ${d.calendar_id}`); return; }
  if (type === 'calendar.update') {
    try {
      const c = await recall2('GET', `/calendars/${d.calendar_id}/`);
      const last = ((c && c.status_changes) || []).slice(-1)[0];
      user.calendar.status = (last && last.status) || (c && c.status) || user.calendar.status;
      if (c && c.platform_email) user.calendar.email = c.platform_email;
      user.calendar.lastError = user.calendar.status === 'disconnected' ? 'Your calendar was disconnected. Please connect it again.' : null;
      await save();
    } catch (e) { log('warn', 'calendar.update lookup failed', e.message); }
    return;
  }
  await syncCalendar(user, { since: d.last_updated_ts });
}

function calendarView(user) {
  const cal = userCal(user);
  if (!cal) return null;
  const upcoming = (cal.upcoming || []).filter((e) => new Date(e.end || e.start).getTime() > calNow()).map((e) => {
    const m = findCalMeeting(user, e.key);
    return { ...e, join: wantsEvent(user, e), meetingId: m ? m.id : null, stage: m ? m.stage : null };
  });
  let icsHost = null;
  if (cal.icsUrl) { try { icsHost = new URL(cal.icsUrl).hostname; } catch { /* ignore */ } }
  return { type: cal.type, provider: cal.provider || null, email: cal.email || null, status: cal.status || null, icsHost, autoJoin: cal.autoJoin || 'all',
    lastSyncAt: cal.lastSyncAt || null, lastError: cal.lastError || null, upcoming };
}

async function disconnectCalendar(user) {
  const cal = userCal(user);
  if (!cal) return;
  if (cal.type === 'recall' && cal.recallCalendarId) {
    try { await recall2('DELETE', `/calendars/${cal.recallCalendarId}/`); } catch (e) { if (e.status !== 404) log('warn', 'Deleting Recall calendar failed', e.message); }
  }
  for (const m of db.meetings) {
    if (m.calendarKeys && m.calendarKeys[user.id] && m.stage === 'scheduled') {
      try { await unwantMeeting(user, m, { remoteGone: cal.type === 'recall' && m.calendarSource === 'recall' }); } catch (e) { log('warn', 'cancel on disconnect failed', e.message); }
    }
  }
  user.calendar = null;
  await save();
}

/* Background jobs: calendar feeds + a safety net for missed webhooks. */
let jobsRunning = false;
async function backgroundJobs({ recallSweep = false } = {}) {
  if (jobsRunning) return;
  jobsRunning = true;
  try {
    for (const u of db.users) {
      if (u.disabled || !u.calendar) continue;
      if (u.calendar.type === 'ics' || (recallSweep && u.calendar.type === 'recall')) await syncCalendar(u);
    }
    const now = calNow();
    const stuck = db.meetings.filter((m) => m.botId && (
      (['processing', 'transcribing'].includes(m.stage) && now - new Date(m.updatedAt) > 20 * 60000)
      || (m.stage === 'scheduled' && m.joinAt && now - new Date(m.joinAt) > 45 * 60000)
      || (['joining', 'waiting_room', 'in_call', 'recording'].includes(m.stage) && now - new Date(m.updatedAt) > 6 * 3600000)
    ) && (!m.lastAutoSync || now - new Date(m.lastAutoSync) > 30 * 60000)).slice(0, 10);
    for (const m of stuck) {
      m.lastAutoSync = nowIso();
      await withLock(m.id, async () => {
        try { await syncFromRecall(m); log('info', `Auto-recovered meeting ${m.id} (${m.stage})`); }
        catch (e) { log('warn', `Auto-recovery failed for ${m.id}`, e.message); }
        touch(m);
      });
    }
    if (stuck.length) await save();
  } catch (e) {
    log('error', 'Background job failed', e.stack || e.message);
  } finally { jobsRunning = false; }
}

/* ------------------------------------------------------------------ */
/* HTTP plumbing                                                       */
/* ------------------------------------------------------------------ */

const SECURITY_HEADERS = {
  'X-Content-Type-Options': 'nosniff',
  'X-Frame-Options': 'DENY',
  'Referrer-Policy': 'no-referrer',
  'Permissions-Policy': 'camera=(), microphone=(), geolocation=()',
  'Content-Security-Policy': "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'",
};

function send(res, status, body, type = 'application/json; charset=utf-8', extra = {}) {
  if (res.headersSent) return;
  const payload = typeof body === 'string' || Buffer.isBuffer(body) ? body : JSON.stringify(body);
  res.writeHead(status, { ...SECURITY_HEADERS, 'Content-Type': type, 'Cache-Control': 'no-store', ...extra });
  res.end(payload);
}

function readBody(req, limit = 1024 * 1024) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on('data', (c) => {
      size += c.length;
      if (size > limit) { reject(new HttpError(413, 'Request is too large.')); req.destroy(); return; }
      chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

async function readJson(req, limit) {
  const buf = await readBody(req, limit);
  if (!buf.length) return {};
  try { return JSON.parse(buf.toString('utf8')); } catch { throw new HttpError(400, 'Invalid JSON body.'); }
}

function clientIp(req) {
  const xf = String(req.headers['x-forwarded-for'] || '').split(',')[0].trim();
  return xf || req.socket.remoteAddress || 'unknown';
}

function baseUrl(req) {
  if (CFG.publicBaseUrl) return CFG.publicBaseUrl;
  const proto = req.headers['x-forwarded-proto'] || 'http';
  return `${proto}://${req.headers.host}`;
}

let INDEX_HTML = null;
function indexHtml() {
  if (!INDEX_HTML || process.env.NODE_ENV === 'development') INDEX_HTML = fs.readFileSync(path.join(ROOT, 'index.html'));
  return INDEX_HTML;
}

/* ------------------------------------------------------------------ */
/* Views                                                               */
/* ------------------------------------------------------------------ */

function canSee(user, m) { return user.role === 'admin' || m.ownerId === user.id || (m.sharedWith || []).includes(user.id); }

function meetingSummary(m) {
  return {
    id: m.id, title: m.title, platform: m.platform, meetingUrl: m.meetingUrl, joinAt: m.eventStart || m.joinAt || null,
    source: m.source || 'link', botName: m.botName || null,
    stage: m.stage, stageLabel: STAGE_LABEL[m.stage] || m.stage, error: m.error || null,
    createdAt: m.createdAt, updatedAt: m.updatedAt, ownerName: m.ownerName, ownerEmail: m.ownerEmail,
    manual: !!m.manual, hasMom: !!m.momMarkdown,
  };
}

function meetingDetail(m) {
  const segs = combinedSegments(m);
  return {
    ...meetingSummary(m),
    agenda: m.agenda || '',
    warnings: m.warnings || [],
    statusHistory: (m.statusHistory || []).map((h) => ({ ...h, text: h.subCode ? subCodeText(h.subCode) : '' })),
    botId: m.botId || null,
    momMarkdown: m.momMarkdown || '',
    mom: m.mom || null,
    momEdited: !!m.momEdited,
    momGeneratedAt: m.momGeneratedAt || null,
    momError: m.momError || null,
    transcript: segs.map((s) => ({ speaker: s.speaker, start: s.start, time: fmtClock(s.start), text: s.text })),
    canLeave: !!m.botId && ['joining', 'waiting_room', 'in_call', 'recording'].includes(m.stage),
    canCancel: !!m.botId && m.stage === 'scheduled',
    canRetry: !m.manual && (m.stage === 'failed' || ['processing', 'transcribing'].includes(m.stage)),
    canRegenerate: segs.length > 0 && m.stage !== 'generating',
  };
}

/* ------------------------------------------------------------------ */
/* Routes                                                              */
/* ------------------------------------------------------------------ */

const routes = [];
function route(method, pattern, handler, opts = {}) {
  const keys = [];
  const re = new RegExp('^' + pattern.replace(/:(\w+)/g, (_, k) => { keys.push(k); return '([^/]+)'; }) + '/?$');
  routes.push({ method, re, keys, handler, auth: opts.auth || 'user' });
}

function getMeetingFor(user, id) {
  const m = db.meetings.find((x) => x.id === id);
  if (!m || !canSee(user, m)) throw new HttpError(404, 'Meeting not found.');
  return m;
}

/* --- public --- */
route('GET', '/api/health', async () => ({ ok: true }), { auth: 'none' });

route('GET', '/api/config', async ({ user }) => ({
  recallConfigured: recallConfigured(),
  webhookVerificationConfigured: !!(CFG.recallVerificationSecret || CFG.recallSvixSecret),
  aiConfigured: llmConfigured(),
  aiProvider: CFG.llmProvider,
  botName: user ? botNameFor(user) : null,
  calendarSignIn: { google: oauthConfigured('google'), outlook: oauthConfigured('outlook') },
  timezone: CFG.timezone,
  needsSetup: db.users.length === 0,
  publicBaseUrlSet: !!CFG.publicBaseUrl,
  webhookUrl: user && user.role === 'admin' ? `${CFG.publicBaseUrl || '<PUBLIC_BASE_URL>'}/webhooks/recall` : undefined,
}), { auth: 'optional' });

route('POST', '/api/login', async ({ req, res, body }) => {
  const email = str(body.email, 200).toLowerCase();
  rateLimit(`login:${clientIp(req)}`, 20, 15 * 60000);
  rateLimit(`login:${email}`, 8, 15 * 60000);
  const user = db.users.find((u) => u.email === email);
  const ok = await verifyPassword(String(body.password || ''), user && user.passwordHash);
  if (!user || !ok) throw new HttpError(401, 'Wrong email or password.');
  if (user.disabled) throw new HttpError(403, 'This account has been disabled. Contact your administrator.');
  user.lastLoginAt = nowIso();
  save();
  setSessionCookie(req, res, user);
  return { user: publicUser(user) };
}, { auth: 'none' });

route('POST', '/api/logout', async ({ req, res }) => { clearSessionCookie(req, res); return { ok: true }; }, { auth: 'none' });

route('GET', '/api/me', async ({ user }) => ({ user: user ? publicUser(user) : null }), { auth: 'optional' });

route('POST', '/api/me/password', async ({ req, res, user, body }) => {
  if (!(await verifyPassword(String(body.currentPassword || ''), user.passwordHash))) throw new HttpError(400, 'Your current password is wrong.');
  validatePassword(body.newPassword);
  user.passwordHash = await hashPassword(body.newPassword);
  user.sessionVersion = (user.sessionVersion || 0) + 1; // signs out other devices
  await save();
  setSessionCookie(req, res, user);
  return { ok: true };
});

function findInvite(token) {
  const inv = db.invites.find((i) => i.tokenHash === sha256(token || ''));
  if (!inv) throw new HttpError(404, 'This invite link is not valid.');
  if (inv.revokedAt) throw new HttpError(410, 'This invite has been revoked.');
  if (inv.usedAt) throw new HttpError(410, 'This invite link has already been used.');
  if (new Date(inv.expiresAt) < new Date()) throw new HttpError(410, 'This invite link has expired. Ask for a new one.');
  return inv;
}

route('GET', '/api/invites/:token', async ({ params, req }) => {
  rateLimit(`invite:${clientIp(req)}`, 30, 15 * 60000);
  const inv = findInvite(params.token);
  return { email: inv.email || null, role: inv.role, expiresAt: inv.expiresAt };
}, { auth: 'none' });

route('POST', '/api/register', async ({ req, res, body }) => {
  rateLimit(`register:${clientIp(req)}`, 10, 15 * 60000);
  const inv = findInvite(str(body.token, 200));
  if (inv.email && body.email && str(body.email, 200).toLowerCase() !== inv.email) {
    throw new HttpError(400, 'This invite is for a different email address.');
  }
  const email = validateEmail(inv.email || body.email);
  if (db.users.some((u) => u.email === email)) throw new HttpError(409, 'An account with this email already exists. Please sign in.');
  const name = str(body.name, 100);
  if (!name) throw new HttpError(400, 'Enter your name.');
  validatePassword(body.password);
  const user = {
    id: newId(), email, name, role: inv.role === 'admin' ? 'admin' : 'member',
    passwordHash: await hashPassword(body.password), sessionVersion: 0, createdAt: nowIso(), invitedBy: inv.createdBy,
  };
  db.users.push(user);
  inv.usedAt = nowIso();
  inv.usedBy = user.id;
  await save();
  setSessionCookie(req, res, user);
  log('info', `New user registered via invite: ${email}`);
  return { user: publicUser(user) };
}, { auth: 'none' });

/* --- admin --- */
route('GET', '/api/admin/users', async () => ({ users: db.users.map(publicUser) }), { auth: 'admin' });

route('POST', '/api/admin/users/:id/status', async ({ params, body, user }) => {
  const target = db.users.find((u) => u.id === params.id);
  if (!target) throw new HttpError(404, 'User not found.');
  if (target.id === user.id) throw new HttpError(400, 'You cannot disable your own account.');
  if (body.role && ['admin', 'member'].includes(body.role)) {
    if (target.role === 'admin' && body.role === 'member' && db.users.filter((u) => u.role === 'admin' && !u.disabled).length <= 1) {
      throw new HttpError(400, 'At least one active admin is required.');
    }
    target.role = body.role;
  }
  if (typeof body.disabled === 'boolean') {
    target.disabled = body.disabled;
    target.sessionVersion = (target.sessionVersion || 0) + 1;
  }
  await save();
  return { user: publicUser(target) };
}, { auth: 'admin' });

route('GET', '/api/admin/invites', async () => ({
  invites: db.invites.slice().reverse().slice(0, 200).map((i) => ({
    id: i.id, email: i.email, role: i.role, createdAt: i.createdAt, expiresAt: i.expiresAt,
    status: i.revokedAt ? 'revoked' : i.usedAt ? 'used' : new Date(i.expiresAt) < new Date() ? 'expired' : 'active',
    usedBy: i.usedBy ? (publicUser(db.users.find((u) => u.id === i.usedBy)) || {}).email : null,
  })),
}), { auth: 'admin' });

route('POST', '/api/admin/invites', async ({ req, body, user }) => {
  const email = body.email ? validateEmail(body.email) : null;
  if (email && db.users.some((u) => u.email === email)) throw new HttpError(409, 'That person already has an account.');
  const days = Math.min(30, Math.max(1, parseInt(body.days, 10) || CFG.inviteDaysDefault));
  const token = newId(24);
  const inv = {
    id: newId(), tokenHash: sha256(token), email, role: body.role === 'admin' ? 'admin' : 'member',
    createdBy: user.id, createdAt: nowIso(), expiresAt: new Date(Date.now() + days * 864e5).toISOString(),
  };
  db.invites.push(inv);
  await save();
  return { id: inv.id, link: `${baseUrl(req)}/#/register/${token}`, expiresAt: inv.expiresAt, email, role: inv.role };
}, { auth: 'admin' });

route('DELETE', '/api/admin/invites/:id', async ({ params }) => {
  const inv = db.invites.find((i) => i.id === params.id);
  if (!inv) throw new HttpError(404, 'Invite not found.');
  inv.revokedAt = inv.revokedAt || nowIso();
  await save();
  return { ok: true };
}, { auth: 'admin' });

/* --- meetings --- */
route('GET', '/api/meetings', async ({ user }) => ({
  meetings: db.meetings.filter((m) => canSee(user, m))
    .sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)))
    .slice(0, 300).map(meetingSummary),
}));

route('POST', '/api/meetings', async ({ user, body }) => {
  rateLimit(`create:${user.id}`, 30, 60 * 60000);
  if (!recallConfigured()) throw new HttpError(503, 'The meeting bot is not configured yet (RECALL_API_KEY). Ask your administrator.');
  const { url, platform } = parseMeetingUrl(body.meetingUrl);
  const title = str(body.title, 200) || `${platform} meeting`;
  let joinAt = null;
  if (body.joinAt) {
    const t = new Date(body.joinAt);
    if (Number.isNaN(t.getTime())) throw new HttpError(400, 'The scheduled time is not valid.');
    if (t.getTime() < Date.now() - 5 * 60000) throw new HttpError(400, 'The scheduled time is in the past. Choose "Join now" instead.');
    if (t.getTime() > Date.now() + 90 * 864e5) throw new HttpError(400, 'You can schedule at most 90 days ahead.');
    joinAt = t.toISOString();
  }
  const active = db.meetings.find((m) => m.meetingUrl === url && !FINAL_STAGES.has(m.stage) && m.ownerId === user.id && m.stage !== 'scheduled');
  if (active && !joinAt) throw new HttpError(409, 'The bot is already on its way to this meeting.');
  const m = {
    id: newId(), ownerId: user.id, ownerName: user.name, ownerEmail: user.email,
    title, platform, meetingUrl: url, agenda: str(body.agenda, 5000), joinAt,
    stage: 'dispatching', statusHistory: [], recordings: {}, warnings: [],
    createdAt: nowIso(), updatedAt: nowIso(),
  };
  db.meetings.push(m);
  await save();
  dispatchBot(m).catch((e) => log('error', 'dispatchBot crashed', e.message));
  return { meeting: meetingSummary(m) };
});

route('POST', '/api/meetings/manual', async ({ user, body }) => {
  rateLimit(`create:${user.id}`, 30, 60 * 60000);
  const segs = parseManualTranscript(str(body.transcript, 2_000_000));
  if (!segs.length) throw new HttpError(400, 'Paste a transcript (plain text "Name: words" lines, or a Teams/Zoom .vtt file).');
  const m = {
    id: newId(), ownerId: user.id, ownerName: user.name, ownerEmail: user.email,
    title: str(body.title, 200) || 'Imported meeting', platform: 'Imported transcript', meetingUrl: '', agenda: str(body.agenda, 5000),
    stage: 'generating', manual: true, manualSegments: segs, statusHistory: [], recordings: {}, warnings: [],
    createdAt: nowIso(), updatedAt: nowIso(), botDone: true,
  };
  db.meetings.push(m);
  await save();
  withLock(m.id, () => generateMom(m)).catch((e) => log('error', 'generateMom crashed', e.message));
  return { meeting: meetingSummary(m) };
}, { bodyLimit: 3 * 1024 * 1024 });

route('GET', '/api/meetings/:id', async ({ user, params }) => ({ meeting: meetingDetail(getMeetingFor(user, params.id)) }));

route('POST', '/api/meetings/:id/leave', async ({ user, params }) => {
  const m = getMeetingFor(user, params.id);
  if (!m.botId) throw new HttpError(400, 'The bot is not in this meeting.');
  await recall('POST', `/bot/${m.botId}/leave_call/`);
  addHistory(m, 'leave_requested', null, nowIso(), 'app');
  touch(m);
  await save();
  return { meeting: meetingDetail(m) };
});

route('POST', '/api/meetings/:id/cancel', async ({ user, params }) => {
  const m = getMeetingFor(user, params.id);
  if (m.stage !== 'scheduled' && m.stage !== 'dispatching') throw new HttpError(400, 'Only a scheduled bot can be cancelled. Use "Remove bot" if it already joined.');
  const key = m.calendarKeys && m.calendarKeys[user.id];
  if (key && user.calendar) {
    // remember the choice so the next calendar sync does not book it again
    user.calendar.overrides = user.calendar.overrides || {};
    user.calendar.overrides[key] = false;
    await withLock('calendar', () => unwantMeeting(user, m));
  } else {
    try { await cancelBotRemote(m); }
    catch (e) { throw new HttpError(502, `Could not cancel the bot: ${e.message}`); }
    setStage(m, 'cancelled', { force: true });
    addHistory(m, 'cancelled', null, nowIso(), 'app');
  }
  await save();
  return { meeting: meetingDetail(m), stillBookedForTeammates: m.stage !== 'cancelled' };
});

route('POST', '/api/meetings/:id/retry', async ({ user, params }) => {
  const m = getMeetingFor(user, params.id);
  if (m.manual) throw new HttpError(400, 'Use "Regenerate MOM" for imported transcripts.');
  if (!m.botId) {
    if (m.joinAt && new Date(m.joinAt) < new Date()) m.joinAt = null;
    setStage(m, 'dispatching', { force: true });
    m.error = null;
    await save();
    await dispatchBot(m);
    return { meeting: meetingDetail(m) };
  }
  await withLock(m.id, async () => {
    if (m.stage === 'failed') {
      // let transcripts be retried
      for (const r of Object.values(m.recordings || {})) if (r.transcriptStatus === 'failed') r.transcriptStatus = null;
      if (m.everRecorded || Object.keys(m.recordings || {}).length) { m.stage = 'processing'; m.error = null; }
    }
    try { await syncFromRecall(m); }
    catch (e) { addWarning(m, `Sync failed: ${e.message}`); throw new HttpError(502, e.message); }
    finally { touch(m); await save(); }
  });
  return { meeting: meetingDetail(m) };
});

route('POST', '/api/meetings/:id/regenerate', async ({ user, params }) => {
  const m = getMeetingFor(user, params.id);
  if (!combinedSegments(m).length) throw new HttpError(400, 'There is no transcript yet.');
  if (m.stage === 'generating') throw new HttpError(409, 'The minutes are already being written.');
  if (!llmConfigured()) throw new HttpError(503, 'AI is not configured on the server.');
  setStage(m, 'generating', { force: true });
  await save();
  withLock(m.id, () => generateMom(m)).catch((e) => log('error', 'generateMom crashed', e.message));
  return { meeting: meetingDetail(m) };
});

route('PUT', '/api/meetings/:id/mom', async ({ user, params, body }) => {
  const m = getMeetingFor(user, params.id);
  const md = str(body.markdown, 200000);
  if (!md) throw new HttpError(400, 'The minutes cannot be empty.');
  m.momMarkdown = md;
  m.momEdited = true;
  if (body.title) m.title = str(body.title, 200);
  touch(m);
  await save();
  return { meeting: meetingDetail(m) };
});

route('DELETE', '/api/meetings/:id', async ({ user, params }) => {
  const m = getMeetingFor(user, params.id);
  if (m.ownerId !== user.id && (m.sharedWith || []).includes(user.id)) {
    // a teammate's shared copy: just hide it for this user
    m.sharedWith = m.sharedWith.filter((id) => id !== user.id);
    if (m.calendarKeys) delete m.calendarKeys[user.id];
    await save();
    return { ok: true };
  }
  if (m.botId && m.stage === 'scheduled') {
    try { await cancelBotRemote(m); } catch (e) { log('warn', 'cancel before delete failed', e.message); }
    const key = m.calendarKeys && m.calendarKeys[user.id];
    if (key && user.calendar) { user.calendar.overrides = user.calendar.overrides || {}; user.calendar.overrides[key] = false; }
  } else if (m.botId && ['joining', 'waiting_room', 'in_call', 'recording'].includes(m.stage)) {
    try { await recall('POST', `/bot/${m.botId}/leave_call/`); } catch (e) { log('warn', 'leave before delete failed', e.message); }
  } else if (m.botId && m.botDone) {
    // privacy: remove the recording stored at Recall.ai
    try { await recall('POST', `/bot/${m.botId}/delete_media/`); } catch (e) { log('warn', 'delete_media failed', e.message); }
  }
  db.meetings = db.meetings.filter((x) => x.id !== m.id);
  await save();
  return { ok: true };
});

/* --- my notetaker settings + calendar --- */
route('POST', '/api/me/settings', async ({ user, body }) => {
  if (body.botName !== undefined) {
    const n = str(body.botName, 100);
    if (n && n.length < 2) throw new HttpError(400, 'The notetaker name is too short.');
    user.botName = n || null;
  }
  await save();
  return { user: publicUser(user) };
});

const calendarPayload = (user) => ({
  calendar: calendarView(user), google: oauthConfigured('google'), outlook: oauthConfigured('outlook'),
  lookaheadHours: CFG.calendarLookaheadHours, pollMinutes: CFG.icsPollMinutes, oauthReady: !!CFG.publicBaseUrl && recallConfigured(),
});

route('GET', '/api/calendar', async ({ user }) => calendarPayload(user));

route('POST', '/api/calendar/ics', async ({ user, body, req }) => {
  rateLimit(`cal:${user.id}`, 20, 15 * 60000);
  if (!recallConfigured()) throw new HttpError(503, 'The meeting bot is not configured yet (RECALL_API_KEY).');
  const url = normalizeIcsUrl(body.url);
  let text;
  try { text = await fetchIcs(url); } catch (e) { throw new HttpError(400, e.message); }
  const count = icsToMeetings(text, Date.now(), Date.now() + 14 * 864e5, user.email).length;
  if (userCal(user)) await withLock('calendar', () => disconnectCalendar(user));
  user.calendar = { type: 'ics', icsUrl: url, autoJoin: body.autoJoin === 'manual' ? 'manual' : 'all', overrides: {}, connectedAt: nowIso() };
  await save();
  await syncCalendar(user);
  log('info', `Calendar link connected for ${user.email} (${count} meetings with links in the next 14 days, from ${clientIp(req)})`);
  return calendarPayload(user);
});

route('POST', '/api/calendar/settings', async ({ user, body }) => {
  const cal = userCal(user);
  if (!cal) throw new HttpError(400, 'Connect a calendar first.');
  if (body.autoJoin) cal.autoJoin = body.autoJoin === 'manual' ? 'manual' : 'all';
  cal.overrides = {}; // a new rule replaces one-off choices
  await save();
  await syncCalendar(user);
  return calendarPayload(user);
});

route('POST', '/api/calendar/sync', async ({ user }) => {
  rateLimit(`calsync:${user.id}`, 30, 15 * 60000);
  if (!userCal(user)) throw new HttpError(400, 'Connect a calendar first.');
  await syncCalendar(user);
  return calendarPayload(user);
});

route('POST', '/api/calendar/toggle', async ({ user, body }) => {
  const cal = userCal(user);
  if (!cal) throw new HttpError(400, 'Connect a calendar first.');
  const key = str(body.key, 600);
  const ev = (cal.upcoming || []).find((e) => e.key === key);
  if (!ev) throw new HttpError(404, 'That meeting is no longer in your calendar. Refresh and try again.');
  cal.overrides = cal.overrides || {};
  cal.overrides[key] = !!body.join;
  const big = Object.entries(cal.overrides);
  if (big.length > 500) cal.overrides = Object.fromEntries(big.slice(-400));
  await withLock('calendar', async () => {
    const startMs = new Date(ev.start).getTime();
    if (body.join && startMs > Date.now() + CFG.calendarLookaheadHours * 3600000) return; // booked automatically closer to the time
    await applyCalendarEvent(user, ev, cal.type === 'ics' ? 'ics' : 'recall');
  });
  await save();
  return calendarPayload(user);
});

route('DELETE', '/api/calendar', async ({ user }) => {
  await withLock('calendar', () => disconnectCalendar(user));
  return calendarPayload(user);
});

route('GET', '/api/calendar/connect/:provider', async ({ user, params, res }) => {
  const provider = params.provider === 'google' ? 'google' : params.provider === 'outlook' ? 'outlook' : null;
  if (!provider) throw new HttpError(404, 'Unknown calendar.');
  if (!oauthConfigured(provider)) throw new HttpError(503, `${provider === 'google' ? 'Google' : 'Outlook'} calendar sign-in is not set up on this server.`);
  if (!CFG.publicBaseUrl) throw new HttpError(503, 'PUBLIC_BASE_URL must be set for calendar sign-in.');
  res.writeHead(302, { Location: oauthStartUrl(user, provider), 'Cache-Control': 'no-store' });
  res.end();
});

async function handleOauthCallback(req, res, url) {
  const back = (q) => { res.writeHead(302, { Location: `/#/notetaker?${q}`, 'Cache-Control': 'no-store' }); res.end(); };
  try {
    if (url.searchParams.get('error')) return back(`calendarError=${encodeURIComponent('Calendar access was not granted.')}`);
    const st = readState(url.searchParams.get('state'));
    if (!st) return back(`calendarError=${encodeURIComponent('The sign-in link expired. Please try again.')}`);
    const user = db.users.find((u) => u.id === st.u && !u.disabled);
    const cookieUser = userFromRequest(req);
    if (!user || (cookieUser && cookieUser.id !== user.id)) return back(`calendarError=${encodeURIComponent('Please sign in to MOM Bot first, then connect your calendar.')}`);
    const { refreshToken, email } = await oauthExchange(st.p, str(url.searchParams.get('code'), 4000));
    const body = {
      platform: st.p === 'google' ? 'google_calendar' : 'microsoft_outlook',
      oauth_client_id: st.p === 'google' ? CFG.googleClientId : CFG.msClientId,
      oauth_client_secret: st.p === 'google' ? CFG.googleClientSecret : CFG.msClientSecret,
      oauth_refresh_token: refreshToken,
      ...(email ? { platform_email: email } : {}),
    };
    const prev = userCal(user);
    let calId = prev && prev.type === 'recall' && prev.provider === st.p ? prev.recallCalendarId : null;
    if (calId) {
      try { await recall2('PATCH', `/calendars/${calId}/`, body); } catch (e) { if (e.status === 404) calId = null; else throw e; }
    }
    if (prev && !calId) await withLock('calendar', () => disconnectCalendar(user));
    if (!calId) { const c = await recall2('POST', '/calendars/', body); calId = c && c.id; }
    if (!calId) throw new Error('Recall.ai did not return a calendar id.');
    user.calendar = {
      type: 'recall', provider: st.p, recallCalendarId: calId, email, status: 'connecting',
      autoJoin: (prev && prev.autoJoin) || 'all', overrides: (prev && prev.overrides) || {}, connectedAt: nowIso(),
    };
    await save();
    log('info', `${st.p} calendar connected for ${user.email} (${email})`);
    setTimeout(() => syncCalendar(user).catch(() => {}), 3000);
    return back('connected=1');
  } catch (e) {
    log('error', 'Calendar connection failed', e.message);
    return back(`calendarError=${encodeURIComponent(e.message.slice(0, 300))}`);
  }
}

/* --- Recall webhook --- */
async function handleWebhook(req, res) {
  let raw;
  try { raw = (await readBody(req, 5 * 1024 * 1024)).toString('utf8'); }
  catch { return send(res, 413, { error: 'too large' }); }
  const headers = {};
  for (const [k, v] of Object.entries(req.headers)) headers[k.toLowerCase()] = Array.isArray(v) ? v.join(',') : String(v || '');
  const problem = verifyRecallRequest(headers, raw);
  if (problem) {
    log('warn', `Rejected webhook: ${problem}`, raw.slice(0, 500));
    return send(res, 401, { error: 'verification failed' });
  }
  let evt;
  try { evt = JSON.parse(raw); } catch { return send(res, 400, { error: 'invalid json' }); }
  const msgId = headers['webhook-id'] || headers['svix-id'];
  if (db.processedEvents.includes(msgId)) return send(res, 200, { ok: true, duplicate: true });
  db.processedEvents.push(msgId);
  if (db.processedEvents.length > 5000) db.processedEvents = db.processedEvents.slice(-4000);
  send(res, 200, { ok: true }); // acknowledge first, process in the background
  setImmediate(() => handleRecallEvent(evt).catch((e) => log('error', 'Webhook processing crashed', e.stack || e.message)));
}

/* --- dispatcher --- */
const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, 'http://localhost');
  const pathname = url.pathname;
  try {
    if (req.method === 'POST' && (pathname === '/webhooks/recall' || pathname === '/webhooks/recall/')) return await handleWebhook(req, res);
    if (req.method === 'GET' && (pathname === '/health' || pathname === '/healthz')) return send(res, 200, 'ok', 'text/plain');
    if (req.method === 'GET' && pathname === '/oauth/callback') return await handleOauthCallback(req, res, url);
    if (req.method === 'GET' && (pathname === '/' || pathname === '/index.html')) return send(res, 200, indexHtml(), 'text/html; charset=utf-8');
    if (req.method === 'GET' && pathname === '/robots.txt') return send(res, 200, 'User-agent: *\nDisallow: /\n', 'text/plain');

    if (!pathname.startsWith('/api/')) return send(res, 404, 'Not found', 'text/plain');

    const r = routes.find((x) => x.method === req.method && x.re.test(pathname));
    if (!r) throw new HttpError(404, 'Not found.');

    if (req.method !== 'GET' && req.headers['x-requested-with'] !== 'mombot') throw new HttpError(403, 'Missing request header.');

    const user = userFromRequest(req);
    if (r.auth === 'user' && !user) throw new HttpError(401, 'Please sign in.');
    if (r.auth === 'admin' && (!user || user.role !== 'admin')) throw new HttpError(user ? 403 : 401, user ? 'Admins only.' : 'Please sign in.');

    const match = pathname.match(r.re);
    const params = {};
    r.keys.forEach((k, i) => { params[k] = decodeURIComponent(match[i + 1]); });
    const body = ['POST', 'PUT', 'PATCH', 'DELETE'].includes(req.method)
      ? await readJson(req, pathname === '/api/meetings/manual' ? 3 * 1024 * 1024 : 256 * 1024) : {};
    const out = await r.handler({ req, res, user, params, body, query: url.searchParams });
    return send(res, 200, out === undefined ? { ok: true } : out);
  } catch (e) {
    const status = e instanceof HttpError ? e.status : 500;
    if (status >= 500) log('error', `${req.method} ${pathname} failed`, e.stack || e.message);
    return send(res, status, { error: status >= 500 && !(e instanceof HttpError) ? 'Something went wrong on the server. Please try again.' : e.message });
  }
});

server.requestTimeout = 120000;
server.headersTimeout = 30000;

/* ------------------------------------------------------------------ */
/* Startup                                                             */
/* ------------------------------------------------------------------ */

async function main() {
  const major = parseInt(process.versions.node.split('.')[0], 10);
  if (major < 20) { console.error('MOM Bot needs Node.js 20 or newer.'); process.exit(1); }
  loadDb();
  await bootstrapAdmin();

  // Meetings left "generating" by a crash/restart: resume them.
  for (const m of db.meetings) {
    if (m.stage === 'generating' && combinedSegments(m).length) withLock(m.id, () => generateMom(m)).catch(() => {});
    if (m.stage === 'dispatching' && !m.botId) dispatchBot(m).catch(() => {});
  }

  log('info', `Recall.ai region: ${CFG.recallRegion}${CFG.recallApiBase ? ` (override ${RECALL_BASE})` : ''}`);
  if (!RECALL_REGIONS.includes(CFG.recallRegion) && !CFG.recallApiBase) log('warn', `RECALL_REGION should be one of ${RECALL_REGIONS.join(', ')}`);
  const missing = [];
  if (!CFG.recallApiKey) missing.push('RECALL_API_KEY');
  if (!CFG.recallVerificationSecret && !CFG.recallSvixSecret) missing.push('RECALL_VERIFICATION_SECRET');
  if (!CFG.publicBaseUrl) missing.push('PUBLIC_BASE_URL');
  if (!llmConfigured()) missing.push(CFG.llmProvider === 'openai' ? 'OPENAI_API_KEY / OPENAI_MODEL' : 'ANTHROPIC_API_KEY');
  if (missing.length) log('warn', `Not configured yet: ${missing.join(', ')} (see README.md)`);
  log('info', `AI provider: ${CFG.llmProvider} (${CFG.llmProvider === 'openai' ? CFG.openaiModel || 'model not set' : CFG.anthropicModel})`);
  log('info', `Data file: ${CFG.dataFile}`);

  setTimeout(() => backgroundJobs({ recallSweep: true }), 5000).unref();
  setInterval(() => backgroundJobs(), CFG.icsPollMinutes * 60000).unref();
  setInterval(() => backgroundJobs({ recallSweep: true }), 60 * 60000).unref();

  server.listen(CFG.port, '0.0.0.0', () => {
    log('info', `MOM Bot running on port ${CFG.port}`);
    if (CFG.publicBaseUrl) log('info', `Recall webhook URL: ${CFG.publicBaseUrl}/webhooks/recall`);
  });
}

async function shutdown(sig) {
  log('info', `${sig} received, saving and shutting down`);
  server.close();
  await save();
  await saveChain;
  process.exit(0);
}
if (require.main === module) {
  process.on('SIGTERM', () => shutdown('SIGTERM'));
  process.on('SIGINT', () => shutdown('SIGINT'));
  process.on('unhandledRejection', (e) => log('error', 'Unhandled rejection', e && (e.stack || e.message)));
}

if (require.main === module) main().catch((e) => { console.error(e); process.exit(1); });
else module.exports = { icsToMeetings, parseMeetingUrl, findMeetingLink, zonedToUtc, resolveTz, parseManualTranscript, normalizeTranscript };
