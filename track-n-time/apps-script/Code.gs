/**
 * Track’n’Time – Butlin's Skegness (CBRE engineering)
 * Google Apps Script backend. The Google Sheet this script is attached to is the database.
 *
 * The timesheet is the heart of the app: raising a job books it onto the timesheet as a
 * planned block for the time you think it needs (your quote). Actual time is logged against it.
 *
 * Job stages:  Scheduled → In progress → To Admin → Complete
 *   - Engineers move their own jobs to In progress and send them To Admin with the report.
 *   - Only an admin can mark a job Complete, once the paperwork has been checked.
 *
 * Roles
 *   admin    – raise jobs for anyone, assign, schedule anyone, mark jobs Complete, manage people.
 *   engineer – raise jobs (assigned to themselves), work their own jobs, log their own time.
 *
 * Locations come from the Heater PPM checklist (villages, units and venues) – see pullLocations_.
 *
 * Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
 */

const APP_NAME = 'Track’n’Time';
const SESSION_DAYS = 60;

/** Heater PPM checklist – source of village units and venues. */
const HEATER_CHECKLIST_ID = '1VQO_y3zsBxahBsAcHDPyi0HiL2X0tuOWoMypssGxPPw';
const VILLAGES = ['Pacific Wharf', 'Starfish Quay', 'Starfish Quay (FG)', 'Ocean Point', 'Lagoon Bay',
  'Lagoon Bay (SV)', 'The Keys', 'Royal Arthur'];

const SHEETS = {
  Users: ['id', 'name', 'pinHash', 'role', 'active', 'createdAt'],
  Jobs: ['id', 'ref', 'woRef', 'title', 'area', 'location', 'type', 'priority', 'status', 'description',
    'date', 'start', 'estHours',
    'createdBy', 'createdAt', 'assignedTo', 'assignedBy', 'updatedAt',
    'workDone', 'findings', 'furtherWork', 'completedBy', 'completedAt',
    'signedOffBy', 'signedOffAt', 'signOffNote', 'followOnRef', 'parentRef'],
  Time: ['id', 'userId', 'date', 'start', 'end', 'hours', 'status', 'title', 'type', 'area', 'location',
    'jobId', 'jobRef', 'note', 'source', 'createdBy', 'createdAt'],
  Materials: ['id', 'jobId', 'jobRef', 'item', 'qty', 'unit', 'source', 'userId', 'at'],
  Notes: ['id', 'jobId', 'userId', 'text', 'kind', 'at'],
  Locations: ['group', 'area', 'location'],
  Areas: ['name'],
  Presets: ['title', 'type']
};

/** Extra places not in the heater checklist. Edit the Areas tab to change. */
const DEFAULT_AREAS = ['Site-wide', 'Plant Rooms', 'External Lighting & Car Parks', 'Splash Waterworld',
  'Skyline Pavilion', 'Funfair & Rides', 'Staff Village', 'Workshop', 'Other'];

const DEFAULT_PRESETS = [
  ['PPM – Emergency lights', 'PPM'], ['PPM – Fire alarm', 'PPM'], ['PPM – Wall heaters', 'PPM'],
  ['PPM – Lighting', 'PPM'], ['PPM – DB checks', 'PPM'], ['PPM – PAT testing', 'PPM'],
  ['Reactive callout', 'Reactive'], ['Back job', 'Back job'], ['Project work', 'Project'],
  ['Travel', 'Travel'], ['Training', 'Training'], ['Admin / paperwork', 'Admin']];

const STAGES = ['Scheduled', 'In progress', 'To Admin', 'Complete'];

/* ======================= Web app entry ======================= */

function doGet() {
  ensure_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Run once from the editor: creates the tabs and pulls locations from the heater checklist. */
function setup() {
  ensure_();
  let msg = 'Track’n’Time tabs are ready.';
  try { msg += ' ' + writeLocations_() + ' locations pulled from the heater checklist.'; }
  catch (e) { msg += ' Could not read the heater checklist: ' + e.message; }
  SpreadsheetApp.getActive().toast(msg);
  Logger.log(msg);
}

/* ======================= Sheet helpers ======================= */

let ensured_ = false;

function ss_() { return SpreadsheetApp.getActive(); }

function ensure_() {
  if (ensured_) return;
  const ss = ss_();
  Object.keys(SHEETS).forEach(function (name) {
    const want = SHEETS[name];
    let s = ss.getSheetByName(name);
    if (!s) {
      s = ss.insertSheet(name);
      s.getRange(1, 1, s.getMaxRows(), Math.max(want.length, s.getMaxColumns())).setNumberFormat('@');
      s.getRange(1, 1, 1, want.length).setValues([want]).setFontWeight('bold');
      s.setFrozenRows(1);
      if (name === 'Areas') s.getRange(2, 1, DEFAULT_AREAS.length, 1).setValues(DEFAULT_AREAS.map(a => [a]));
      if (name === 'Presets') s.getRange(2, 1, DEFAULT_PRESETS.length, 2).setValues(DEFAULT_PRESETS);
    } else {
      const lastCol = Math.max(1, s.getLastColumn());
      const have = s.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
      const missing = want.filter(h => have.indexOf(h) < 0);
      if (missing.length) {
        const startCol = (have.filter(String).length || 0) + 1;
        if (startCol + missing.length - 1 > s.getMaxColumns()) s.insertColumnsAfter(s.getMaxColumns(), missing.length);
        s.getRange(1, startCol, 1, missing.length).setValues([missing]).setFontWeight('bold');
      }
    }
  });
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  ensured_ = true;
}

function pad2_(n) { return (n < 10 ? '0' : '') + n; }

function norm_(h, v) {
  if (v instanceof Date) {
    const tz = Session.getScriptTimeZone();
    if (h === 'start' || h === 'end') return Utilities.formatDate(v, tz, 'HH:mm');
    if (h === 'date') return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    return v.toISOString();
  }
  if (h === 'qty' || h === 'hours' || h === 'estHours') return v === '' ? 0 : Number(v);
  if (h === 'active') return !(v === false || String(v).toLowerCase() === 'false');
  return v === null || v === undefined ? '' : v;
}

function read_(name) {
  ensure_();
  const s = ss_().getSheetByName(name);
  const vals = s.getDataRange().getValues();
  if (vals.length < 2) return [];
  const heads = vals[0].map(String);
  const out = [];
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    if (row.every(c => c === '' || c === null)) continue;
    const o = { _row: r + 1 };
    heads.forEach((h, i) => { if (h) o[h] = norm_(h, row[i]); });
    out.push(o);
  }
  return out;
}

function heads_(s) { return s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(String); }

function toRow_(heads, obj) {
  return heads.map(h => {
    const v = obj[h];
    if (v === undefined || v === null) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return String(v);
  });
}

function append_(name, obj) {
  ensure_();
  const s = ss_().getSheetByName(name);
  const heads = heads_(s);
  const r = s.getLastRow() + 1;
  if (r > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), 50);
  s.getRange(r, 1, 1, heads.length).setNumberFormat('@').setValues([toRow_(heads, obj)]);
  return obj;
}

function find_(name, id) { return read_(name).filter(o => String(o.id) === String(id))[0] || null; }

function update_(name, id, patch) {
  const cur = find_(name, id);
  if (!cur) throw new Error('That record no longer exists. Pull down to refresh.');
  const s = ss_().getSheetByName(name);
  const heads = heads_(s);
  const next = Object.assign({}, cur, patch);
  s.getRange(cur._row, 1, 1, heads.length).setNumberFormat('@').setValues([toRow_(heads, next)]);
  return next;
}

function remove_(name, id) {
  const cur = find_(name, id);
  if (cur) ss_().getSheetByName(name).deleteRow(cur._row);
}

function strip_(o) { if (!o) return o; const c = Object.assign({}, o); delete c._row; delete c.pinHash; return c; }
function uid_(prefix) { return prefix + Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function now_() { return new Date().toISOString(); }
function clean_(v, max) { return String(v === undefined || v === null ? '' : v).trim().slice(0, max || 500); }

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function makeRef_() {
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let r = '';
  for (let i = 0; i < 3; i++) r += c.charAt(Math.floor(Math.random() * c.length));
  return 'BS-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd') + '-' + r;
}

function dayStr_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

/* ======================= Locations (from the heater checklist) ======================= */

function tidy_(s) {
  s = String(s || '').replace(/\s+/g, ' ').trim();
  if (s && s === s.toUpperCase() && /[A-Z]{3}/.test(s)) s = s.toLowerCase().replace(/\b[a-z]/g, c => c.toUpperCase());
  return s;
}

/** Reads every village unit and venue floor/room from the Heater PPM checklist. */
function pullLocations_() {
  const src = SpreadsheetApp.openById(HEATER_CHECKLIST_ID);
  const rows = [], seen = {};
  const add = (g, a, l) => {
    a = tidy_(a); l = tidy_(l);
    if (!a || !l) return;
    const k = a + '|' + l;
    if (seen[k]) return;
    seen[k] = 1; rows.push([g, a, l]);
  };
  VILLAGES.forEach(v => {
    const s = src.getSheetByName(v);
    if (!s || s.getLastRow() < 2) return;
    s.getRange(1, 1, s.getLastRow(), 1).getValues().forEach(r => {
      const u = String(r[0]).trim();
      if (/\d/.test(u) && !/unit number|total|schedule|apartments/i.test(u)) add('Village', v, u);
    });
  });
  const ven = src.getSheets().filter(s => /^venues/i.test(s.getName()))[0];
  if (ven) {
    const vals = ven.getDataRange().getValues();
    let hdr = -1;
    for (let i = 0; i < vals.length; i++) if (/^building/i.test(String(vals[i][0]).trim())) { hdr = i; break; }
    for (let i = hdr + 1; i < vals.length; i++) {
      const r = vals[i], name = String(r[0]).trim();
      if (!name) continue;
      add('Venue', name, [tidy_(r[1]), tidy_(r[2])].filter(Boolean).join(' – ') || 'Throughout');
    }
  }
  return rows;
}

function writeLocations_() {
  const rows = pullLocations_();
  if (!rows.length) throw new Error('No locations found in the heater checklist.');
  ensure_();
  const s = ss_().getSheetByName('Locations');
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 3).clearContent();
  if (rows.length + 1 > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), rows.length + 1 - s.getMaxRows());
  s.getRange(2, 1, rows.length, 3).setNumberFormat('@').setValues(rows);
  return rows.length;
}

/** Admin button: refresh locations from the heater checklist. */
function syncLocations(token) {
  requireAdmin_(auth_(token));
  return withLock_(() => ({ count: writeLocations_(), places: places_() }));
}

/** Areas grouped for the app: villages (with units), venues (with floors/rooms), then extra areas. */
function places_() {
  const map = {}, order = [];
  read_('Locations').forEach(l => {
    const key = l.area;
    if (!map[key]) { map[key] = { name: l.area, group: l.group || 'Other', locs: [] }; order.push(key); }
    map[key].locs.push(l.location);
  });
  read_('Areas').forEach(a => {
    if (a.name && !map[a.name]) { map[a.name] = { name: a.name, group: 'Other', locs: [] }; order.push(a.name); }
  });
  const rank = { Village: 0, Venue: 1, Other: 2 };
  return order.map(k => map[k]).sort((a, b) => (rank[a.group] - rank[b.group]) ||
    (a.group === 'Village' ? VILLAGES.indexOf(a.name) - VILLAGES.indexOf(b.name) : a.name.localeCompare(b.name)));
}

/* ======================= Auth ======================= */

function hash_(userId, pin) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, userId + ':' + pin + ':site-jobs'));
}

function checkPin_(pin) {
  pin = String(pin || '');
  if (!/^\d{4,6}$/.test(pin)) throw new Error('PIN must be 4 to 6 digits.');
  return pin;
}

function props_() { return PropertiesService.getScriptProperties(); }

function newSession_(userId) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  props_().setProperty('tok_' + token, JSON.stringify({ u: userId, exp: Date.now() + SESSION_DAYS * 864e5 }));
  return token;
}

function auth_(token) {
  if (!token) throw new Error('SESSION');
  const raw = props_().getProperty('tok_' + token);
  if (!raw) throw new Error('SESSION');
  const s = JSON.parse(raw);
  if (s.exp < Date.now()) { props_().deleteProperty('tok_' + token); throw new Error('SESSION'); }
  const u = find_('Users', s.u);
  if (!u || !u.active) throw new Error('SESSION');
  return u;
}

function isAdmin_(u) { return u.role === 'admin'; }
function requireAdmin_(u) { if (!isAdmin_(u)) throw new Error('Only an admin can do that.'); }

function pruneSessions_() {
  const all = props_().getProperties();
  Object.keys(all).forEach(k => {
    if (k.indexOf('tok_') !== 0) return;
    try { if (JSON.parse(all[k]).exp < Date.now()) props_().deleteProperty(k); } catch (e) { props_().deleteProperty(k); }
  });
}

function getStart(token) {
  ensure_();
  if (token) { try { auth_(token); return { loggedIn: true }; } catch (e) { /* show login */ } }
  const users = read_('Users').filter(u => u.active).map(u => ({ id: u.id, name: u.name }))
    .sort((a, b) => a.name.localeCompare(b.name));
  return { loggedIn: false, needsSetup: users.length === 0, users: users };
}

function createFirstAdmin(name, pin) {
  return withLock_(() => {
    if (read_('Users').length) throw new Error('The app is already set up. Ask an admin to add you.');
    name = clean_(name, 60); pin = checkPin_(pin);
    if (!name) throw new Error('Enter your name.');
    const id = uid_('u_');
    append_('Users', { id: id, name: name, pinHash: hash_(id, pin), role: 'admin', active: true, createdAt: now_() });
    return { token: newSession_(id) };
  });
}

function login(userId, pin) {
  const u = find_('Users', userId);
  if (!u || !u.active || u.pinHash !== hash_(u.id, String(pin))) {
    Utilities.sleep(800);
    throw new Error('Wrong PIN. Try again or ask an admin to reset it.');
  }
  if (Math.random() < 0.1) pruneSessions_();
  return { token: newSession_(u.id) };
}

function logout(token) { if (token) props_().deleteProperty('tok_' + token); return true; }

function changePin(token, oldPin, newPin) {
  const u = auth_(token);
  if (u.pinHash !== hash_(u.id, String(oldPin))) throw new Error('Current PIN is wrong.');
  newPin = checkPin_(newPin);
  withLock_(() => update_('Users', u.id, { pinHash: hash_(u.id, newPin) }));
  return true;
}

/* ======================= Data for the app ======================= */

function boot(token) {
  const me = auth_(token);
  const admin = isAdmin_(me);
  const users = read_('Users').map(u => ({ id: u.id, name: u.name, role: u.role, active: u.active }));
  const cutoff = new Date(Date.now() - 60 * 864e5).toISOString();
  const jobs = read_('Jobs').filter(j => j.status !== 'Complete' || String(j.signedOffAt) > cutoff).map(strip_);
  const from = dayStr_(new Date(Date.now() - 42 * 864e5));
  const to = dayStr_(new Date(Date.now() + 42 * 864e5));
  const time = read_('Time').filter(t => t.date >= from && t.date <= to && (admin || t.userId === me.id)).map(strip_);
  return {
    me: { id: me.id, name: me.name, role: me.role },
    users: users, jobs: jobs, time: time, timeFrom: from, timeTo: to,
    places: places_(),
    presets: read_('Presets').filter(p => p.title).map(p => ({ title: p.title, type: p.type || 'PPM' })),
    stages: STAGES, today: dayStr_(new Date())
  };
}

function getTime(token, fromDate, toDate, userId) {
  const me = auth_(token);
  const who = isAdmin_(me) ? userId : me.id;
  return read_('Time').filter(t => t.date >= fromDate && t.date <= toDate && (!who || t.userId === who)).map(strip_);
}

function getJob(token, jobId) {
  const me = auth_(token);
  const job = find_('Jobs', jobId);
  if (!job) throw new Error('That job no longer exists.');
  return {
    job: strip_(job),
    time: read_('Time').filter(t => t.jobId === jobId).map(strip_),
    materials: read_('Materials').filter(m => m.jobId === jobId).map(strip_),
    notes: read_('Notes').filter(n => n.jobId === jobId).map(strip_).sort((a, b) => String(b.at).localeCompare(String(a.at))),
    canEdit: canEditJob_(me, job), isAdmin: isAdmin_(me)
  };
}

function canEditJob_(u, job) { return isAdmin_(u) || (!!job.assignedTo && job.assignedTo === u.id); }

function requireJobEdit_(u, job) {
  if (canEditJob_(u, job)) return;
  throw new Error(job.assignedTo ? 'This job is assigned to someone else. Ask an admin to reassign it.' : 'This job isn\'t assigned yet. Ask an admin to assign it to you.');
}

function note_(jobId, userId, text, kind) {
  append_('Notes', { id: uid_('n_'), jobId: jobId, userId: userId, text: text, kind: kind || 'note', at: now_() });
}

function userName_(id) { const u = find_('Users', id); return u ? u.name : 'someone'; }

/* ======================= Time maths ======================= */

function toMin_(hhmm) {
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(hhmm || ''));
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
}

function fromMin_(m) { m = ((m % 1440) + 1440) % 1440; return pad2_(Math.floor(m / 60)) + ':' + pad2_(m % 60); }

function hours_(start, end) {
  const a = toMin_(start), b = toMin_(end);
  if (a === null || b === null) return 0;
  let d = b - a; if (d < 0) d += 1440;
  return Math.round(d / 60 * 100) / 100;
}

function nowHHMM_(roundTo) {
  const tz = Session.getScriptTimeZone(), d = new Date();
  let m = Number(Utilities.formatDate(d, tz, 'H')) * 60 + Number(Utilities.formatDate(d, tz, 'm'));
  m = Math.round(m / roundTo) * roundTo; if (m >= 1440) m = 1435;
  return fromMin_(m);
}

/* ======================= Jobs ======================= */

/**
 * Keeps the job's planned block on the timesheet in line with the job:
 * one Planned entry for the assignee on the job's date, start and quoted hours.
 * allowCreate: only when the job is raised, rescheduled or reassigned – otherwise a block that was
 * already worked (confirmed or clocked) is not booked again.
 */
function syncPlan_(job, allowCreate) {
  const plans = read_('Time').filter(t => t.jobId === job.id && t.source === 'job' && t.status === 'Planned');
  const want = job.assignedTo && job.date && toMin_(job.start) !== null && Number(job.estHours) > 0 &&
    (job.status === 'Scheduled' || job.status === 'In progress');
  if (!want) { plans.forEach(p => remove_('Time', p.id)); return; }
  const rec = {
    userId: job.assignedTo, date: job.date, start: job.start,
    end: fromMin_(toMin_(job.start) + Math.round(Number(job.estHours) * 60)), hours: Number(job.estHours),
    status: 'Planned', title: job.title, type: job.type, area: job.area, location: job.location,
    jobId: job.id, jobRef: job.ref, source: 'job'
  };
  if (plans.length) {
    update_('Time', plans[0].id, rec);
    plans.slice(1).forEach(p => remove_('Time', p.id));
  } else if (allowCreate) {
    rec.id = uid_('t_'); rec.createdBy = job.assignedBy || job.createdBy; rec.createdAt = now_();
    append_('Time', rec);
  }
}

function jobFields_(j) {
  const est = Number(j.estHours);
  const f = {
    title: clean_(j.title, 140), area: clean_(j.area, 80), location: clean_(j.location, 140),
    type: clean_(j.type, 40) || 'Reactive', priority: ['P1', 'P2', 'P3', 'P4'].indexOf(j.priority) >= 0 ? j.priority : 'P3',
    description: clean_(j.description, 4000), woRef: clean_(j.woRef, 40),
    date: /^\d{4}-\d{2}-\d{2}$/.test(j.date) ? j.date : '', start: toMin_(j.start) !== null ? j.start : '',
    estHours: est > 0 ? Math.min(Math.round(est * 4) / 4, 24) : ''
  };
  if (!f.title) throw new Error('Give the job a title.');
  if (!f.area) throw new Error('Choose an area.');
  return f;
}

/** Raise or edit a job. Raising it books the quoted time onto the timesheet. */
function saveJob(token, j) {
  const me = auth_(token);
  return withLock_(() => {
    const fields = jobFields_(j);
    if (j.id) {
      const cur = find_('Jobs', j.id);
      if (!cur) throw new Error('That job no longer exists.');
      requireJobEdit_(me, cur);
      fields.updatedAt = now_();
      const moved = fields.date !== cur.date || fields.start !== cur.start || Number(fields.estHours) !== Number(cur.estHours);
      const next = update_('Jobs', j.id, fields);
      syncPlan_(next, moved);
      return getJob(token, j.id);
    }
    if (!fields.date || !fields.start || !fields.estHours) throw new Error('Pick a day, start time and how long it will take.');
    const assignee = isAdmin_(me) ? clean_(j.assignedTo, 40) : me.id;
    const job = Object.assign({
      id: uid_('j_'), ref: makeRef_(), status: 'Scheduled', createdBy: me.id, createdAt: now_(), updatedAt: now_(),
      assignedTo: assignee, assignedBy: assignee ? me.id : '', parentRef: clean_(j.parentRef, 40)
    }, fields);
    append_('Jobs', job);
    syncPlan_(job, true);
    note_(job.id, me.id, 'Job raised and scheduled' + (assignee && assignee !== me.id ? ' for ' + userName_(assignee) : ''), 'sys');
    return getJob(token, job.id);
  });
}

function assignJob(token, jobId, userId) {
  const me = auth_(token); requireAdmin_(me);
  return withLock_(() => {
    if (!find_('Jobs', jobId)) throw new Error('That job no longer exists.');
    userId = clean_(userId, 40);
    const next = update_('Jobs', jobId, { assignedTo: userId, assignedBy: me.id, updatedAt: now_() });
    syncPlan_(next, true);
    note_(jobId, me.id, userId ? 'Assigned to ' + userName_(userId) : 'Unassigned', 'sys');
    return getJob(token, jobId);
  });
}

/** Engineers: Scheduled ↔ In progress on their own jobs. Admins: any stage except Complete (use signOff). */
function setStatus(token, jobId, status) {
  const me = auth_(token);
  return withLock_(() => {
    const cur = find_('Jobs', jobId); if (!cur) throw new Error('That job no longer exists.');
    requireJobEdit_(me, cur);
    const admin = isAdmin_(me);
    const allowed = admin ? ['Scheduled', 'In progress', 'To Admin'] : ['Scheduled', 'In progress'];
    if (allowed.indexOf(status) < 0) throw new Error(status === 'To Admin' ? 'Use "Send to Admin" with the job report.' : 'Only an admin can mark a job Complete.');
    if ((cur.status === 'To Admin' || cur.status === 'Complete') && !admin) throw new Error('This job is with the admin now. Ask them to send it back if it needs more work.');
    const next = update_('Jobs', jobId, { status: status, updatedAt: now_() });
    syncPlan_(next, false);
    note_(jobId, me.id, 'Moved to ' + status, 'sys');
    return getJob(token, jobId);
  });
}

function setWo(token, jobId, wo) {
  const me = auth_(token);
  return withLock_(() => {
    const cur = find_('Jobs', jobId); if (!cur) throw new Error('That job no longer exists.');
    if (!isAdmin_(me) && cur.assignedTo !== me.id && cur.createdBy !== me.id) throw new Error('Only the assigned engineer or an admin can change this.');
    wo = clean_(wo, 40);
    update_('Jobs', jobId, { woRef: wo, updatedAt: now_() });
    note_(jobId, me.id, wo ? 'IFMHub WO set to ' + wo : 'IFMHub WO removed', 'sys');
    return getJob(token, jobId);
  });
}

function saveReport(token, jobId, r) {
  const me = auth_(token);
  return withLock_(() => {
    const cur = find_('Jobs', jobId); if (!cur) throw new Error('That job no longer exists.');
    requireJobEdit_(me, cur);
    update_('Jobs', jobId, { workDone: clean_(r.workDone, 4000), findings: clean_(r.findings, 4000), furtherWork: clean_(r.furtherWork, 4000), updatedAt: now_() });
    return getJob(token, jobId);
  });
}

/** Engineer finishes the work and sends it to the admin (like sending it for invoicing). */
function sendToAdmin(token, jobId, r) {
  const me = auth_(token);
  return withLock_(() => {
    const cur = find_('Jobs', jobId); if (!cur) throw new Error('That job no longer exists.');
    requireJobEdit_(me, cur);
    if (cur.status === 'Complete') throw new Error('This job is already complete.');
    const workDone = clean_(r.workDone, 4000);
    if (!workDone) throw new Error('Add the work carried out before sending to Admin.');
    const patch = { workDone: workDone, findings: clean_(r.findings, 4000), furtherWork: clean_(r.furtherWork, 4000),
      status: 'To Admin', completedBy: me.id, completedAt: now_(), updatedAt: now_(), signedOffBy: '', signedOffAt: '', signOffNote: '' };
    read_('Time').filter(t => t.jobId === jobId && t.status === 'Running').forEach(t => stopEntry_(t));
    if (r.followOn && patch.furtherWork) {
      const ref = makeRef_();
      const nj = { id: uid_('j_'), ref: ref, woRef: '', title: 'Back job: ' + cur.title, area: cur.area, location: cur.location,
        type: 'Back job', priority: cur.priority === 'P1' ? 'P2' : cur.priority, status: 'Scheduled', description: patch.furtherWork,
        date: '', start: '', estHours: '', createdBy: me.id, createdAt: now_(), updatedAt: now_(),
        assignedTo: isAdmin_(me) ? '' : me.id, assignedBy: isAdmin_(me) ? '' : me.id, parentRef: cur.ref };
      append_('Jobs', nj);
      note_(nj.id, me.id, 'Raised as a back job from ' + cur.ref + ' – needs a day and time', 'sys');
      patch.followOnRef = ref;
    }
    const next = update_('Jobs', jobId, patch);
    syncPlan_(next, false); // clears any planned block left over
    note_(jobId, me.id, 'Sent to Admin' + (patch.followOnRef ? ' (back job ' + patch.followOnRef + ' raised)' : ''), 'sys');
    return getJob(token, jobId);
  });
}

/** Admin only: mark Complete once paperwork is checked, or send it back to the engineer. */
function signOff(token, jobId, approve, comment, paperworkChecked) {
  const me = auth_(token); requireAdmin_(me);
  return withLock_(() => {
    const cur = find_('Jobs', jobId); if (!cur) throw new Error('That job no longer exists.');
    comment = clean_(comment, 1000);
    if (approve) {
      if (cur.status !== 'To Admin') throw new Error('Only jobs sent To Admin can be marked Complete.');
      if (!paperworkChecked) throw new Error('Tick that the paperwork has been checked first.');
      update_('Jobs', jobId, { status: 'Complete', signedOffBy: me.id, signedOffAt: now_(), signOffNote: comment, updatedAt: now_() });
      note_(jobId, me.id, 'Marked Complete – paperwork checked' + (comment ? ': ' + comment : ''), 'sys');
    } else {
      if (!comment) throw new Error('Say what is missing before sending it back.');
      update_('Jobs', jobId, { status: 'In progress', signedOffBy: '', signedOffAt: '', signOffNote: comment, updatedAt: now_() });
      note_(jobId, me.id, 'Sent back: ' + comment, 'sys');
    }
    return getJob(token, jobId);
  });
}

function addMaterial(token, jobId, m) {
  const me = auth_(token);
  return withLock_(() => {
    const cur = find_('Jobs', jobId); if (!cur) throw new Error('That job no longer exists.');
    requireJobEdit_(me, cur);
    const item = clean_(m.item, 160); if (!item) throw new Error('Enter the item.');
    append_('Materials', { id: uid_('m_'), jobId: jobId, jobRef: cur.ref, item: item, qty: Number(m.qty) || 0,
      unit: clean_(m.unit, 20) || 'each', source: clean_(m.source, 40), userId: me.id, at: now_() });
    return getJob(token, jobId);
  });
}

function deleteMaterial(token, id) {
  const me = auth_(token);
  return withLock_(() => {
    const m = find_('Materials', id); if (!m) return true;
    if (!isAdmin_(me) && m.userId !== me.id) throw new Error('You can only remove materials you added.');
    remove_('Materials', id);
    return getJob(token, m.jobId);
  });
}

function addNote(token, jobId, text) {
  const me = auth_(token);
  text = clean_(text, 2000); if (!text) throw new Error('Write a note first.');
  withLock_(() => note_(jobId, me.id, text, 'note'));
  return getJob(token, jobId);
}

/* ======================= Timesheet ======================= */

function startJob_(job) {
  if (job && job.status === 'Scheduled') update_('Jobs', job.id, { status: 'In progress', updatedAt: now_() });
}

/**
 * Create or update a timesheet entry.
 * status: 'Done' (worked), 'Planned' (booked ahead), 'Running' (clocked on).
 */
function saveTime(token, e) {
  const me = auth_(token);
  const admin = isAdmin_(me);
  return withLock_(() => {
    const userId = admin && e.userId ? clean_(e.userId, 40) : me.id;
    const status = ['Done', 'Planned', 'Running'].indexOf(e.status) >= 0 ? e.status : 'Done';
    const date = /^\d{4}-\d{2}-\d{2}$/.test(e.date) ? e.date : dayStr_(new Date());
    const start = toMin_(e.start) !== null ? e.start : '08:00';
    const end = status === 'Running' ? '' : (toMin_(e.end) !== null ? e.end : start);
    let job = null;
    if (e.jobId) { job = find_('Jobs', e.jobId); if (!job) throw new Error('That job no longer exists.'); }
    if (job && job.status === 'Complete' && !admin) throw new Error('That job is complete. Ask an admin if time needs adding.');
    const rec = {
      userId: userId, date: date, start: start, end: end, hours: status === 'Running' ? 0 : hours_(start, end), status: status,
      title: clean_(e.title, 140) || (job ? job.title : 'Work'), type: clean_(e.type, 40) || (job ? job.type : ''),
      area: clean_(e.area, 80) || (job ? job.area : ''), location: clean_(e.location, 140) || (job ? job.location : ''),
      jobId: job ? job.id : '', jobRef: job ? job.ref : '', note: clean_(e.note, 500)
    };
    if (status !== 'Running' && rec.hours > 16) throw new Error('That is more than 16 hours. Check the start and finish times.');
    if (e.id) {
      const cur = find_('Time', e.id); if (!cur) throw new Error('That entry no longer exists.');
      if (!admin && cur.userId !== me.id) throw new Error('You can only change your own time.');
      // A worked/confirmed block stops being the job's planned quote block.
      if (cur.source === 'job' && status !== 'Planned') rec.source = 'worked';
      update_('Time', e.id, rec);
      // Moving the planned block moves the job's booked slot with it.
      if (cur.source === 'job' && status === 'Planned' && job) {
        update_('Jobs', job.id, { date: date, start: start, estHours: rec.hours, updatedAt: now_() });
      }
    } else {
      if (status === 'Running') read_('Time').filter(t => t.userId === userId && t.status === 'Running').forEach(t => stopEntry_(t));
      rec.id = uid_('t_'); rec.createdBy = me.id; rec.createdAt = now_(); rec.source = 'manual';
      append_('Time', rec);
    }
    if (job && status !== 'Planned') startJob_(job);
    return true;
  });
}

function stopEntry_(t) {
  const end = nowHHMM_(5);
  update_('Time', t.id, { end: end, hours: hours_(t.start, end), status: 'Done' });
}

/** Clock on now. On a job with a planned block today, that block becomes the live clock. */
function clockOn(token, e) {
  const me = auth_(token);
  e = e || {};
  const today = dayStr_(new Date());
  if (e.jobId) {
    const plan = withLock_(() => read_('Time').filter(t => t.jobId === e.jobId && t.userId === me.id && t.status === 'Planned' && t.date === today)[0]);
    if (plan) {
      return withLock_(() => {
        read_('Time').filter(t => t.userId === me.id && t.status === 'Running').forEach(t => stopEntry_(t));
        update_('Time', plan.id, { status: 'Running', start: nowHHMM_(5), end: '', hours: 0, source: 'worked' });
        startJob_(find_('Jobs', e.jobId));
        return true;
      });
    }
  }
  e.status = 'Running'; e.start = nowHHMM_(5); e.date = today; e.userId = me.id;
  return saveTime(token, e);
}

function clockOff(token, id) {
  const me = auth_(token);
  return withLock_(() => {
    const t = find_('Time', id); if (!t) throw new Error('That entry no longer exists.');
    if (!isAdmin_(me) && t.userId !== me.id) throw new Error('You can only stop your own clock.');
    stopEntry_(t);
    return true;
  });
}

function deleteTime(token, id) {
  const me = auth_(token);
  return withLock_(() => {
    const t = find_('Time', id); if (!t) return true;
    if (!isAdmin_(me) && t.userId !== me.id) throw new Error('You can only remove your own time.');
    if (t.source === 'job' && t.status === 'Planned' && t.jobId) {
      update_('Jobs', t.jobId, { date: '', start: '', updatedAt: now_() }); // job becomes unscheduled
    }
    remove_('Time', id);
    return true;
  });
}

/* ======================= Admin: people ======================= */

function saveUser(token, u) {
  const me = auth_(token); requireAdmin_(me);
  return withLock_(() => {
    const name = clean_(u.name, 60); if (!name) throw new Error('Enter a name.');
    const role = u.role === 'admin' ? 'admin' : 'engineer';
    if (u.id) {
      const cur = find_('Users', u.id); if (!cur) throw new Error('That person no longer exists.');
      const active = u.active !== false;
      if (cur.id === me.id && (role !== 'admin' || !active)) throw new Error("You can't remove your own admin access. Ask another admin.");
      const patch = { name: name, role: role, active: active };
      if (u.pin) patch.pinHash = hash_(cur.id, checkPin_(u.pin));
      update_('Users', u.id, patch);
    } else {
      const pin = checkPin_(u.pin);
      if (read_('Users').some(x => x.active && x.name.toLowerCase() === name.toLowerCase())) throw new Error('Someone already has that name. Add a surname or initial.');
      const id = uid_('u_');
      append_('Users', { id: id, name: name, pinHash: hash_(id, pin), role: role, active: true, createdAt: now_() });
    }
    return read_('Users').map(x => ({ id: x.id, name: x.name, role: x.role, active: x.active }));
  });
}
