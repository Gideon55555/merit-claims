/**
 * Merit Properties – Customer Claim System
 * Google Apps Script web app, bound to the claims spreadsheet.
 *
 * Rules:
 *  - A phone number registered in ANY category (Expecting / Office Visit / Site Visit)
 *    cannot be registered again by anyone for 30 days. It is free again on day 31.
 *  - Numbers are compared on their first 8 digits (e.g. 0911 51 83), because most
 *    records keep the last 2 digits hidden.
 *  - Roles: admin (everything), entry (register + check), viewer (records + check).
 */

const CFG = {
  TZ: 'Africa/Addis_Ababa',   // "today" is calculated in this time zone
  LOCK_DAYS: 30,              // locked for 30 days, free from day 31
  KEY_DIGITS: 8,              // digits compared when matching phone numbers
  SESSION_SECONDS: 6 * 60 * 60,
  MAX_LOGIN_FAILS: 5,
  AGENTS_SHEET: 'Agents',
  USERS_SHEET: 'Users',
  MAX_RECORDS: 500,
};

const CATEGORIES = [
  { key: 'expecting', label: 'Expecting', match: /expect/i },
  { key: 'office', label: 'Office Visit', match: /office/i },
  { key: 'site', label: 'Site Visit', match: /site/i },
];

const AGENT_TYPES = ['Sales', 'Elite Sales', 'Supervisor', 'Freelancer', 'Agent'];
const ROLES = ['admin', 'entry', 'viewer'];
const EXTRA_HEADERS = ['Agent Type', 'Supervisor', 'Entered By', 'Entered At'];
const AGENT_HEADERS = ['ID', 'Type', 'Name', 'Supervisor', 'Phone', 'Active'];
const USER_HEADERS = ['Username', 'Full Name', 'Role', 'Active', 'Salt', 'Hash', 'Must Change Password', 'Created', 'Last Login'];

/* ───────────────────────── Web app entry ───────────────────────── */

/** Public number-check page for sales (no login): set enabled:false to turn it off, showAgent:false to hide who claimed it. */
const PUBLIC_CHECK = { enabled: true, showAgent: true, maxPerMinute: 120 };

function doGet(e) {
  const isCheck = e && e.parameter && String(e.parameter.page || '').toLowerCase() === 'check';
  return HtmlService.createTemplateFromFile(isCheck ? 'Check' : 'Index')
    .evaluate()
    .setTitle(isCheck ? 'Merit Properties · Check a Number' : 'Merit Properties · Customer Claims')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);   // allows embedding in Google Sites
}

/**
 * JSON API used by the website on GitHub Pages.
 * The site sends POST {fn, args}; only the functions listed here can be called.
 */
function doPost(e) {
  const API = {
    login: login, logout: logout, whoAmI: whoAmI, changePassword: changePassword,
    getAgents: getAgents, checkPhone: checkPhone, register: register, listRecords: listRecords,
    listUsers: listUsers, saveUser: saveUser, resetUserPassword: resetUserPassword, saveAgent: saveAgent, addAgentQuick: addAgentQuick,
    publicCheck: publicCheck,
  };
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = API.hasOwnProperty(req.fn) ? API[req.fn] : null;
    if (!fn) throw new Error('Unknown action.');
    out = { result: fn.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { error: String((err && err.message) || err).replace(/^Error:\s*/, '') };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Used by the public Check page. Shows only status, never the customer's name. */
function publicCheck(phone) {
  if (!PUBLIC_CHECK.enabled) throw new Error('Number checking is turned off. Ask the office.');
  const cache = CacheService.getScriptCache();
  const k = 'pub:' + Utilities.formatDate(new Date(), CFG.TZ, 'yyyyMMddHHmm');
  const n = Number(cache.get(k) || 0) + 1;
  cache.put(k, String(n), 120);
  if (n > PUBLIC_CHECK.maxPerMinute) throw new Error('Too many checks right now. Wait a minute and try again.');
  const p = parseNewPhone_(phone);
  if (!p.valid) return { valid: false, error: p.error };
  const all = readAllClaims_();
  const hit = findLock_(p.key, all.rows, all.today);
  if (!hit) return { valid: true, formatted: p.formatted, claimed: false };
  return {
    valid: true, formatted: p.formatted, claimed: true, category: hit.categoryLabel,
    date: dayStr_(hit.day), lockedUntil: dayStr_(hit.day + CFG.LOCK_DAYS), freeFrom: dayStr_(hit.day + CFG.LOCK_DAYS + 1),
    agent: PUBLIC_CHECK.showAgent ? hit.via : '',
  };
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Claims System')
    .addItem('Run setup', 'setupFromMenu')
    .addItem('Reset admin password', 'resetAdminFromMenu')
    .addToUi();
}

/* ───────────────────────── Setup ───────────────────────── */

/** Creates the Agents and Users tabs (only if missing) and the first admin user. Safe to run again. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const msgs = [];

  CATEGORIES.forEach(function (cat) {
    const sh = categorySheet_(cat, true);
    layout_(sh, true);
    msgs.push('Category tab ready: "' + sh.getName() + '"');
  });

  let ag = ss.getSheetByName(CFG.AGENTS_SHEET);
  if (!ag) {
    ag = ss.insertSheet(CFG.AGENTS_SHEET);
    ag.getRange(1, 1, 1, AGENT_HEADERS.length).setValues([AGENT_HEADERS]).setFontWeight('bold');
    const rows = SEED_AGENTS.map(function (a, i) { return ['A' + (i + 1), a[0], a[1], a[2], a[3], true]; });
    if (rows.length) {
      ag.getRange(2, 5, rows.length, 1).setNumberFormat('@');
      ag.getRange(2, 1, rows.length, AGENT_HEADERS.length).setValues(rows);
    }
    ag.setFrozenRows(1);
    msgs.push('Agents tab created with ' + rows.length + ' names');
  }

  let us = ss.getSheetByName(CFG.USERS_SHEET);
  if (!us) {
    us = ss.insertSheet(CFG.USERS_SHEET);
    us.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]).setFontWeight('bold');
    us.setFrozenRows(1);
    us.hideSheet();
    msgs.push('Users tab created (hidden)');
  }
  if (!readUsers_().some(function (u) { return u.role === 'admin' && u.active; })) {
    createUser_('admin', 'Administrator', 'admin', 'ChangeMe123', true);
    msgs.push('Admin user created → username: admin  password: ChangeMe123 (must be changed at first login)');
  }
  Logger.log(msgs.join('\n'));
  return msgs.join('\n');
}

function setupFromMenu() {
  SpreadsheetApp.getUi().alert(setup());
}

function resetAdminFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Reset admin password', 'Username of the admin to reset:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const u = findUser_(r.getResponseText().trim());
  if (!u) { ui.alert('User not found'); return; }
  setPassword_(u, 'ChangeMe123', true);
  ui.alert('Password for "' + u.username + '" reset to ChangeMe123. It must be changed at next login.');
}

/* ───────────────────────── Auth ───────────────────────── */

function login(username, password) {
  username = String(username || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const failKey = 'fail:' + username;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= CFG.MAX_LOGIN_FAILS) throw new Error('Too many wrong attempts. Try again in 15 minutes.');

  const u = findUser_(username);
  if (!u || !u.active || hash_(String(password || ''), u.salt) !== u.hash) {
    cache.put(failKey, String(fails + 1), 15 * 60);
    throw new Error('Wrong username or password.');
  }
  cache.remove(failKey);
  usersSheet_().getRange(u.row, 9).setValue(nowStr_());
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('sess:' + token, u.username, CFG.SESSION_SECONDS);
  return { token: token, user: publicUser_(u) };
}

function logout(token) {
  CacheService.getScriptCache().remove('sess:' + token);
  return true;
}

function whoAmI(token) {
  return publicUser_(auth_(token, null, true));
}

function changePassword(token, oldPw, newPw) {
  const u = auth_(token, null, true);
  if (hash_(String(oldPw || ''), u.salt) !== u.hash) throw new Error('Current password is wrong.');
  checkPasswordStrength_(newPw);
  setPassword_(u, newPw, false);
  return publicUser_(findUser_(u.username));
}

/** Returns the user for a valid session, or throws. roles = allowed roles (null = any). */
function auth_(token, roles, allowMustChange) {
  const cache = CacheService.getScriptCache();
  const username = token ? cache.get('sess:' + token) : null;
  if (!username) throw new Error('SESSION_EXPIRED');
  const u = findUser_(username);
  if (!u || !u.active) { cache.remove('sess:' + token); throw new Error('SESSION_EXPIRED'); }
  if (u.mustChange && !allowMustChange) throw new Error('PASSWORD_CHANGE_REQUIRED');
  if (roles && roles.indexOf(u.role) < 0) throw new Error('You do not have permission for this.');
  cache.put('sess:' + token, u.username, CFG.SESSION_SECONDS); // sliding expiry
  return u;
}

function hash_(pw, salt) {
  let h = salt + '|' + pw;
  for (let i = 0; i < 1000; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  }
  return h;
}

function checkPasswordStrength_(pw) {
  pw = String(pw || '');
  if (pw.length < 6) throw new Error('Password must be at least 6 characters.');
}

/* ───────────────────────── Users (admin) ───────────────────────── */

function usersSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.USERS_SHEET);
  if (!sh) throw new Error('System not set up yet. Open the spreadsheet and run Claims System → Run setup.');
  return sh;
}

function readUsers_() {
  const sh = usersSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, USER_HEADERS.length).getValues().map(function (r, i) {
    return {
      row: i + 2, username: String(r[0]).trim().toLowerCase(), name: String(r[1]), role: String(r[2]),
      active: r[3] === true || String(r[3]).toUpperCase() === 'TRUE', salt: String(r[4]), hash: String(r[5]),
      mustChange: r[6] === true || String(r[6]).toUpperCase() === 'TRUE', created: fmtCell_(r[7]), lastLogin: fmtCell_(r[8]),
    };
  }).filter(function (u) { return u.username; });
}

function findUser_(username) {
  username = String(username || '').trim().toLowerCase();
  return readUsers_().filter(function (u) { return u.username === username; })[0] || null;
}

function publicUser_(u) {
  return { username: u.username, name: u.name, role: u.role, active: u.active, mustChange: u.mustChange, created: u.created, lastLogin: u.lastLogin };
}

function createUser_(username, name, role, pw, mustChange) {
  const salt = Utilities.getUuid();
  usersSheet_().appendRow([username, name, role, true, salt, hash_(pw, salt), mustChange, nowStr_(), '']);
}

function setPassword_(u, pw, mustChange) {
  const salt = Utilities.getUuid();
  usersSheet_().getRange(u.row, 5, 1, 3).setValues([[salt, hash_(pw, salt), mustChange]]);
}

function listUsers(token) {
  auth_(token, ['admin']);
  return readUsers_().map(publicUser_);
}

/** Create (isNew) or update a user. Password only needed when creating. */
function saveUser(token, data) {
  const me = auth_(token, ['admin']);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const username = String(data.username || '').trim().toLowerCase();
    const name = String(data.name || '').trim();
    const role = String(data.role || '');
    const active = data.active !== false;
    if (!/^[a-z0-9._-]{3,30}$/.test(username)) throw new Error('Username: 3–30 letters/numbers, no spaces.');
    if (!name) throw new Error('Full name is required.');
    if (ROLES.indexOf(role) < 0) throw new Error('Choose a role.');
    const existing = findUser_(username);
    if (data.isNew) {
      if (existing) throw new Error('Username "' + username + '" already exists.');
      checkPasswordStrength_(data.password);
      createUser_(username, name, role, String(data.password), true);
    } else {
      if (!existing) throw new Error('User not found.');
      if (existing.username === me.username && (role !== 'admin' || !active)) throw new Error('You cannot remove your own admin access.');
      usersSheet_().getRange(existing.row, 2, 1, 3).setValues([[name, role, active]]);
    }
    return listUsers(token);
  } finally { lock.releaseLock(); }
}

function resetUserPassword(token, username, newPw) {
  auth_(token, ['admin']);
  checkPasswordStrength_(newPw);
  const u = findUser_(username);
  if (!u) throw new Error('User not found.');
  setPassword_(u, String(newPw), true);
  return true;
}

/* ───────────────────────── Agents ───────────────────────── */

function agentsSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.AGENTS_SHEET);
  if (!sh) throw new Error('System not set up yet. Open the spreadsheet and run Claims System → Run setup.');
  return sh;
}

function readAgents_() {
  const sh = agentsSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, AGENT_HEADERS.length).getValues().map(function (r, i) {
    return {
      row: i + 2, id: String(r[0]), type: String(r[1]).trim(), name: String(r[2]).trim(),
      supervisor: String(r[3]).trim(), phone: String(r[4]).trim(),
      active: !(r[5] === false || String(r[5]).toUpperCase() === 'FALSE'),
    };
  }).filter(function (a) { return a.id && a.name; });
}

function getAgents(token) {
  auth_(token);
  return { types: AGENT_TYPES, agents: readAgents_().map(function (a) { delete a.row; return a; }) };
}

function saveAgent(token, data) {
  auth_(token, ['admin']);
  saveAgentCore_(data, '');
  return getAgents(token);
}

/**
 * Add a new person straight from the Register form (admin + data entry).
 * Only adds — editing existing names stays admin-only. If the name already exists, that entry is returned.
 */
function addAgentQuick(token, data) {
  const u = auth_(token, ['admin', 'entry']);
  data = { type: data.type, name: data.name, supervisor: data.supervisor, phone: data.phone, active: true };
  const agent = saveAgentCore_(data, u.name || u.username, true);
  return { agent: agent, agents: getAgents(token).agents };
}

function saveAgentCore_(data, addedBy, reuseExisting) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const type = String(data.type || '');
    const name = String(data.name || '').trim().replace(/\s+/g, ' ');
    let supervisor = String(data.supervisor || '').trim();
    const phone = String(data.phone || '').trim();
    const active = data.active !== false;
    if (AGENT_TYPES.indexOf(type) < 0) throw new Error('Choose a type.');
    if (!name) throw new Error('Name is required.');
    if (name.length > 60) throw new Error('Name is too long.');
    if (type === 'Sales' && !supervisor) throw new Error('Sales agents need a supervisor.');
    if (type === 'Elite Sales' || type === 'Supervisor' || type === 'Agent') supervisor = '';
    const all = readAgents_();
    if (supervisor && !all.some(function (a) { return a.type === 'Supervisor' && a.name === supervisor; })) throw new Error('Choose a supervisor from the list.');
    const dup = all.filter(function (a) { return a.type === type && a.name.toLowerCase() === name.toLowerCase() && a.id !== data.id; })[0];
    if (dup) {
      if (reuseExisting && dup.active) { delete dup.row; return dup; }
      throw new Error('"' + name + '" already exists under ' + type + (dup.active ? '.' : ' (inactive — an admin can turn it back on in Agents).'));
    }
    const sh = agentsSheet_();
    if (data.id) {
      const a = all.filter(function (x) { return x.id === data.id; })[0];
      if (!a) throw new Error('Agent not found.');
      sh.getRange(a.row, 5).setNumberFormat('@');
      sh.getRange(a.row, 2, 1, 5).setValues([[type, name, supervisor, phone, active]]);
      return { id: a.id, type: type, name: name, supervisor: supervisor, phone: phone, active: active };
    }
    const maxId = all.reduce(function (m, a) { return Math.max(m, Number(a.id.replace(/\D/g, '')) || 0); }, 0);
    const id = 'A' + (maxId + 1);
    const row = sh.getLastRow() + 1;
    if (!String(sh.getRange(1, 7).getValue())) sh.getRange(1, 7, 1, 2).setValues([['Added By', 'Added At']]).setFontWeight('bold');
    sh.getRange(row, 5).setNumberFormat('@');
    sh.getRange(row, 8).setNumberFormat('@');
    sh.getRange(row, 1, 1, 8).setValues([[id, type, name, supervisor, phone, active, addedBy || '', nowStr_()]]);
    return { id: id, type: type, name: name, supervisor: supervisor, phone: phone, active: active };
  } finally { lock.releaseLock(); }
}

/* ───────────────────────── Category tabs ───────────────────────── */

function categorySheet_(cat, create) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheets().filter(function (s) {
    const n = s.getName();
    return n !== CFG.AGENTS_SHEET && n !== CFG.USERS_SHEET && cat.match.test(n);
  })[0];
  if (sh) return sh;
  if (!create) return null;
  const ns = ss.insertSheet(cat.label);
  ns.getRange(1, 1, 1, 5).setValues([['S/N', 'Date', 'Customer Name', 'Phone No', 'Via']]).setFontWeight('bold');
  ns.setFrozenRows(1);
  return ns;
}

/** Finds the header row and the column of each field. Adds the extra columns if asked. */
function layout_(sh, addExtras) {
  const lastCol = Math.max(sh.getLastColumn(), 5);
  const top = sh.getRange(1, 1, Math.min(6, Math.max(sh.getLastRow(), 1)), lastCol).getValues();
  let hr = -1;
  for (let i = 0; i < top.length && hr < 0; i++) {
    if (top[i].some(function (v) { return /phone/i.test(String(v)); })) hr = i;
  }
  if (hr < 0) throw new Error('Could not find the header row (with "Phone") in tab "' + sh.getName() + '".');
  const h = top[hr].map(function (v) { return String(v).trim(); });
  const find = function (test) { for (let c = 0; c < h.length; c++) if (h[c] && test(h[c])) return c + 1; return 0; };
  const L = { headerRow: hr + 1, sheet: sh };
  L.phone = find(function (s) { return /phone/i.test(s); });
  L.date = find(function (s) { return /date/i.test(s); });
  L.name = find(function (s) { return /name|customer/i.test(s) && !/phone/i.test(s); });
  L.sn = find(function (s) { return /^(s\s*\/?\s*n|sr\.?\s*no|no\.?)$/i.test(s); });
  L.via = find(function (s) { return /via|sales|agent/i.test(s) && !/type/i.test(s); }) || (L.phone + 1);
  if (addExtras && !h[L.via - 1]) { sh.getRange(L.headerRow, L.via).setValue('Via').setFontWeight('bold'); h[L.via - 1] = 'Via'; }
  const extraKeys = ['type', 'supervisor', 'enteredBy', 'enteredAt'];
  let next = L.via + 1;
  EXTRA_HEADERS.forEach(function (title, i) {
    let c = find(function (s) { return s.toLowerCase() === title.toLowerCase(); });
    if (!c) {
      while (h[next - 1]) next++;           // skip columns that already have another header
      c = next;
      if (addExtras) {
        sh.getRange(L.headerRow, c).setValue(title).setFontWeight('bold');
        h[c - 1] = title;
      }
    }
    L[extraKeys[i]] = c;
    next = Math.max(next, c + 1);
  });
  L.width = Math.max(L.phone, L.date, L.name, L.via, L.type, L.supervisor, L.enteredBy, L.enteredAt, L.sn);
  return L;
}

/** Reads every claim row of a category tab. */
function readClaims_(cat, today, ssTz) {
  const sh = categorySheet_(cat, false);
  if (!sh) return { rows: [], lastDataRow: 0, layout: null };
  const L = layout_(sh, false);
  const last = sh.getLastRow();
  const out = [];
  let lastDataRow = L.headerRow;
  if (last > L.headerRow) {
    const vals = sh.getRange(L.headerRow + 1, 1, last - L.headerRow, L.width).getValues();
    const cands = [];
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      const g = function (c) { return c ? r[c - 1] : ''; };
      const name = clean_(g(L.name)), phoneRaw = g(L.phone);
      const phoneStr = cellStr_(phoneRaw);
      if (!name && !phoneStr) continue;
      lastDataRow = L.headerRow + 1 + i;
      const enteredAt = cellStr_(g(L.enteredAt));
      const stamp = enteredAt ? cellDayCands_(enteredAt, ssTz, today) : [];   // set by this system: exact
      cands.push(stamp.length ? stamp : cellDayCands_(g(L.date), ssTz, today));
      out.push({
        category: cat.key, categoryLabel: cat.label, row: lastDataRow,
        day: null, name: name, phone: phoneStr, keys: phoneKeys_(phoneRaw),
        via: clean_(g(L.via)), type: clean_(g(L.type)), supervisor: clean_(g(L.supervisor)),
        enteredBy: clean_(g(L.enteredBy)), enteredAt: enteredAt,
      });
    }
    resolveDays_(cands).forEach(function (d, i) { out[i].day = d; });
  }
  return { rows: out, lastDataRow: lastDataRow, layout: L };
}

function readAllClaims_() {
  const ssTz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const today = todayDay_();
  let rows = [];
  CATEGORIES.forEach(function (cat) { rows = rows.concat(readClaims_(cat, today, ssTz).rows); });
  return { rows: rows, today: today };
}

/* ───────────────────────── Phone + date logic ───────────────────────── */

/** All match keys (first 8 digits) found in a phone cell. A cell may hold several numbers. */
function phoneKeys_(v) {
  if (v === null || v === undefined || v === '') return [];
  const s = typeof v === 'number' ? String(Math.round(v)) : String(v);
  const keys = [];
  s.split(/[\/,;\n()&]|\bor\b/i).forEach(function (p) {
    const k = phoneKey_(p);
    if (k && keys.indexOf(k) < 0) keys.push(k);
  });
  return keys;
}

function phoneKey_(p) {
  const d = normDigits_(p);
  return d && d.length >= CFG.KEY_DIGITS ? d.slice(0, CFG.KEY_DIGITS) : null;
}

/** Digits in local form starting with 09/07, or null. Handles +251 and a lost leading zero. */
function normDigits_(p) {
  let d = String(p).replace(/\D/g, '');
  if (d.indexOf('00251') === 0) d = '0' + d.slice(5);
  else if (d.indexOf('251') === 0 && d.length >= 11) d = '0' + d.slice(3);
  if (/^[79]/.test(d)) d = '0' + d;
  return /^0[79]/.test(d) ? d : null;
}

/** Validates a newly typed number: 8 digits (last 2 hidden) or full 10 digits. */
function parseNewPhone_(input) {
  const d = normDigits_(input || '');
  if (!d || (d.length !== 8 && d.length !== 10)) {
    return { valid: false, error: 'Enter a mobile number starting with 09 or 07 — the first 8 digits (0911 51 83) or all 10.' };
  }
  return { valid: true, digits: d, key: d.slice(0, CFG.KEY_DIGITS), formatted: fmtPhone_(d) };
}

function fmtPhone_(d) {
  return d.slice(0, 4) + ' ' + d.slice(4, 6) + ' ' + d.slice(6, 8) + (d.length > 8 ? ' ' + d.slice(8) : '');
}

/**
 * Possible day numbers (days since 1970-01-01) for a date cell.
 * Real date cells with day and month both ≤ 12 were sometimes saved swapped (US order),
 * so they get two candidates; resolveDays_ picks the one that fits the neighbouring rows.
 * Dates after today are typos (e.g. 2035) and are dropped.
 */
function cellDayCands_(v, ssTz, today) {
  let y, m, d, ambiguous = false;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return [];
    const p = Utilities.formatDate(v, ssTz, 'yyyy-M-d').split('-');
    y = +p[0]; m = +p[1]; d = +p[2];
    ambiguous = d <= 12 && d !== m;
  } else if (typeof v === 'number') {
    const n = v > 20000 && v < 80000 ? Math.floor(v) - 25569 : null; // spreadsheet serial number
    return n !== null && n <= today ? [n] : [];
  } else {
    const mt = String(v || '').trim().match(/^(\d{1,2})\s*[\/.\-]\s*(\d{1,2})\s*[\/.\-]\s*(\d{2,4})/);
    if (!mt) return [];
    d = +mt[1]; m = +mt[2]; y = +mt[3];                      // typed text is day/month/year
    if (y < 100) y += 2000;
    if (m > 12 && d <= 12) { const t = d; d = m; m = t; }
  }
  if (!(m >= 1 && m <= 12 && d >= 1 && d <= 31 && y > 2000 && y < 2100)) return [];
  const out = [dayNum_(y, m, d)];
  if (ambiguous) out.push(dayNum_(y, d, m));
  return out.filter(function (x) { return x <= today; });
}

/** Picks one day per row: ambiguous rows take the candidate closest to nearby clear dates (sheets are in date order). */
function resolveDays_(cands) {
  const clear = [];
  cands.forEach(function (c, i) { if (c.length === 1) clear.push(i); });
  let p = 0;
  return cands.map(function (c, i) {
    if (c.length < 2) return c.length ? c[0] : null;
    while (p < clear.length && clear[p] < i) p++;
    const near = clear.slice(Math.max(0, p - 8), p).concat(clear.slice(p, p + 8)).map(function (j) { return cands[j][0]; });
    if (!near.length) return c[0];
    near.sort(function (a, b) { return a - b; });
    const med = near[Math.floor(near.length / 2)];
    return Math.abs(c[1] - med) < Math.abs(c[0] - med) ? c[1] : c[0];
  });
}

/** Single date cell → day number (used for text/serial values where order is known). */
function cellDay_(v, ssTz, today) {
  return resolveDays_([cellDayCands_(v, ssTz, today)])[0];
}

function dayNum_(y, m, d) { return Math.floor(Date.UTC(y, m - 1, d) / 864e5); }

function todayDay_() {
  const p = Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-M-d').split('-');
  return dayNum_(+p[0], +p[1], +p[2]);
}

function dayStr_(day) {
  if (day === null || day === undefined) return '';
  const dt = new Date(day * 864e5);
  const mon = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][dt.getUTCMonth()];
  return dt.getUTCDate() + ' ' + mon + ' ' + dt.getUTCFullYear();
}

/** Latest registration of this key that is still inside the lock window, or null. */
function findLock_(key, rows, today) {
  let best = null;
  rows.forEach(function (r) {
    if (r.day === null || r.keys.indexOf(key) < 0) return;
    if (today - r.day <= CFG.LOCK_DAYS && (!best || r.day > best.day)) best = r;
  });
  return best;
}

function claimView_(r, today) {
  const v = {
    category: r.categoryLabel, date: dayStr_(r.day), name: r.name, phone: r.phone, via: r.via,
    type: r.type, enteredBy: r.enteredBy, day: r.day,
    supervisor: r.supervisor || ((r.via.match(/\(([^)]+)\)/) || [])[1] || '').trim(),   // old rows: "name(supervisor)"
  };
  if (r.day !== null && today - r.day <= CFG.LOCK_DAYS) {
    v.lockedUntil = dayStr_(r.day + CFG.LOCK_DAYS);
    v.freeFrom = dayStr_(r.day + CFG.LOCK_DAYS + 1);
    v.daysLeft = r.day + CFG.LOCK_DAYS - today;   // 0 = last locked day
  }
  return v;
}

/* ───────────────────────── Check + Register ───────────────────────── */

function checkPhone(token, phone) {
  auth_(token);
  const p = parseNewPhone_(phone);
  if (!p.valid) return p;
  const all = readAllClaims_();
  const lockRow = findLock_(p.key, all.rows, all.today);
  const history = all.rows.filter(function (r) { return r.keys.indexOf(p.key) >= 0; })
    .sort(function (a, b) { return (b.day || 0) - (a.day || 0); })
    .slice(0, 15).map(function (r) { return claimView_(r, all.today); });
  return { valid: true, formatted: p.formatted, key: p.key, lock: lockRow ? claimView_(lockRow, all.today) : null, history: history };
}

function register(token, data) {
  const u = auth_(token, ['admin', 'entry']);
  const cat = CATEGORIES.filter(function (c) { return c.key === data.category; })[0];
  if (!cat) throw new Error('Choose a category.');
  const name = String(data.name || '').trim().replace(/\s+/g, ' ');
  if (!name) throw new Error('Customer name is required.');
  if (name.length > 80) throw new Error('Customer name is too long.');
  const p = parseNewPhone_(data.phone);
  if (!p.valid) throw new Error(p.error);
  const agent = readAgents_().filter(function (a) { return a.id === data.agentId && a.active; })[0];
  if (!agent) throw new Error('Choose the agent from the list.');
  const supervisor = agent.type === 'Supervisor' ? agent.name : agent.supervisor;
  const via = agent.name + (supervisor && agent.type !== 'Supervisor' ? ' (' + supervisor + ')' : '');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssTz = ss.getSpreadsheetTimeZone();
    const today = todayDay_();
    let rows = [], target = null;
    CATEGORIES.forEach(function (c) {
      const res = readClaims_(c, today, ssTz);
      rows = rows.concat(res.rows);
      if (c.key === cat.key) target = res;
    });
    const hit = findLock_(p.key, rows, today);
    if (hit) return { ok: false, lock: claimView_(hit, today) };

    const sh = categorySheet_(cat, true);
    const L = layout_(sh, true);
    const r = Math.max(target && target.layout ? target.lastDataRow : L.headerRow, L.headerRow) + 1;
    if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
    const range = sh.getRange(r, 1, 1, L.width);
    const vals = range.getValues()[0];
    const set = function (c, v) { if (c) vals[c - 1] = v; };
    const todayStr = Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd');
    set(L.date, Utilities.parseDate(todayStr, ssTz, 'yyyy-MM-dd'));
    set(L.name, name);
    set(L.phone, p.formatted);
    set(L.via, via);
    set(L.type, agent.type);
    set(L.supervisor, supervisor || '');
    set(L.enteredBy, u.name || u.username);
    set(L.enteredAt, nowStr_());
    if (L.phone) sh.getRange(r, L.phone).setNumberFormat('@');
    if (L.enteredAt) sh.getRange(r, L.enteredAt).setNumberFormat('@');
    range.setValues([vals]);
    if (L.date) sh.getRange(r, L.date).setNumberFormat('dd/mm/yyyy');
    SpreadsheetApp.flush();
    return {
      ok: true, category: cat.label, row: r, name: name, phone: p.formatted, agent: agent.name, agentType: agent.type,
      supervisor: supervisor || '', date: dayStr_(today), lockedUntil: dayStr_(today + CFG.LOCK_DAYS),
    };
  } finally {
    lock.releaseLock();
  }
}

/* ───────────────────────── Records (viewer/admin) ───────────────────────── */

function listRecords(token, f) {
  auth_(token, ['admin', 'viewer']);
  f = f || {};
  const all = readAllClaims_();
  const today = all.today;
  const days = Number(f.days || 0);
  const q = String(f.q || '').trim().toLowerCase();
  const qDigits = q.replace(/\D/g, '');
  const qKey = qDigits.length >= 4 ? (normDigits_(qDigits) || qDigits) : '';
  let rows = all.rows.filter(function (r) {
    if (f.category && f.category !== 'all' && r.category !== f.category) return false;
    if (days && (r.day === null || today - r.day >= days)) return false;
    if (f.lockedOnly && (r.day === null || today - r.day > CFG.LOCK_DAYS)) return false;
    if (!q) return true;
    if (qKey && r.keys.some(function (k) { return k.indexOf(qKey.slice(0, CFG.KEY_DIGITS)) === 0 || qKey.indexOf(k) === 0; })) return true;
    return (r.name + ' ' + r.via + ' ' + r.supervisor + ' ' + r.enteredBy).toLowerCase().indexOf(q) >= 0;
  });
  rows.sort(function (a, b) { return (b.day === null ? -1e9 : b.day) - (a.day === null ? -1e9 : a.day) || b.row - a.row; });
  const total = rows.length;
  const stats = { today: {}, last30: {} };
  CATEGORIES.forEach(function (c) { stats.today[c.label] = 0; stats.last30[c.label] = 0; });
  all.rows.forEach(function (r) {
    if (r.day === today) stats.today[r.categoryLabel]++;
    if (r.day !== null && today - r.day < 30 && today - r.day >= 0) stats.last30[r.categoryLabel]++;
  });
  return { total: total, shown: Math.min(total, CFG.MAX_RECORDS), stats: stats, rows: rows.slice(0, CFG.MAX_RECORDS).map(function (r) { return claimView_(r, today); }) };
}

/* ───────────────────────── Helpers ───────────────────────── */

function nowStr_() { return Utilities.formatDate(new Date(), CFG.TZ, 'dd/MM/yyyy HH:mm'); }
function clean_(v) { return v === null || v === undefined ? '' : String(v).replace(/\s+/g, ' ').trim(); }
function cellStr_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') return String(Math.round(v));
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, CFG.TZ, 'dd/MM/yyyy HH:mm');
  return clean_(v);
}
function fmtCell_(v) { return cellStr_(v); }
