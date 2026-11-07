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
  EDIT_HOURS: 24,             // data entry can edit their own entries for this long (admins: any time)
  MAX_PHONES: 3,              // phone numbers per customer
};

const CATEGORIES = [
  { key: 'expecting', label: 'Expecting', match: /expect/i },
  { key: 'office', label: 'Office Visit', match: /office/i },
  { key: 'site', label: 'Site Visit', match: /site/i },
];

const AGENT_TYPES = ['Sales', 'Elite Sales', 'Supervisor', 'Freelancer', 'Agent'];
const ROLES = ['admin', 'entry', 'viewer'];
const EXTRA_HEADERS = ['Agent Type', 'Supervisor', 'Entered By', 'Entered At', 'Edited', 'Note'];
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
    publicCheck: publicCheck, editRecord: editRecord, listMyEntries: listMyEntries, phoneHistory: phoneHistory,
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
  const all = recentClaims_();
  const hit = findLock_({ key: p.key, digits: p.digits, intl: p.intl }, all.rows, all.today);
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
  clearCaches_();
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
  try { usersSheet_().getRange(userRow_(u.username), 9).setValue(nowStr_()); cache.remove('users'); } catch (e) {}
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
  const c = CacheService.getScriptCache(), hit = c.get('users');
  if (hit) return JSON.parse(hit);
  const list = readUsersSheet_();
  c.put('users', JSON.stringify(list), 300);
  return list;
}

function readUsersSheet_() {
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
  CacheService.getScriptCache().remove('users');
}

/** Current sheet row of a user, read fresh from the sheet (never from cache) before writing. */
function userRow_(username) {
  const u = readUsersSheet_().filter(function (x) { return x.username === username; })[0];
  if (!u) throw new Error('User not found.');
  return u.row;
}

function setPassword_(u, pw, mustChange) {
  const salt = Utilities.getUuid();
  usersSheet_().getRange(userRow_(u.username), 5, 1, 3).setValues([[salt, hash_(pw, salt), mustChange]]);
  CacheService.getScriptCache().remove('users');
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
      usersSheet_().getRange(userRow_(existing.username), 2, 1, 3).setValues([[name, role, active]]);
      CacheService.getScriptCache().remove('users');
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
  const c = CacheService.getScriptCache(), hit = c.get('agents');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const list = readAgentsSheet_();
  const str = JSON.stringify(list);
  if (str.length < 90000) c.put('agents', str, 300);
  return list;
}

function readAgentsSheet_() {
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
    const all = readAgentsSheet_();
    CacheService.getScriptCache().remove('agents');
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
  } finally { CacheService.getScriptCache().remove('agents'); lock.releaseLock(); }
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
  const extraKeys = ['type', 'supervisor', 'enteredBy', 'enteredAt', 'edited', 'note'];
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
  L.width = Math.max(L.phone, L.date, L.name, L.via, L.type, L.supervisor, L.enteredBy, L.enteredAt, L.edited || 0, L.note || 0, L.sn);
  return L;
}

/** Turns raw sheet values (starting at sheet row firstRow) into claim objects. */
function parseClaims_(cat, L, vals, firstRow, today, ssTz) {
  const out = [], cands = [];
  let lastDataRow = 0;
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    const g = function (c) { return c ? r[c - 1] : ''; };
    const name = clean_(g(L.name)), phoneRaw = g(L.phone);
    const phoneStr = cellStr_(phoneRaw);
    if (!name && !phoneStr) continue;
    lastDataRow = firstRow + i;
    const enteredAt = cellStr_(g(L.enteredAt));
    const stamp = enteredAt ? cellDayCands_(enteredAt, ssTz, today) : [];   // set by this system: exact
    cands.push(stamp.length ? stamp : cellDayCands_(g(L.date), ssTz, today));
    out.push({
      category: cat.key, categoryLabel: cat.label, row: lastDataRow,
      day: null, name: name, phone: phoneStr, keys: phoneKeys_(phoneRaw),
      via: clean_(g(L.via)), type: clean_(g(L.type)), supervisor: clean_(g(L.supervisor)),
      enteredBy: clean_(g(L.enteredBy)), enteredAt: enteredAt, note: clean_(g(L.note)),
    });
  }
  resolveDays_(cands).forEach(function (d, i) { out[i].day = d; });
  return { rows: out, lastDataRow: lastDataRow };
}

/** Reads every claim row of a category tab (used for full history / "All time" records). */
function readClaims_(cat, today, ssTz) {
  const sh = categorySheet_(cat, false);
  if (!sh) return { rows: [], lastDataRow: 0, layout: null };
  const L = layout_(sh, false);
  const last = sh.getLastRow();
  if (last <= L.headerRow) return { rows: [], lastDataRow: L.headerRow, layout: L };
  const res = parseClaims_(cat, L, sh.getRange(L.headerRow + 1, 1, last - L.headerRow, L.width).getValues(), L.headerRow + 1, today, ssTz);
  return { rows: res.rows, lastDataRow: res.lastDataRow || L.headerRow, layout: L };
}

/**
 * Reads only the recent end of a tab: the sheets are in date order, so it reads blocks
 * from the bottom up and stops once a block's top rows are clearly older than ~45 days.
 */
function readClaimsRecent_(cat, today, ssTz) {
  const sh = categorySheet_(cat, false);
  if (!sh) return [];
  const L = layout_(sh, false);
  const CHUNK = 1500, OLD = today - (CFG.LOCK_DAYS + 15);
  let end = sh.getLastRow(), vals = [], first = end + 1;
  while (end > L.headerRow) {
    const start = Math.max(L.headerRow + 1, end - CHUNK + 1);
    const chunk = sh.getRange(start, 1, end - start + 1, L.width).getValues();
    vals = chunk.concat(vals); first = start; end = start - 1;
    const clear = [];
    for (let i = 0; i < chunk.length && clear.length < 40; i++) {      // clear (unambiguous) dates at the top of the block
      const c = cellDayCands_(chunk[i][L.date - 1], ssTz, today);
      if (c.length === 1) clear.push(c[0]);
    }
    if (clear.length >= 10) {
      clear.sort(function (a, b) { return a - b; });
      if (clear[Math.floor(clear.length / 2)] < OLD) break;
    }
  }
  return parseClaims_(cat, L, vals, first, today, ssTz).rows;
}

function readAllClaims_() {
  const ssTz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const today = todayDay_();
  let rows = [];
  CATEGORIES.forEach(function (cat) { rows = rows.concat(readClaims_(cat, today, ssTz).rows); });
  return { rows: rows, today: today };
}

/* Recent claims (last 31 days) kept in the script cache so checks don't have to open the sheet.
   Cleared when someone edits the sheet by hand (onEdit) and updated on every registration. */
const IDX_TTL = 300;          // seconds
const IDX_FIELDS = ['day', 'keys', 'category', 'categoryLabel', 'row', 'name', 'phone', 'via', 'type', 'supervisor', 'enteredBy', 'enteredAt', 'note'];

function recentClaims_() {
  const today = todayDay_();
  const cached = idxGet_(today);
  if (cached) return { rows: cached, today: today };
  const ssTz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  let rows = [];
  CATEGORIES.forEach(function (cat) {
    rows = rows.concat(readClaimsRecent_(cat, today, ssTz).filter(function (r) { return r.day !== null && today - r.day <= 31; }));
  });
  idxPut_(rows, today);
  return { rows: rows, today: today };
}

function idxGet_(today) {
  const c = CacheService.getScriptCache();
  const meta = c.get('idx:meta');
  if (!meta) return null;
  const m = JSON.parse(meta);
  if (m.today !== today) return null;
  const keys = [];
  for (let i = 0; i < m.n; i++) keys.push('idx:' + m.ver + ':' + i);
  const parts = c.getAll(keys);
  let str = '';
  for (let i = 0; i < keys.length; i++) { if (parts[keys[i]] == null) return null; str += parts[keys[i]]; }
  return JSON.parse(str).map(function (a) { const o = {}; IDX_FIELDS.forEach(function (f, i) { o[f] = a[i]; }); return o; });
}

function idxPut_(rows, today) {
  const c = CacheService.getScriptCache();
  const str = JSON.stringify(rows.map(function (r) { return IDX_FIELDS.map(function (f) { return r[f]; }); }));
  const ver = Utilities.getUuid().slice(0, 8), obj = {}, SIZE = 25000;   // chars per piece (cache limit 100 KB per value)
  let n = 0;
  for (let i = 0; i < str.length; i += SIZE) obj['idx:' + ver + ':' + (n++)] = str.slice(i, i + SIZE);
  if (n > 40) return;                                                     // unusually large: just don't cache
  c.putAll(obj, IDX_TTL + 60);
  c.put('idx:meta', JSON.stringify({ ver: ver, n: n, today: today }), IDX_TTL);
}

function clearCaches_() { CacheService.getScriptCache().removeAll(['idx:meta', 'hist:meta', 'users', 'agents']); }

/** Simple trigger: someone edited the sheet by hand → drop cached data so the next check re-reads it. */
function onEdit(e) {
  try {
    const n = e && e.range ? e.range.getSheet().getName() : '';
    const c = CacheService.getScriptCache();
    if (n === CFG.AGENTS_SHEET) c.remove('agents');
    else if (n === CFG.USERS_SHEET) c.remove('users');
    else c.remove('idx:meta');
  } catch (x) { /* never block an edit */ }
}

/* All-time history per phone number (last few claims each), kept in cache for 6 hours.
   Recent claims always come from recentClaims_(), so new registrations show up immediately. */
const HIST_TTL = 21600;
const HIST_KEEP = 4;

function historyIndex_() {
  const c = CacheService.getScriptCache();
  const meta = c.get('hist:meta');
  if (meta) {
    const m = JSON.parse(meta), keys = [];
    for (let i = 0; i < m.n; i++) keys.push('hist:' + m.ver + ':' + i);
    const parts = c.getAll(keys);
    let str = '', ok = true;
    for (let i = 0; i < keys.length; i++) { if (parts[keys[i]] == null) { ok = false; break; } str += parts[keys[i]]; }
    if (ok) return JSON.parse(str);
  }
  const all = readAllClaims_(), idx = {};
  const catKeys = CATEGORIES.map(function (x) { return x.key; });
  all.rows.forEach(function (r) {
    if (r.day === null) return;
    const e = [r.day, catKeys.indexOf(r.category), r.via, r.row, r.name, r.phone];
    r.keys.forEach(function (k) { (idx[k] = idx[k] || []).push(e); });
  });
  Object.keys(idx).forEach(function (k) { idx[k] = idx[k].sort(function (a, b) { return b[0] - a[0]; }).slice(0, HIST_KEEP); });
  const str = JSON.stringify(idx), ver = Utilities.getUuid().slice(0, 8), obj = {}, SIZE = 25000;
  let n = 0;
  for (let i = 0; i < str.length; i += SIZE) obj['hist:' + ver + ':' + (n++)] = str.slice(i, i + SIZE);
  if (n <= 300) { c.putAll(obj, HIST_TTL); c.put('hist:meta', JSON.stringify({ ver: ver, n: n }), HIST_TTL); }
  return idx;
}

/**
 * Earlier registrations of a number (newest first), for the line under the phone check.
 * opts.category: only that category (e.g. previous Office Visits). opts.customerName: when digits are hidden,
 * only people whose name matches (same rule as the 30-day check).
 */
function phoneHistory(token, phone, skip, opts) {
  auth_(token);
  opts = opts || {};
  const p = parseNewPhone_(phone);
  if (!p.valid) return [];
  const item = { key: p.key, digits: p.digits, intl: p.intl };
  const custName = opts.customerName ? String(opts.customerName) : null;
  const idx = historyIndex_(), recent = recentClaims_(), today = recent.today;
  const found = {}, add = function (day, catKey, via, row, name, ph) {
    if (opts.category && catKey !== opts.category) return;
    if (skip && skip.catKey === catKey && Number(skip.row) === row) return;
    if (matchReason_({ keys: rowNumbers_(ph).map(function (n) { return n.key; }), phone: ph, name: name }, item, custName, true) === null) return;
    const id = catKey + ':' + row;
    if (!found[id]) found[id] = { day: day, catKey: catKey, via: via, row: row, name: name };
  };
  recent.rows.forEach(function (r) { if (r.day !== null && hasKey_(r.keys, p.key)) add(r.day, r.category, r.via, r.row, r.name, r.phone); });
  Object.keys(idx).forEach(function (k) {
    if (!keyMatch_(k, p.key)) return;
    idx[k].forEach(function (e) {
      const cat = CATEGORIES[e[1]];
      if (!cat || today - e[0] <= 31) return;              // last 31 days come from recent (fresh) data
      add(e[0], cat.key, e[2], e[3], e[4], e[5]);
    });
  });
  return Object.keys(found).map(function (id) { return found[id]; })
    .sort(function (a, b) { return b.day - a.day; }).slice(0, 5)
    .map(function (h) {
      const cat = CATEGORIES.filter(function (c) { return c.key === h.catKey; })[0];
      return { via: h.via, name: h.name, category: cat ? cat.label : '', date: dayStr_(h.day), endsOn: dayStr_(h.day + CFG.LOCK_DAYS),
        active: today - h.day <= CFG.LOCK_DAYS };
    });
}

/** Last row that has a customer name or phone (reads two columns from the bottom up). */
function lastDataRow_(sh, L) {
  let end = sh.getLastRow();
  while (end > L.headerRow) {
    const start = Math.max(L.headerRow + 1, end - 499);
    const n = end - start + 1;
    const names = L.name ? sh.getRange(start, L.name, n, 1).getValues() : [];
    const phones = sh.getRange(start, L.phone, n, 1).getValues();
    for (let i = n - 1; i >= 0; i--) {
      if (String(phones[i][0]).trim() || (names[i] && String(names[i][0]).trim())) return start + i;
    }
    end = start - 1;
  }
  return L.headerRow;
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
  const n = normPhone_(p);
  if (!n) return null;
  if (n.intl) return '+' + n.digits;                       // other countries: full number
  return n.digits.length >= CFG.KEY_DIGITS ? n.digits.slice(0, CFG.KEY_DIGITS) : null;   // Ethiopia: first 8 digits
}

/**
 * Normalises a phone number.
 *  - Ethiopian (local) numbers → { digits: '0911518375', intl: false }. Handles +251 / 00251 / 251 and a lost leading 0.
 *  - Other countries (typed with + or 00, or 11+ digits not starting with 0) → { digits: '971501234567', intl: true }.
 */
function normPhone_(p) {
  const raw = String(p == null ? '' : p).trim();
  let d = raw.replace(/\D/g, '');
  if (!d) return null;
  let intl = false;
  if (/^\+/.test(raw)) intl = true;
  else if (d.indexOf('00') === 0) { d = d.slice(2); intl = true; }
  if (d.indexOf('251') === 0 && d.length >= 11) { d = '0' + d.slice(3); intl = false; }
  if (!intl) {
    if (/^[1-9]/.test(d) && d.length <= 9) d = '0' + d;      // leading 0 lost (number stored as a number)
    else if (/^[1-9]/.test(d) && d.length >= 11) intl = true; // e.g. 971501234567 typed without +
  }
  if (intl) return d.length >= 7 && d.length <= 15 ? { digits: d, intl: true } : null;
  return /^0[1-9]\d{5,}$/.test(d) ? { digits: d, intl: false } : null;
}

/** Old helper kept for search: local digits or null. */
function normDigits_(p) {
  const n = normPhone_(p);
  return n && !n.intl ? n.digits : null;
}

/** Two match keys refer to the same customer. International: also when one has the last digits hidden. */
function keyMatch_(a, b) {
  if (a === b) return true;
  if (a.charAt(0) !== '+' || b.charAt(0) !== '+') return false;
  const s = a.length < b.length ? a : b, l = a.length < b.length ? b : a;
  return s.length >= 9 && l.length - s.length <= 2 && l.indexOf(s) === 0;
}

function hasKey_(keys, key) {
  for (let i = 0; i < keys.length; i++) if (keyMatch_(keys[i], key)) return true;
  return false;
}

/** Validates a newly typed number. */
function parseNewPhone_(input) {
  const n = normPhone_(input || '');
  if (n && n.intl) {
    if (n.digits.length < 8) return { valid: false, error: 'That international number looks too short. Include the country code, e.g. +971 50 123 4567.' };
    return { valid: true, digits: n.digits, intl: true, key: '+' + n.digits, formatted: '+' + n.digits };
  }
  if (!n || (n.digits.length !== 8 && n.digits.length !== 10)) {
    return { valid: false, error: 'Enter an Ethiopian number as the first 8 digits (0911 51 83) or all 10, or a foreign number starting with + and the country code (+971 50 123 4567).' };
  }
  return { valid: true, digits: n.digits, intl: false, key: n.digits.slice(0, CFG.KEY_DIGITS), formatted: fmtPhone_(n.digits) };
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

/**
 * Picks one day per row: ambiguous rows take the candidate closest to nearby clear dates (sheets are in date order).
 * Rows with a blank, unreadable or impossible date take the date of the nearest dated row above (or below),
 * so they still lock their number instead of being ignored.
 */
function resolveDays_(cands) {
  const clear = [];
  cands.forEach(function (c, i) { if (c.length === 1) clear.push(i); });
  let p = 0;
  const days = cands.map(function (c, i) {
    if (c.length < 2) return c.length ? c[0] : null;
    while (p < clear.length && clear[p] < i) p++;
    const near = clear.slice(Math.max(0, p - 8), p).concat(clear.slice(p, p + 8)).map(function (j) { return cands[j][0]; });
    if (!near.length) return c[0];
    near.sort(function (a, b) { return a - b; });
    const med = near[Math.floor(near.length / 2)];
    return Math.abs(c[1] - med) < Math.abs(c[0] - med) ? c[1] : c[0];
  });
  if (days.length > 1) {                                      // fill missing dates from neighbours
    let last = null;
    const filled = days.slice();
    for (let i = 0; i < days.length; i++) { if (days[i] !== null) last = days[i]; else if (last !== null) filled[i] = last; }
    let next = null;
    for (let i = days.length - 1; i >= 0; i--) { if (days[i] !== null) next = days[i]; else if (filled[i] === null && next !== null) filled[i] = next; }
    return filled;
  }
  return days;
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
/**
 * Latest claim (within 30 days) for a number. `item` is a match key string, or a parsed number
 * { key, digits, intl } — then hidden-digit matches are decided by the customer's name (custName).
 */
/** All claims within 30 days that match (newest first). */
function findLocks_(item, rows, today, skip, custName) {
  const key = typeof item === 'string' ? item : item.key;
  return rows.filter(function (r) {
    if (r.day === null || !hasKey_(r.keys, key)) return false;
    if (typeof item !== 'string' && !claimMatches_(r, item, custName)) return false;
    if (skip && r.category === skip.category && r.row === skip.row) return false;
    return today - r.day <= CFG.LOCK_DAYS;
  }).sort(function (a, b) { return b.day - a.day || stampMs_(b.enteredAt) - stampMs_(a.enteredAt); });
}

function findLock_(item, rows, today, skip, custName) {
  const key = typeof item === 'string' ? item : item.key;
  let best = null;
  rows.forEach(function (r) {
    if (r.day === null || !hasKey_(r.keys, key)) return;
    if (typeof item !== 'string' && !claimMatches_(r, item, custName)) return;
    if (skip && r.category === skip.category && r.row === skip.row) return;   // the record being edited
    if (today - r.day > CFG.LOCK_DAYS) return;
    if (!best || r.day > best.day || (r.day === best.day && stampMs_(r.enteredAt) >= stampMs_(best.enteredAt))) best = r;   // same day: newest entry wins
  });
  return best;
}

/** The numbers in a phone cell, parsed: [{ key, digits, intl }]. */
function rowNumbers_(phone) {
  const out = [];
  String(phone == null ? '' : phone).split(/[\/,;\n()&]|\bor\b/i).forEach(function (part) {
    const n = normPhone_(part);
    if (!n) return;
    const key = n.intl ? '+' + n.digits : (n.digits.length >= CFG.KEY_DIGITS ? n.digits.slice(0, CFG.KEY_DIGITS) : null);
    if (key) out.push({ key: key, digits: n.digits, intl: n.intl });
  });
  return out;
}

/** 'same' | 'different' | 'unknown' (some digits are hidden, can't tell from the number alone). */
function compareNumbers_(a, b) {
  if (a.digits === b.digits) return a.intl || a.digits.length >= 10 ? 'same' : 'unknown';   // both hidden: can't tell
  if (a.intl !== b.intl) return 'different';
  if (a.intl) return a.digits.length === b.digits.length ? 'different' : 'unknown';
  return a.digits.length >= 10 && b.digits.length >= 10 ? 'different' : 'unknown';
}

/**
 * Is this claim row the same customer as the new number?
 *  - full numbers on both sides → compare all digits
 *  - hidden digits on either side → compare the customer's first name (no name given → assume same, to be safe)
 */
function claimMatches_(row, item, custName) {
  return matchReason_(row, item, custName) !== null;
}

/**
 * Why a claim row matches a new number:
 *  'number' – same number; 'hidden' – same first 8 digits and digits are hidden (counted as the same customer,
 *  whatever name is typed); 'name' – (only when useName) hidden digits and a similar name.
 * Rows that an admin approved as a different person ("override") only match by name.
 */
function matchReason_(row, item, custName, useName) {
  if (!hasKey_(row.keys, item.key)) return null;
  if (isOverride_(row)) useName = true;
  if (!item.digits) return 'number';
  const nums = rowNumbers_(row.phone).filter(function (n) { return keyMatch_(n.key, item.key); });
  if (!nums.length) return 'number';
  let unknown = false;
  for (let i = 0; i < nums.length; i++) {
    const c = compareNumbers_(nums[i], item);
    if (c === 'same') return 'number';
    if (c === 'unknown') unknown = true;
  }
  if (!unknown) return null;                                 // full numbers that differ in the last digits
  if (!useName) return 'hidden';
  if (custName === null || custName === undefined || !String(custName).trim()) return 'hidden';
  return sameCustomerName_(row.name, custName) ? 'name' : null;
}

function isOverride_(row) { return /^override/i.test(String(row.note || '')); }

/**
 * Names to compare: the first name, plus any name written in brackets ("jemal(kader)" → jemal, kader).
 * Titles (Ato, W/ro…) and placeholders ("unknown") are ignored.
 */
function nameTokens_(s) {
  const low = String(s || '').toLowerCase();
  const parts = [low.replace(/\(.*?\)/g, ' ')].concat((low.match(/\(([^)]*)\)/g) || []).map(function (x) { return x.slice(1, -1); }));
  const out = [];
  parts.forEach(function (p) {
    const w = p.trim().replace(/^(ato|w\/ro|w\/rt|wro|weyzero|mr|mrs|ms|dr)\.?\s+/, '').split(/\s+/)[0] || '';
    const t = w.replace(/[^a-z\u1200-\u137f]/g, '');
    if (t && !/^(unknown|unkown|unknow|unk|customer|client|noname|none|na|x+)$/.test(t)) out.push(t);
  });
  return out;
}

/** First-letter group: vowels count as one group (a/e/i/o/u, አ/ዐ); Ethiopic letters by consonant family (ሀ ሁ ሂ ሃ … are one). */
function initialGroup_(t) {
  const ch = t.charAt(0), cp = ch.charCodeAt(0);
  if ('aeiou'.indexOf(ch) >= 0) return 'V';
  if (cp >= 0x1200 && cp <= 0x137f) {
    const fam = Math.floor((cp - 0x1200) / 8);
    return fam === (0x12a0 - 0x1200) / 8 || fam === (0x12d0 - 0x1200) / 8 ? 'V' : 'E' + fam;
  }
  return ch;
}

/**
 * Same customer? Decided by the first letter of the first name (spelling of the rest varies too much):
 * same first letter → same customer. A missing name counts as the same, to be safe.
 */
function sameCustomerName_(a, b) {
  const x = nameTokens_(a), y = nameTokens_(b);
  if (!x.length || !y.length) return true;
  return x.some(function (p) { return y.some(function (q) { return initialGroup_(p) === initialGroup_(q); }); });
}

/** Kept for reference: simple spelling distance. */
function editDistance_(a, b) {
  const m = a.length, n = b.length, d = [];
  for (let i = 0; i <= m; i++) { d[i] = [i]; }
  for (let j = 0; j <= n; j++) d[0][j] = j;
  for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++) {
    d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
  }
  return d[m][n];
}

function claimView_(r, today) {
  const v = {
    category: r.categoryLabel, catKey: r.category, row: r.row, enteredAt: r.enteredAt || '', note: r.note || '',
    date: dayStr_(r.day), name: r.name, phone: r.phone, via: r.via,
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

/** who = { agentId } or { agentName }: the agent this number would be registered for (optional). */
function checkPhone(token, phone, withHistory, skip, who) {
  auth_(token);
  const p = parseNewPhone_(phone);
  if (!p.valid) return p;
  const recent = recentClaims_();
  const custName = who && who.customerName ? String(who.customerName) : null;
  const item = { key: p.key, digits: p.digits, intl: p.intl };
  const sk = skip && skip.catKey ? { category: skip.catKey, row: Number(skip.row) } : null;
  const agentName = who ? (who.agentId ? agentNameById_(who.agentId) : String(who.agentName || '')) : '';
  const blocking = firstLock_([item], recent.rows, recent.today, sk, agentName, custName);
  const lockRow = blocking || findLocks_(item, recent.rows, recent.today, sk, custName)[0] || null;
  const out = { valid: true, formatted: p.formatted, key: p.key, lock: lockRow ? claimView_(lockRow, recent.today) : null };
  if (lockRow) {
    out.lock.match = matchReason_(lockRow, item, custName);     // 'number' | 'hidden' | 'name'
    out.lock.sameAgent = !blocking;
  }

  if (withHistory) {
    const all = readAllClaims_();
    out.history = all.rows.filter(function (r) { return hasKey_(r.keys, p.key); })
      .sort(function (a, b) { return (b.day || 0) - (a.day || 0); })
      .slice(0, 15).map(function (r) { return claimView_(r, all.today); });
  }
  return out;
}

/** Validates 1–3 phone numbers. Returns { keys, formatted: '0911 51 83 75 / +971501234567' } or throws. */
function parsePhones_(list) {
  list = (Array.isArray(list) ? list : [list]).map(function (x) { return String(x || '').trim(); }).filter(String);
  if (!list.length) throw new Error('Enter the customer phone number.');
  if (list.length > CFG.MAX_PHONES) throw new Error('Up to ' + CFG.MAX_PHONES + ' phone numbers per customer.');
  const keys = [], shown = [], items = [];
  list.forEach(function (x) {
    const p = parseNewPhone_(x);
    if (!p.valid) throw new Error(x + ': ' + p.error);
    if (items.some(function (it) { return it.digits === p.digits; })) return;   // same number twice
    keys.push(p.key); shown.push(p.formatted); items.push({ key: p.key, digits: p.digits, intl: p.intl });
  });
  return { keys: keys, items: items, formatted: shown.join(' / ') };
}

function agentFields_(agentId) {
  const agent = readAgents_().filter(function (a) { return a.id === agentId && a.active; })[0];
  if (!agent) throw new Error('Choose the agent from the list.');
  const supervisor = agent.type === 'Supervisor' ? agent.name : agent.supervisor;
  return { agent: agent, supervisor: supervisor || '', via: agent.name + (supervisor && agent.type !== 'Supervisor' ? ' (' + supervisor + ')' : '') };
}

/**
 * First claim that blocks these numbers. A claim by the SAME agent does not block:
 * the agent who owns the customer may register them again (e.g. Expecting → Office Visit → Site Visit).
 */
function firstLock_(items, rows, today, skip, agentName, custName) {
  for (let i = 0; i < items.length; i++) {
    const all = findLocks_(items[i], rows, today, skip, custName);
    let hits = all.filter(function (r) { return !sameAgent_(r, agentName); });
    // an admin already approved this agent's customer as a different person → other claims only block if the name matches
    if (hits.length && agentName && custName && all.some(function (r) { return sameAgent_(r, agentName) && isOverride_(r) && sameCustomerName_(r.name, custName); })) {
      hits = hits.filter(function (r) { return sameCustomerName_(r.name, custName); });
    }
    if (hits.length) return hits[0];
  }
  return null;
}

/** Name used for matching agents: lower case, letters/digits only, "(supervisor)" removed. */
function agentKey_(s) {
  return String(s || '').toLowerCase().replace(/\(.*?\)/g, '').replace(/[^a-z0-9\u1200-\u137f]/g, '');
}

/** Lower-case words of a name, brackets included ("betlham(alemu)" → betlham, alemu). */
function nameWords_(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9\u1200-\u137f]+/g, ' ').trim().split(' ').filter(String);
}

/**
 * Does this claim row belong to the agent with this name?
 *  - rows made by this system: the agent's full name must match
 *  - old hand-typed rows ("betlham(alemu)", "abdi"): all of the agent's names appear → yes;
 *    only the first name appears → yes only if no other agent has that first name
 *    (and never if the row names another agent's surname).
 */
function sameAgent_(row, agentName) {
  const a = agentKey_(row.via), b = agentKey_(agentName);
  if (!a || !b) return false;
  if (a === b) return true;
  if (row.type) return false;
  const tok = nameWords_(row.via), aw = nameWords_(agentName);
  if (!tok.length || !aw.length) return false;
  if (aw.every(function (w) { return tok.indexOf(w) >= 0; })) return true;
  if (tok[0] !== aw[0]) return false;
  const others = readAgents_().filter(function (x) {
    return x.active && nameWords_(x.name)[0] === aw[0] && agentKey_(x.name) !== b;
  });
  if (others.some(function (o) { return nameWords_(o.name).slice(1).some(function (w) { return tok.indexOf(w) >= 0; }); })) return false;
  return others.length === 0;
}

function agentNameById_(agentId) {
  const a = readAgents_().filter(function (x) { return x.id === agentId; })[0];
  return a ? a.name : '';
}

function register(token, data) {
  const u = auth_(token, ['admin', 'entry']);
  const cat = CATEGORIES.filter(function (c) { return c.key === data.category; })[0];
  if (!cat) throw new Error('Choose a category.');
  const name = String(data.name || '').trim().replace(/\s+/g, ' ');
  if (!name) throw new Error('Customer name is required.');
  if (name.length > 80) throw new Error('Customer name is too long.');
  const ph = parsePhones_(data.phones || data.phone);
  const a = agentFields_(data.agentId);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ssTz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const recent = recentClaims_();
    const today = recent.today;
    let note = '';
    const hit = firstLock_(ph.items, recent.rows, today, null, a.agent.name, name);
    if (hit) {
      const ov = data.override;
      if (!ov) return { ok: false, lock: claimView_(hit, today), canOverride: u.role === 'admin' };
      if (u.role !== 'admin') throw new Error('Only an admin can register a claimed number.');
      const reason = String(ov.reason || '').trim();
      if (reason.length < 5) throw new Error('Write the reason (e.g. "different person, confirmed by supervisor").');
      note = 'OVERRIDE by ' + (u.name || u.username) + ': ' + reason + ' (claimed by ' + hit.via + ', ' + hit.name + ')';
    }

    const sh = categorySheet_(cat, true);
    const L = layout_(sh, true);
    const r = lastDataRow_(sh, L) + 1;
    if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
    const todayStr = Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd');
    const enteredBy = u.name || u.username, enteredAt = nowStr_();
    writeClaimRow_(sh, L, r, {
      date: Utilities.parseDate(todayStr, ssTz, 'yyyy-MM-dd'), name: name, phone: ph.formatted, via: a.via,
      type: a.agent.type, supervisor: a.supervisor, enteredBy: enteredBy, enteredAt: enteredAt, edited: '', note: note,
    });
    recent.rows.push({ day: today, keys: ph.keys, category: cat.key, categoryLabel: cat.label, row: r, name: name, phone: ph.formatted,
      via: a.via, type: a.agent.type, supervisor: a.supervisor, enteredBy: enteredBy, enteredAt: enteredAt, note: note });
    idxPut_(recent.rows, today);
    return {
      ok: true, category: cat.label, catKey: cat.key, row: r, name: name, phone: ph.formatted, agent: a.agent.name, agentType: a.agent.type,
      supervisor: a.supervisor, date: dayStr_(today), lockedUntil: dayStr_(today + CFG.LOCK_DAYS), enteredAt: enteredAt, override: !!note,
    };
  } finally {
    lock.releaseLock();
  }
}

/** Writes one claim row (only the columns this system manages). */
function writeClaimRow_(sh, L, r, v) {
  const range = sh.getRange(r, 1, 1, L.width);
  const vals = range.getValues()[0];
  const set = function (c, x) { if (c && x !== undefined) vals[c - 1] = x; };
  set(L.date, v.date); set(L.name, v.name); set(L.phone, v.phone); set(L.via, v.via);
  set(L.type, v.type); set(L.supervisor, v.supervisor); set(L.enteredBy, v.enteredBy); set(L.enteredAt, v.enteredAt); set(L.edited, v.edited); set(L.note, v.note);
  if (L.phone) sh.getRange(r, L.phone).setNumberFormat('@');
  if (L.enteredAt) sh.getRange(r, L.enteredAt).setNumberFormat('@');
  range.setValues([vals]);
  if (L.date && v.date) sh.getRange(r, L.date).setNumberFormat('dd/mm/yyyy');
  SpreadsheetApp.flush();
}

function canEdit_(u, rec) {
  if (u.role === 'admin') return true;
  if (u.role !== 'entry') return false;
  if (!rec.enteredAt || rec.enteredBy !== (u.name || u.username)) return false;
  try {
    const t = Utilities.parseDate(rec.enteredAt, CFG.TZ, 'dd/MM/yyyy HH:mm').getTime();
    return Date.now() - t <= CFG.EDIT_HOURS * 3600 * 1000;
  } catch (e) { return false; }
}

/** Data entry: their own entries from the last 24 hours (editable). Admin: everything they entered recently. */
function listMyEntries(token) {
  const u = auth_(token, ['admin', 'entry']);
  const recent = recentClaims_();
  const me = u.name || u.username;
  return recent.rows.filter(function (r) { return r.enteredBy === me && r.enteredAt && canEdit_(u, r); })
    .sort(function (a, b) { return stampMs_(b.enteredAt) - stampMs_(a.enteredAt); })
    .slice(0, 60).map(function (r) { return claimView_(r, recent.today); });
}

function stampMs_(s) { if (!s) return 0; try { return Utilities.parseDate(s, CFG.TZ, 'dd/MM/yyyy HH:mm').getTime(); } catch (e) { return 0; } }

/**
 * Fix a registration. data = { catKey, row, expectName, expectPhone, category, name, phones, agentId (optional) }.
 * The registration date is never changed, so editing cannot restart the 30-day lock.
 */
function editRecord(token, data) {
  const u = auth_(token, ['admin', 'entry']);
  const fromCat = CATEGORIES.filter(function (c) { return c.key === data.catKey; })[0];
  const toCat = CATEGORIES.filter(function (c) { return c.key === (data.category || data.catKey); })[0];
  if (!fromCat || !toCat) throw new Error('Choose a category.');
  const name = String(data.name || '').trim().replace(/\s+/g, ' ');
  if (!name) throw new Error('Customer name is required.');
  if (name.length > 80) throw new Error('Customer name is too long.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet(), ssTz = ss.getSpreadsheetTimeZone();
    const sh = categorySheet_(fromCat, false);
    if (!sh) throw new Error('Tab not found.');
    const L = layout_(sh, true);
    const row = Number(data.row);
    if (!(row > L.headerRow)) throw new Error('Record not found.');
    const cur = sh.getRange(row, 1, 1, L.width).getValues()[0];
    const g = function (c) { return c ? cur[c - 1] : ''; };
    const rec = { name: clean_(g(L.name)), phone: cellStr_(g(L.phone)), via: clean_(g(L.via)), type: clean_(g(L.type)), supervisor: clean_(g(L.supervisor)),
      enteredBy: clean_(g(L.enteredBy)), enteredAt: cellStr_(g(L.enteredAt)), date: g(L.date), note: clean_(g(L.note)) };
    if (rec.name !== String(data.expectName || '') || rec.phone !== String(data.expectPhone || '')) {
      throw new Error('This record was changed or moved in the sheet. Refresh the list and try again.');
    }
    if (!canEdit_(u, rec)) throw new Error('You can only edit your own entries from the last ' + CFG.EDIT_HOURS + ' hours. Ask an admin to change this one.');

    // phones: re-validate only if they were changed
    const phonesIn = (Array.isArray(data.phones) ? data.phones : [data.phones]).map(function (x) { return String(x || '').trim(); }).filter(String);
    const unchanged = phonesIn.join(' / ') === rec.phone || (phonesIn.length === 1 && phonesIn[0] === rec.phone);
    let phone = rec.phone, items = rowNumbers_(rec.phone);
    if (!unchanged) { const ph = parsePhones_(phonesIn); phone = ph.formatted; items = ph.items; }

    let via = rec.via, type = rec.type, supervisor = rec.supervisor;
    if (data.agentId) { const a = agentFields_(data.agentId); via = a.via; type = a.agent.type; supervisor = a.supervisor; }

    const recent = recentClaims_();
    const today = recent.today;
    if (!unchanged || data.agentId || name !== rec.name) {
      const hit = firstLock_(items, recent.rows, today, { category: fromCat.key, row: row }, via.replace(/\(.*?\)/g, '').trim(), name);
      if (hit) return { ok: false, lock: claimView_(hit, today) };
    }
    const edited = (u.name || u.username) + ', ' + nowStr_();
    let outCat = fromCat, outRow = row;
    if (toCat.key === fromCat.key) {
      writeClaimRow_(sh, L, row, { name: name, phone: phone, via: via, type: type, supervisor: supervisor, edited: edited });
    } else {
      // move to the other tab, keeping the original date and who entered it; the old row is blanked
      const sh2 = categorySheet_(toCat, true), L2 = layout_(sh2, true);
      const r2 = lastDataRow_(sh2, L2) + 1;
      if (r2 > sh2.getMaxRows()) sh2.insertRowsAfter(sh2.getMaxRows(), 50);
      writeClaimRow_(sh2, L2, r2, { date: rec.date, name: name, phone: phone, via: via, type: type, supervisor: supervisor,
        enteredBy: rec.enteredBy, enteredAt: rec.enteredAt, edited: edited + ' (moved from ' + fromCat.label + ')', note: rec.note });
      [L.date, L.name, L.phone, L.via, L.type, L.supervisor, L.enteredBy, L.enteredAt, L.edited, L.note].forEach(function (c) { if (c) sh.getRange(row, c).clearContent(); });
      outCat = toCat; outRow = r2;
    }
    CacheService.getScriptCache().removeAll(['idx:meta', 'hist:meta']);
    const fresh = recentClaims_();
    const updated = fresh.rows.filter(function (r) { return r.category === outCat.key && r.row === outRow; })[0];
    return { ok: true, record: updated ? claimView_(updated, fresh.today) : null };
  } finally {
    lock.releaseLock();
  }
}

/* ───────────────────────── Records (viewer/admin) ───────────────────────── */

function listRecords(token, f) {
  const me = auth_(token, ['admin', 'viewer']);
  f = f || {};
  const days = Number(f.days || 0);
  const all = days > 0 && days <= 31 ? recentClaims_() : readAllClaims_();   // recent periods: fast path
  const today = all.today;
  const q = String(f.q || '').trim().toLowerCase();
  const qDigits = q.replace(/\D/g, '');
  const qLocal = qDigits.length >= 4 ? (qDigits.indexOf('251') === 0 ? '0' + qDigits.slice(3) : /^[79]/.test(qDigits) ? '0' + qDigits : qDigits) : '';
  let rows = all.rows.filter(function (r) {
    if (f.category && f.category !== 'all' && r.category !== f.category) return false;
    if (days && (r.day === null || today - r.day >= days)) return false;
    if (f.lockedOnly && (r.day === null || today - r.day > CFG.LOCK_DAYS)) return false;
    if (!q) return true;
    if (qLocal) {
      const pd = String(r.phone).replace(/\D/g, '');
      if (pd.indexOf(qDigits) >= 0 || pd.indexOf(qLocal) >= 0 || r.keys.some(function (k) { return qLocal.indexOf(k) === 0 || k.replace('+', '').indexOf(qDigits) === 0; })) return true;
    }
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
  return { canEdit: me.role === 'admin', total: total, shown: Math.min(total, CFG.MAX_RECORDS), stats: stats, rows: rows.slice(0, CFG.MAX_RECORDS).map(function (r) { return claimView_(r, today); }) };
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
