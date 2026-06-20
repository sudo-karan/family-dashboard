/**
 * Family FD Tracker — Google Apps Script backend.
 *
 * Paste this into the script editor of the Google Sheet that holds the data
 * (Extensions -> Apps Script), set the FAMILY_PASSWORD script property, then
 * deploy as a Web App (execute as Me, access: Anyone with the link).
 * Full steps: see SETUP.md.
 *
 * The sheet is the source of truth. Two tabs:
 *   FDs      — one row per deposit
 *   Accounts — one row per account (name, person, bank)
 *
 * Columns are matched by HEADER NAME, not position, so the family can reorder
 * or add columns without breaking anything. Both the friendly headers used by
 * the family's original workbook ("FD Account Number", "Start Date",
 * "Interest %", ...) and the camelCase field names ("accountNumber",
 * "startDate", "rate", ...) are recognised — see HEADER_ALIASES below.
 */

var FD_SHEET = 'FDs';
var ACCOUNTS_SHEET = 'Accounts';
// Auto-created change log. Don't edit it by hand: the Before/After JSON
// columns are what make Undo work.
var LOG_SHEET = 'Log';
var LOG_HEADERS = ['Id', 'Time', 'Device', 'Action', 'Target', 'Summary', 'Before', 'After', 'Undone'];
// Optional device name sent with each mutation ("Dad's phone"); set per request.
var CURRENT_DEVICE = '';

// Canonical field order used when talking JSON to the frontend.
var FD_FIELDS = ['id', 'account', 'accountNumber', 'startDate', 'endDate',
  'tenureDays', 'rate', 'principal', 'maturity', 'interestOnMaturity',
  'compounding', 'status', 'showInDashboard', 'remarks'];
var ACCOUNT_FIELDS = ['name', 'person', 'bank'];

// Normalised sheet header -> canonical field. Headers are compared after
// trimming, lower-casing and collapsing whitespace.
var HEADER_ALIASES = {
  'id': 'id', 's.no': 'id', 'sno': 'id', 's no': 'id', 'sr no': 'id', 'sr.no': 'id',
  'account': 'account', 'account name': 'account',
  'accountnumber': 'accountNumber', 'account number': 'accountNumber',
  'fd account number': 'accountNumber', 'fd account no': 'accountNumber',
  'fd no': 'accountNumber', 'fd number': 'accountNumber',
  'startdate': 'startDate', 'start date': 'startDate',
  'enddate': 'endDate', 'end date': 'endDate', 'maturity date': 'endDate',
  'tenuredays': 'tenureDays', 'tenure (days)': 'tenureDays',
  'tenure days': 'tenureDays', 'tenure': 'tenureDays',
  'rate': 'rate', 'interest %': 'rate', 'interest%': 'rate', 'interest': 'rate',
  'interest ()': 'rate', 'interest rate': 'rate', 'rate %': 'rate',
  'principal': 'principal', 'principal amount': 'principal',
  'maturity': 'maturity', 'maturity amount': 'maturity',
  'interestonmaturity': 'interestOnMaturity', 'interest on maturity': 'interestOnMaturity',
  'compounding': 'compounding',
  'status': 'status',
  'showindashboard': 'showInDashboard', 'show in dashboard': 'showInDashboard',
  'remarks': 'remarks', 'remark': 'remarks', 'notes': 'remarks',
  // Accounts tab
  'name': 'name', 'person': 'person', 'bank': 'bank', 'bank name': 'bank'
};

// Fields that must be stored as plain text so Sheets never mangles them
// (huge FD numbers -> scientific notation, ISO dates -> locale date serials).
var TEXT_FIELDS = { accountNumber: true, startDate: true, endDate: true };
var BOOL_FIELDS = { interestOnMaturity: true, showInDashboard: true };
var NUMBER_FIELDS = { id: true, tenureDays: true, rate: true, principal: true, maturity: true };

function doGet() {
  return jsonOut({
    ok: true,
    app: 'fd-tracker',
    hint: "POST JSON like {\"action\":\"list\",\"password\":\"...\"} to use the API."
  });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return jsonOut({ ok: false, error: 'busy: another request is writing — try again' });
  }
  try {
    var req;
    try {
      req = JSON.parse(e && e.postData && e.postData.contents ? e.postData.contents : '');
    } catch (parseErr) {
      return jsonOut({ ok: false, error: 'bad request: body must be JSON' });
    }

    var stored = PropertiesService.getScriptProperties().getProperty('FAMILY_PASSWORD');
    if (!stored) {
      return jsonOut({ ok: false, error: 'setup incomplete: FAMILY_PASSWORD script property is not set' });
    }
    if (!req.password || String(req.password) !== stored) {
      return jsonOut({ ok: false, error: 'unauthorized' });
    }

    CURRENT_DEVICE = String(req.device || '').trim().substring(0, 40);

    var action = String(req.action || '');
    if (action === 'list') {
      // fall through: every action returns the full fresh state
    } else if (action === 'create') {
      upsertAccount(req.newAccount);
      createFd(req.fd || {});
    } else if (action === 'update') {
      upsertAccount(req.newAccount);
      updateFd(req.fd || {});
    } else if (action === 'delete') {
      deleteFd(req.id);
    } else if (action === 'accountCreate') {
      accountCreate(req.account || {});
    } else if (action === 'accountUpdate') {
      accountUpdate(req.account || {}, req.originalName);
    } else if (action === 'accountDelete') {
      accountDelete(req.name);
    } else if (action === 'transferFds') {
      transferFds(req.ids, req.to, req.from);
    } else if (action === 'undo') {
      undoLog(req.logId);
    } else if (action === 'subscribe') {
      subscribePush(req.subscription, CURRENT_DEVICE);
      return jsonOut({ ok: true });
    } else if (action === 'unsubscribe') {
      unsubscribePush(req.endpoint);
      return jsonOut({ ok: true });
    } else {
      return jsonOut({ ok: false, error: 'unknown action: ' + action });
    }

    // Commit buffered sheet writes before reading the state back — without
    // this, the read in the same request can still see pre-change data
    // (classic Apps Script behaviour, most visible after deleteRow).
    SpreadsheetApp.flush();
    return jsonOut({
      ok: true, fds: readFds(), accounts: readAccounts(), log: readLog(15),
      vapidPublic: PropertiesService.getScriptProperties().getProperty('VAPID_PUBLIC') || ''
    });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

function jsonOut(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------- sheet I/O

function getSheet(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error("sheet '" + name + "' not found — check the tab name");
  return sh;
}

function normalizeHeader(h) {
  return String(h == null ? '' : h).trim().toLowerCase().replace(/\s+/g, ' ');
}

/** Map canonical field -> 0-based column index, by reading the header row. */
function headerMap(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) throw new Error("sheet '" + sheet.getName() + "' has no header row");
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var field = HEADER_ALIASES[normalizeHeader(headers[i])];
    if (field && !(field in map)) map[field] = i;
  }
  return map;
}

function toIsoDate(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return String(v).trim();
}

/** Opaque strings (FD account numbers): never render as scientific notation. */
function toOpaqueString(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) return toIsoDate(v);
  if (typeof v === 'number') {
    if (isFinite(v) && Math.floor(v) === v) return v.toFixed(0);
    var s = String(v);
    return (s.indexOf('e') >= 0 || s.indexOf('E') >= 0) ? v.toFixed(0) : s;
  }
  return String(v).trim();
}

function toBool(v) {
  return v === true || /^true$/i.test(String(v).trim());
}

function toNumber(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v).replace(/[,₹\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

function coerceField(field, v) {
  if (field === 'accountNumber') return toOpaqueString(v);
  if (field === 'startDate' || field === 'endDate') return toIsoDate(v);
  if (BOOL_FIELDS[field]) return toBool(v);
  if (NUMBER_FIELDS[field]) return toNumber(v);
  return v === null || v === undefined ? '' : String(v).trim();
}

function readFds() {
  var sheet = getSheet(FD_SHEET);
  var map = headerMap(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var fds = [];
  var maxId = 0;
  var r, fd, i, field;
  for (r = 0; r < values.length; r++) {
    var row = values[r];
    var blank = true;
    for (i = 0; i < row.length; i++) {
      if (row[i] !== '' && row[i] !== null) { blank = false; break; }
    }
    if (blank) continue;
    fd = { _row: r + 2 };
    for (i = 0; i < FD_FIELDS.length; i++) {
      field = FD_FIELDS[i];
      fd[field] = (field in map) ? coerceField(field, row[map[field]]) : coerceField(field, '');
    }
    if (fd.id > maxId) maxId = fd.id;
    fds.push(fd);
  }
  // Self-heal: rows added by hand in the sheet may lack an ID — assign one so
  // the app can edit/delete them.
  if ('id' in map) {
    for (r = 0; r < fds.length; r++) {
      if (!fds[r].id) {
        maxId += 1;
        fds[r].id = maxId;
        sheet.getRange(fds[r]._row, map.id + 1).setValue(maxId);
      }
    }
  }
  for (r = 0; r < fds.length; r++) delete fds[r]._row;
  return fds;
}

function readAccounts() {
  var sheet = getSheet(ACCOUNTS_SHEET);
  var map = headerMap(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var out = [];
  for (var r = 0; r < values.length; r++) {
    var acc = {};
    for (var i = 0; i < ACCOUNT_FIELDS.length; i++) {
      var field = ACCOUNT_FIELDS[i];
      acc[field] = (field in map) ? String(values[r][map[field]] == null ? '' : values[r][map[field]]).trim() : '';
    }
    if (acc.name) out.push(acc);
  }
  return out;
}

/** Build a full sheet row (array) for an FD using the sheet's own column layout. */
function fdToRow(fd, map, width) {
  var row = [];
  for (var c = 0; c < width; c++) row.push('');
  for (var i = 0; i < FD_FIELDS.length; i++) {
    var field = FD_FIELDS[i];
    if (!(field in map)) continue;
    var v = fd[field];
    if (BOOL_FIELDS[field]) {
      row[map[field]] = toBool(v);
    } else if (NUMBER_FIELDS[field]) {
      row[map[field]] = (v === '' || v === null || v === undefined) ? '' : toNumber(v);
    } else if (field === 'accountNumber') {
      row[map[field]] = toOpaqueString(v);
    } else {
      row[map[field]] = v === null || v === undefined ? '' : String(v);
    }
  }
  return row;
}

/** Force plain-text format on opaque columns of one row, then write values. */
function writeFdRow(sheet, map, rowIndex, row) {
  for (var field in TEXT_FIELDS) {
    if (field in map) sheet.getRange(rowIndex, map[field] + 1).setNumberFormat('@');
  }
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
}

function nextId(fds) {
  var max = 0;
  for (var i = 0; i < fds.length; i++) {
    if (fds[i].id > max) max = fds[i].id;
  }
  return max + 1;
}

function findFdById(id) {
  var fds = readFds();
  var want = toNumber(id);
  for (var i = 0; i < fds.length; i++) {
    if (fds[i].id === want) return fds[i];
  }
  return null;
}

/* Raw row operations (no logging) — used by the wired mutations and by undo. */
function insertFdRow(fd) {
  var sheet = getSheet(FD_SHEET);
  var map = headerMap(sheet);
  var row = fdToRow(fd, map, sheet.getLastColumn());
  writeFdRow(sheet, map, sheet.getLastRow() + 1, row);
}

function rawUpdateFd(fd) {
  var sheet = getSheet(FD_SHEET);
  var map = headerMap(sheet);
  var rowIndex = findRowById(sheet, map, fd.id);
  if (!rowIndex) throw new Error('FD with id ' + fd.id + ' not found');
  writeFdRow(sheet, map, rowIndex, fdToRow(fd, map, sheet.getLastColumn()));
}

function rawDeleteFd(id) {
  var sheet = getSheet(FD_SHEET);
  var map = headerMap(sheet);
  var rowIndex = findRowById(sheet, map, id);
  if (!rowIndex) throw new Error('FD with id ' + id + ' not found');
  sheet.deleteRow(rowIndex);
}

function createFd(fd) {
  fd = applyDefaults(fd);
  fd.id = nextId(readFds());
  insertFdRow(fd);
  appendLog('create', 'fd:' + fd.id,
    'Added FD #' + fd.id + ' · ' + fd.account + moneyBit(fd), null, pickFd(fd));
}

function updateFd(fd) {
  fd = applyDefaults(fd);
  var before = findFdById(fd.id);
  if (!before) throw new Error('FD with id ' + fd.id + ' not found');
  rawUpdateFd(fd);
  appendLog('update', 'fd:' + before.id,
    'Edited FD #' + before.id + ' · ' + fd.account, before, pickFd(fd));
}

function deleteFd(id) {
  var before = findFdById(id);
  if (!before) throw new Error('FD with id ' + id + ' not found');
  rawDeleteFd(id);
  appendLog('delete', 'fd:' + before.id,
    'Deleted FD #' + before.id + ' · ' + before.account + moneyBit(before), before, null);
}

function findRowById(sheet, map, id) {
  if (!('id' in map)) throw new Error("the FDs sheet has no ID column");
  var want = toNumber(id);
  if (!want) throw new Error('missing id');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var ids = sheet.getRange(2, map.id + 1, lastRow - 1, 1).getValues();
  for (var r = 0; r < ids.length; r++) {
    if (toNumber(ids[r][0]) === want) return r + 2;
  }
  return 0;
}

function applyDefaults(fd) {
  if (fd.compounding === undefined || fd.compounding === '') fd.compounding = 'quarterly';
  if (fd.status === undefined || fd.status === '') fd.status = 'Active';
  if (fd.interestOnMaturity === undefined) fd.interestOnMaturity = true;
  if (fd.showInDashboard === undefined) fd.showInDashboard = true;
  return fd;
}

/** Silent upsert used when an FD is saved with an inline-added account. */
function upsertAccount(acc) {
  if (!acc || !acc.name || !String(acc.name).trim()) return;
  var name = String(acc.name).trim();
  var sheet = getSheet(ACCOUNTS_SHEET);
  var map = headerMap(sheet);
  if (findAccountRow(sheet, map, name)) return;
  appendAccountRow(sheet, map, acc);
  appendLog('accountCreate', accountTarget(name), 'Added account "' + name + '"', null, pickAccount(acc));
}

function accountTarget(name) {
  return 'account:' + String(name || '').trim().toLowerCase();
}

function pickAccount(acc) {
  return {
    name: String(acc.name || '').trim(),
    person: String(acc.person || '').trim(),
    bank: String(acc.bank || '').trim()
  };
}

function findAccountByName(name) {
  var want = String(name || '').trim().toLowerCase();
  var accounts = readAccounts();
  for (var i = 0; i < accounts.length; i++) {
    if (accounts[i].name.toLowerCase() === want) return accounts[i];
  }
  return null;
}

/** 1-based row index of the account with this name (case-insensitive), or 0. */
function findAccountRow(sheet, map, name) {
  if (!('name' in map)) throw new Error("the Accounts sheet has no Name column");
  var want = String(name == null ? '' : name).trim().toLowerCase();
  if (!want) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var vals = sheet.getRange(2, map.name + 1, lastRow - 1, 1).getValues();
  for (var r = 0; r < vals.length; r++) {
    if (String(vals[r][0]).trim().toLowerCase() === want) return r + 2;
  }
  return 0;
}

function accountRowValues(sheet, map, acc) {
  var row = [];
  for (var c = 0; c < sheet.getLastColumn(); c++) row.push('');
  if ('name' in map) row[map.name] = String(acc.name || '').trim();
  if ('person' in map) row[map.person] = String(acc.person || '').trim();
  if ('bank' in map) row[map.bank] = String(acc.bank || '').trim();
  return row;
}

function appendAccountRow(sheet, map, acc) {
  var row = accountRowValues(sheet, map, acc);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

function accountCreate(acc) {
  var name = String(acc.name || '').trim();
  if (!name) throw new Error('account name is required');
  var sheet = getSheet(ACCOUNTS_SHEET);
  var map = headerMap(sheet);
  if (findAccountRow(sheet, map, name)) {
    throw new Error('an account named "' + name + '" already exists');
  }
  appendAccountRow(sheet, map, acc);
  appendLog('accountCreate', accountTarget(name), 'Added account "' + name + '"', null, pickAccount(acc));
}

function accountUpdate(acc, originalName) {
  var before = findAccountByName(originalName);
  if (!before) throw new Error('account "' + originalName + '" not found');
  rawAccountUpdate(acc, originalName);
  var after = pickAccount(acc);
  var summary = before.name !== after.name
    ? 'Renamed account "' + before.name + '" → "' + after.name + '"'
    : 'Edited account "' + after.name + '"';
  appendLog('accountUpdate', accountTarget(after.name), summary, before, after);
}

function rawAccountUpdate(acc, originalName) {
  var orig = String(originalName || '').trim();
  var name = String(acc.name || '').trim();
  if (!name) throw new Error('account name is required');
  var sheet = getSheet(ACCOUNTS_SHEET);
  var map = headerMap(sheet);
  var rowIndex = findAccountRow(sheet, map, orig);
  if (!rowIndex) throw new Error('account "' + orig + '" not found');
  var dup = findAccountRow(sheet, map, name);
  if (dup && dup !== rowIndex) {
    throw new Error('an account named "' + name + '" already exists');
  }
  var row = accountRowValues(sheet, map, acc);
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  // Renaming must not orphan deposits: cascade to the FDs tab.
  if (name !== orig) renameFdAccounts(orig, name);
}

function renameFdAccounts(oldName, newName) {
  var sheet = getSheet(FD_SHEET);
  var map = headerMap(sheet);
  if (!('account' in map)) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var rng = sheet.getRange(2, map.account + 1, lastRow - 1, 1);
  var vals = rng.getValues();
  var want = String(oldName || '').trim().toLowerCase();
  var changed = false;
  for (var r = 0; r < vals.length; r++) {
    if (String(vals[r][0]).trim().toLowerCase() === want) {
      vals[r][0] = newName;
      changed = true;
    }
  }
  if (changed) rng.setValues(vals);
}

function accountDelete(name) {
  var before = findAccountByName(name);
  if (!before) throw new Error('account "' + name + '" not found');
  rawAccountDelete(name);
  appendLog('accountDelete', accountTarget(before.name), 'Deleted account "' + before.name + '"', before, null);
}

/** Move the chosen FDs (by id) onto another account — reassigns `account`. */
function transferFds(ids, to, from) {
  to = String(to || '').trim();
  if (!to) throw new Error('choose a destination account');
  if (!findAccountByName(to)) throw new Error('account "' + to + '" not found');
  if (!ids || !ids.length) throw new Error('select at least one FD to move');
  var want = {};
  for (var k = 0; k < ids.length; k++) want[toNumber(ids[k])] = true;
  var fds = readFds();
  var moved = [];
  var srcSet = {};
  for (var i = 0; i < fds.length; i++) {
    var f = fds[i];
    if (want[f.id] && String(f.account).trim().toLowerCase() !== to.toLowerCase()) {
      srcSet[f.account] = true;
      rawUpdateFd(applyDefaults(Object.assign({}, f, { account: to })));
      moved.push(f.id);
    }
  }
  if (moved.length === 0) throw new Error('nothing to move');
  var srcKeys = Object.keys(srcSet);
  var src = String(from || '').trim() || srcKeys.join(', ');
  var payload = { from: src, to: to, ids: moved };
  appendLog('transfer', 'account:' + String(srcKeys[0] || src).toLowerCase(),
    'Moved ' + moved.length + ' FD' + (moved.length === 1 ? '' : 's') + ' from "' + src + '" to "' + to + '"',
    payload, payload);
}

function rawAccountDelete(name) {
  var want = String(name || '').trim().toLowerCase();
  if (!want) throw new Error('missing account name');
  var fds = readFds();
  var used = 0;
  for (var i = 0; i < fds.length; i++) {
    if (String(fds[i].account).trim().toLowerCase() === want) used += 1;
  }
  if (used > 0) {
    throw new Error('cannot delete "' + name + '": ' + used + ' FD(s) still use this account');
  }
  var sheet = getSheet(ACCOUNTS_SHEET);
  var map = headerMap(sheet);
  var rowIndex = findAccountRow(sheet, map, name);
  if (!rowIndex) throw new Error('account "' + name + '" not found');
  sheet.deleteRow(rowIndex);
}

// ---------------------------------------------------------------- change log

function ensureLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  }
  return sheet;
}

function moneyBit(fd) {
  return fd.principal ? ' · ' + inr(fd.principal) : '';
}

function inr(n) {
  try {
    return '₹' + Number(n).toLocaleString('en-IN', { maximumFractionDigits: 2 });
  } catch (e) {
    return '₹' + n;
  }
}

function pickFd(fd) {
  var out = {};
  for (var i = 0; i < FD_FIELDS.length; i++) {
    var k = FD_FIELDS[i];
    out[k] = fd[k] === undefined ? '' : fd[k];
  }
  return out;
}

/** Append one change to the Log tab. Must never make the actual save fail. */
function appendLog(action, target, summary, before, after) {
  try {
    var sheet = ensureLogSheet();
    var lastRow = sheet.getLastRow();
    var id = lastRow >= 2 ? toNumber(sheet.getRange(lastRow, 1).getValue()) + 1 : 1;
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var time = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
    sheet.getRange(lastRow + 1, 2).setNumberFormat('@'); // keep Time as plain text
    sheet.getRange(lastRow + 1, 1, 1, LOG_HEADERS.length).setValues([[
      id, time, CURRENT_DEVICE, action, target, summary,
      before ? JSON.stringify(before) : '', after ? JSON.stringify(after) : '', false
    ]]);
  } catch (err) {
    // logging is best-effort by design
  }
}

/** Newest-first recent entries with canUndo computed: an entry is undoable
 *  only while it is the latest change to its target (never clobber a newer
 *  edit), is not an undo itself, and has not already been undone. */
function parseJsonSafe(s) {
  if (!s) return null;
  try { return JSON.parse(String(s)); } catch (e) { return null; }
}

function readLog(limit) {
  var sheet;
  try { sheet = ensureLogSheet(); } catch (e) { return []; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var n = Math.min(lastRow - 1, 50);
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var values = sheet.getRange(lastRow - n + 1, 1, n, LOG_HEADERS.length).getValues();
  var entries = [];
  for (var r = values.length - 1; r >= 0; r--) { // newest first
    var v = values[r];
    entries.push({
      id: toNumber(v[0]),
      time: v[1] instanceof Date ? Utilities.formatDate(v[1], tz, 'yyyy-MM-dd HH:mm') : String(v[1] || ''),
      device: String(v[2] == null ? '' : v[2]),
      action: String(v[3] || ''),
      target: String(v[4] || ''),
      summary: String(v[5] || ''),
      // full snapshots so the app can show exactly what was recorded
      before: parseJsonSafe(v[6]),
      after: parseJsonSafe(v[7]),
      undone: toBool(v[8])
    });
  }
  for (var i = 0; i < entries.length; i++) {
    var e = entries[i];
    // transfers move many rows and are reversible by another transfer, not undo
    var blocked = e.undone || e.action === 'undo' || e.action === 'transfer';
    for (var j = 0; j < i && !blocked; j++) { // entries[j] are newer
      if (entries[j].target === e.target) blocked = true;
    }
    e.canUndo = !blocked;
  }
  return entries.slice(0, limit || 15);
}

/** Revert one logged change using its stored Before/After images. */
function undoLog(logId) {
  var sheet = ensureLogSheet();
  var lastRow = sheet.getLastRow();
  var want = toNumber(logId);
  if (!want || lastRow < 2) throw new Error('log entry not found');
  var values = sheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues();
  var idx = -1;
  for (var r = 0; r < values.length; r++) {
    if (toNumber(values[r][0]) === want) { idx = r; break; }
  }
  if (idx < 0) throw new Error('log entry ' + logId + ' not found');
  var action = String(values[idx][3] || '');
  var target = String(values[idx][4] || '');
  var summary = String(values[idx][5] || '');
  if (toBool(values[idx][8])) throw new Error('this change was already undone');
  if (action === 'undo') throw new Error('an undo cannot be undone');
  for (var r2 = idx + 1; r2 < values.length; r2++) {
    if (String(values[r2][4] || '') === target) {
      throw new Error('cannot undo: this item was changed again afterwards');
    }
  }
  var before = null, after = null;
  try {
    before = values[idx][6] ? JSON.parse(String(values[idx][6])) : null;
    after = values[idx][7] ? JSON.parse(String(values[idx][7])) : null;
  } catch (e) {
    throw new Error('log entry is unreadable');
  }

  if (action === 'create') {
    rawDeleteFd(after.id);
  } else if (action === 'update') {
    rawUpdateFd(before);
  } else if (action === 'delete') {
    if (findFdById(before.id)) throw new Error('cannot undo: FD id ' + before.id + ' exists again');
    insertFdRow(before);
  } else if (action === 'accountCreate') {
    rawAccountDelete(after.name); // refuses if FDs were attached meanwhile
  } else if (action === 'accountUpdate') {
    rawAccountUpdate(before, after.name); // rename cascades back to the FDs
  } else if (action === 'accountDelete') {
    var s = getSheet(ACCOUNTS_SHEET);
    if (findAccountByName(before.name)) throw new Error('cannot undo: account "' + before.name + '" exists again');
    appendAccountRow(s, headerMap(s), before);
  } else {
    throw new Error('cannot undo action: ' + action);
  }

  sheet.getRange(idx + 2, 9).setValue(true); // Undone column
  appendLog('undo', target, 'Undid: ' + summary, after, before);
}

// ===================================================================
//  Push reminders (Web Push to the installed PWA)
//
//  Sends a payload-less push (no FD details on the lock screen — privacy)
//  when an Active FD matures in 2 days, 1 day, or today, IST. Driven by a
//  daily time trigger around 10:00 in the project's timezone (set it to
//  Asia/Kolkata — see SETUP.md). VAPID JWTs are signed with the pure ES5
//  ES256 implementation below (Apps Script has no native ECDSA and no BigInt);
//  the same algorithm is validated against Node's crypto in the repo's tests.
// ===================================================================

var PUSH_SHEET = 'PushSubs';
var PUSH_HEADERS = ['Endpoint', 'P256dh', 'Auth', 'Device', 'Added'];

/* ---- byte helpers (Apps Script Byte[] are signed -128..127) ---- */
function toSigned(arr) { return arr.map(function (v) { return v > 127 ? v - 256 : v; }); }
function toUnsigned(arr) { return arr.map(function (v) { return v < 0 ? v + 256 : v; }); }
function strBytes(s) { return toUnsigned(Utilities.newBlob(s).getBytes()); }
function gsSha256(u) { return toUnsigned(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, toSigned(u))); }
function gsHmac(keyU, msgU) { return toUnsigned(Utilities.computeHmacSha256Signature(toSigned(msgU), toSigned(keyU))); }
function b64urlEncodeBytes(u) { return Utilities.base64EncodeWebSafe(toSigned(u)).replace(/=+$/, ''); }
function b64urlDecodeBytes(s) { return toUnsigned(Utilities.base64DecodeWebSafe(s)); }

/* ---- pure ES5 P-256 ECDSA (ES256) — Apps Script has no BigInt; big ints
 * are 16-bit limb arrays. Validated against Node crypto in es256_es5.test.js. ---- */
var ES256 = (function () {
  // limbs: little-endian 16-bit. value = sum limb[i] * 65536^i
  function norm(a) { while (a.length > 1 && a[a.length - 1] === 0) a.pop(); return a; }
  function isZero(a) { return a.length === 1 && a[0] === 0; }
  function cmp(a, b) {
    if (a.length !== b.length) return a.length < b.length ? -1 : 1;
    for (var i = a.length - 1; i >= 0; i--) if (a[i] !== b[i]) return a[i] < b[i] ? -1 : 1;
    return 0;
  }
  function add(a, b) {
    var r = [], c = 0, n = Math.max(a.length, b.length);
    for (var i = 0; i < n; i++) { var s = (a[i] || 0) + (b[i] || 0) + c; r[i] = s & 0xffff; c = s >>> 16; }
    if (c) r[i] = c;
    return r;
  }
  function sub(a, b) { // assumes a >= b
    var r = [], br = 0;
    for (var i = 0; i < a.length; i++) {
      var s = a[i] - (b[i] || 0) - br;
      if (s < 0) { s += 0x10000; br = 1; } else br = 0;
      r[i] = s;
    }
    return norm(r);
  }
  function mul(a, b) {
    var r = []; for (var i = 0; i < a.length + b.length; i++) r[i] = 0;
    for (i = 0; i < a.length; i++) {
      var c = 0, ai = a[i];
      for (var j = 0; j < b.length; j++) { var s = ai * b[j] + r[i + j] + c; r[i + j] = s & 0xffff; c = Math.floor(s / 0x10000); }
      var k = i + b.length;
      while (c) { var s2 = (r[k] || 0) + c; r[k] = s2 & 0xffff; c = Math.floor(s2 / 0x10000); k++; }
    }
    return norm(r);
  }
  function shl1(a) {
    var r = [], c = 0;
    for (var i = 0; i < a.length; i++) { var v = (a[i] << 1) | c; r[i] = v & 0xffff; c = (v >>> 16) & 1; }
    if (c) r[i] = c;
    return r;
  }
  function bitLength(a) {
    var i = a.length - 1; while (i > 0 && a[i] === 0) i--;
    var w = a[i], b = 0; while (w) { w >>>= 1; b++; }
    return i * 16 + b;
  }
  function testBit(a, i) { var limb = i >> 4, off = i & 15; return ((a[limb] || 0) >> off) & 1; }

  function fromBytes(bytes) { // big-endian byte array -> limbs
    var r = [0], k = 0;
    for (var i = bytes.length - 1; i >= 0; i--) {
      var li = k >> 1;
      if (k & 1) r[li] = (r[li] || 0) | ((bytes[i] & 0xff) << 8);
      else r[li] = (bytes[i] & 0xff);
      k++;
    }
    return norm(r);
  }
  function to32Bytes(a) {
    var out = []; for (var i = 0; i < 32; i++) out[i] = 0;
    for (var li = 0; li < a.length; li++) {
      var lo = a[li] & 0xff, hi = (a[li] >> 8) & 0xff;
      var pos = 31 - li * 2;
      if (pos >= 0) out[pos] = lo;
      if (pos - 1 >= 0) out[pos - 1] = hi;
    }
    return out;
  }

  // curve constants as limb arrays
  var P = fromBytes(hexBytes('ffffffff00000001000000000000000000000000ffffffffffffffffffffffff'));
  var N = fromBytes(hexBytes('ffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551'));
  var Gx = fromBytes(hexBytes('6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296'));
  var Gy = fromBytes(hexBytes('4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5'));
  var P_M2 = sub(P, [2]);
  var N_M2 = sub(N, [2]);
  function hexBytes(h) { var b = []; for (var i = 0; i < h.length; i += 2) b.push(parseInt(h.substr(i, 2), 16)); return b; }

  // NIST P-256 fast reduction of a product (<= 512 bits) mod P
  function word(a, i) { return (a[2 * i] || 0) + (a[2 * i + 1] || 0) * 65536; } // 32-bit word i
  function fromWords(w) { // w[0]=most significant ... w[7]=least significant (8 words = 256 bits)
    var r = [];
    for (var pos = 0; pos < 8; pos++) {
      var val = w[7 - pos]; // limb position pos <-> word (7-pos)
      r[2 * pos] = val & 0xffff;
      r[2 * pos + 1] = Math.floor(val / 65536) & 0xffff;
    }
    return norm(r);
  }
  function reduceP(prod) {
    var c = [];
    for (var i = 0; i < 16; i++) c[i] = word(prod, i);
    var s1 = fromWords([c[7], c[6], c[5], c[4], c[3], c[2], c[1], c[0]]);
    var s2 = fromWords([c[15], c[14], c[13], c[12], c[11], 0, 0, 0]);
    var s3 = fromWords([0, c[15], c[14], c[13], c[12], 0, 0, 0]);
    var s4 = fromWords([c[15], c[14], 0, 0, 0, c[10], c[9], c[8]]);
    var s5 = fromWords([c[8], c[13], c[15], c[14], c[13], c[11], c[10], c[9]]);
    var s6 = fromWords([c[10], c[8], 0, 0, 0, c[13], c[12], c[11]]);
    var s7 = fromWords([c[11], c[9], 0, 0, c[15], c[14], c[13], c[12]]);
    var s8 = fromWords([c[12], 0, c[10], c[9], c[8], c[15], c[14], c[13]]);
    var s9 = fromWords([c[13], 0, c[11], c[10], c[9], 0, c[15], c[14]]);
    var t = s1;
    t = add(t, shl1(s2)); t = add(t, shl1(s3)); t = add(t, s4); t = add(t, s5);
    var negs = [s6, s7, s8, s9];
    for (var k = 0; k < 4; k++) { while (cmp(t, negs[k]) < 0) t = add(t, P); t = sub(t, negs[k]); }
    while (cmp(t, P) >= 0) t = sub(t, P);
    return norm(t);
  }

  // generic mod (binary long division), used for mod N
  function mod(a, m) {
    if (cmp(a, m) < 0) return a.slice();
    var r = [0], bl = bitLength(a);
    for (var i = bl - 1; i >= 0; i--) {
      r = shl1(r);
      if (testBit(a, i)) r[0] |= 1;
      if (cmp(r, m) >= 0) r = sub(r, m);
    }
    return norm(r);
  }

  // field (mod P) helpers
  function fAdd(a, b) { var t = add(a, b); if (cmp(t, P) >= 0) t = sub(t, P); return t; }
  function fSub(a, b) { return cmp(a, b) >= 0 ? sub(a, b) : sub(add(a, P), b); }
  function fMul(a, b) { return reduceP(mul(a, b)); }
  function modPow(base, exp, m, mulmod) {
    var r = [1], b = mod(base, m), bl = bitLength(exp);
    for (var i = 0; i < bl; i++) { if (testBit(exp, i)) r = mulmod(r, b); b = mulmod(b, b); }
    return r;
  }
  function fInv(a) { return modPow(a, P_M2, P, fMul); }
  function nMul(a, b) { return mod(mul(a, b), N); }
  function nInv(a) { return modPow(a, N_M2, N, nMul); }

  // Jacobian point ops mod P. Point = [X,Y,Z]; infinity has Z = [0].
  function jDouble(Pt) {
    var X1 = Pt[0], Y1 = Pt[1], Z1 = Pt[2];
    if (isZero(Y1) || isZero(Z1)) return [[1], [1], [0]];
    var A = fMul(Y1, Y1);
    var B = fMul(X1, A); B = fAdd(B, B); B = fAdd(B, B);          // 4*X1*A
    var C = fMul(A, A); C = fAdd(C, C); C = fAdd(C, C); C = fAdd(C, C); // 8*A^2
    var ZZ = fMul(Z1, Z1);
    // M = 3*(X1 - Z1^2)*(X1 + Z1^2)
    var M = fMul(fSub(X1, ZZ), fAdd(X1, ZZ)); M = fAdd(fAdd(M, M), M);
    var X3 = fSub(fSub(fMul(M, M), B), B);                       // M^2 - 2B
    var Y3 = fSub(fMul(M, fSub(B, X3)), C);                      // M*(B - X3) - C
    var Z3 = fMul(fAdd(Y1, Y1), Z1);                             // 2*Y1*Z1
    return [X3, Y3, Z3];
  }
  function jAdd(Pa, Pb) {
    if (isZero(Pa[2])) return Pb;
    if (isZero(Pb[2])) return Pa;
    var X1 = Pa[0], Y1 = Pa[1], Z1 = Pa[2], X2 = Pb[0], Y2 = Pb[1], Z2 = Pb[2];
    var Z1Z1 = fMul(Z1, Z1), Z2Z2 = fMul(Z2, Z2);
    var U1 = fMul(X1, Z2Z2), U2 = fMul(X2, Z1Z1);
    var S1 = fMul(fMul(Y1, Z2), Z2Z2), S2 = fMul(fMul(Y2, Z1), Z1Z1);
    if (cmp(U1, U2) === 0) { if (cmp(S1, S2) !== 0) return [[1], [1], [0]]; return jDouble(Pa); }
    var H = fSub(U2, U1);
    var I = fMul(fAdd(H, H), fAdd(H, H));
    var J = fMul(H, I);
    var r = fAdd(fSub(S2, S1), fSub(S2, S1));
    var V = fMul(U1, I);
    var X3 = fSub(fSub(fSub(fMul(r, r), J), V), V);
    var Y3 = fSub(fMul(r, fSub(V, X3)), fAdd(fMul(S1, J), fMul(S1, J)));
    var Z3 = fMul(fSub(fSub(fMul(fAdd(Z1, Z2), fAdd(Z1, Z2)), Z1Z1), Z2Z2), H);
    return [X3, Y3, Z3];
  }
  function jMul(k, Pt) {
    var R = [[1], [1], [0]], bl = bitLength(k);
    for (var i = bl - 1; i >= 0; i--) { R = jDouble(R); if (testBit(k, i)) R = jAdd(R, Pt); }
    return R;
  }
  function toAffine(Pt) {
    if (isZero(Pt[2])) return null;
    var zi = fInv(Pt[2]), zi2 = fMul(zi, zi), zi3 = fMul(zi2, zi);
    return [fMul(Pt[0], zi2), fMul(Pt[1], zi3)];
  }

  function bytesConcat(a, b) { return a.concat(b); }

  /* RFC 6979 deterministic k (hmac injected: hmac(keyBytes,msgBytes)->bytes) */
  function rfc6979k(hBytes, dBytes, hmac) {
    var z = mod(fromBytes(hBytes), N);
    var z2 = to32Bytes(z), x = to32Bytes(fromBytes(dBytes));
    var V = [], K = [];
    for (var i = 0; i < 32; i++) { V.push(1); K.push(0); }
    K = hmac(K, V.concat([0x00], x, z2)); V = hmac(K, V);
    K = hmac(K, V.concat([0x01], x, z2)); V = hmac(K, V);
    for (;;) {
      V = hmac(K, V);
      var k = fromBytes(V);
      if (!isZero(k) && cmp(k, N) < 0) return k;
      K = hmac(K, V.concat([0x00])); V = hmac(K, V);
    }
  }

  /* sign message bytes -> raw 64-byte signature r||s (low-S) */
  function signRaw(msgBytes, dBytes, sha256, hmac) {
    var hBytes = sha256(msgBytes);
    var z = mod(fromBytes(hBytes), N);
    var d = fromBytes(dBytes);
    for (;;) {
      var k = rfc6979k(hBytes, dBytes, hmac);
      var Raff = toAffine(jMul(k, [Gx, Gy, [1]]));
      var r = mod(Raff[0], N);
      if (isZero(r)) continue;
      var s = nMul(nInv(k), mod(add(z, mod(mul(r, d), N)), N));
      if (isZero(s)) continue;
      if (cmp(shl1(s), N) > 0) s = sub(N, s); // low-S: s > N/2  <=> 2s > N
      return to32Bytes(r).concat(to32Bytes(s));
    }
  }
  function publicKey(dBytes) {
    var Q = toAffine(jMul(fromBytes(dBytes), [Gx, Gy, [1]]));
    return [0x04].concat(to32Bytes(Q[0]), to32Bytes(Q[1]));
  }

  return { signRaw: signRaw, publicKey: publicKey,
    _internals: { fromBytes: fromBytes, to32Bytes: to32Bytes, fMul: fMul, mod: mod, fInv: fInv, nInv: nInv, mul: mul, P: P, N: N } };
})();

/** Run once in the editor; copy the two logged values into Script Properties. */
function generateVapidKeys() {
  var seed = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid() + String(Date.now());
  var d = gsSha256(strBytes(seed)); // 32 bytes, ~256-bit entropy
  var pub = ES256.publicKey(d);
  Logger.log('VAPID_PUBLIC = ' + b64urlEncodeBytes(pub));
  Logger.log('VAPID_PRIVATE = ' + b64urlEncodeBytes(d));
  Logger.log('Add both as Script Properties, plus VAPID_SUBJECT = mailto:you@example.com');
}

function vapidJwt(aud) {
  var props = PropertiesService.getScriptProperties();
  var privB64 = props.getProperty('VAPID_PRIVATE');
  var sub = props.getProperty('VAPID_SUBJECT') || 'mailto:fd-tracker@example.com';
  if (!privB64) throw new Error('VAPID keys not set');
  var header = b64urlEncodeBytes(strBytes(JSON.stringify({ typ: 'JWT', alg: 'ES256' })));
  var payload = b64urlEncodeBytes(strBytes(JSON.stringify({ aud: aud, exp: Math.floor(Date.now() / 1000) + 43200, sub: sub })));
  var signingInput = header + '.' + payload;
  var sig = ES256.signRaw(strBytes(signingInput), b64urlDecodeBytes(privB64), gsSha256, gsHmac);
  return signingInput + '.' + b64urlEncodeBytes(sig);
}

/** Payload-less Web Push to one endpoint; returns the HTTP status code. */
function webPushSend(endpoint) {
  var aud = (String(endpoint).match(/^(https?:\/\/[^\/]+)/) || [])[1];
  if (!aud) return 0;
  var jwt = vapidJwt(aud);
  var pub = PropertiesService.getScriptProperties().getProperty('VAPID_PUBLIC');
  var res = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    headers: { 'TTL': '86400', 'Authorization': 'vapid t=' + jwt + ', k=' + pub },
    muteHttpExceptions: true
  });
  return res.getResponseCode();
}

// ---- subscription storage ----
function ensurePushSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PUSH_SHEET);
  if (!sheet) { sheet = ss.insertSheet(PUSH_SHEET); sheet.getRange(1, 1, 1, PUSH_HEADERS.length).setValues([PUSH_HEADERS]); }
  return sheet;
}
function readSubs() {
  var sheet = ensurePushSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var v = sheet.getRange(2, 1, lastRow - 1, PUSH_HEADERS.length).getValues();
  var out = [];
  for (var r = 0; r < v.length; r++) { if (v[r][0]) out.push({ endpoint: String(v[r][0]), p256dh: String(v[r][1]), auth: String(v[r][2]), row: r + 2 }); }
  return out;
}
function subscribePush(sub, device) {
  if (!sub || !sub.endpoint) throw new Error('missing push subscription');
  var keys = sub.keys || {};
  var sheet = ensurePushSheet();
  var subs = readSubs();
  for (var i = 0; i < subs.length; i++) {
    if (subs[i].endpoint === sub.endpoint) {
      sheet.getRange(subs[i].row, 1, 1, PUSH_HEADERS.length)
        .setValues([[sub.endpoint, keys.p256dh || '', keys.auth || '', device || '', new Date()]]);
      return;
    }
  }
  sheet.appendRow([sub.endpoint, keys.p256dh || '', keys.auth || '', device || '', new Date()]);
}
function unsubscribePush(endpoint) {
  if (!endpoint) return;
  var sheet = ensurePushSheet();
  var subs = readSubs();
  for (var i = subs.length - 1; i >= 0; i--) { if (subs[i].endpoint === endpoint) sheet.deleteRow(subs[i].row); }
}

// ---- the daily reminder job ----
function istDateString(offsetDays) {
  var d = new Date(Date.now() + offsetDays * 86400000);
  return Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd');
}
/** Trigger target: push if any Active FD matures in {0,1,2} days IST. */
function sendMaturityPush() {
  if (!PropertiesService.getScriptProperties().getProperty('VAPID_PUBLIC')) return;
  var targets = {};
  targets[istDateString(0)] = 1; targets[istDateString(1)] = 1; targets[istDateString(2)] = 1;
  var fds = readFds();
  var due = false;
  for (var i = 0; i < fds.length; i++) {
    if (fds[i].status === 'Active' && targets[fds[i].endDate]) { due = true; break; }
  }
  if (!due) return;
  var subs = readSubs();
  for (var s = 0; s < subs.length; s++) {
    try {
      var code = webPushSend(subs[s].endpoint);
      if (code === 404 || code === 410) unsubscribePush(subs[s].endpoint); // subscription expired
    } catch (e) { /* one bad endpoint shouldn't stop the rest */ }
  }
}
function installMaturityReminders() {
  removeMaturityReminders();
  ScriptApp.newTrigger('sendMaturityPush').timeBased().atHour(10).everyDays(1).create();
}
function removeMaturityReminders() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendMaturityPush') ScriptApp.deleteTrigger(t);
  });
}
