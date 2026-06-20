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
//  Asia/Kolkata — see SETUP.md). VAPID JWTs are signed with the pure-BigInt
//  ES256 implementation below (Apps Script has no native ECDSA); the same
//  algorithm is validated against Node's crypto in the repo's tests.
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

/* ---- pure-BigInt P-256 ECDSA (ES256), no native crypto needed ---- */
var ES256 = (function () {
  var p = 0xffffffff00000001000000000000000000000000ffffffffffffffffffffffffn;
  var a = p - 3n;
  var n = 0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551n;
  var Gx = 0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296n;
  var Gy = 0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5n;
  function mod(x, m) { var r = x % m; return r < 0n ? r + m : r; }
  function modInv(x, m) {
    var lo = mod(x, m), hi = m, a0 = 1n, a1 = 0n;
    while (lo > 0n) { var q = hi / lo; var t = hi - q * lo; hi = lo; lo = t; t = a1 - q * a0; a1 = a0; a0 = t; }
    return mod(a1, m);
  }
  function add(P, Q) {
    if (!P) return Q; if (!Q) return P;
    var x1 = P[0], y1 = P[1], x2 = Q[0], y2 = Q[1];
    if (x1 === x2 && mod(y1 + y2, p) === 0n) return null;
    var m;
    if (x1 === x2 && y1 === y2) m = mod((3n * x1 * x1 + a) * modInv(2n * y1, p), p);
    else m = mod((y2 - y1) * modInv(mod(x2 - x1, p), p), p);
    var x3 = mod(m * m - x1 - x2, p);
    return [x3, mod(m * (x1 - x3) - y1, p)];
  }
  function mul(k, P) { var R = null, A = P; while (k > 0n) { if (k & 1n) R = add(R, A); A = add(A, A); k >>= 1n; } return R; }
  function bytesToBig(b) { var x = 0n; for (var i = 0; i < b.length; i++) x = (x << 8n) | BigInt(b[i]); return x; }
  function bigTo32(x) { var o = new Array(32); for (var i = 31; i >= 0; i--) { o[i] = Number(x & 0xffn); x >>= 8n; } return o; }
  function rfc6979k(h, x, hmac) {
    var z2 = bigTo32(mod(bytesToBig(h), n));
    var V = []; for (var i = 0; i < 32; i++) V.push(0x01);
    var K = []; for (var j = 0; j < 32; j++) K.push(0x00);
    K = hmac(K, V.concat([0x00], x, z2)); V = hmac(K, V);
    K = hmac(K, V.concat([0x01], x, z2)); V = hmac(K, V);
    for (;;) { V = hmac(K, V); var k = mod(bytesToBig(V), n); if (k >= 1n && k < n) return k; K = hmac(K, V.concat([0x00])); V = hmac(K, V); }
  }
  function signRaw(msgBytes, dBytes, sha256, hmac) {
    var h = sha256(msgBytes), z = mod(bytesToBig(h), n), d = bytesToBig(dBytes), r, s;
    for (;;) {
      var k = rfc6979k(h, dBytes, hmac);
      var Rp = mul(k, [Gx, Gy]); r = mod(Rp[0], n); if (r === 0n) continue;
      s = mod(modInv(k, n) * mod(z + r * d, n), n); if (s === 0n) continue;
      if (s > n / 2n) s = n - s; break;
    }
    return bigTo32(r).concat(bigTo32(s));
  }
  function publicKey(dBytes) { var Q = mul(bytesToBig(dBytes), [Gx, Gy]); return [0x04].concat(bigTo32(Q[0]), bigTo32(Q[1])); }
  return { signRaw: signRaw, publicKey: publicKey };
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
