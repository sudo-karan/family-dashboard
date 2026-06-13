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
      transferFds(req.from, req.to);
    } else if (action === 'undo') {
      undoLog(req.logId);
    } else {
      return jsonOut({ ok: false, error: 'unknown action: ' + action });
    }

    // Commit buffered sheet writes before reading the state back — without
    // this, the read in the same request can still see pre-change data
    // (classic Apps Script behaviour, most visible after deleteRow).
    SpreadsheetApp.flush();
    return jsonOut({ ok: true, fds: readFds(), accounts: readAccounts(), log: readLog(15) });
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

/** Move every FD from one account to another (reassign the `account` field). */
function transferFds(from, to) {
  from = String(from || '').trim();
  to = String(to || '').trim();
  if (!from || !to) throw new Error('choose both accounts');
  if (from.toLowerCase() === to.toLowerCase()) throw new Error('pick a different destination account');
  if (!findAccountByName(to)) throw new Error('account "' + to + '" not found');
  var fds = readFds();
  var moved = [];
  for (var i = 0; i < fds.length; i++) {
    if (String(fds[i].account).trim().toLowerCase() === from.toLowerCase()) {
      rawUpdateFd(applyDefaults(Object.assign({}, fds[i], { account: to })));
      moved.push(fds[i].id);
    }
  }
  if (moved.length === 0) throw new Error('"' + from + '" has no FDs to move');
  var payload = { from: from, to: to, ids: moved };
  appendLog('transfer', 'account:' + from.toLowerCase(),
    'Moved ' + moved.length + ' FD' + (moved.length === 1 ? '' : 's') + ' from "' + from + '" to "' + to + '"',
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
