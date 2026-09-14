/**
 * OLEA Move In — Sheet API
 * ---------------------------------------------------------------
 * Paste this into Extensions ▸ Apps Script in your spreadsheet,
 * then Deploy ▸ New deployment ▸ Web app
 *   Execute as:      Me
 *   Who has access:  Anyone
 * Copy the /exec URL into the app.
 *
 * The script writes into the existing table (header row is found by
 * searching for "Room / Space"), keeps your totals row intact, and
 * never overwrites a cell that contains a formula.
 */

// Leave this empty and the app opens with no passcode — that's the
// default. If you ever want to lock it down, put any string here,
// redeploy a new version, and the app will start asking for it.
const API_TOKEN = '';

const SHEET_NAME = 'Move-In Purchases';
const HEADER_ANCHOR = 'Room / Space';
const DESC_HEADER = 'Item Description';
const PRICE_HEADER = 'Estimated Price';
const STATUS_HEADER = 'Order Status';
const PRIORITY_HEADER = 'Priority';

/* ============================ router ============================ */

function doGet(e) { return route(e); }
function doPost(e) { return route(e); }

function route(e) {
  var out;
  try {
    var p = {};
    if (e && e.parameter) { for (var k in e.parameter) p[k] = e.parameter[k]; }
    if (e && e.postData && e.postData.contents) {
      var body = JSON.parse(e.postData.contents);
      for (var b in body) p[b] = body[b];
    }

    if (API_TOKEN && String(p.token || '') !== API_TOKEN) {
      return ContentService
        .createTextOutput(JSON.stringify({
          ok: false, needAuth: true,
          error: p.token ? 'That passcode doesn\u2019t match.' : 'This plan is passcode protected.'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var action = p.action || 'list';
    var item = p.item;
    if (typeof item === 'string' && item) item = JSON.parse(item);

    if (action === 'list') {
      out = { ok: true, data: readAll() };
    } else {
      var lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        if (action === 'create')      { createRow(item); }
        else if (action === 'update') { updateRow(item); }
        else if (action === 'delete') { deleteRow(item); }
        else { throw new Error('Unknown action: ' + action); }
        refreshTotals();
        out = { ok: true, data: readAll() };
      } finally {
        lock.releaseLock();
      }
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================ layout ============================ */

function sheet_() {
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

/** Locates the header row, the column block, and the data extent. */
function layout_() {
  var sh = sheet_();
  var grid = sh.getDataRange().getValues();
  var hr = -1, c0 = -1;

  for (var r = 0; r < Math.min(grid.length, 40) && hr < 0; r++) {
    for (var c = 0; c < grid[r].length; c++) {
      if (norm_(grid[r][c]) === norm_(HEADER_ANCHOR)) { hr = r; c0 = c; break; }
    }
  }
  if (hr < 0) throw new Error('Could not find a "' + HEADER_ANCHOR + '" header cell in "' + sh.getName() + '".');

  var headers = [];
  for (var hc = c0; hc < grid[hr].length; hc++) {
    var h = String(grid[hr][hc]).trim();
    if (!h) break;
    headers.push(h);
  }

  var rows = [], blanks = 0, last = hr, totalRow = -1;
  for (var d = hr + 1; d < grid.length; d++) {
    var room = String(grid[d][c0] || '').trim();
    var desc = String(grid[d][c0 + 1] || '').trim();
    if (/^total\b/i.test(room)) { totalRow = d; break; }
    if (!room && !desc) { if (++blanks >= 5) break; continue; }
    blanks = 0;
    rows.push(d);
    last = d;
  }

  return {
    sheet: sh, grid: grid, headers: headers,
    headerRow: hr + 1, col0: c0 + 1,
    dataRows: rows.map(function (i) { return i + 1; }),
    lastRow: last + 1,
    totalRow: totalRow < 0 ? -1 : totalRow + 1
  };
}

function norm_(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().toLowerCase(); }

function colOf_(L, header) {
  for (var i = 0; i < L.headers.length; i++) {
    if (norm_(L.headers[i]) === norm_(header)) return L.col0 + i;
  }
  return -1;
}

/* ============================= read ============================= */

function readAll() {
  var L = layout_();
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var n = L.headers.length;
  var items = [];

  for (var i = 0; i < L.dataRows.length; i++) {
    var row = L.dataRows[i];
    var vals = L.sheet.getRange(row, L.col0, 1, n).getValues()[0];
    var cells = {};
    for (var j = 0; j < n; j++) cells[L.headers[j]] = cellOut_(vals[j], tz);
    items.push({ row: row, cells: cells });
  }

  return {
    sheetName: L.sheet.getName(),
    headers: L.headers,
    items: items,
    closingDate: findClosingDate_(L, tz),
    updated: Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss")
  };
}

function cellOut_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  if (v === null || v === undefined) return '';
  return v;
}

/** Reads "Target Closing: Oct 6, 2026" from the dashboard block. */
function findClosingDate_(L, tz) {
  for (var r = 0; r < L.headerRow - 1; r++) {
    var line = L.grid[r] || [];
    for (var c = 0; c < line.length; c++) {
      var raw = line[c];
      if (raw instanceof Date) {
        var near = String(line[c - 1] || '') + String(line[c - 2] || '');
        if (/closing/i.test(near)) return Utilities.formatDate(raw, tz, 'yyyy-MM-dd');
      }
      var s = String(raw || '');
      if (/closing/i.test(s)) {
        var m = s.match(/([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})|(\d{4}-\d{2}-\d{2})/);
        if (m) {
          var d = new Date(m[0]);
          if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
        }
      }
    }
  }
  return '';
}

/* ============================ writes ============================ */

function createRow(item) {
  if (!item) throw new Error('No item supplied.');
  var L = layout_();
  var at = L.lastRow;                      // last populated data row
  L.sheet.insertRowsAfter(at, 1);          // pushes the totals row down
  var target = at + 1;
  L.sheet.getRange(at, L.col0, 1, L.headers.length)
    .copyTo(L.sheet.getRange(target, L.col0, 1, L.headers.length), { formatOnly: true });
  writeCells_(L, target, item.cells);
}

function updateRow(item) {
  if (!item) throw new Error('No item supplied.');
  var L = layout_();
  var row = resolveRow_(L, item);
  writeCells_(L, row, item.cells);
}

function deleteRow(item) {
  var L = layout_();
  var row = resolveRow_(L, item);
  L.sheet.deleteRow(row);
}

/**
 * Trusts the row number only if the description still matches.
 * Otherwise re-finds the item — so a stale tab can't clobber the wrong line.
 */
function resolveRow_(L, item) {
  var descCol = colOf_(L, DESC_HEADER);
  var want = norm_(item.match || (item.cells ? item.cells[DESC_HEADER] : ''));
  var row = Number(item.row || 0);

  if (row && L.dataRows.indexOf(row) > -1) {
    if (!want) return row;
    if (descCol > 0 && norm_(L.sheet.getRange(row, descCol).getValue()) === want) return row;
  }
  if (want && descCol > 0) {
    for (var i = 0; i < L.dataRows.length; i++) {
      if (norm_(L.sheet.getRange(L.dataRows[i], descCol).getValue()) === want) return L.dataRows[i];
    }
  }
  throw new Error('That row is no longer in the sheet. Refresh and try again.');
}

function writeCells_(L, row, cells) {
  if (!cells) return;
  for (var header in cells) {
    var col = colOf_(L, header);
    if (col < 0) continue;
    var cell = L.sheet.getRange(row, col);
    if (cell.getFormula()) continue;                 // never stomp a formula
    cell.setValue(coerce_(header, cells[header]));
  }
}

function coerce_(header, value) {
  var s = String(value == null ? '' : value).trim();
  if (s === '') return '';

  if (norm_(header).indexOf('date') > -1) {
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return s;
  }
  if (norm_(header) === norm_(PRICE_HEADER)) {
    var num = Number(s.replace(/[^0-9.\-]/g, ''));
    return isNaN(num) ? s : num;
  }
  return value;
}

/* =========================== totals =========================== */

/** Refreshes the dashboard strip and totals row — formulas are left alone. */
function refreshTotals() {
  var L = layout_();
  var priceCol = colOf_(L, PRICE_HEADER);
  var statusCol = colOf_(L, STATUS_HEADER);
  var prioCol = colOf_(L, PRIORITY_HEADER);

  var total = 0, delivered = 0, pending = 0, essential = 0, count = L.dataRows.length;

  for (var i = 0; i < L.dataRows.length; i++) {
    var r = L.dataRows[i];
    if (priceCol > 0) {
      var p = Number(String(L.sheet.getRange(r, priceCol).getValue()).replace(/[^0-9.\-]/g, ''));
      if (!isNaN(p)) total += p;
    }
    if (statusCol > 0) {
      var st = norm_(L.sheet.getRange(r, statusCol).getValue());
      if (st === 'delivered') delivered++;
      else if (st === 'ordered' || st === 'in transit') pending++;
    }
    if (prioCol > 0 && /essential/i.test(String(L.sheet.getRange(r, prioCol).getValue()))) essential++;
  }

  var days = '';
  var closing = findClosingDate_(L, SpreadsheetApp.getActive().getSpreadsheetTimeZone());
  if (closing) {
    var cm = closing.split('-');
    var cd = new Date(Number(cm[0]), Number(cm[1]) - 1, Number(cm[2]));
    var today = new Date(); today.setHours(0, 0, 0, 0);
    days = Math.max(0, Math.round((cd - today) / 86400000)) + ' Days';
  }

  setUnderLabel_(L, 'days until closing', days);
  setUnderLabel_(L, 'total estimated spend', total);
  setUnderLabel_(L, 'essential items', essential + ' Items');
  setUnderLabel_(L, 'delivered', delivered + ' of ' + count);
  setUnderLabel_(L, 'orders pending', pending + ' Pending');

  if (L.totalRow > 0) {
    if (priceCol > 0) setIfNotFormula_(L.sheet.getRange(L.totalRow, priceCol), total);
    if (statusCol > 0) setIfNotFormula_(L.sheet.getRange(L.totalRow, statusCol), delivered + ' of ' + count + ' Delivered');
  }
}

function setUnderLabel_(L, label, value) {
  for (var r = 0; r < L.headerRow - 1; r++) {
    var line = L.grid[r] || [];
    for (var c = 0; c < line.length; c++) {
      if (norm_(line[c]) === label) {
        setIfNotFormula_(L.sheet.getRange(r + 2, c + 1), value);
        return;
      }
    }
  }
}

function setIfNotFormula_(range, value) {
  if (range.getFormula()) return;
  range.setValue(value);
}

/* ============================ helper ============================ */

/** Run once from the editor to confirm the script can see your table. */
function testConnection() {
  var d = readAll();
  Logger.log('Sheet: %s · %s items · closing %s', d.sheetName, d.items.length, d.closingDate);
  Logger.log(JSON.stringify(d.items[0], null, 2));
}
