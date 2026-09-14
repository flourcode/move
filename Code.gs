/**
 * OLEA Move In — Sheet API
 * ---------------------------------------------------------------
 * Paste into Extensions ▸ Apps Script, then Deploy ▸ New deployment
 * ▸ Web app (Execute as: Me · Who has access: Anyone).
 *
 * Reads the sheet in two calls and writes a whole row in one, rather
 * than touching cells one at a time. Formulas are never overwritten.
 */

// Leave empty for no passcode. Put any string here to require one.
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
      return json_({ ok: false, needAuth: true, error: 'Passcode required.' });
    }

    var action = p.action || 'list';
    var item = p.item;
    if (typeof item === 'string' && item) item = JSON.parse(item);

    if (action === 'list') {
      out = { ok: true, data: readAll(layout_()) };
    } else {
      var lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        var L = layout_();
        if (action === 'update') {
          updateRow_(L, item);         // same rows, grid patched in place
        } else if (action === 'create' || action === 'delete') {
          if (action === 'create') createRow_(L, item); else deleteRow_(L, item);
          L = layout_();               // rows shifted, so re-read once
        } else {
          throw new Error('Unknown action: ' + action);
        }
        refreshTotals_(L);
        out = { ok: true, data: readAll(L) };
      } finally {
        lock.releaseLock();
      }
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return json_(out);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================ layout ============================ */

function sheet_() {
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

/**
 * Two service calls total — the value grid and the formula grid.
 * Everything downstream reads from these in memory.
 */
function layout_() {
  var sh = sheet_();
  var rng = sh.getDataRange();
  var grid = rng.getValues();
  var fx = rng.getFormulas();

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
    rows.push(d + 1);
    last = d;
  }

  var idx = {};
  for (var i = 0; i < headers.length; i++) idx[norm_(headers[i])] = c0 + 1 + i;

  return {
    sheet: sh, grid: grid, fx: fx, headers: headers, idx: idx,
    headerRow: hr + 1, col0: c0 + 1, n: headers.length,
    dataRows: rows, lastRow: last + 1,
    totalRow: totalRow < 0 ? -1 : totalRow + 1,
    tz: SpreadsheetApp.getActive().getSpreadsheetTimeZone()
  };
}

function norm_(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().toLowerCase(); }
function colOf_(L, header) { return L.idx[norm_(header)] || -1; }
function cellAt_(L, row, col) { return L.grid[row - 1][col - 1]; }
function fxAt_(L, row, col) {
  var r = L.fx[row - 1];
  return r ? (r[col - 1] || '') : '';
}

/* ============================= read ============================= */

function readAll(L) {
  var items = [];
  for (var i = 0; i < L.dataRows.length; i++) {
    var row = L.dataRows[i];
    var cells = {};
    for (var j = 0; j < L.n; j++) cells[L.headers[j]] = out_(L.grid[row - 1][L.col0 - 1 + j], L.tz);
    items.push({ row: row, cells: cells });
  }
  return {
    sheetName: L.sheet.getName(),
    headers: L.headers,
    items: items,
    closingDate: findClosingDate_(L),
    updated: Utilities.formatDate(new Date(), L.tz, "yyyy-MM-dd'T'HH:mm:ss")
  };
}

function out_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return (v === null || v === undefined) ? '' : v;
}

function findClosingDate_(L) {
  for (var r = 0; r < L.headerRow - 1; r++) {
    var line = L.grid[r] || [];
    for (var c = 0; c < line.length; c++) {
      var raw = line[c];
      if (raw instanceof Date) {
        if (/closing/i.test(String(line[c - 1] || '') + String(line[c - 2] || ''))) {
          return Utilities.formatDate(raw, L.tz, 'yyyy-MM-dd');
        }
      }
      var s = String(raw || '');
      if (/closing/i.test(s)) {
        var m = s.match(/([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})|(\d{4}-\d{2}-\d{2})/);
        if (m) {
          var d = new Date(m[0]);
          if (!isNaN(d.getTime())) return Utilities.formatDate(d, L.tz, 'yyyy-MM-dd');
        }
      }
    }
  }
  return '';
}

/* ============================ writes ============================ */

/** One setValues call for the whole row. Formulas are written back untouched. */
function writeRow_(L, row, cells, fresh) {
  var byHeader = {};
  for (var k in cells) byHeader[norm_(k)] = cells[k];

  var out = [];
  for (var i = 0; i < L.n; i++) {
    var col = L.col0 + i;
    var f = fresh ? '' : fxAt_(L, row, col);
    if (f) { out.push(f); continue; }             // put the formula straight back
    var h = L.headers[i];
    var key = norm_(h);
    if (byHeader.hasOwnProperty(key)) out.push(coerce_(h, byHeader[key]));
    else out.push(fresh ? '' : cellAt_(L, row, col));
  }
  L.sheet.getRange(row, L.col0, 1, L.n).setValues([out]);

  // keep the cached grid honest so totals and the reply reflect this write
  if (L.grid[row - 1]) {
    for (var j = 0; j < L.n; j++) {
      var cj = L.col0 + j;
      if (!fresh && fxAt_(L, row, cj)) continue;   // computed value is unknown
      L.grid[row - 1][cj - 1] = out[j];
    }
  }
}

function createRow_(L, item) {
  if (!item) throw new Error('No item supplied.');
  var at = L.lastRow;
  L.sheet.insertRowsAfter(at, 1);
  L.sheet.getRange(at, L.col0, 1, L.n)
    .copyTo(L.sheet.getRange(at + 1, L.col0, 1, L.n), { formatOnly: true });
  writeRow_(L, at + 1, item.cells, true);
}

function updateRow_(L, item) {
  if (!item) throw new Error('No item supplied.');
  writeRow_(L, resolveRow_(L, item), item.cells, false);
}

function deleteRow_(L, item) {
  L.sheet.deleteRow(resolveRow_(L, item));
}

/** Trusts the row number only while the description still matches. */
function resolveRow_(L, item) {
  var descCol = colOf_(L, DESC_HEADER);
  var want = norm_(item.match || (item.cells ? item.cells[DESC_HEADER] : ''));
  var row = Number(item.row || 0);

  if (row > 0 && L.dataRows.indexOf(row) > -1) {
    if (!want) return row;
    if (descCol > 0 && norm_(cellAt_(L, row, descCol)) === want) return row;
  }
  if (want && descCol > 0) {
    for (var i = 0; i < L.dataRows.length; i++) {
      if (norm_(cellAt_(L, L.dataRows[i], descCol)) === want) return L.dataRows[i];
    }
  }
  throw new Error('That row is no longer in the sheet. Refresh and try again.');
}

function coerce_(header, value) {
  var s = String(value == null ? '' : value).trim();
  if (s === '') return '';
  if (norm_(header).indexOf('date') > -1) {
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : s;
  }
  if (norm_(header) === norm_(PRICE_HEADER)) {
    var n = Number(s.replace(/[^0-9.\-]/g, ''));
    return isNaN(n) ? s : n;
  }
  return value;
}

/* =========================== totals =========================== */

/** Counts from the in-memory grid, then writes only the summary cells. */
function refreshTotals_(L) {
  var priceCol = colOf_(L, PRICE_HEADER);
  var statusCol = colOf_(L, STATUS_HEADER);
  var prioCol = colOf_(L, PRIORITY_HEADER);
  var total = 0, delivered = 0, pending = 0, essential = 0;
  var count = L.dataRows.length;

  for (var i = 0; i < count; i++) {
    var r = L.dataRows[i];
    if (priceCol > 0) {
      var p = Number(String(cellAt_(L, r, priceCol)).replace(/[^0-9.\-]/g, ''));
      if (!isNaN(p)) total += p;
    }
    if (statusCol > 0) {
      var st = norm_(cellAt_(L, r, statusCol));
      if (st === 'delivered') delivered++;
      else if (st === 'ordered' || st === 'in transit') pending++;
    }
    if (prioCol > 0 && /essential/i.test(String(cellAt_(L, r, prioCol)))) essential++;
  }

  var days = '';
  var closing = findClosingDate_(L);
  if (closing) {
    var c = closing.split('-');
    var cd = new Date(Number(c[0]), Number(c[1]) - 1, Number(c[2]));
    var today = new Date(); today.setHours(0, 0, 0, 0);
    days = Math.max(0, Math.round((cd - today) / 86400000)) + ' Days';
  }

  put_(L, 'days until closing', days);
  put_(L, 'total estimated spend', total);
  put_(L, 'essential items', essential + ' Items');
  put_(L, 'delivered', delivered + ' of ' + count);
  put_(L, 'orders pending', pending + ' Pending');

  if (L.totalRow > 0) {
    if (priceCol > 0 && !fxAt_(L, L.totalRow, priceCol)) {
      L.sheet.getRange(L.totalRow, priceCol).setValue(total);
    }
    if (statusCol > 0 && !fxAt_(L, L.totalRow, statusCol)) {
      L.sheet.getRange(L.totalRow, statusCol).setValue(delivered + ' of ' + count + ' Delivered');
    }
  }
}

/** Writes under a dashboard label, skipping any cell holding a formula. */
function put_(L, label, value) {
  for (var r = 0; r < L.headerRow - 1; r++) {
    var line = L.grid[r] || [];
    for (var c = 0; c < line.length; c++) {
      if (norm_(line[c]) === label) {
        if (!fxAt_(L, r + 2, c + 1)) L.sheet.getRange(r + 2, c + 1).setValue(value);
        return;
      }
    }
  }
}

/* ============================ helper ============================ */

/** Run once from the editor to confirm the script can see your table. */
function testConnection() {
  var t = new Date();
  var d = readAll(layout_());
  Logger.log('%s · %s items · closing %s · read in %sms',
    d.sheetName, d.items.length, d.closingDate, new Date() - t);
}
