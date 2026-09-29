/**
 * Branding Trip Live Status sync bridge.
 * Deploy: Deploy → New deployment → Web app
 * Execute as: Me
 * Who has access: Anyone
 * Then paste the /exec URL into the HTML SHEET.syncUrl (or localStorage wb-sync-url).
 */
var SHEET_ID = '1sM0GIPORgPaD2YQXQKGKeXGow6EfnHvQMVZY_2TmXyM';
var SHEET_NAME = ''; // first sheet if blank

function _sheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
}

function _headers(sh) {
  var row = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < row.length; i++) map[String(row[i]).toLowerCase().trim()] = i;
  return map;
}

function doGet(e) {
  e = e || { parameter: {} };
  var p = e.parameter || {};
  var action = p.action || 'get';
  var out;
  try {
    if (action === 'get') out = { ok: true, rows: readRows_() };
    else if (action === 'set') {
      var rows = JSON.parse(p.rows || '[]');
      out = { ok: true, updated: writeRows_(rows) };
    } else out = { ok: false, error: 'unknown action' };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  var cb = p.callback;
  var body = JSON.stringify(out);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

function readRows_() {
  var sh = _sheet();
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var h = _headers(sh);
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var code = h.code != null ? row[h.code] : '';
    var date = h.date != null ? row[h.date] : '';
    if (!code && !date) continue;
    // Dates may be Date objects
    if (date instanceof Date) {
      date = Utilities.formatDate(date, Session.getScriptTimeZone() || 'Europe/London', 'yyyy-MM-dd');
    } else {
      date = String(date).slice(0, 10);
    }
    rows.push({
      date: date,
      code: String(code).trim(),
      name: h.name != null ? String(row[h.name] || '') : '',
      status: h.status != null ? String(row[h.status] || 'pending').toLowerCase() : 'pending',
      notes: h.notes != null ? String(row[h.notes] || '') : '',
      updated_at: h.updated_at != null ? String(row[h.updated_at] || '') : ''
    });
  }
  return rows;
}

function writeRows_(rows) {
  var sh = _sheet();
  var h = _headers(sh);
  if (h.date == null || h.code == null || h.status == null) {
    throw new Error('Sheet needs date, code, status columns');
  }
  var values = sh.getDataRange().getValues();
  var index = {}; // date|code -> row number (1-based)
  for (var r = 1; r < values.length; r++) {
    var d = values[r][h.date];
    if (d instanceof Date) d = Utilities.formatDate(d, Session.getScriptTimeZone() || 'Europe/London', 'yyyy-MM-dd');
    else d = String(d).slice(0, 10);
    var c = String(values[r][h.code] || '').trim();
    index[d + '|' + c] = r + 1;
  }
  var updated = 0;
  var now = new Date().toISOString();
  rows.forEach(function(item) {
    var key = String(item.date).slice(0,10) + '|' + String(item.code).trim();
    var rowNum = index[key];
    if (!rowNum) return; // only update existing visit rows
    sh.getRange(rowNum, h.status + 1).setValue(item.status || 'pending');
    if (h.notes != null) sh.getRange(rowNum, h.notes + 1).setValue(item.notes || '');
    if (h.updated_at != null) sh.getRange(rowNum, h.updated_at + 1).setValue(item.updated_at || now);
    updated++;
  });
  return updated;
}
