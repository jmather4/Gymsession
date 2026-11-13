// Weights Log -> Google Sheet.
// Paste into the sheet's own script editor (Extensions > Apps Script), set KEY, then
// Deploy > New deployment > Web app, "Execute as: Me", "Who has access: Anyone".
const KEY = 'CHANGE-ME';   // any phrase you like; type the same phrase into the app's Setup tab
const TAB = 'Sheet1';      // the tab that receives the rows

function doPost(e) {
  let out;
  try {
    const body = JSON.parse(e.postData.contents);
    if (KEY === 'CHANGE-ME' || body.key !== KEY) throw new Error('wrong key');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB);
    if (!sheet) throw new Error('no tab called ' + TAB);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const last = sheet.getLastRow();
      // Column I holds the session id, so a session that is sent twice is only stored once.
      const seen = new Set(last > 1 ? sheet.getRange(2, 9, last - 1, 1).getValues().map(r => String(r[0])) : []);
      const rows = (body.rows || []).filter(r => Array.isArray(r) && r.length === 9 && !seen.has(String(r[8])));
      if (rows.length) sheet.getRange(last + 1, 1, rows.length, 9).setValues(rows);
      out = { ok: true, added: rows.length };
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
