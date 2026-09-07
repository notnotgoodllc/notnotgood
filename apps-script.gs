/**
 * Google Apps Script backend for the manga survey.
 *
 * It receives one POST per survey step and upserts a single row per respondent
 * (keyed by `id`), so the row fills in as the person progresses — you see partial
 * and abandoned responses, not just completed ones.
 *
 * SETUP
 * 1. Create a Google Sheet.
 * 2. Extensions → Apps Script. Delete the sample code, paste ALL of this in.
 * 3. Deploy → New deployment → type "Web app".
 *      - Execute as:  Me
 *      - Who has access:  Anyone
 *    Click Deploy, authorize when prompted, and copy the Web app URL
 *    (it ends in /exec).
 * 4. Paste that URL into SHEET_ENDPOINT at the top of index.html.
 *
 * To change the survey later you don't need to touch this file — new columns are
 * added automatically from whatever keys the page sends.
 */

const SHEET_NAME = "Responses";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // avoid races when two people submit at once

  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // Ensure the header row covers every key in this payload (add new columns).
    let headers = sheet.getLastColumn()
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      : [];
    let headersChanged = false;
    Object.keys(data).forEach(function (k) {
      if (headers.indexOf(k) === -1) { headers.push(k); headersChanged = true; }
    });
    if (headersChanged) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Find the existing row for this respondent id, if any.
    const idCol = headers.indexOf("id") + 1;
    let rowIndex = -1;
    if (sheet.getLastRow() > 1 && idCol > 0) {
      const ids = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0] === data.id) { rowIndex = i + 2; break; }
      }
    }
    const isNew = rowIndex === -1;
    if (isNew) rowIndex = sheet.getLastRow() + 1;

    // Merge: keep existing cell values for any column not present in this payload.
    const existing = (!isNew)
      ? sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0]
      : [];
    const out = headers.map(function (h, i) {
      if (Object.prototype.hasOwnProperty.call(data, h)) return data[h];
      return existing[i] !== undefined ? existing[i] : "";
    });
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([out]);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// A GET is handy for a quick "is it deployed?" check in the browser.
function doGet() {
  return json({ ok: true, service: "manga-survey" });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
