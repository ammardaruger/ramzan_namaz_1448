/**
 * Ramzan 1448 Azan & Takbira — live backend
 *
 * This is the actual data store for the dashboard. Every booking, release,
 * and availability check goes through this script and lands directly as a
 * row in the "Bookings" sheet — there is no separate export step anymore.
 *
 * SETUP (one time)
 *  1. Create a new Google Sheet (or open the one you want to use).
 *  2. Extensions > Apps Script.
 *  3. Delete the starter code, paste this whole file in.
 *  4. Save the project (any name is fine).
 *  5. Click Deploy > New deployment.
 *     - Select type: Web app.
 *     - Description: anything.
 *     - Execute as: Me.
 *     - Who has access: Anyone.
 *  6. Click Deploy, then authorize the permissions it asks for
 *     (it's your own script, acting on your own Sheet).
 *  7. Copy the "Web app URL" it gives you — it ends in /exec.
 *  8. Paste that URL into WEB_APP_URL near the top of the dashboard's
 *     <script> section.
 *
 * If you ever change the deployment (not just edit the code), you'll get a
 * new URL and need to update the dashboard again. Editing the code and
 * clicking Deploy > Manage deployments > Edit > Deploy (same deployment)
 * keeps the same URL.
 */

const SHEET_NAME = "Bookings";
const RAMZAN_START = new Date(2027, 1, 6); // 6 Feb 2027
const DOW = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const HEADERS = ["Date", "Option", "ITS Number", "Email", "Reserved At", "Ramzan Day", "Weekday"];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function formatDateStr_(v) {
  if (v instanceof Date) {
    return v.getFullYear() + "-" + String(v.getMonth() + 1).padStart(2, "0") + "-" + String(v.getDate()).padStart(2, "0");
  }
  return String(v).split("T")[0];
}

function readAllRows_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const out = [];
  values.forEach(function (row, idx) {
    if (row[0] === "" && row[1] === "" && row[2] === "") return;
    out.push({
      rowIndex: idx + 2,
      date: formatDateStr_(row[0]),
      option: row[1],
      its: String(row[2]),
      email: String(row[3]),
      reservedAt: row[4] instanceof Date ? row[4].toISOString() : String(row[4])
    });
  });
  return out;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const rows = readAllRows_();
    return jsonOut_({ ok: true, bookings: rows });
  } catch (err) {
    return jsonOut_({ ok: false, reason: "error", detail: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return jsonOut_({ ok: false, reason: "error", detail: "Server busy, try again." });
  }

  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === "claim") {
      return jsonOut_(claim_(body.date, body.option, body.its, body.email));
    } else if (action === "release") {
      return jsonOut_(release_(body.date, body.option, body.its));
    } else if (action === "list") {
      return jsonOut_({ ok: true, bookings: readAllRows_() });
    }
    return jsonOut_({ ok: false, reason: "error", detail: "Unknown action." });
  } catch (err) {
    return jsonOut_({ ok: false, reason: "error", detail: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function claim_(dateIso, option, its, email) {
  if (!/^\d{8}$/.test(String(its))) {
    return { ok: false, reason: "error", detail: "ITS number must be exactly 8 digits." };
  }
  const emailStr = String(email || "").trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailStr)) {
    return { ok: false, reason: "error", detail: "A valid email address is required." };
  }
  const rows = readAllRows_();

  for (let i = 0; i < rows.length; i++) {
    if (rows[i].date === dateIso && rows[i].option === option) {
      return { ok: false, reason: "taken" };
    }
  }
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].its === String(its)) {
      return { ok: false, reason: "limit", existingDate: rows[i].date, existingOption: rows[i].option };
    }
  }

  const sheet = getSheet_();
  const dObj = new Date(dateIso + "T00:00:00");
  const dayNum = Math.round((dObj - RAMZAN_START) / 86400000) + 1;
  const dow = DOW[dObj.getDay()];
  const now = new Date();

  sheet.appendRow([dateIso, option, String(its), emailStr, now.toISOString(), dayNum, dow]);
  return { ok: true, bookings: readAllRows_() };
}

function release_(dateIso, option, its) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: false, reason: "not_found" };

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    const rDate = formatDateStr_(values[i][0]);
    const rOption = values[i][1];
    const rIts = String(values[i][2]);
    if (rDate === dateIso && rOption === option && rIts === String(its)) {
      sheet.deleteRow(i + 2);
      return { ok: true, bookings: readAllRows_() };
    }
  }
  return { ok: false, reason: "not_found" };
}
