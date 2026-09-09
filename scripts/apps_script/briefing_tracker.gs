/*
 * briefing_tracker.gs -- Backend for store-briefing.html's view/print tracking.
 *
 * This is NOT run by any build step -- it's the reference copy of the code that must be pasted
 * into a Google Apps Script project bound to a Google Sheet, then deployed as a Web App, since
 * GitHub Pages (where the dashboard is hosted) has no server of its own.
 *
 * Setup (one-time, done from a Google account with edit access):
 *   1. Create a new Google Sheet (any name, e.g. "BSC Briefing Tracker").
 *   2. Extensions > Apps Script. Delete the default Code.gs contents and paste this whole file in.
 *   3. Deploy > New deployment > type "Web app". Execute as: Me. Who has access: Anyone.
 *   4. Authorize when prompted. Copy the deployed /exec URL.
 *   5. Paste that URL into APPS_SCRIPT_URL near the top of store-briefing.html.
 *
 * Data model: every view/print is one appended row in a "Log" sheet (created automatically on
 * first write): Timestamp | Store | Event ("view" or "print"). store-briefing.html fetches the
 * full log on load (GET) and derives "last viewed"/"last printed" per store client-side --
 * keeping this script a thin, dumb logger rather than something that needs updating whenever the
 * store list changes.
 */

const SHEET_NAME = 'Log';

function doGet(e) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).map(function(r) {
    return {
      timestamp: (r[0] instanceof Date) ? r[0].toISOString() : String(r[0]),
      store: r[1],
      event: r[2],
    };
  });
  return jsonOutput({ok: true, rows: rows});
}

function doPost(e) {
  try {
    // Sent as Content-Type: text/plain from the browser to avoid a CORS preflight that Apps
    // Script Web Apps can't answer -- parse the JSON body manually.
    const body = JSON.parse(e.postData.contents);
    const store = body.store;
    const event = body.event; // 'view' or 'print'
    if (!store || !event) throw new Error('missing store or event');
    getSheet().appendRow([new Date(), store, event]);
    return jsonOutput({ok: true});
  } catch (err) {
    return jsonOutput({ok: false, error: String(err)});
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Store', 'Event']);
  }
  return sheet;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
