/**
 * Board Remuneration Benchmarking Survey — intake script.
 *
 * SETUP
 * 1. Open your Google Sheet (built from board_remuneration_google_sheet_template.xlsx).
 * 2. Extensions > Apps Script.
 * 3. Delete any starter code, paste this whole file in, and save.
 * 4. Deploy > New deployment > select type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 5. Click Deploy, approve the permissions Google asks for, and copy the
 *    "Web app" URL it gives you.
 * 6. Paste that URL into SCRIPT_URL near the top of the survey HTML file.
 *
 * The column order below must match the "Responses" sheet header row exactly.
 */

const SHEET_NAME = 'Responses';

const COLUMN_ORDER = [
  'submittedAt',
  'q1',
  'q2',
  'q3',
  'q4',
  'q5',
  'q6',
  'q7',
  'q8',
  'q9',
  'q10',
  'comment'
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    const row = COLUMN_ORDER.map(key => {
      let v = data[key];
      if (Array.isArray(v)) v = v.join('; ');
      if (v === undefined || v === null) v = '';
      return v;
    });

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Lets you sanity-check the deployment by visiting the Web app URL in a browser.
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Survey intake endpoint is live.' }))
    .setMimeType(ContentService.MimeType.JSON);
}
