/**
 * Salt & Shine Co. Open House Lead Capture → Google Sheets
 *
 * How to use:
 * 1) Create a Google Sheet named something like "Salt & Shine Leads".
 * 2) In the Sheet, go to Extensions → Apps Script.
 * 3) Delete any starter code and paste this entire file.
 * 4) Optional: set REQUIRED_SECRET to the same secret you enter in the website setup.
 * 5) Run setupSheet once, approve permissions, then deploy as a Web App.
 */

const SHEET_NAME = 'Open House Leads';
const REQUIRED_SECRET = ''; // Optional. Example: 'saltshine2026'

const HEADERS = [
  'Received At',
  'Name',
  'Phone',
  'Email',
  'Timeline',
  'Financing',
  'Price Range',
  'Working With Agent',
  'Notes',
  'Property',
  'Lead Source',
  'Submitted At',
  'Page URL',
  'User Agent'
];

function setupSheet() {
  const sheet = getOrCreateSheet_();
  ensureHeaders_(sheet);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = parseRequest_(e);

    if (REQUIRED_SECRET && data.secret !== REQUIRED_SECRET) {
      return json_({ ok: false, error: 'Unauthorized' });
    }

    const sheet = getOrCreateSheet_();
    ensureHeaders_(sheet);

    sheet.appendRow([
      new Date(),
      clean_(data.name),
      clean_(data.phone),
      clean_(data.email),
      clean_(data.timeline),
      clean_(data.financing),
      clean_(data.priceRange),
      clean_(data.hasAgent),
      clean_(data.notes),
      clean_(data.property),
      clean_(data.leadSource),
      clean_(data.submittedAt),
      clean_(data.pageUrl),
      clean_(data.userAgent)
    ]);

    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, message: 'Salt & Shine lead endpoint is live.' });
}

function parseRequest_(e) {
  if (!e) return {};

  if (e.postData && e.postData.contents) {
    const body = e.postData.contents;
    try {
      return JSON.parse(body);
    } catch (err) {
      // Fall through to form-style parameters below.
    }
  }

  return e.parameter || {};
}

function getOrCreateSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function ensureHeaders_(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = firstRow.some(value => String(value).trim() !== '');

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    return;
  }

  // If headers exist but are outdated, only fill missing header cells without wiping your sheet.
  const updated = HEADERS.map((header, index) => firstRow[index] || header);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([updated]);
}

function clean_(value) {
  if (value === undefined || value === null) return '';
  return String(value).trim();
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
