/**
 * Salt & Shine Co. Open House Lead + Property Capture → Google Sheets
 *
 * This version supports SHORT QR codes:
 * - Leads save to the "Open House Leads" tab.
 * - Property setup saves to the "Open House Properties" tab.
 * - QR codes only carry a short property ID, then the live site loads house details from this Sheet.
 */

const LEADS_SHEET_NAME = 'Open House Leads';
const PROPERTIES_SHEET_NAME = 'Open House Properties';
const REQUIRED_SECRET = ''; // Optional. Example: 'saltshine2026'

const LEAD_HEADERS = [
  'Received At', 'Property ID', 'Name', 'Phone', 'Email', 'Timeline', 'Financing',
  'Price Range', 'Working With Agent', 'Notes', 'Property', 'Lead Source',
  'Submitted At', 'Page URL', 'User Agent'
];

const PROPERTY_HEADERS = [
  'Updated At', 'Property ID', 'Address', 'Price', 'Open House Date', 'Beds', 'Baths',
  'Sq Ft', 'Year Built', 'Highlights', 'Agent', 'Agent Phone', 'Booking Link',
  'Showing Text Message', 'Listing URL', 'Property Message', 'Lead Secret', 'Main Image URL'
];

function setupSheet() {
  const leads = getOrCreateSheet_(LEADS_SHEET_NAME);
  ensureHeaders_(leads, LEAD_HEADERS);
  leads.setFrozenRows(1);
  leads.autoResizeColumns(1, LEAD_HEADERS.length);

  const props = getOrCreateSheet_(PROPERTIES_SHEET_NAME);
  ensureHeaders_(props, PROPERTY_HEADERS);
  props.setFrozenRows(1);
  props.autoResizeColumns(1, PROPERTY_HEADERS.length);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = parseRequest_(e);
    if (data.action === 'saveProperty') return json_(saveProperty_(data));

    if (REQUIRED_SECRET && data.secret !== REQUIRED_SECRET) {
      return json_({ ok: false, error: 'Unauthorized' });
    }

    const sheet = getOrCreateSheet_(LEADS_SHEET_NAME);
    ensureHeaders_(sheet, LEAD_HEADERS);
    sheet.appendRow([
      new Date(), clean_(data.propertyId), clean_(data.name), clean_(data.phone), clean_(data.email),
      clean_(data.timeline), clean_(data.financing), clean_(data.priceRange), clean_(data.hasAgent),
      clean_(data.notes), clean_(data.property), clean_(data.leadSource), clean_(data.submittedAt),
      clean_(data.pageUrl), clean_(data.userAgent)
    ]);
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const action = params.action || '';
    const callback = params.callback || '';

    if (action === 'saveProperty') {
      const data = decodePayload_(params.payload || '');
      return json_(saveProperty_(data), callback);
    }

    if (action === 'getProperty') {
      return json_(getProperty_(params.propertyId || ''), callback);
    }

    if (action === 'resolveListingImage') {
      return json_({
        ok: true,
        imageUrl: resolveListingImage_(params.listingUrl || '')
      }, callback);
    }

    if (action === 'listLeads') {
      return json_(listLeads_(params.propertyId || ''), callback);
    }

    return json_({ ok: true, message: 'Salt & Shine endpoint is live.' }, callback);
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) }, e && e.parameter && e.parameter.callback);
  }
}

function saveProperty_(data) {
  const propertyId = clean_(data.propertyId || data.id);
  if (!propertyId) return { ok: false, error: 'Missing property ID' };
  const imageUrl = clean_(data.propertyImageUrl || data.imageUrl) || resolveListingImage_(data.listingUrl);

  const sheet = getOrCreateSheet_(PROPERTIES_SHEET_NAME);
  ensureHeaders_(sheet, PROPERTY_HEADERS);

  const row = [
    new Date(), propertyId, clean_(data.address), clean_(data.price), clean_(data.date),
    clean_(data.beds), clean_(data.baths), clean_(data.sqft), clean_(data.year),
    clean_(data.highlights), clean_(data.agent), clean_(data.agentPhone), clean_(data.bookingLink),
    clean_(data.showingText), clean_(data.listingUrl), clean_(data.propertyMessage), clean_(data.leadSecret),
    imageUrl
  ];

  const existingRow = findPropertyRow_(sheet, propertyId);
  if (existingRow) {
    sheet.getRange(existingRow, 1, 1, PROPERTY_HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { ok: true, propertyId, imageUrl };
}

function getProperty_(propertyId) {
  propertyId = clean_(propertyId);
  if (!propertyId) return { ok: false, error: 'Missing property ID' };
  const sheet = getOrCreateSheet_(PROPERTIES_SHEET_NAME);
  ensureHeaders_(sheet, PROPERTY_HEADERS);
  const rowNum = findPropertyRow_(sheet, propertyId);
  if (!rowNum) return { ok: false, error: 'Property not found' };
  const row = sheet.getRange(rowNum, 1, 1, PROPERTY_HEADERS.length).getValues()[0];
  return {
    ok: true,
    property: {
      propertyId: row[1], address: row[2], price: row[3], date: row[4], beds: row[5], baths: row[6],
      sqft: row[7], year: row[8], highlights: row[9], agent: row[10], agentPhone: row[11],
      bookingLink: row[12], showingText: row[13], listingUrl: row[14], propertyMessage: row[15],
      leadSecret: row[16], propertyImageUrl: row[17]
    }
  };
}

function resolveListingImage_(listingUrl) {
  listingUrl = clean_(listingUrl);
  if (!listingUrl) return '';

  const urlMatch = listingUrl.match(/^https?:\/\/([^/:?#]+)/i);
  if (!urlMatch) return '';
  const host = urlMatch[1].toLowerCase();
  if (!host || host === 'localhost' || /^127\./.test(host) || /^10\./.test(host) ||
      /^192\.168\./.test(host) || /^172\.(1[6-9]|2\d|3[01])\./.test(host)) return '';

  try {
    const response = UrlFetchApp.fetch(listingUrl, {
      followRedirects: true,
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; SaltShineOpenHouse/1.0)',
        'Accept': 'text/html,application/xhtml+xml'
      }
    });
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 400) return '';
    const html = response.getContentText();
    const patterns = [
      /<meta[^>]+property=["']og:image(?::secure_url)?["'][^>]+content=["']([^"']+)["']/i,
      /<meta[^>]+content=["']([^"']+)["'][^>]+property=["']og:image(?::secure_url)?["']/i,
      /<meta[^>]+name=["']twitter:image(?::src)?["'][^>]+content=["']([^"']+)["']/i,
      /<meta[^>]+content=["']([^"']+)["'][^>]+name=["']twitter:image(?::src)?["']/i,
      /<link[^>]+rel=["']image_src["'][^>]+href=["']([^"']+)["']/i
    ];
    for (let i = 0; i < patterns.length; i++) {
      const match = html.match(patterns[i]);
      if (match && match[1]) return decodeHtmlEntities_(match[1]);
    }
  } catch (err) {
    console.warn('Could not resolve listing image: ' + err);
  }
  return '';
}

function decodeHtmlEntities_(value) {
  return clean_(value)
    .replace(/&amp;/g, '&')
    .replace(/&#x2F;/gi, '/')
    .replace(/&#47;/g, '/')
    .replace(/&quot;/g, '"');
}

function listLeads_(propertyId) {
  propertyId = clean_(propertyId);
  const sheet = getOrCreateSheet_(LEADS_SHEET_NAME);
  ensureHeaders_(sheet, LEAD_HEADERS);
  const last = sheet.getLastRow();
  if (last < 2) return { ok: true, leads: [] };

  const rows = sheet.getRange(2, 1, last - 1, LEAD_HEADERS.length).getValues();
  const leads = rows.map(row => ({
    receivedAt: serializeDate_(row[0]),
    propertyId: clean_(row[1]),
    name: clean_(row[2]),
    phone: clean_(row[3]),
    email: clean_(row[4]),
    timeline: clean_(row[5]),
    financing: clean_(row[6]),
    priceRange: clean_(row[7]),
    hasAgent: clean_(row[8]),
    notes: clean_(row[9]),
    property: clean_(row[10]),
    leadSource: clean_(row[11]),
    submittedAt: serializeDate_(row[12]),
    pageUrl: clean_(row[13]),
    userAgent: clean_(row[14]),
    date: row[12] ? formatDate_(row[12]) : formatDate_(row[0])
  })).filter(lead => !propertyId || lead.propertyId === propertyId);

  leads.reverse();
  return { ok: true, leads };
}

function serializeDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return value.toISOString();
  return clean_(value);
}

function formatDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'M/d/yyyy');
  }
  return clean_(value);
}

function findPropertyRow_(sheet, propertyId) {
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const ids = sheet.getRange(2, 2, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (clean_(ids[i][0]) === propertyId) return i + 2;
  }
  return null;
}

function parseRequest_(e) {
  if (!e) return {};
  if (e.postData && e.postData.contents) {
    const body = e.postData.contents;
    try { return JSON.parse(body); } catch (err) {}
  }
  return e.parameter || {};
}

function decodePayload_(payload) {
  if (!payload) return {};
  try {
    const text = Utilities.newBlob(Utilities.base64Decode(payload)).getDataAsString();
    return JSON.parse(text);
  } catch (err) {
    const text2 = Utilities.newBlob(Utilities.base64DecodeWebSafe(payload)).getDataAsString();
    return JSON.parse(text2);
  }
}

function getOrCreateSheet_(name) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

function ensureHeaders_(sheet, headers) {
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = firstRow.some(value => String(value).trim() !== '');
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return;
  }
  const updated = headers.map((header, index) => firstRow[index] || header);
  sheet.getRange(1, 1, 1, headers.length).setValues([updated]);
}

function clean_(value) {
  if (value === undefined || value === null) return '';
  return String(value).trim();
}

function json_(payload, callback) {
  const text = JSON.stringify(payload);
  if (callback) {
    return ContentService
      .createTextOutput(String(callback).replace(/[^a-zA-Z0-9_.$]/g, '') + '(' + text + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}
