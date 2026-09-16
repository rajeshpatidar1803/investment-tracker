/**
 * Investment Tracker backend — Google Apps Script Web App.
 *
 * Setup:
 * 1. Open your Google Sheet, then Extensions > Apps Script.
 * 2. Delete the default code, paste this whole file in as Code.gs.
 * 3. Set ADMIN_KEY below to a long random string you choose.
 * 4. In the toolbar, pick "setup" from the function dropdown and click Run.
 *    Authorize when prompted — this creates the "Clients" and "Entries"
 *    tabs with headers and a sample client row.
 * 5. Deploy > New deployment > type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy the deployment URL into js/app.js as WEB_APP_URL.
 */

const CLIENTS_SHEET = 'Clients';
const ENTRIES_SHEET = 'Entries';
const ADMIN_KEY = 'CHANGE_ME_ADMIN_KEY';
const DEFAULT_RATE = 0.13; // 13% p.a., used when a client has no Rate column value

function doGet(e) {
  return handle(e.parameter);
}

function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  return handle(params);
}

function handle(p) {
  try {
    let result;
    if (p.action === 'login') result = login(p);
    else if (p.action === 'addEntry') result = addEntry(p);
    else result = { error: 'Unknown action' };
    return json(result);
  } catch (err) {
    return json({ error: err.message });
  }
}

function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Run this once from the Apps Script editor (function dropdown > setup > Run)
 * to create the Clients and Entries tabs with headers and a sample client.
 */
function setup() {
  const spreadsheet = ss();

  let clients = spreadsheet.getSheetByName(CLIENTS_SHEET);
  if (!clients) clients = spreadsheet.insertSheet(CLIENTS_SHEET);
  clients.clear();
  clients.getRange(1, 1, 1, 4).setValues([['ClientID', 'ClientName', 'AccessCode', 'Rate']]);
  clients.getRange(2, 1, 1, 4).setValues([['C001', 'Sample Client', 'change-me-123', DEFAULT_RATE]]);
  clients.setFrozenRows(1);

  let entries = spreadsheet.getSheetByName(ENTRIES_SHEET);
  if (!entries) entries = spreadsheet.insertSheet(ENTRIES_SHEET);
  entries.clear();
  entries.getRange(1, 1, 1, 6).setValues([['ClientID', 'Date', 'Type', 'Amount', 'Notes', 'Timestamp']]);
  entries.setFrozenRows(1);

  const blank = spreadsheet.getSheetByName('Sheet1');
  if (blank && spreadsheet.getSheets().length > 2) spreadsheet.deleteSheet(blank);

  Logger.log('Setup complete: Clients and Entries tabs are ready.');
}

/**
 * One-time migration: adds a "Rate" column to the Clients tab (default
 * DEFAULT_RATE for existing rows) if it doesn't already exist. Run once
 * from the function dropdown after pulling in this version of the script.
 */
function addRateColumn() {
  const sheet = ss().getSheetByName(CLIENTS_SHEET);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('Rate') !== -1) {
    Logger.log('Rate column already present.');
    return;
  }
  const col = sheet.getLastColumn() + 1;
  sheet.getRange(1, col).setValue('Rate');
  const numRows = sheet.getLastRow() - 1;
  if (numRows > 0) {
    const defaults = Array.from({ length: numRows }, () => [DEFAULT_RATE]);
    sheet.getRange(2, col, numRows, 1).setValues(defaults);
  }
  Logger.log('Rate column added with default ' + DEFAULT_RATE + ' for ' + numRows + ' clients.');
}

/**
 * Compounds `principal` annually at `rate` from `startDate` up to `asOfDate`,
 * then applies simple interest for the partial year since the last
 * anniversary. Matches the original spreadsheet's interest calculation.
 */
function compoundedValue(principal, startDate, rate, asOfDate) {
  let value = principal;
  let cursor = new Date(startDate);
  while (true) {
    const nextAnniversary = new Date(cursor);
    nextAnniversary.setFullYear(cursor.getFullYear() + 1);
    if (nextAnniversary > asOfDate) break;
    value *= (1 + rate);
    cursor = nextAnniversary;
  }
  const msPerDay = 24 * 60 * 60 * 1000;
  const daysSince = Math.max(0, (asOfDate - cursor) / msPerDay);
  return value * (1 + rate * daysSince / 365);
}

/** Sums compounded Investment entries minus face-value Withdrawals. */
function tallyCurrentValue(clientEntries, rate) {
  const today = new Date();
  let total = 0;
  clientEntries.forEach(en => {
    if (en.Type === 'Investment') {
      total += compoundedValue(Number(en.Amount), new Date(en.Date), rate, today);
    } else if (en.Type === 'Withdrawal') {
      total -= Number(en.Amount);
    }
  });
  return total;
}

/**
 * Strips any stored "Current Value" rows and replaces them with a freshly
 * computed one per client, dated today, so the dashboard always shows an
 * up-to-date value without needing manual entries.
 */
function withComputedValues(clients, allEntries) {
  const realEntries = allEntries.filter(en => en.Type !== 'Current Value');
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const synthetic = clients.map(c => {
    const clientEntries = realEntries.filter(en => String(en.ClientID) === String(c.ClientID));
    const rate = Number(c.Rate) || DEFAULT_RATE;
    const value = tallyCurrentValue(clientEntries, rate);
    return {
      ClientID: c.ClientID,
      Date: todayStr,
      Type: 'Current Value',
      Amount: Math.round(value * 100) / 100,
      Notes: 'Auto-calculated at ' + (rate * 100) + '% p.a.',
    };
  });
  return realEntries.concat(synthetic);
}

function sheetToObjects(sheetName) {
  const sheet = ss().getSheetByName(sheetName);
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  return rows
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => {
        const v = r[i];
        obj[h] = v instanceof Date
          ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : v;
      });
      return obj;
    });
}

function login(p) {
  const clients = sheetToObjects(CLIENTS_SHEET);
  const allEntries = sheetToObjects(ENTRIES_SHEET);

  if (p.adminKey === ADMIN_KEY) {
    const entries = withComputedValues(clients, allEntries);
    return { role: 'admin', clients, entries };
  }

  const client = clients.find(
    c => String(c.ClientID) === String(p.clientId) &&
         String(c.AccessCode) === String(p.accessCode)
  );
  if (!client) return { error: 'Invalid Client ID or Access Code' };

  const entries = withComputedValues([client], allEntries).filter(
    en => String(en.ClientID) === String(client.ClientID)
  );
  return { role: 'client', client, entries };
}

function addEntry(p) {
  if (p.adminKey !== ADMIN_KEY) return { error: 'Not authorized' };
  if (!p.clientId || !p.date || !p.type || p.amount === undefined) {
    return { error: 'Missing required fields' };
  }
  const sheet = ss().getSheetByName(ENTRIES_SHEET);
  sheet.appendRow([p.clientId, p.date, p.type, Number(p.amount), p.notes || '', new Date()]);
  return { success: true };
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
