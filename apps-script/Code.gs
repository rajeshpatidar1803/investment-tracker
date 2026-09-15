/**
 * Investment Tracker backend — Google Apps Script Web App.
 *
 * Setup:
 * 1. Create a Google Sheet with two tabs:
 *
 *    "Clients" tab, headers in row 1:
 *      ClientID | ClientName | AccessCode
 *
 *    "Entries" tab, headers in row 1:
 *      ClientID | Date | Type | Amount | Notes | Timestamp
 *      (Type is one of: Investment, Withdrawal, Current Value)
 *
 * 2. Extensions > Apps Script, paste this file in as Code.gs.
 * 3. Set ADMIN_KEY below to a long random string you choose.
 * 4. Deploy > New deployment > type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the deployment URL into js/app.js as WEB_APP_URL.
 */

const CLIENTS_SHEET = 'Clients';
const ENTRIES_SHEET = 'Entries';
const ADMIN_KEY = 'CHANGE_ME_ADMIN_KEY';

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

  if (p.adminKey === ADMIN_KEY) {
    const entries = sheetToObjects(ENTRIES_SHEET);
    return { role: 'admin', clients, entries };
  }

  const client = clients.find(
    c => String(c.ClientID) === String(p.clientId) &&
         String(c.AccessCode) === String(p.accessCode)
  );
  if (!client) return { error: 'Invalid Client ID or Access Code' };

  const entries = sheetToObjects(ENTRIES_SHEET).filter(
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
