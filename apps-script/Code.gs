/**
 * Vegacapital Investment Tracker backend — Google Apps Script Web App.
 *
 * Data model:
 *  - "Clients" tab: ClientID | ClientName | AccessCode | Rate
 *  - "Ledger" tab: ClientID | Date | Type | Amount | Notes | Timestamp | TrancheID
 *      Type is one of: Investment, Withdrawal, Interest Paid, Reinvested
 *      Each "Investment" row starts a new tranche and gets its own TrancheID
 *      (e.g. C001-1, C001-2, ...). "Interest Paid" / "Reinvested" rows record
 *      what happened to a specific tranche's interest at a 1-year cycle
 *      boundary; a boundary with no such row yet defaults to reinvest
 *      (compounding), so nothing breaks if you never take any action.
 *  - "Main" tab: auto-generated report, one row per tranche, rebuilt after
 *    every login/add/settle so it always reflects the current numbers.
 *
 * Setup (fresh install):
 * 1. Open your Google Sheet, then Extensions > Apps Script.
 * 2. Delete the default code, paste this whole file in as Code.gs.
 * 3. Set ADMIN_KEY below to a long random string you choose.
 * 4. Function dropdown > setup > Run. Authorize when prompted.
 * 5. Deploy > New deployment > type "Web app". Execute as: Me. Access: Anyone.
 * 6. Copy the deployment URL into js/app.js as WEB_APP_URL.
 *
 * Upgrading an existing sheet (Entries -> Ledger, adds TrancheID, builds Main):
 *   Function dropdown > upgradeToLedgerSystem > Run, once.
 *
 * Adding the Nifty 50 comparison benchmark tab (one-time, or if you skipped it):
 *   Function dropdown > addBenchmarksTab > Run, once. Edit the "Benchmarks"
 *   tab's CAGR (%) column any time to refresh the figures — no redeploy needed.
 *
 * Adding the "Other Investments" tab (lets clients log MF/Stocks/Gold/Real
 * Estate holdings themselves, from their own dashboard):
 *   Function dropdown > addOtherInvestmentsTab > Run, once.
 */

const CLIENTS_SHEET = 'Clients';
const ENTRIES_SHEET = 'Ledger';
const MAIN_SHEET = 'Main';
const BENCHMARKS_SHEET = 'Benchmarks';
const OTHER_INVESTMENTS_SHEET = 'OtherInvestments';
const ADMIN_KEY = 'CHANGE_ME_ADMIN_KEY';
const DEFAULT_RATE = 0.13; // 13% p.a., used when a client has no Rate column value
const MS_PER_DAY = 24 * 60 * 60 * 1000;

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
    else if (p.action === 'settleCycle') result = settleCycle(p);
    else if (p.action === 'addOtherInvestment') result = addOtherInvestment(p);
    else if (p.action === 'updateOtherInvestmentValue') result = updateOtherInvestmentValue(p);
    else if (p.action === 'deleteOtherInvestment') result = deleteOtherInvestment(p);
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
 * on a brand-new sheet to create the Clients, Ledger and Main tabs.
 */
function setup() {
  const spreadsheet = ss();

  let clients = spreadsheet.getSheetByName(CLIENTS_SHEET);
  if (!clients) clients = spreadsheet.insertSheet(CLIENTS_SHEET);
  clients.clear();
  clients.getRange(1, 1, 1, 4).setValues([['ClientID', 'ClientName', 'AccessCode', 'Rate']]);
  clients.getRange(2, 1, 1, 4).setValues([['C001', 'Sample Client', 'change-me-123', DEFAULT_RATE]]);
  clients.setFrozenRows(1);

  let ledger = spreadsheet.getSheetByName(ENTRIES_SHEET);
  if (!ledger) ledger = spreadsheet.insertSheet(ENTRIES_SHEET);
  ledger.clear();
  ledger.getRange(1, 1, 1, 7).setValues([['ClientID', 'Date', 'Type', 'Amount', 'Notes', 'Timestamp', 'TrancheID']]);
  ledger.setFrozenRows(1);

  const blank = spreadsheet.getSheetByName('Sheet1');
  if (blank && spreadsheet.getSheets().length > 2) spreadsheet.deleteSheet(blank);

  addBenchmarksTab();
  addOtherInvestmentsTab();
  refreshMainTab();
  Logger.log('Setup complete: Clients, Ledger, Main, Benchmarks and OtherInvestments tabs are ready.');
}

/**
 * Creates the "OtherInvestments" tab, where clients log their own holdings
 * outside Vegacapital (Mutual Funds, Stocks, Gold, Real Estate, etc.) for
 * comparison. Safe to run again later.
 */
function addOtherInvestmentsTab() {
  const spreadsheet = ss();
  let sheet = spreadsheet.getSheetByName(OTHER_INVESTMENTS_SHEET);
  if (!sheet) sheet = spreadsheet.insertSheet(OTHER_INVESTMENTS_SHEET);
  if (sheet.getLastRow() > 0) return;
  sheet.getRange(1, 1, 1, 9).setValues([[
    'HoldingID', 'ClientID', 'Type', 'Name', 'InvestedAmount', 'CurrentValue', 'DateAdded', 'Notes', 'LastUpdated',
  ]]);
  sheet.setFrozenRows(1);
  Logger.log('OtherInvestments tab ready.');
}

/** Returns [] instead of throwing when the sheet doesn't exist yet (not migrated). */
function safeSheetToObjects(sheetName) {
  const sheet = ss().getSheetByName(sheetName);
  if (!sheet) return [];
  return sheetToObjects(sheetName);
}

/**
 * Creates (or resets) the "Benchmarks" tab used for the Nifty 50 comparison
 * shown on client dashboards. Safe to run again later — edit the CAGR (%)
 * column any time in the sheet itself to refresh the figures.
 */
function addBenchmarksTab() {
  const spreadsheet = ss();
  let sheet = spreadsheet.getSheetByName(BENCHMARKS_SHEET);
  if (!sheet) sheet = spreadsheet.insertSheet(BENCHMARKS_SHEET);
  if (sheet.getLastRow() > 0) return; // don't overwrite figures you've already edited

  const note = 'Nifty 50 trailing CAGR (price return), data through Sep-2026 '
    + '(source: returnscreener.com). This is an actual point-in-time trailing '
    + 'return, not a long-run average — it moves with the market. Edit the '
    + 'CAGR column here any time to refresh — no redeploy needed.';
  sheet.getRange(1, 1, 1, 3).setValues([['Period', 'CAGR (%)', 'Notes']]);
  sheet.getRange(2, 1, 3, 3).setValues([
    ['1Y', -6.4, note],
    ['3Y', 5.4, note],
    ['5Y', 6.2, note],
  ]);
  sheet.setFrozenRows(1);
  Logger.log('Benchmarks tab ready.');
}

function getBenchmarks() {
  const sheet = ss().getSheetByName(BENCHMARKS_SHEET);
  if (!sheet) return null;
  const rows = sheet.getDataRange().getValues();
  rows.shift();
  const map = {};
  let note = '';
  rows.forEach(r => {
    if (r[0]) {
      map[r[0]] = Number(r[1]);
      if (r[2]) note = r[2];
    }
  });
  return { oneYear: map['1Y'], threeYear: map['3Y'], fiveYear: map['5Y'], note };
}

/** Weighted-average effective ROI across one client's tranches, weighted by initial investment. */
function blendedRoi(report, clientId) {
  const rows = report.filter(r => String(r.ClientID) === String(clientId));
  const totalInitial = rows.reduce((sum, r) => sum + r.InitialInvestment, 0);
  if (!totalInitial) return null;
  const weighted = rows.reduce((sum, r) => sum + r.EffectiveROIPercent * r.InitialInvestment, 0);
  return Math.round((weighted / totalInitial) * 100) / 100;
}

/**
 * One-time upgrade for a sheet still using the old "Entries" tab: renames
 * it to "Ledger", adds the TrancheID column, backfills TrancheIDs for
 * existing Investment rows, and builds the Main report tab. Safe to run
 * more than once.
 */
function upgradeToLedgerSystem() {
  const spreadsheet = ss();

  const oldEntries = spreadsheet.getSheetByName('Entries');
  if (oldEntries && !spreadsheet.getSheetByName(ENTRIES_SHEET)) {
    oldEntries.setName(ENTRIES_SHEET);
  }

  const ledger = spreadsheet.getSheetByName(ENTRIES_SHEET);
  if (!ledger) {
    Logger.log('No Ledger/Entries sheet found — run setup() instead for a fresh install.');
    return;
  }

  const lastCol = ledger.getLastColumn();
  const headers = ledger.getRange(1, 1, 1, Math.max(lastCol, 7)).getValues()[0];
  if (headers[6] !== 'TrancheID') {
    ledger.getRange(1, 7).setValue('TrancheID');
  }

  const clientsSheet = spreadsheet.getSheetByName(CLIENTS_SHEET);
  const clientHeaders = clientsSheet.getRange(1, 1, 1, clientsSheet.getLastColumn()).getValues()[0];
  if (clientHeaders.indexOf('Rate') === -1) {
    const col = clientsSheet.getLastColumn() + 1;
    clientsSheet.getRange(1, col).setValue('Rate');
    const numRows = clientsSheet.getLastRow() - 1;
    if (numRows > 0) {
      const defaults = Array.from({ length: numRows }, () => [DEFAULT_RATE]);
      clientsSheet.getRange(2, col, numRows, 1).setValues(defaults);
    }
  }

  const lastRow = ledger.getLastRow();
  if (lastRow > 1) {
    const data = ledger.getRange(2, 1, lastRow - 1, 7).getValues();
    const counters = {};
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const clientId = row[0];
      const type = row[2];
      let trancheId = row[6];
      if (type === 'Investment') {
        if (!trancheId) {
          counters[clientId] = (counters[clientId] || 0) + 1;
          trancheId = clientId + '-' + counters[clientId];
          ledger.getRange(2 + i, 7).setValue(trancheId);
        } else {
          const n = parseInt(String(trancheId).split('-').pop(), 10) || 0;
          counters[clientId] = Math.max(counters[clientId] || 0, n);
        }
      }
    }
  }

  addBenchmarksTab();
  refreshMainTab();
  Logger.log('Upgrade complete: Ledger tab ready with TrancheID, Main and Benchmarks tabs built.');
}

function addYears(date, n) {
  const d = new Date(date);
  d.setFullYear(d.getFullYear() + n);
  return d;
}

/**
 * Walks a tranche's yearly cycle boundaries from its start date up to
 * `asOfDate`. At each boundary, applies the matching logged decision
 * (by position, in chronological order) if one exists; otherwise defaults
 * to reinvest (compounding), which is also what an explicit "Reinvested"
 * entry does. An "Interest Paid" entry leaves the principal unchanged for
 * that cycle instead.
 */
function walkTrancheCycles(startDate, initialAmount, rate, decisions, asOfDate) {
  let principal = Number(initialAmount);
  let cycleStart = new Date(startDate);
  let idx = 0;
  while (true) {
    const nextBoundary = addYears(cycleStart, 1);
    if (nextBoundary > asOfDate) break;
    const decision = decisions[idx];
    const interest = principal * rate;
    if (!decision || decision.Type !== 'Interest Paid') {
      principal += interest;
    }
    cycleStart = nextBoundary;
    idx++;
  }
  const daysIntoCycle = Math.max(0, (asOfDate - cycleStart) / MS_PER_DAY);
  const accrued = principal * rate * daysIntoCycle / 365;
  return {
    cycleStartDate: cycleStart,
    cycleStartPrincipal: principal,
    accruedThisCycle: accrued,
    totalValue: principal + accrued,
    pendingCount: Math.max(0, idx - decisions.length),
  };
}

/**
 * Finds the earliest cycle boundary that has occurred but has no decision
 * logged yet, and returns what settling it now would look like. Returns
 * null if no cycle is currently due.
 */
function findCycleToSettle(startDate, initialAmount, rate, decisions, asOfDate) {
  let principal = Number(initialAmount);
  let cycleStart = new Date(startDate);
  let idx = 0;
  while (true) {
    const nextBoundary = addYears(cycleStart, 1);
    if (nextBoundary > asOfDate) return null;
    if (idx === decisions.length) {
      return {
        cycleStartDate: cycleStart,
        cycleEndDate: nextBoundary,
        principalAtCycleStart: principal,
        interestAmount: principal * rate,
      };
    }
    const decision = decisions[idx];
    const interest = principal * rate;
    if (!decision || decision.Type !== 'Interest Paid') principal += interest;
    cycleStart = nextBoundary;
    idx++;
  }
}

function decisionsForTranche(allEntries, trancheId) {
  return allEntries
    .filter(e => e.TrancheID && String(e.TrancheID) === String(trancheId) &&
      (e.Type === 'Interest Paid' || e.Type === 'Reinvested'))
    .sort((a, b) => new Date(a.Date) - new Date(b.Date));
}

/**
 * Annualized rate that discounts a series of dated cash flows to zero NPV
 * (standard XIRR). Used for effective ROI so that interest paid out in cash
 * is credited at the time it was received, instead of vanishing from the
 * return just because it's no longer sitting in the tranche.
 */
function xirr(cashflows) {
  const t0 = cashflows[0].date;
  const years = d => (d - t0) / MS_PER_DAY / 365;
  const npv = rate => cashflows.reduce((sum, cf) => sum + cf.amount / Math.pow(1 + rate, years(cf.date)), 0);

  let lo = -0.5, hi = 5;
  let npvLo = npv(lo), npvHi = npv(hi);
  if (npvLo * npvHi > 0) return null;
  for (let i = 0; i < 100; i++) {
    const mid = (lo + hi) / 2;
    const npvMid = npv(mid);
    if (Math.abs(npvMid) < 1e-6) return mid;
    if ((npvLo < 0) === (npvMid < 0)) { lo = mid; npvLo = npvMid; } else { hi = mid; }
  }
  return (lo + hi) / 2;
}

/** One row per tranche with everything needed for the Main tab and the UI. */
function computeTrancheReport(clients, allEntries) {
  const today = new Date();
  const investments = allEntries.filter(e => e.Type === 'Investment' && e.TrancheID);
  return investments.map(inv => {
    const client = clients.find(c => String(c.ClientID) === String(inv.ClientID));
    const rate = Number(client && client.Rate) || DEFAULT_RATE;
    const decisions = decisionsForTranche(allEntries, inv.TrancheID);
    const startDate = new Date(inv.Date);
    const walk = walkTrancheCycles(startDate, inv.Amount, rate, decisions, today);
    const ageDays = Math.max(1, Math.round((today - startDate) / MS_PER_DAY));

    const paidFlows = decisions
      .filter(d => d.Type === 'Interest Paid')
      .map(d => ({ amount: Number(d.Amount), date: new Date(d.Date) }));
    const totalPaidOut = paidFlows.reduce((sum, f) => sum + f.amount, 0);

    const cashflows = [{ amount: -Number(inv.Amount), date: startDate }]
      .concat(paidFlows)
      .concat([{ amount: walk.totalValue, date: today }]);
    const xirrRate = xirr(cashflows);
    // Fallback (shouldn't normally trigger): plain CAGR of the remaining balance only.
    const effectiveRoi = xirrRate !== null
      ? xirrRate * 100
      : (Math.pow(walk.totalValue / Number(inv.Amount), 365 / ageDays) - 1) * 100;

    const totalReturnPercent = ((walk.totalValue + totalPaidOut) / Number(inv.Amount) - 1) * 100;

    let pendingPreview = null;
    if (walk.pendingCount > 0) {
      const cycle = findCycleToSettle(startDate, inv.Amount, rate, decisions, today);
      if (cycle) pendingPreview = Math.round(cycle.interestAmount * 100) / 100;
    }

    return {
      ClientID: inv.ClientID,
      ClientName: client ? client.ClientName : '',
      TrancheID: inv.TrancheID,
      InitialInvestment: Number(inv.Amount),
      InitialDate: inv.Date,
      CurrentCycleAmount: Math.round(walk.cycleStartPrincipal * 100) / 100,
      CurrentCycleStartDate: Utilities.formatDate(walk.cycleStartDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      InterestAccruedCurrentCycle: Math.round(walk.accruedThisCycle * 100) / 100,
      TotalValue: Math.round(walk.totalValue * 100) / 100,
      TotalPaidOut: Math.round(totalPaidOut * 100) / 100,
      AgeDays: ageDays,
      EffectiveROIPercent: Math.round(effectiveRoi * 100) / 100,
      TotalReturnPercent: Math.round(totalReturnPercent * 100) / 100,
      PendingCycles: walk.pendingCount,
      PendingInterestPreview: pendingPreview,
    };
  });
}

function refreshMainTab() {
  const clients = sheetToObjects(CLIENTS_SHEET);
  const allEntries = sheetToObjects(ENTRIES_SHEET);
  const report = computeTrancheReport(clients, allEntries);

  const spreadsheet = ss();
  let sheet = spreadsheet.getSheetByName(MAIN_SHEET);
  if (!sheet) sheet = spreadsheet.insertSheet(MAIN_SHEET);
  sheet.clear();

  const headers = ['Investor Name', 'ClientID', 'TrancheID', 'Initial Investment', 'Date of Initial Investment',
    'Current Cycle Invested Amount', 'Current Cycle Investment Date', 'Interest Accrued (Current Cycle)',
    'Total Value', 'Total Interest Paid Out', 'Investment Age (days)', 'Effective ROI % (annualized)',
    'Total Return % (cumulative)', 'Cycles Awaiting Decision'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  if (report.length) {
    const rows = report.map(r => [r.ClientName, r.ClientID, r.TrancheID, r.InitialInvestment, r.InitialDate,
      r.CurrentCycleAmount, r.CurrentCycleStartDate, r.InterestAccruedCurrentCycle, r.TotalValue, r.TotalPaidOut,
      r.AgeDays, r.EffectiveROIPercent, r.TotalReturnPercent, r.PendingCycles]);
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
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

/**
 * Strips any stored "Current Value" rows and replaces them with a freshly
 * computed one per client (sum of all their tranches' current value, minus
 * withdrawals), dated today.
 */
function withComputedValues(clients, allEntries) {
  const realEntries = allEntries.filter(en => en.Type !== 'Current Value');
  const report = computeTrancheReport(clients, realEntries);
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const synthetic = clients.map(c => {
    const trancheTotal = report
      .filter(r => String(r.ClientID) === String(c.ClientID))
      .reduce((sum, r) => sum + r.TotalValue, 0);
    const withdrawals = realEntries
      .filter(e => String(e.ClientID) === String(c.ClientID) && e.Type === 'Withdrawal')
      .reduce((sum, e) => sum + Number(e.Amount), 0);
    const value = trancheTotal - withdrawals;
    return {
      ClientID: c.ClientID,
      Date: todayStr,
      Type: 'Current Value',
      Amount: Math.round(value * 100) / 100,
      Notes: 'Auto-calculated',
    };
  });
  return realEntries.concat(synthetic);
}

function login(p) {
  const clients = sheetToObjects(CLIENTS_SHEET);
  const allEntries = sheetToObjects(ENTRIES_SHEET);

  const benchmarks = getBenchmarks();

  if (p.adminKey === ADMIN_KEY) {
    const entries = withComputedValues(clients, allEntries);
    const realEntries = allEntries.filter(e => e.Type !== 'Current Value');
    const report = computeTrancheReport(clients, realEntries);
    const dueCycles = report.filter(r => r.PendingCycles > 0);
    const totalWithdrawals = realEntries
      .filter(e => e.Type === 'Withdrawal')
      .reduce((sum, e) => sum + Number(e.Amount), 0);
    const round2 = n => Math.round(n * 100) / 100;
    const summary = {
      totalClients: clients.length,
      totalInitialInvestment: round2(report.reduce((sum, r) => sum + r.InitialInvestment, 0)),
      totalCurrentCycleInvestment: round2(report.reduce((sum, r) => sum + r.CurrentCycleAmount, 0)),
      totalAccruedInterest: round2(report.reduce((sum, r) => sum + r.InterestAccruedCurrentCycle, 0)),
      totalOutstanding: round2(report.reduce((sum, r) => sum + r.TotalValue, 0) - totalWithdrawals),
    };
    const clientROI = {};
    clients.forEach(c => { clientROI[c.ClientID] = blendedRoi(report, c.ClientID); });
    const otherInvestments = computeOtherInvestmentReport(safeSheetToObjects(OTHER_INVESTMENTS_SHEET));
    return { role: 'admin', clients, entries, dueCycles, summary, clientROI, benchmarks, otherInvestments };
  }

  const client = authenticateClient(p.clientId, p.accessCode);
  if (!client) return { error: 'Invalid Client ID or Access Code' };

  const realEntries = allEntries.filter(e => e.Type !== 'Current Value');
  const report = computeTrancheReport(clients, realEntries);
  const portfolioROI = blendedRoi(report, client.ClientID);
  const entries = withComputedValues([client], allEntries).filter(
    en => String(en.ClientID) === String(client.ClientID)
  );

  const otherInvestments = computeOtherInvestmentReport(safeSheetToObjects(OTHER_INVESTMENTS_SHEET))
    .filter(h => String(h.ClientID) === String(client.ClientID));
  const otherTypeROI = {};
  OTHER_INVESTMENT_TYPES.forEach(t => { otherTypeROI[t] = blendedTypeROI(otherInvestments, t); });

  return { role: 'client', client, entries, portfolioROI, benchmarks, otherInvestments, otherTypeROI };
}

function authenticateClient(clientId, accessCode) {
  const clients = sheetToObjects(CLIENTS_SHEET);
  return clients.find(
    c => String(c.ClientID) === String(clientId) && String(c.AccessCode) === String(accessCode)
  ) || null;
}

const OTHER_INVESTMENT_TYPES = ['Mutual Fund', 'Stocks', 'Gold', 'Real Estate', 'Other'];

/** Adds annualized CAGR and simple cumulative gain % to each holding row. */
function computeOtherInvestmentReport(holdings) {
  const today = new Date();
  return holdings.map(h => {
    const invested = Number(h.InvestedAmount);
    const current = Number(h.CurrentValue);
    const dateAdded = new Date(h.DateAdded);
    const ageDays = Math.max(1, Math.round((today - dateAdded) / MS_PER_DAY));
    const cagr = invested > 0 ? (Math.pow(current / invested, 365 / ageDays) - 1) * 100 : null;
    const gainPercent = invested > 0 ? (current / invested - 1) * 100 : null;
    return Object.assign({}, h, {
      InvestedAmount: invested,
      CurrentValue: current,
      AgeDays: ageDays,
      CAGRPercent: cagr !== null ? Math.round(cagr * 100) / 100 : null,
      GainPercent: gainPercent !== null ? Math.round(gainPercent * 100) / 100 : null,
    });
  });
}

/** Weighted-average CAGR (by invested amount) across a client's holdings of one type. */
function blendedTypeROI(report, type) {
  const rows = report.filter(r => r.Type === type && r.CAGRPercent !== null);
  const totalInvested = rows.reduce((sum, r) => sum + r.InvestedAmount, 0);
  if (!totalInvested) return null;
  const weighted = rows.reduce((sum, r) => sum + r.CAGRPercent * r.InvestedAmount, 0);
  return Math.round((weighted / totalInvested) * 100) / 100;
}

function nextHoldingId(clientId, allHoldings) {
  const count = allHoldings.filter(h => String(h.ClientID) === String(clientId)).length;
  return clientId + '-H' + (count + 1);
}

/** Client action: log a new holding outside Vegacapital (MF, Stocks, Gold, Real Estate, ...). */
function addOtherInvestment(p) {
  const client = authenticateClient(p.clientId, p.accessCode);
  if (!client) return { error: 'Invalid Client ID or Access Code' };
  if (!p.type || !p.name || p.investedAmount === undefined || p.currentValue === undefined || !p.date) {
    return { error: 'Missing required fields' };
  }
  const sheet = ss().getSheetByName(OTHER_INVESTMENTS_SHEET);
  if (!sheet) return { error: 'OtherInvestments tab not set up yet. Ask your advisor to run addOtherInvestmentsTab().' };
  const holdingId = nextHoldingId(p.clientId, safeSheetToObjects(OTHER_INVESTMENTS_SHEET));
  const now = new Date();
  sheet.appendRow([holdingId, p.clientId, p.type, p.name, Number(p.investedAmount), Number(p.currentValue),
    p.date, p.notes || '', now]);
  return { success: true, holdingId };
}

/** Client action: update a holding's current value (e.g. checked its latest worth elsewhere). */
function updateOtherInvestmentValue(p) {
  const client = authenticateClient(p.clientId, p.accessCode);
  if (!client) return { error: 'Invalid Client ID or Access Code' };
  if (!p.holdingId || p.currentValue === undefined) return { error: 'Missing required fields' };
  const sheet = ss().getSheetByName(OTHER_INVESTMENTS_SHEET);
  if (!sheet) return { error: 'Holding not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('HoldingID');
  const clientCol = headers.indexOf('ClientID');
  const valueCol = headers.indexOf('CurrentValue');
  const updatedCol = headers.indexOf('LastUpdated');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(p.holdingId) && String(data[i][clientCol]) === String(p.clientId)) {
      sheet.getRange(i + 1, valueCol + 1).setValue(Number(p.currentValue));
      sheet.getRange(i + 1, updatedCol + 1).setValue(new Date());
      return { success: true };
    }
  }
  return { error: 'Holding not found' };
}

/** Client action: remove a holding they no longer want tracked. */
function deleteOtherInvestment(p) {
  const client = authenticateClient(p.clientId, p.accessCode);
  if (!client) return { error: 'Invalid Client ID or Access Code' };
  if (!p.holdingId) return { error: 'Missing holdingId' };
  const sheet = ss().getSheetByName(OTHER_INVESTMENTS_SHEET);
  if (!sheet) return { error: 'Holding not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('HoldingID');
  const clientCol = headers.indexOf('ClientID');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(p.holdingId) && String(data[i][clientCol]) === String(p.clientId)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Holding not found' };
}

function nextTrancheId(clientId, allEntries) {
  const count = allEntries.filter(
    e => String(e.ClientID) === String(clientId) && e.Type === 'Investment'
  ).length;
  return clientId + '-' + (count + 1);
}

function addEntry(p) {
  if (p.adminKey !== ADMIN_KEY) return { error: 'Not authorized' };
  if (!p.clientId || !p.date || !p.type || p.amount === undefined) {
    return { error: 'Missing required fields' };
  }
  const sheet = ss().getSheetByName(ENTRIES_SHEET);
  let trancheId = '';
  if (p.type === 'Investment') {
    trancheId = nextTrancheId(p.clientId, sheetToObjects(ENTRIES_SHEET));
  }
  sheet.appendRow([p.clientId, p.date, p.type, Number(p.amount), p.notes || '', new Date(), trancheId]);
  refreshMainTab();
  return { success: true, trancheId };
}

/**
 * Advisor action: settle the earliest interest cycle currently due for a
 * tranche, either paying it out (principal unchanged) or reinvesting it
 * (principal grows) — this is also what happens by default if you never
 * call this, so calling it is about keeping an explicit record, and about
 * correcting a tranche where you actually paid cash instead of letting it
 * compound.
 */
function settleCycle(p) {
  if (p.adminKey !== ADMIN_KEY) return { error: 'Not authorized' };
  if (!p.trancheId || (p.decision !== 'paid' && p.decision !== 'reinvested')) {
    return { error: 'Missing or invalid trancheId/decision' };
  }

  const allEntries = sheetToObjects(ENTRIES_SHEET);
  const invEntry = allEntries.find(e => e.Type === 'Investment' && String(e.TrancheID) === String(p.trancheId));
  if (!invEntry) return { error: 'Tranche not found' };

  const clients = sheetToObjects(CLIENTS_SHEET);
  const client = clients.find(c => String(c.ClientID) === String(invEntry.ClientID));
  const rate = Number(client && client.Rate) || DEFAULT_RATE;

  const decisions = decisionsForTranche(allEntries, p.trancheId);
  const today = new Date();
  const cycle = findCycleToSettle(new Date(invEntry.Date), Number(invEntry.Amount), rate, decisions, today);
  if (!cycle) return { error: 'No cycle is currently due for this investment.' };

  const type = p.decision === 'paid' ? 'Interest Paid' : 'Reinvested';
  const dateStr = Utilities.formatDate(cycle.cycleEndDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const cycleStartStr = Utilities.formatDate(cycle.cycleStartDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const amount = Math.round(cycle.interestAmount * 100) / 100;
  const notes = (type === 'Interest Paid' ? 'Interest paid out' : 'Interest reinvested')
    + ' for cycle starting ' + cycleStartStr;

  const sheet = ss().getSheetByName(ENTRIES_SHEET);
  sheet.appendRow([invEntry.ClientID, dateStr, type, amount, notes, new Date(), p.trancheId]);
  refreshMainTab();
  return { success: true, amount };
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
