// Paste your deployed Apps Script Web App URL here (see apps-script/Code.gs).
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwZXEJsYmuM6DF5j0_4VDkZo5xqNRkm1YiQaxbI1CGt0DU5m2AmjsXpIEJcFZ5fpZoU/exec';

const state = {
  role: null,       // 'admin' | 'client'
  adminKey: null,
  clientAccessCode: null, // client only: kept so they can add/edit/delete their own other-investments
  clients: [],
  entries: [],       // entries currently displayed
  dueCycles: [],      // admin only: tranches with an interest cycle awaiting a decision
  summary: null,      // admin only: portfolio-wide totals
  clientROI: {},      // admin only: blended effective ROI % per ClientID
  portfolioROI: null, // client only: their own blended effective ROI %
  benchmarks: null,   // Nifty 50 comparison figures (both roles)
  otherInvestments: [], // admin: all clients' holdings; client: just their own
  currentClientId: null,
  currentClientName: '',
};

const OTHER_INVESTMENT_COLORS = {
  'Mutual Fund': '#5b3fae',
  'Stocks': '#2563a8',
  'Gold': '#b8860b',
  'Real Estate': '#52616b',
  'Other': '#667085',
};

const $ = (id) => document.getElementById(id);

// ---------- API ----------

async function api(payload) {
  const res = await fetch(WEB_APP_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // avoids CORS preflight on Apps Script
    body: JSON.stringify(payload),
  });
  if (!res.ok) throw new Error('Network error');
  return res.json();
}

// ---------- Tabs on login screen ----------

document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    $(btn.dataset.tab === 'clientLogin' ? 'clientLoginForm' : 'adminLoginForm').classList.add('active');
  });
});

// ---------- Login ----------

$('clientLoginForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  await doLogin({
    action: 'login',
    clientId: $('clientId').value.trim(),
    accessCode: $('accessCode').value.trim(),
  });
});

$('adminLoginForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  await doLogin({
    action: 'login',
    adminKey: $('adminKey').value.trim(),
  });
});

async function doLogin(payload, opts = {}) {
  setLoading(true);
  hideError();
  try {
    const res = await api(payload);
    if (res.error) {
      // A restored session with stale credentials just drops back to the login screen.
      if (opts.silent) clearSession(); else showError(res.error);
      return;
    }
    saveSession(payload);
    state.role = res.role;
    if (res.role === 'admin') {
      state.adminKey = payload.adminKey;
      state.clients = res.clients;
      state.entries = res.entries;
      state.dueCycles = res.dueCycles || [];
      state.summary = res.summary || null;
      state.clientROI = res.clientROI || {};
      state.benchmarks = res.benchmarks || null;
      state.otherInvestments = res.otherInvestments || [];
      enterAdminDashboard();
    } else {
      state.currentClientId = res.client.ClientID;
      state.currentClientName = res.client.ClientName;
      state.clientAccessCode = payload.accessCode;
      state.entries = res.entries;
      state.portfolioROI = res.portfolioROI ?? null;
      state.benchmarks = res.benchmarks || null;
      state.otherInvestments = res.otherInvestments || [];
      enterClientDashboard();
    }
  } catch (err) {
    if (!opts.silent) showError('Could not reach the server. Check WEB_APP_URL in js/app.js.');
  } finally {
    setLoading(false);
  }
}

// ---------- Session persistence ----------
// Kept in sessionStorage so a refresh doesn't log you out. It clears itself
// when the tab is closed, so credentials don't linger on a shared machine.

const SESSION_KEY = 'vega-session';

function saveSession(payload) {
  try { sessionStorage.setItem(SESSION_KEY, JSON.stringify(payload)); } catch (err) { /* private mode */ }
}

function clearSession() {
  try { sessionStorage.removeItem(SESSION_KEY); } catch (err) { /* private mode */ }
}

function restoreSession() {
  let saved = null;
  try {
    const raw = sessionStorage.getItem(SESSION_KEY);
    saved = raw ? JSON.parse(raw) : null;
  } catch (err) {
    return;
  }
  if (saved && saved.action === 'login') doLogin(saved, { silent: true });
}

// ---------- Dashboard entry points ----------

function enterAdminDashboard(preserveClientId) {
  $('loginView').hidden = true;
  $('dashboardView').hidden = false;
  $('adminControls').hidden = false;
  $('logoutBtn').hidden = false;

  const select = $('clientSelect');
  select.innerHTML = state.clients
    .map(c => `<option value="${c.ClientID}">${c.ClientName} (${c.ClientID})</option>`)
    .join('');
  select.onchange = renderForSelectedClient;
  if (preserveClientId && state.clients.some(c => String(c.ClientID) === String(preserveClientId))) {
    select.value = preserveClientId;
  }

  $('entryDate').valueAsDate = new Date();

  renderSummary();
  renderDueCycles();
  renderClientsTable();

  if (state.clients.length) {
    renderForSelectedClient();
  }
}

/** Per-client totals, using the same math as the single-client view so the numbers agree. */
function clientTotals(clientId) {
  let invested = 0, withdrawn = 0, interestPaid = 0, currentValue = 0;
  for (const en of state.entries) {
    if (String(en.ClientID) !== String(clientId)) continue;
    const amt = Number(en.Amount) || 0;
    if (en.Type === 'Investment') invested += amt;
    else if (en.Type === 'Withdrawal') withdrawn += amt;
    else if (en.Type === 'Interest Paid') interestPaid += amt;
    else if (en.Type === 'Current Value') currentValue = amt;
  }
  const net = invested - withdrawn;
  return { net, currentValue, gain: (currentValue + interestPaid) - net };
}

function renderClientsTable() {
  const section = $('clientsTableSection');
  if (!state.clients.length) {
    section.hidden = true;
    return;
  }
  section.hidden = false;

  const rows = state.clients
    .map(c => {
      const totals = clientTotals(c.ClientID);
      const cyclesDue = state.dueCycles
        .filter(d => String(d.ClientID) === String(c.ClientID))
        .reduce((sum, d) => sum + (d.PendingCycles || 0), 0);
      return { client: c, ...totals, roi: state.clientROI[c.ClientID], cyclesDue };
    })
    .sort((a, b) => b.currentValue - a.currentValue);

  $('clientsTableBody').innerHTML = rows.map(r => `
    <tr data-client="${r.client.ClientID}">
      <td>${r.client.ClientName} <span class="hint">(${r.client.ClientID})</span></td>
      <td>${fmt(r.net)}</td>
      <td>${fmt(r.currentValue)}</td>
      <td class="${r.gain >= 0 ? 'gain-positive' : 'gain-negative'}">${fmt(r.gain)}</td>
      <td>${r.roi === null || r.roi === undefined ? '-' : r.roi.toFixed(1) + '%'}</td>
      <td>${r.cyclesDue || '-'}</td>
    </tr>
  `).join('');
}

$('clientsTableBody').addEventListener('click', (e) => {
  const row = e.target.closest('tr[data-client]');
  if (!row) return;
  $('clientSelect').value = row.dataset.client;
  renderForSelectedClient();
  $('clientNameHeading').scrollIntoView({ behavior: 'smooth', block: 'start' });
});

function renderSummary() {
  const s = state.summary;
  const section = $('summarySection');
  if (!s) {
    section.hidden = true;
    return;
  }
  section.hidden = false;
  $('summaryOutstanding').textContent = fmt(s.totalOutstanding);
  $('summaryClients').textContent = s.totalClients;
  $('summaryInitial').textContent = fmt(s.totalInitialInvestment);
  $('summaryCycle').textContent = fmt(s.totalCurrentCycleInvestment);
  $('summaryAccrued').textContent = fmt(s.totalAccruedInterest);
}

function renderDueCycles() {
  const section = $('dueCyclesSection');
  const list = $('dueCyclesList');
  if (!state.dueCycles.length) {
    section.hidden = true;
    list.innerHTML = '';
    return;
  }
  section.hidden = false;
  list.innerHTML = state.dueCycles.map(d => `
    <div class="due-cycle-row">
      <div>
        <strong>${d.ClientName} (${d.ClientID})</strong> &middot; tranche ${d.TrancheID}
        <div class="hint">Invested ${fmt(d.InitialInvestment)} on ${d.InitialDate}
          &middot; ${d.PendingCycles} cycle${d.PendingCycles > 1 ? 's' : ''} awaiting a decision
          &middot; next cycle interest ${d.PendingInterestPreview !== null ? fmt(d.PendingInterestPreview) : '-'}
        </div>
      </div>
      <div class="due-cycle-actions">
        <button type="button" class="btn-ghost" data-tranche="${d.TrancheID}" data-decision="paid">Mark Paid</button>
        <button type="button" class="btn-primary" data-tranche="${d.TrancheID}" data-decision="reinvested">Reinvest</button>
      </div>
    </div>
  `).join('');
}

$('dueCyclesList').addEventListener('click', async (e) => {
  const btn = e.target.closest('button[data-tranche]');
  if (!btn) return;
  const trancheId = btn.dataset.tranche;
  const decision = btn.dataset.decision;
  const row = btn.closest('.due-cycle-row');
  const buttons = row.querySelectorAll('button');
  buttons.forEach(b => b.disabled = true);

  try {
    const res = await api({ action: 'settleCycle', adminKey: state.adminKey, trancheId, decision });
    if (res.error) {
      alert(res.error);
      buttons.forEach(b => b.disabled = false);
      return;
    }
    const currentClientId = $('clientSelect').value;
    const refreshed = await api({ action: 'login', adminKey: state.adminKey });
    if (refreshed.error) {
      alert(refreshed.error);
      return;
    }
    state.clients = refreshed.clients;
    state.entries = refreshed.entries;
    state.dueCycles = refreshed.dueCycles || [];
    state.summary = refreshed.summary || null;
    state.clientROI = refreshed.clientROI || {};
    state.benchmarks = refreshed.benchmarks || null;
    state.otherInvestments = refreshed.otherInvestments || [];
    enterAdminDashboard(currentClientId);
  } catch (err) {
    alert('Could not save this action. Please try again.');
    buttons.forEach(b => b.disabled = false);
  }
});

function renderForSelectedClient() {
  const clientId = $('clientSelect').value;
  const client = state.clients.find(c => String(c.ClientID) === String(clientId));
  const entries = state.entries.filter(en => String(en.ClientID) === String(clientId));
  $('clientNameHeading').textContent = client ? `${client.ClientName} (${client.ClientID})` : '';
  renderSummaryAndTable(entries, state.clientROI[clientId], clientId);
  renderOtherInvestments(clientId, false);
}

function enterClientDashboard() {
  $('loginView').hidden = true;
  $('dashboardView').hidden = false;
  $('adminControls').hidden = true;
  $('logoutBtn').hidden = false;
  $('clientNameHeading').textContent = `${state.currentClientName} (${state.currentClientId})`;
  renderSummaryAndTable(state.entries, state.portfolioROI, state.currentClientId);
  renderOtherInvestments(state.currentClientId, true);
}

// ---------- Add entry (admin only) ----------

$('addEntryForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const clientId = $('clientSelect').value;
  if (!clientId) return;

  const payload = {
    action: 'addEntry',
    adminKey: state.adminKey,
    clientId,
    date: $('entryDate').value,
    type: $('entryType').value,
    amount: $('entryAmount').value,
    notes: $('entryNotes').value.trim(),
  };

  const msg = $('addEntryMsg');
  msg.hidden = false;
  msg.textContent = 'Saving...';

  try {
    const res = await api(payload);
    if (res.error) {
      msg.textContent = res.error;
      return;
    }
    const refreshed = await api({ action: 'login', adminKey: state.adminKey });
    if (refreshed.error) {
      msg.textContent = refreshed.error;
      return;
    }
    state.clients = refreshed.clients;
    state.entries = refreshed.entries;
    state.dueCycles = refreshed.dueCycles || [];
    state.summary = refreshed.summary || null;
    state.clientROI = refreshed.clientROI || {};
    state.benchmarks = refreshed.benchmarks || null;
    state.otherInvestments = refreshed.otherInvestments || [];
    enterAdminDashboard(clientId);
    $('addEntryForm').reset();
    $('entryDate').valueAsDate = new Date();
    msg.hidden = false;
    msg.textContent = 'Entry added.';
    setTimeout(() => { msg.hidden = true; }, 2500);
  } catch (err) {
    msg.textContent = 'Could not save entry.';
  }
});

// ---------- Rendering ----------

function renderSummaryAndTable(entries, roi, clientId) {
  const sorted = [...entries].sort((a, b) => new Date(a.Date) - new Date(b.Date));

  let invested = 0, withdrawn = 0, interestPaid = 0, latestValue = null;
  for (const en of sorted) {
    const amt = Number(en.Amount) || 0;
    if (en.Type === 'Investment') invested += amt;
    else if (en.Type === 'Withdrawal') withdrawn += amt;
    else if (en.Type === 'Interest Paid') interestPaid += amt;
    else if (en.Type === 'Current Value') latestValue = amt; // last one wins, list is sorted by date
  }
  const net = invested - withdrawn;
  const currentValue = latestValue !== null ? latestValue : net;
  // Gain/Loss = growth of what's still invested, plus interest already paid out in
  // cash (it's still money you earned — it just isn't sitting in the balance above).
  const gain = (currentValue + interestPaid) - net;

  $('heroValue').textContent = fmt(currentValue);
  $('sumInvested').textContent = fmt(invested);
  $('sumWithdrawn').textContent = fmt(withdrawn);
  $('sumNet').textContent = fmt(net);

  const gainEl = $('sumGain');
  const gainPct = net !== 0 ? ` (${(gain / net * 100).toFixed(1)}%)` : '';
  gainEl.textContent = fmt(gain) + gainPct;
  gainEl.classList.toggle('positive', gain > 0);
  gainEl.classList.toggle('negative', gain < 0);

  const body = $('entriesBody');
  const descending = [...sorted].reverse();
  body.innerHTML = descending.map(en => `
    <tr>
      <td>${en.Date}</td>
      <td><span class="type-badge ${typeBadgeClass(en.Type)}">${en.Type}</span></td>
      <td>${fmt(Number(en.Amount))}</td>
      <td>${en.TrancheID || ''}</td>
      <td>${en.Notes || ''}</td>
    </tr>
  `).join('');

  $('noEntries').hidden = descending.length > 0;
  $('entriesTable').hidden = descending.length === 0;

  renderROICompare(roi, clientId);
}

function typeBadgeClass(type) {
  return {
    'Investment': 'investment',
    'Withdrawal': 'withdrawal',
    'Interest Paid': 'interest-paid',
    'Reinvested': 'reinvested',
    'Current Value': 'current-value',
  }[type] || '';
}

/** Weighted-average CAGR (by invested amount) for one investment type, from a holdings list. */
function computeTypeROI(holdings, type) {
  const rows = holdings.filter(h => h.Type === type && h.CAGRPercent !== null && h.CAGRPercent !== undefined);
  const totalInvested = rows.reduce((sum, r) => sum + Number(r.InvestedAmount), 0);
  if (!totalInvested) return null;
  const weighted = rows.reduce((sum, r) => sum + r.CAGRPercent * Number(r.InvestedAmount), 0);
  return weighted / totalInvested;
}

function renderROICompare(roi, clientId) {
  const section = $('roiCompare');
  const bench = state.benchmarks;

  const holdings = clientId
    ? state.otherInvestments.filter(h => String(h.ClientID) === String(clientId))
    : [];
  const typeRows = Object.keys(OTHER_INVESTMENT_COLORS)
    .map(type => ({ label: type, value: computeTypeROI(holdings, type), color: OTHER_INVESTMENT_COLORS[type] }))
    .filter(r => r.value !== null);

  if ((roi === null || roi === undefined) && !bench && typeRows.length === 0) {
    section.hidden = true;
    return;
  }
  section.hidden = false;

  const rows = [];
  if (roi !== null && roi !== undefined) rows.push({ label: 'Your effective return (CAGR)', value: roi, cls: 'roi-bar-yours' });
  if (bench) {
    rows.push({ label: 'Nifty 50 — 1 Yr', value: bench.oneYear, cls: 'roi-bar-nifty' });
    rows.push({ label: 'Nifty 50 — 3 Yr', value: bench.threeYear, cls: 'roi-bar-nifty' });
    rows.push({ label: 'Nifty 50 — 5 Yr', value: bench.fiveYear, cls: 'roi-bar-nifty' });
  }
  typeRows.forEach(r => rows.push({ label: r.label, value: r.value, color: r.color }));

  const validRows = rows.filter(r => r.value !== null && r.value !== undefined && !isNaN(r.value));
  const max = Math.max(...validRows.map(r => Math.abs(r.value)), 1);
  $('roiBars').innerHTML = validRows.map(r => {
    const isNeg = r.value < 0;
    const widthPct = Math.max(4, Math.abs(r.value) / max * 100);
    const barClass = isNeg ? 'roi-bar-negative' : (r.cls || '');
    const barStyle = `width:${widthPct}%` + (r.color && !isNeg ? `;background:${r.color}` : '');
    return `
    <div class="roi-row">
      <span class="roi-label">${r.label}</span>
      <div class="roi-track"><div class="roi-fill ${barClass}" style="${barStyle}"></div></div>
      <span class="roi-value${isNeg ? ' negative' : ''}">${r.value.toFixed(1)}%</span>
    </div>
  `;
  }).join('');

  $('roiNote').textContent = (bench && bench.note) || 'Figures shown are annualized (CAGR).';
}

// ---------- SIP calculator ----------

$('sipToggle').addEventListener('click', () => {
  const body = $('sipBody');
  body.hidden = !body.hidden;
  $('sipToggleIcon').innerHTML = body.hidden ? '&#9662;' : '&#9652;';
});

$('sipCalcBtn').addEventListener('click', () => {
  const amount = Number($('sipAmount').value);
  const rate = Number($('sipRate').value);
  const years = Number($('sipYears').value);
  const errEl = $('sipError');

  if (!amount || amount <= 0 || !years || years <= 0 || rate < 0 || $('sipRate').value === '') {
    errEl.textContent = 'Enter a monthly amount, annual return, and duration first.';
    errEl.hidden = false;
    $('sipResult').hidden = true;
    return;
  }
  errEl.hidden = true;

  const i = rate / 100 / 12;
  const n = years * 12;
  const fv = i === 0 ? amount * n : amount * ((Math.pow(1 + i, n) - 1) / i) * (1 + i);
  const invested = amount * n;

  $('sipInvested').textContent = fmt(invested);
  $('sipReturns').textContent = fmt(fv - invested);
  $('sipFV').textContent = fmt(fv);
  $('sipResult').hidden = false;
});

// ---------- Other Investments ----------

function renderOtherInvestments(clientId, editable) {
  const holdings = state.otherInvestments.filter(h => String(h.ClientID) === String(clientId));
  $('addHoldingForm').hidden = !editable;

  const body = $('holdingsBody');
  body.innerHTML = holdings.map(h => {
    const gain = h.GainPercent;
    const gainCls = gain === null || gain === undefined ? '' : (gain >= 0 ? 'gain-positive' : 'gain-negative');
    const gainText = gain === null || gain === undefined ? '-' : `${gain.toFixed(1)}%`;
    const cagrText = h.CAGRPercent === null || h.CAGRPercent === undefined ? '-' : `${h.CAGRPercent.toFixed(1)}%`;
    const actions = editable
      ? `<div class="holding-actions">
          <button type="button" class="btn-ghost" data-action="edit" data-id="${h.HoldingID}">Edit value</button>
          <button type="button" class="btn-ghost" data-action="delete" data-id="${h.HoldingID}">Delete</button>
        </div>`
      : '';
    return `
    <tr>
      <td>${h.Type}</td>
      <td>${h.Name}</td>
      <td>${fmt(h.InvestedAmount)}</td>
      <td>${fmt(h.CurrentValue)}</td>
      <td class="${gainCls}">${gainText}</td>
      <td>${cagrText}</td>
      <td>${actions}</td>
    </tr>
  `;
  }).join('');

  $('noHoldings').hidden = holdings.length > 0;
  $('holdingsTable').hidden = holdings.length === 0;
}

$('addHoldingForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const payload = {
    action: 'addOtherInvestment',
    clientId: state.currentClientId,
    accessCode: state.clientAccessCode,
    type: $('holdingType').value,
    name: $('holdingName').value.trim(),
    investedAmount: $('holdingInvested').value,
    currentValue: $('holdingCurrent').value,
    date: $('holdingDate').value,
    notes: $('holdingNotes').value.trim(),
  };

  const msg = $('holdingMsg');
  if (!payload.name || !payload.investedAmount || !payload.currentValue || !payload.date) {
    msg.hidden = false;
    msg.textContent = 'Please fill in name, invested amount, current value and date.';
    return;
  }

  msg.hidden = false;
  msg.textContent = 'Saving...';
  try {
    const res = await api(payload);
    if (res.error) {
      msg.textContent = res.error;
      return;
    }
    await refreshClientOtherInvestments();
    $('addHoldingForm').reset();
    msg.textContent = 'Investment added.';
    setTimeout(() => { msg.hidden = true; }, 2500);
  } catch (err) {
    msg.textContent = 'Could not save this investment.';
  }
});

$('holdingsBody').addEventListener('click', async (e) => {
  const btn = e.target.closest('button[data-action]');
  if (!btn) return;
  const holdingId = btn.dataset.id;

  if (btn.dataset.action === 'edit') {
    const holding = state.otherInvestments.find(h => String(h.HoldingID) === String(holdingId));
    const input = prompt(`New current value for "${holding ? holding.Name : holdingId}":`, holding ? holding.CurrentValue : '');
    if (input === null || input.trim() === '' || isNaN(Number(input))) return;
    btn.disabled = true;
    try {
      const res = await api({
        action: 'updateOtherInvestmentValue',
        clientId: state.currentClientId,
        accessCode: state.clientAccessCode,
        holdingId,
        currentValue: input,
      });
      if (res.error) { alert(res.error); return; }
      await refreshClientOtherInvestments();
    } catch (err) {
      alert('Could not update this investment.');
    } finally {
      btn.disabled = false;
    }
  } else if (btn.dataset.action === 'delete') {
    if (!confirm('Remove this investment from your records?')) return;
    btn.disabled = true;
    try {
      const res = await api({
        action: 'deleteOtherInvestment',
        clientId: state.currentClientId,
        accessCode: state.clientAccessCode,
        holdingId,
      });
      if (res.error) { alert(res.error); return; }
      await refreshClientOtherInvestments();
    } catch (err) {
      alert('Could not delete this investment.');
    } finally {
      btn.disabled = false;
    }
  }
});

async function refreshClientOtherInvestments() {
  const refreshed = await api({
    action: 'login',
    clientId: state.currentClientId,
    accessCode: state.clientAccessCode,
  });
  if (refreshed.error) return;
  state.otherInvestments = refreshed.otherInvestments || [];
  renderOtherInvestments(state.currentClientId, true);
  renderROICompare(state.portfolioROI, state.currentClientId);
}

function fmt(n) {
  return (Number(n) || 0).toLocaleString('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 });
}

// ---------- Logout ----------

$('logoutBtn').addEventListener('click', () => {
  Object.assign(state, {
    role: null, adminKey: null, clientAccessCode: null, clients: [], entries: [],
    dueCycles: [], summary: null, clientROI: {}, portfolioROI: null, benchmarks: null,
    otherInvestments: [], currentClientId: null, currentClientName: '',
  });
  clearSession();
  $('dashboardView').hidden = true;
  $('logoutBtn').hidden = true;
  $('loginView').hidden = false;
  $('clientLoginForm').reset();
  $('adminLoginForm').reset();
});

// ---------- Helpers ----------

function setLoading(v) { $('loading').hidden = !v; }
function showError(msg) { const el = $('loginError'); el.textContent = msg; el.hidden = false; }
function hideError() { $('loginError').hidden = true; }

restoreSession();
