// Paste your deployed Apps Script Web App URL here (see apps-script/Code.gs).
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwZXEJsYmuM6DF5j0_4VDkZo5xqNRkm1YiQaxbI1CGt0DU5m2AmjsXpIEJcFZ5fpZoU/exec';

const state = {
  role: null,       // 'admin' | 'client'
  adminKey: null,
  clients: [],
  entries: [],       // entries currently displayed
  dueCycles: [],      // admin only: tranches with an interest cycle awaiting a decision
  summary: null,      // admin only: portfolio-wide totals
  currentClientId: null,
  currentClientName: '',
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

async function doLogin(payload) {
  setLoading(true);
  hideError();
  try {
    const res = await api(payload);
    if (res.error) {
      showError(res.error);
      return;
    }
    state.role = res.role;
    if (res.role === 'admin') {
      state.adminKey = payload.adminKey;
      state.clients = res.clients;
      state.entries = res.entries;
      state.dueCycles = res.dueCycles || [];
      state.summary = res.summary || null;
      enterAdminDashboard();
    } else {
      state.currentClientId = res.client.ClientID;
      state.currentClientName = res.client.ClientName;
      state.entries = res.entries;
      enterClientDashboard();
    }
  } catch (err) {
    showError('Could not reach the server. Check WEB_APP_URL in js/app.js.');
  } finally {
    setLoading(false);
  }
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

  if (state.clients.length) {
    renderForSelectedClient();
  }
}

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
  renderSummaryAndTable(entries);
}

function enterClientDashboard() {
  $('loginView').hidden = true;
  $('dashboardView').hidden = false;
  $('adminControls').hidden = true;
  $('logoutBtn').hidden = false;
  $('clientNameHeading').textContent = `${state.currentClientName} (${state.currentClientId})`;
  renderSummaryAndTable(state.entries);
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

function renderSummaryAndTable(entries) {
  const sorted = [...entries].sort((a, b) => new Date(a.Date) - new Date(b.Date));

  let invested = 0, withdrawn = 0, latestValue = null;
  for (const en of sorted) {
    const amt = Number(en.Amount) || 0;
    if (en.Type === 'Investment') invested += amt;
    else if (en.Type === 'Withdrawal') withdrawn += amt;
    else if (en.Type === 'Current Value') latestValue = amt; // last one wins, list is sorted by date
  }
  const net = invested - withdrawn;
  const currentValue = latestValue !== null ? latestValue : net;
  const gain = currentValue - net;

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
      <td>${en.Type}</td>
      <td>${fmt(Number(en.Amount))}</td>
      <td>${en.TrancheID || ''}</td>
      <td>${en.Notes || ''}</td>
    </tr>
  `).join('');

  $('noEntries').hidden = descending.length > 0;
  $('entriesTable').hidden = descending.length === 0;
}

function fmt(n) {
  return (Number(n) || 0).toLocaleString('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 });
}

// ---------- Logout ----------

$('logoutBtn').addEventListener('click', () => {
  Object.assign(state, {
    role: null, adminKey: null, clients: [], entries: [],
    dueCycles: [], summary: null,
    currentClientId: null, currentClientName: '',
  });
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
