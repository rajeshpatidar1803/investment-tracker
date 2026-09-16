// Paste your deployed Apps Script Web App URL here (see apps-script/Code.gs).
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwZXEJsYmuM6DF5j0_4VDkZo5xqNRkm1YiQaxbI1CGt0DU5m2AmjsXpIEJcFZ5fpZoU/exec';

const state = {
  role: null,       // 'admin' | 'client'
  adminKey: null,
  clients: [],
  entries: [],       // entries currently displayed
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

function enterAdminDashboard() {
  $('loginView').hidden = true;
  $('dashboardView').hidden = false;
  $('adminControls').hidden = false;
  $('logoutBtn').hidden = false;

  const select = $('clientSelect');
  select.innerHTML = state.clients
    .map(c => `<option value="${c.ClientID}">${c.ClientName} (${c.ClientID})</option>`)
    .join('');
  select.addEventListener('change', renderForSelectedClient);

  $('entryDate').valueAsDate = new Date();

  if (state.clients.length) {
    renderForSelectedClient();
  }
}

function renderForSelectedClient() {
  const clientId = $('clientSelect').value;
  const client = state.clients.find(c => String(c.ClientID) === String(clientId));
  const entries = state.entries.filter(en => String(en.ClientID) === String(clientId));
  $('clientNameHeading').textContent = client ? client.ClientName : '';
  renderSummaryAndTable(entries);
}

function enterClientDashboard() {
  $('loginView').hidden = true;
  $('dashboardView').hidden = false;
  $('adminControls').hidden = true;
  $('logoutBtn').hidden = false;
  $('clientNameHeading').textContent = state.currentClientName;
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
    state.entries.push({
      ClientID: clientId,
      Date: payload.date,
      Type: payload.type,
      Amount: Number(payload.amount),
      Notes: payload.notes,
    });
    renderForSelectedClient();
    $('addEntryForm').reset();
    $('entryDate').valueAsDate = new Date();
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

  $('sumInvested').textContent = fmt(invested);
  $('sumWithdrawn').textContent = fmt(withdrawn);
  $('sumNet').textContent = fmt(net);
  $('sumCurrentValue').textContent = fmt(currentValue);

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
