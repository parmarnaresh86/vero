import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import {
  BellRing,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  Filter,
  KeyRound,
  MessageSquareText,
  Printer,
  ReceiptText,
  RefreshCw,
  Search,
  Send,
  Settings,
  ShieldCheck,
  UploadCloud,
  Users
} from 'lucide-react';
import './styles.css';

const API = import.meta.env.VITE_API_URL || 'http://localhost:4200/api';

const importTypes = [
  { kind: 'demand', label: 'Demand Register', file: 'KrManganaragisterNew.xlsx' },
  { kind: 'recovery', label: 'Vasulat Register', file: 'VasulatRegisterNew.xlsx' },
  { kind: 'pending', label: 'Baki Register', file: 'BakiRegisterNew.xlsx' },
  { kind: 'mobile', label: 'Mobile Mapping', file: 'MilkatnoswisemobilenoReport.xlsx' }
];

const excelReportPages = {
  'mobile-detail': { kind: 'mobile', label: 'Mobile Detail', description: 'Property wise mobile details from MilkatnoswisemobilenoReport.xlsx' },
  'baki-register': { kind: 'pending', label: 'Baki Register', description: 'Pending tax register rows from BakiRegisterNew.xlsx' },
  'demand-register': { kind: 'demand', label: 'Demand Register', description: 'Demand register rows from KrManganaragisterNew.xlsx' },
  'vasulat-register': { kind: 'recovery', label: 'Vasulat Register', description: 'Receipt and recovery rows from VasulatRegisterNew.xlsx' }
};

const sourceColumns = {
  mobile: [
    { label: 'Property', key: 'property_no' },
    { label: 'Owner', key: 'holder_name' },
    { label: 'Occupant', key: 'occupant_name' },
    { label: 'Mobile', key: 'mobile' },
    { label: 'Category', key: 'category' },
    { label: 'Pending', key: 'grand_total', type: 'money' },
    { label: 'Description', key: 'description' }
  ],
  pending: [
    { label: 'Property', key: 'property_no' },
    { label: 'House', key: 'house_no' },
    { label: 'Owner', key: 'holder_name' },
    { label: 'Area', key: 'area' },
    { label: 'Previous', key: 'previous_total', type: 'money' },
    { label: 'Current', key: 'current_total', type: 'money' },
    { label: 'Total', key: 'grand_total', type: 'money' }
  ],
  demand: [
    { label: 'Property', key: 'property_no' },
    { label: 'Owner', key: 'holder_name' },
    { label: 'Occupant', key: 'occupant_name' },
    { label: 'Previous', key: 'previous_total', type: 'money' },
    { label: 'Current', key: 'current_total', type: 'money' },
    { label: 'Demand Total', key: 'grand_total', type: 'money' }
  ],
  recovery: [
    { label: 'Receipt', key: 'receipt_no' },
    { label: 'Date', key: 'receipt_date' },
    { label: 'Property', key: 'property_no' },
    { label: 'Owner', key: 'holder_name' },
    { label: 'Area', key: 'area' },
    { label: 'Amount', key: 'grand_total', type: 'money' }
  ]
};

const pageList = [
  { id: 'dashboard', label: 'Dashboard', icon: Users, permission: 'dashboard.view' },
  { id: 'import', label: 'Excel Import', icon: UploadCloud, permission: 'excel.import' },
  { id: 'mobile-detail', label: 'Mobile Detail', icon: FileSpreadsheet, permission: 'excel.view' },
  { id: 'baki-register', label: 'Baki Register', icon: FileSpreadsheet, permission: 'excel.view' },
  { id: 'demand-register', label: 'Demand Register', icon: FileSpreadsheet, permission: 'excel.view' },
  { id: 'vasulat-register', label: 'Vasulat Register', icon: FileSpreadsheet, permission: 'excel.view' },
  { id: 'billing', label: 'Billing', icon: Printer, permission: 'billing.view' },
  { id: 'broadcast', label: 'Broadcast', icon: MessageSquareText, permission: 'broadcast.send' },
  { id: 'whatsapp-status', label: 'WhatsApp Status', icon: MessageSquareText, permission: 'reports.view' },
  { id: 'reports', label: 'Reports', icon: BellRing, permission: 'reports.view' },
  { id: 'admin', label: 'Admin', icon: ShieldCheck, permission: 'admin.manage' }
];

function currency(value) {
  return new Intl.NumberFormat('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 }).format(value || 0);
}

function App() {
  const [user, setUser] = useState(() => JSON.parse(localStorage.getItem('billingUser') || 'null'));
  const [page, setPage] = useState('dashboard');
  const [dashboard, setDashboard] = useState({});
  const [taxpayers, setTaxpayers] = useState([]);
  const [reports, setReports] = useState([]);
  const [whatsappStatus, setWhatsappStatus] = useState({ summary: {}, rows: [] });
  const [importHistory, setImportHistory] = useState([]);
  const [sourceRows, setSourceRows] = useState([]);
  const [filters, setFilters] = useState({ q: '', area: '', category: '', status: '' });
  const emptyAdvancedFilters = { q: '', propertyNo: '', houseNo: '', ownerName: '', occupantName: '', area: '', category: '', mobile: '', status: '', messageStatus: '', minAmount: '', maxAmount: '' };
  const [broadcastFilters, setBroadcastFilters] = useState(emptyAdvancedFilters);
  const [statusFilters, setStatusFilters] = useState(emptyAdvancedFilters);
  const [message, setMessage] = useState('તમારું ગ્રામપંચાયત બિલ તૈયાર છે. કૃપા કરીને સમયસર ચુકવણી કરો.');
  const [notice, setNotice] = useState('Ready');
  const [admin, setAdmin] = useState({ permissions: [], roles: [], users: [], settings: {} });
  const [villages, setVillages] = useState([]);
  const [activeVillageId, setActiveVillageId] = useState('');
  const [newVillageName, setNewVillageName] = useState('');
  const [couponCode, setCouponCode] = useState('');

  const allowedPages = useMemo(() => pageList.filter((item) => has(user, item.permission)), [user]);
  const areas = useMemo(() => [...new Set(taxpayers.map((item) => item.area).filter(Boolean))], [taxpayers]);
  const categories = useMemo(() => [...new Set(taxpayers.map((item) => item.category).filter(Boolean))], [taxpayers]);

  function headers() {
    return { 'Content-Type': 'application/json', 'x-user-id': String(user?.id || '') };
  }

  async function api(path, options = {}) {
    const response = await fetch(`${API}${path}`, {
      ...options,
      headers: options.body instanceof FormData ? { 'x-user-id': String(user?.id || '') } : { ...headers(), ...(options.headers || {}) }
    });
    const data = await response.json().catch(() => null);
    if (!response.ok) throw new Error(data?.error || 'Request failed');
    return data;
  }

  async function load() {
    if (!user) return;
    const query = new URLSearchParams(filters).toString();
    const villageParam = activeVillageId ? `villageId=${activeVillageId}` : '';
    const tasks = [];
    if (has(user, 'excel.view')) tasks.push(api('/villages').then((rows) => {
      setVillages(rows);
      if (!activeVillageId && rows[0]) setActiveVillageId(String(rows[0].id));
    }));
    if (has(user, 'dashboard.view')) tasks.push(api(`/dashboard?${villageParam}`).then(setDashboard));
    if (has(user, 'billing.view')) tasks.push(api(`/taxpayers?${[query, villageParam].filter(Boolean).join('&')}`).then(setTaxpayers));
    if (has(user, 'reports.view')) tasks.push(api('/messages/report').then(setReports));
    if (has(user, 'reports.view') && page === 'whatsapp-status') {
      tasks.push(api(`/whatsapp/status?${new URLSearchParams(statusFilters).toString()}`).then(setWhatsappStatus));
    }
    if (has(user, 'excel.view')) {
      tasks.push(api(`/imports?${villageParam}`).then(setImportHistory));
      if (excelReportPages[page]) {
        tasks.push(api(`/source-rows?kind=${excelReportPages[page].kind}${villageParam ? `&${villageParam}` : ''}`).then(setSourceRows));
      }
    }
    if (has(user, 'admin.manage')) tasks.push(api('/admin/bootstrap').then(setAdmin));
    await Promise.all(tasks);
  }

  useEffect(() => {
    if (user && !has(user, pageList.find((item) => item.id === page)?.permission)) {
      setPage(allowedPages[0]?.id || 'dashboard');
    }
  }, [user, allowedPages, page]);

  useEffect(() => {
    load().catch((error) => setNotice(error.message));
  }, [user, filters, page, statusFilters, activeVillageId]);

  async function login(username, password) {
    const loggedIn = await fetch(`${API}/auth/login`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password })
    }).then(async (response) => {
      const data = await response.json();
      if (!response.ok) throw new Error(data.error);
      return data;
    });
    localStorage.setItem('billingUser', JSON.stringify(loggedIn));
    setUser(loggedIn);
    setPage(pageList.find((item) => loggedIn.permissions.includes(item.permission))?.id || 'dashboard');
  }

  async function importFile(kind, file) {
    if (!file) return;
    const form = new FormData();
    form.append('file', file);
    form.append('villageId', activeVillageId || villages[0]?.id || 1);
    form.append('villageName', villages.find((item) => String(item.id) === String(activeVillageId))?.name || '');
    setNotice(`Importing ${file.name}`);
    const result = await api(`/import/${kind}`, { method: 'POST', body: form });
    setNotice(`Imported ${result.imported || 0} records. Total ${result.total || 0}.`);
    await load();
  }

  async function togglePaid(item) {
    await api(`/taxpayers/${item.propertyNo}/status`, {
      method: 'PATCH',
      body: JSON.stringify({ paid: !item.paid })
    });
    await load();
  }

  async function sendMessages() {
    const result = await api('/messages/send', {
      method: 'POST',
      body: JSON.stringify({ message, filters: { ...broadcastFilters, villageId: activeVillageId } })
    });
    setNotice(`Queued ${result.sent || 0} WhatsApp messages.`);
    await load();
  }

  function logout() {
    localStorage.removeItem('billingUser');
    setUser(null);
  }

  async function createVillage() {
    if (!newVillageName.trim()) return;
    const village = await api('/villages', { method: 'POST', body: JSON.stringify({ name: newVillageName }) });
    setNewVillageName('');
    setActiveVillageId(String(village.id));
    await load();
  }

  async function redeemCurrentCoupon() {
    const result = await api('/coupons/redeem', { method: 'POST', body: JSON.stringify({ code: couponCode }) });
    localStorage.setItem('billingUser', JSON.stringify(result.user));
    setUser(result.user);
    setCouponCode('');
    setNotice(`Recharge added: ${result.coupon.message_credit} messages`);
  }

  if (!user) return <LoginPage onLogin={login} notice={notice} setNotice={setNotice} />;

  return (
    <main className="shell">
      <aside className="sidebar">
        <div className="brand">
          <ReceiptText size={30} />
          <div>
            <strong>WhatsApp Billing</strong>
            <span>{user.name} · {user.roleName}</span>
          </div>
        </div>
        <nav>
          {allowedPages.map((item) => {
            const Icon = item.icon;
            return <button className={page === item.id ? 'navButton active' : 'navButton'} key={item.id} onClick={() => setPage(item.id)}><Icon size={18} /> {item.label}</button>;
          })}
        </nav>
        <div className="roleBox">
          <label>Active Rights</label>
          <small>{user.permissions.length} permissions · package {user.packageLimit || 0}+{user.bonusLimit || 0}</small>
          {has(user, 'excel.view') && (
            <>
              <select value={activeVillageId} onChange={(event) => setActiveVillageId(event.target.value)}>
                <option value="">All villages</option>
                {villages.map((village) => <option key={village.id} value={village.id}>{village.name}</option>)}
              </select>
              <input placeholder="New village" value={newVillageName} onChange={(event) => setNewVillageName(event.target.value)} />
              <button className="secondary" onClick={createVillage}>Add Village</button>
            </>
          )}
          {has(user, 'broadcast.send') && (
            <>
              <input placeholder="Recharge coupon" value={couponCode} onChange={(event) => setCouponCode(event.target.value)} />
              <button className="secondary" onClick={redeemCurrentCoupon}>Redeem Coupon</button>
            </>
          )}
          <button className="secondary" onClick={logout}>Logout</button>
        </div>
      </aside>

      <section className="content">
        <header className="topbar">
          <div>
            <h1>{pageTitle(page)}</h1>
            <p>Pages load as per user role and assigned rights.</p>
          </div>
          <button className="iconButton" onClick={load} title="Refresh"><RefreshCw size={18} /></button>
        </header>

        {page === 'dashboard' && <DashboardPage dashboard={dashboard} />}
        {page === 'import' && <ImportPage importFile={importFile} importHistory={importHistory} villages={villages} activeVillageId={activeVillageId} />}
        {excelReportPages[page] && <ExcelReportPage pageInfo={excelReportPages[page]} sourceRows={sourceRows} />}
        {page === 'billing' && <BillingPage taxpayers={taxpayers} filters={filters} setFilters={setFilters} areas={areas} categories={categories} togglePaid={togglePaid} canUpdate={has(user, 'billing.update')} />}
        {page === 'broadcast' && <BroadcastPage message={message} setMessage={setMessage} sendMessages={sendMessages} notice={notice} filters={broadcastFilters} setFilters={setBroadcastFilters} emptyFilters={emptyAdvancedFilters} areas={areas} categories={categories} />}
        {page === 'whatsapp-status' && <WhatsappStatusPage data={whatsappStatus} filters={statusFilters} setFilters={setStatusFilters} emptyFilters={emptyAdvancedFilters} areas={areas} categories={categories} />}
        {page === 'reports' && <ReportsPage taxpayers={taxpayers} reports={reports} filters={filters} user={user} />}
        {page === 'admin' && <AdminPage admin={admin} api={api} reload={load} setNotice={setNotice} />}
      </section>
    </main>
  );
}

function has(user, permission) {
  return Boolean(user?.permissions?.includes(permission));
}

function pageTitle(page) {
  return {
    dashboard: 'Dashboard',
    import: 'Excel Import',
    'mobile-detail': 'Mobile Detail',
    'baki-register': 'Baki Register',
    'demand-register': 'Demand Register',
    'vasulat-register': 'Vasulat Register',
    billing: 'Billing',
    broadcast: 'WhatsApp Broadcast',
    'whatsapp-status': 'WhatsApp Status',
    reports: 'Reports',
    admin: 'Admin Panel'
  }[page] || 'Dashboard';
}

function LoginPage({ onLogin, notice, setNotice }) {
  const [username, setUsername] = useState('admin');
  const [password, setPassword] = useState('admin123');
  return (
    <main className="loginShell">
      <section className="loginPanel">
        <ReceiptText size={34} />
        <h1>WhatsApp Billing Login</h1>
        <input value={username} onChange={(event) => setUsername(event.target.value)} placeholder="Username" />
        <input value={password} onChange={(event) => setPassword(event.target.value)} placeholder="Password" type="password" />
        <button className="primary" onClick={() => onLogin(username, password).catch((error) => setNotice(error.message))}><KeyRound size={18} /> Login</button>
        <p className="notice">{notice}</p>
      </section>
    </main>
  );
}

function DashboardPage({ dashboard }) {
  return (
    <section className="metrics">
      <Metric icon={<Users />} label="Records" value={dashboard.total || 0} />
      <Metric icon={<CheckCircle2 />} label="Paid" value={dashboard.paid || 0} />
      <Metric icon={<BellRing />} label="Unpaid" value={dashboard.unpaid || 0} />
      <Metric icon={<ReceiptText />} label="Pending Amount" value={currency(dashboard.pendingAmount)} />
      <Metric icon={<MessageSquareText />} label="Mobile Ready" value={dashboard.mobileReady || 0} />
    </section>
  );
}

function ImportPage({ importFile, importHistory, villages, activeVillageId }) {
  const activeVillage = villages.find((item) => String(item.id) === String(activeVillageId));
  return (
    <>
      <section className="band">
        <div className="sectionTitle"><FileSpreadsheet size={22} /><h2>Excel Import</h2></div>
        <p className="notice">Uploading into village: {activeVillage?.name || 'All villages / default'}</p>
        <div className="importGrid">
          {importTypes.map((item) => (
            <label className="uploadTile" key={item.kind}>
              <UploadCloud size={24} />
              <strong>{item.label}</strong>
              <span>{item.file}</span>
              <input type="file" accept=".xlsx,.xls" onChange={(event) => importFile(item.kind, event.target.files?.[0])} />
            </label>
          ))}
        </div>
      </section>
      <ImportHistory importHistory={importHistory} />
    </>
  );
}

function ImportHistory({ importHistory }) {
  return (
    <section className="band">
      <div className="sectionTitle"><UploadCloud size={22} /><h2>Excel Import History</h2></div>
      <div className="importHistory">
        {importHistory.map((item) => (
          <div className="historyItem" key={item.id}>
            <strong>{item.kind}</strong>
            <span>{item.file_name}</span>
            <small>{item.village_name || '-'} · {item.user_name || '-'} · {item.row_count} rows</small>
          </div>
        ))}
      </div>
    </section>
  );
}

function ExcelReportPage({ pageInfo, sourceRows }) {
  return (
    <section className="band">
      <div className="sectionTitle"><FileSpreadsheet size={22} /><h2>{pageInfo.label}</h2></div>
      <div className="sourceSummary">
        <strong>{pageInfo.description}</strong>
        <span>{sourceRows.length} rows loaded from SQLite</span>
      </div>
      <DataTable columns={sourceColumns[pageInfo.kind] || []} rows={sourceRows} />
    </section>
  );
}

function BillingPage({ taxpayers, filters, setFilters, areas, categories, togglePaid, canUpdate }) {
  const [columnFilters, setColumnFilters] = useState({});
  const [pageNo, setPageNo] = useState(1);
  const columns = [
    { label: 'Property', key: 'propertyNo' },
    { label: 'Name', key: 'holderName' },
    { label: 'Area', key: 'area' },
    { label: 'Mobile', key: 'mobile' },
    { label: 'Amount', key: 'pendingTax', type: 'money' },
    { label: 'Status', key: 'paidText' }
  ];
  const enrichedRows = taxpayers.map((item) => ({ ...item, paidText: item.paid ? 'Paid' : 'Unpaid' }));
  const filteredRows = useMemo(() => filterRowsByColumns(enrichedRows, columns, columnFilters), [taxpayers, columnFilters]);
  const pagedRows = useMemo(() => paginateRows(filteredRows, pageNo), [filteredRows, pageNo]);

  return (
    <section className="band">
      <div className="sectionTitle"><Filter size={22} /><h2>Client Filter & Bills</h2></div>
      <FilterBar filters={filters} setFilters={setFilters} areas={areas} categories={categories} />
      <ColumnFilterBar columns={columns} filters={columnFilters} setFilters={(next) => { setColumnFilters(next); setPageNo(1); }} />
      <div className="tableWrap">
        <table>
          <thead><tr><th>Property</th><th>Name</th><th>Area</th><th>Mobile</th><th>Amount</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody>
            {pagedRows.map((item) => (
              <tr key={item.propertyNo}>
                <td>{item.propertyNo}</td>
                <td><strong>{item.holderName || '-'}</strong><span>{item.description}</span></td>
                <td>{item.area || '-'}</td>
                <td>{item.mobile || '-'}</td>
                <td>{currency(item.pendingTax || item.currentTax)}</td>
                <td><button disabled={!canUpdate} className={item.paid ? 'status paid' : 'status unpaid'} onClick={() => togglePaid(item)}>{item.paid ? 'Paid' : 'Unpaid'}</button></td>
                <td className="actions"><a className="iconButton" href={`${API}/bills/${item.propertyNo}.pdf`} target="_blank"><ReceiptText size={17} /></a><button className="iconButton" onClick={() => window.open(`${API}/bills/${item.propertyNo}.pdf`, '_blank')}><Printer size={17} /></button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <TablePagination pageNo={pageNo} setPageNo={setPageNo} total={filteredRows.length} />
    </section>
  );
}

function FilterBar({ filters, setFilters, areas, categories }) {
  return (
    <div className="filters">
      <label className="searchBox"><Search size={17} /><input placeholder="Property, name, mobile" value={filters.q} onChange={(event) => setFilters({ ...filters, q: event.target.value })} /></label>
      <select value={filters.area} onChange={(event) => setFilters({ ...filters, area: event.target.value })}><option value="">All areas</option>{areas.map((area) => <option key={area}>{area}</option>)}</select>
      <select value={filters.category} onChange={(event) => setFilters({ ...filters, category: event.target.value })}><option value="">All categories</option>{categories.map((category) => <option key={category}>{category}</option>)}</select>
      <select value={filters.status} onChange={(event) => setFilters({ ...filters, status: event.target.value })}><option value="">All status</option><option value="paid">Paid</option><option value="unpaid">Unpaid</option></select>
    </div>
  );
}

function BroadcastPage({ message, setMessage, sendMessages, notice, filters, setFilters, emptyFilters, areas, categories }) {
  return (
    <>
      <section className="band">
        <div className="sectionTitle"><Filter size={22} /><h2>Broadcast Filters</h2></div>
        <AdvancedFilterPanel filters={filters} setFilters={setFilters} emptyFilters={emptyFilters} areas={areas} categories={categories} />
      </section>
      <section className="split">
        <div className="band">
        <div className="sectionTitle"><MessageSquareText size={22} /><h2>WhatsApp Broadcast</h2></div>
        <textarea value={message} onChange={(event) => setMessage(event.target.value)} rows={5} />
        <button className="primary" onClick={sendMessages}><Send size={18} /> Send to Filtered Clients</button>
        <p className="notice">{notice}</p>
        </div>
        <div className="preview"><img src="/register-format.png" alt="Register format preview" /></div>
      </section>
    </>
  );
}

function AdvancedFilterPanel({ filters, setFilters, emptyFilters, areas, categories }) {
  return (
    <div className="advancedFilters">
      <label><span>Search</span><input placeholder="Any text" value={filters.q} onChange={(event) => setFilters({ ...filters, q: event.target.value })} /></label>
      <label><span>Milkat No</span><input placeholder="Property no" value={filters.propertyNo} onChange={(event) => setFilters({ ...filters, propertyNo: event.target.value })} /></label>
      <label><span>House No</span><input placeholder="House no" value={filters.houseNo} onChange={(event) => setFilters({ ...filters, houseNo: event.target.value })} /></label>
      <label><span>Owner</span><input placeholder="Owner name" value={filters.ownerName} onChange={(event) => setFilters({ ...filters, ownerName: event.target.value })} /></label>
      <label><span>Occupant</span><input placeholder="Occupant name" value={filters.occupantName} onChange={(event) => setFilters({ ...filters, occupantName: event.target.value })} /></label>
      <label><span>Mobile</span><input placeholder="Mobile no" value={filters.mobile} onChange={(event) => setFilters({ ...filters, mobile: event.target.value })} /></label>
      <label><span>Area</span><select value={filters.area} onChange={(event) => setFilters({ ...filters, area: event.target.value })}><option value="">All areas</option>{areas.map((area) => <option key={area}>{area}</option>)}</select></label>
      <label><span>Category</span><select value={filters.category} onChange={(event) => setFilters({ ...filters, category: event.target.value })}><option value="">All categories</option>{categories.map((category) => <option key={category}>{category}</option>)}</select></label>
      <label><span>Payment</span><select value={filters.status} onChange={(event) => setFilters({ ...filters, status: event.target.value })}><option value="">All</option><option value="paid">Paid</option><option value="unpaid">Unpaid</option></select></label>
      <label><span>WhatsApp</span><select value={filters.messageStatus} onChange={(event) => setFilters({ ...filters, messageStatus: event.target.value })}><option value="">All</option><option value="not_sent">Not Sent</option><option value="sent">Sent Any</option><option value="pending">Pending</option><option value="delivered">Delivered</option><option value="failed">Failed</option></select></label>
      <label><span>Min Amount</span><input type="number" placeholder="From" value={filters.minAmount} onChange={(event) => setFilters({ ...filters, minAmount: event.target.value })} /></label>
      <label><span>Max Amount</span><input type="number" placeholder="To" value={filters.maxAmount} onChange={(event) => setFilters({ ...filters, maxAmount: event.target.value })} /></label>
      <button className="secondary" onClick={() => setFilters(emptyFilters)}>Clear Filters</button>
    </div>
  );
}

function WhatsappStatusPage({ data, filters, setFilters, emptyFilters, areas, categories }) {
  const rows = data.rows || [];
  const summary = data.summary || {};
  return (
    <>
      <section className="metrics">
        <Metric icon={<Users />} label="Total" value={summary.total || 0} />
        <Metric icon={<MessageSquareText />} label="Sent" value={summary.sent || 0} />
        <Metric icon={<BellRing />} label="Not Sent" value={summary.notSent || 0} />
        <Metric icon={<CheckCircle2 />} label="Delivered" value={summary.delivered || 0} />
        <Metric icon={<BellRing />} label="Failed" value={summary.failed || 0} />
      </section>
      <section className="band">
        <div className="sectionTitle"><Filter size={22} /><h2>Status Filters</h2></div>
        <AdvancedFilterPanel filters={filters} setFilters={setFilters} emptyFilters={emptyFilters} areas={areas} categories={categories} />
      </section>
      <section className="band">
        <div className="sectionTitle"><MessageSquareText size={22} /><h2>Milkat Wise WhatsApp Status</h2></div>
        <DataTable
          columns={[
            { label: 'Milkat', key: 'propertyNo' },
            { label: 'Owner', key: 'holderName' },
            { label: 'Area', key: 'area' },
            { label: 'Mobile', key: 'mobile' },
            { label: 'Amount', key: 'pendingTax', type: 'money' },
            { label: 'Payment', key: 'paidText' },
            { label: 'WhatsApp', key: 'messageStatus' },
            { label: 'Sent Count', key: 'totalSent' },
            { label: 'Last Sent', key: 'lastSentAt' },
            { label: 'Last Message', key: 'lastMessage' }
          ]}
          rows={rows.map((row) => ({ ...row, paidText: row.paid ? 'Paid' : 'Unpaid', lastSentAt: row.lastSentAt ? new Date(row.lastSentAt).toLocaleString() : '-' }))}
        />
      </section>
    </>
  );
}

function ReportsPage({ taxpayers, reports, filters, user }) {
  const query = new URLSearchParams(filters).toString();
  return (
    <>
      <section className="band">
        <div className="sectionTitle"><FileSpreadsheet size={22} /><h2>Excel Data Report</h2><a className="primary compact" href={`${API}/reports/taxpayers.xlsx?${query}&userId=${user.id}`} target="_blank" onClick={(event) => { if (!has(user, 'reports.view')) event.preventDefault(); }}><Download size={17} /> Download Excel</a></div>
        <DataTable columns={[{ label: 'Property', key: 'propertyNo' }, { label: 'Name', key: 'holderName' }, { label: 'Area', key: 'area' }, { label: 'Mobile', key: 'mobile' }, { label: 'Category', key: 'category' }, { label: 'Amount', key: 'pendingTax', type: 'money' }]} rows={taxpayers} />
      </section>
      <section className="band">
        <div className="sectionTitle"><BellRing size={22} /><h2>Message Sending Report</h2></div>
        <div className="reportList">
          {reports.length === 0 && <p>No messages sent yet.</p>}
          {reports.map((item, index) => <div className="reportItem" key={`${item.propertyNo}-${index}`}><strong>{item.mobile}</strong><span>{item.status}</span><small>{item.propertyNo} · {new Date(item.sentAt).toLocaleString()}</small></div>)}
        </div>
      </section>
    </>
  );
}

function AdminPage({ admin, api, reload, setNotice }) {
  const emptyRoleForm = { name: '', permissions: [], messageLimit: 100 };
  const emptyUserForm = { name: '', username: '', password: '', roleId: '', active: true, accessExpiresAt: '', packageType: 'monthly', packageLimit: 1000, bonusLimit: 0, packageStartedAt: '', packageExpiresAt: '' };
  const [roleForm, setRoleForm] = useState({ name: '', permissions: [], messageLimit: 100 });
  const [userForm, setUserForm] = useState({ name: '', username: '', password: '', roleId: '', active: true, accessExpiresAt: '', packageType: 'monthly', packageLimit: 1000, bonusLimit: 0, packageStartedAt: '', packageExpiresAt: '' });
  const [couponForm, setCouponForm] = useState({ code: '', messageCredit: 1000, assignedUserId: '', expiresAt: '' });
  const [settings, setSettings] = useState(admin.settings || {});

  useEffect(() => setSettings(admin.settings || {}), [admin.settings]);

  async function saveRoleForm() {
    await api('/admin/roles', { method: 'POST', body: JSON.stringify(roleForm) });
    setRoleForm(emptyRoleForm);
    setNotice(roleForm.id ? 'Role updated' : 'Role saved');
    await reload();
  }

  async function saveUserForm() {
    await api('/admin/users', { method: 'POST', body: JSON.stringify({ ...userForm, roleId: Number(userForm.roleId || admin.roles[0]?.id) }) });
    setUserForm(emptyUserForm);
    setNotice(userForm.id ? 'User updated' : 'User saved');
    await reload();
  }

  function editRole(role) {
    setRoleForm({
      id: role.id,
      name: role.name,
      permissions: role.permissions || [],
      messageLimit: role.messageLimit || 0
    });
  }

  async function deleteRoleItem(role) {
    if (!window.confirm(`Delete role ${role.name}?`)) return;
    await api(`/admin/roles/${role.id}`, { method: 'DELETE' });
    if (roleForm.id === role.id) setRoleForm(emptyRoleForm);
    setNotice('Role deleted');
    await reload();
  }

  function editUser(item) {
    setUserForm({
      id: item.id,
      name: item.name || '',
      username: item.username || '',
      password: '',
      roleId: item.roleId || '',
      active: Boolean(item.active),
      accessExpiresAt: item.accessExpiresAt || '',
      packageType: item.packageType || 'monthly',
      packageLimit: item.packageLimit || 0,
      bonusLimit: item.bonusLimit || 0,
      packageStartedAt: item.packageStartedAt || '',
      packageExpiresAt: item.packageExpiresAt || ''
    });
  }

  async function deleteUserItem(item) {
    if (!window.confirm(`Delete user ${item.username}?`)) return;
    await api(`/admin/users/${item.id}`, { method: 'DELETE' });
    if (userForm.id === item.id) setUserForm(emptyUserForm);
    setNotice('User deleted');
    await reload();
  }

  async function resetPassword(item) {
    const password = window.prompt(`New password for ${item.username}`);
    if (!password) return;
    await api(`/admin/users/${item.id}/reset-password`, { method: 'POST', body: JSON.stringify({ password }) });
    setNotice('Password reset');
  }

  async function saveCouponForm() {
    await api('/admin/coupons', { method: 'POST', body: JSON.stringify({ ...couponForm, assignedUserId: Number(couponForm.assignedUserId || 0) }) });
    setCouponForm({ code: '', messageCredit: 1000, assignedUserId: '', expiresAt: '' });
    setNotice('Coupon created');
    await reload();
  }

  async function saveSettingsForm() {
    await api('/admin/settings', { method: 'POST', body: JSON.stringify(settings) });
    setNotice('WhatsApp configuration saved');
    await reload();
  }

  return (
    <div className="adminGrid">
      <section className="band">
        <div className="sectionTitle"><ShieldCheck size={22} /><h2>User Roles & Rights</h2></div>
        <input placeholder="Role name" value={roleForm.name} onChange={(event) => setRoleForm({ ...roleForm, name: event.target.value })} />
        <input placeholder="WhatsApp daily limit, 0 unlimited" type="number" value={roleForm.messageLimit} onChange={(event) => setRoleForm({ ...roleForm, messageLimit: Number(event.target.value) })} />
        <div className="permissionGrid">
          {admin.permissions.map((permission) => (
            <label key={permission}><input type="checkbox" checked={roleForm.permissions.includes(permission)} onChange={(event) => setRoleForm({ ...roleForm, permissions: event.target.checked ? [...roleForm.permissions, permission] : roleForm.permissions.filter((item) => item !== permission) })} /> {permission}</label>
          ))}
        </div>
        <div className="formActions">
          <button className="primary" onClick={saveRoleForm}>{roleForm.id ? 'Update Role' : 'Save Role'}</button>
          {roleForm.id && <button className="secondary" onClick={() => setRoleForm(emptyRoleForm)}>Cancel Edit</button>}
        </div>
        <div className="miniList">
          {admin.roles.map((role) => (
            <div key={role.id}>
              <strong>{role.name}</strong>
              <span>{role.permissions.length} rights - limit {role.messageLimit || 'unlimited'}</span>
              <div className="rowActions">
                <button className="secondary" onClick={() => editRole(role)}>Edit</button>
                <button className="danger" onClick={() => deleteRoleItem(role)}>Delete</button>
              </div>
            </div>
          ))}
        </div>
      </section>

      <section className="band">
        <div className="sectionTitle"><Users size={22} /><h2>User Creation</h2></div>
        <input placeholder="Name" value={userForm.name} onChange={(event) => setUserForm({ ...userForm, name: event.target.value })} />
        <input placeholder="Username" value={userForm.username} onChange={(event) => setUserForm({ ...userForm, username: event.target.value })} />
        <input placeholder="Password" value={userForm.password} onChange={(event) => setUserForm({ ...userForm, password: event.target.value })} />
        <select value={userForm.roleId} onChange={(event) => setUserForm({ ...userForm, roleId: event.target.value })}><option value="">Select role</option>{admin.roles.map((role) => <option value={role.id} key={role.id}>{role.name}</option>)}</select>
        <select value={userForm.packageType} onChange={(event) => setUserForm({ ...userForm, packageType: event.target.value })}><option value="monthly">Monthly</option><option value="yearly">Yearly</option></select>
        <input placeholder="Package message limit" type="number" value={userForm.packageLimit} onChange={(event) => setUserForm({ ...userForm, packageLimit: Number(event.target.value) })} />
        <input placeholder="Bonus messages" type="number" value={userForm.bonusLimit} onChange={(event) => setUserForm({ ...userForm, bonusLimit: Number(event.target.value) })} />
        <input type="date" value={userForm.packageStartedAt} onChange={(event) => setUserForm({ ...userForm, packageStartedAt: event.target.value })} />
        <input type="date" value={userForm.packageExpiresAt} onChange={(event) => setUserForm({ ...userForm, packageExpiresAt: event.target.value })} />
        <input type="date" value={userForm.accessExpiresAt} onChange={(event) => setUserForm({ ...userForm, accessExpiresAt: event.target.value })} />
        <label className="checkLine"><input type="checkbox" checked={userForm.active} onChange={(event) => setUserForm({ ...userForm, active: event.target.checked })} /> Active user</label>
        <div className="formActions">
          <button className="primary" onClick={saveUserForm}>{userForm.id ? 'Update User' : 'Save User'}</button>
          {userForm.id && <button className="secondary" onClick={() => setUserForm(emptyUserForm)}>Cancel Edit</button>}
        </div>
        <div className="miniList">
          {admin.users.map((item) => (
            <div key={item.id}>
              <strong>{item.name}</strong>
              <span>{item.username} - {item.roleName} - {item.packageType} {item.packageLimit}+{item.bonusLimit} - {item.active ? 'Active' : 'Blocked'}</span>
              <div className="rowActions">
                <button className="secondary" onClick={() => editUser(item)}>Edit</button>
                <button className="secondary" onClick={() => resetPassword(item)}>Reset Password</button>
                <button className="danger" onClick={() => deleteUserItem(item)}>Delete</button>
              </div>
            </div>
          ))}
        </div>
      </section>

      <section className="band wide">
        <div className="sectionTitle"><MessageSquareText size={22} /><h2>Recharge Coupon</h2></div>
        <div className="settingsGrid">
          <input placeholder="Coupon code" value={couponForm.code} onChange={(event) => setCouponForm({ ...couponForm, code: event.target.value })} />
          <input placeholder="Message credit" type="number" value={couponForm.messageCredit} onChange={(event) => setCouponForm({ ...couponForm, messageCredit: Number(event.target.value) })} />
          <select value={couponForm.assignedUserId} onChange={(event) => setCouponForm({ ...couponForm, assignedUserId: event.target.value })}><option value="">Any user</option>{admin.users.map((user) => <option value={user.id} key={user.id}>{user.name}</option>)}</select>
          <input type="date" value={couponForm.expiresAt} onChange={(event) => setCouponForm({ ...couponForm, expiresAt: event.target.value })} />
        </div>
        <button className="primary" onClick={saveCouponForm}>Create Coupon</button>
        <div className="miniList">{(admin.coupons || []).map((coupon) => <div key={coupon.id}><strong>{coupon.code}</strong><span>{coupon.message_credit} messages · {coupon.status}</span></div>)}</div>
      </section>

      <section className="band wide">
        <div className="sectionTitle"><Settings size={22} /><h2>WhatsApp API Configuration</h2></div>
        <div className="settingsGrid">
          <input placeholder="Phone Number ID" value={settings.whatsapp_phone_number_id || ''} onChange={(event) => setSettings({ ...settings, whatsapp_phone_number_id: event.target.value })} />
          <input placeholder="Access Token" value={settings.whatsapp_access_token || ''} onChange={(event) => setSettings({ ...settings, whatsapp_access_token: event.target.value })} />
          <input placeholder="Template Name" value={settings.whatsapp_template_name || ''} onChange={(event) => setSettings({ ...settings, whatsapp_template_name: event.target.value })} />
          <input placeholder="Language Code" value={settings.whatsapp_language_code || ''} onChange={(event) => setSettings({ ...settings, whatsapp_language_code: event.target.value })} />
          <input placeholder="Global Daily Limit" type="number" value={settings.daily_message_limit || ''} onChange={(event) => setSettings({ ...settings, daily_message_limit: event.target.value })} />
        </div>
        <button className="primary" onClick={saveSettingsForm}>Save WhatsApp Config</button>
      </section>
    </div>
  );
}

function filterRowsByColumns(rows, columns, filters) {
  return rows.filter((row) => columns.every((column) => {
    const needle = String(filters[column.key] || '').trim().toLowerCase();
    if (!needle) return true;
    return String(row[column.key] ?? '').toLowerCase().includes(needle);
  }));
}

function paginateRows(rows, pageNo, pageSize = 50) {
  return rows.slice((pageNo - 1) * pageSize, pageNo * pageSize);
}

function ColumnFilterBar({ columns, filters, setFilters }) {
  return (
    <div className="columnFilterBar">
      {columns.map((column) => (
        <label key={column.key}>
          <span>{column.label}</span>
          <input
            placeholder={`Search ${column.label}`}
            value={filters[column.key] || ''}
            onChange={(event) => setFilters({ ...filters, [column.key]: event.target.value })}
          />
        </label>
      ))}
      <button className="secondary" onClick={() => setFilters({})}>Clear</button>
    </div>
  );
}

function TablePagination({ pageNo, setPageNo, total }) {
  const pageSize = 50;
  const totalPages = Math.max(1, Math.ceil(total / pageSize));
  const start = total === 0 ? 0 : ((pageNo - 1) * pageSize) + 1;
  const end = Math.min(total, pageNo * pageSize);

  return (
    <div className="paginationBar">
      <span>{start}-{end} of {total}</span>
      <button className="secondary" disabled={pageNo <= 1} onClick={() => setPageNo(Math.max(1, pageNo - 1))}>Previous</button>
      <strong>Page {pageNo} / {totalPages}</strong>
      <button className="secondary" disabled={pageNo >= totalPages} onClick={() => setPageNo(Math.min(totalPages, pageNo + 1))}>Next</button>
    </div>
  );
}

function DataTable({ columns, rows }) {
  const [columnFilters, setColumnFilters] = useState({});
  const [pageNo, setPageNo] = useState(1);
  const filteredRows = useMemo(() => filterRowsByColumns(rows, columns, columnFilters), [rows, columns, columnFilters]);
  const pagedRows = useMemo(() => paginateRows(filteredRows, pageNo), [filteredRows, pageNo]);

  return (
    <>
      <ColumnFilterBar columns={columns} filters={columnFilters} setFilters={(next) => { setColumnFilters(next); setPageNo(1); }} />
      <div className="tableWrap reportTable">
        <table>
          <thead><tr>{columns.map((column) => <th key={column.key}>{column.label}</th>)}</tr></thead>
          <tbody>
            {pagedRows.map((row, index) => (
              <tr key={row.id || row.propertyNo || index}>
                {columns.map((column) => <td key={`${column.key}-${row.id || index}`}>{column.type === 'money' ? currency(row[column.key]) : row[column.key] || '-'}</td>)}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <TablePagination pageNo={pageNo} setPageNo={setPageNo} total={filteredRows.length} />
    </>
  );
}

function Metric({ icon, label, value }) {
  return <article className="metric">{React.cloneElement(icon, { size: 22 })}<span>{label}</span><strong>{value}</strong></article>;
}

createRoot(document.getElementById('root')).render(<App />);
