
const { jsPDF } = window.jspdf;

const STORAGE_KEY = 'roman-landscaping-clean-v7';
const DEFAULT_SYNC_URL = 'https://script.google.com/macros/s/AKfycbyJWNDE0zM4_7zOoAZq5CO5C7GnQB8impRCkgI3fmjy7WZNFwumHKUj5eHh-x9hDkf5Bg/exec';
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const TAB_META = {
  dashboard: { icon: '📊', title: 'Dashboard', subtitle: 'Panel de Control', blurbUS: 'Overview of billed, collected, pending, and overdue balances.', blurbMX: 'Resumen de facturado, cobrado, pendiente y vencido.' },
  invoices: { icon: '📄', title: 'Invoices', subtitle: 'Facturas', blurbUS: 'Track invoices by vendor and month. Download, mark paid, or batch follow up.', blurbMX: 'Rastree facturas por cliente y mes. Descargue, marque pagadas o haga seguimiento por lote.' },
  overdue: { icon: '⏰', title: 'Overdue Payments', subtitle: 'Pagos Vencidos', blurbUS: 'Focus only on unpaid and overdue invoices. Download follow-up packs by vendor.', blurbMX: 'Enfóquese solo en facturas pendientes y vencidas. Descargue paquetes de seguimiento por cliente.' },
  vendors: { icon: '👥', title: 'Vendors / Clients', subtitle: 'Clientes', blurbUS: 'Manage clients, prefixes, property names, and contact details.', blurbMX: 'Administre clientes, prefijos, propiedades y datos de contacto.' },
  appraisals: { icon: '📋', title: 'Appraisals', subtitle: 'Presupuestos', blurbUS: 'Create service appraisals and convert approved ones into invoices.', blurbMX: 'Cree presupuestos y convierta los aprobados en facturas.' },
  createInvoice: { icon: '🖨️', title: 'Invoice Generator', subtitle: 'Generador de Facturas', blurbUS: 'Build clean invoice PDFs with sequence numbering, image-to-lines, and print-safe layout.', blurbMX: 'Cree PDFs limpios con numeración secuencial, imagen a líneas y diseño listo para imprimir.' },
  services: { icon: '🌿', title: 'Service Catalog', subtitle: 'Catálogo de Servicios', blurbUS: 'Store reusable services with preset prices for faster invoice creation.', blurbMX: 'Guarde servicios reutilizables con precios preestablecidos para crear facturas más rápido.' },
  settings: { icon: '⚙️', title: 'Settings', subtitle: 'Ajustes', blurbUS: 'Sync, print, theme, payment, and OCR configuration in one place.', blurbMX: 'Sincronización, impresión, tema, pagos y OCR en un solo lugar.' }
};

const sampleServices = [
  ['Spring clean up – one-time service', 480],
  ['Disposal of trash from spring clean up', 40],
  ['Trimming first time – includes disposal of trash', 380],
  ['Trimming second time – includes disposal of trash', 300],
  ['Monthly cut grass', 290],
  ['6 yards red mulch (material only)', 540],
  ['Labor to place mulch and deliver', 480],
  ['Grass reseeding; fertilize; grub worm control (labor+material)', 500],
  ['Fall clean up Oct/Nov – each month', 480],
];

let state = loadState();
let currentTab = 'dashboard';
let selectedYear = new Date().getFullYear();
let selectedMonth = MONTHS[new Date().getMonth()];
let toastTimer = null;
let draftSeed = null;
let appRoot = document.getElementById('app');
const imageCache = {};

bootstrap();

async function bootstrap() {
  applyTheme();
  await warmImages();
  renderApp();
  if (state.settings.syncUrl) {
    await loadFromRemote({ silent: true });
  }
}

function defaultState() {
  const vendorId = uid('vendor');
  const appraisalId = uid('appraisal');
  const now = new Date();
  const year = now.getFullYear();
  const createdDate = toInputDate(now);
  const sampleVendor = {
    id: vendorId,
    name: 'Yajardo',
    property: '1 David Lane',
    prefix: 'DL',
    contact: '',
    email: '',
    address: '',
    notes: 'Set property'
  };
  return {
    vendors: [sampleVendor],
    invoices: {},
    appraisals: [{
      id: appraisalId,
      vendorId,
      property: '1 David Lane',
      status: 'Draft',
      createdDate,
      notes: 'Imported starting point',
      items: [
        { name: 'Spring clean up – one-time service in April', amount: 700 },
        { name: 'Maintenance 1 time per month (May thru September)', amount: 400 },
      ],
    }],
    services: sampleServices.map(([name, price]) => ({ id: uid('service'), name, price })),
    settings: {
      syncUrl: DEFAULT_SYNC_URL,
      autoSync: false,
      themeMode: 'dark',
      taxRate: 8.375,
      companyName: 'ROMAN DESIGN & LANDSCAPING INC.',
      ownerName: 'Roman Cuateco',
      companyAddress: '516 Tarrytown Road, White Plains, NY 10607',
      companyEmail: 'RomanD.Landscaping@gmail.com',
      phone: '914-467-9927',
      paymentNote: 'Please send payment to P.O.box: P.O. Box 767 Elmsford NY 10523',
      anthropicApiKey: '',
      anthropicModel: 'claude-sonnet-4-6',
      promoImage: '/assets/promo-default-v7.png',
      headerLogo: '/assets/header-logo-v7.png',
      defaultDueDays: 14,
      compactView: false,
      defaultPrefixMode: 'vendor',
    }
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultState();
    const parsed = JSON.parse(raw);
    const merged = deepMerge(defaultState(), parsed || {});
    normalizeState(merged);
    return merged;
  } catch (err) {
    console.warn(err);
    return defaultState();
  }
}

function saveState({ sync = false, silent = false } = {}) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  applyTheme();
  if (sync || state.settings.autoSync) {
    syncToRemote({ silent });
  }
}

function normalizeState(target) {
  target.vendors = Array.isArray(target.vendors) ? target.vendors : [];
  target.services = Array.isArray(target.services) ? target.services : [];
  target.appraisals = Array.isArray(target.appraisals) ? target.appraisals : [];
  target.invoices = target.invoices && typeof target.invoices === 'object' ? target.invoices : {};
  target.settings = deepMerge(defaultState().settings, target.settings || {});
  target.vendors.forEach(v => { if (!v.prefix) v.prefix = derivePrefix(v); });
  Object.values(target.invoices).flat().forEach(inv => {
    inv.items = Array.isArray(inv.items) && inv.items.length ? inv.items : [{ name: inv.description || 'Invoice total', amount: Number(inv.amount || 0) }];
    inv.invoiceNumber = inv.invoiceNumber || nextInvoiceNumber(inv.prefix || derivePrefix(findVendor(inv.vendorId)) || 'INV');
    inv.prefix = inv.prefix || deriveSuffix(inv.invoiceNumber) || derivePrefix(findVendor(inv.vendorId));
    inv.amount = Number(inv.amount || invoiceItemsTotal(inv.items));
    inv.billingHeader = inv.billingHeader || inv.snapshot?.billingHeader || '';
  });
  target.appraisals.forEach(ap => {
    ap.items = Array.isArray(ap.items) ? ap.items : [];
    ap.notes = ap.notes || '';
    ap.status = ap.status || 'Draft';
  });
}

function deepMerge(base, incoming) {
  if (Array.isArray(base)) return Array.isArray(incoming) ? incoming : base;
  const out = { ...base };
  for (const [key, value] of Object.entries(incoming || {})) {
    if (value && typeof value === 'object' && !Array.isArray(value) && base[key] && typeof base[key] === 'object' && !Array.isArray(base[key])) {
      out[key] = deepMerge(base[key], value);
    } else {
      out[key] = value;
    }
  }
  return out;
}

function uid(prefix = 'id') {
  return `${prefix}_${Math.random().toString(36).slice(2, 9)}${Date.now().toString(36).slice(-4)}`;
}

function money(value) {
  const number = Number(value || 0);
  return number.toLocaleString('en-US', { style: 'currency', currency: 'USD' });
}

function shortMoney(value) {
  return Number(value || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function parseMoney(value) {
  if (value == null) return 0;
  return Number(String(value).replace(/[^0-9.-]/g, '')) || 0;
}

function toInputDate(date) {
  const d = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(d.getTime())) return '';
  return d.toISOString().slice(0, 10);
}

function todayString() {
  return toInputDate(new Date());
}

function addDays(input, days) {
  const d = new Date(input || new Date());
  d.setDate(d.getDate() + Number(days || 0));
  return toInputDate(d);
}

function getAllInvoices() {
  return Object.values(state.invoices || {}).flat();
}

function getInvoicesForVendorYear(vendorId, year) {
  return state.invoices[`${vendorId}-${year}`] || [];
}

function invoiceItemsTotal(items = []) {
  return items.reduce((sum, item) => sum + parseMoney(item.amount), 0);
}

function buildInvoiceKey(vendorId, year) {
  return `${vendorId}-${year}`;
}

function findVendor(id) {
  return state.vendors.find(v => v.id === id) || null;
}

function derivePrefix(vendor) {
  if (!vendor) return 'INV';
  if (vendor.prefix) return vendor.prefix.toUpperCase().replace(/[^A-Z0-9]/g, '').slice(-6) || 'INV';
  const source = (vendor.property || vendor.name || 'INV').replace(/[^A-Za-z0-9 ]/g, ' ').trim();
  const parts = source.split(/\s+/).filter(Boolean).filter(part => !/^\d+$/.test(part));
  const letters = parts.slice(0, 3).map(part => part[0].toUpperCase()).join('');
  return letters || 'INV';
}

function deriveSuffix(invoiceNumber) {
  const match = String(invoiceNumber || '').trim().toUpperCase().match(/^(\d+)([A-Z]+)$/);
  return match ? match[2] : '';
}

function nextInvoiceNumber(prefix) {
  const clean = (prefix || 'INV').toUpperCase().replace(/[^A-Z0-9]/g, '') || 'INV';
  let max = 0;
  getAllInvoices().forEach(inv => {
    const match = String(inv.invoiceNumber || '').toUpperCase().match(/^(\d+)([A-Z0-9]+)$/);
    if (match && match[2] === clean) {
      max = Math.max(max, Number(match[1]));
    }
  });
  return `${max + 1}${clean}`;
}

function invoiceStatus(invoice) {
  if (invoice.paid) return 'paid';
  if (isOverdue(invoice)) return 'overdue';
  return 'pending';
}

function isOverdue(invoice) {
  if (invoice.paid || !invoice.dueDate) return false;
  const due = new Date(invoice.dueDate);
  const today = new Date();
  due.setHours(0,0,0,0);
  today.setHours(0,0,0,0);
  return due < today;
}

function totals() {
  const all = getAllInvoices().filter(inv => Number(inv.year) === Number(selectedYear));
  const billed = all.reduce((sum, inv) => sum + parseMoney(inv.amount), 0);
  const collected = all.filter(inv => inv.paid).reduce((sum, inv) => sum + parseMoney(inv.amount), 0);
  const outstanding = all.filter(inv => !inv.paid).reduce((sum, inv) => sum + parseMoney(inv.amount), 0);
  const overdue = all.filter(inv => isOverdue(inv)).reduce((sum, inv) => sum + parseMoney(inv.amount), 0);
  return { billed, collected, outstanding, overdue };
}

function monthlyBreakdown() {
  return MONTHS.map(month => {
    const invoices = getAllInvoices().filter(inv => Number(inv.year) === Number(selectedYear) && inv.month === month);
    return {
      month,
      billed: invoices.reduce((sum, inv) => sum + parseMoney(inv.amount), 0),
      collected: invoices.filter(inv => inv.paid).reduce((sum, inv) => sum + parseMoney(inv.amount), 0),
    };
  });
}

function applyTheme() {
  const mode = state.settings.themeMode || 'dark';
  const finalMode = mode === 'system'
    ? (window.matchMedia('(prefers-color-scheme: light)').matches ? 'light' : 'dark')
    : mode;
  document.body.setAttribute('data-theme', finalMode);
}

async function warmImages() {
  const paths = ['/assets/header-logo-v7.png', '/assets/promo-default-v7.png'];
  await Promise.all(paths.map(loadImageAsDataUrl));
}

async function loadImageAsDataUrl(path) {
  if (!path) return null;
  if (imageCache[path]) return imageCache[path];
  try {
    const response = await fetch(path);
    if (!response.ok) throw new Error(`Image not found: ${path}`);
    const blob = await response.blob();
    const data = await blobToDataUrl(blob);
    imageCache[path] = data;
    return data;
  } catch (error) {
    console.warn(error);
    return null;
  }
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function inferImageFormat(dataUrl) {
  if (!dataUrl || typeof dataUrl !== 'string') return 'PNG';
  const match = dataUrl.match(/^data:image\/([a-zA-Z0-9+.-]+);base64,/i);
  if (!match) return 'PNG';
  const subtype = match[1].toLowerCase();
  if (subtype.includes('png')) return 'PNG';
  if (subtype.includes('jpe') || subtype.includes('jpg')) return 'JPEG';
  if (subtype.includes('webp')) return 'WEBP';
  return 'PNG';
}

function renderApp() {
  const meta = TAB_META[currentTab];
  appRoot.innerHTML = `
    <div class="app-shell">
      <header class="topbar">
        <div class="brand-wrap">
          <div class="brand-side"><img src="/assets/header-logo-v7.png" alt="Roman logo"></div>
          <div class="brand-center">
            <h1 class="brand-title">Roman Design & Landscaping Inc.</h1>
            <p class="brand-subtitle">Invoice Manager — ${selectedYear}</p>
            <div class="topbar-actions">
              <div class="year-switch">
                <button data-action="year-prev">◀ ${selectedYear - 1}</button>
                <button class="year-current" type="button">${selectedYear}</button>
                <button data-action="year-next">${selectedYear + 1} ▶</button>
              </div>
              <button class="primary-btn" data-action="open-vendor-modal">+ Vendor</button>
              <button class="ghost-btn" data-action="sync-now">⭮ Sync</button>
            </div>
          </div>
          <div class="brand-side"><img src="/assets/header-logo-v7.png" alt="Roman logo"></div>
        </div>
        <div class="utility-row">
          <span class="utility-chip">${state.settings.autoSync ? '🟢 Auto-Sync ON' : '⚪ Auto-Sync OFF'}</span>
          <span class="utility-chip">Theme: ${displayThemeLabel()}</span>
          <span class="utility-chip">OCR: ${state.settings.anthropicApiKey ? 'Ready' : 'Set API key'}</span>
          <span class="utility-chip">Sequence sample: <span class="code-pill">1DL → 2DL → 3DL</span></span>
        </div>
        <nav class="navbar">
          ${Object.entries(TAB_META).map(([key, value]) => `
            <button class="tab-btn ${key === currentTab ? 'active' : ''}" data-action="switch-tab" data-tab="${key}">${value.icon} ${value.title}</button>
          `).join('')}
        </nav>
      </header>
      <main class="main-wrap">
        <section class="hero">
          <h2>${meta.title}</h2>
          <div class="subtitle">${meta.subtitle}</div>
          <div class="accent-line"></div>
        </section>
        ${renderCurrentPage()}
      </main>
      ${renderVendorModal()}
      ${renderAppraisalModal()}
      <div id="toast" class="toast"></div>
    </div>
  `;
  bindEvents();
  if (currentTab === 'createInvoice') {
    hydrateInvoiceForm();
  }
}

function displayThemeLabel() {
  const mode = state.settings.themeMode || 'dark';
  return mode[0].toUpperCase() + mode.slice(1);
}

function renderCurrentPage() {
  switch (currentTab) {
    case 'dashboard': return renderDashboard();
    case 'invoices': return renderInvoicesPage();
    case 'overdue': return renderOverduePage();
    case 'vendors': return renderVendorsPage();
    case 'appraisals': return renderAppraisalsPage();
    case 'createInvoice': return renderCreateInvoicePage();
    case 'services': return renderServicesPage();
    case 'settings': return renderSettingsPage();
    default: return renderDashboard();
  }
}

function renderDashboard() {
  const t = totals();
  const breakdown = monthlyBreakdown();
  const maxBar = Math.max(1, ...breakdown.map(row => row.billed));
  return `
    <section class="section-card section-copy">
      <div><strong>US</strong> ${TAB_META.dashboard.blurbUS}</div>
      <div><strong>MX</strong> ${TAB_META.dashboard.blurbMX}</div>
    </section>
    <section class="grid-4">
      <div class="metric-card blue"><div>💰</div><div class="metric-value">${money(t.billed)}</div><h4>Billed / Facturado</h4></div>
      <div class="metric-card green"><div>✅</div><div class="metric-value">${money(t.collected)}</div><h4>Collected / Cobrado</h4></div>
      <div class="metric-card gold"><div>⏳</div><div class="metric-value">${money(t.outstanding)}</div><h4>Outstanding / Pendiente</h4></div>
      <div class="metric-card red"><div>🚨</div><div class="metric-value">${money(t.overdue)}</div><h4>Overdue / Vencido</h4></div>
    </section>
    <section class="section-card">
      <h3>Monthly Revenue / Ingresos Mensuales</h3>
      <div class="progress-bars">
        ${breakdown.map(row => `
          <div class="bar-row">
            <div>${row.month}</div>
            <div class="bar-track"><div class="bar-fill" style="width:${(row.billed / maxBar) * 100}%"></div></div>
            <div>${money(row.billed)}</div>
          </div>
        `).join('')}
      </div>
    </section>
    <section class="grid-2">
      <div class="section-card">
        <h3>Pending Next Actions</h3>
        <div class="stack-list">
          <div class="stack-item"><strong>${getAllInvoices().filter(i => !i.paid).length}</strong><span class="small-copy">Open invoices</span></div>
          <div class="stack-item"><strong>${getAllInvoices().filter(i => isOverdue(i)).length}</strong><span class="small-copy">Overdue invoices</span></div>
          <div class="stack-item"><strong>${state.vendors.length}</strong><span class="small-copy">Active vendors / clients</span></div>
        </div>
      </div>
      <div class="section-card">
        <h3>Readability upgrades already built in</h3>
        <div class="stack-list">
          <div class="stack-item">Higher contrast, larger body text, clearer spacing, and stronger card separation.</div>
          <div class="stack-item">Dark / Light / System theme selector and lighter hero overlays for easier reading.</div>
          <div class="stack-item">Outline is used only on key headings and number-heavy display areas — not paragraph text.</div>
        </div>
      </div>
    </section>
  `;
}

function renderInvoicesPage() {
  const vendorGroups = state.vendors.map(vendor => ({ vendor, invoices: getInvoicesForVendorYear(vendor.id, selectedYear).filter(inv => inv.month === selectedMonth) }))
    .filter(group => group.invoices.length);
  return `
    <section class="section-card section-copy">
      <div><strong>US</strong> ${TAB_META.invoices.blurbUS}</div>
      <div><strong>MX</strong> ${TAB_META.invoices.blurbMX}</div>
    </section>
    <section class="section-card">
      <div class="actions-row">
        ${MONTHS.map(month => `<button class="pill-btn ${month === selectedMonth ? 'active' : ''}" data-action="select-month" data-month="${month}">${month}</button>`).join('')}
      </div>
      ${vendorGroups.length ? vendorGroups.map(group => renderInvoiceGroup(group.vendor, group.invoices)).join('') : `<div class="notice">No invoices found for ${selectedMonth} ${selectedYear}. Create one from the Invoice Generator.</div>`}
    </section>
  `;
}

function renderInvoiceGroup(vendor, invoices) {
  const total = invoices.reduce((sum, inv) => sum + parseMoney(inv.amount), 0);
  return `
    <details class="invoice-group" open>
      <summary>
        <div>${vendor.name}</div>
        <div>${money(total)} · ${invoices.length} item${invoices.length > 1 ? 's' : ''}</div>
      </summary>
      <div class="invoice-list">
        ${invoices.map(inv => renderInvoiceRow(inv, vendor)).join('')}
        <button class="primary-btn" data-action="open-create-for-vendor" data-vendor-id="${vendor.id}">+ Add invoice for ${selectedMonth}</button>
      </div>
    </details>
  `;
}

function renderInvoiceRow(inv, vendor) {
  return `
    <div class="invoice-row">
      <div class="invoice-meta">
        <div class="actions-row">
          <span class="invoice-number">${inv.invoiceNumber || '—'}</span>
          <span class="status-pill ${invoiceStatus(inv)}">${invoiceStatus(inv).toUpperCase()}</span>
          <span class="small-copy">${money(inv.amount)}</span>
        </div>
        <div class="small-copy">Sent ${inv.dateSent || '—'} · Due ${inv.dueDate || '—'} · ${inv.description || 'Invoice'}</div>
      </div>
      <div class="actions-row">
        <button class="ghost-btn" data-action="download-invoice" data-id="${inv.id}">Download PDF</button>
        <button class="ghost-btn" data-action="toggle-paid" data-id="${inv.id}">${inv.paid ? 'Mark Unpaid' : 'Mark Paid'}</button>
        <button class="ghost-btn" data-action="delete-invoice" data-id="${inv.id}">Delete</button>
      </div>
    </div>
  `;
}

function renderOverduePage() {
  const vendorRows = state.vendors.map(vendor => {
    const related = getInvoicesForVendorYear(vendor.id, selectedYear).filter(inv => !inv.paid);
    if (!related.length) return null;
    const overdue = related.filter(isOverdue);
    return { vendor, related, overdue };
  }).filter(Boolean).sort((a, b) => b.overdue.length - a.overdue.length || b.related.length - a.related.length);
  return `
    <section class="section-card section-copy">
      <div><strong>US</strong> ${TAB_META.overdue.blurbUS}</div>
      <div><strong>MX</strong> ${TAB_META.overdue.blurbMX}</div>
    </section>
    <section class="grid-3">
      <div class="metric-card gold"><div class="metric-value">${getAllInvoices().filter(i => !i.paid).length}</div><h4>Open invoices</h4></div>
      <div class="metric-card red"><div class="metric-value">${getAllInvoices().filter(isOverdue).length}</div><h4>Overdue invoices</h4></div>
      <div class="metric-card blue"><div class="metric-value">${money(getAllInvoices().filter(i => !i.paid).reduce((s, i) => s + parseMoney(i.amount), 0))}</div><h4>Follow-up balance</h4></div>
    </section>
    <section class="section-card stack-list">
      ${vendorRows.length ? vendorRows.map(({ vendor, related, overdue }) => `
        <div class="overdue-row">
          <div>
            <div class="invoice-number">${vendor.name}</div>
            <div class="small-copy">${related.length} open · ${overdue.length} overdue · ${money(related.reduce((s, inv) => s + parseMoney(inv.amount), 0))}</div>
          </div>
          <div class="actions-row">
            <button class="ghost-btn" data-action="download-followup-pack" data-vendor-id="${vendor.id}">Download pending pack</button>
            <button class="ghost-btn" data-action="copy-followup-text" data-vendor-id="${vendor.id}">Copy follow-up text</button>
            <button class="ghost-btn" data-action="open-create-for-vendor" data-vendor-id="${vendor.id}">New invoice</button>
          </div>
        </div>
      `).join('') : `<div class="notice">No pending or overdue invoices right now.</div>`}
    </section>
  `;
}

function renderVendorsPage() {
  return `
    <section class="section-card section-copy">
      <div><strong>US</strong> ${TAB_META.vendors.blurbUS}</div>
      <div><strong>MX</strong> ${TAB_META.vendors.blurbMX}</div>
    </section>
    <section class="section-card">
      <div class="actions-row"><button class="primary-btn" data-action="open-vendor-modal">+ Add New Vendor / Agregar Nuevo Cliente</button></div>
      <div class="stack-list" style="margin-top:14px;">
        ${state.vendors.map(vendor => renderVendorCard(vendor)).join('')}
      </div>
    </section>
  `;
}

function renderVendorCard(vendor) {
  const invoices = getInvoicesForVendorYear(vendor.id, selectedYear);
  const billed = invoices.reduce((sum, inv) => sum + parseMoney(inv.amount), 0);
  const collected = invoices.filter(inv => inv.paid).reduce((sum, inv) => sum + parseMoney(inv.amount), 0);
  const outstanding = billed - collected;
  return `
    <div class="vendor-card">
      <div class="vendor-avatar">${(vendor.name || 'V').trim()[0]?.toUpperCase() || 'V'}</div>
      <div class="vendor-head">
        <div class="invoice-number">${vendor.name}</div>
        <div class="small-copy">${vendor.property || 'No property set'} · Prefix ${vendor.prefix || derivePrefix(vendor)}</div>
        <div class="small-copy">Contact: ${vendor.contact || '—'} · Email: ${vendor.email || '—'}</div>
        <div class="kpi-strip">
          <div class="kpi-pill">Billed <strong>${money(billed)}</strong></div>
          <div class="kpi-pill">Collected <strong>${money(collected)}</strong></div>
          <div class="kpi-pill">Outstanding <strong>${money(outstanding)}</strong></div>
        </div>
      </div>
      <div class="actions-row">
        <button class="ghost-btn" data-action="open-create-for-vendor" data-vendor-id="${vendor.id}">Invoice</button>
        <button class="ghost-btn" data-action="edit-vendor" data-id="${vendor.id}">Edit</button>
        <button class="ghost-btn" data-action="delete-vendor" data-id="${vendor.id}">Delete</button>
      </div>
    </div>
  `;
}

function renderAppraisalsPage() {
  return `
    <section class="section-card section-copy">
      <div><strong>US</strong> ${TAB_META.appraisals.blurbUS}</div>
      <div><strong>MX</strong> ${TAB_META.appraisals.blurbMX}</div>
    </section>
    <section class="section-card">
      <div class="actions-row"><button class="primary-btn" data-action="open-appraisal-modal">+ New Appraisal / Nuevo Presupuesto</button></div>
      <div class="stack-list" style="margin-top:14px;">
        ${state.appraisals.length ? state.appraisals.map(ap => renderAppraisalCard(ap)).join('') : `<div class="notice">No appraisals yet.</div>`}
      </div>
    </section>
  `;
}

function renderAppraisalCard(ap) {
  const vendor = findVendor(ap.vendorId);
  const total = invoiceItemsTotal(ap.items);
  return `
    <div class="stack-item">
      <div class="actions-row">
        <span class="invoice-number">${vendor?.name || 'Unknown vendor'}</span>
        <span class="status-pill neutral">${ap.status}</span>
      </div>
      <div class="small-copy">${ap.property || ''} · ${ap.createdDate || ''}</div>
      <div class="small-copy">${ap.items.length} line${ap.items.length > 1 ? 's' : ''} · ${money(total)}</div>
      <div class="actions-row">
        <button class="ghost-btn" data-action="convert-appraisal" data-id="${ap.id}">Convert to invoice</button>
        <button class="ghost-btn" data-action="delete-appraisal" data-id="${ap.id}">Delete</button>
      </div>
    </div>
  `;
}

function renderCreateInvoicePage() {
  return `
    <section class="section-card section-copy">
      <div><strong>US</strong> ${TAB_META.createInvoice.blurbUS}</div>
      <div><strong>MX</strong> ${TAB_META.createInvoice.blurbMX}</div>
    </section>
    <section class="grid-2">
      <div class="section-card">
        <h3>Invoice Details / Detalles</h3>
        <div class="inline-note">The PDF generator now always tries to place the QR promo image on the final page. If the services are short, it shrinks the promo to keep everything on page 1. If there are many lines, the invoice automatically continues to page 2.</div>
        <form id="invoice-form" class="form-grid" style="margin-top:16px;">
          <div class="field span-6">
            <label>Vendor / Cliente *</label>
            <select id="invoiceVendorId">${renderVendorOptions()}</select>
          </div>
          <div class="field span-6">
            <label>Property / Propiedad</label>
            <input id="invoiceProperty" type="text" />
          </div>
          <div class="field span-4">
            <label>Prefix / Prefijo</label>
            <input id="invoicePrefix" type="text" maxlength="6" />
          </div>
          <div class="field span-4">
            <label>Invoice Number / Número</label>
            <input id="invoiceNumber" type="text" />
          </div>
          <div class="field span-4">
            <label>Description / Descripción</label>
            <input id="invoiceDescription" type="text" value="Appraisal For Services" />
          </div>
          <div class="field span-4">
            <label>Date Sent / Fecha</label>
            <input id="invoiceDateSent" type="date" value="${todayString()}" />
          </div>
          <div class="field span-4">
            <label>Due Date / Vencimiento</label>
            <input id="invoiceDueDate" type="date" value="${addDays(todayString(), state.settings.defaultDueDays)}" />
          </div>
          <div class="field span-4">
            <label>Notes / Notas</label>
            <input id="invoiceNotes" type="text" placeholder="Optional note" />
          </div>
          <div class="field span-12">
            <label>Billing Header / Bill To (optional)</label>
            <textarea id="invoiceBillingHeader" rows="3" placeholder="Only use this when a vendor needs a billing header above the invoice details, for example:
Brentwood Condominium
c/o Garthchester Realty 440 Mamaroneck Ave Suite S-512, Harrison NY 10528"></textarea>
          </div>
        </form>
        <div class="section-card" style="margin-top:16px; padding:14px; background: var(--bg-soft);">
          <div class="actions-row" style="justify-content:space-between;">
            <h4 style="margin:0;">Service Lines / Líneas de Servicio</h4>
            <div class="actions-row">
              <button class="ghost-btn" data-action="add-line-row">+ Add line</button>
              <button class="ghost-btn" data-action="clear-line-rows">Clear</button>
            </div>
          </div>
          <div id="lineItemsWrap" class="service-list" style="margin-top:12px;"></div>
        </div>
        <div class="actions-row" style="margin-top:16px;">
          <button class="primary-btn" data-action="save-invoice">Save invoice</button>
          <button class="ghost-btn" data-action="save-download-invoice">Save + Download PDF</button>
          <button class="ghost-btn" data-action="preview-current-invoice">Download PDF only</button>
        </div>
      </div>
      <div class="stack-list">
        <div class="section-card">
          <h3>Image to Lines / Imagen a Líneas</h3>
          <p class="helper">Upload a photo or scan of an invoice/list. The Netlify OCR function sends the image to Anthropic's Messages API using image input blocks and extracts clean JSON line items.</p>
          <div class="field" style="margin-top:12px;">
            <label>Invoice image</label>
            <input id="ocrImageInput" type="file" accept="image/png,image/jpeg,image/webp,image/gif" />
          </div>
          <div class="actions-row" style="margin-top:12px;">
            <button class="primary-btn" data-action="extract-lines">Extract lines</button>
            <span id="ocrStatus" class="small-copy">Ready</span>
          </div>
        </div>
        <div class="section-card">
          <h3>Quick Add Services</h3>
          <div class="service-chip-wrap">
            ${state.services.map(service => `<button class="chip" data-action="quick-add-service" data-id="${service.id}">${service.name} — ${shortMoney(service.price)}</button>`).join('')}
          </div>
        </div>
        <div class="section-card">
          <h3>Sequence preview</h3>
          <div class="inline-note">When a vendor prefix is set to <span class="code-pill">DL</span>, the next invoice number automatically becomes <span class="code-pill">${nextInvoiceNumber('DL')}</span>. You can still edit it manually before saving.</div>
        </div>
      </div>
    </section>
  `;
}

function renderServicesPage() {
  return `
    <section class="section-card section-copy">
      <div><strong>US</strong> ${TAB_META.services.blurbUS}</div>
      <div><strong>MX</strong> ${TAB_META.services.blurbMX}</div>
    </section>
    <section class="section-card">
      <div id="servicesWrap" class="service-list">
        ${state.services.map(service => renderServiceRow(service)).join('')}
      </div>
      <div class="service-row" style="margin-top:14px;">
        <input id="newServiceName" type="text" placeholder="New service / Nuevo servicio" />
        <input id="newServicePrice" type="number" step="0.01" placeholder="$" />
        <button class="primary-btn" data-action="add-service">+</button>
      </div>
    </section>
  `;
}

function renderServiceRow(service) {
  return `
    <div class="service-row">
      <input type="text" data-service-field="name" data-id="${service.id}" value="${escapeHtml(service.name)}" />
      <input type="number" step="0.01" data-service-field="price" data-id="${service.id}" value="${service.price}" />
      <button class="ghost-btn" data-action="delete-service" data-id="${service.id}">🗑️</button>
    </div>
  `;
}

function renderSettingsPage() {
  return `
    <section class="section-card section-copy">
      <div><strong>US</strong> ${TAB_META.settings.blurbUS}</div>
      <div><strong>MX</strong> ${TAB_META.settings.blurbMX}</div>
    </section>
    <section class="grid-2">
      <div class="section-card">
        <h3>Google Sheets Sync</h3>
        <div class="field">
          <label>Web App URL</label>
          <input id="settingsSyncUrl" type="url" value="${escapeHtml(state.settings.syncUrl)}" />
        </div>
        <div class="actions-row" style="margin-top:14px;">
          <button class="ghost-btn" data-action="toggle-autosync">${state.settings.autoSync ? 'Disable auto-sync' : 'Enable auto-sync'}</button>
          <button class="primary-btn" data-action="sync-now">Sync now</button>
          <button class="ghost-btn" data-action="load-remote">Load remote</button>
        </div>
        <div class="notice">Use the updated <span class="code-pill">Code.gs</span> in this package. It now preserves invoice numbers, line items JSON, vendor prefixes, and settings.</div>
      </div>
      <div class="section-card">
        <h3>Theme + readability</h3>
        <div class="theme-picker">
          ${['dark','light','system'].map(mode => `<button class="ghost-btn ${state.settings.themeMode === mode ? 'active' : ''}" data-action="set-theme" data-theme="${mode}">${mode[0].toUpperCase()+mode.slice(1)}</button>`).join('')}
        </div>
        <div class="notice">For a 100/100 feel: keep outlines only on headings, body text at 16px+, one primary CTA per screen, strong status colors, and enough whitespace around money and due dates.</div>
      </div>
      <div class="section-card">
        <h3>Company + invoice defaults</h3>
        <div class="form-grid">
          <div class="field span-6"><label>Owner Name</label><input id="settingsOwnerName" type="text" value="${escapeHtml(state.settings.ownerName)}" /></div>
          <div class="field span-6"><label>Company Name</label><input id="settingsCompanyName" type="text" value="${escapeHtml(state.settings.companyName)}" /></div>
          <div class="field span-6"><label>Phone</label><input id="settingsPhone" type="text" value="${escapeHtml(state.settings.phone)}" /></div>
          <div class="field span-6"><label>Email</label><input id="settingsEmail" type="email" value="${escapeHtml(state.settings.companyEmail)}" /></div>
          <div class="field span-12"><label>Address</label><input id="settingsAddress" type="text" value="${escapeHtml(state.settings.companyAddress)}" /></div>
          <div class="field span-8"><label>Payment Note</label><input id="settingsPaymentNote" type="text" value="${escapeHtml(state.settings.paymentNote)}" /></div>
          <div class="field span-2"><label>Tax Rate %</label><input id="settingsTaxRate" type="number" step="0.001" value="${state.settings.taxRate}" /></div>
          <div class="field span-2"><label>Default due days</label><input id="settingsDueDays" type="number" step="1" value="${state.settings.defaultDueDays}" /></div>
        </div>
      </div>
      <div class="section-card">
        <h3>Anthropic OCR</h3>
        <div class="field"><label>API Key (optional if set in Netlify env)</label><input id="settingsAnthropicApiKey" type="password" value="${escapeHtml(state.settings.anthropicApiKey)}" /></div>
        <div class="field"><label>Model</label><input id="settingsAnthropicModel" type="text" value="${escapeHtml(state.settings.anthropicModel)}" /></div>
        <div class="helper">Best practice: set <span class="code-pill">ANTHROPIC_API_KEY</span> as a Netlify environment variable so the browser never sees the key. The function will fall back to this field only if the env variable is not set.</div>
      </div>
      <div class="section-card">
        <h3>Promo image for PDF footer</h3>
        <div class="preview-box"><img id="promoPreview" src="${state.settings.promoImage || '/assets/promo-default-v7.png'}" alt="Promo preview"></div>
        <div class="field" style="margin-top:12px;"><label>Upload replacement promo image</label><input id="settingsPromoImage" type="file" accept="image/png,image/jpeg,image/webp" /></div>
        <div class="helper">The PDF generator uses this image on the final page and scales it to fit the remaining space. If there are too many service lines, it moves the image to a new page automatically.</div>
      </div>
    </section>
    <section class="actions-row" style="margin-top:18px;"><button class="primary-btn" data-action="save-settings">Save settings</button></section>
  `;
}

function renderVendorOptions(selected = '') {
  return state.vendors.map(v => `<option value="${v.id}" ${selected === v.id ? 'selected' : ''}>${escapeHtml(v.name)}${v.property ? ` — ${escapeHtml(v.property)}` : ''}</option>`).join('');
}

function renderVendorModal() {
  return `
    <div id="vendorModalBackdrop" class="modal-backdrop">
      <div class="modal">
        <div class="modal-header">
          <h3 id="vendorModalTitle">Vendor / Client</h3>
          <button class="close-btn" data-action="close-vendor-modal">✕</button>
        </div>
        <form id="vendorForm" class="form-grid">
          <input type="hidden" id="vendorIdField" />
          <div class="field span-6"><label>Name / Nombre</label><input id="vendorNameField" type="text" required /></div>
          <div class="field span-6"><label>Property / Propiedad</label><input id="vendorPropertyField" type="text" /></div>
          <div class="field span-3"><label>Prefix</label><input id="vendorPrefixField" type="text" maxlength="6" /></div>
          <div class="field span-3"><label>Contact</label><input id="vendorContactField" type="text" /></div>
          <div class="field span-3"><label>Email</label><input id="vendorEmailField" type="email" /></div>
          <div class="field span-3"><label>Address</label><input id="vendorAddressField" type="text" /></div>
          <div class="field span-12"><label>Notes</label><textarea id="vendorNotesField"></textarea></div>
        </form>
        <div class="actions-row" style="margin-top:16px;">
          <button class="primary-btn" data-action="save-vendor">Save vendor</button>
          <button class="ghost-btn" data-action="close-vendor-modal">Cancel</button>
        </div>
      </div>
    </div>
  `;
}

function renderAppraisalModal() {
  return `
    <div id="appraisalModalBackdrop" class="modal-backdrop">
      <div class="modal">
        <div class="modal-header">
          <h3>New Appraisal / Nuevo Presupuesto</h3>
          <button class="close-btn" data-action="close-appraisal-modal">✕</button>
        </div>
        <form id="appraisalForm" class="form-grid">
          <div class="field span-6"><label>Vendor</label><select id="appraisalVendorId">${renderVendorOptions()}</select></div>
          <div class="field span-6"><label>Property</label><input id="appraisalProperty" type="text" /></div>
          <div class="field span-12"><label>Notes</label><input id="appraisalNotes" type="text" placeholder="Optional note" /></div>
        </form>
        <div id="appraisalItemsWrap" class="service-list" style="margin-top:14px;"></div>
        <div class="actions-row" style="margin-top:14px;">
          <button class="ghost-btn" data-action="add-appraisal-line">+ Add line</button>
          <button class="primary-btn" data-action="save-appraisal">Save appraisal</button>
          <button class="ghost-btn" data-action="close-appraisal-modal">Cancel</button>
        </div>
      </div>
    </div>
  `;
}

function bindEvents() {
  appRoot.querySelectorAll('[data-action]').forEach(el => el.addEventListener('click', handleAction));
  appRoot.querySelectorAll('[data-service-field]').forEach(el => el.addEventListener('change', handleServiceFieldChange));
  const vendorSelect = document.getElementById('invoiceVendorId');
  if (vendorSelect) vendorSelect.addEventListener('change', handleInvoiceVendorChange);
  const prefix = document.getElementById('invoicePrefix');
  if (prefix) prefix.addEventListener('change', handlePrefixChange);
  const promoInput = document.getElementById('settingsPromoImage');
  if (promoInput) promoInput.addEventListener('change', handlePromoUpload);
}

function handleAction(event) {
  const button = event.currentTarget;
  const action = button.dataset.action;
  if (!action) return;
  switch (action) {
    case 'switch-tab': currentTab = button.dataset.tab; renderApp(); break;
    case 'year-prev': selectedYear -= 1; renderApp(); break;
    case 'year-next': selectedYear += 1; renderApp(); break;
    case 'select-month': selectedMonth = button.dataset.month; renderApp(); break;
    case 'open-vendor-modal': openVendorModal(); break;
    case 'close-vendor-modal': closeModal('vendorModalBackdrop'); break;
    case 'save-vendor': saveVendorFromModal(); break;
    case 'edit-vendor': openVendorModal(button.dataset.id); break;
    case 'delete-vendor': deleteVendor(button.dataset.id); break;
    case 'open-appraisal-modal': openAppraisalModal(); break;
    case 'close-appraisal-modal': closeModal('appraisalModalBackdrop'); break;
    case 'add-appraisal-line': addAppraisalLineRow(); break;
    case 'save-appraisal': saveAppraisalFromModal(); break;
    case 'delete-appraisal': deleteAppraisal(button.dataset.id); break;
    case 'convert-appraisal': convertAppraisal(button.dataset.id); break;
    case 'add-service': addService(); break;
    case 'delete-service': deleteService(button.dataset.id); break;
    case 'open-create-for-vendor': openCreateForVendor(button.dataset.vendorId); break;
    case 'add-line-row': addLineItemRow(); break;
    case 'clear-line-rows': clearInvoiceRows(); break;
    case 'quick-add-service': quickAddService(button.dataset.id); break;
    case 'save-invoice': saveInvoice({ download: false }); break;
    case 'save-download-invoice': saveInvoice({ download: true }); break;
    case 'preview-current-invoice': previewCurrentInvoice(); break;
    case 'download-invoice': downloadInvoiceById(button.dataset.id); break;
    case 'delete-invoice': deleteInvoice(button.dataset.id); break;
    case 'toggle-paid': togglePaid(button.dataset.id); break;
    case 'download-followup-pack': downloadFollowupPack(button.dataset.vendorId); break;
    case 'copy-followup-text': copyFollowupText(button.dataset.vendorId); break;
    case 'sync-now': syncToRemote(); break;
    case 'load-remote': loadFromRemote(); break;
    case 'toggle-autosync': state.settings.autoSync = !state.settings.autoSync; saveState(); renderApp(); toast(state.settings.autoSync ? 'Auto-sync enabled.' : 'Auto-sync disabled.'); break;
    case 'save-settings': saveSettings(); break;
    case 'set-theme': state.settings.themeMode = button.dataset.theme; saveState(); renderApp(); break;
    case 'extract-lines': extractLinesFromImage(); break;
  }
}

function handleServiceFieldChange(event) {
  const { serviceField, id } = event.target.dataset;
  const service = state.services.find(item => item.id === id);
  if (!service) return;
  service[serviceField] = serviceField === 'price' ? parseMoney(event.target.value) : event.target.value;
  saveState({ silent: true });
}

function handleInvoiceVendorChange() {
  const select = document.getElementById('invoiceVendorId');
  const vendor = findVendor(select.value);
  if (!vendor) return;
  document.getElementById('invoiceProperty').value = vendor.property || '';
  const prefixInput = document.getElementById('invoicePrefix');
  prefixInput.value = vendor.prefix || derivePrefix(vendor);
  document.getElementById('invoiceNumber').value = nextInvoiceNumber(prefixInput.value);
}

function handlePrefixChange() {
  const prefix = document.getElementById('invoicePrefix').value.trim().toUpperCase();
  if (!document.getElementById('invoiceNumber').value.trim()) {
    document.getElementById('invoiceNumber').value = nextInvoiceNumber(prefix);
  }
}

function openVendorModal(id = '') {
  const modal = document.getElementById('vendorModalBackdrop');
  modal.classList.add('open');
  const vendor = state.vendors.find(v => v.id === id);
  document.getElementById('vendorModalTitle').textContent = vendor ? 'Edit Vendor / Client' : 'Vendor / Client';
  document.getElementById('vendorIdField').value = vendor?.id || '';
  document.getElementById('vendorNameField').value = vendor?.name || '';
  document.getElementById('vendorPropertyField').value = vendor?.property || '';
  document.getElementById('vendorPrefixField').value = vendor?.prefix || '';
  document.getElementById('vendorContactField').value = vendor?.contact || '';
  document.getElementById('vendorEmailField').value = vendor?.email || '';
  document.getElementById('vendorAddressField').value = vendor?.address || '';
  document.getElementById('vendorNotesField').value = vendor?.notes || '';
}

function saveVendorFromModal() {
  const id = document.getElementById('vendorIdField').value;
  const vendor = {
    id: id || uid('vendor'),
    name: document.getElementById('vendorNameField').value.trim(),
    property: document.getElementById('vendorPropertyField').value.trim(),
    prefix: document.getElementById('vendorPrefixField').value.trim().toUpperCase(),
    contact: document.getElementById('vendorContactField').value.trim(),
    email: document.getElementById('vendorEmailField').value.trim(),
    address: document.getElementById('vendorAddressField').value.trim(),
    notes: document.getElementById('vendorNotesField').value.trim(),
  };
  if (!vendor.name) { toast('Vendor name is required.'); return; }
  if (!vendor.prefix) vendor.prefix = derivePrefix(vendor);
  const index = state.vendors.findIndex(item => item.id === vendor.id);
  if (index >= 0) state.vendors[index] = vendor; else state.vendors.push(vendor);
  saveState();
  closeModal('vendorModalBackdrop');
  renderApp();
  toast('Vendor saved.');
}

function deleteVendor(id) {
  if (!confirm('Delete this vendor?')) return;
  state.vendors = state.vendors.filter(v => v.id !== id);
  delete state.invoices[`${id}-${selectedYear}`];
  state.appraisals = state.appraisals.filter(ap => ap.vendorId !== id);
  saveState();
  renderApp();
  toast('Vendor deleted.');
}

function openAppraisalModal() {
  const modal = document.getElementById('appraisalModalBackdrop');
  modal.classList.add('open');
  document.getElementById('appraisalVendorId').value = state.vendors[0]?.id || '';
  document.getElementById('appraisalProperty').value = state.vendors[0]?.property || '';
  document.getElementById('appraisalNotes').value = '';
  const wrap = document.getElementById('appraisalItemsWrap');
  wrap.innerHTML = '';
  addAppraisalLineRow();
}

function addAppraisalLineRow(item = { name: '', amount: '' }) {
  const wrap = document.getElementById('appraisalItemsWrap');
  if (!wrap) return;
  const row = document.createElement('div');
  row.className = 'service-row';
  row.innerHTML = `
    <input type="text" class="appraisal-name" placeholder="Service line" value="${escapeHtml(item.name || '')}" />
    <input type="number" step="0.01" class="appraisal-amount" placeholder="0.00" value="${item.amount || ''}" />
    <button class="ghost-btn" type="button">✕</button>
  `;
  row.querySelector('button').addEventListener('click', () => row.remove());
  wrap.appendChild(row);
}

function saveAppraisalFromModal() {
  const vendorId = document.getElementById('appraisalVendorId').value;
  const property = document.getElementById('appraisalProperty').value.trim();
  const notes = document.getElementById('appraisalNotes').value.trim();
  const items = [...document.querySelectorAll('#appraisalItemsWrap .service-row')].map(row => ({
    name: row.querySelector('.appraisal-name').value.trim(),
    amount: parseMoney(row.querySelector('.appraisal-amount').value)
  })).filter(item => item.name && item.amount);
  if (!vendorId || !items.length) { toast('Add a vendor and at least one service line.'); return; }
  state.appraisals.unshift({ id: uid('appraisal'), vendorId, property, notes, status: 'Draft', createdDate: todayString(), items });
  saveState();
  closeModal('appraisalModalBackdrop');
  renderApp();
  toast('Appraisal saved.');
}

function deleteAppraisal(id) {
  if (!confirm('Delete this appraisal?')) return;
  state.appraisals = state.appraisals.filter(ap => ap.id !== id);
  saveState();
  renderApp();
}

function convertAppraisal(id) {
  const appraisal = state.appraisals.find(ap => ap.id === id);
  if (!appraisal) return;
  const vendor = findVendor(appraisal.vendorId);
  draftSeed = {
    vendorId: appraisal.vendorId,
    property: appraisal.property || vendor?.property || '',
    prefix: vendor?.prefix || derivePrefix(vendor),
    description: 'Appraisal For Services',
    notes: appraisal.notes || '',
    items: appraisal.items,
  };
  appraisal.status = 'Converted';
  saveState();
  currentTab = 'createInvoice';
  renderApp();
  toast('Appraisal loaded into invoice form.');
}

function addService() {
  const name = document.getElementById('newServiceName').value.trim();
  const price = parseMoney(document.getElementById('newServicePrice').value);
  if (!name) { toast('Service name is required.'); return; }
  state.services.push({ id: uid('service'), name, price });
  saveState();
  renderApp();
  toast('Service added.');
}

function deleteService(id) {
  state.services = state.services.filter(service => service.id !== id);
  saveState();
  renderApp();
}

function openCreateForVendor(vendorId) {
  const vendor = findVendor(vendorId);
  draftSeed = {
    vendorId,
    property: vendor?.property || '',
    prefix: vendor?.prefix || derivePrefix(vendor),
    description: 'Appraisal For Services',
    notes: '',
    items: [],
  };
  currentTab = 'createInvoice';
  renderApp();
}

function hydrateInvoiceForm() {
  const select = document.getElementById('invoiceVendorId');
  if (!select) return;
  const vendor = draftSeed?.vendorId ? findVendor(draftSeed.vendorId) : state.vendors[0];
  if (vendor) select.value = vendor.id;
  document.getElementById('invoiceProperty').value = draftSeed?.property || vendor?.property || '';
  document.getElementById('invoicePrefix').value = draftSeed?.prefix || vendor?.prefix || derivePrefix(vendor);
  document.getElementById('invoiceNumber').value = draftSeed?.invoiceNumber || nextInvoiceNumber(document.getElementById('invoicePrefix').value);
  document.getElementById('invoiceDescription').value = draftSeed?.description || 'Appraisal For Services';
  document.getElementById('invoiceDateSent').value = draftSeed?.dateSent || todayString();
  document.getElementById('invoiceDueDate').value = draftSeed?.dueDate || addDays(todayString(), state.settings.defaultDueDays);
  document.getElementById('invoiceNotes').value = draftSeed?.notes || '';
  const wrap = document.getElementById('lineItemsWrap');
  wrap.innerHTML = '';
  (draftSeed?.items?.length ? draftSeed.items : [{ name: '', amount: '' }]).forEach(item => addLineItemRow(item));
  draftSeed = null;
}

function addLineItemRow(item = { name: '', amount: '' }) {
  const wrap = document.getElementById('lineItemsWrap');
  if (!wrap) return;
  const row = document.createElement('div');
  row.className = 'service-row';
  row.innerHTML = `
    <input type="text" class="line-name" placeholder="Service line" value="${escapeHtml(item.name || '')}" />
    <input type="number" step="0.01" class="line-amount" placeholder="0.00" value="${item.amount || ''}" />
    <button class="ghost-btn" type="button">✕</button>
  `;
  row.querySelector('button').addEventListener('click', () => row.remove());
  wrap.appendChild(row);
}

function clearInvoiceRows() {
  const wrap = document.getElementById('lineItemsWrap');
  if (!wrap) return;
  wrap.innerHTML = '';
  addLineItemRow();
}

function quickAddService(id) {
  const service = state.services.find(item => item.id === id);
  if (!service) return;
  addLineItemRow({ name: service.name, amount: service.price });
}

function collectInvoiceForm() {
  const vendorId = document.getElementById('invoiceVendorId').value;
  const vendor = findVendor(vendorId);
  const prefix = (document.getElementById('invoicePrefix').value.trim().toUpperCase() || vendor?.prefix || derivePrefix(vendor));
  const items = [...document.querySelectorAll('#lineItemsWrap .service-row')].map(row => ({
    name: row.querySelector('.line-name').value.trim(),
    amount: parseMoney(row.querySelector('.line-amount').value)
  })).filter(item => item.name && item.amount);
  const amount = invoiceItemsTotal(items);
  return {
    id: uid('invoice'),
    invoiceNumber: document.getElementById('invoiceNumber').value.trim() || nextInvoiceNumber(prefix),
    prefix,
    vendorId,
    vendorName: vendor?.name || '',
    month: selectedMonth,
    year: selectedYear,
    property: document.getElementById('invoiceProperty').value.trim() || vendor?.property || '',
    dateSent: document.getElementById('invoiceDateSent').value,
    dueDate: document.getElementById('invoiceDueDate').value,
    description: document.getElementById('invoiceDescription').value.trim() || 'Invoice',
    notes: document.getElementById('invoiceNotes').value.trim(),
    billingHeader: (document.getElementById('invoiceBillingHeader')?.value || '').trim(),
    paid: false,
    paidDate: '',
    amount,
    items,
  };
}

function createInvoiceSnapshot(invoice) {
  const vendor = findVendor(invoice.vendorId) || { name: invoice.vendorName || '', property: invoice.property || '' };
  return {
    invoiceNumber: invoice.invoiceNumber || '',
    prefix: invoice.prefix || '',
    vendorName: vendor.name || '',
    property: invoice.property || vendor.property || '',
    description: invoice.description || '',
    dateSent: invoice.dateSent || '',
    dueDate: invoice.dueDate || '',
    billingHeader: invoice.billingHeader || '',
    items: (invoice.items || []).map(item => ({ name: item.name, amount: Number(item.amount || 0) })),
    amount: Number(invoice.amount || 0),
    ownerName: state.settings.ownerName || '',
    companyAddress: state.settings.companyAddress || '',
    companyEmail: state.settings.companyEmail || '',
    companyName: state.settings.companyName || '',
    phone: state.settings.phone || '',
    paymentNote: state.settings.paymentNote || '',
    taxRate: Number(state.settings.taxRate || 0),
    promoImage: state.settings.promoImage || '/assets/promo-default-v7.png',
    headerLogo: state.settings.headerLogo || '/assets/header-logo-v7.png',
    capturedAt: new Date().toISOString()
  };
}

function formatPdfDate(value) {
  if (!value) return '';
  const simple = String(value).split('T')[0];
  if (/^\d{4}-\d{2}-\d{2}$/.test(simple)) {
    const [y, m, d] = simple.split('-').map(Number);
    return `${m}/${d}/${y}`;
  }
  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) return `${parsed.getMonth() + 1}/${parsed.getDate()}/${parsed.getFullYear()}`;
  return String(value);
}

async function saveInvoice({ download }) {
  const invoice = collectInvoiceForm();
  if (!invoice.vendorId || !invoice.items.length) {
    toast('Choose a vendor and add at least one service line.');
    return;
  }
  invoice.snapshot = createInvoiceSnapshot(invoice);
  const key = buildInvoiceKey(invoice.vendorId, invoice.year);
  state.invoices[key] = state.invoices[key] || [];
  state.invoices[key].unshift(invoice);
  saveState();
  try {
    if (download) await generateInvoicePdf(invoice, { download: true });
    renderApp();
    toast(download ? 'Invoice saved and PDF downloaded.' : 'Invoice saved.');
  } catch (error) {
    console.error(error);
    renderApp();
    toast(`Save worked, but PDF failed: ${error.message}`);
  }
}

async function previewCurrentInvoice() {
  const invoice = collectInvoiceForm();
  if (!invoice.vendorId || !invoice.items.length) {
    toast('Choose a vendor and add at least one service line.');
    return;
  }
  try {
    await generateInvoicePdf(invoice, { download: true });
  } catch (error) {
    console.error(error);
    toast(`PDF error: ${error.message}`);
  }
}

function findInvoiceById(id) {
  return getAllInvoices().find(inv => inv.id === id) || null;
}

function deleteInvoice(id) {
  if (!confirm('Delete this invoice?')) return;
  Object.keys(state.invoices).forEach(key => {
    state.invoices[key] = state.invoices[key].filter(inv => inv.id !== id);
    if (!state.invoices[key].length) delete state.invoices[key];
  });
  saveState();
  renderApp();
}

function togglePaid(id) {
  const invoice = findInvoiceById(id);
  if (!invoice) return;
  invoice.paid = !invoice.paid;
  invoice.paidDate = invoice.paid ? todayString() : '';
  saveState();
  renderApp();
  toast(invoice.paid ? 'Marked paid.' : 'Marked unpaid.');
}

async function downloadInvoiceById(id) {
  const invoice = findInvoiceById(id);
  if (!invoice) return;
  try {
    await generateInvoicePdf(invoice, { download: true });
  } catch (error) {
    console.error(error);
    toast(`PDF error: ${error.message}`);
  }
}

async function generateInvoicePdf(invoice, { download = true, doc = null, addPage = false } = {}) {
  const vendor = findVendor(invoice.vendorId) || { name: invoice.vendorName || '', property: invoice.property || '' };
  const snap = invoice.snapshot || createInvoiceSnapshot(invoice);
  const localDoc = doc || new jsPDF({ unit: 'pt', format: 'letter' });
  if (addPage) localDoc.addPage();

  const pageWidth = localDoc.internal.pageSize.getWidth();
  const pageHeight = localDoc.internal.pageSize.getHeight();
  const margin = 24;
  const contentWidth = pageWidth - margin * 2;

  const headerLogoPath = snap.headerLogo || state.settings.headerLogo || '/assets/header-logo-v7.png';
  const headerLogo = headerLogoPath?.startsWith('data:') ? headerLogoPath : await loadImageAsDataUrl(headerLogoPath);
  const promoPath = snap.promoImage || state.settings.promoImage || '/assets/promo-default-v7.png';
  const promo = promoPath?.startsWith('data:') ? promoPath : await loadImageAsDataUrl(promoPath);

  const headerOwner = snap.ownerName || state.settings.ownerName || 'Roman Cuateco';
  const headerAddress = snap.companyAddress || state.settings.companyAddress || '516 Tarrytown Road, White Plains, NY 10607';
  const headerEmail = snap.companyEmail || state.settings.companyEmail || 'RomanD.Landscaping@gmail.com';
  const paymentNote = snap.paymentNote || state.settings.paymentNote || '';
  const taxRate = Number(snap.taxRate ?? state.settings.taxRate ?? 0);
  const descriptionText = `${snap.description || invoice.description || ''} ${snap.property || invoice.property || vendor.property || ''}`.trim();
  const displayDate = formatPdfDate(snap.dateSent || invoice.dateSent || '');
  const bodyItems = (snap.items && snap.items.length ? snap.items : (invoice.items?.length ? invoice.items : [{ name: invoice.description || 'Invoice total', amount: invoice.amount }]))
    .map(item => ({ name: item.name || '', amount: Number(item.amount || 0) }));

  localDoc.setFillColor(241, 241, 241);
  localDoc.rect(0, 0, pageWidth, pageHeight, 'F');

  let y = 10;
  if (headerLogo) {
    const logoProps = localDoc.getImageProperties(headerLogo);
    const logoRatio = logoProps.width / logoProps.height;
    const logoWidth = 118;
    const logoHeight = logoWidth / logoRatio;
    localDoc.addImage(headerLogo, inferImageFormat(headerLogo), pageWidth / 2 - logoWidth / 2, y, logoWidth, logoHeight, undefined, 'FAST');
    y += logoHeight + 8;
  } else {
    y += 110;
  }

  const colWidths = [contentWidth * 0.28, contentWidth * 0.44, contentWidth * 0.28];
  const rowHeight = 16;
  let x = margin;
  localDoc.setDrawColor(40, 40, 40);
  localDoc.setLineWidth(1);
  [headerOwner, headerAddress, headerEmail].forEach((cell, index) => {
    const width = colWidths[index];
    localDoc.rect(x, y, width, rowHeight);
    localDoc.setTextColor(0, 0, 0);
    localDoc.setFont('helvetica', 'bold');
    localDoc.setFontSize(10.5);
    localDoc.text(String(cell || ''), x + 8, y + 12);
    x += width;
  });
  y += rowHeight + 14;

  const billingHeader = (snap.billingHeader || invoice.billingHeader || '').trim();
  if (billingHeader) {
    localDoc.setTextColor(0, 104, 84);
    localDoc.setFont('times', 'bold');
    localDoc.setFontSize(12.5);
    const lines = localDoc.splitTextToSize(billingHeader, contentWidth - 70);
    lines.forEach((line, idx) => {
      localDoc.text(line, pageWidth / 2, y + idx * 15, { align: 'center' });
    });
    y += lines.length * 15 + 6;
  }

  localDoc.setTextColor(0, 104, 84);
  localDoc.setFont('helvetica', 'bold');
  localDoc.setFontSize(11.5);
  localDoc.text(`DATE: ${displayDate}`, margin + 145, y);
  const maxDescWidth = pageWidth - (margin + 290) - margin;
  const desc = localDoc.splitTextToSize(`DESCRIPTION: ${descriptionText}`, maxDescWidth)[0] || 'DESCRIPTION:';
  localDoc.text(desc, margin + 290, y);
  y += 10;
  localDoc.setDrawColor(48, 48, 48);
  localDoc.line(margin + 58, y, pageWidth - margin - 58, y);
  y += 18;

  localDoc.setTextColor(0, 0, 0);
  localDoc.setFont('helvetica', 'normal');
  localDoc.setFontSize(12);
  const numX = margin + 20;
  const textX = margin + 44;
  const amountX = pageWidth - margin - 62;
  const textWidth = amountX - textX - 88;

  for (let i = 0; i < bodyItems.length; i++) {
    const item = bodyItems[i];
    const lines = localDoc.splitTextToSize(item.name || '', textWidth);
    const rowH = Math.max(18, lines.length * 14);
    const neededSpace = rowH + 70 + 220;
    if (y + neededSpace > pageHeight) {
      localDoc.addPage();
      localDoc.setFillColor(241, 241, 241);
      localDoc.rect(0, 0, pageWidth, pageHeight, 'F');
      y = 40;
    }
    localDoc.setFont('helvetica', 'bold');
    localDoc.text(`${i + 1}.`, numX, y + 12);
    localDoc.setFont('helvetica', 'normal');
    localDoc.text(lines, textX, y + 12);
    localDoc.text(money(item.amount), amountX, y + 12, { align: 'right' });
    y += rowH + 8;
  }

  const subtotal = invoiceItemsTotal(bodyItems);
  const tax = subtotal * (taxRate / 100);
  const total = subtotal + tax;
  y += 2;
  localDoc.line(margin + 58, y, pageWidth - margin - 58, y);
  y += 18;

  localDoc.setFontSize(10.5);
  const totalsLabelX = pageWidth - margin - 190;
  const totalsValueX = amountX;
  const totalLineGap = 14;
  localDoc.setFont('helvetica', 'normal');
  localDoc.text('Subtotal:', totalsLabelX, y);
  localDoc.text(money(subtotal), totalsValueX, y, { align: 'right' });
  y += totalLineGap;
  localDoc.text('Tax:', totalsLabelX, y);
  localDoc.text(money(tax), totalsValueX, y, { align: 'right' });
  y += totalLineGap;
  localDoc.setFont('helvetica', 'bold');
  localDoc.text('Total:', totalsLabelX, y);
  localDoc.text(money(total), totalsValueX, y, { align: 'right' });
  y += 18;

  if (paymentNote) {
    localDoc.setTextColor(210, 20, 20);
    localDoc.setFont('helvetica', 'bold');
    localDoc.setFontSize(10);
    localDoc.text(`!!! ${paymentNote} !!!`, pageWidth / 2, y, { align: 'center' });
    localDoc.setTextColor(0, 0, 0);
    y += 14;
  }

  if (promo) {
    const promoProps = localDoc.getImageProperties(promo);
    const promoRatio = promoProps.width / promoProps.height;
    let promoWidth = Math.min(contentWidth - 100, 458);
    let promoHeight = promoWidth / promoRatio;
    const maxPromoHeight = pageHeight - y - margin;
    if (promoHeight > maxPromoHeight) {
      promoHeight = maxPromoHeight;
      promoWidth = promoHeight * promoRatio;
    }
    if (promoHeight < 150) {
      localDoc.addPage();
      localDoc.setFillColor(241, 241, 241);
      localDoc.rect(0, 0, pageWidth, pageHeight, 'F');
      y = 40;
      promoWidth = Math.min(contentWidth - 100, 458);
      promoHeight = promoWidth / promoRatio;
    }
    localDoc.addImage(promo, inferImageFormat(promo), pageWidth / 2 - promoWidth / 2, y + 4, promoWidth, promoHeight, undefined, 'FAST');
  }

  if (download) localDoc.save(`${sanitizeFileName((invoice.invoiceNumber || 'invoice') + '-' + (vendor.name || 'client'))}.pdf`);
  return localDoc;
}

async function downloadFollowupPack(vendorId) {
  const vendor = findVendor(vendorId);
  if (!vendor) return;
  const invoices = getAllInvoices().filter(inv => inv.vendorId === vendorId && !inv.paid);
  if (!invoices.length) { toast('No pending invoices for this vendor.'); return; }
  const bundle = new jsPDF({ unit: 'pt', format: 'letter' });
  let first = true;
  for (const inv of invoices) {
    await generateInvoicePdf(inv, { download: false, doc: bundle, addPage: !first });
    first = false;
  }
  bundle.save(`${sanitizeFileName((vendor.name || 'vendor') + '-pending-pack')}.pdf`);
  toast(`Downloaded ${invoices.length} pending invoice(s) for ${vendor.name}.`);
}

async function copyFollowupText(vendorId) {
  const vendor = findVendor(vendorId);
  const invoices = getAllInvoices().filter(inv => inv.vendorId === vendorId && !inv.paid);
  if (!vendor || !invoices.length) return;
  const lines = invoices.map(inv => `• ${inv.invoiceNumber || 'Invoice'} — ${money(inv.amount)} due ${inv.dueDate || '—'}`);
  const text = `Hello ${vendor.contact || vendor.name},\n\nThis is a friendly follow-up regarding the following open invoices:\n${lines.join('\n')}\n\nPlease let us know if payment has already been sent.\n\nThank you,\n${state.settings.companyName}`;
  await navigator.clipboard.writeText(text);
  toast('Follow-up text copied.');
}

function saveSettings() {
  state.settings.syncUrl = document.getElementById('settingsSyncUrl').value.trim();
  state.settings.ownerName = document.getElementById('settingsOwnerName').value.trim();
  state.settings.companyName = document.getElementById('settingsCompanyName').value.trim();
  state.settings.phone = document.getElementById('settingsPhone').value.trim();
  state.settings.companyEmail = document.getElementById('settingsEmail').value.trim();
  state.settings.companyAddress = document.getElementById('settingsAddress').value.trim();
  state.settings.paymentNote = document.getElementById('settingsPaymentNote').value.trim();
  state.settings.taxRate = parseMoney(document.getElementById('settingsTaxRate').value);
  state.settings.defaultDueDays = Number(document.getElementById('settingsDueDays').value || 14);
  state.settings.anthropicApiKey = document.getElementById('settingsAnthropicApiKey').value.trim();
  state.settings.anthropicModel = document.getElementById('settingsAnthropicModel').value.trim() || 'claude-sonnet-4-6';
  saveState();
  renderApp();
  toast('Settings saved.');
}

async function handlePromoUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  state.settings.promoImage = await blobToDataUrl(file);
  saveState({ silent: true });
  const preview = document.getElementById('promoPreview');
  if (preview) preview.src = state.settings.promoImage;
  toast('Promo image updated.');
}

async function extractLinesFromImage() {
  const input = document.getElementById('ocrImageInput');
  const file = input?.files?.[0];
  if (!file) { toast('Choose an image first.'); return; }
  const status = document.getElementById('ocrStatus');
  status.textContent = 'Extracting…';
  try {
    const base64 = (await blobToDataUrl(file)).split(',')[1];
    const payload = {
      apiKey: state.settings.anthropicApiKey || '',
      model: state.settings.anthropicModel || 'claude-sonnet-4-6',
      mediaType: file.type,
      imageBase64: base64,
    };
    const response = await fetch('/anthropic-ocr', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.error || 'OCR failed');
    applyExtractedData(result.data || {});
    status.textContent = 'Done';
    toast('Image converted into invoice lines.');
  } catch (error) {
    console.error(error);
    status.textContent = 'Error';
    toast(`OCR failed: ${error.message}`);
  }
}

function applyExtractedData(data) {
  const description = document.getElementById('invoiceDescription');
  const dateSent = document.getElementById('invoiceDateSent');
  const dueDate = document.getElementById('invoiceDueDate');
  const property = document.getElementById('invoiceProperty');
  const invoiceNumber = document.getElementById('invoiceNumber');
  if (data.description && description) description.value = data.description;
  if (data.dateSent && dateSent) dateSent.value = normalizePossibleDate(data.dateSent);
  if (data.dueDate && dueDate) dueDate.value = normalizePossibleDate(data.dueDate);
  if (data.property && property) property.value = data.property;
  if (data.invoiceNumber && invoiceNumber) invoiceNumber.value = data.invoiceNumber;
  if (Array.isArray(data.items) && data.items.length) {
    clearInvoiceRows();
    const wrap = document.getElementById('lineItemsWrap');
    wrap.innerHTML = '';
    data.items.forEach(item => addLineItemRow({ name: item.name, amount: parseMoney(item.amount) }));
  }
}

function normalizePossibleDate(value) {
  if (!value) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? '' : toInputDate(parsed);
}

async function syncToRemote({ silent = false } = {}) {
  if (!state.settings.syncUrl) {
    if (!silent) toast('Add the Google Apps Script Web App URL in Settings first.');
    return;
  }
  try {
    const payload = serializeForRemote();
    const response = await fetch(state.settings.syncUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload)
    });
    const result = await response.json();
    if (!result.success) throw new Error(result.error || 'Sync failed');
    if (!silent) toast('Synced to Google Sheets.');
  } catch (error) {
    console.error(error);
    if (!silent) toast(`Sync error: ${error.message}`);
  }
}

async function loadFromRemote({ silent = false } = {}) {
  if (!state.settings.syncUrl) return;
  try {
    const response = await fetch(state.settings.syncUrl);
    const result = await response.json();
    if (!result.success) throw new Error(result.error || 'Load failed');
    hydrateFromRemote(result);
    saveState({ silent: true });
    renderApp();
    if (!silent) toast('Loaded latest data from Google Sheets.');
  } catch (error) {
    console.error(error);
    if (!silent) toast(`Load error: ${error.message}`);
  }
}

function serializeForRemote() {
  return {
    vendors: state.vendors,
    invoices: state.invoices,
    appraisals: state.appraisals,
    services: state.services,
    settings: state.settings,
  };
}

function hydrateFromRemote(remote) {
  if (Array.isArray(remote.vendors)) state.vendors = remote.vendors.map(v => ({ ...v, prefix: v.prefix || derivePrefix(v) }));
  if (remote.invoices && typeof remote.invoices === 'object') {
    state.invoices = remote.invoices;
  }
  if (Array.isArray(remote.appraisals)) state.appraisals = remote.appraisals;
  if (Array.isArray(remote.services)) state.services = remote.services;
  if (remote.settings && typeof remote.settings === 'object') state.settings = deepMerge(state.settings, remote.settings);
  normalizeState(state);
}

function closeModal(id) {
  document.getElementById(id)?.classList.remove('open');
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function sanitizeFileName(text) {
  return String(text || 'file').replace(/[^a-z0-9-_]+/gi, '-').replace(/-+/g, '-').replace(/^-|-$/g, '');
}

function toast(message) {
  const node = document.getElementById('toast');
  if (!node) return;
  node.textContent = message;
  node.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => node.classList.remove('show'), 2600);
}
