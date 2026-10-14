import { GraphStore } from './store-graph.js?v=2026101402';
import { DemoStore } from './store-demo.js?v=2026101402';
import { compressImage, readReceipt } from './receipt.js?v=2026101402';

const SCHEMA = window.BOOKS_SCHEMA;
const DEMO = new URLSearchParams(location.search).has('demo');
const TABLES = ['Projects', 'Clients', 'Workers', 'Categories', 'Tasks', 'Expenses', 'Mileage', 'Invoices', 'Disbursements'];
const OPTIONAL_TABLES = ['Status', 'Table12']; // drop-down lists; the app has defaults if they're missing
const loadEverything = () => store.loadAll(TABLES).then(() => Promise.allSettled(OPTIONAL_TABLES.map(t => store.loadTable(t))));

// ---------- small utilities ----------
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const num = v => Number(v) || 0;
const EPOCH = Date.UTC(1899, 11, 30);
const toSerial = iso => iso ? Math.round((Date.parse(iso + 'T00:00:00Z') - EPOCH) / 864e5) : null;
const toISO = s => (typeof s === 'number' && s > 0) ? new Date(EPOCH + s * 864e5).toISOString().slice(0, 10) : (typeof s === 'string' ? s : '');
const todayISO = () => { const d = new Date(); return new Date(d - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10); };
const longDate = iso => iso ? new Date(iso + 'T12:00:00').toLocaleDateString('en-CA', { year: 'numeric', month: 'long', day: 'numeric' }) : '';
const money = v => num(v).toLocaleString('en-CA', { style: 'currency', currency: 'CAD' });
const hrs = v => num(v).toLocaleString('en-CA', { maximumFractionDigits: 2 });
const ls = {
  get(k, d) { try { return JSON.parse(localStorage.getItem('books:' + k)) ?? d; } catch { return d; } },
  set(k, v) { try { localStorage.setItem('books:' + k, JSON.stringify(v)); } catch {} },
};
const rows = name => (store.tables[name]?.rows || []).filter(r => r._raw.some(v => v !== '' && v !== null));
const clean = s => String(s || '').trim();

// ---------- app state ----------
const CFG = { ...window.BOOKS_CONFIG, ...ls.get('settings', {}) };
const store = DEMO ? new DemoStore(CFG, SCHEMA) : new GraphStore(CFG, SCHEMA);
if (DEMO) window.demoStore = store; // for poking at sample data in the browser console
const cfg = () => store.cfg; // demo store may override company details
let user = null;
let me = null; // matching Workers row name

// ---------- feedback ----------
function toast(msg, kind = 'ok', ms = 4500) {
  const t = document.createElement('div');
  t.className = 'toast ' + kind;
  t.innerHTML = msg;
  $('#toasts').appendChild(t);
  setTimeout(() => t.remove(), ms);
}
function logActivity(action, address, detail = '') {
  const log = ls.get('log', []);
  log.unshift({ when: new Date().toISOString(), action, address, detail });
  ls.set('log', log.slice(0, 100));
}
async function busy(label, fn) {
  const b = $('#busy');
  b.querySelector('span').textContent = label;
  b.hidden = false;
  try { return await fn(); }
  catch (e) { console.error(e); toast('⚠ ' + esc(e.message || e), 'warn', 9000); throw e; }
  finally { b.hidden = true; }
}

// ---------- boot ----------
async function boot() {
  try {
    user = await store.init();
  } catch (e) {
    return screen(`<h1>Sign-in problem</h1><p>${esc(e.message)}</p><p>Check <code>clientId</code> and <code>tenantId</code> in <code>js/config.js</code>.</p>`);
  }
  if (!user) {
    return screen(`
      <div class="brandmark big">${brandMark()}</div>
      <h1>${esc(CFG.appName)}</h1>
      <p class="muted">Time, expenses, receipts, mileage and invoices, written straight into your bookkeeping workbook.</p>
      <button class="btn primary lg" id="signin">Sign in with Microsoft</button>
      <p class="small muted">Only authorized accounts can open the workbook. <a href="?demo">Try the demo with sample data</a></p>`,
      () => $('#signin').onclick = () => store.login());
  }
  // Ennovantiv accounts only: the tenant is enforced by Microsoft sign-in;
  // the domain check also turns away guest accounts from other organizations.
  const allowed = (cfg().allowedUsers || []).map(s => s.toLowerCase());
  const domains = (cfg().allowedDomains || []).map(s => s.toLowerCase());
  const domainOk = !domains.length || domains.some(d => user.email.endsWith('@' + d));
  if (!DEMO && (!domainOk || (allowed.length && !allowed.includes(user.email)))) {
    return screen(`<h1>Not authorized</h1><p>${esc(user.email)} is not an Ennovantiv account, so it can't use this app.</p><button class="btn" id="so">Sign out</button>`,
      () => $('#so').onclick = () => store.logout());
  }
  renderShell();
  const url = CFG.fileShareUrl;
  if (!DEMO && !url) { location.hash = '#/settings'; toast('Paste the link to your workbook to get started.', 'info'); return route(); }
  await connect();
}

async function connect() {
  try {
    await busy('Opening workbook…', async () => {
      await store.openFile(CFG.fileShareUrl);
      await loadEverything();
    });
  } catch {
    location.hash = '#/settings';
    return route();
  }
  me = rows('Workers').find(w => clean(w.email).toLowerCase() === user.email)?.Name || rows('Workers')[0]?.Name || '';
  $('#filechip').innerHTML = `<span class="dot"></span>${esc(store.file.name)}`;
  $('#filechip').href = store.file.webUrl;
  route();
}

async function refresh() {
  RATES = null;
  await busy('Refreshing from Excel…', loadEverything);
  route();
  toast('Up to date with the workbook.', 'info', 2000);
}

function screen(html, after) {
  document.body.innerHTML = `<main class="center-screen"><div class="card narrow">${html}</div></main>`;
  after?.();
}

function brandMark() {
  return `<img src="${esc(CFG.logoUrl)}" alt="" onerror="this.remove()"><span class="wm">Ennovantiv<small>SERVICES INC</small></span>`;
}

// ---------- shell & routing ----------
const NAV = [
  ['dashboard', 'Dashboard', 'M3 13h8V3H3zm10 8h8V11h-8zM3 21h8v-6H3zm10-18v6h8V3z'],
  ['time', 'Time', 'M12 7v5l3 3M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0z'],
  ['expenses', 'Expenses', 'M4 3h16v18l-3-2-3 2-2-2-2 2-3-2-3 2zM8 8h8M8 12h8M8 16h5'],
  ['mileage', 'Mileage', 'M5 17h14l-2-7H7zM7 17v2M17 17v2M8 13h.01M16 13h.01M7 10l2-4h6l2 4'],
  ['invoices', 'Invoices', 'M6 2h9l5 5v15H6zM14 2v6h6M9 13h8M9 17h8M9 9h3'],
  ['projects', 'Projects', 'M3 7h18v13H3zM8 7V4h8v3M3 12h18'],
  ['settings', 'Settings', 'M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6zM19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z'],
];
const GEAR = NAV.find(n => n[0] === 'settings')[2];
const icon = d => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="${d}"/></svg>`;

function renderShell() {
  document.body.innerHTML = `
    <header class="topbar">
      <a class="brandmark" href="#/dashboard">${brandMark()}</a>
      <div class="topbar-right">
        ${DEMO ? '<span class="pill demo" title="Sample data, nothing is saved">DEMO</span>' : ''}
        <a id="filechip" class="filechip" target="_blank" rel="noopener" title="Open the workbook in Excel"><span class="dot off"></span>Not connected</a>
        <button class="iconbtn" id="refresh" title="Reload from Excel">${icon('M21 12a9 9 0 1 1-3-6.7L21 8M21 3v5h-5')}</button>
        <a class="iconbtn mobile-only" href="#/settings" title="Settings">${icon(GEAR)}</a>
      </div>
    </header>
    <nav class="nav">${NAV.map(([k, l, d]) => `<a href="#/${k}" data-k="${k}" class="${k === 'settings' ? 'desktop-only' : ''}">${icon(d)}<span>${l === 'Dashboard' ? '<i class="desktop-only">Dashboard</i><i class="mobile-only">Home</i>' : l}</span></a>`).join('')}</nav>
    <main id="view" class="view"></main>
    <div id="toasts" class="toasts" aria-live="polite"></div>
    <div id="busy" class="busy" hidden><div class="spinner"></div><span></span></div>
    <div id="print-root"></div>`;
  $('#refresh').onclick = () => store.file && refresh();
  window.onhashchange = route;
}

function route() {
  const [k, sub] = (location.hash.replace(/^#\/?/, '') || 'dashboard').split('/');
  const navKey = { invoice: 'invoices', report: 'expenses' }[k] || k;
  $$('.nav a').forEach(a => a.classList.toggle('active', a.dataset.k === navKey));
  const view = $('#view');
  if (!store.file && k !== 'settings') { view.innerHTML = '<div class="card"><p class="muted">Connecting to the workbook…</p></div>'; return; }
  const views = { dashboard: viewDashboard, time: () => viewEntry(TIME), expenses: () => viewEntry(EXPENSE), mileage: () => viewEntry(MILEAGE), invoices: viewInvoices, invoice: viewInvoiceDetail, projects: viewProjects, report: viewExpenseReport, settings: viewSettings };
  (views[k] || viewDashboard)(view, sub && decodeURIComponent(sub));
  window.scrollTo(0, 0);
}

// ---------- period filter ----------
// Each workbook holds one year, so "This year" shows every row.
const PERIODS = { today: 'Today', week: 'This week', lastweek: 'Last week', d7: 'Last 7 days', d14: 'Last 14 days', month: 'This month', lastmonth: 'Last month', year: 'This year' };
function inPeriod(serial, p) {
  const v = num(serial), iso = toISO(serial), t = todayISO(), s = toSerial(t);
  const dow = (new Date(t + 'T12:00:00').getDay() + 6) % 7; // weeks start Monday
  const between = (a, b) => v >= a && v <= b;
  switch (p) {
    case 'today': return v === s;
    case 'week': return between(s - dow, s - dow + 6);
    case 'lastweek': return between(s - dow - 7, s - dow - 1);
    case 'd7': return between(s - 6, s);
    case 'd14': return between(s - 13, s);
    case 'month': return iso.slice(0, 7) === t.slice(0, 7);
    case 'lastmonth': { const d = new Date(t.slice(0, 8) + '15T12:00:00'); d.setMonth(d.getMonth() - 1); return iso.slice(0, 7) === d.toISOString().slice(0, 7); }
    default: return true;
  }
}

// ---------- billing day ----------
// Ontario statutory holidays, plus the Civic Holiday (first Monday of August)
// which most businesses observe. Weekend holidays move to the next weekday.
function holidays(y) {
  const iso = (m, d) => new Date(Date.UTC(y, m - 1, d)).toISOString().slice(0, 10);
  const dow = (m, d) => new Date(Date.UTC(y, m - 1, d)).getUTCDay();
  const nthMonday = (m, n) => 1 + ((8 - dow(m, 1)) % 7) + (n - 1) * 7;
  const observed = (m, d) => [iso(m, d), dow(m, d) === 6 ? iso(m, d + 2) : dow(m, d) === 0 ? iso(m, d + 1) : null];
  // Easter (Anonymous Gregorian algorithm) for Good Friday
  const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, mm = Math.floor((a + 11 * h + 22 * l) / 451);
  const em = Math.floor((h + l - 7 * mm + 114) / 31), ed = ((h + l - 7 * mm + 114) % 31) + 1;
  const goodFriday = new Date(Date.UTC(y, em - 1, ed - 2)).toISOString().slice(0, 10);
  const dec25 = dow(12, 25);
  return new Set([
    ...observed(1, 1), iso(2, nthMonday(2, 3)), goodFriday,
    iso(5, 24 - ((dow(5, 24) + 6) % 7)),              // Victoria Day
    ...observed(7, 1), iso(8, nthMonday(8, 1)), iso(9, nthMonday(9, 1)), iso(10, nthMonday(10, 2)),
    iso(12, 25), iso(12, 26),
    dec25 === 6 ? iso(12, 27) : null, dec25 === 6 || dec25 === 5 ? iso(12, 28) : null, dec25 === 0 ? iso(12, 27) : null,
    ...(cfg().extraHolidays || []),
  ].filter(Boolean));
}
const holidayNames = { '01': 'New Year’s Day', '02': 'Family Day', '05': 'Victoria Day', '07': 'Canada Day', '08': 'Civic Holiday', '09': 'Labour Day', '10': 'Thanksgiving' };
// First Monday of the month, or the next business day if that is a holiday.
function billingDay(y, m) {
  const off = holidays(y);
  const first = new Date(Date.UTC(y, m - 1, 1)).getUTCDay();
  let d = new Date(Date.UTC(y, m - 1, 1 + ((8 - first) % 7)));
  const monday = d.toISOString().slice(0, 10);
  while (off.has(d.toISOString().slice(0, 10)) || [0, 6].includes(d.getUTCDay())) d.setUTCDate(d.getUTCDate() + 1);
  const iso = d.toISOString().slice(0, 10);
  return { iso, shifted: iso !== monday, reason: iso !== monday ? (holidayNames[monday.slice(5, 7)] || 'holiday') : '' };
}
function nextBillingDay() {
  const t = todayISO(), y = Number(t.slice(0, 4)), m = Number(t.slice(5, 7));
  const b = billingDay(y, m);
  return b.iso >= t ? b : billingDay(m === 12 ? y + 1 : y, m === 12 ? 1 : m + 1);
}
const weekday = iso => new Date(iso + 'T12:00:00').toLocaleDateString('en-CA', { weekday: 'long', month: 'long', day: 'numeric' });

// =====================================================================
// Dashboard
// =====================================================================
function viewDashboard(view) {
  const yr = todayISO().slice(0, 4), mo = todayISO().slice(0, 7);
  const own = cfg().companyClientRow;
  const projects = rows('Projects').filter(p => p.Client !== own);
  const disb = rows('Disbursements').filter(d => d.Invoiced !== true);
  const byClient = {};
  for (const p of projects) if (num(p['This Invoice']) > 0) byClient[p.Client] = (byClient[p.Client] || 0) + num(p['This Invoice']);
  for (const d of disb) byClient[d.Client] = (byClient[d.Client] || 0) + num(d.Billable);
  const unbilled = Object.values(byClient).reduce((a, b) => a + b, 0);
  const open = rows('Invoices').filter(i => !i.Paid);
  const overdue = open.filter(i => toISO(num(i.Date) + cfg().invoiceTermsDays) < todayISO());
  const tasksMo = rows('Tasks').filter(t => toISO(t.Date).startsWith(mo));
  const myHrs = tasksMo.filter(t => t.Worker === me).reduce((s, t) => s + num(t.Hours), 0);
  const allHrs = tasksMo.reduce((s, t) => s + num(t.Hours), 0);
  const expYr = rows('Expenses').filter(e => toISO(e.Date).startsWith(yr));
  const expTot = expYr.reduce((s, e) => s + num(e['Bus Total']), 0);
  const kmYr = rows('Mileage').filter(m => toISO(m.Date).startsWith(yr)).reduce((s, m) => s + num(m.km), 0);
  const log = ls.get('log', []).slice(0, 8);

  view.innerHTML = `
    <div class="page-head"><h1>Hello${me ? ', ' + esc(me.split(' ')[0]) : ''}</h1><p class="muted">${longDate(todayISO())}</p></div>
    <div class="quick">
      <a class="btn primary" href="#/time">+ Time</a>
      <a class="btn primary" href="#/expenses/scan">Scan receipt</a>
      <a class="btn" href="#/expenses">+ Expense</a>
      <a class="btn" href="#/mileage">+ Trip</a>
      <a class="btn" href="#/invoices/new">New invoice</a>
      <a class="btn" href="#/report">Expense report</a>
    </div>
    <div class="stats">
      <div class="stat accent"><label>Ready to invoice</label><b>${money(unbilled)}</b><small>next billing day ${weekday(nextBillingDay().iso)}</small></div>
      <div class="stat"><label>Receivables</label><b>${money(open.reduce((s, i) => s + num(i.Amount), 0))}</b><small>${open.length} open · ${overdue.length} overdue</small></div>
      <div class="stat"><label>Hours this month</label><b>${hrs(myHrs)}</b><small>${hrs(allHrs)} for all workers</small></div>
      <div class="stat"><label>Expenses ${yr}</label><b>${money(expTot)}</b><small>business portion, incl. HST</small></div>
      <div class="stat"><label>Business km ${yr}</label><b>${num(kmYr).toLocaleString('en-CA')}</b><small>${rows('Mileage').filter(m => toISO(m.Date).startsWith(yr)).length} trips</small></div>
    </div>
    <div class="grid2">
      <div class="card">
        <h2>Ready to invoice by client</h2>
        ${Object.keys(byClient).length ? `<table class="list"><tbody>${Object.entries(byClient).sort((a, b) => b[1] - a[1]).map(([c, v]) =>
          `<tr><td>${esc(c)}</td><td class="r">${money(v)}</td><td class="r"><a class="btn sm" href="#/invoices/new" data-client="${esc(c)}">Invoice</a></td></tr>`).join('')}</tbody></table>`
          : '<p class="muted">Nothing waiting to be invoiced.</p>'}
      </div>
      <div class="card">
        <h2>Recent changes made from this device</h2>
        ${log.length ? `<ul class="log">${log.map(l => `<li><b>${esc(l.action)}</b> <code>${esc(l.address)}</code><br><small class="muted">${new Date(l.when).toLocaleString('en-CA')} ${esc(l.detail)}</small></li>`).join('')}</ul>`
          : '<p class="muted">Every row the app writes shows up here with its exact sheet and cell address.</p>'}
      </div>
    </div>`;
  $$('[data-client]', view).forEach(a => a.onclick = () => ls.set('invClient', a.dataset.client));
}

// =====================================================================
// Generic entry screen (Time / Expenses / Mileage)
// =====================================================================
// A spec describes one Excel table: its fields, defaults, and list columns.
// Fields marked `virtual` are form-only; `derive` turns them into columns and
// fills in the columns that don't need your input.
const isInternal = p => p?.Type === 'Internal';
const isOpportunity = p => (cfg().opportunityStatuses || []).includes(p?.Status);
// Internal and opportunity (pre-PO) work isn't billable by default
const defaultBillable = p => !!p && !isInternal(p) && !isOpportunity(p);
// Internal projects first (most used), then client projects grouped by
// client, then opportunities / on-hold work waiting for a PO.
function projectOptions(all) {
  const open = [...cfg().activeStatuses, ...(cfg().opportunityStatuses || [])];
  const ps = rows('Projects').filter(p => all || open.includes(p.Status));
  const groups = [{ group: 'Internal', items: ps.filter(isInternal).map(p => p.Name) }];
  const byClient = {};
  ps.filter(p => !isInternal(p) && !isOpportunity(p)).forEach(p => (byClient[p.Client] ||= []).push(p.Name));
  Object.keys(byClient).sort().forEach(c => groups.push({ group: c, items: byClient[c] }));
  groups.push({ group: 'Opportunities & on hold (not billable until PO)', items: ps.filter(p => !isInternal(p) && isOpportunity(p)).map(p => p.Name) });
  return groups.filter(g => g.items.length);
}
const projectRow = name => rows('Projects').find(p => p.Name === name);
const distinct = (table, col, filter = () => true) => {
  const seen = new Map();
  for (const r of [...rows(table)].reverse()) if (filter(r) && clean(r[col]) && !seen.has(clean(r[col]).toLowerCase())) seen.set(clean(r[col]).toLowerCase(), clean(r[col]));
  return [...seen.values()].slice(0, 60);
};

const TIME = {
  key: 'time', table: 'Tasks', title: 'Time', addLabel: 'Log time', noun: 'time entry', defPeriod: 'week',
  headActions: '<a class="btn" href="#/projects/time">+ New project</a>',
  fields: [
    { col: 'Date', type: 'date', def: todayISO, required: true },
    { col: 'Project', type: 'select', options: () => projectOptions(false), required: true, editOptions: () => projectOptions(true), extra: '<div id="budget"></div>' },
    { col: 'Hours', type: 'number', step: '0.25', min: '0', required: true },
    { col: 'Task', type: 'text', list: () => distinct('Tasks', 'Task'), placeholder: 'What did you work on?', required: true },
    { col: 'Worker', type: 'select', options: () => rows('Workers').map(w => w.Name), def: () => me, sticky: true },
    { col: 'Billable', type: 'check' },
    { col: 'Invoiced', type: 'check', editOnly: true },
    { col: 'Paid', type: 'check', editOnly: true },
    { col: 'Invoice', type: 'text', editOnly: true, label: 'Invoice #', hint: 'Filled in when the invoice is recorded.' },
  ],
  list: [
    { label: 'Date', get: r => toISO(r.Date) },
    { label: 'Project', get: r => esc(r.Project) },
    { label: 'Task', get: r => esc(r.Task), cls: 'grow' },
    { label: 'Worker', get: r => esc(String(r.Worker).split(' ')[0]), cls: 'hide-sm' },
    { label: 'Invoice', cls: 'hide-sm', get: r => timeStatus(r) },
    { label: 'Hours', get: r => hrs(r.Hours), cls: 'r' },
  ],
  summary: (shown) => `${hrs(shown.reduce((s, r) => s + num(r.Hours), 0))} h`,
  filterCol: 'Project',
  // PO approved: turn a project's non-billable, uninvoiced time billable so it
  // goes on the next invoice, whatever period it was worked in.
  bulk(shown, project) {
    if (!project) return null;
    const p = projectRow(project);
    if (isInternal(p)) return null;
    const list = shown.filter(r => r.Billable !== true && r.Invoiced !== true);
    if (!list.length) return null;
    const h = list.reduce((s, r) => s + num(r.Hours), 0);
    return {
      html: `<div class="bulk"><span><b>${list.length}</b> non-billable ${list.length === 1 ? 'entry' : 'entries'} (${hrs(h)} h) shown for ${esc(project)}. PO approved?</span><button class="btn sm primary" id="bulkgo">Make billable</button></div>`,
      async run() {
        if (!confirm(`Set Billable = TRUE on ${list.length} time entries (${hrs(h)} h) for ${project}?\n\nThey will go on the next invoice for ${p?.Client || 'the client'}.`)) return false;
        await busy('Updating time entries…', async () => {
          await store.setColumn('Tasks', list, 'Billable', true, k => { $('#busy span').textContent = `Updating time entries… ${k}/${list.length}`; });
          logActivity('Made billable', `Tasks ×${list.length}`, `${project} · ${hrs(h)} h`);
        });
        toast(`${list.length} entries on ${esc(project)} are now billable.${isOpportunity(p) ? ` The project's status is still <b>${esc(p.Status)}</b>; change it to Active in the Projects table when the PO is in.` : ''}`, 'ok', 9000);
        return true;
      },
    };
  },
  afterRender(form, editing) {
    const proj = form.elements.Project, bill = form.elements.Billable;
    // Internal (Internal, Training, …) and opportunity projects are non-billable by default
    const setBillable = () => { const p = projectRow(proj.value); bill.checked = p ? defaultBillable(p) : true; };
    const setList = () => {
      const dl = $('#dl-Task'); if (!dl) return;
      const opts = [...distinct('Tasks', 'Task', r => r.Project === proj.value), ...distinct('Tasks', 'Task')];
      dl.innerHTML = [...new Set(opts)].slice(0, 60).map(o => `<option value="${esc(o)}">`).join('');
    };
    // Budget: hours left on fixed-fee projects, updating as you type the hours.
    // When editing, this entry is already counted, so only the change matters.
    const paintBudget = () => {
      const p = projectRow(proj.value);
      const counted = editing && editing.Project === proj.value && editing.Billable === true && editing.Invoiced !== true ? num(editing.Hours) : 0;
      $('#budget').innerHTML = budgetHTML(p, bill.checked ? num(form.elements.Hours.value) - counted : -counted);
    };
    proj.onchange = () => { setBillable(); setList(); paintBudget(); };
    form.elements.Hours.oninput = paintBudget;
    bill.onchange = paintBudget;
    const pick = !editing && ls.get('pickProject', '');
    if (pick) { ls.set('pickProject', ''); proj.value = pick; }
    if (!editing) setBillable();
    setList();
    paintBudget();
  },
};

// Invoice status of a time entry. Uninvoiced billable time shows the invoice
// it is expected to go on (display only; nothing is written until recorded).
function timeStatus(r) {
  const col = linkCol('Tasks');
  const no = col && clean(r[col]);
  if (no) return `<a class="pill ${r.Paid ? '' : 'solid'}" href="#/invoice/${encodeURIComponent(no)}" onclick="event.stopPropagation()" title="${r.Paid ? 'Paid' : 'Invoiced, not yet paid'}">${esc(no)}${r.Paid ? ' ✓' : ''}</a>`;
  if (r.Invoiced && col && trackingStart() !== null && num(r.Date) >= trackingStart()) return `<a class="pill warn" href="#/invoices" onclick="event.stopPropagation()" title="Marked Invoiced without an invoice number. See Invoice check">Invoiced, no #</a>`;
  if (r.Invoiced) return `<span class="pill solid">${r.Paid ? 'Paid' : 'Invoiced'}</span>`;
  if (!r.Billable) return '';
  const client = rows('Clients').find(c => c.Name === projectRow(r.Project)?.Client);
  const bd = nextBillingDay().iso;
  return client ? `<span class="pill" title="Expected on the ${weekday(bd)} invoice">→ ${esc(clean(client.Short) + bd.slice(2, 4) + bd.slice(5, 7))}</span>` : '<span class="pill">Billable</span>';
}

const HST_FORMULA = '=Expenses[[#This Row],[Amount]]*13/113';
const hasHst = r => r.Tax === '' || Math.abs(num(r.Tax) - num(r.Amount) * 13 / 113) < 0.005;

const EXPENSE = {
  key: 'expenses', table: 'Expenses', title: 'Expenses', addLabel: 'Add expense', noun: 'expense', defPeriod: 'month',
  headActions: '<a class="btn" href="#/report">Expense report</a>',
  fields: [
    { col: 'Date', type: 'date', def: todayISO, required: true },
    { col: 'Vendor', type: 'text', list: () => distinct('Expenses', 'Vendor'), required: true },
    { col: 'Description', type: 'text', list: () => distinct('Expenses', 'Description') },
    { col: 'Category', type: 'select', options: () => rows('Categories').map(c => c['CRA Categories']), required: true },
    { col: 'Amount', type: 'number', step: '0.01', label: 'Total paid (incl. tax)', required: true },
    { col: '_hst', type: 'check', virtual: true, inline: true, label: 'Includes 13% HST', def: () => true, load: hasHst, hint: '<span id="taxhint"></span>' },
    { col: '_tax', type: 'number', virtual: true, step: '0.01', min: '0', label: 'Tax included ($)', def: () => '', load: r => num(r.Tax), hint: 'Use 0 for no tax, or the GST/PST/other amount shown on the receipt.' },
    { col: 'Business', type: 'percent', label: 'Business use %', def: () => 1, sticky: true },
    { col: 'Project', type: 'select', options: () => ['', ...projectOptions(false)], editOptions: () => ['', ...projectOptions(true)], label: 'Project (optional)' },
    { col: 'Billable', type: 'check', label: 'Billable to client', def: () => true },
    { col: 'Reimburse', type: 'check', label: 'Paid personally (reimburse me)', sticky: true },
  ],
  list: [
    { label: 'Date', get: r => toISO(r.Date) },
    { label: 'Vendor', get: r => esc(r.Vendor) },
    { label: 'Description', get: r => esc(r.Description), cls: 'grow hide-sm' },
    { label: 'Category', get: r => esc(r.Category), cls: 'hide-sm' },
    { label: '', get: r => (r.Scan ? '<span class="pill" title="Receipt scanned">🧾</span>' : '') + (r.Reimburse && !r.Paid ? '<span class="pill warn" title="Not yet reimbursed">Owed</span>' : '') },
    { label: 'Amount', get: r => money(r.Amount), cls: 'r' },
  ],
  summary: (shown) => `${money(shown.reduce((s, r) => s + num(r['Bus Total']), 0))} business`,
  top: () => `
    <div class="receipt" id="receipt">
      <div class="receipt-btns">
        <label class="btn primary">📷 Take photo<input type="file" accept="image/*" capture="environment" hidden id="cam"></label>
        <label class="btn">Choose file<input type="file" accept="image/*,application/pdf" hidden id="pick"></label>
      </div>
      <div id="rcpt-preview" class="rcpt-preview" hidden></div>
      <p class="small muted">The receipt is read on this device to pre-fill the form, then saved to <code>${esc(CFG.receiptFolder)}</code> next to the workbook.</p>
    </div>`,
  afterRender(form, editing) {
    const f = form.elements;
    const show = (name, on) => { f[name].closest('.field, .check-wrap').hidden = !on; };
    const sync = () => {
      show('_tax', !f._hst.checked);
      $('#taxhint').textContent = f._hst.checked && num(f.Amount.value) ? `HST ${money(num(f.Amount.value) * 13 / 113)}, calculated by Excel` : '';
      // Without a project the expense can't be billed to a client
      show('Billable', !!f.Project.value);
    };
    f.Amount.oninput = sync; f._hst.onchange = () => { if (!f._hst.checked && f._tax.value === '') f._tax.value = 0; sync(); };
    f.Project.onchange = () => { const p = projectRow(f.Project.value); if (!editing || !editing.Project) f.Billable.checked = defaultBillable(p); sync(); };
    if (!editing) f.Billable.checked = false;
    sync();
    // remember category per vendor
    f.Vendor.onchange = () => {
      const last = [...rows('Expenses')].reverse().find(e => clean(e.Vendor).toLowerCase() === clean(f.Vendor.value).toLowerCase());
      if (last && !editing) f.Category.value = last.Category;
    };
    EXPENSE.receipt = null;
    const onFile = async (file) => {
      if (!file) return;
      const blob = await compressImage(file);
      EXPENSE.receipt = blob;
      const pv = $('#rcpt-preview');
      pv.hidden = false;
      pv.innerHTML = (blob.type.startsWith('image/') ? `<img src="${URL.createObjectURL(blob)}" alt="Receipt">` : `<div class="pill">${esc(file.name)}</div>`) + '<div class="ocr" id="ocr">Reading receipt…</div>';
      if (!blob.type.startsWith('image/')) { $('#ocr').textContent = 'PDF attached. It will be saved and the row marked as scanned.'; return; }
      try {
        const r = await readReceipt(blob, p => { const o = $('#ocr'); if (o) o.textContent = `Reading receipt… ${p}%`; });
        const filled = [];
        if (r.total && !f.Amount.value) { f.Amount.value = r.total.toFixed(2); filled.push('amount'); }
        if (r.date && /^\d{4}-\d\d-\d\d$/.test(r.date) && f.Date.value === todayISO()) { f.Date.value = r.date; filled.push('date'); }
        if (r.vendor && !f.Vendor.value) {
          const known = distinct('Expenses', 'Vendor').find(v => r.vendor.toLowerCase().includes(v.toLowerCase()) || v.toLowerCase().includes(r.vendor.toLowerCase()));
          f.Vendor.value = known || r.vendor; f.Vendor.onchange(); filled.push('vendor');
        }
        sync();
        $('#ocr').innerHTML = (filled.length ? `Pre-filled ${filled.join(', ')} from the receipt. <b>Please check them.</b>` : 'Could not read details; please fill in the form.') + ' The row will be marked as scanned.';
      } catch (e) { $('#ocr').textContent = 'Could not read the receipt automatically: ' + e.message; }
    };
    $('#cam').onchange = e => onFile(e.target.files[0]);
    $('#pick').onchange = e => onFile(e.target.files[0]);
    if (location.hash.endsWith('/scan') && !editing) setTimeout(() => $('#cam').click(), 150);
  },
  derive(v, editing) {
    const hst = v._hst, tax = v._tax;
    delete v._hst; delete v._tax;
    // Tax: leave Excel's HST formula alone; a custom amount replaces it for this row only
    if (!hst) v.Tax = num(tax);
    else if (editing && !hasHst(editing)) v.Tax = HST_FORMULA;
    if (!v.Project) v.Billable = false;
    // Not reimbursable means the company paid it, so it's already paid
    if (!editing) { v.Paid = !v.Reimburse; v.Scan = false; }
    else if (v.Reimburse !== editing.Reimburse) v.Paid = !v.Reimburse;
  },
  async beforeSave(values) {
    if (!EXPENSE.receipt) return '';
    const ext = EXPENSE.receipt.type === 'application/pdf' ? 'pdf' : 'jpg';
    const name = `${toISO(values.Date)}_${clean(values.Vendor).replace(/[^\w-]+/g, '-')}_${num(values.Amount).toFixed(2)}.${ext}`;
    const up = await store.uploadFile(EXPENSE.receipt, name);
    values.Scan = true;
    values.Receipt = up.webUrl; // only written if your Expenses table has a "Receipt" column
    EXPENSE.receipt = null;
    return ` · receipt saved as <a href="${esc(up.webUrl)}" target="_blank" rel="noopener">${esc(up.name)}</a>`;
  },
};

const PER_DIEM = [['PD AM', 'Breakfast'], ['PD Noon', 'Lunch'], ['PD PM', 'Dinner'], ['PD Inc', 'Incidentals'], ['PD PA', 'Private accommodations']];
const MILEAGE = {
  key: 'mileage', table: 'Mileage', title: 'Mileage', addLabel: 'Log a trip', noun: 'trip', defPeriod: 'month',
  fields: [
    { col: 'Date', type: 'date', def: todayISO, required: true },
    { col: 'Destination', type: 'text', list: () => distinct('Mileage', 'Destination'), required: true },
    { col: 'Reason For Trip', type: 'text', list: () => distinct('Mileage', 'Reason For Trip'), required: true },
    { col: 'km', type: 'number', step: '1', min: '0', label: 'Distance (km, round trip)', required: true,
      extra: '<details class="odo"><summary>Use odometer</summary><div class="row2"><input type="number" id="odo1" placeholder="Start"><input type="number" id="odo2" placeholder="End"></div></details>' },
    { col: 'EV', type: 'check', label: 'Electric vehicle', sticky: true },
    // PD Paid is set on the spreadsheet after per diems are paid out
    ...PER_DIEM.map(([col, label]) => ({ col, label, type: 'check', group: 'Per diem' })),
  ],
  list: [
    { label: 'Date', get: r => toISO(r.Date) },
    { label: 'Destination', get: r => esc(r.Destination) },
    { label: 'Reason', get: r => esc(r['Reason For Trip']), cls: 'grow hide-sm' },
    { label: 'Per diem', cls: 'hide-sm', get: r => PER_DIEM.filter(([c]) => r[c] === true).map(([, l]) => `<span class="pill" title="${l}">${l === 'Private accommodations' ? 'PA' : l}</span>`).join('') },
    { label: 'km', get: r => num(r.km).toLocaleString('en-CA'), cls: 'r' },
  ],
  summary: (shown) => `${shown.reduce((s, r) => s + num(r.km), 0).toLocaleString('en-CA')} km`,
  afterRender(form, editing) {
    const f = form.elements;
    const odo = () => { const a = num($('#odo1').value), b = num($('#odo2').value); if (a && b > a) f.km.value = b - a; };
    $('#odo1').oninput = odo; $('#odo2').oninput = odo;
    f.Destination.onchange = () => {
      const last = [...rows('Mileage')].reverse().find(m => clean(m.Destination).toLowerCase() === clean(f.Destination.value).toLowerCase());
      if (last && !editing && !f.km.value) { f.km.value = last.km; if (!f['Reason For Trip'].value) f['Reason For Trip'].value = last['Reason For Trip'] || ''; }
    };
  },
  derive(v, editing) { if (!editing) v['PD Paid'] = false; },
};

function fieldHTML(fd, editing) {
  const label = esc(fd.label || fd.col);
  const id = 'f-' + fd.col.replace(/\W/g, '');
  const req = fd.required ? 'required' : '';
  if (fd.type === 'check') return `<div class="${fd.inline ? 'field' : 'check-wrap'}"><label class="check"><input type="checkbox" name="${esc(fd.col)}"> ${label}</label>${fd.hint ? `<small>${fd.hint}</small>` : ''}</div>`;
  let input;
  if (fd.type === 'select') {
    const opts = (editing && fd.editOptions ? fd.editOptions() : fd.options());
    const opt = o => `<option value="${esc(o)}">${esc(o || '—')}</option>`;
    input = `<select id="${id}" name="${esc(fd.col)}" ${req}>${fd.required ? '<option value="">Choose…</option>' : ''}${opts.map(o =>
      typeof o === 'object' ? `<optgroup label="${esc(o.group)}">${o.items.map(opt).join('')}</optgroup>` : opt(o)).join('')}</select>`;
  } else if (fd.type === 'percent') {
    input = `<div class="suffix"><input id="${id}" name="${esc(fd.col)}" type="number" min="0" max="100" step="1" inputmode="numeric"><span>%</span></div>`;
  } else {
    const list = fd.list ? `list="dl-${esc(fd.col)}"` : '';
    const mode = fd.type === 'number' ? `inputmode="decimal" step="${fd.step || 'any'}" ${fd.min ? `min="${fd.min}"` : ''}` : '';
    input = `<input id="${id}" name="${esc(fd.col)}" type="${fd.type}" ${mode} ${list} ${req} placeholder="${esc(fd.placeholder || '')}" autocomplete="off">` +
      (fd.list ? `<datalist id="dl-${esc(fd.col)}">${fd.list().map(o => `<option value="${esc(o)}">`).join('')}</datalist>` : '');
  }
  return `<div class="field"><label for="${id}">${label}</label>${input}${fd.extra || ''}${fd.hint ? `<small>${fd.hint}</small>` : ''}</div>`;
}

function viewEntry(spec, view = $('#view'), editing = null) {
  const t = store.tables[spec.table];
  // skip fields for columns this workbook doesn't have (e.g. the optional Invoice column)
  const fields = spec.fields.filter(f => (!f.editOnly || editing) && (f.virtual || t.headers.includes(f.col)));
  const groups = {};
  fields.filter(f => f.group).forEach(f => (groups[f.group] ||= []).push(f));
  const main = fields.filter(f => !f.group && (f.type !== 'check' || f.inline));
  const checks = fields.filter(f => !f.group && f.type === 'check' && !f.inline);
  const saved = ls.get('period:' + spec.key, '');
  const period = PERIODS[saved] ? saved : spec.defPeriod;

  view.innerHTML = `
    <div class="page-head"><h1>${spec.title}</h1>${spec.headActions || ''}</div>
    <div class="grid2 entry">
      <form class="card form ${editing ? 'editing' : ''}" id="entry" novalidate>
        <h2>${editing ? `Edit ${spec.noun} <span class="muted small">· ${esc(t.sheet)} row ${editing._row}</span>` : spec.addLabel}</h2>
        ${spec.top ? spec.top() : ''}
        ${main.map(f => fieldHTML(f, editing)).join('')}
        <div class="checks">${checks.map(f => fieldHTML(f, editing)).join('')}</div>
        ${Object.entries(groups).map(([g, fs]) => `<details class="group" ${fs.some(f => editing?.[f.col]) ? 'open' : ''}><summary>${g}</summary><div class="checks">${fs.map(f => fieldHTML(f, editing)).join('')}</div></details>`).join('')}
        <p class="dest">${icon('M4 4h16v16H4zM4 9h16M9 4v16')}<span>${editing ? 'Updates' : 'Adds a row to'} table <b>${spec.table}</b> on sheet <b>${esc(t.sheet)}</b>. Formula columns are left to Excel.</span></p>
        <div class="actions">
          <button class="btn primary" type="submit">${editing ? 'Save changes' : 'Save'}</button>
          ${editing ? '<button class="btn" type="button" id="cancel">Cancel</button><button class="btn danger" type="button" id="del">Delete row</button>' : '<button class="btn" type="reset" id="clear">Clear</button>'}
        </div>
      </form>
      <div class="card">
        <div class="list-head">
          <select id="period">${Object.entries(PERIODS).map(([k, l]) => `<option value="${k}" ${k === period ? 'selected' : ''}>${l}</option>`).join('')}</select>
          ${spec.filterCol ? `<select id="lf"><option value="">All ${spec.filterCol.toLowerCase()}s</option>${[...new Set(rows(spec.table).map(r => clean(r[spec.filterCol])).filter(Boolean))].sort().map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('')}</select>` : ''}
          <input type="search" id="q" placeholder="Search…">
          <span class="muted small" id="sum"></span>
        </div>
        <div id="bulkbar"></div>
        <div id="listwrap"></div>
      </div>
    </div>`;

  const form = $('#entry');
  // populate
  const sticky = ls.get('sticky:' + spec.key, {});
  for (const fd of fields) {
    const el = form.elements[fd.col];
    let v = editing ? (fd.load ? fd.load(editing) : editing[fd.col]) : (fd.sticky && fd.col in sticky ? sticky[fd.col] : fd.def?.());
    if (fd.type === 'check') el.checked = v === true;
    else if (fd.type === 'date') el.value = editing ? toISO(v) : (v || '');
    else if (fd.type === 'percent') el.value = v === '' || v == null ? '' : Math.round(num(v) * 100);
    else el.value = v ?? '';
    if (fd.type === 'select' && editing && v && el.value !== String(v)) {
      el.insertAdjacentHTML('beforeend', `<option value="${esc(v)}">${esc(v)}</option>`); el.value = v;
    }
  }
  spec.afterRender?.(form, editing);

  // list
  const drawList = () => {
    const p = $('#period').value, q = $('#q').value.toLowerCase(), lf = $('#lf')?.value || '';
    ls.set('period:' + spec.key, p);
    const shown = rows(spec.table)
      .filter(r => inPeriod(r.Date, p))
      .filter(r => !lf || clean(r[spec.filterCol]) === lf)
      .filter(r => !q || r._raw.join(' ').toLowerCase().includes(q))
      .sort((a, b) => num(b.Date) - num(a.Date) || b._row - a._row);
    $('#sum').textContent = `${shown.length} rows · ${spec.summary(shown)}`;
    const bulk = spec.bulk?.(shown, lf);
    $('#bulkbar').innerHTML = bulk ? bulk.html : '';
    if (bulk) $('#bulkgo').onclick = async () => { if (await bulk.run()) drawList(); };
    $('#listwrap').innerHTML = shown.length ? `
      <table class="list clickable"><thead><tr>${spec.list.map(c => `<th class="${c.cls || ''}">${c.label}</th>`).join('')}</tr></thead>
      <tbody>${shown.slice(0, 300).map(r => `<tr data-row="${r._row}" class="${editing?._row === r._row ? 'sel' : ''}" title="${esc(t.sheet)} row ${r._row}">${spec.list.map(c => `<td class="${c.cls || ''}">${c.get(r)}</td>`).join('')}</tr>`).join('')}</tbody></table>
      ${shown.length > 300 ? '<p class="small muted">Showing the 300 most recent. Use search or the period filter.</p>' : ''}`
      : '<p class="muted">No entries for this period.</p>';
    $$('#listwrap tr[data-row]').forEach(tr => tr.onclick = () => {
      const r = rows(spec.table).find(x => x._row === Number(tr.dataset.row));
      viewEntry(spec, view, r);
      $('#entry').scrollIntoView({ behavior: 'smooth' });
    });
  };
  $('#period').onchange = drawList;
  if ($('#lf')) $('#lf').onchange = drawList;
  $('#q').oninput = drawList;
  drawList();

  // actions
  if (editing) {
    $('#cancel').onclick = () => viewEntry(spec, view);
    $('#del').onclick = async () => {
      if (!confirm(`Delete ${esc(t.sheet)} row ${editing._row} from the workbook? This removes the row in Excel.`)) return;
      const res = await busy('Deleting…', () => store.deleteRow(spec.table, editing));
      logActivity(`Deleted ${spec.noun}`, res.address);
      toast(`Deleted row ${esc(res.address)}`, 'info');
      viewEntry(spec, view);
    };
  } else {
    $('#clear').onclick = (e) => { e.preventDefault(); viewEntry(spec, view); };
  }

  form.onsubmit = async (e) => {
    e.preventDefault();
    const missing = fields.filter(fd => fd.required && !clean(form.elements[fd.col].value));
    if (form.elements._hst && !form.elements._hst.checked && form.elements._tax.value === '') form.elements._tax.value = 0;
    if (missing.length) { toast('Please fill in: ' + missing.map(m => m.label || m.col).join(', '), 'warn'); form.elements[missing[0].col].focus(); return; }
    const values = {};
    for (const fd of fields) {
      const el = form.elements[fd.col];
      if (fd.type === 'check') values[fd.col] = el.checked;
      else if (fd.type === 'date') values[fd.col] = toSerial(el.value);
      else if (fd.type === 'number') values[fd.col] = el.value === '' ? '' : Number(el.value);
      else if (fd.type === 'percent') values[fd.col] = el.value === '' ? 1 : Number(el.value) / 100;
      else values[fd.col] = clean(el.value);
    }
    if (!editing) for (const fd of spec.fields) if (fd.editOnly) values[fd.col] = false;
    spec.derive?.(values, editing);
    // remember sticky values
    const st = {};
    fields.filter(f => f.sticky).forEach(f => st[f.col] = values[f.col]);
    ls.set('sticky:' + spec.key, st);

    await busy(editing ? 'Saving changes…' : 'Saving to Excel…', async () => {
      const extra = spec.beforeSave ? await spec.beforeSave(values) : '';
      const res = editing ? await store.updateRow(spec.table, editing, values) : await store.addRow(spec.table, values);
      logActivity(editing ? `Updated ${spec.noun}` : `Added ${spec.noun}`, res.address, summaryLine(spec, values));
      toast(`${editing ? 'Updated' : 'Added'} <code>${esc(res.address)}</code>${extra}`, 'ok', 7000);
    });
    viewEntry(spec, view);
  };
}

function summaryLine(spec, v) {
  if (spec.table === 'Tasks') return `${v.Hours} h · ${v.Project}`;
  if (spec.table === 'Expenses') return `${money(v.Amount)} · ${v.Vendor}`;
  if (spec.table === 'Mileage') return `${v.km} km · ${v.Destination}`;
  return '';
}

// =====================================================================
// Invoices
// =====================================================================
const linkCol = table => {
  const c = cfg().invoiceLinkColumn;
  return c && store.tables[table]?.headers.includes(c) ? c : null;
};

function computeInvoice(clientName, dateISO) {
  const c = cfg();
  const client = rows('Clients').find(x => x.Name === clientName) || {};
  const projs = rows('Projects').filter(p => p.Client === clientName);
  const ref = p => (p.ClientID && p.ClientID !== 'None') ? p.ClientID : '';
  // Same rules as the FILTER formulas on the "Invoice Creator" sheet
  const hourly = projs.filter(p => num(p.Hours) > 0 && !num(p.Fee))
    .map(p => ({ name: p.Name, ref: ref(p), rate: num(p.Rate), hours: num(p.Hours), amount: num(p['This Invoice']) }));
  const fixed = projs.filter(p => num(p.Fee) > 0 && (num(p['This Invoice']) > 0 || p.Status === 'Active*'))
    .map(p => ({ name: p.Name, ref: ref(p), fee: num(p.Fee), comp: num(p['%Comp']), amount: num(p['This Invoice']), prev: num(p.Billed), remaining: num(p.Remaining) }));
  const disb = rows('Disbursements').filter(d => d.Client === clientName && d.Invoiced !== true)
    .map(d => ({ row: d, project: d.Project, desc: d.Description, cost: num(d.Cost), markup: num(d.Markup), amount: num(d.Billable) }));
  const subtotal = [...hourly, ...fixed, ...disb].reduce((s, l) => s + l.amount, 0);
  const tax = Math.round(subtotal * c.taxRate * 100) / 100;
  const yymm = dateISO.slice(2, 4) + dateISO.slice(5, 7);
  const names = new Set([...hourly, ...fixed].map(l => l.name));
  const tasks = rows('Tasks').filter(t => names.has(t.Project) && t.Billable === true && t.Invoiced !== true);
  return {
    client, dateISO, number: clean(client.Short) + yymm,
    due: toISO(toSerial(dateISO) + c.invoiceTermsDays),
    hourly, fixed, disb, subtotal, tax, total: Math.round((subtotal + tax) * 100) / 100, tasks,
  };
}

const rowClient = (table, r) => table === 'Tasks' ? projectRow(r.Project)?.Client : r.Client;
const stampedWith = (table, number) => { const col = linkCol(table); return col ? rows(table).filter(r => clean(r[col]) === clean(number)) : []; };
const isTracked = inv => stampedWith('Tasks', inv.Number).length + stampedWith('Disbursements', inv.Number).length > 0;

// Rows that belong to an issued invoice.
// - Recorded with invoice tracking: the rows carrying its number.
// - Recorded before tracking: dates can't tell (time can become billable long
//   after it was worked, e.g. once a PO arrives), so rows are only known when
//   this is the client's one unpaid pre-tracking invoice: every invoiced,
//   unpaid row without a number must then be on it. Otherwise, none.
function invoiceItems(inv, { unpaidOnly = true } = {}) {
  let tasks = stampedWith('Tasks', inv.Number), disb = stampedWith('Disbursements', inv.Number);
  let how = 'number';
  if (!tasks.length && !disb.length) {
    const legacyUnpaid = rows('Invoices').filter(i => i.Client === inv.Client && !i.Paid && !isTracked(i));
    const loose = table => rows(table).filter(r => r.Invoiced === true && r.Paid !== true && rowClient(table, r) === inv.Client && !(linkCol(table) && clean(r[linkCol(table)])));
    if (!inv.Paid && legacyUnpaid.length === 1 && legacyUnpaid[0]._row === inv._row) {
      how = 'only-unpaid';
      tasks = loose('Tasks'); disb = loose('Disbursements');
    } else how = !inv.Paid && legacyUnpaid.length > 1 ? 'ambiguous' : 'untracked';
  }
  if (unpaidOnly) { tasks = tasks.filter(r => r.Paid !== true); disb = disb.filter(r => r.Paid !== true); }
  return { tasks: { list: tasks }, disb: { list: disb }, how };
}

function invoiceHTML(inv) {
  const c = cfg();
  const co = rows('Clients').find(x => x.Name === c.companyClientRow) || {};
  const cl = inv.client;
  const section = (title, head, lines, cells) => lines.length ? `
    <h3>${title}</h3>
    <table class="inv-table"><thead><tr>${head.map((h, i) => `<th class="${i > 1 ? 'r' : ''}">${h}</th>`).join('')}</tr></thead>
    <tbody>${lines.map(l => `<tr>${cells(l).map((v, i) => `<td class="${i > 1 ? 'r' : ''}${i === head.length - 1 ? ' amt' : ''}">${v}</td>`).join('')}</tr>`).join('')}</tbody></table>` : '';
  return `
  <article class="invoice">
    <header class="inv-head">
      <div class="inv-brand brandmark">${brandMark()}</div>
      <div class="inv-title"><h1>Invoice</h1>
        <dl><dt>Invoice #</dt><dd><b>${esc(inv.number)}</b></dd>
        ${cl.Notes ? `<dt>Reference</dt><dd>${esc(cl.Notes)}</dd>` : ''}
        <dt>Date of issue</dt><dd>${longDate(inv.dateISO)}</dd>
        <dt>Due date</dt><dd>${longDate(inv.due)}</dd></dl>
      </div>
    </header>
    <div class="inv-parties">
      <div><h4>From</h4><p><b>${esc(c.companyName)}</b><br>${esc(co.Address)}<br>${esc(co.Phone)}<br>${esc(co.email)}</p></div>
      <div><h4>Bill to</h4><p>${cl.Contact ? esc(cl.Contact) + '<br>' : ''}<b>${esc(cl.Name)}</b><br>${esc(cl.Address)}<br>${esc(cl.email)}<br>${esc(cl.Phone)}</p></div>
    </div>
    ${section('Hourly projects', ['Project', 'Reference #', 'Rate', 'Hours', 'This invoice'], inv.hourly,
      l => [esc(l.name), esc(l.ref), money(l.rate), hrs(l.hours), money(l.amount)])}
    ${section('Fixed fee projects', ['Project', 'Reference #', 'Fee', '% Complete', 'Previously billed', 'Remaining', 'This invoice'], inv.fixed,
      l => [esc(l.name), esc(l.ref), money(l.fee), Math.round(l.comp * 100) + '%', money(l.prev), money(l.remaining), money(l.amount)])}
    ${section('Disbursements', ['Project', 'Description', 'Cost', 'Markup', 'This invoice'], inv.disb,
      l => [esc(l.project), esc(l.desc), money(l.cost), Math.round(l.markup * 100) + '%', money(l.amount)])}
    <table class="inv-totals">
      <tr><td>Subtotal</td><td class="r">${money(inv.subtotal)}</td></tr>
      <tr><td>${esc(c.taxLabel)} ${Math.round(c.taxRate * 100)}% <small>#${esc(c.taxNumber)}</small></td><td class="r">${money(inv.tax)}</td></tr>
      <tr class="due"><td>Amount due (CAD)</td><td class="r">${money(inv.total)}</td></tr>
    </table>
    <footer class="inv-foot">${c.invoiceFooter.map(esc).join('<br>')}<div class="inv-rule">${esc(c.companyName).toUpperCase()}</div></footer>
  </article>`;
}

// Small dialog that resolves to FormData on confirm, or null on cancel.
function modal(html, size = '') {
  return new Promise(resolve => {
    const d = document.createElement('dialog');
    d.className = 'modal ' + size;
    d.innerHTML = `<form method="dialog">${html}</form>`;
    document.body.appendChild(d);
    const form = d.querySelector('form');
    let done = false;
    const finish = (data) => { if (done) return; done = true; d.close(); d.remove(); resolve(data); };
    // Handle the button press directly; the dialog "close" event can be delayed
    form.addEventListener('submit', e => {
      e.preventDefault();
      if (e.submitter?.value !== 'ok') return finish(null);
      if (!form.reportValidity()) return;
      finish(new FormData(form));
    });
    d.addEventListener('cancel', () => finish(null)); // Esc key
    d.querySelectorAll('a[data-close]').forEach(a => a.addEventListener('click', () => finish(null)));
    d.showModal();
  });
}

function viewInvoices(view, sub) {
  const invs = rows('Invoices').sort((a, b) => num(b.Date) - num(a.Date) || b._row - a._row);
  const open = invs.filter(i => !i.Paid);
  const clients = rows('Clients').filter(c => c.Name !== cfg().companyClientRow);
  const preset = ls.get('invClient', '');
  const t = store.tables.Invoices;
  const now = todayISO();
  const bd = billingDay(Number(now.slice(0, 4)), Number(now.slice(5, 7)));

  view.innerHTML = `
    <div class="page-head"><h1>Invoices</h1></div>
    <div class="card" id="newinv">
      <h2>Create an invoice</h2>
      <div class="row-inline">
        <div class="field"><label for="ic">Client</label><select id="ic"><option value="">Choose…</option>${clients.map(c => {
          const amt = computeInvoice(c.Name, now).subtotal;
          return `<option value="${esc(c.Name)}" ${c.Name === preset ? 'selected' : ''}>${esc(c.Name)}${amt ? ' · ' + money(amt) + ' ready' : ''}</option>`;
        }).join('')}</select></div>
        <div class="field"><label for="idate">Date of issue</label><input type="date" id="idate" value="${bd.iso}">
          <small>This month's billing day: ${weekday(bd.iso)}${bd.shifted ? ` (Monday is ${esc(bd.reason)})` : ''}.</small></div>
      </div>
      <div id="invprev"></div>
    </div>
    ${invoiceCheckCard()}
    <div class="card">
      <div class="list-head"><h2>Issued</h2><span class="muted small">${open.length} open · ${money(open.reduce((s, i) => s + num(i.Amount), 0))} outstanding</span></div>
      <table class="list clickable"><thead><tr><th>Date</th><th>Number</th><th class="hide-sm">Client</th><th class="r">Amount</th><th>Status</th><th></th></tr></thead>
      <tbody>${invs.map(i => {
        const due = toISO(num(i.Date) + cfg().invoiceTermsDays);
        const status = i.Paid ? `<span class="pill">Paid ${toISO(i.Paid)}</span>` : (due < now ? `<span class="pill warn">Overdue · ${num(i.Days)} d</span>` : `<span class="pill solid">Due ${due}</span>`);
        return `<tr data-inv="${esc(i.Number)}" title="Open ${esc(i.Number)} · ${esc(t.sheet)} row ${i._row}"><td>${toISO(i.Date)}</td><td><b>${esc(i.Number)}</b></td><td class="hide-sm">${esc(i.Client)}</td><td class="r">${money(i.Amount)}</td><td>${status}</td>
          <td class="r">${i.Paid ? '' : `<button class="btn sm" data-paid="${i._row}">Mark paid</button>`}</td></tr>`;
      }).join('')}</tbody></table>
    </div>`;

  $$('[data-paid]', view).forEach(b => b.onclick = (e) => { e.stopPropagation(); markPaid(rows('Invoices').find(x => x._row === Number(b.dataset.paid)), () => viewInvoices(view)); });
  $$('tr[data-inv]', view).forEach(tr => tr.onclick = () => { location.hash = '#/invoice/' + encodeURIComponent(tr.dataset.inv); });
  wireInvoiceCheck(view);

  const draw = () => {
    const client = $('#ic').value, date = $('#idate').value || bd.iso;
    ls.set('invClient', client);
    if (!client) { $('#invprev').innerHTML = '<p class="muted">Pick a client to preview the invoice. Lines come from the Projects and Disbursements tables, using the same rules as your Invoice Creator sheet.</p>'; return; }
    const inv = computeInvoice(client, date);
    const dupe = rows('Invoices').some(i => clean(i.Number) === inv.number);
    if (!inv.hourly.length && !inv.fixed.length && !inv.disb.length) { $('#invprev').innerHTML = `<p class="muted">Nothing to invoice for ${esc(client)} right now.</p>`; return; }
    const link = linkCol('Tasks');
    $('#invprev').innerHTML = `
      <div class="preview-wrap">${invoiceHTML(inv)}</div>
      <div class="record">
        ${dupe ? `<p class="warnbox">Invoice number <b>${esc(inv.number)}</b> already exists in the Invoices table. Recording again will create a duplicate.</p>` : ''}
        <label class="check"><input type="checkbox" id="markinv" checked> When recording, mark the ${inv.tasks.length} included time entries and ${inv.disb.length} disbursements as <b>Invoiced</b>${link ? ` and write <b>${esc(inv.number)}</b> in their ${esc(link)} column` : ''}</label>
        <p class="dest"><span>Recording adds 1 row to table <b>Invoices</b> (${esc(inv.number)}, ${money(inv.total)})${inv.tasks.length || inv.disb.length ? ` and sets Invoiced = TRUE on ${inv.tasks.length} row(s) in <b>Tasks</b> and ${inv.disb.length} in <b>Disbursements</b>` : ''}.</span></p>
        <div class="actions">
          <button class="btn" id="print">Print / Save PDF</button>
          <button class="btn primary" id="rec">Record invoice in Excel</button>
        </div>
      </div>`;
    $('#print').onclick = async () => {
      await busy('Checking the workbook…', () => store.loadAll(['Projects', 'Tasks', 'Disbursements', 'Invoices']));
      const fresh = computeInvoice(client, date);
      if (await preInvoiceCheck(client, date, fresh, 'print')) printInvoice(fresh);
    };
    $('#rec').onclick = () => recordInvoice(client, date, view);
  };
  $('#ic').onchange = draw;
  $('#idate').onchange = draw;
  draw();
  if (sub === 'new') $('#newinv').scrollIntoView();
}

async function markPaid(inv, after) {
  await busy('Checking the workbook…', () => store.loadAll(['Invoices', 'Tasks', 'Disbursements']));
  inv = rows('Invoices').find(x => x._row === inv._row && clean(x.Number) === clean(inv.Number)) || inv;
  const { tasks, disb, how } = invoiceItems(inv);
  const n = tasks.list.length, m = disb.list.length;
  const hours = tasks.list.reduce((s, r) => s + num(r.Hours), 0);
  const data = await modal(`
    <h2>Mark ${esc(inv.Number)} paid</h2>
    <p class="muted small">${esc(inv.Client)} · ${money(inv.Amount)} · issued ${toISO(inv.Date)}</p>
    <div class="field"><label for="pd">Payment received</label><input type="date" id="pd" name="date" value="${todayISO()}" required></div>
    ${n || m ? `
      <label class="check"><input type="checkbox" name="cascade" checked> Also set <b>Paid</b> on ${n} time entries (${hrs(hours)} h)${m ? ` and ${m} disbursements` : ''}</label>
      <p class="small muted">${how === 'number' ? 'Matched by invoice number.' : `Recorded before invoice tracking. It's the only unpaid invoice from then for ${esc(inv.Client)}, so these are all of that client's invoiced, unpaid entries.`}</p>
      ${n ? `<details><summary>Show time entries</summary><table class="list"><tbody>${tasks.list.map(r => `<tr><td>${toISO(r.Date)}</td><td>${esc(r.Project)}</td><td class="r">${hrs(r.Hours)} h</td></tr>`).join('')}</tbody></table></details>` : ''}`
      : how === 'ambiguous'
        ? `<p class="warnbox small">This invoice was recorded before invoice tracking, and ${esc(inv.Client)} has other unpaid invoices from then, so the app can't tell which time entries were on it. Mark those paid in the spreadsheet as before.</p>`
        : '<p class="small muted">No unpaid time entries are linked to this invoice.</p>'}
    <div class="actions"><button class="btn" value="cancel" formnovalidate>Cancel</button><button class="btn primary" value="ok">Mark paid</button></div>`);
  if (!data) return;
  const d = data.get('date');
  if (!/^\d{4}-\d\d-\d\d$/.test(d || '')) return toast('Please enter a payment date.', 'warn');
  await busy('Recording payment…', async () => {
    const res = await store.updateRow('Invoices', inv, { Paid: toSerial(d) });
    logActivity('Invoice paid', res.address, `${inv.Number} · ${d}`);
    if (data.get('cascade')) {
      if (n) await store.setColumn('Tasks', tasks.list, 'Paid', true, k => { $('#busy span').textContent = `Marking time entries paid… ${k}/${n}`; });
      if (m) await store.setColumn('Disbursements', disb.list, 'Paid', true);
      logActivity('Marked paid', `Tasks ×${n}, Disbursements ×${m}`, inv.Number);
    }
    toast(`${esc(inv.Number)} marked paid in <code>${esc(res.address)}</code>${data.get('cascade') && (n || m) ? ` · ${n} time entries${m ? ` and ${m} disbursements` : ''} marked paid` : ''}`, 'ok', 7000);
  });
  after();
}

function printDoc(html, title) {
  $('#print-root').innerHTML = html;
  const old = document.title;
  document.title = title; // default PDF file name
  window.print();
  document.title = old;
}
const printInvoice = inv => printDoc(invoiceHTML(inv), `Invoice ${inv.number}`);

async function recordInvoice(client, date, view) {
  // Re-read the source tables so we record exactly what Excel has now
  await busy('Checking the workbook…', () => store.loadAll(['Projects', 'Tasks', 'Disbursements', 'Invoices']));
  const inv = computeInvoice(client, date);
  const mark = $('#markinv')?.checked;
  if (!await preInvoiceCheck(client, date, inv, 'record')) return;
  await busy('Recording invoice…', async () => {
    const res = await store.addRow('Invoices', { Date: toSerial(date), Client: client, Number: inv.number, Amount: inv.total });
    logActivity('Recorded invoice', res.address, `${inv.number} · ${money(inv.total)}`);
    if (mark) {
      const disbRows = inv.disb.map(d => d.row);
      const progress = n => { $('#busy span').textContent = `Marking time entries… ${n}/${inv.tasks.length}`; };
      if (inv.tasks.length) await store.setColumn('Tasks', inv.tasks, 'Invoiced', true, progress);
      if (disbRows.length) await store.setColumn('Disbursements', disbRows, 'Invoiced', true);
      // Optional link column: lets "Mark paid" find exactly this invoice's rows later
      const tl = linkCol('Tasks'), dl = linkCol('Disbursements');
      if (tl && inv.tasks.length) await store.setColumn('Tasks', inv.tasks, tl, inv.number);
      if (dl && disbRows.length) await store.setColumn('Disbursements', disbRows, dl, inv.number);
      logActivity('Marked invoiced', `Tasks ×${inv.tasks.length}, Disbursements ×${inv.disb.length}`, inv.number);
    }
    await store.loadAll(['Projects', 'Tasks', 'Disbursements', 'Invoices']);
    toast(`Invoice ${esc(inv.number)} recorded in <code>${esc(res.address)}</code>`, 'ok', 8000);
    // Double-check every included entry now carries the number
    if (mark && linkCol('Tasks')) {
      const stamped = stampedWith('Tasks', inv.number).length;
      if (stamped !== inv.tasks.length) toast(`⚠ Only ${stamped} of ${inv.tasks.length} time entries show invoice number ${esc(inv.number)}. Check the Invoice column in the Tasks table.`, 'warn', 12000);
    }
  });
  viewInvoices(view);
}

// =====================================================================
// Invoice detail: everything billed on one invoice. Serves as the audit trail
// and as a time breakdown you can send the client.
// =====================================================================
function invoiceDetail(inv) {
  const items = invoiceItems(inv, { unpaidOnly: false });
  const { tasks, disb } = items;
  tasks.list.sort((a, b) => num(a.Date) - num(b.Date) || a._row - b._row);
  const byProject = {};
  for (const r of tasks.list) (byProject[r.Project] ||= []).push(r);
  return { tasks, disb, how: items.how, byProject, hours: tasks.list.reduce((s, r) => s + num(r.Hours), 0) };
}

function timeDetailHTML(inv, d) {
  const c = cfg();
  const client = rows('Clients').find(x => x.Name === inv.Client) || {};
  const dates = d.tasks.list.map(r => num(r.Date));
  const period = dates.length ? `${longDate(toISO(Math.min(...dates)))} – ${longDate(toISO(Math.max(...dates)))}` : '';
  return `
  <article class="invoice">
    <header class="inv-head">
      <div class="inv-brand brandmark">${brandMark()}</div>
      <div class="inv-title"><h1>Time detail</h1>
        <dl><dt>Invoice #</dt><dd><b>${esc(inv.Number)}</b></dd>
        <dt>Client</dt><dd>${esc(inv.Client)}</dd>
        ${client.Notes ? `<dt>Reference</dt><dd>${esc(client.Notes)}</dd>` : ''}
        <dt>Invoice date</dt><dd>${longDate(toISO(inv.Date))}</dd>
        ${period ? `<dt>Work period</dt><dd>${period}</dd>` : ''}</dl>
      </div>
    </header>
    ${Object.entries(d.byProject).map(([proj, list]) => {
      const p = projectRow(proj) || {};
      const ref = p.ClientID && p.ClientID !== 'None' ? ` <small>· ${esc(p.ClientID)}</small>` : '';
      return `<h3>${esc(proj)}${ref}</h3>
      <table class="inv-table"><thead><tr><th style="width:7.5em">Date</th><th>Description</th><th style="width:10em">By</th><th class="r" style="width:5em">Hours</th></tr></thead>
      <tbody>${list.map(r => `<tr><td>${toISO(r.Date)}</td><td>${esc(r.Task)}</td><td>${esc(r.Worker)}</td><td class="r">${hrs(r.Hours)}</td></tr>`).join('')}
      <tr class="sub"><td></td><td></td><td>Project total</td><td class="r">${hrs(list.reduce((s, r) => s + num(r.Hours), 0))}</td></tr></tbody></table>`;
    }).join('')}
    ${d.disb.list.length ? `<h3>Disbursements</h3>
      <table class="inv-table"><thead><tr><th style="width:7.5em">Date</th><th>Description</th><th>Project</th><th class="r">Amount</th></tr></thead>
      <tbody>${d.disb.list.map(r => `<tr><td>${toISO(r.Date)}</td><td>${esc(r.Description)}${r.Vendor ? ` (${esc(r.Vendor)})` : ''}</td><td>${esc(r.Project)}</td><td class="r">${money(r.Billable)}</td></tr>`).join('')}</tbody></table>` : ''}
    <table class="inv-totals"><tr class="due"><td>Total hours</td><td class="r">${hrs(d.hours)}</td></tr></table>
    <footer class="inv-foot"><div class="inv-rule">${esc(c.companyName).toUpperCase()}</div></footer>
  </article>`;
}

function viewInvoiceDetail(view, number) {
  const inv = rows('Invoices').find(i => clean(i.Number) === clean(number));
  if (!inv) { view.innerHTML = `<div class="card"><p>Invoice ${esc(number)} was not found. <a href="#/invoices">Back to invoices</a></p></div>`; return; }
  const d = invoiceDetail(inv);
  const due = toISO(num(inv.Date) + cfg().invoiceTermsDays);
  const status = inv.Paid ? `Paid ${toISO(inv.Paid)}` : due < todayISO() ? `Overdue (due ${due})` : `Due ${due}`;
  const linked = d.tasks.list.length || d.disb.list.length;
  view.innerHTML = `
    <div class="page-head">
      <div><a class="small" href="#/invoices">← Invoices</a><h1>${esc(inv.Number)}</h1></div>
      <div class="actions" style="margin:0">
        ${inv.Paid ? '' : '<button class="btn" id="paid">Mark paid</button>'}
        <button class="btn primary" id="ptd" ${d.tasks.list.length ? '' : 'disabled'}>Print time detail</button>
      </div>
    </div>
    <div class="stats">
      <div class="stat"><label>Client</label><b class="sm">${esc(inv.Client)}</b><small>issued ${longDate(toISO(inv.Date))}</small></div>
      <div class="stat"><label>Amount</label><b>${money(inv.Amount)}</b><small>incl. ${esc(cfg().taxLabel)}</small></div>
      <div class="stat"><label>Time billed</label><b>${hrs(d.hours)} h</b><small>${d.tasks.list.length} entries · ${Object.keys(d.byProject).length} projects</small></div>
      <div class="stat ${inv.Paid ? '' : 'accent'}"><label>Status</label><b class="sm">${status}</b><small>${esc(store.tables.Invoices.sheet)} row ${inv._row}</small></div>
    </div>
    ${d.how === 'only-unpaid' ? `<p class="warnbox">Recorded before invoice tracking. It's the only unpaid invoice from then for ${esc(inv.Client)}, so these invoiced, unpaid entries must be the ones on it.</p>` : ''}
    ${linked ? `<div class="preview-wrap">${timeDetailHTML(inv, d)}</div>`
      : `<div class="card"><p class="muted">${d.how === 'number' ? 'No time entries are linked to this invoice.' : `This invoice was recorded before invoice tracking started, so its time entries aren't linked. ${linkCol('Tasks') ? '' : 'Add the Invoice column in <a href="#/settings">Settings → Invoice tracking</a> so future invoices are.'}`}</p></div>`}`;
  $('#ptd').onclick = () => printDoc(timeDetailHTML(inv, d), `Time detail ${inv.Number}`);
  if ($('#paid')) $('#paid').onclick = () => markPaid(inv, () => viewInvoiceDetail(view, number));
}

// =====================================================================
// Projects: overview of open projects + add a new one. Editing and deleting
// projects stays in the spreadsheet.
// =====================================================================
// Fixed-fee budget in hours, from the Projects table's own formulas:
// Total = billed (incl. carryover) + uninvoiced billable time, in dollars.
function budget(p) {
  if (!p) return null;
  const rate = num(p.Rate), fee = num(p.Fee);
  if (p.Type === 'Fixed Fee' && fee > 0 && rate > 0) {
    const total = num(p.Total);
    return { kind: 'fixed', fee, rate, budgetH: fee / rate, usedH: total / rate, remainingH: (fee - total) / rate };
  }
  if (rate > 0 && !isInternal(p)) return { kind: 'hourly', rate, unbilledH: num(p.Hours) };
  return null;
}
function budgetHTML(p, addH = 0) {
  const b = budget(p);
  if (!b) return '';
  if (b.kind === 'hourly') return `<div class="budget"><div class="budget-top"><b>${esc(p.Type)}</b> <span>at ${money(b.rate)}/h · ${hrs(b.unbilledH)} h not yet invoiced</span></div></div>`;
  const after = b.remainingH - addH;
  const pct = Math.min(100, Math.max(0, (b.usedH + Math.max(0, addH)) / b.budgetH * 100));
  const head = b.remainingH >= 0 ? `${hrs(b.remainingH)} h remaining` : `${hrs(-b.remainingH)} h over budget`;
  return `<div class="budget fixed ${after < 0 ? 'over' : ''}">
    <div class="budget-top"><b>${head}</b> <span>of ${hrs(b.budgetH)} h · ${money(b.fee)} fee at ${money(b.rate)}/h</span></div>
    <div class="bar"><i style="width:${pct}%"></i></div>
    <small>${hrs(b.usedH)} h used (billed + not yet invoiced)${addH ? ` · after this entry: <b>${after >= 0 ? hrs(after) + ' h left' : hrs(-after) + ' h over budget'}</b>` : ''}</small>
  </div>`;
}

const listOf = (table, col, fallback) => { const v = rows(table).map(r => clean(r[col])).filter(Boolean); return v.length ? v : fallback; };
const STATUS_LIST = () => listOf('Status', 'Status', ['Active', 'Active*', 'Opportunity', 'Lost', 'Hold', 'Archive']);
const TYPE_LIST = () => listOf('Table12', 'Type', ['Internal', 'Retainer', 'Hourly', 'Fixed Fee']);

// Next Project ID for a client, following the existing pattern (WF-023 → WF-024).
// The prefix is shared across clients (FWC-… for both FWC purchase orders).
function suggestProjectId(client) {
  const ids = rows('Projects').map(p => clean(p.ProjectID));
  const own = rows('Projects').filter(p => p.Client === client).map(p => clean(p.ProjectID).match(/^([A-Za-z]+)-\d+/)).find(Boolean);
  const prefix = own ? own[1] : (clean(rows('Clients').find(c => c.Name === client)?.Short) || 'PRJ').toUpperCase();
  const max = Math.max(0, ...ids.map(id => id.match(new RegExp(`^${prefix}-(\\d+)`))).filter(Boolean).map(m => Number(m[1])));
  return `${prefix}-${String(max + 1).padStart(3, '0')}`;
}

function viewProjects(view, from) {
  const open = [...cfg().activeStatuses, ...(cfg().opportunityStatuses || [])];
  const projects = rows('Projects');
  const t = store.tables.Projects;
  view.innerHTML = `
    <div class="page-head"><h1>Projects</h1><a class="small" href="${esc(store.file.webUrl)}" target="_blank" rel="noopener">Edit or remove projects in Excel ↗</a></div>
    <div class="grid2 entry">
      <form class="card form" id="pform" novalidate>
        <h2>Add a project</h2>
        <div class="field"><label for="p-name">Project name</label><input id="p-name" name="Name" required autocomplete="off" placeholder="e.g. PowerCo Substation Ph2"></div>
        <div class="field"><label for="p-client">Client</label><select id="p-client" name="Client" required><option value="">Choose…</option>${rows('Clients').map(c => `<option>${esc(c.Name)}</option>`).join('')}</select>
          <small>New client? Add it to the Clients table in Excel first, then refresh (↻).</small></div>
        <div class="row2">
          <div class="field"><label for="p-id">Project ID</label><input id="p-id" name="ProjectID" required autocomplete="off"></div>
          <div class="field"><label for="p-ref">Client ref / PO #</label><input id="p-ref" name="ClientID" autocomplete="off" placeholder="optional"></div>
        </div>
        <div class="row2">
          <div class="field"><label for="p-status">Status</label><select id="p-status" name="Status">${STATUS_LIST().map(s => `<option ${s === 'Opportunity' ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select></div>
          <div class="field"><label for="p-type">Type</label><select id="p-type" name="Type">${TYPE_LIST().map(s => `<option ${s === 'Hourly' ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select></div>
        </div>
        <div class="row2">
          <div class="field"><label for="p-rate">Rate ($/h)</label><input id="p-rate" name="Rate" type="number" step="5" min="0" inputmode="decimal"></div>
          <div class="field" id="feefield"><label for="p-fee">Fixed fee ($)</label><input id="p-fee" name="Fee" type="number" step="100" min="0" inputmode="decimal"><small id="feehint"></small></div>
        </div>
        <div class="field"><label for="p-prime">Prime</label><select id="p-prime" name="Prime">${rows('Workers').map(w => `<option ${w.Name === me ? 'selected' : ''}>${esc(w.Name)}</option>`).join('')}</select></div>
        <div class="field"><label for="p-notes">Notes</label><input id="p-notes" name="Notes" autocomplete="off"></div>
        <p class="dest">${icon('M4 4h16v16H4zM4 9h16M9 4v16')}<span>Adds a row to table <b>Projects</b> on sheet <b>${esc(t.sheet)}</b>. Hours, billed and budget columns are Excel formulas and fill in by themselves.</span></p>
        <div class="actions"><button class="btn primary" type="submit">Add project</button></div>
      </form>
      <div class="card">
        <div class="list-head"><h2>Open projects</h2><span class="muted small">${projects.filter(p => open.includes(p.Status)).length} projects</span></div>
        ${open.map(st => {
          const list = projects.filter(p => p.Status === st);
          if (!list.length) return '';
          return `<h3 class="group-h">${esc(st)}</h3><table class="list"><tbody>${list.map(p => {
            const b = budget(p);
            const info = !b ? `<span class="muted">${esc(p.Type)}</span>`
              : b.kind === 'hourly' ? `${money(b.rate)}/h · ${hrs(b.unbilledH)} h unbilled`
              : `<div class="mini"><div class="bar ${b.remainingH < 0 ? 'over' : ''}"><i style="width:${Math.min(100, b.usedH / b.budgetH * 100)}%"></i></div><span>${b.remainingH >= 0 ? `${hrs(b.remainingH)} h left` : `${hrs(-b.remainingH)} h over`} of ${hrs(b.budgetH)}</span></div>`;
            return `<tr data-name="${esc(p.Name)}"><td><b>${esc(p.Name)}</b><br><small class="muted">${esc(p.ProjectID)} · ${esc(p.Client)}</small></td><td class="hide-sm">${esc(p.Type)}</td><td class="r proj-info">${info}</td></tr>`;
          }).join('')}</tbody></table>`;
        }).join('')}
      </div>
    </div>`;

  const f = $('#pform').elements;
  let idTouched = false;
  f.ProjectID.oninput = () => { idTouched = true; };
  f.Client.onchange = () => {
    if (!idTouched) f.ProjectID.value = f.Client.value ? suggestProjectId(f.Client.value) : '';
    // default rate: the client's most recent project rate
    const last = [...projects].reverse().find(p => p.Client === f.Client.value && num(p.Rate));
    if (last && !f.Rate.value) f.Rate.value = num(last.Rate);
  };
  const syncType = () => {
    const fixed = f.Type.value === 'Fixed Fee';
    $('#feefield').hidden = !fixed;
    $('#feehint').textContent = fixed && num(f.Fee.value) && num(f.Rate.value) ? `${hrs(num(f.Fee.value) / num(f.Rate.value))} h budget` : '';
  };
  f.Type.onchange = syncType; f.Fee.oninput = syncType; f.Rate.oninput = syncType;
  syncType();

  $('#pform').onsubmit = async (e) => {
    e.preventDefault();
    const v = name => clean(f[name].value);
    const missing = ['Name', 'Client', 'ProjectID'].filter(n => !v(n));
    if (missing.length) return toast('Please fill in: ' + missing.join(', '), 'warn');
    if (projects.some(p => clean(p.Name).toLowerCase() === v('Name').toLowerCase())) return toast(`A project named “${esc(v('Name'))}” already exists. Project names must be unique; time and invoices look projects up by name.`, 'warn', 8000);
    if (projects.some(p => clean(p.ProjectID).toLowerCase() === v('ProjectID').toLowerCase()) && !confirm(`Project ID ${v('ProjectID')} is already used. Add anyway?`)) return;
    if (f.Type.value === 'Fixed Fee' && !num(f.Fee.value)) return toast('Enter the fixed fee for a Fixed Fee project.', 'warn');
    const row = {
      Name: v('Name'), ProjectID: v('ProjectID'), Client: v('Client'), ClientID: v('ClientID'),
      Status: f.Status.value, Type: f.Type.value, Prime: f.Prime.value,
      Rate: f.Rate.value === '' ? '' : Number(f.Rate.value),
      Fee: f.Type.value === 'Fixed Fee' ? Number(f.Fee.value) : '',
      Notes: v('Notes'),
    };
    let added = false;
    await busy('Adding project…', async () => {
      const res = await store.addRow('Projects', row);
      logActivity('Added project', res.address, `${row.ProjectID} · ${row.Name}`);
      // Read the table back and make sure the project is really there
      await store.loadTable('Projects');
      added = rows('Projects').some(p => clean(p.Name) === row.Name);
      const listed = [...cfg().activeStatuses, ...(cfg().opportunityStatuses || [])].includes(row.Status);
      if (!added) toast(`⚠ Excel accepted the row (<code>${esc(res.address)}</code>) but the project didn't come back when the Projects table was re-read. Press ↻ to refresh, and check the Projects table in Excel.`, 'warn', 15000);
      else toast(`Added <b>${esc(row.Name)}</b> in <code>${esc(res.address)}</code>. ${listed ? 'It’s now in the Time project list.' : `Projects with status ${esc(row.Status)} aren’t offered for new time.`}`, 'ok', 8000);
    });
    if (added && from === 'time') { ls.set('pickProject', row.Name); location.hash = '#/time'; return; }
    viewProjects(view);
    if (added) {
      const tr = $$('#view tr[data-name]').find(t => t.dataset.name === row.Name);
      if (tr) { tr.classList.add('sel'); tr.scrollIntoView({ block: 'center', behavior: 'smooth' }); }
    }
  };
}

// =====================================================================
// Expense report: personally-paid expenses not yet reimbursed (same rule as
// the "Expense Rep Creator" sheet), plus per diems from the Mileage table and,
// optionally, vehicle expenses and the mileage allowance.
// =====================================================================
let RATES = null;
async function loadRates() {
  if (RATES) return RATES;
  const c = cfg();
  const fallback = { pd: [29.05, 29.6, 60.75, 17.5, 50], first: 0.73, threshold: 5000, after: 0.67, fromSheet: false };
  try {
    const pd = await store.readRange(c.perDiemRange);
    const mr = await store.readRange(c.mileageRateRange);
    RATES = { pd: pd.map(r => num(r[1])), first: num(mr[0][0]), threshold: num(mr[0][2]), after: num(mr[1][0]), fromSheet: true };
    if (RATES.pd.length < 5 || RATES.pd.some(x => !x)) throw new Error('rates incomplete');
  } catch { RATES = fallback; }
  return RATES;
}

function buildReport(o, R) {
  const c = cfg();
  const from = toSerial(o.from), to = toSerial(o.to);
  const inRange = r => num(r.Date) >= from && num(r.Date) <= to;
  const expenses = rows('Expenses').filter(r => r.Reimburse === true && r.Paid !== true && inRange(r) && (o.vehicle || r.Category !== c.vehicleCategory))
    .sort((a, b) => num(a.Date) - num(b.Date));
  const trips = rows('Mileage').filter(inRange).sort((a, b) => num(a.Date) - num(b.Date));
  const perDiem = !o.perDiem ? [] : trips.filter(r => r['PD Paid'] !== true).flatMap(r =>
    PER_DIEM.map(([col, label], i) => r[col] === true ? { row: r, date: r.Date, trip: r.Destination, item: label, amount: R.pd[i] } : null).filter(Boolean));
  // Allowance tiers apply to the year's cumulative km, so count every earlier trip
  let mileage = [];
  if (o.mileage) {
    let cum = rows('Mileage').filter(r => num(r.Date) < from).reduce((s, r) => s + num(r.km), 0);
    mileage = trips.map(r => {
      const km = num(r.km), lowKm = Math.max(0, Math.min(km, R.threshold - cum));
      const amount = lowKm * R.first + (km - lowKm) * R.after;
      cum += km;
      return { date: r.Date, trip: r.Destination, reason: r['Reason For Trip'], km, amount };
    });
  }
  const sum = (a, k) => a.reduce((s, x) => s + num(x[k]), 0);
  const tot = { expenses: sum(expenses, 'Bus Total'), perDiem: sum(perDiem, 'amount'), mileage: sum(mileage, 'amount') };
  return { ...o, expenses, perDiem, mileage, tot, total: tot.expenses + tot.perDiem + tot.mileage, perDiemTrips: [...new Set(perDiem.map(p => p.row))] };
}

function reportHTML(r) {
  const c = cfg();
  const table = (title, head, lines, cells, total) => lines.length ? `
    <h3>${title}</h3>
    <table class="inv-table"><thead><tr>${head.map(([h, cls]) => `<th class="${cls || ''}">${h}</th>`).join('')}</tr></thead>
    <tbody>${lines.map(l => `<tr>${cells(l).map((v, i) => `<td class="${head[i][1] || ''}">${v}</td>`).join('')}</tr>`).join('')}
    <tr class="sub">${head.map((_, i) => `<td class="${i === head.length - 1 ? 'r' : ''}">${i === head.length - 2 ? 'Subtotal' : i === head.length - 1 ? money(total) : ''}</td>`).join('')}</tr></tbody></table>` : '';
  return `
  <article class="invoice">
    <header class="inv-head">
      <div class="inv-brand brandmark">${brandMark()}</div>
      <div class="inv-title"><h1>Expense report</h1>
        <dl><dt>Submitted by</dt><dd><b>${esc(r.worker)}</b></dd>
        <dt>Period</dt><dd>${longDate(r.from)} – ${longDate(r.to)}</dd>
        <dt>Next invoice</dt><dd>${longDate(nextBillingDay().iso)}</dd>
        <dt>Date</dt><dd>${longDate(todayISO())}</dd></dl>
      </div>
    </header>
    ${table('Expenses' + (r.vehicle ? '' : ' (excluding motor vehicle)'), [['Date'], ['Vendor'], ['Description'], ['Amount', 'r'], ['Business', 'r'], ['Total', 'r']], r.expenses,
      l => [toISO(l.Date), esc(l.Vendor), esc(l.Description), money(l.Amount), Math.round(num(l.Business) * 100) + '%', money(l['Bus Total'])], r.tot.expenses)}
    ${table('Per diems', [['Date'], ['Trip'], ['Item'], ['Amount', 'r']], r.perDiem,
      l => [toISO(l.date), esc(l.trip), esc(l.item), money(l.amount)], r.tot.perDiem)}
    ${table('Mileage allowance', [['Date'], ['Destination'], ['Reason'], ['km', 'r'], ['Amount', 'r']], r.mileage,
      l => [toISO(l.date), esc(l.trip), esc(l.reason), num(l.km).toLocaleString('en-CA'), money(l.amount)], r.tot.mileage)}
    ${!r.expenses.length && !r.perDiem.length && !r.mileage.length ? '<p class="muted" style="margin-top:1.5rem">Nothing to reimburse for this period.</p>' : ''}
    <table class="inv-totals">
      ${r.tot.expenses ? `<tr><td>Expenses</td><td class="r">${money(r.tot.expenses)}</td></tr>` : ''}
      ${r.tot.perDiem ? `<tr><td>Per diems</td><td class="r">${money(r.tot.perDiem)}</td></tr>` : ''}
      ${r.tot.mileage ? `<tr><td>Mileage allowance</td><td class="r">${money(r.tot.mileage)}</td></tr>` : ''}
      <tr class="due"><td>Total to reimburse (CAD)</td><td class="r">${money(r.total)}</td></tr>
    </table>
    <div class="signoff"><div>Submitted by</div><div>Approved by</div></div>
    <footer class="inv-foot"><div class="inv-rule">${esc(c.companyName).toUpperCase()}</div></footer>
  </article>`;
}

async function viewExpenseReport(view) {
  const R = await loadRates();
  // Reports run alongside invoicing, so they default to the next billing day
  const nb = nextBillingDay();
  const o = { worker: me, from: todayISO().slice(0, 4) + '-01-01', vehicle: false, perDiem: true, mileage: false, ...ls.get('report', {}) };
  o.to = nb.iso.slice(0, 4) === todayISO().slice(0, 4) ? nb.iso : todayISO().slice(0, 4) + '-12-31';
  o.from = o.to.slice(0, 4) + '-01-01';
  view.innerHTML = `
    <div class="page-head"><div><a class="small" href="#/expenses">← Expenses</a><h1>Expense report</h1></div></div>
    <p class="nextbill">${icon('M3 5h18v16H3zM3 10h18M8 3v4M16 3v4')}<span>Next invoice: <b>${weekday(nb.iso)}</b>${nb.shifted ? ` (Monday is ${esc(nb.reason)})` : ''}. The report runs to that day by default.</span></p>
    <div class="card" id="rform">
      <div class="row-inline">
        <div class="field"><label for="r-w">Submitted by</label><select id="r-w">${rows('Workers').map(w => `<option ${w.Name === o.worker ? 'selected' : ''}>${esc(w.Name)}</option>`).join('')}</select></div>
        <div class="field"><label for="r-f">From</label><input type="date" id="r-f" value="${o.from}"></div>
        <div class="field"><label for="r-t">To</label><input type="date" id="r-t" value="${o.to}"></div>
      </div>
      <div class="checks">
        <label class="check"><input type="checkbox" id="r-pd" ${o.perDiem ? 'checked' : ''}> Per diems (from Mileage)</label>
        <label class="check"><input type="checkbox" id="r-veh" ${o.vehicle ? 'checked' : ''}> Motor vehicle expenses</label>
        <label class="check"><input type="checkbox" id="r-km" ${o.mileage ? 'checked' : ''}> Mileage allowance</label>
      </div>
      <p class="small muted">Includes expenses marked <b>Reimburse</b> that aren't <b>Paid</b> yet, per diems on trips where <b>PD Paid</b> is unticked, and (if ticked) the mileage allowance for trips in the period: ${money(R.first)}/km for the first ${num(R.threshold).toLocaleString('en-CA')} km of the year, then ${money(R.after)}/km.
        ${R.fromSheet ? `Rates come from <code>${esc(cfg().perDiemRange)}</code> and <code>${esc(cfg().mileageRateRange)}</code>.` : '<b>Couldn’t read the rates from the Mileage sheet; using the 2026 defaults.</b>'}</p>
    </div>
    <div id="rprev"></div>`;
  const draw = () => {
    Object.assign(o, { worker: $('#r-w').value, from: $('#r-f').value, to: $('#r-t').value, perDiem: $('#r-pd').checked, vehicle: $('#r-veh').checked, mileage: $('#r-km').checked });
    ls.set('report', { perDiem: o.perDiem, vehicle: o.vehicle, mileage: o.mileage });
    const r = buildReport(o, R);
    const n = r.expenses.length, m = r.perDiemTrips.length;
    $('#rprev').innerHTML = `
      <div class="preview-wrap">${reportHTML(r)}</div>
      <div class="actions">
        <button class="btn" id="r-print">Print / Save PDF</button>
        ${n || m ? `<button class="btn primary" id="r-paid">Mark as reimbursed…</button>` : ''}
      </div>`;
    $('#r-print').onclick = () => printDoc(reportHTML(r), `Expense report ${o.worker} ${o.to}`);
    if ($('#r-paid')) $('#r-paid').onclick = async () => {
      if (!confirm(`Mark this report as reimbursed?\n\n• Sets Paid = TRUE on ${n} expense${n === 1 ? '' : 's'}\n• Sets PD Paid = TRUE on ${m} trip${m === 1 ? '' : 's'}${r.mileage.length ? '\n\nThe Mileage table has no paid flag for the km allowance, so nothing is changed for it.' : ''}`)) return;
      await busy('Updating…', async () => {
        if (n) await store.setColumn('Expenses', r.expenses, 'Paid', true);
        if (m) await store.setColumn('Mileage', r.perDiemTrips, 'PD Paid', true);
        logActivity('Expense report reimbursed', `Expenses ×${n}, Mileage ×${m}`, `${o.worker} · ${money(r.total)}`);
      });
      toast(`Marked ${n} expenses and ${m} trips' per diems as reimbursed (${money(r.total)}).`, 'ok', 7000);
      draw();
    };
  };
  ['#r-w', '#r-f', '#r-t', '#r-pd', '#r-veh', '#r-km'].forEach(s => { $(s).onchange = draw; });
  draw();
}

// =====================================================================
// Invoice checks: before printing/recording, confirm the period's time is
// all in, and find invoiced time that has no invoice number (e.g. ticked
// Invoiced by hand in the spreadsheet).
// =====================================================================
// Work covered by an invoice: after the previous month's billing day, up to the invoice date.
function billingPeriod(dateISO) {
  const y = Number(dateISO.slice(0, 4)), m = Number(dateISO.slice(5, 7));
  const prev = billingDay(m === 1 ? y - 1 : y, m === 1 ? 12 : m - 1).iso;
  return { from: toSerial(prev) + 1, to: toSerial(dateISO) };
}

// Tracking begins with the period of the first invoice recorded with numbers.
// Invoiced time from then on should carry a number; earlier time is left alone.
let _track = { key: null, start: null };
function trackingStart() {
  const col = linkCol('Tasks');
  if (!col) return null;
  if (_track.key === store.tables.Tasks) return _track.start;
  const numbers = new Set(rows('Tasks').map(r => clean(r[col])).filter(Boolean));
  const dates = rows('Invoices').filter(i => numbers.has(clean(i.Number))).map(i => num(i.Date));
  const first = dates.length ? Math.min(...dates) : null;
  _track = { key: store.tables.Tasks, start: first === null ? null : billingPeriod(toISO(first)).from };
  return _track.start;
}
function unlinkedTime() {
  const col = linkCol('Tasks'), start = trackingStart();
  if (!col || start === null) return [];
  return rows('Tasks').filter(r => r.Invoiced === true && !clean(r[col]) && num(r.Date) >= start);
}

async function preInvoiceCheck(client, dateISO, inv, action) {
  const per = billingPeriod(dateISO);
  const projects = new Set(rows('Projects').filter(p => p.Client === client).map(p => p.Name));
  const year = rows('Tasks');
  const inPer = year.filter(r => num(r.Date) >= per.from && num(r.Date) <= per.to);
  const workers = rows('Workers').map(w => w.Name).filter(n => year.some(r => r.Worker === n));
  const last = n => Math.max(0, ...year.filter(r => r.Worker === n).map(r => num(r.Date)));
  const nonBill = inPer.filter(r => projects.has(r.Project) && r.Billable !== true && r.Invoiced !== true && !isInternal(projectRow(r.Project)));
  const later = inv.tasks.filter(r => num(r.Date) > per.to);
  const earlier = inv.tasks.filter(r => num(r.Date) < per.from);
  const unlinked = unlinkedTime().filter(r => projects.has(r.Project));
  const warn = [];
  if (unlinked.length) warn.push(`<b>${unlinked.length} ${esc(client)} time ${unlinked.length === 1 ? 'entry is' : 'entries are'} marked Invoiced but ${unlinked.length === 1 ? 'has' : 'have'} no invoice number</b>, probably ticked by hand in the spreadsheet. ${unlinked.length === 1 ? 'It isn’t' : 'They aren’t'} on this invoice. <a href="#/invoices" data-close>Review them under Invoice check</a>.`);
  if (nonBill.length) warn.push(`${nonBill.length} non-billable ${nonBill.length === 1 ? 'entry' : 'entries'} (${hrs(nonBill.reduce((s, r) => s + num(r.Hours), 0))} h) on ${esc(client)} projects this period won't be invoiced: ${[...new Set(nonBill.map(r => esc(r.Project)))].join(', ')}.`);
  if (later.length) warn.push(`${later.length} included ${later.length === 1 ? 'entry is' : 'entries are'} dated after ${toISO(per.to)}, the invoice date.`);
  if (earlier.length) warn.push(`${earlier.length} included ${earlier.length === 1 ? 'entry is' : 'entries are'} from before this period (e.g. made billable after a PO).`);
  const data = await modal(`
    <h2>Before you ${action} ${esc(inv.number)}</h2>
    <p class="small muted">Billing period ${longDate(toISO(per.from))} – ${longDate(toISO(per.to))}. Check everyone's time is in before invoicing.</p>
    <div class="scroll"><table class="list"><thead><tr><th>Worker</th><th class="r">Hours this period</th><th class="r">On ${esc(client)}</th><th>Last entry</th></tr></thead><tbody>
      ${workers.map(n => { const l = last(n); const stale = l < per.to - 3; return `<tr><td>${esc(n)}</td><td class="r">${hrs(inPer.filter(r => r.Worker === n).reduce((s, r) => s + num(r.Hours), 0))}</td><td class="r">${hrs(inPer.filter(r => r.Worker === n && projects.has(r.Project)).reduce((s, r) => s + num(r.Hours), 0))}</td><td>${l ? `<span class="${stale ? 'pill warn' : ''}">${toISO(l)}</span>` : '—'}</td></tr>`; }).join('')}
    </tbody></table></div>
    ${warn.map(w => `<p class="warnbox small">${w}</p>`).join('')}
    ${action === 'record' ? `<p class="dest"><span>Recording adds <b>${esc(inv.number)}</b> (${money(inv.total)}) to the Invoices table${$('#markinv')?.checked ? ` and marks ${inv.tasks.length} time entries and ${inv.disb.length} disbursements as Invoiced${linkCol('Tasks') ? ', with the invoice number' : ''}` : ''}.</span></p>` : ''}
    <label class="check"><input type="checkbox" name="ok" required> All time for this period has been entered</label>
    <div class="actions"><button class="btn" value="cancel" formnovalidate>Cancel</button><button class="btn primary" value="ok">${action === 'print' ? 'Print' : 'Record invoice'}</button></div>`, 'wide');
  return !!data;
}

// Invoices page card: invoiced time without an invoice number, with a choice
// of that client's invoices to assign. You pick; nothing is guessed.
function invoiceCheckCard() {
  if (!linkCol('Tasks')) return '';
  const start = trackingStart();
  if (start === null) return `<div class="card"><h2>Invoice check</h2><p class="small muted">Starts with the first invoice recorded with invoice tracking on. From then on, any invoiced time without an invoice number shows up here.</p></div>`;
  const list = unlinkedTime().sort((a, b) => num(a.Date) - num(b.Date));
  if (!list.length) return `<div class="card"><h2>Invoice check</h2><p class="small muted">✓ All invoiced time since ${longDate(toISO(start))} has an invoice number.</p></div>`;
  return `<div class="card" id="invcheck">
    <h2>Invoice check <span class="pill warn">${list.length}</span></h2>
    <p class="small muted">These entries are marked <b>Invoiced</b> but have no invoice number, usually because they were entered or ticked by hand in the spreadsheet. Pick the invoice each one was billed on, or leave it blank to skip it for now.</p>
    <table class="list"><thead><tr><th>Date</th><th>Project</th><th class="hide-sm">Task</th><th class="r">Hours</th><th>Invoice</th></tr></thead><tbody>
    ${list.map(r => {
      const client = projectRow(r.Project)?.Client;
      const invs = rows('Invoices').filter(i => i.Client === client).sort((a, b) => num(b.Date) - num(a.Date));
      return `<tr><td>${toISO(r.Date)}</td><td>${esc(r.Project)}</td><td class="hide-sm">${esc(r.Task)}</td><td class="r">${hrs(r.Hours)}</td>
        <td><select data-fix="${r._row}"><option value="">—</option>${invs.map(i => `<option value="${esc(i.Number)}">${esc(i.Number)} · ${toISO(i.Date)}</option>`).join('')}</select></td></tr>`;
    }).join('')}</tbody></table>
    <div class="actions"><button class="btn primary" id="fixsave">Save invoice numbers</button></div>
  </div>`;
}
function wireInvoiceCheck(view) {
  if (!$('#fixsave', view)) return;
  $('#fixsave', view).onclick = async () => {
    const items = $$('select[data-fix]', view).filter(s => s.value).map(s => ({ row: rows('Tasks').find(r => r._row === Number(s.dataset.fix)), value: s.value })).filter(x => x.row);
    if (!items.length) return toast('Pick an invoice for at least one entry.', 'info');
    await busy('Saving invoice numbers…', async () => {
      await store.setCells('Tasks', linkCol('Tasks'), items);
      logActivity('Assigned invoice numbers', `Tasks ×${items.length}`, [...new Set(items.map(i => i.value))].join(', '));
    });
    toast(`Saved invoice numbers on ${items.length} time ${items.length === 1 ? 'entry' : 'entries'}.`);
    viewInvoices(view);
  };
}

// =====================================================================
// Settings
// =====================================================================
function trackingCard() {
  const name = cfg().invoiceLinkColumn;
  const tables = ['Tasks', 'Disbursements'];
  const missing = tables.filter(t => !linkCol(t));
  return `
    <div class="card" id="track" style="margin-top:1rem">
      <h2>Invoice tracking</h2>
      <p class="small muted">An <b>${esc(name)}</b> column on the Tasks and Disbursements tables records which invoice each row was billed on. Recording an invoice fills it in, "Mark paid" uses it to mark exactly those rows paid, and each invoice's page lists its time so you can send the client a breakdown. Tracking starts with the next invoice you record; earlier invoices are left as they are.</p>
      <ul class="small tight">${tables.map(t => `<li><b>${t}</b> (sheet ${esc(store.tables[t]?.sheet)}): ${linkCol(t) ? '✓ has the column' : 'no column yet'}</li>`).join('')}</ul>
      <div class="actions">
        ${missing.length ? `<button class="btn primary" id="addcol">Add ${esc(name)} column to ${missing.join(' and ')}</button>` : ''}
        ${missing.length ? '' : `<span class="small muted">On. Next billing day: ${weekday(nextBillingDay().iso)}.</span>`}
      </div>
    </div>`;
}
function wireTracking(view) {
  const name = cfg().invoiceLinkColumn;
  const add = $('#addcol', view);
  if (add) add.onclick = async () => {
    const missing = ['Tasks', 'Disbursements'].filter(t => !linkCol(t));
    if (!confirm(`Excel will add an empty column named "${name}" at the right edge of the ${missing.join(' and ')} table(s). Cells directly to the right of a table move over one column; formulas that refer to them update automatically.`)) return;
    await busy('Adding column…', async () => {
      for (const t of missing) {
        await store.addColumn(t, name);
        logActivity('Added column', `${store.tables[t].sheet}: ${t}[${name}]`);
      }
    });
    toast(`Added the ${esc(name)} column. New invoices will fill it in.`, 'ok');
    viewSettings(view);
  };
}

function viewSettings(view) {
  const s = ls.get('settings', {});
  view.innerHTML = `
    <div class="page-head"><h1>Settings</h1></div>
    <div class="grid2">
      <form class="card form" id="sform">
        <h2>Workbook</h2>
        <div class="field"><label for="s-url">Link to the bookkeeping workbook</label>
          <input id="s-url" type="url" value="${esc(s.fileShareUrl ?? window.BOOKS_CONFIG.fileShareUrl)}" placeholder="https://…sharepoint.com/…/Bookkeeping%202026.xlsx" ${DEMO ? 'disabled' : ''}>
          <small>In Excel for the web: <b>Share → Copy link</b>. Change this each year when you start a new workbook.</small></div>
        <div class="field"><label for="s-rf">Receipts folder (next to the workbook)</label>
          <input id="s-rf" value="${esc(s.receiptFolder ?? window.BOOKS_CONFIG.receiptFolder)}"></div>
        <div class="actions"><button class="btn primary" ${DEMO ? 'disabled' : ''}>Save &amp; reconnect</button></div>
      </form>
      <div class="card">
        <h2>Account</h2>
        <p>Signed in as <b>${esc(user?.name)}</b><br><span class="muted">${esc(user?.email)}</span></p>
        ${me ? `<p class="small muted">Matched to worker <b>${esc(me)}</b> in the Workers table.</p>` : ''}
        ${store.file ? `<p>Connected to <a href="${esc(store.file.webUrl)}" target="_blank" rel="noopener">${esc(store.file.name)}</a></p>` : ''}
        <div class="actions"><button class="btn" id="so">Sign out</button><button class="btn" id="clr">Clear activity log</button></div>
        <h2 style="margin-top:1.5rem">How your data is handled</h2>
        <ul class="small muted tight">
          <li>The app talks only to Microsoft 365. There is no other server or database.</li>
          <li>Each entry is written as a new row in a named Excel table. Formula columns are left for Excel to calculate.</li>
          <li>Before editing or deleting a row, the app checks it still matches Excel, so it won't overwrite changes you made in Excel.</li>
          <li>Receipts are read on this device and saved to your own receipts folder.</li>
        </ul>
      </div>
    </div>
    ${store.file ? trackingCard() : ''}`;
  wireTracking(view);
  $('#sform').onsubmit = async (e) => {
    e.preventDefault();
    const next = { fileShareUrl: clean($('#s-url').value), receiptFolder: clean($('#s-rf').value) || 'Receipts' };
    ls.set('settings', next);
    Object.assign(CFG, next); Object.assign(store.cfg, next);
    await connect();
    location.hash = '#/dashboard';
  };
  $('#so').onclick = () => store.logout();
  $('#clr').onclick = () => { ls.set('log', []); toast('Activity log cleared.', 'info'); };
}

boot();
