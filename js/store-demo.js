// In-memory store with made-up sample data, same interface as GraphStore.
// Lets you try the app (?demo) without Microsoft sign-in. Nothing is saved.
import { parseTable, rowAddress } from './store-graph.js?v=2026101502';

const today = new Date();
const serial = (d) => Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 864e5);
const daysAgo = (n) => { const d = new Date(today); d.setDate(d.getDate() - n); return serial(d); };

const HEAD = {
  Tasks: ['Date', 'Project', 'Hours', 'Task', 'Worker', 'Billable', 'Invoiced', 'Paid', 'Short'],
  Expenses: ['Date', 'Vendor', 'Description', 'Category', 'Amount', 'Tax', 'Business', 'Reimburse', 'Paid', 'Scan', 'Billable', 'Bus Total', 'Bus Tax', 'Project'],
  Mileage: ['Date', 'km', 'Destination', 'Reason For Trip', 'EV', 'PD AM', 'PD Noon', 'PD PM', 'PD Inc', 'PD PA', 'PD Paid'],
  Invoices: ['Date', 'Client', 'Number', 'Amount', 'Tax', 'Due Date', 'Paid', 'Days', 'Interest', 'Overdue'],
  Disbursements: ['Date', 'Client', 'Project', 'Description', 'Vendor', 'Cost', 'Markup', 'Markup  $', 'Billable', 'Invoiced', 'Paid'],
  Projects: ['Name', 'ProjectID', 'Client', 'ClientID', 'Status', 'Type', 'Prime', 'Rate', 'Fee', '%Comp', 'Notes', 'Hours', 'This Invoice', 'Billed', 'Total', 'Remaining', 'Overage', 'Total Hours', 'Short', 'Carryover'],
  Clients: ['Name', 'Contact', 'email', 'Phone', 'Address', 'Notes', 'Short'],
  Workers: ['Name', 'email', 'Address', 'Phone', 'Start Date', 'Hours'],
  Categories: ['CRA Categories', 'Total', 'Bus Total', 'Examples'],
  Status: ['Status'],
  Table12: ['Type'],
};
const SHEET = { Tasks: 'Time', Categories: 'Expenses', Status: 'Projects', Table12: 'Projects' };
const RANGES = {
  'Mileage!Q22:R26': [['Breakfast', 29.05], ['Lunch', 29.6], ['Dinner', 60.75], ['Incedental', 17.5], ['Private Acc', 50]],
  'Mileage!T7:V8': [[0.73, 'First', 5000], [0.67, 'After', '']],
};
const effTax = r => (r[5] === '' || r[5] == null || String(r[5]).startsWith('=')) ? r[4] * 13 / 113 : Number(r[5]) || 0;

function seed() {
  const W = 'Demo User';
  return {
    Clients: [
      ['Internal', 'Demo User', 'demo@example.com', '555-0100', '100 Main St, Anytown, ON A1A 1A1', '', 'INT'],
      ['Northwind Utilities', 'Pat Lee', 'ap@northwind.example', '555-0101', '200 Grid Rd, Hamilton, ON L8L 1A1', 'PO# 4471', 'NW'],
      ['Contoso Health', 'Sam Diaz', 'billing@contoso.example', '555-0102', '9 Care Way, London, ON N6A 1A1', '', 'CH'],
    ],
    Workers: [[W, 'demo@example.com', '', '', 45000, 0], ['Alex Helper', 'alex@example.com', '', '', 45500, 0]],
    Projects: [
      ['Internal', 'INT-001', 'Internal', 'None', 'Active', 'Internal', W, '', ''],
      ['Training', 'INT-002', 'Internal', 'None', 'Active', 'Internal', W, '', ''],
      ['Substation Upgrade', 'NW-001', 'Northwind Utilities', '2026-0101-10', 'Active', 'Hourly', W, 195, ''],
      ['Feeder Study', 'NW-002', 'Northwind Utilities', '2026-0144-10', 'Active', 'Fixed Fee', W, 200, 12000],
      ['Hospital Generator', 'CH-001', 'Contoso Health', 'PO-7781', 'Active*', 'Fixed Fee', W, 210, 18000],
      ['Arc Flash Review', 'CH-002', 'Contoso Health', '', 'Opportunity', 'Fixed Fee', W, 200, 9000],
    ],
    Status: ['Active', 'Active*', 'Opportunity', 'Lost', 'Hold', 'Archive'].map(x => [x]),
    Table12: ['Internal', 'Retainer', 'Hourly', 'Fixed Fee'].map(x => [x]),
    Categories: ['Home Office Expenses', 'Motor Vehicle Expenses', 'Business Taxes, Licenses, Dues', 'Office Expenses', 'Supplies', 'Meals and Entertainment', 'Travel', 'Telephone and Utilities', 'Training and Education'].map(c => [c, 0, 0, '']),
    Tasks: [
      [daysAgo(85), 'Substation Upgrade', 12, 'Site survey and existing conditions', W, true, true, true],
      [daysAgo(78), 'Substation Upgrade', 8, 'Preliminary single line diagram', W, true, true, true],
      [daysAgo(40), 'Feeder Study', 6, 'Load flow model', W, true, true, false],
      [daysAgo(12), 'Substation Upgrade', 3.5, 'Single line diagram revisions', W, true, false, false],
      [daysAgo(9), 'Feeder Study', 5, 'Short circuit study', W, true, false, false],
      [daysAgo(6), 'Hospital Generator', 4, 'Site visit and ATS review', W, true, false, false],
      [daysAgo(5), 'Internal', 1.5, 'Bookkeeping', 'Alex Helper', false, false, false],
      [daysAgo(2), 'Substation Upgrade', 2, 'Protection coordination', W, true, false, false],
      [daysAgo(1), 'Feeder Study', 3, 'Report writing', W, true, false, false],
      [daysAgo(52), 'Arc Flash Review', 4, 'Proposal and site walk', W, false, false, false],
      [daysAgo(33), 'Arc Flash Review', 2.5, 'Scope clarifications', W, false, false, false],
    ].map(r => [...r, '']),
    Expenses: [
      [daysAgo(20), 'Staples', 'Printer paper', 'Office Expenses', 22.6, '', 1, true, false, true, false, 0, 0, ''],
      [daysAgo(8), 'Esso', 'Fuel', 'Motor Vehicle Expenses', 80, '', 1, true, false, true, false, 0, 0, ''],
      [daysAgo(15), 'Rogers', 'Internet', 'Telephone and Utilities', 90.39, '', 0.5, true, false, false, false, 0, 0, ''],
      [daysAgo(6), 'Hilton', 'Hotel, generator inspection', 'Travel', 189.5, '', 1, true, false, true, false, 0, 0, 'Hospital Generator'],
      [daysAgo(3), 'IEEE', 'Membership', 'Business Taxes, Licenses, Dues', 254.25, 0, 1, false, true, false, false, 0, 0, ''],
    ],
    Mileage: [
      [daysAgo(12), 84, 'Northwind Substation', 'Site visit', true, false, true, false, false, false, false],
      [daysAgo(6), 132, 'Contoso Health', 'Generator inspection', true, true, true, true, true, false, false],
    ],
    Invoices: [
      [daysAgo(70), 'Northwind Utilities', 'NW2607', 4407, 0, 0, daysAgo(48), 0, '', ''],
      [daysAgo(38), 'Northwind Utilities', 'NW2608', 1356, 0, 0, '', 0, '', ''],
    ],
    Disbursements: [
      [daysAgo(10), 'Contoso Health', 'Hospital Generator', 'Permit drawings print', 'PrintCo', 150, 0.1, 0, 0, false, false],
    ],
  };
}

export class DemoStore {
  constructor(cfg, schema) { this.cfg = cfg; this.schema = schema; this.tables = {}; this.data = seed(); }
  async init() { return { name: 'Demo User', email: 'demo@example.com' }; }
  login() {} logout() { location.href = location.pathname; }
  async openFile() {
    this.file = { name: 'Bookkeeping DEMO.xlsx', webUrl: '#' };
    // Pretend we are the company itself
    this.cfg = { ...this.cfg, companyName: 'Your Company Inc.', taxNumber: '000000000-RT0001', allowedUsers: [] };
    return this.file;
  }

  recalc() {
    const d = this.data;
    for (const r of d.Tasks) r[8] = d.Clients.find(c => c[0] === (d.Projects.find(p => p[0] === r[1]) || [])[2])?.[6] || '';
    for (const r of d.Expenses) { r[11] = r[4] * r[6]; r[12] = effTax(r) * r[6]; }
    for (const r of d.Invoices) { r[4] = r[3] * 13 / 113; r[5] = r[0] + 30; r[7] = (r[6] ? r[6] : serial(today)) - r[0]; }
    for (const r of d.Disbursements) { r[7] = r[5] * r[6]; r[8] = r[5] + r[7]; }
    d.Projects = d.Projects.map(p => {
      const [name, , , , , type, , rate, fee] = p;
      const rt = Number(rate) || 0, f = Number(fee) || 0;
      const tasks = d.Tasks.filter(t => t[1] === name && t[5]);
      const hrs = tasks.filter(t => !t[6]).reduce((s, t) => s + t[2], 0);
      const billedH = tasks.filter(t => t[6]).reduce((s, t) => s + t[2], 0);
      const billed = f ? Math.min(billedH * rt, f) : billedH * rt;
      const total = hrs * rt + billed;
      const comp = type === 'Fixed Fee' ? Math.min(1, f ? total / f : 0) : '';
      const thisInv = comp === 1 ? f - billed : hrs * rt;
      const remaining = type === 'Fixed Fee' && comp < 1 ? f - billed - thisInv : 0;
      const totalH = d.Tasks.filter(t => t[1] === name).reduce((s, t) => s + t[2], 0);
      return [...p.slice(0, 9), comp, p[10] ?? '', hrs, thisInv, billed, total, remaining, 0, totalH, '', p[19] ?? ''];
    });
    for (const w of d.Workers) w[5] = d.Tasks.filter(t => t[4] === w[0]).reduce((s, t) => s + t[2], 0);
  }

  async loadTable(name) {
    this.recalc();
    // Expenses[Tax] holds the formula (or a custom amount); show its result like Excel does
    const out = r => HEAD[name].map((_, i) => (name === 'Expenses' && i === 5 ? effTax(r) : r[i] ?? ''));
    const t = parseTable(`${SHEET[name] || name}!A1:A1`, [HEAD[name], ...this.data[name].map(out)]);
    this.tables[name] = t;
    return t;
  }
  async loadAll(names) { for (const n of names) await this.loadTable(n); }

  async addRow(name, obj) {
    this.data[name].push(HEAD[name].map(h => (h in obj && obj[h] !== null ? obj[h] : '')));
    const t = await this.loadTable(name);
    const row = t.rows[t.rows.length - 1];
    return { sheet: t.sheet, row: row._row, address: `${t.sheet}!${rowAddress(t, row._row)} (demo)` };
  }
  async updateRow(name, row, changes) {
    const r = this.data[name][row._i];
    HEAD[name].forEach((h, i) => { if (h in changes) r[i] = changes[h]; });
    const t = await this.loadTable(name);
    return { address: `${t.sheet}!${rowAddress(t, row._row)} (demo)` };
  }
  async deleteRow(name, row) {
    this.data[name].splice(row._i, 1);
    const t = await this.loadTable(name);
    return { address: `${t.sheet}!${rowAddress(t, row._row)} (demo)` };
  }
  async setCells(name, header, items) {
    const i = HEAD[name].indexOf(header);
    for (const { row, value } of items) this.data[name][row._i][i] = value;
    await this.loadTable(name);
  }
  setColumn(name, rowsToSet, header, value) { return this.setCells(name, header, rowsToSet.map(row => ({ row, value }))); }
  async addColumn(name, header) {
    if (HEAD[name].includes(header)) return;
    HEAD[name].push(header);
    this.data[name].forEach(r => { r.length = HEAD[name].length - 1; r.push(''); });
    await this.loadTable(name);
  }
  async readRange(ref) { return RANGES[ref] || []; }
  async uploadFile(blob, fileName) { return { name: fileName, webUrl: URL.createObjectURL(blob) }; }
}
