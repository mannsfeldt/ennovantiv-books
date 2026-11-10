// End-to-end tests: drive the real app in demo mode (sample data, no sign-in)
// inside an iframe and check what it writes to the demo workbook.
import { test, assert, equal, until, sleep } from './harness.js';

// A fresh copy of the app for each test, so sample data never leaks between tests.
export async function openApp(hash = '/dashboard') {
  for (const k of Object.keys(localStorage)) if (/^books:(pendingInvoice|printed:)/.test(k)) localStorage.removeItem(k);
  const frame = document.createElement('iframe');
  frame.style.cssText = 'width:1200px;height:900px;border:0;position:absolute;left:-9999px;top:0';
  frame.src = `../index.html?demo&t=${Date.now()}#${hash}`;
  document.body.appendChild(frame);
  const win = await until(() => frame.contentWindow?.demoStore?.tables?.Tasks && frame.contentDocument.querySelector('#view')?.children.length && frame.contentWindow, 10000, 'the app to load');
  const doc = win.document, store = win.demoStore;
  const app = {
    win, doc, store,
    confirms: [], answer: true,                     // every confirm() is recorded; answer decides OK/Cancel
    $: s => doc.querySelector(s),
    $$: s => [...doc.querySelectorAll(s)],
    rows: name => store.tables[name].rows,
    async go(h) { win.location.hash = h; await sleep(60); },
    form: () => doc.querySelector('#entry'),
    set(name, value) {                              // fill a form field the way a person would
      const el = doc.querySelector('#entry').elements[name];
      if (el.type === 'checkbox') el.checked = value; else el.value = value;
      el.dispatchEvent(new win.Event('input', { bubbles: true }));
      el.dispatchEvent(new win.Event('change', { bubbles: true }));
    },
    async save() { doc.querySelector('#entry').requestSubmit(); await app.idle(); },
    async idle() { await sleep(30); await until(() => doc.querySelector('#busy')?.hidden !== false, 8000, 'the app to finish saving'); await sleep(30); },
    async edit(table, row) { doc.querySelector(`#listwrap tr[data-row="${row._row}"]`).click(); await sleep(40); },
    async showAll() { const p = doc.querySelector('#period'); p.value = 'year'; p.dispatchEvent(new win.Event('change')); await sleep(30); },
    // change a cell in the demo workbook directly, as if it had been edited in Excel
    async poke(table, rowIndex, column, value) { store.data[table][rowIndex][store.tables[table].headers.indexOf(column)] = value; await store.loadTable(table); },
    toasts: () => [...doc.querySelectorAll('.toast')].map(t => t.textContent),
    close() { frame.remove(); },
  };
  win.confirm = msg => { app.confirms.push(String(msg)); return app.answer; };
  win.print = () => { app.printed = (app.printed || 0) + 1; };
  return app;
}
// run a test with an app that is always cleaned up
export const withApp = (hash, fn) => async () => { const app = await openApp(hash); try { await fn(app); } finally { app.close(); } };

test('app: logging time adds one row to the Tasks table', withApp('/time', async app => {
  const before = app.rows('Tasks').length;
  app.set('Project', 'Substation Upgrade');
  app.set('Hours', '2.5');
  app.set('Task', 'Cable sizing');
  await app.save();
  equal(app.rows('Tasks').length, before + 1, 'row count');
  const row = app.rows('Tasks').at(-1);
  equal([row.Project, row.Hours, row.Task, row.Billable, row.Invoiced], ['Substation Upgrade', 2.5, 'Cable sizing', true, false]);
}));

test('A-01: editing an expense keeps a fractional Business % (25.8%)', withApp('/time', async app => {
  await app.poke('Expenses', 0, 'Business', 0.258);         // a home-office expense
  await app.go('/expenses'); await app.showAll();
  await app.edit('Expenses', app.rows('Expenses')[0]);
  equal(app.form().elements.Business.value, '25.8', 'form shows the exact percentage');
  app.set('Description', 'typo fixed');
  await app.save();
  const row = app.rows('Expenses')[0];
  equal([row.Description, row.Business], ['typo fixed', 0.258], 'Business is unchanged after an unrelated edit');
}));

const fillTime = app => { app.set('Project', 'Substation Upgrade'); app.set('Hours', '1.5'); app.set('Task', 'Site call'); };

test('A-04: submitting the form twice saves the entry once', withApp('/time', async app => {
  const before = app.rows('Tasks').length;
  fillTime(app);
  app.form().requestSubmit(); app.form().requestSubmit();     // Enter pressed twice
  await app.idle();
  equal(app.store.data.Tasks.length, before + 1, 'rows in the workbook');
}));

test('A-04: when the row is saved but cannot be re-read, the app says saved, not failed', withApp('/time', async app => {
  const before = app.store.data.Tasks.length, load = app.store.loadTable.bind(app.store);
  let failed = false;
  app.store.loadTable = async name => { if (name === 'Tasks' && !failed) { failed = true; throw new Error('Failed to fetch'); } return load(name); };
  fillTime(app);
  await app.save();
  equal(app.store.data.Tasks.length, before + 1, 'rows in the workbook');
  const said = app.toasts().join(' | ');
  assert(/Saved/.test(said) && /again/i.test(said), `toast should say it was saved and not to save again; got: ${said}`);
  assert(!/⚠ Failed to fetch/.test(said), 'no bare failure message');
}));

test('A-02: "Make billable" stops if rows moved since the list was shown', withApp('/time', async app => {
  await app.showAll();
  const lf = app.$('#lf'); lf.value = 'Arc Flash Review'; lf.dispatchEvent(new app.win.Event('change')); await sleep(30);
  app.store.data.Tasks.splice(0, 1);                    // meanwhile someone deletes an earlier entry in Excel
  const before = JSON.stringify(app.store.data.Tasks);
  app.$('#bulkgo').click();
  await app.idle();
  equal(JSON.stringify(app.store.data.Tasks), before, 'nothing was written');
  assert(/changed in Excel/.test(app.toasts().join(' ')), 'the app explains why it stopped');
}));

// ---- invoice helpers ----
const CLIENT = 'Northwind Utilities';
const addTracking = async app => { await app.store.addColumn('Tasks', 'Invoice'); await app.store.addColumn('Disbursements', 'Invoice'); };
async function previewInvoice(app, client = CLIENT) {
  await app.go('/dashboard'); await app.go('/invoices');
  const s = app.$('#ic'); s.value = client; s.dispatchEvent(new app.win.Event('change')); await sleep(30);
  return app.$('#invprev .inv-title dd b').textContent;          // the invoice number
}
// tick every required box in the open dialog and confirm it
async function confirmDialog(app) {
  const dlg = await until(() => app.$('dialog.modal'), 6000, 'the dialog');
  dlg.querySelectorAll('input[type=checkbox][required]').forEach(c => { c.checked = true; });
  const text = dlg.textContent;
  dlg.querySelector('button[value=ok]').click();
  await app.idle();
  return text;
}
async function record(app, client = CLIENT) {
  const number = await previewInvoice(app, client);
  app.$('#rec').click();
  await confirmDialog(app);
  return number;
}
const invoiceRows = (app, number) => app.rows('Invoices').filter(r => r.Number === number);
const clientTasks = app => app.rows('Tasks').filter(r => ['Substation Upgrade', 'Feeder Study'].includes(r.Project) && r.Billable === true);

test('app: recording an invoice adds one Invoices row and ticks and numbers its time entries', withApp('/dashboard', async app => {
  await addTracking(app);
  const due = clientTasks(app).filter(r => r.Invoiced !== true).length;
  assert(due > 0, 'sample data has time to invoice');
  const number = await record(app);
  equal(invoiceRows(app, number).length, 1, 'Invoices rows for this number');
  equal(clientTasks(app).filter(r => r.Invoiced === true && r.Invoice === number).length, due, 'entries ticked and numbered');
  equal(clientTasks(app).filter(r => r.Invoiced !== true).length, 0, 'entries left unbilled');
}));

// make the next bulk write fail, optionally after writing the first `partial` rows
function failNextBulkWrite(app, partial = 0) {
  const real = app.store.setRows.bind(app.store);
  let armed = true;
  app.store.setRows = async (name, items, onProgress) => {
    if (!armed) return real(name, items, onProgress);
    armed = false;
    if (partial) await real(name, items.slice(0, partial));
    throw new Error('Failed to fetch');
  };
}
const pending = () => JSON.parse(localStorage.getItem('books:pendingInvoice') || 'null');

test('A-03: an interrupted recording can be resumed, and ends complete with one Invoices row', withApp('/dashboard', async app => {
  await addTracking(app);
  const due = clientTasks(app).filter(r => r.Invoiced !== true).length;
  failNextBulkWrite(app);                                   // connection drops after the Invoices row is added
  const number = await record(app);
  equal(invoiceRows(app, number).length, 1, 'Invoices row was added');
  equal(clientTasks(app).filter(r => r.Invoiced === true && r.Invoice === number).length, 0, 'no time ticked yet');
  assert(pending()?.number === number, 'the interrupted invoice is remembered');
  const resume = await until(() => app.$('#resume'), 3000, 'a Resume button on the Invoices page');
  resume.click(); await app.idle();
  equal(invoiceRows(app, number).length, 1, 'still exactly one Invoices row');
  equal(clientTasks(app).filter(r => r.Invoiced === true && r.Invoice === number).length, due, 'all entries ticked and numbered');
  equal(pending(), null, 'nothing left pending');
  assert(!app.$('#resume'), 'Resume button gone');
}));

test('A-03: resuming after some entries were already ticked only does the rest', withApp('/dashboard', async app => {
  await addTracking(app);
  const due = clientTasks(app).filter(r => r.Invoiced !== true).length;
  failNextBulkWrite(app, 1);                                // one entry ticked, then the connection drops
  const number = await record(app);
  equal(clientTasks(app).filter(r => r.Invoiced === true).length - clientTasks(app).filter(r => r.Invoiced === true && r.Invoice !== number).length, 1, 'one entry ticked before the failure');
  (await until(() => app.$('#resume'), 3000, 'a Resume button')).click(); await app.idle();
  equal(clientTasks(app).filter(r => r.Invoiced === true && r.Invoice === number).length, due, 'all entries ticked and numbered');
  equal(invoiceRows(app, number).length, 1, 'exactly one Invoices row');
}));

test('A-03: recording an invoice number that already exists asks first', withApp('/dashboard', async app => {
  await addTracking(app);
  const number = await record(app);
  await app.store.addRow('Tasks', { Date: app.rows('Tasks').at(-1).Date, Project: 'Substation Upgrade', Hours: 1, Task: 'late entry', Worker: 'Demo User', Billable: true, Invoiced: false, Paid: false });
  await app.store.loadTable('Projects');                    // the demo recalculates project totals on load
  equal(await previewInvoice(app), number, 'same month, same number');
  app.confirms.length = 0; app.answer = false;              // decline
  app.$('#rec').click(); await app.idle(); await sleep(100);
  assert(app.confirms.some(c => /already/i.test(c)), `asked before recording a second ${number}; confirms: ${JSON.stringify(app.confirms)}`);
  app.$('dialog.modal')?.querySelector('button[value=cancel]')?.click();
  equal(invoiceRows(app, number).length, 1, 'no second Invoices row');
}));

test('A-03: if marking time paid fails, the invoice stays unpaid so Mark paid can be retried', withApp('/dashboard', async app => {
  await addTracking(app);
  const number = await record(app);
  const inv = () => invoiceRows(app, number)[0];
  const paidTasks = () => clientTasks(app).filter(r => r.Invoice === number && r.Paid === true).length;
  const total = clientTasks(app).filter(r => r.Invoice === number).length;
  failNextBulkWrite(app);
  app.$(`[data-paid="${inv()._row}"]`).click(); await confirmDialog(app);
  assert(!inv().Paid, 'invoice is not marked paid when its time entries could not be');
  equal(paidTasks(), 0);
  app.$(`[data-paid="${inv()._row}"]`).click(); await confirmDialog(app);     // retry
  assert(inv().Paid, 'invoice paid after the retry');
  equal(paidTasks(), total, 'all its time entries paid');
}));

test('A-03: two identical time entries are both ticked after a resumed recording', withApp('/dashboard', async app => {
  await addTracking(app);
  const twin = { Date: app.rows('Tasks').at(-1).Date, Project: 'Substation Upgrade', Hours: 2, Task: 'site visit', Worker: 'Demo User', Billable: true, Invoiced: false, Paid: false };
  await app.store.addRow('Tasks', twin); await app.store.addRow('Tasks', twin);
  await app.store.loadTable('Projects');
  const due = clientTasks(app).filter(r => r.Invoiced !== true).length;
  failNextBulkWrite(app, 1);
  const number = await record(app);
  (await until(() => app.$('#resume'), 3000, 'a Resume button')).click(); await app.idle();
  equal(clientTasks(app).filter(r => r.Invoiced === true && r.Invoice === number).length, due, 'every entry ticked, including both twins');
  equal(clientTasks(app).filter(r => r.Task === 'site visit' && r.Invoice === number).length, 2, 'both identical entries carry the number');
}));

async function printInvoice(app) {
  const number = await previewInvoice(app);
  app.$('#print').click();
  await confirmDialog(app);
  return number;
}
async function openRecordDialog(app) {
  await previewInvoice(app);
  app.$('#rec').click();
  return until(() => app.$('dialog.modal'), 6000, 'the record checklist');
}

test('A-06: recording warns, and needs an extra tick, when the invoice changed since it was printed', withApp('/dashboard', async app => {
  await addTracking(app);
  await printInvoice(app);
  equal(app.printed, 1, 'printed once');
  // a billable entry is added after the PDF went to the client
  await app.store.addRow('Tasks', { Date: app.rows('Tasks').at(-1).Date, Project: 'Substation Upgrade', Hours: 3, Task: 'added after printing', Worker: 'Demo User', Billable: true, Invoiced: false, Paid: false });
  await app.store.loadTable('Projects');
  const dlg = await openRecordDialog(app);
  assert(/not the invoice you printed/i.test(dlg.textContent), 'the checklist says the invoice changed since printing');
  assert(/1 (time )?entr(y|ies) added/i.test(dlg.textContent), 'it says what changed');
  assert(dlg.querySelector('input[name=changed][required]'), 'recording the new amount needs its own tick');
  dlg.querySelector('button[value=cancel]').click();
}));

test('A-06: no warning when the invoice is the same as the one printed', withApp('/dashboard', async app => {
  await addTracking(app);
  await printInvoice(app);
  const dlg = await openRecordDialog(app);
  assert(!/not the invoice you printed/i.test(dlg.textContent), 'no change warning');
  assert(!dlg.querySelector('input[name=changed]'), 'no extra tick');
  dlg.querySelector('button[value=cancel]').click();
}));

const invoicedEntry = app => app.rows('Tasks').find(r => r.Invoiced === true);
const sameRow = (app, row) => app.rows('Tasks').find(r => r._row === row._row);

test('A-07: changing the hours of an invoiced time entry asks first, and leaves it alone if declined', withApp('/time', async app => {
  await app.showAll();
  const row = invoicedEntry(app);
  await app.edit('Tasks', row);
  app.set('Hours', String(row.Hours + 1));
  app.confirms.length = 0; app.answer = false;
  await app.save();
  assert(app.confirms.some(c => /billed|invoice/i.test(c)), `asked before changing billed hours; confirms: ${JSON.stringify(app.confirms)}`);
  equal(sameRow(app, row).Hours, row.Hours, 'hours unchanged');
}));

test('A-07: correcting only the description of an invoiced entry does not ask', withApp('/time', async app => {
  await app.showAll();
  const row = invoicedEntry(app);
  await app.edit('Tasks', row);
  app.set('Task', 'typo fixed');
  app.confirms.length = 0;
  await app.save();
  equal(app.confirms, [], 'no confirmation');
  equal(sameRow(app, row).Task, 'typo fixed');
}));

test('A-07: deleting an invoiced time entry warns that it was billed', withApp('/time', async app => {
  await app.showAll();
  const row = invoicedEntry(app), before = app.rows('Tasks').length;
  await app.edit('Tasks', row);
  app.confirms.length = 0; app.answer = false;
  app.$('#del').click(); await app.idle();
  assert(/billed|invoice/i.test(app.confirms.join(' ')), `delete warning mentions the invoice; got: ${JSON.stringify(app.confirms)}`);
  equal(app.rows('Tasks').length, before, 'row kept');
}));

test('A-10: 75 hours in one time entry asks before saving', withApp('/time', async app => {
  const before = app.rows('Tasks').length;
  fillTime(app); app.set('Hours', '75');                    // meant 7.5
  app.answer = false;
  await app.save();
  assert(app.confirms.some(c => /75/.test(c) && /hours/i.test(c)), `asked about 75 hours; confirms: ${JSON.stringify(app.confirms)}`);
  equal(app.rows('Tasks').length, before, 'not saved when declined');
}));

test('A-10: zero hours asks before saving', withApp('/time', async app => {
  const before = app.rows('Tasks').length;
  fillTime(app); app.set('Hours', '0');
  app.answer = false;
  await app.save();
  assert(app.confirms.length === 1, 'asked once');
  equal(app.rows('Tasks').length, before);
}));

test('A-10: an ordinary time entry saves without any question', withApp('/time', async app => {
  fillTime(app);
  await app.save();
  equal(app.confirms, []);
}));

test('A-10: Business use of 150% asks before saving an expense', withApp('/expenses', async app => {
  const before = app.rows('Expenses').length;
  app.set('Vendor', 'Staples'); app.set('Category', 'Office Expenses'); app.set('Amount', '20'); app.set('Business', '150');
  app.answer = false;
  await app.save();
  assert(app.confirms.some(c => /150/.test(c)), `asked about 150%; confirms: ${JSON.stringify(app.confirms)}`);
  equal(app.rows('Expenses').length, before);
}));

test('A-10: a billable project cannot be added without a rate', withApp('/projects', async app => {
  const before = app.rows('Projects').length, f = app.$('#pform').elements;
  f.Name.value = 'No Rate Job'; f.Client.value = CLIENT; f.Client.dispatchEvent(new app.win.Event('change'));
  f.Type.value = 'Hourly'; f.Rate.value = '';
  app.$('#pform').requestSubmit(); await app.idle();
  equal(app.rows('Projects').length, before, 'project not added');
  assert(/rate/i.test(app.toasts().join(' ')), 'says a rate is needed');
}));

test('A-10: a payment date before the invoice date asks first', withApp('/dashboard', async app => {
  await addTracking(app);
  const number = await record(app);
  const inv = () => invoiceRows(app, number)[0];
  app.$(`[data-paid="${inv()._row}"]`).click();
  const dlg = await until(() => app.$('dialog.modal'), 6000, 'the Mark paid dialog');
  dlg.querySelector('input[name=date]').value = '2020-01-01';
  app.confirms.length = 0; app.answer = false;
  dlg.querySelector('button[value=ok]').click(); await app.idle();
  assert(app.confirms.some(c => /before/i.test(c)), `asked about the early date; confirms: ${JSON.stringify(app.confirms)}`);
  assert(!inv().Paid, 'not marked paid when declined');
}));

test('A-10: the checklist flags an hourly project that has hours but no rate', withApp('/dashboard', async app => {
  const i = app.rows('Projects').findIndex(p => p.Name === 'Substation Upgrade');
  await app.poke('Projects', i, 'Rate', '');
  const dlg = await openRecordDialog(app);
  assert(/no rate/i.test(dlg.textContent) && /Substation Upgrade/.test(dlg.textContent), 'checklist names the project with no rate');
  dlg.querySelector('button[value=cancel]').click();
}));

test('A-09: an entry dated in another year asks before saving into this workbook', withApp('/time', async app => {
  const before = app.rows('Tasks').length, next = new Date().getFullYear() + 1;
  fillTime(app); app.set('Date', `${next}-01-05`);
  app.answer = false;
  await app.save();
  assert(app.confirms.some(c => c.includes(String(next)) && /workbook/i.test(c)), `asked about the year; confirms: ${JSON.stringify(app.confirms)}`);
  equal(app.rows('Tasks').length, before, 'not saved when declined');
}));

test('A-09: the workbook year comes from the file name, so last year\'s file is caught in January', withApp('/time', async app => {
  app.store.file.name = 'Bookkeeping 2031.xlsx';            // device still linked to another year's workbook
  const before = app.rows('Tasks').length;
  fillTime(app);                                            // dated today
  app.answer = false;
  await app.save();
  assert(app.confirms.some(c => /2031/.test(c)), `asked; confirms: ${JSON.stringify(app.confirms)}`);
  equal(app.rows('Tasks').length, before);
}));

test('A-09: the dashboard warns when the connected workbook is for another year', withApp('/time', async app => {
  app.store.file.name = 'Bookkeeping 2031.xlsx';
  await app.go('/dashboard');
  assert(/2031/.test(app.$('#view .warnbox')?.textContent || ''), 'warning names the workbook year');
}));

test('A-09: entry screens name the workbook they write to', withApp('/time', async app => {
  assert(app.$('#entry .dest').textContent.includes(app.store.file.name), 'workbook name shown on the form');
}));
