// Starting a new year (2026 -> 2027): the rules, then the whole flow in the demo.
import { test, assert, equal, until, sleep } from './harness.js';
import { openApp, withApp } from './app.test.js';
import { carryoverFor, planProjects, newFileName, carryoverChanges } from '../js/rollover.js';

// ---- the rules ----
test('rollover rules: a fixed-fee project carries what it has spent (billed + not yet invoiced + billed beyond the fee)', () => {
  equal(carryoverFor({ Type: 'Fixed Fee', Fee: 15000, Total: 13668.75, Overage: 0 }), 13668.75);
  equal(carryoverFor({ Type: 'Fixed Fee', Fee: 15250, Total: 15587.5, Overage: 162.5 }), 15750, 'Overage is added: Total stops at the fee');
  equal(carryoverFor({ Type: 'Fixed Fee', Fee: 100, Total: 0.1, Overage: 0.2 }), 0.3, 'no floating point crumbs');
});

test('rollover rules: hourly, retainer and internal projects, and fixed fees with nothing spent, carry nothing', () => {
  equal(carryoverFor({ Type: 'Hourly', Fee: '', Total: 5000, Overage: 0 }), '');
  equal(carryoverFor({ Type: 'Retainer', Fee: '', Total: 1400, Overage: 0 }), '');
  equal(carryoverFor({ Type: 'Fixed Fee', Fee: 9000, Total: 0, Overage: 0 }), '');
  equal(carryoverFor({ Type: 'Fixed Fee', Fee: '', Total: 10, Overage: 0 }), '', 'no fee: nothing to carry against');
});

test('rollover rules: Lost and Archive projects are not carried by default', () => {
  const plan = planProjects([{ Name: 'A', Status: 'Active', Type: 'Hourly' }, { Name: 'B', Status: 'Lost', Type: 'Fixed Fee' }, { Name: 'C', Status: 'Opportunity', Type: 'Fixed Fee' }, { Name: 'D', Status: 'Archive', Type: 'Hourly' }], ['Active', 'Active*', 'Opportunity', 'Hold']);
  equal(plan.map(p => [p.name, p.include]), [['A', true], ['B', false], ['C', true], ['D', false]]);
});

test('rollover rules: the new file name replaces the year, or adds one', () => {
  equal(newFileName('Bookkeeping 2026.xlsx', 2027), 'Bookkeeping 2027.xlsx');
  equal(newFileName('Books.xlsx', 2027), 'Books 2027.xlsx');
});

test('rollover rules: carryover changes list only projects whose carryover differs, and those missing from the target', () => {
  const source = [{ Name: 'F1', Type: 'Fixed Fee', Fee: 100, Total: 100, Overage: 20 }, { Name: 'F2', Type: 'Fixed Fee', Fee: 50, Total: 25, Overage: 0 }, { Name: 'F3', Type: 'Fixed Fee', Fee: 50, Total: 5, Overage: 0 }, { Name: 'H', Type: 'Hourly', Total: 900 }];
  const target = [{ Name: 'F1', Carryover: 100 }, { Name: 'F2', Carryover: 25 }, { Name: 'H', Carryover: '' }];
  const { changes, missing } = carryoverChanges(source, target);
  equal(changes.map(c => [c.name, c.from, c.to]), [['F1', 100, 120]]);
  equal(missing, ['F3']);
});

// ---- the flow, in the demo ----
const NEW_NAME = 'Bookkeeping 2027 DEMO.xlsx';
const files = app => app.store.constructor.files;
const newData = app => files(app).get(NEW_NAME)?.data;
const spentOf = (app, name) => { const p = app.rows('Projects').find(r => r.Name === name); return Math.round((Number(p.Total) + Number(p.Overage)) * 100) / 100; };
const cad = n => n.toLocaleString('en-CA', { style: 'currency', currency: 'CAD' });
// blank = nothing typed in the first columns (formula columns such as Expenses' tax legitimately show 0)
const isBlankRow = r => r.slice(0, 4).every(v => v === '' || v == null);

async function openPreview(app) {
  await app.go('/settings');
  app.$('#startroll').click();
  return until(() => app.$('dialog.modal'), 8000, 'the new-year preview');
}
async function createNewYear(app, dlg) {
  dlg.querySelector('input[name=ok]').checked = true;
  dlg.querySelector('button[value=ok]').click();
  return until(() => { const d = app.$('dialog.modal'); return d && /is ready/i.test(d.textContent) ? d : null; }, 12000, 'the finished dialog');
}
const closeResult = (app, switchNow = false) => { const d = app.$('dialog.modal'); d.querySelector('input[name=switch]').checked = switchNow; d.querySelector('button[value=ok]').click(); };

test('new year: the preview shows what will carry, with carryover for fixed-fee projects and none for hourly ones', withApp('/settings', async app => {
  await app.store.addRow('Projects', { Name: 'Old Lost Job', ProjectID: 'X-1', Client: 'Northwind Utilities', Status: 'Lost', Type: 'Fixed Fee', Rate: 200, Fee: 5000 });
  await app.store.loadTable('Projects');
  const dlg = await openPreview(app);
  assert(/Bookkeeping 2027 DEMO\.xlsx/.test(dlg.textContent), 'names the new file');
  assert(/is not changed/i.test(dlg.textContent), 'promises the current workbook is not changed');
  equal(dlg.querySelector('[data-carry="Feeder Study"]').textContent.trim(), cad(spentOf(app, 'Feeder Study')), 'fixed fee carries what was spent');
  equal(dlg.querySelector('[data-carry="Substation Upgrade"]').textContent.trim(), '—', 'hourly carries nothing');
  assert(dlg.querySelector('input[name=proj][value="Feeder Study"]').checked, 'active projects are ticked');
  assert(!dlg.querySelector('input[name=proj][value="Old Lost Job"]').checked, 'lost projects are not');
  assert(/Calendar!A1/.test(dlg.textContent) && /Mileage!Q2/.test(dlg.textContent), 'the by-hand checklist is shown');
  dlg.querySelector('button[value=cancel]').click();
}));

test('new year: creating it copies the workbook, empties the year\'s tables, keeps the lists and sets carryover', withApp('/settings', async app => {
  await app.store.addRow('Projects', { Name: 'Old Lost Job', ProjectID: 'X-1', Client: 'Northwind Utilities', Status: 'Lost', Type: 'Fixed Fee', Rate: 200, Fee: 5000 });
  await app.store.loadTable('Projects');
  const feeder = spentOf(app, 'Feeder Study');
  const dlg = await openPreview(app);
  await createNewYear(app, dlg);
  const d = newData(app);
  assert(d, 'the new file exists');
  for (const t of ['Tasks', 'Expenses', 'Mileage', 'Invoices', 'Disbursements']) assert(d[t].length === 1 && isBlankRow(d[t][0]), `${t} is empty (one blank row)`);
  equal(d.Clients.length, app.store.data.Clients.length, 'clients kept');
  equal(d.Workers.length, app.store.data.Workers.length, 'workers kept');
  assert(!d.Projects.some(r => r[0] === 'Old Lost Job'), 'the lost project was not carried');
  const col = app.store.tables.Projects.headers.indexOf('Carryover');
  equal(d.Projects.find(r => r[0] === 'Feeder Study')[col], feeder, 'carryover = spent');
  equal(d.Projects.find(r => r[0] === 'Substation Upgrade')[col], '', 'hourly carries none');
  const text = app.$('dialog.modal').textContent;
  assert(/Calendar!A1/.test(text) && /demo:\/\/file/.test(app.$('dialog.modal input[readonly]').value), 'checklist and the link for other devices');
  closeResult(app);
}));

test('new year: the current workbook is not changed at all', withApp('/settings', async app => {
  const before = JSON.stringify(app.store.data);
  const dlg = await openPreview(app);
  await createNewYear(app, dlg);
  closeResult(app);
  await sleep(100);
  equal(JSON.stringify(app.store.data), before);
  assert(app.store.file.name.includes('2026'), 'still on 2026');
}));

test('new year: it appears in the workbook menu with its own nickname and colour, and can be switched to', withApp('/settings', async app => {
  const dlg = await openPreview(app);
  await createNewYear(app, dlg);
  closeResult(app, true);
  await app.idle(); await sleep(300);
  assert(app.store.file.name === NEW_NAME, `switched: ${app.store.file.name}`);
  equal(app.$('#filechip').textContent.trim(), '2027');
  const list = JSON.parse(localStorage.getItem('books-demo:workbooks'));
  equal(list.length, 2);
  assert(list[0].color !== list[1].color || !list[0].color, 'a different colour');
  await app.go('/time'); await app.showAll();
  assert(/No entries/.test(app.$('#listwrap').textContent), 'the empty Time table shows no entries, not a ghost row');
}));

test('new year: the first entry in the new workbook goes into its blank row, not below it', withApp('/settings', async app => {
  const dlg = await openPreview(app);
  await createNewYear(app, dlg);
  closeResult(app, true);
  await app.idle(); await sleep(300);
  await app.go('/time');
  app.set('Project', 'Substation Upgrade'); app.set('Hours', '2'); app.set('Task', 'First entry of 2027');
  await app.save();
  equal(app.store.data.Tasks.length, 1, 'one row, no blank row above it');
  equal(app.rows('Tasks').length, 1);
}));

test('new year: a file with that name already in the folder is never touched', withApp('/settings', async app => {
  files(app).set(NEW_NAME, { data: { Sentinel: [['keep me']] } });
  const dlg = await openPreview(app);
  dlg.querySelector('input[name=ok]').checked = true;
  dlg.querySelector('button[value=ok]').click(); await app.idle(); await sleep(200);
  assert(/already exists/i.test(app.toasts().join(' ')), `says so: ${app.toasts().join(' | ')}`);
  equal(JSON.stringify(newData(app)), JSON.stringify({ Sentinel: [['keep me']] }), 'left exactly as it was');
  assert(!JSON.parse(localStorage.getItem('books-demo:rollover') || 'null'), 'nothing left half-started');
}));

test('new year: an interrupted setup can be resumed on the same file, and finishes once', withApp('/settings', async app => {
  const dlg = await openPreview(app);
  app.win.__demoFault = 'clearTable';                         // the connection drops while emptying a table
  dlg.querySelector('input[name=ok]').checked = true;
  dlg.querySelector('button[value=ok]').click(); await app.idle(); await sleep(300);
  assert(app.store.constructor.files.has(NEW_NAME), 'the copy exists');
  assert(JSON.parse(localStorage.getItem('books-demo:rollover')), 'the setup is remembered');
  await app.go('/dashboard'); await app.go('/settings');
  const resume = await until(() => app.$('#resumeroll'), 4000, 'a Continue button');
  resume.click();
  await until(() => { const d = app.$('dialog.modal'); return d && /is ready/i.test(d.textContent); }, 12000, 'the finished dialog');
  const d = newData(app);
  assert(d.Tasks.length === 1 && isBlankRow(d.Tasks[0]), 'tables emptied after resuming');
  equal([...files(app).keys()].filter(k => /2027/.test(k)).length, 1, 'still one new file');
  equal(JSON.parse(localStorage.getItem('books-demo:rollover') || 'null'), null, 'nothing left pending');
  closeResult(app);
}));

test('new year: after January\'s invoice is recorded here, carryover in the new workbook can be updated, and only where it changed', withApp('/settings', async app => {
  const dlg = await openPreview(app);
  await createNewYear(app, dlg);
  closeResult(app);
  await sleep(100);
  // billed since: 100 more hours invoiced on Hospital Generator in this workbook
  await app.store.addRow('Tasks', { Date: 46300, Project: 'Hospital Generator', Hours: 100, Task: 'billed in January', Worker: 'Demo User', Billable: true, Invoiced: true, Paid: false });
  await app.store.loadAll(['Projects', 'Tasks']);
  await app.go('/dashboard'); await app.go('/settings');
  app.$('#updcarry').click();
  const pre = await until(() => app.$('dialog.modal'), 6000, 'the update preview');
  assert(/Hospital Generator/.test(pre.textContent) && !/Feeder Study/.test(pre.textContent), `only the changed project is listed: ${pre.textContent.slice(0, 200)}`);
  pre.querySelector('input[name=ok]').checked = true;
  pre.querySelector('button[value=ok]').click(); await app.idle(); await sleep(200);
  const col = app.store.tables.Projects.headers.indexOf('Carryover');
  equal(newData(app).Projects.find(r => r[0] === 'Hospital Generator')[col], spentOf(app, 'Hospital Generator'), 'updated to what is now spent');
  app.$('#updcarry').click(); await app.idle(); await sleep(200);
  assert(/up to date/i.test(app.toasts().join(' ')), 'running it again finds nothing to change');
}));
