// Version 2 features: expense report, shared change log, project order, repeat entries, budget alerts.
import { test, assert, equal, until, sleep } from './harness.js';
import { openApp, withApp } from './app.test.js';

const localISO = (d = new Date()) => new Date(d - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10);

// ---- expense report: no "next invoice" date ----
test('report: the expense report page does not mention the next invoice', withApp('/report', async app => {
  await until(() => app.$('#rprev .invoice'), 5000, 'the report');
  assert(!app.$('.nextbill'), 'no "next invoice" banner');
  assert(!/next invoice/i.test(app.$('#view').textContent), 'no mention anywhere on the page');
}));

test('report: it runs to today by default, not to a billing day', withApp('/report', async app => {
  await until(() => app.$('#r-t'), 5000, 'the form');
  equal(app.$('#r-t').value, localISO(), 'To = today');
  equal(app.$('#r-f').value, localISO().slice(0, 4) + '-01-01', 'From = start of the year');
}));

test('report: the printed report has no "next invoice" line', withApp('/report', async app => {
  await until(() => app.$('#r-print'), 5000, 'the print button');
  app.$('#r-print').click(); await sleep(100);
  const printed = app.$('#print-root').textContent;
  assert(/expense report/i.test(printed), 'printed something');
  assert(!/next invoice/i.test(printed), 'no next-invoice line');
}));

// ---- shared change log: data layer ----
import { fakeGraph, sampleTasks, storeOn } from './fake-graph.js';
const realFetch = globalThis.fetch;
const LOG_HEADERS = ['When', 'Who', 'Action', 'Where', 'Detail'];
const withLogStore = fn => async () => {
  const fake = fakeGraph(sampleTasks());
  try { await fn(await storeOn(fake), fake); } finally { globalThis.fetch = realFetch; }
};

test('log: addRows writes several rows in one request', withLogStore(async (store, fake) => {
  fake.tables.ChangeLog = { sheet: 'ChangeLog', headers: LOG_HEADERS, rows: [] };
  await store.addRows('ChangeLog', [['2026-11-23 10:00', 'David', 'Added time entry', 'Time!A9:I9', ''], ['2026-11-23 10:01', 'David', 'Invoice paid', 'Invoices!A4:J4', 'NW2611']]);
  equal(fake.tables.ChangeLog.rows.length, 2, 'rows written');
  equal(fake.calls.filter(c => c.includes('rows/add')).length, 1, 'requests sent');
}));

test('log: createTable makes the sheet, the header row and a named table', withLogStore(async (store, fake) => {
  const t = await store.createTable('ChangeLog', 'ChangeLog', LOG_HEADERS);
  equal(t.headers, LOG_HEADERS);
  equal(fake.tables.ChangeLog.sheet, 'ChangeLog', 'table named ChangeLog on its own sheet');
  assert(fake.calls.some(c => c.endsWith('/worksheets/add')), 'a sheet was added');
}));

// ---- shared change log: in the app ----
const logRows = app => (app.store.data.ChangeLog || []).map(r => ({ when: r[0], who: r[1], action: r[2], where: r[3], detail: r[4] }));
async function enableLog(app) {
  await app.go('/settings');
  app.$('#addlog').click(); await app.idle(); await sleep(100);
}
const logTime = app => { app.set('Project', 'Substation Upgrade'); app.set('Hours', '1.5'); app.set('Task', 'Site call'); };

test('log: Settings adds the ChangeLog table to the workbook, once', withApp('/settings', async app => {
  assert(!app.store.data.ChangeLog, 'not there at first');
  await enableLog(app);
  assert(app.store.data.ChangeLog, 'table created');
  assert(!app.$('#addlog'), 'button gone once it is on');
  assert(/On/.test(app.$('#logcard').textContent), 'says it is on');
}));

test('log: once on, every save is recorded with who, when, what and where', withApp('/settings', async app => {
  await enableLog(app);
  await app.go('/time'); logTime(app); await app.save();
  const row = await until(() => logRows(app).find(r => r.action === 'Added time entry'), 4000, 'the log row');
  equal(logRows(app)[0].action, 'Added shared change log', 'turning the log on is itself recorded');
  equal([row.who, row.action], ['Demo User', 'Added time entry']);
  assert(/^Time!A\d+:I\d+/.test(row.where), `where is the sheet range: ${row.where}`);
  assert(/^\d{4}-\d\d-\d\d \d\d:\d\d$/.test(row.when), `when is a readable timestamp: ${row.when}`);
}));

test('log: a log write that fails never breaks the save, and is sent with the next one', withApp('/settings', async app => {
  await enableLog(app);
  const real = app.store.addRows.bind(app.store); let failed = false;
  app.store.addRows = async (...a) => { if (!failed) { failed = true; throw new Error('Failed to fetch'); } return real(...a); };
  const before = app.rows('Tasks').length;
  await app.go('/time'); logTime(app); await app.save();
  equal(app.store.data.Tasks.length, before + 1, 'the entry was saved');
  assert(!/⚠/.test(app.toasts().join(' ')), 'no warning shown for a log problem');
  await sleep(300);
  equal(logRows(app).filter(r => r.action === 'Added time entry').length, 0, 'the entry is not logged yet');
  logTime(app); await app.save();                           // the next save carries the missed row too
  await until(() => logRows(app).filter(r => r.action === 'Added time entry').length >= 2, 4000, 'both log rows');
  equal(logRows(app).filter(r => r.action === 'Added time entry').length, 2);
}));

test('log: entries waiting for a different workbook are not written into this one', withApp('/settings', async app => {
  await enableLog(app);
  localStorage.setItem('books-demo:logQueue', JSON.stringify([{ wb: 'demo://elsewhere', row: ['2026-01-01 00:00', 'X', 'ELSEWHERE', '', ''] }]));
  await app.go('/time'); logTime(app); await app.save();
  await until(() => logRows(app).find(r => r.action === 'Added time entry'), 4000, 'the log row');
  assert(!logRows(app).some(r => r.action === 'ELSEWHERE'), 'the other workbook\'s entry stayed out');
  assert(JSON.parse(localStorage.getItem('books-demo:logQueue')).some(q => q.wb === 'demo://elsewhere'), 'and is still waiting for its own workbook');
}));

test('log: the dashboard shows what other people changed', withApp('/settings', async app => {
  await enableLog(app);
  app.store.data.ChangeLog.push(['2026-11-20 09:15', 'Shari Mannsfeldt', 'Added expense', 'Expenses!A40:N40', '$12.00 · Staples']);
  await app.store.loadTable('ChangeLog');
  await app.go('/dashboard');
  const text = app.$('#view').textContent;
  assert(text.includes('Shari Mannsfeldt') && text.includes('Added expense'), 'her change is listed with her name');
}));

test('log: without the table the dashboard says how to turn the shared log on', withApp('/dashboard', async app => {
  assert(/shared change log/i.test(app.$('#view').textContent), 'hint shown');
}));

// ---- dashboard: days since the last expense report ----
const daysAgoStamp = n => { const d = new Date(); d.setDate(d.getDate() - n); return `${localISO(d)} 09:00`; };
const reportTile = app => [...app.$$('.stat')].find(t => /expense report/i.test(t.querySelector('label')?.textContent || ''));
async function withLog(app, rowsToAdd) {
  await enableLog(app);
  for (const r of rowsToAdd) app.store.data.ChangeLog.push(r);
  await app.store.loadTable('ChangeLog');
  await app.go('/dashboard');
}

test('report days: the dashboard counts the days since the last expense report', withApp('/settings', async app => {
  await withLog(app, [[daysAgoStamp(10), 'Shari Mannsfeldt', 'Printed expense report', '', 'Shari · 2026-01-01 to 2026-11-13 · $120.00']]);
  const t = reportTile(app);
  assert(t, 'a tile for it');
  assert(/10 days/.test(t.textContent), `says 10 days: ${t.textContent}`);
  assert(/Shari/.test(t.textContent), 'and who filed it');
}));

test('report days: the most recent of printing and marking reimbursed counts', withApp('/settings', async app => {
  await withLog(app, [
    [daysAgoStamp(10), 'David', 'Printed expense report', '', ''],
    [daysAgoStamp(3), 'David', 'Expense report reimbursed', 'Expenses ×4, Mileage ×2', 'David · $400.00'],
  ]);
  assert(/3 days/.test(reportTile(app).textContent), reportTile(app).textContent);
}));

test('report days: says so when no report has been filed yet', withApp('/dashboard', async app => {
  assert(/none on record/i.test(reportTile(app).textContent), reportTile(app).textContent);
}));

test('report days: printing a report starts the count again from today (even without the shared log)', withApp('/report', async app => {
  await until(() => app.$('#r-print'), 5000, 'the print button');
  app.$('#r-print').click(); await sleep(100);
  await app.go('/dashboard');
  assert(/today/i.test(reportTile(app).textContent), reportTile(app).textContent);
}));

// ---- project order in the Time form ----
const projectOrder = app => [...app.form().elements.Project.querySelectorAll('optgroup')].map(g => ({ group: g.label, items: [...g.children].map(o => o.value) }));
const daysAgoSerial = n => { const d = new Date(); d.setDate(d.getDate() - n); return Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 864e5); };

test('projects: the list runs from most recently used to least, with new projects first, not grouped by client', withApp('/dashboard', async app => {
  app.store.data.Tasks = [];                                // start from a known history
  const log = (project, daysAgo, worker = 'Demo User') => app.store.addRow('Tasks', { Date: daysAgoSerial(daysAgo), Project: project, Hours: 1, Task: 't', Worker: worker, Billable: true, Invoiced: false, Paid: false });
  await log('Feeder Study', 20); await log('Internal', 2); await log('Substation Upgrade', 0);
  await log('Feeder Study', 0, 'Alex Helper');              // someone else's work must not count as mine
  await app.store.addRow('Projects', { Name: 'Brand New', ProjectID: 'NW-009', Client: 'Northwind Utilities', Status: 'Active', Type: 'Hourly', Rate: 195 });
  await app.store.loadAll(['Projects', 'Tasks']);
  await app.go('/time');
  const groups = projectOrder(app);
  equal(groups[0].items, ['Brand New', 'Hospital Generator', 'Substation Upgrade', 'Internal', 'Feeder Study', 'Training'], 'one list, newest and most recently used first');
  assert(!groups.some(g => /Northwind|Contoso/.test(g.group)), 'no client groups');
  equal(groups[1].items, ['Arc Flash Review'], 'opportunities stay in their own group at the end');
}));

// ---- repeat and favourites (time) ----
async function seedHistory(app) {
  app.store.data.Tasks = [];
  const log = (project, task, hours, daysAgo, worker = 'Demo User') => app.store.addRow('Tasks', { Date: daysAgoSerial(daysAgo), Project: project, Hours: hours, Task: task, Worker: worker, Billable: true, Invoiced: false, Paid: false });
  await log('Feeder Study', 'Report writing', 3, 9); await log('Feeder Study', 'Report writing', 2, 4); await log('Feeder Study', 'report writing', 2.5, 1);
  await log('Substation Upgrade', 'SLD revisions', 1.5, 2);
  await log('Hospital Generator', 'Site visit', 6, 0, 'Alex Helper');           // someone else's favourite is not mine
  await app.store.loadAll(['Projects', 'Tasks']);
  await app.go('/dashboard'); await app.go('/time');
}
const chips = app => app.$$('#quick .chip');

test('repeat: quick-add chips offer my most used project and task pairs, most used first', withApp('/dashboard', async app => {
  await seedHistory(app);
  const labels = chips(app).map(c => c.textContent.replace(/\s+/g, ' ').trim());
  equal(labels.length, 2, 'two pairs, not the other person\'s');
  assert(/Feeder Study/.test(labels[0]) && /Report writing/i.test(labels[0]), `most used first: ${labels[0]}`);
  assert(/Substation Upgrade/.test(labels[1]), labels[1]);
}));

test('repeat: tapping a chip fills the form from the latest time that pair was logged', withApp('/dashboard', async app => {
  await seedHistory(app);
  chips(app)[0].click(); await sleep(60);
  const f = app.form().elements;
  equal([f.Project.value, f.Task.value, f.Hours.value], ['Feeder Study', 'report writing', '2.5'], 'project, the latest wording of the task, and the latest hours');
  assert(/remaining/.test(app.$('#budget').textContent), 'the budget panel updates too');
  await app.save();
  const row = app.rows('Tasks').at(-1);
  equal([row.Project, row.Task, row.Hours, row.Billable], ['Feeder Study', 'report writing', 2.5, true], 'one more tap saves it');
}));

test('repeat: no chips while editing an entry', withApp('/dashboard', async app => {
  await seedHistory(app);
  await app.showAll();
  await app.edit('Tasks', app.rows('Tasks')[0]);
  equal(chips(app).length, 0);
}));

test('repeat: a project that is no longer open is not offered', withApp('/dashboard', async app => {
  await seedHistory(app);
  const i = app.rows('Projects').findIndex(p => p.Name === 'Feeder Study');
  await app.poke('Projects', i, 'Status', 'Lost');
  await app.go('/dashboard'); await app.go('/time');
  equal(chips(app).map(c => c.textContent).filter(t => /Feeder/.test(t)).length, 0);
}));

test('repeat: the Repeat button on a list row copies it into the form, dated today, without opening it for edit', withApp('/dashboard', async app => {
  await seedHistory(app);
  await app.showAll();
  const row = app.rows('Tasks').find(r => r.Project === 'Substation Upgrade');
  app.$(`[data-repeat="${row._row}"]`).click(); await sleep(60);
  const f = app.form().elements;
  equal([f.Project.value, f.Task.value, f.Hours.value, f.Date.value], ['Substation Upgrade', 'SLD revisions', '1.5', localISO()]);
  assert(/Log time/.test(app.$('#entry h2').textContent), 'still the new-entry form, not an edit');
}));

// ---- fixed-fee budget alerts ----
const budgetCard = app => app.$('#budgets');
async function setHours(app, project, hours, { invoiced = false } = {}) {
  await app.store.addRow('Tasks', { Date: daysAgoSerial(5), Project: project, Hours: hours, Task: 'work', Worker: 'Demo User', Billable: true, Invoiced: invoiced, Paid: false });
  await app.store.loadAll(['Projects', 'Tasks']);
  await app.go('/time'); await app.go('/dashboard');
}

test('budget: the dashboard lists fixed-fee projects that are past 80% of their fee, most used first', withApp('/dashboard', async app => {
  app.store.data.Tasks = [];
  await setHours(app, 'Feeder Study', 51);                  // 60 h budget: 85%
  const card = budgetCard(app);
  assert(card, 'a budget card');
  assert(/Feeder Study/.test(card.textContent) && /85%/.test(card.textContent), card.textContent);
  assert(!/Hospital Generator/.test(card.textContent), 'projects under 80% are not listed');
}));

test('budget: hours already billed beyond the fee count as over budget', withApp('/dashboard', async app => {
  app.store.data.Tasks = [];
  await setHours(app, 'Hospital Generator', 100, { invoiced: true });   // $21,000 billed on an $18,000 fee
  const row = [...budgetCard(app).querySelectorAll('[data-project]')].find(r => r.dataset.project === 'Hospital Generator');
  assert(row, 'listed');
  assert(/117%/.test(row.textContent) && /over/i.test(row.textContent), row.textContent);
  await app.go('/time');
  app.set('Project', 'Hospital Generator');
  assert(/over budget/.test(app.$('#budget').textContent), `the Time form agrees: ${app.$('#budget').textContent}`);
}));

test('budget: over-budget projects come first, then approaching ones, then fully used ones', withApp('/dashboard', async app => {
  app.store.data.Tasks = [];
  await setHours(app, 'Hospital Generator', 100, { invoiced: true });   // over: 117%
  await setHours(app, 'Feeder Study', 51);                              // approaching: 85%
  const names = [...budgetCard(app).querySelectorAll('[data-project]')].map(r => r.dataset.project);
  equal(names, ['Hospital Generator', 'Feeder Study']);
}));

test('budget: says so when every fixed-fee project is comfortably within its fee', withApp('/dashboard', async app => {
  app.store.data.Tasks = [];
  await setHours(app, 'Feeder Study', 5);
  assert(/within/i.test(budgetCard(app).textContent) && !/Feeder Study/.test(budgetCard(app).textContent), budgetCard(app).textContent);
}));

test('budget: opportunities and on-hold projects are not alerted', withApp('/dashboard', async app => {
  app.store.data.Tasks = [];
  await setHours(app, 'Arc Flash Review', 44);              // 44 of 45 h, but still an opportunity
  assert(!/Arc Flash Review/.test(budgetCard(app).textContent));
}));

test('budget: tapping a project goes to the Time form with it selected', withApp('/dashboard', async app => {
  app.store.data.Tasks = [];
  await setHours(app, 'Feeder Study', 51);
  app.$('#budgets [data-project="Feeder Study"]').click(); await sleep(150);
  equal(app.win.location.hash, '#/time');
  equal(app.form().elements.Project.value, 'Feeder Study');
}));
