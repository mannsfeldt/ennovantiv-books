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
