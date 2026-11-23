// Several workbooks (2026 / 2027): the switcher in the header and the list in Settings.
import { test, assert, equal, sleep } from './harness.js';
import { openApp, withApp } from './app.test.js';

const openMenu = async app => { app.$('#filechip').click(); await sleep(40); return app.$('.wbmenu'); };
const menuItems = app => app.$$('.wbmenu [data-wb]').map(b => b.textContent.replace(/\s+/g, ' ').trim());
async function addWorkbook(app, url, nickname = '') {
  await app.go('/settings');
  app.$('#s-url').value = url; app.$('#s-nick').value = nickname;
  app.$('#wbadd').requestSubmit(); await app.idle(); await sleep(150);
}

test('wb: the header names the open workbook and the menu lists the saved ones', withApp('/dashboard', async app => {
  assert(app.$('#filechip').textContent.includes('DEMO'), 'chip shows the workbook');
  const menu = await openMenu(app);
  assert(menu, 'menu opens');
  equal(menuItems(app).length, 1, 'one saved workbook');
}));

test('wb: adding a second workbook switches to it, and the menu switches back', withApp('/dashboard', async app => {
  const tasks2026 = app.rows('Tasks').length;
  assert(tasks2026 > 0, 'sample data');
  await addWorkbook(app, 'demo://2027', '2027');
  equal(app.$('#filechip').textContent.trim(), '2027', 'chip shows the nickname');
  assert(app.store.file.name.includes('2027'), 'the 2027 workbook is open');
  equal(app.rows('Tasks').length, 0, 'its data, not 2026\'s');
  await openMenu(app);
  equal(menuItems(app).length, 2, 'both listed');
  app.$$('.wbmenu [data-wb]').find(b => !b.textContent.includes('2027')).click();
  await app.idle(); await sleep(200);
  assert(!app.store.file.name.includes('2027'), 'back on the first workbook');
  equal(app.rows('Tasks').length, tasks2026, '2026 data is back');
}));

test('wb: a nickname can be changed in Settings and shows in the header', withApp('/settings', async app => {
  const nick = app.$('input.nick'); nick.value = 'Current year'; nick.dispatchEvent(new app.win.Event('change', { bubbles: true })); await sleep(60);
  equal(app.$('#filechip').textContent.trim(), 'Current year');
  equal(JSON.parse(localStorage.getItem('books-demo:workbooks'))[0].nickname, 'Current year', 'saved on this device');
}));

test('wb: a workbook that will not open leaves you on the one you were using', withApp('/dashboard', async app => {
  const name = app.store.file.name;
  await addWorkbook(app, 'demo://fail', 'broken');
  equal(app.store.file.name, name, 'still on the original workbook');
  assert(/couldn.t open/i.test(app.toasts().join(' ')), 'says so');
  equal(JSON.parse(localStorage.getItem('books-demo:workbooks')).length, 1, 'the broken link was not saved');
}));

test('wb: an interrupted invoice belongs to its workbook and cannot be resumed from another', withApp('/dashboard', async app => {
  localStorage.setItem('books-demo:pendingInvoice', JSON.stringify({ number: 'NW2611', client: 'Northwind Utilities', date: '2026-11-02', total: 100, mark: true, when: new Date().toISOString(), added: true, invoicesBefore: 0, wb: 'demo://2026', tasks: { need: {}, already: {} }, disb: { need: {}, already: {} } }));
  await addWorkbook(app, 'demo://2027', '2027');
  await app.go('/invoices');
  assert(!app.$('#resume'), 'no Resume in the other workbook');
  const card = app.$('.card.pending')?.textContent || '';
  assert(/NW2611/.test(card) && /switch/i.test(card), 'explains which workbook it belongs to');
}));

test('wb: trying the demo never touches the saved workbook links of the live app', async () => {
  localStorage.setItem('books:settings', JSON.stringify({ fileShareUrl: 'https://live.example/book.xlsx' }));
  localStorage.setItem('books:workbooks', JSON.stringify([{ id: 'live', url: 'https://live.example/book.xlsx', nickname: 'Live' }]));
  const app = await openApp('/dashboard');
  let settings, list;
  try { await addWorkbook(app, 'demo://2027', '2027'); }
  finally {
    app.close();
    settings = localStorage.getItem('books:settings'); list = localStorage.getItem('books:workbooks');
    localStorage.setItem('books:settings', ''); localStorage.setItem('books:workbooks', '');   // tidy up the stand-ins
  }
  equal(JSON.parse(settings).fileShareUrl, 'https://live.example/book.xlsx', 'live link untouched');
  equal(JSON.parse(list).length, 1, 'live list untouched');
});
