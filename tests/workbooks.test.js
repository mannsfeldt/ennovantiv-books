// Several workbooks (2026 / 2027): the switcher in the header and the list in Settings.
import { test, assert, equal, until, sleep } from './harness.js';
import { openApp, withApp } from './app.test.js';

const openMenu = async app => { app.$('#filechip').click(); await sleep(40); return app.$('.wbmenu'); };
const menuItems = app => app.$$('.wbmenu [data-wb]').map(b => b.textContent.replace(/\s+/g, ' ').trim());
async function addWorkbook(app, url, nickname = '') {
  await app.go('/settings');
  app.$('#s-url').value = url; app.$('#s-nick').value = nickname;
  app.$('#wbadd').requestSubmit(); await app.idle(); await sleep(150);
}

test('wb: the header names the open workbook and the menu lists the saved ones', withApp('/dashboard', async app => {
  assert(app.$('#filechip').textContent.includes('2026'), 'chip shows the workbook (its year, until it has a nickname)');
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

// ---- a colour for each workbook ----
const accent = app => app.win.getComputedStyle(app.doc.documentElement).getPropertyValue('--blue').trim().toLowerCase();
const hue = hex => {
  const [r, g, b] = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16) / 255), mx = Math.max(r, g, b), mn = Math.min(r, g, b), d = mx - mn;
  if (!d) return { h: 0, s: 0 };
  const h = mx === r ? ((g - b) / d) % 6 : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return { h: (h * 60 + 360) % 360, s: d / (1 - Math.abs(mx + mn - 1) || 1) };
};

test('colour: each workbook gets its own accent colour, and the app changes colour when you switch', withApp('/dashboard', async app => {
  const first = accent(app);
  assert(/^#[0-9a-f]{6}$/.test(first), `an accent is set: ${first}`);
  await addWorkbook(app, 'demo://2027', '2027');
  const second = accent(app);
  assert(second !== first, `the new workbook has a different colour (${first} vs ${second})`);
  await openMenu(app);
  app.$$('.wbmenu [data-wb]').find(b => !b.textContent.includes('2027')).click();
  await app.idle(); await sleep(200);
  equal(accent(app), first, 'back to the first workbook\'s colour');
}));

test('colour: pick one next to the nickname in Settings; it applies at once and is remembered', withApp('/settings', async app => {
  const radio = app.$('.wb-list input[type=radio][value=magenta]');
  assert(radio, 'a magenta swatch beside the current workbook');
  radio.click(); await sleep(80);
  equal(accent(app), '#c2569b');
  equal(JSON.parse(localStorage.getItem('books-demo:workbooks'))[0].color, 'magenta', 'saved on this device');
}));

test('colour: the palette has no red or green (the brand reserves them for electrical safety)', withApp('/settings', async app => {
  const colours = app.$$('.wb-list tr:first-child .sw i').map(i => app.win.getComputedStyle(i).backgroundColor);
  assert(colours.length >= 5, `swatches shown: ${colours.length}`);
  for (const c of colours) {
    const [r, g, b] = c.match(/\d+/g).map(Number), { h, s } = hue('#' + [r, g, b].map(v => v.toString(16).padStart(2, '0')).join(''));
    assert(s < .2 || !(h < 20 || h > 340 || (h > 80 && h < 170)), `${c} (hue ${Math.round(h)}) is red or green`);
  }
}));

// ---- the year boundary: December's work is invoiced in January, in the same workbook as the time ----
const bar = app => app.$('.infobar')?.textContent || '';

test('year: the January invoice for December\'s work is recorded in the workbook that holds the time', withApp('/dashboard', async app => {
  await app.store.addColumn('Tasks', 'Invoice'); await app.store.addColumn('Disbursements', 'Invoice');
  await app.go('/invoices');
  const s = app.$('#ic'); s.value = 'Northwind Utilities'; s.dispatchEvent(new app.win.Event('change'));
  app.$('#idate').value = '2027-01-04'; app.$('#idate').dispatchEvent(new app.win.Event('change')); await sleep(40);
  assert(/NW2701/.test(app.$('#invprev').textContent), 'numbered for January 2027');
  app.$('#rec').click();
  const dlg = await until(() => app.$('dialog.modal'), 6000, 'the checklist');
  dlg.querySelectorAll('input[type=checkbox][required]').forEach(c => { c.checked = true; });
  dlg.querySelector('button[value=ok]').click(); await app.idle();
  const inv = app.rows('Invoices').find(r => r.Number === 'NW2701');
  assert(inv, 'recorded in this (2026) workbook');
  equal(new Date(Date.UTC(1899, 11, 30) + inv.Date * 864e5).toISOString().slice(0, 10), '2027-01-04', 'dated in January 2027');
  assert(app.rows('Tasks').some(r => r.Invoice === 'NW2701' && r.Invoiced === true), 'the 2026 time entries are ticked and numbered');
}));

test('year: last year\'s workbook says so, and offers to switch to this year\'s saved workbook', withApp('/dashboard', async app => {
  const thisYear = String(new Date().getFullYear());
  await addWorkbook(app, `demo://${thisYear - 1}`, 'last year');          // now in last year's workbook
  await app.go('/dashboard');
  assert(new RegExp(`${thisYear - 1} workbook`).test(app.$('.warnbox')?.textContent || ''), 'explains this is last year\'s workbook');
  const sw = app.$('.warnbox [data-switch]');
  assert(sw, 'a switch button for this year\'s workbook');
  sw.click(); await app.idle(); await sleep(200);
  assert(app.store.file.name.includes(thisYear), 'switched to this year\'s workbook');
}));

test('year: unpaid invoices in another workbook are shown, with a way to open it', withApp('/dashboard', async app => {
  const unpaid = app.rows('Invoices').filter(i => !i.Paid);
  assert(unpaid.length > 0, 'sample data has an unpaid invoice');
  await addWorkbook(app, 'demo://2027', '2027');
  await app.go('/dashboard');
  assert(/unpaid in 2026/i.test(bar(app)) && /1 invoice/.test(bar(app)), `the dashboard says what is still unpaid in 2026: ${bar(app)}`);
  await openMenu(app);
  assert(menuItems(app).some(t => /1 unpaid/.test(t)), `the menu shows it too: ${menuItems(app)}`);
  app.$('.infobar [data-switch]').click(); await app.idle(); await sleep(200);
  assert(app.store.file.name.includes('2026'), 'opened the 2026 workbook');
}));
