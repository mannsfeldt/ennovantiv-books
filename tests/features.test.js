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
