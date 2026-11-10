// Unit tests for the Excel data layer (js/store-graph.js) against a fake Graph.
import { test, assert, equal, rejects } from './harness.js';
import { fakeGraph, sampleTasks, storeOn, HEADERS } from './fake-graph.js';

const realFetch = globalThis.fetch;
const col = name => HEADERS.indexOf(name);
const entry = { Date: 46400, Project: 'Alpha', Hours: 2.5, Task: 'new entry', Worker: 'Pat', Billable: true, Invoiced: false, Paid: false };
// run a test against a fresh fake workbook, restoring the real fetch afterwards
const withStore = fn => async () => {
  const fake = fakeGraph(sampleTasks());
  try { await fn(await storeOn(fake), fake); } finally { globalThis.fetch = realFetch; }
};

test('store: addRow appends one row and reports its address', withStore(async (store, fake) => {
  const res = await store.addRow('Tasks', entry);
  equal(fake.tables.Tasks.rows.length, 9, 'row count');
  equal(res.address, 'Time!A10:I10');
  equal(fake.tables.Tasks.rows[8][col('Task')], 'new entry');
}));

test('store: setColumn writes the value on exactly the rows asked for', withStore(async (store, fake) => {
  const rows = store.tables.Tasks.rows.filter(r => r.Project === 'Alpha');   // rows 2, 4, 6, 8
  await store.setColumn('Tasks', rows, 'Invoiced', true);
  equal(fake.tables.Tasks.rows.map(r => r[col('Invoiced')]), [true, false, true, false, true, false, true, false]);
}));

test('store: updateRow refuses when the row changed in Excel', withStore(async (store, fake) => {
  const row = store.tables.Tasks.rows[2];
  fake.tables.Tasks.rows[2][col('Hours')] = 99;       // edited in Excel after we loaded
  await rejects(() => store.updateRow('Tasks', row, { Task: 'x' }), /changed in Excel/);
  equal(fake.tables.Tasks.rows[2][col('Task')], 'task 3', 'row left untouched');
}));
