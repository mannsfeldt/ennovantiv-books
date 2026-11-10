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

const afterAdd = fake => fake.calls.some(c => c.includes('rows/add'));

test('A-04: a failed re-read after adding a row is not reported as a failed save', withStore(async (store, fake) => {
  // the add reaches Excel, then the connection drops while re-reading the table
  fake.failNext({ when: (m, p) => m === 'GET' && p.includes("tables('Tasks')/range") && afterAdd(fake), network: true });
  const res = await store.addRow('Tasks', entry);        // must not throw: the row is in the workbook
  equal(fake.tables.Tasks.rows.length, 9, 'row count');
  equal(res.address, 'Time!A10:I10');
  assert(res.unverified, 'result says the table could not be re-read');
}));

const isAdd = (m, p) => m === 'POST' && p.includes('rows/add');
const isDelete = (m, p) => m === 'POST' && p.endsWith('/delete');

test('A-05: an add that times out after Excel did it is not sent again', withStore(async (store, fake) => {
  fake.failNext({ when: isAdd, status: 504, apply: true });
  const res = await store.addRow('Tasks', entry);
  equal(fake.tables.Tasks.rows.length, 9, 'row count (one new row, not two)');
  equal(res.address, 'Time!A10:I10');
}));

test('A-05: an add that times out without arriving reports an error and is not silently repeated', withStore(async (store, fake) => {
  fake.failNext({ when: isAdd, status: 504, apply: false });
  await rejects(() => store.addRow('Tasks', entry), /timed out/i);
  equal(fake.calls.filter(c => c.includes('rows/add')).length, 1, 'add requests sent');
  equal(fake.tables.Tasks.rows.length, 8, 'row count');
}));

test('A-05: a throttled (429) add is still retried, since nothing was done', withStore(async (store, fake) => {
  fake.failNext({ when: isAdd, status: 429, apply: false });
  await store.addRow('Tasks', entry);
  equal(fake.tables.Tasks.rows.length, 9, 'row count');
}));

test('A-05: a delete that times out after Excel did it does not delete the next row too', withStore(async (store, fake) => {
  fake.failNext({ when: isDelete, status: 504, apply: true });
  await store.deleteRow('Tasks', store.tables.Tasks.rows[2]);          // "task 3"
  equal(fake.tables.Tasks.rows.map(r => r[col('Task')]), ['task 1', 'task 2', 'task 4', 'task 5', 'task 6', 'task 7', 'task 8']);
}));

const flags = (fake, name) => fake.tables.Tasks.rows.map(r => r[col(name)]);

test('A-02: a bulk update is refused if a row was deleted in Excel meanwhile (rows shifted up)', withStore(async (store, fake) => {
  const rows = store.tables.Tasks.rows.filter(r => r.Project === 'Beta');   // tasks 2, 4, 6, 8
  fake.tables.Tasks.rows.splice(0, 1);                                      // someone deletes task 1
  await rejects(() => store.setColumn('Tasks', rows, 'Invoiced', true), /changed in Excel/);
  equal(flags(fake, 'Invoiced'), [false, false, false, false, false, false, false], 'nothing was ticked');
}));

test('A-02: a bulk update is refused if the table was sorted in Excel meanwhile', withStore(async (store, fake) => {
  const rows = store.tables.Tasks.rows.slice(0, 3);
  fake.tables.Tasks.rows.reverse();
  await rejects(() => store.setColumn('Tasks', rows, 'Paid', true), /changed in Excel/);
  assert(flags(fake, 'Paid').every(v => v === false), 'nothing was ticked');
}));

test('A-02: setColumns writes several columns on the same rows after one check', withStore(async (store, fake) => {
  const rows = store.tables.Tasks.rows.filter(r => r.Project === 'Alpha');  // tasks 1, 3, 5, 7
  await store.setColumns('Tasks', rows, { Invoice: 'AL2611', Invoiced: true });
  equal(flags(fake, 'Invoiced'), [true, false, true, false, true, false, true, false]);
  equal(flags(fake, 'Invoice'), ['AL2611', '', 'AL2611', '', 'AL2611', '', 'AL2611', '']);
  const writes = fake.calls.filter(c => c.startsWith('PATCH'));
  assert(writes[0].includes("'I"), 'the invoice number (column I) is written before the Invoiced tick, so an interrupted write never leaves "invoiced, no number"');
}));
