// Data layer for starting a new year: copying the workbook file, opening a second
// workbook, emptying tables safely, and the blank row Excel leaves in an empty table.
import { test, assert, equal, rejects } from './harness.js';
import { fakeGraph, sampleTasks, storeOn, HEADERS } from './fake-graph.js';

const realFetch = globalThis.fetch;
const withStore = (fn, schema) => async () => {
  const fake = fakeGraph(sampleTasks());
  try { await fn(await storeOn(fake, schema), fake); } finally { globalThis.fetch = realFetch; }
};
const blankRow = () => HEADERS.map(() => '');
const entry = { Date: 46400, Project: 'Alpha', Hours: 2.5, Task: 'first of the year', Worker: 'Pat', Billable: true, Invoiced: false, Paid: false };

test('copy: copies the workbook into the same folder under the new name and waits for it to appear', withStore(async (store, fake) => {
  fake.copyPolls = 2;                                        // the copy is not there for the first look-ups
  const item = await store.copyFile('Bookkeeping 2027.xlsx', { pollMs: 1 });
  equal(item.name, 'Bookkeeping 2027.xlsx');
  assert(item.id !== 'i', 'a new file, not the source');
  equal(fake.calls.filter(c => c.startsWith('POST') && c.endsWith('/copy')).length, 1, 'one copy request');
}));

test('copy: refuses when a file with that name is already in the folder, without sending a copy', withStore(async (store, fake) => {
  fake.files.x = { id: 'x', name: 'Bookkeeping 2027.xlsx', parentId: 'p' };
  const err = await rejects(() => store.copyFile('Bookkeeping 2027.xlsx', { pollMs: 1 }), /already exists/);
  assert(err.exists, 'flagged as an existing file');
  equal(fake.calls.filter(c => c.endsWith('/copy')).length, 0, 'nothing was copied over it');
}));

test('copy: gives up with a clear message if the copy never appears', withStore(async (store, fake) => {
  fake.copyPolls = 1e9;
  await rejects(() => store.copyFile('Bookkeeping 2027.xlsx', { pollMs: 1, waitMs: 30 }), /did not appear/i);
}));

test('open: a second workbook is opened with the same sign-in, and is a different store', withStore(async (store, fake) => {
  const item = await store.copyFile('Bookkeeping 2027.xlsx', { pollMs: 1 });
  const other = await store.openItem(item);
  assert(other !== store && other.file.itemId === item.id, 'its own file');
  equal(await other.token(), 'test-token', 'shares the signed-in session');
  equal(store.file.itemId, 'i', 'the first workbook is untouched');
}));

test('clear: empties the data rows of one table, keeps its header, and leaves Excel\'s one blank row', withStore(async (store, fake) => {
  const n = await store.clearTable('Tasks', { expectItemId: 'i' });
  equal(n, 8, 'rows removed');
  equal(fake.tables.Tasks.headers, HEADERS, 'header kept');
  equal(fake.tables.Tasks.rows, [blankRow()], 'one blank row remains');
}));

test('clear: refuses unless told which workbook it must be, and that is the one it is on', withStore(async (store, fake) => {
  await rejects(() => store.clearTable('Tasks'), /not the workbook/i);
  await rejects(() => store.clearTable('Tasks', { expectItemId: 'somewhere-else' }), /not the workbook/i);
  equal(fake.tables.Tasks.rows.length, 8, 'nothing was deleted');
}));

test('clear: a table that is already empty is left alone', withStore(async (store, fake) => {
  await store.clearTable('Tasks', { expectItemId: 'i' });
  fake.calls.length = 0;
  equal(await store.clearTable('Tasks', { expectItemId: 'i' }), 0);
  assert(!fake.calls.some(c => c.endsWith('/delete')), 'no delete sent');
}));

test('blank row: the first entry goes into the blank row Excel leaves in an emptied table', withStore(async (store, fake) => {
  await store.clearTable('Tasks', { expectItemId: 'i' });
  const res = await store.addRow('Tasks', entry);
  equal(fake.tables.Tasks.rows.length, 1, 'no blank row left above the entry');
  equal(fake.tables.Tasks.rows[0][3], 'first of the year');
  equal(res.address, 'Time!A2:I2');
  await store.addRow('Tasks', { ...entry, Task: 'second' });
  equal(fake.tables.Tasks.rows.map(r => r[3]), ['first of the year', 'second'], 'later entries are added below');
}));

test('blank row: formula columns that show errors in a blank row do not stop it counting as blank', withStore(async (store, fake) => {
  await store.clearTable('Tasks', { expectItemId: 'i' });
  fake.tables.Tasks.rows[0][8] = '#N/A';                      // the Invoice column pretends to be a formula
  await store.addRow('Tasks', entry);
  equal(fake.tables.Tasks.rows.length, 1);
}, { Tasks: { computed: ['Invoice'] } }));
