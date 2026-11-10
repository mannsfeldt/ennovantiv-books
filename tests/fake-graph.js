// A small stand-in for the Microsoft Graph Excel API: tables held as arrays,
// plus switches to inject the failures the real service produces (time-outs
// that still did the work, dropped connections, rows moved by someone else).
import { GraphStore } from '../js/store-graph.js';

const colIndex = l => [...l].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const colName = n => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };

// tables: { Tasks: { sheet: 'Time', headers: [...], rows: [[...], ...] } }
export function fakeGraph(tables) {
  const fake = { tables, calls: [], faults: [] };
  const bySheet = sheet => Object.values(tables).find(t => t.sheet === sheet);
  const parse = a => { const m = a.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/); return { c1: colIndex(m[1]) - 1, r1: +m[2], c2: colIndex(m[3] || m[1]) - 1, r2: +(m[4] || m[2]) }; };

  function handle(method, path, body) {
    let m;
    if (path === '/createSession') return { id: 'session-1' };
    if ((m = path.match(/^\/tables\('(.+?)'\)\/range/))) {
      const t = tables[m[1]];
      return { address: `${t.sheet}!A1:${colName(t.headers.length)}${t.rows.length + 1}`, values: [[...t.headers], ...t.rows.map(r => [...r])] };
    }
    if ((m = path.match(/^\/tables\('(.+?)'\)\/rows\/add/))) {
      const t = tables[m[1]];
      t.rows.push(body.values[0].map(v => v ?? ''));
      return { index: t.rows.length - 1 };
    }
    if ((m = path.match(/^\/worksheets\('(.+?)'\)\/range\(address='(.+?)'\)(\/delete)?/))) {
      const t = bySheet(m[1]), a = parse(m[2]);
      if (m[3]) { t.rows.splice(a.r1 - 2, a.r2 - a.r1 + 1); return null; }
      if (method === 'PATCH') {
        (body.values || body.formulas || body.formulasR1C1).forEach((row, i) => row.forEach((v, k) => { if (v !== null) t.rows[a.r1 - 2 + i][a.c1 + k] = v; }));
        return {};
      }
      const out = [];
      for (let r = a.r1; r <= a.r2; r++) out.push((t.rows[r - 2] || t.headers.map(() => '')).slice(a.c1, a.c2 + 1));
      return { values: out, formulasR1C1: out };
    }
    throw new Error(`fake Graph: unhandled ${method} ${path}`);
  }

  // fault: { when(method, path), status?, network?, apply?, times? }
  //   apply: true  = the server did the work, then failed to say so
  fake.failNext = fault => fake.faults.push({ times: 1, ...fault });
  fake.fetch = async (url, init = {}) => {
    const method = init.method || 'GET';
    const path = decodeURIComponent(String(url).split('/workbook')[1] || '');
    const body = init.body ? JSON.parse(init.body) : undefined;
    fake.calls.push(`${method} ${path}`);
    const fault = fake.faults.find(f => f.times > 0 && f.when(method, path));
    if (fault) {
      fault.times--;
      if (fault.apply) handle(method, path, body);
      if (fault.network) throw new TypeError('Failed to fetch');
      return new Response(JSON.stringify({ error: { message: 'Gateway Timeout' } }), { status: fault.status, headers: { 'Retry-After': '0' } });
    }
    const res = handle(method, path, body);
    return res === null ? new Response(null, { status: 204 }) : new Response(JSON.stringify(res), { status: 200 });
  };
  return fake;
}

export const HEADERS = ['Date', 'Project', 'Hours', 'Task', 'Worker', 'Billable', 'Invoiced', 'Paid', 'Invoice'];
export const sampleTasks = () => ({
  Tasks: {
    sheet: 'Time', headers: [...HEADERS],
    rows: [1, 2, 3, 4, 5, 6, 7, 8].map(i => [46300 + i, i % 2 ? 'Alpha' : 'Beta', i, `task ${i}`, 'Pat', true, false, false, '']),
  },
});

// A GraphStore wired to the fake: no sign-in, no network.
export async function storeOn(fake, schema = { Tasks: { computed: [] } }) {
  const store = new GraphStore({}, schema);
  store.token = async () => 'test-token';
  store.file = { driveId: 'd', itemId: 'i', parentId: 'p' };
  store.session = 'session-1';
  globalThis.fetch = fake.fetch;
  await store.loadTable('Tasks');
  return store;
}
