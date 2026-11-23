// A small stand-in for the Microsoft Graph Excel API: tables held as arrays,
// plus switches to inject the failures the real service produces (time-outs
// that still did the work, dropped connections, rows moved by someone else),
// and a tiny drive so file copies behave like the real, asynchronous ones.
import { GraphStore } from '../js/store-graph.js';

const colIndex = l => [...l].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const colName = n => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
const GRAPH = 'https://graph.microsoft.com/v1.0';

// tables: { Tasks: { sheet: 'Time', headers: [...], rows: [[...], ...] } }
export function fakeGraph(tables) {
  const fake = { tables, calls: [], faults: [], sheets: {} };
  const bySheet = sheet => Object.values(tables).find(t => t.sheet === sheet);
  const parse = a => { const m = a.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/); return { c1: colIndex(m[1]) - 1, r1: +m[2], c2: colIndex(m[3] || m[1]) - 1, r2: +(m[4] || m[2]) }; };

  // ---- the workbook ----
  function handle(method, path, body) {
    let m;
    if (path === '/createSession') return { id: 'session-1' };
    // creating a sheet, writing its header cells, turning them into a named table
    if (path === '/worksheets/add') { fake.sheets[body.name] = {}; return { name: body.name }; }
    if (method === 'POST' && path === '/tables/add') {
      const [sheet, addr] = body.address.split('!'), a = parse(addr), cells = fake.sheets[sheet] || {};
      const headers = (cells[addr] || [[]])[0].slice(0, a.c2 - a.c1 + 1);
      const id = 'tbl' + (Object.keys(tables).length + 1);
      tables[id] = { sheet, headers, rows: [] };
      return { id, name: id };
    }
    if (method === 'PATCH' && (m = path.match(/^\/tables\('(.+?)'\)$/))) { tables[body.name] = tables[m[1]]; delete tables[m[1]]; return {}; }
    if (method === 'PATCH' && (m = path.match(/^\/worksheets\('(.+?)'\)\/range\(address='(.+?)'\)/)) && fake.sheets[m[1]] && !bySheet(m[1])) {
      fake.sheets[m[1]][m[2]] = body.values; return {};
    }
    if ((m = path.match(/^\/tables\('(.+?)'\)\/range/))) {
      const t = tables[m[1]];
      return { address: `${t.sheet}!A1:${colName(t.headers.length)}${t.rows.length + 1}`, values: [[...t.headers], ...t.rows.map(r => [...r])] };
    }
    if ((m = path.match(/^\/tables\('(.+?)'\)\/rows\/add/))) {
      const t = tables[m[1]];
      const first = t.rows.length;
      body.values.forEach(v => t.rows.push(v.map(c => c ?? '')));
      return { index: first };
    }
    if ((m = path.match(/^\/worksheets\('(.+?)'\)\/range\(address='(.+?)'\)(\/delete)?/))) {
      const t = bySheet(m[1]), a = parse(m[2]);
      if (m[3]) {
        t.rows.splice(a.r1 - 2, a.r2 - a.r1 + 1);
        if (!t.rows.length) t.rows.push(t.headers.map(() => ''));   // like Excel: a table always keeps one (blank) row
        return null;
      }
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

  // ---- the drive: copying a file is asynchronous (202 first, the file later) ----
  fake.files = { i: { id: 'i', name: 'Bookkeeping 2026.xlsx', parentId: 'p' } };
  fake.copyPolls = 0;            // how many look-ups fail before a copy shows up
  fake.pendingCopies = [];
  const item = f => ({ id: f.id, name: f.name, webUrl: `https://example.sharepoint.com/Finance/${f.name}`, parentReference: { driveId: 'd', id: f.parentId } });
  const reply = (obj, status = 200) => new Response(JSON.stringify(obj), { status });
  const drivePath = url => decodeURIComponent(String(url).replace(GRAPH, '').split('?')[0]);
  function driveRoute(method, path, body) {
    let m;
    if (method === 'POST' && (m = path.match(/^\/drives\/d\/items\/([^/]+)\/copy$/))) {
      if (Object.values(fake.files).some(f => f.name === body.name && f.parentId === body.parentReference.id)) return reply({ error: { message: 'nameAlreadyExists' } }, 409);
      fake.pendingCopies.push({ id: 'copy' + (Object.keys(fake.files).length + fake.pendingCopies.length), name: body.name, parentId: body.parentReference.id, polls: fake.copyPolls });
      return new Response(null, { status: 202 });            // no body, as in the real service
    }
    if (method === 'GET' && (m = path.match(/^\/drives\/d\/items\/([^/]+):\/(.+)$/))) {
      for (const c of [...fake.pendingCopies]) {
        if (c.parentId === m[1] && c.name === m[2] && c.polls-- <= 0) {
          fake.files[c.id] = { id: c.id, name: c.name, parentId: c.parentId };
          fake.pendingCopies.splice(fake.pendingCopies.indexOf(c), 1);
        }
      }
      const f = Object.values(fake.files).find(x => x.parentId === m[1] && x.name === m[2]);
      return f ? reply(item(f)) : reply({ error: { message: 'itemNotFound' } }, 404);
    }
    if (method === 'GET' && (m = path.match(/^\/drives\/d\/items\/([^/]+)$/))) {
      const f = fake.files[m[1]];
      return f ? reply(item(f)) : reply({ error: { message: 'itemNotFound' } }, 404);
    }
    throw new Error(`fake Graph: unhandled drive request ${method} ${path}`);
  }

  // fault: { when(method, path), status?, network?, apply?, times? }
  //   apply: true  = the server did the work, then failed to say so
  fake.failNext = fault => fake.faults.push({ times: 1, ...fault });
  fake.fetch = async (url, init = {}) => {
    const method = init.method || 'GET';
    const body = init.body ? JSON.parse(init.body) : undefined;
    if (!String(url).includes('/workbook')) {
      const path = drivePath(url);
      fake.calls.push(`${method} ${path}`);
      return driveRoute(method, path, body);
    }
    const path = decodeURIComponent(String(url).split('/workbook')[1] || '');
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
  store.file = { driveId: 'd', itemId: 'i', parentId: 'p', name: 'Bookkeeping 2026.xlsx' };
  store.session = 'session-1';
  globalThis.fetch = fake.fetch;
  await store.loadTable('Tasks');
  return store;
}
