// Data layer: reads/writes Excel Tables in the workbook via Microsoft Graph.
// Every write is to a named Excel Table, by column header, so the workbook's
// own formulas and sheets keep working unchanged.

const GRAPH = 'https://graph.microsoft.com/v1.0';
const SCOPES = ['User.Read', 'Files.ReadWrite.All'];

export class GraphStore {
  constructor(cfg, schema) {
    this.cfg = cfg;
    this.schema = schema;
    this.tables = {};      // name -> { sheet, startCol, startRow, headers, rows }
    this.session = null;
    this.file = null;      // { driveId, itemId, parentId, name, webUrl }
  }

  // ---------- auth ----------
  async init() {
    this.msal = new msal.PublicClientApplication({
      auth: {
        clientId: this.cfg.clientId,
        authority: `https://login.microsoftonline.com/${this.cfg.tenantId}`,
        redirectUri: location.origin + location.pathname,
      },
      cache: { cacheLocation: 'localStorage' },
    });
    await this.msal.initialize();
    const res = await this.msal.handleRedirectPromise();
    if (res?.account) this.msal.setActiveAccount(res.account);
    const acct = this.msal.getActiveAccount() || this.msal.getAllAccounts()[0];
    if (!acct) return null;
    this.msal.setActiveAccount(acct);
    this.account = acct;
    return { name: acct.name, email: acct.username.toLowerCase() };
  }
  login() { return this.msal.loginRedirect({ scopes: SCOPES, prompt: 'select_account' }); }
  logout() { return this.msal.logoutRedirect({ account: this.account }); }

  async token() {
    try {
      const r = await this.msal.acquireTokenSilent({ scopes: SCOPES, account: this.account });
      return r.accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) {
        await this.msal.acquireTokenRedirect({ scopes: SCOPES });
      }
      throw e;
    }
  }

  // ---------- low-level request with retry ----------
  async api(method, path, body, opts = {}) {
    const url = path.startsWith('http') ? path : GRAPH + path;
    for (let attempt = 0; attempt < 4; attempt++) {
      const headers = { Authorization: 'Bearer ' + await this.token(), ...(opts.headers || {}) };
      if (opts.workbook && this.session) headers['workbook-session-id'] = this.session;
      let payload = body;
      if (body && !(body instanceof Blob)) { headers['Content-Type'] = 'application/json'; payload = JSON.stringify(body); }
      const r = await fetch(url, { method, headers, body: payload });
      if (r.ok) return r.status === 204 ? null : r.json();

      const text = await r.text();
      if ([429, 503, 504].includes(r.status)) {
        const wait = Number(r.headers.get('Retry-After') || 2) * 1000;
        await new Promise(res => setTimeout(res, wait));
        continue;
      }
      if (opts.workbook && /session/i.test(text) && attempt === 0) {
        await this.createSession();   // session expired: make a new one and retry once
        continue;
      }
      let msg = text;
      try { msg = JSON.parse(text).error.message; } catch {}
      throw new Error(`${r.status}: ${msg}`);
    }
    throw new Error('Microsoft Graph is busy, please try again in a minute.');
  }

  get wb() { return `/drives/${this.file.driveId}/items/${this.file.itemId}/workbook`; }
  wbApi(method, path, body) { return this.api(method, this.wb + path, body, { workbook: true }); }

  // ---------- file ----------
  async openFile(shareUrl) {
    if (!shareUrl) throw new Error('No workbook link configured. Add it in Settings.');
    const enc = 'u!' + btoa(unescape(encodeURIComponent(shareUrl.trim())))
      .replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
    const item = await this.api('GET', `/shares/${enc}/driveItem?$select=id,name,webUrl,parentReference`);
    this.file = {
      driveId: item.parentReference.driveId,
      itemId: item.id,
      parentId: item.parentReference.id,
      name: item.name,
      webUrl: item.webUrl,
    };
    await this.createSession();
    return this.file;
  }

  async createSession() {
    this.session = null;
    const s = await this.api('POST', this.wb + '/createSession', { persistChanges: true });
    this.session = s.id;
  }

  // ---------- tables ----------
  async loadTable(name) {
    const r = await this.wbApi('GET', `/tables('${encodeURIComponent(name)}')/range?$select=address,values`);
    const t = parseTable(r.address, r.values);
    this.tables[name] = t;
    return t;
  }

  async loadAll(names) {
    // a few at a time: Excel Online throttles parallel calls to one workbook
    const queue = [...names];
    const worker = async () => { while (queue.length) await this.loadTable(queue.shift()); };
    await Promise.all([worker(), worker(), worker()]);
  }

  rangeApi(sheet, address) {
    return `/worksheets('${encodeURIComponent(sheet)}')/range(address='${address}')`;
  }

  // Read a plain cell range, e.g. 'Mileage!Q22:R26'
  async readRange(ref) {
    const [, sheet, addr] = ref.match(/^'?(.*?)'?!(.+)$/);
    const r = await this.wbApi('GET', this.rangeApi(sheet, addr) + '?$select=values');
    return r.values;
  }

  // Confirm the row on the server still matches what we displayed before
  // changing it (guards against the sheet being sorted/edited in Excel meanwhile).
  async verify(tableName, row) {
    const t = this.tables[tableName];
    const addr = rowAddress(t, row._row);
    const r = await this.wbApi('GET', this.rangeApi(t.sheet, addr) + '?$select=values');
    const computed = this.schema[tableName]?.computed || [];
    const live = r.values[0];
    const same = t.headers.every((h, i) => computed.includes(h) || String(live[i]) === String(row._raw[i]));
    if (!same) throw new Error(`Row ${row._row} of "${t.sheet}" changed in Excel since it was loaded. Refresh and try again.`);
  }

  async addRow(tableName, obj) {
    let t = this.tables[tableName] || await this.loadTable(tableName);
    const values = t.headers.map(h => (h in obj ? obj[h] : null));
    const res = await this.wbApi('POST', `/tables('${encodeURIComponent(tableName)}')/rows/add`, { index: null, values: [values] });
    t = await this.loadTable(tableName);
    const row = t.rows[res.index];
    await this.repairComputed(tableName, row, Object.keys(obj));
    return { sheet: t.sheet, row: row._row, address: `${t.sheet}!${rowAddress(t, row._row)}` };
  }

  // If Excel didn't auto-fill a calculated column for a new row, copy the
  // formula from the row above (R1C1 keeps relative references correct).
  async repairComputed(tableName, row, provided = []) {
    const computed = (this.schema[tableName]?.computed || []).filter(h => !provided.includes(h));
    const t = this.tables[tableName];
    if (!computed.length || row._row <= t.startRow + 1) return;
    const cur = await this.wbApi('GET', this.rangeApi(t.sheet, rowAddress(t, row._row)) + '?$select=formulasR1C1');
    const prev = await this.wbApi('GET', this.rangeApi(t.sheet, rowAddress(t, row._row - 1)) + '?$select=formulasR1C1');
    for (const h of computed) {
      const i = t.headers.indexOf(h);
      if (i < 0) continue;
      const have = String(cur.formulasR1C1[0][i] || '');
      const want = String(prev.formulasR1C1[0][i] || '');
      if (!have.startsWith('=') && want.startsWith('=')) {
        const cell = colName(t.startCol + i) + row._row;
        await this.wbApi('PATCH', this.rangeApi(t.sheet, cell), { formulasR1C1: [[want]] });
      }
    }
  }

  // Writes only editable columns, in contiguous runs. Formula columns are
  // skipped unless a change for them is passed explicitly (e.g. custom tax).
  async updateRow(tableName, row, changes) {
    const t = this.tables[tableName];
    await this.verify(tableName, row);
    const computed = this.schema[tableName]?.computed || [];
    const runs = [];
    let run = null;
    t.headers.forEach((h, i) => {
      if (computed.includes(h) && !(h in changes)) { run = null; return; }
      if (!run) { run = { start: i, values: [] }; runs.push(run); }
      run.values.push(h in changes ? changes[h] : row._raw[i]);
    });
    for (const r of runs) {
      const changed = r.values.some((v, k) => v !== row._raw[r.start + k]);
      if (!changed) continue;
      const a = colName(t.startCol + r.start) + row._row + ':' + colName(t.startCol + r.start + r.values.length - 1) + row._row;
      // a value like "=…" must be sent as a formula
      const key = r.values.some(v => typeof v === 'string' && v.startsWith('=')) ? 'formulas' : 'values';
      await this.wbApi('PATCH', this.rangeApi(t.sheet, a), { [key]: [r.values] });
    }
    await this.loadTable(tableName);
    return { address: `${t.sheet}!${rowAddress(t, row._row)}` };
  }

  async deleteRow(tableName, row) {
    const t = this.tables[tableName];
    await this.verify(tableName, row);
    await this.wbApi('POST', this.rangeApi(t.sheet, rowAddress(t, row._row)) + '/delete', { shift: 'Up' });
    await this.loadTable(tableName);
    return { address: `${t.sheet}!${rowAddress(t, row._row)}` };
  }

  // Write one column on many rows. Consecutive rows go in a single request.
  async setCells(tableName, header, items, onProgress) {
    const t = this.tables[tableName];
    const i = t.headers.indexOf(header);
    if (i < 0) throw new Error(`Column "${header}" not found in table ${tableName}`);
    const col = colName(t.startCol + i);
    const sorted = [...items].sort((a, b) => a.row._row - b.row._row);
    let done = 0;
    for (let k = 0; k < sorted.length;) {
      let j = k;
      while (j + 1 < sorted.length && sorted[j + 1].row._row === sorted[j].row._row + 1) j++;
      const block = sorted.slice(k, j + 1);
      await this.wbApi('PATCH', this.rangeApi(t.sheet, `${col}${block[0].row._row}:${col}${block[block.length - 1].row._row}`), { values: block.map(b => [b.value]) });
      done += block.length;
      onProgress?.(done);
      k = j + 1;
    }
    await this.loadTable(tableName);
  }
  setColumn(tableName, rowsToSet, header, value, onProgress) {
    return this.setCells(tableName, header, rowsToSet.map(row => ({ row, value })), onProgress);
  }

  // Add an empty column at the right edge of a table (Excel does the insert).
  async addColumn(tableName, name) {
    const t = this.tables[tableName] || await this.loadTable(tableName);
    if (t.headers.includes(name)) return;
    await this.wbApi('POST', `/tables('${encodeURIComponent(tableName)}')/columns/add`, { index: null, values: [[name], ...t.rows.map(() => [''])] });
    await this.loadTable(tableName);
  }

  // ---------- files ----------
  async uploadFile(blob, fileName) {
    const folder = (this.cfg.receiptFolder || 'Receipts').split('/').map(encodeURIComponent).join('/');
    const path = `/drives/${this.file.driveId}/items/${this.file.parentId}:/${folder}/${encodeURIComponent(fileName)}:/content?@microsoft.graph.conflictBehavior=rename`;
    const r = await this.api('PUT', path, blob, { headers: { 'Content-Type': blob.type || 'image/jpeg' } });
    return { name: r.name, webUrl: r.webUrl };
  }
}

// ---------- helpers shared with the demo store ----------
export function parseTable(address, values) {
  // address like  Time!A1:I754  or  'Main Act'!A1:H13
  const m = address.match(/^'?(.*?)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/);
  const sheet = m[1].replace(/''/g, "'");
  const startCol = colIndex(m[2]);
  const startRow = Number(m[3]);
  const headers = values[0].map(String);
  const rows = values.slice(1).map((raw, i) => {
    const o = { _i: i, _row: startRow + 1 + i, _raw: raw };
    headers.forEach((h, k) => { o[h] = raw[k]; });
    return o;
  });
  return { sheet, startCol, startRow, headers, rows };
}
export function colIndex(letters) { return [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0); }
export function colName(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
export function rowAddress(t, rowNum) { return `${colName(t.startCol)}${rowNum}:${colName(t.startCol + t.headers.length - 1)}${rowNum}`; }
