// Data layer: reads/writes Excel Tables in the workbook via Microsoft Graph.
// Every write is to a named Excel Table, by column header, so the workbook's
// own formulas and sheets keep working unchanged.

const GRAPH = 'https://graph.microsoft.com/v1.0';
const SCOPES = ['User.Read', 'Files.ReadWrite.All'];
const LAST_USER = 'books:lastUser';     // account to sign back in automatically
const AUTO = 'books:autoSignIn';        // set while an automatic sign-in is under way

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
    let res = null;
    try { res = await this.msal.handleRedirectPromise(); }
    catch (e) {
      // An automatic sign-in that Microsoft couldn't finish just falls back to the button
      if (!sessionStorage.getItem(AUTO)) throw e;
    }
    if (res?.account) this.msal.setActiveAccount(res.account);
    let acct = this.msal.getActiveAccount() || this.msal.getAllAccounts()[0];

    // MSAL keeps its saved sign-in encrypted with a key that is dropped when the
    // browser or home-screen app closes, so a fresh open has no account. If you
    // have signed in here before, sign in again automatically: Microsoft still
    // has its own session (the "stay signed in" one), so this needs no password
    // until your organization's session policy asks for one.
    const hint = localStorage.getItem(LAST_USER);
    if (!acct && hint) {
      try {
        const r = await this.msal.ssoSilent({ scopes: SCOPES, loginHint: hint }); // works where 3rd-party cookies are allowed
        acct = r.account;
      } catch {
        if (!sessionStorage.getItem(AUTO)) {   // once per app session, so it can't loop
          sessionStorage.setItem(AUTO, '1');
          await this.msal.loginRedirect({ scopes: SCOPES, loginHint: hint });
          return new Promise(() => {});          // the page is navigating to Microsoft
        }
      }
    }
    if (!acct) return null;
    sessionStorage.removeItem(AUTO);
    this.msal.setActiveAccount(acct);
    this.account = acct;
    localStorage.setItem(LAST_USER, acct.username);
    return { name: acct.name, email: acct.username.toLowerCase() };
  }
  login() { sessionStorage.removeItem(AUTO); return this.msal.loginRedirect({ scopes: SCOPES, prompt: 'select_account' }); }
  logout() { localStorage.removeItem(LAST_USER); return this.msal.logoutRedirect({ account: this.account }); }

  async token() {
    try {
      const r = await this.msal.acquireTokenSilent({ scopes: SCOPES, account: this.account });
      return r.accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) {
        await this.msal.acquireTokenRedirect({ scopes: SCOPES, loginHint: this.account?.username });
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
      // 429 means throttled before anything happened, so repeating is always safe.
      // 503/504 can arrive after Excel already did the work, so only repeat
      // requests that are safe to run twice (opts.once marks the ones that aren't).
      if (r.status === 429 || ([503, 504].includes(r.status) && !opts.once)) {
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
      const err = new Error(`${r.status}: ${msg}`);
      err.status = r.status;
      throw err;
    }
    throw new Error('Microsoft Graph is busy, please try again in a minute.');
  }

  get wb() { return `/drives/${this.file.driveId}/items/${this.file.itemId}/workbook`; }
  wbApi(method, path, body, opts = {}) { return this.api(method, this.wb + path, body, { workbook: true, ...opts }); }

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
    let res;
    try {
      res = await this.wbApi('POST', `/tables('${encodeURIComponent(tableName)}')/rows/add`, { index: null, values: [values] }, { once: true });
    } catch (e) {
      if (![503, 504].includes(e.status)) throw e;
      // Timed out: Excel may or may not have added the row. Look for it among
      // the rows that are new since we last read the table before deciding.
      const before = t.rows.length;
      const fresh = await this.loadTable(tableName);
      const arrived = fresh.rows.slice(before).find(r => values.every((v, i) => v === null || (typeof v === 'string' && v.startsWith('=')) || String(r._raw[i]) === String(v)));
      if (!arrived) throw new Error(`Excel timed out (${e.status}) and the row has not appeared. Wait a minute, press ↻ and check the list before saving again.`);
      res = { index: arrived._i };
    }
    // From here on the row is in the workbook. A problem re-reading it must not
    // be reported as a failed save: the natural reaction is to save again, which
    // would add the row twice. Report it as saved-but-unverified instead.
    const rowNum = t.startRow + 1 + res.index;
    const out = { sheet: t.sheet, row: rowNum, address: `${t.sheet}!${rowAddress(t, rowNum)}` };
    try {
      t = await this.loadTable(tableName);
      await this.repairComputed(tableName, t.rows[res.index], Object.keys(obj));
    } catch (e) { out.unverified = e.message || String(e); }
    return out;
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
    try {
      await this.wbApi('POST', this.rangeApi(t.sheet, rowAddress(t, row._row)) + '/delete', { shift: 'Up' }, { once: true });
    } catch (e) {
      if (![503, 504].includes(e.status)) throw e;
      // Timed out: deleting again could remove the row that moved up into its
      // place. If the row is no longer there, Excel did delete it.
      let stillThere = true;
      try { await this.verify(tableName, row); } catch (v) { if (/changed in Excel/.test(v.message)) stillThere = false; else throw e; }
      if (stillThere) throw new Error(`Excel timed out (${e.status}) and the row is still there. Press ↻ and try again.`);
    }
    await this.loadTable(tableName);
    return { address: `${t.sheet}!${rowAddress(t, row._row)}` };
  }

  // Write cells on many rows: items = [{ row, values: { Header: value, … } }].
  // Rows are addressed by sheet row number, so first re-read the table and make
  // sure every target row is still where, and what, it was when the caller's
  // list was built: a sort or a deleted row in Excel moves them. If anything
  // differs, nothing is written. Consecutive rows go in a single request.
  async setRows(tableName, items, onProgress) {
    const computed = this.schema[tableName]?.computed || [];
    const fresh = await this.loadTable(tableName);
    const headers = [...new Set(items.flatMap(it => Object.keys(it.values)))];
    for (const h of headers) if (!fresh.headers.includes(h)) throw new Error(`Column "${h}" not found in table ${tableName}`);
    for (const { row } of items) {
      const live = fresh.rows[row._row - fresh.startRow - 1];
      const same = live && fresh.headers.every((h, i) => computed.includes(h) || String(live._raw[i]) === String(row._raw[i]));
      if (!same) throw new Error(`Row ${row._row} of "${fresh.sheet}" changed in Excel since this list was loaded (rows may have been sorted, added or deleted). Nothing was written. Press ↻ and try again.`);
    }
    const sorted = [...items].sort((a, b) => a.row._row - b.row._row);
    let done = 0;
    for (const header of headers) {
      const col = colName(fresh.startCol + fresh.headers.indexOf(header));
      const todo = sorted.filter(it => header in it.values);
      for (let k = 0; k < todo.length;) {
        let j = k;
        while (j + 1 < todo.length && todo[j + 1].row._row === todo[j].row._row + 1) j++;
        const block = todo.slice(k, j + 1);
        await this.wbApi('PATCH', this.rangeApi(fresh.sheet, `${col}${block[0].row._row}:${col}${block[block.length - 1].row._row}`), { values: block.map(b => [b.values[header]]) });
        done += block.length;
        onProgress?.(Math.min(items.length, Math.ceil(done / headers.length)));
        k = j + 1;
      }
    }
    await this.loadTable(tableName);
  }
  // one column, a value per row
  setCells(tableName, header, items, onProgress) {
    return this.setRows(tableName, items.map(({ row, value }) => ({ row, values: { [header]: value } })), onProgress);
  }
  // one column, the same value on every row
  setColumn(tableName, rowsToSet, header, value, onProgress) {
    return this.setRows(tableName, rowsToSet.map(row => ({ row, values: { [header]: value } })), onProgress);
  }
  // several columns, the same values on every row; written in the order given
  setColumns(tableName, rowsToSet, values, onProgress) {
    return this.setRows(tableName, rowsToSet.map(row => ({ row, values })), onProgress);
  }

  // Add an empty column at the right edge of a table (Excel does the insert).
  async addColumn(tableName, name) {
    const t = this.tables[tableName] || await this.loadTable(tableName);
    if (t.headers.includes(name)) return;
    await this.wbApi('POST', `/tables('${encodeURIComponent(tableName)}')/columns/add`, { index: null, values: [[name], ...t.rows.map(() => [''])] }, { once: true });
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
