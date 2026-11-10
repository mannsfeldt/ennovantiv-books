import { tests } from './harness.js';
import './store-graph.test.js';
import './app.test.js';

const summary = document.querySelector('#summary'), list = document.querySelector('#results');
// The app tests clear a few of the app's saved values, so never run them on the live site.
if (!['localhost', '127.0.0.1'].includes(location.hostname)) {
  summary.textContent = 'Tests only run locally: python3 tests/serve.py';
} else {
  const only = new URLSearchParams(location.search).get('only');
  const results = [];
  for (const t of tests.filter(t => !only || t.name.includes(only))) {
    let error = null;
    try { await t.fn(); } catch (e) { error = String(e?.message || e); console.error(t.name, e); }
    results.push({ name: t.name, ok: !error, error });
    list.insertAdjacentHTML('beforeend', `<li class="${error ? 'fail' : 'ok'}"></li>`);
    list.lastElementChild.textContent = t.name;
    if (error) { const pre = document.createElement('pre'); pre.textContent = error; list.lastElementChild.appendChild(pre); }
    summary.textContent = `Running… ${results.length}`;
  }
  const failed = results.filter(r => !r.ok).length;
  summary.textContent = `${results.length - failed} passed, ${failed} failed`;
  window.TEST_RESULTS = { done: true, passed: results.length - failed, failed, results };
}
