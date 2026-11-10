// Minimal test harness: no dependencies, runs in the browser.
export const tests = [];
export const test = (name, fn) => tests.push({ name, fn });

export function assert(cond, msg) { if (!cond) throw new Error(msg || 'assertion failed'); }
export function equal(got, expected, msg) {
  const a = JSON.stringify(got), b = JSON.stringify(expected);
  if (a !== b) throw new Error(`${msg || 'values differ'}: got ${a}, expected ${b}`);
}
export const sleep = ms => new Promise(r => setTimeout(r, ms));
export async function until(cond, ms = 5000, what = 'condition') {
  const t0 = Date.now();
  for (;;) {
    const v = await cond();
    if (v) return v;
    if (Date.now() - t0 > ms) throw new Error('timed out waiting for ' + what);
    await sleep(20);
  }
}
export async function rejects(fn, pattern, msg) {
  let err = null;
  try { await fn(); } catch (e) { err = e; }
  if (!err) throw new Error(msg || 'expected an error, but the call succeeded');
  if (pattern && !pattern.test(String(err.message || err))) throw new Error(`error "${err.message}" does not match ${pattern}`);
  return err;
}
