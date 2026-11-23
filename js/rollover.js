// Starting a new year: what carries over, and what the new workbook is called.
// Pure functions, no workbook access, so they can be tested on their own.

const num = v => Number(v) || 0;
export const round2 = n => Math.round((n + Number.EPSILON) * 100) / 100;

// What a fixed-fee project has already spent, in dollars, as far as this
// workbook knows. The Projects table's Total is billed (capped at the fee, and
// including last year's Carryover) plus billable time not yet invoiced; Overage
// holds hours billed beyond the fee, which Total leaves out. Everything the fee
// has been used for is Total + Overage. Carrying that into next year's
// Carryover column keeps its "previously billed", % complete and hours
// remaining right. Time worked in December but invoiced in January is counted
// here, because that invoice is recorded in this workbook.
export function carryoverFor(p) {
  if (p.Type !== 'Fixed Fee' || !(num(p.Fee) > 0)) return '';
  const spent = round2(num(p.Total) + num(p.Overage));
  return spent > 0 ? spent : '';
}

// One row per project for the preview. `include` says whether it is carried over
// into the new workbook (Lost and Archive projects are not, by default).
export function planProjects(projects, keepStatuses) {
  return projects.map(p => ({
    name: p.Name, id: p.ProjectID, client: p.Client, status: p.Status, type: p.Type, fee: num(p.Fee),
    spent: p.Type === 'Fixed Fee' ? round2(num(p.Total) + num(p.Overage)) : 0,
    carryover: carryoverFor(p),
    include: keepStatuses.includes(p.Status),
  }));
}

// "Bookkeeping 2026.xlsx" -> "Bookkeeping 2027.xlsx"; a name with no year gets one added.
export function newFileName(name, toYear) {
  const n = String(name || 'Bookkeeping.xlsx');
  return /20\d\d/.test(n) ? n.replace(/20\d\d/, String(toYear)) : n.replace(/(\.xlsx)?$/i, ` ${toYear}.xlsx`);
}

// Which projects in `target` need a different Carryover than `source` now says?
// Returns the changes (by project name) and the source projects the target lacks.
export function carryoverChanges(sourceProjects, targetProjects) {
  const changes = [], missing = [];
  for (const p of sourceProjects) {
    const want = carryoverFor(p);
    if (want === '') continue;
    const t = targetProjects.find(x => x.Name === p.Name);
    if (!t) { missing.push(p.Name); continue; }
    const have = t.Carryover === '' || t.Carryover == null ? '' : round2(num(t.Carryover));
    if (have !== want) changes.push({ name: p.Name, row: t, from: have, to: want });
  }
  return { changes, missing };
}
