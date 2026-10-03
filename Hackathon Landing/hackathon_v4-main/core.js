// PRS Quick View core: parsing, dedupe/conflict detection, tiered sources, daily overview, summary. No DOM, no network.
(function (root) {
  const clean = s => String(s == null ? '' : s).replace(/\ufeff/g, '').replace(/\s+/g, ' ').trim();
  const PH = /^(|-|n\/a|#n\/a|na|null)$/i;
  const FIELDS = ['start', 'end', 'target', 'units', 'hours', 'prod', 'variance', 'remarks'];
  const TIER_NAMES = { 1: 'primary', 2: 'prs', 3: 'logs' };

  function num(v) { // '17%' -> 17 (percent as displayed); placeholders -> null
    if (v == null) return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    const s = clean(v); if (PH.test(s)) return null;
    if (!/^[-+]?[\d,]*\.?\d+%?$/.test(s)) return null;
    const n = parseFloat(s.replace(/,/g, '').replace('%', '')); return isNaN(n) ? null : n;
  }
  function ymd(y, m, d) {
    y = +y; m = +m; d = +d; const t = new Date(Date.UTC(y, m - 1, d));
    return (m >= 1 && m <= 12 && t.getUTCMonth() === m - 1 && t.getUTCDate() === d) ? `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}` : null;
  }
  // rule: 'dmy' (default, DD/MM/YYYY), 'mdy' (M/D/YYYY, US display), 'isoDayMonthSwapped'. ISO and serial numbers are always accepted.
  // year: used for dates shown without a year ("7/1"). Nothing else is guessed: anything unreadable -> null.
  function parseDate(v, rule, year) {
    if (v == null || clean(v) === '') return null;
    if (typeof v === 'number' && v > 20000) { const t = new Date(Math.round((v - 25569) * 864e5)); const iso = ymd(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()); return iso ? { iso } : null; }
    const s = clean(v); let m;
    if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) {
      if (rule === 'isoDayMonthSwapped') { const iso = ymd(m[1], m[3], m[2]); return iso ? { iso, repaired: true } : null; }
      const iso = ymd(m[1], m[2], m[3]); return iso ? { iso } : null;
    }
    if ((m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/))) { const iso = rule === 'mdy' ? ymd(m[3], m[1], m[2]) : ymd(m[3], m[2], m[1]); return iso ? { iso } : null; }
    if (year && (m = s.match(/^(\d{1,2})\/(\d{1,2})$/))) { const iso = rule === 'mdy' ? ymd(year, m[1], m[2]) : ymd(year, m[2], m[1]); return iso ? { iso } : null; }
    return null;
  }
  function parseTime(v) { // -> minutes since midnight | null
    if (v == null) return null;
    if (typeof v === 'number') return v >= 0 && v < 1 ? Math.round(v * 1440) : null;
    const m = clean(v).match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(am|pm)?$/i); if (!m) return null;
    let h = +m[1]; const mi = +m[2]; if (m[3]) h = (h % 12) + (/pm/i.test(m[3]) ? 12 : 0);
    return h < 24 && mi < 60 ? h * 60 + mi : null;
  }
  function parseDuration(v) { // -> hours | null. Accepts decimal hours (1.25) or clock durations (1:15:00, 25:10:00).
    if (v == null) return null;
    if (typeof v === 'number') return isFinite(v) && v >= 0 ? v : null;
    const s = clean(v), m = s.match(/^(\d+):(\d{2})(?::(\d{2}))?$/);
    if (m) return +m[1] + +m[2] / 60 + (+m[3] || 0) / 3600;
    const n = num(s); return n != null && n >= 0 ? n : null;
  }
  const fmtH = h => { const t = Math.round(h * 60); return `${Math.floor(t / 60)}h ${String(t % 60).padStart(2, '0')}m`; };
  const fmtT = m => m == null ? '—' : `${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;
  const fmtD = iso => iso.split('-').reverse().join('/');
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const fmtMonth = ym => `${MONTHS[+ym.slice(5, 7) - 1]} ${ym.slice(0, 4)}`;

  // Match a configured tab name to the real tab titles. XLSX exports cut tab names to 31 characters, so long names match by prefix.
  function resolveSheet(titles, src) {
    const names = [src.sheet, ...(src.aliases || [])], low = s => clean(s).toLowerCase();
    for (const n of names) if (titles.includes(n)) return n;
    for (const n of names) { const t = titles.find(x => low(x) === low(n)); if (t) return t; }
    for (const n of names) { const t = titles.find(x => (x.length >= 31 || n.length >= 31) && low(x.slice(0, 31)) === low(n.slice(0, 31))); if (t) return t; }
    return null;
  }
  function headerRow(rows, cols) { // the header is not always row 1 (some tabs start with a blank row)
    const want = cols.map(c => c.toLowerCase());
    for (let i = 0; i < Math.min(6, rows.length); i++) { const h = (rows[i] || []).map(c => clean(c).toLowerCase()); if (want.every(w => h.includes(w))) return i; }
    return 0;
  }
  function buildRecords(rows, sc, name, idOf) {
    const label = sc.label || name, out = (o = {}) => ({ records: [], errors: [], warnings: [], skipped: 0, label, ...o });
    if (!rows || !rows.length) return out({ warnings: [`"${name}" is empty or was not found`] });
    const C = sc.columns, h0 = headerRow(rows, [C.date, C.empId || C.email]), hdr = rows[h0].map(h => clean(h).toLowerCase()), idx = {}, warnings = [];
    for (const [f, col] of Object.entries(C)) { idx[f] = hdr.indexOf(col.toLowerCase()); if (idx[f] < 0) warnings.push(`"${name}": column "${col}" (${f}) not found`); }
    const hasId = idx.empId >= 0 || idx.email >= 0;
    if (!hasId || !(idx.date >= 0)) return out({ warnings, errors: [`"${name}" needs a date column and an Employee ID or email column, so it was not used`] });
    const records = []; let skipped = 0;
    rows.slice(h0 + 1).forEach((r, i) => {
      const g = f => idx[f] >= 0 ? r[idx[f]] : null, email = clean(g('email')).toLowerCase();
      let id = num(g('empId')); if (id == null && email && idOf) id = idOf(email);
      const d = parseDate(g('date'), sc.dateRule, sc.year);
      if (id == null || !d) { if (r.some(c => clean(c) !== '')) skipped++; return; }
      const rm = clean(g('remarks'));
      const rec = { source: name, sourceLabel: label, tier: sc.tier, row: h0 + i + 2, empId: String(Math.trunc(id)), empName: clean(g('empName')), email, date: d.iso, dateRepaired: !!d.repaired,
        project: clean(g('project')), task: sc.workflow ? (sc.task || '') : clean(g('task')), subtask: sc.workflow || clean(g('subtask')), start: parseTime(g('start')), end: parseTime(g('end')),
        target: num(g('target')), units: num(g('units')), hours: parseDuration(g('hours')), prod: num(g('prod')), variance: num(g('variance')), remarks: PH.test(rm) ? '' : rm, from: {}, notes: [] };
      if (rec.hours == null && rec.start != null && rec.end != null && rec.end > rec.start) { rec.hours = (rec.end - rec.start) / 60; rec.from.hours = 'calculated from start/end'; }
      if (sc.computeProd && rec.prod == null && rec.units != null && rec.hours > 0) { rec.prod = Math.round(rec.units / rec.hours * 100) / 100; rec.from.prod = 'calculated from units/duration'; }
      records.push(rec);
    });
    return out({ records, warnings, skipped });
  }
  // Exact duplicates (all FIELDS equal) collapse. Same key but different values => conflict (kept, flagged, never auto-resolved).
  function dedupe(recs, keyFn) {
    const g = new Map(); recs.forEach(r => { const k = keyFn(r); (g.get(k) || g.set(k, []).get(k)).push(r); });
    const records = [], conflicts = []; let removed = 0;
    for (const rows of g.values()) {
      const uniq = []; rows.forEach(r => { if (uniq.some(u => FIELDS.every(f => u[f] === r[f]))) removed++; else uniq.push(r); });
      const diff = uniq.length > 1 ? FIELDS.filter(f => new Set(uniq.map(u => u[f])).size > 1) : [];
      const info = { empId: uniq[0].empId, emp: uniq[0].empName, date: uniq[0].date, label: uniq[0].subtask, source: uniq[0].sourceLabel, rows: uniq.map(u => u.row), diff };
      if (uniq.length > 1 && diff.every(f => f === 'variance' || f === 'remarks')) { // same activity; only variance/remarks differ: count once, keep a visible note
        uniq[0].notes.push(`Duplicate rows ${info.rows.join(', ')} differ only in ${diff.join('/')}; counted once, first row shown`);
        records.push(uniq[0]); conflicts.push({ ...info, soft: true });
      } else if (uniq.length > 1) { uniq.forEach(u => { u.conflict = diff; records.push(u); }); conflicts.push(info); }
      else records.push(uniq[0]);
    }
    return { records, conflicts, removed };
  }
  function buildIndex(cfg, sheets) {
    const canon = l => (cfg.labelAliases[clean(l).toLowerCase()] || clean(l)).toLowerCase();
    const emailToId = new Map(); // email -> Employee ID, for workflow-log rows whose ID cell is empty or #N/A
    if (cfg.idLookup) {
      const rows = sheets[cfg.idLookup.sheet] || [], c = cfg.idLookup.columns, h0 = headerRow(rows, [c.email, c.empId]);
      const hdr = (rows[h0] || []).map(x => clean(x).toLowerCase()), ei = hdr.indexOf(c.email.toLowerCase()), ii = hdr.indexOf(c.empId.toLowerCase());
      if (ei >= 0 && ii >= 0) rows.slice(h0 + 1).forEach(r => { const e = clean(r[ei]).toLowerCase(), id = num(r[ii]); if (e && id != null && !emailToId.has(e)) emailToId.set(e, String(Math.trunc(id))); });
    }
    const idOf = e => emailToId.get(e) || null;
    const tiers = new Map(), conflicts = [], sources = [], errors = [], warnings = [], named = new Map(), mailOnly = new Map(); let skipped = 0;
    for (const sc of cfg.sources) {
      const B = buildRecords(sheets[sc.sheet], sc, sc.sheet, idOf);
      errors.push(...B.errors); warnings.push(...B.warnings); skipped += B.skipped;
      const keyFn = sc.tier === 1 ? r => [r.empId, r.date, canon(r.subtask), r.start, r.end].join('|') : r => [r.empId, r.date, canon(r.subtask)].join('|');
      const D = dedupe(B.records, keyFn); conflicts.push(...D.conflicts);
      sources.push({ sheet: sc.sheet, label: B.label, tier: sc.tier, loaded: !!(sheets[sc.sheet] && sheets[sc.sheet].length), records: D.records.length, skipped: B.skipped });
      const m = tiers.get(sc.tier) || tiers.set(sc.tier, new Map()).get(sc.tier);
      D.records.forEach(r => { const k = r.empId + '|' + r.date; (m.get(k) || m.set(k, []).get(k)).push(r);
        if (r.empName && !named.has(r.empId)) named.set(r.empId, r.empName); if (r.email && !mailOnly.has(r.empId)) mailOnly.set(r.empId, r.email); });
    }
    const emps = new Map(); [...new Set([...named.keys(), ...mailOnly.keys()])].forEach(id => emps.set(id, named.get(id) || mailOnly.get(id)));
    return { cfg, canon, emps, tiers, conflicts, sources, errors, warnings, skipped, emailToId };
  }
  function datesFor(ix, empId) {
    const s = new Set(); for (const m of ix.tiers.values()) for (const k of m.keys()) if (k.startsWith(empId + '|')) s.add(k.split('|')[1]);
    return [...s].sort();
  }
  function dateRange(ix, empId) { const d = datesFor(ix, empId); return d.length ? [d[0], d[d.length - 1]] : null; }

  // Uses the most detailed tier that has rows for that employee and day: 1 Detailed Productivity, 2 PRS sheets, 3 workflow logs.
  function query(ix, empId, date) {
    const k = empId + '|' + date, get = t => (ix.tiers.get(t) && ix.tiers.get(t).get(k)) || [];
    const copy = r => ({ ...r, from: { ...r.from }, notes: [...r.notes] }), fb = get(2);
    let tier = get(1).length ? 1 : fb.length ? 2 : get(3).length ? 3 : 0, rows;
    if (tier !== 1) rows = tier ? get(tier).map(copy) : [];
    else rows = get(1).map(r0 => { // enrich Detailed Productivity rows with target/variance from the PRS sheets, labelled
      const r = copy(r0), m = fb.filter(f => ix.canon(f.subtask) === ix.canon(r.subtask));
      const one = m.length === 1 ? m[0] : null, take = f => { if (one && one[f] != null) { r[f] = one[f]; r.from[f] = one.sourceLabel; } };
      if (!(r.target > 0)) { // primary target missing/0 => its variance cannot be verified, so it is not used
        const hadVar = r.variance != null; r.target = null; r.variance = null;
        if (m.length > 1) r.notes.push(`${m.length} PRS records match; target/variance not filled`); else if (!m.length) { if (hadVar) r.notes.push('Source variance hidden: no target in source'); } else { take('target'); take('variance'); }
      }
      ['units', 'hours', 'prod'].forEach(f => { if (r[f] == null) take(f); });
      if (!r.remarks && one && one.remarks) { r.remarks = one.remarks; r.from.remarks = one.sourceLabel; }
      return r;
    });
    rows.sort((a, b) => (a.start ?? 1e9) - (b.start ?? 1e9) || a.subtask.localeCompare(b.subtask));
    const srcs = new Set(rows.map(r => r.sourceLabel)), mine = ix.conflicts.filter(c => c.date === date && c.empId === empId && srcs.has(c.source));
    return { mode: tier ? TIER_NAMES[tier] : 'none', tier, rows, empId, empName: ix.emps.get(empId) || '', date,
      sources: [...srcs], repaired: rows.some(r => r.dateRepaired), conflicts: mine.filter(c => !c.soft), softDups: mine.filter(c => c.soft) };
  }
  function summarize(rows, includeConflicts) {
    const use = rows.filter(r => includeConflicts || !r.conflict), sum = k => use.reduce((a, r) => a + (r[k] || 0), 0);
    const iv = use.filter(r => r.start != null && r.end != null && r.end > r.start).map(r => [r.start, r.end]).sort((a, b) => a[0] - b[0]);
    let merged = 0, cur = null; iv.forEach(([s, e]) => { if (cur && s <= cur[1]) cur[1] = Math.max(cur[1], e); else { if (cur) merged += cur[1] - cur[0]; cur = [s, e]; } }); if (cur) merged += cur[1] - cur[0];
    return { n: use.length, units: sum('units'), hours: sum('hours'), workflows: new Set(use.map(r => r.subtask.toLowerCase())).size, mergedMin: use.length && iv.length === use.length ? merged : null,
      excluded: rows.length - use.length, missingUnits: use.filter(r => r.units == null).length, missingHours: use.filter(r => r.hours == null).length };
  }
  // One entry per day that has data in any source, for one employee. Optional from/to (ISO) limits the range.
  function daily(ix, empId, from, to) {
    return datesFor(ix, empId).filter(d => (!from || d >= from) && (!to || d <= to)).map(d => {
      const q = query(ix, empId, d), t = summarize(q.rows, false);
      return { date: d, mode: q.mode, sources: q.sources, n: t.n, units: t.units, hours: t.hours, workflows: [...new Set(q.rows.filter(r => !r.conflict).map(r => r.subtask))], excluded: t.excluded, conflicts: q.conflicts.length };
    });
  }
  function totals(days) {
    const months = new Map();
    days.forEach(d => { const k = d.date.slice(0, 7), m = months.get(k) || months.set(k, { month: k, days: 0, units: 0, hours: 0 }).get(k); m.days++; m.units += d.units; m.hours += d.hours; });
    return { days: days.length, units: days.reduce((a, d) => a + d.units, 0), hours: days.reduce((a, d) => a + d.hours, 0),
      workflows: new Set(days.flatMap(d => d.workflows.map(w => w.toLowerCase()))).size, months: [...months.values()] };
  }
  // Template summary: every number comes from the retrieved rows. No causes are ever suggested.
  function aiSummary(res, t) {
    const use = res.rows.filter(r => !r.conflict); if (!use.length) return res.rows.length ? 'All records for this day are in data-quality conflict, so no summary was generated.' : '';
    const names = [...new Set(use.map(r => r.subtask))], top = [...use].filter(r => r.units != null).sort((a, b) => b.units - a.units)[0];
    let s = `On ${fmtD(res.date)}, ${res.empName} has ${t.n} record${t.n > 1 ? 's' : ''} across ${t.workflows} workflow${t.workflows > 1 ? 's' : ''} (${names.join(', ')}), with ${t.units.toLocaleString()} units over ${fmtH(t.hours)} of logged time.`;
    if (top) s += ` The largest contribution by units was ${top.subtask} (${top.units.toLocaleString()}).`;
    const v = use.filter(r => r.variance != null && r.target > 0).map(r => `${r.subtask} ${r.variance >= 0 ? '+' : ''}${r.variance}% vs target ${r.target}/h`);
    if (v.length) s += ` Recorded variance: ${v.join('; ')}.`;
    if (t.missingUnits || t.missingHours) s += ` Note: ${Math.max(t.missingUnits, t.missingHours)} record(s) are missing units or duration, so totals may be understated.`;
    return s;
  }
  const api = { clean, num, parseDate, parseTime, parseDuration, resolveSheet, buildRecords, dedupe, buildIndex, datesFor, dateRange, query, summarize, daily, totals, aiSummary, fmtH, fmtT, fmtD, fmtMonth };
  if (typeof module !== 'undefined') module.exports = api; else root.PRS = api;
})(this);
