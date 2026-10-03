// node tests/run-tests.js  -- synthetic cases for every rule (no real data needed)
const C = require('../config.js'), P = require('../core.js'); let fail = 0;
const ok = (c, m) => { console.log((c ? 'PASS ' : 'FAIL ') + m); if (!c) fail++; };
const src = n => C.sources.find(s => s.sheet === n), PRI = C.sources[0], SEPT = src('Sept. PRS'), JULY = src('July PRS'), AV2 = C.sources.find(s => s.workflow === 'AV2'), SVG = C.sources.find(s => s.workflow === 'SVm-GAP');
const PH = Object.values(PRI.columns), FH = Object.values(SEPT.columns);
// primary row: [project,task,subtask,id,name,date,start,end,target,units,hours,prod,var,remarks]
const pr = (o = {}) => { const d = { subtask: 'Alert Validation', id: '1', name: 'Ann', date: '2026-09-01', start: '07:30', end: '08:30', target: 0, units: 300, hours: 1, prod: 300, variance: 5, remarks: '', ...o };
  return ['Proj', 'Label Cloud', d.subtask, d.id, d.name, d.date, d.start, d.end, d.target, d.units, d.hours, d.prod, d.variance, d.remarks]; };
// PRS row: [project,task,subtask,id,name,date,target,units,hours,prod,variance,remarks]
const fr = (o = {}) => { const d = { subtask: 'Alert Validation', id: '1', name: 'Ann', date: '14/09/2026', target: 240, units: 300, hours: 1, prod: 300, variance: '25%', remarks: '', ...o };
  return ['Proj', 'Label Cloud', d.subtask, d.id, d.name, d.date, d.target, d.units, d.hours, d.prod, d.variance, d.remarks]; };
const AVH = ['Employee ID', 'DATE', 'USERNAME', 'Productivity', 'Mins. time', 'Avarage time()', 'Duration'];
const av = (o = {}) => { const d = { id: '1', date: '8/5', email: 'ann@x.com', units: 120, hours: '1.5', ...o }; return [d.id, d.date, d.email, d.units, '1s', '10', d.hours]; };
const SVH = ['Employee ID', 'Date', 'Email Address', ' Productivity', 'Duration'];
const sv = (o = {}) => { const d = { id: '', date: '8/6/2026', email: 'ann@x.com', units: 40, hours: '0:45:00', ...o }; return [d.id, d.date, d.email, d.units, d.hours]; };
const LK = [['', '', '', '', '', '', '', 'Employee Email Address', 'Employee ID'], ['', '', '', '', '', '', '', 'Ann@x.com', '1']];
const ix = (p = [], f = [], extra = {}) => P.buildIndex(C, { [PRI.sheet]: [PH, ...p], [SEPT.sheet]: [FH, ...f], ...extra });
const q = (i, d = '2026-09-01') => P.query(i, '1', d);

let r = q(ix([pr()])); ok(r.rows.length === 1 && P.summarize(r.rows).units === 300, 'one workflow');
r = q(ix([pr(), pr({ subtask: 'GC-LC', start: '08:45', end: '09:45', units: 50 })])); let t = P.summarize(r.rows); ok(t.workflows === 2 && t.units === 350 && t.hours === 2, 'multiple workflows totals');
r = q(ix([pr(), pr({ start: '09:00', end: '10:00', units: 100 })])); ok(r.rows.length === 2 && !r.rows[0].conflict, 'same workflow, different times = two real records');
ok(q(ix([pr()]), '2026-09-02').mode === 'none', 'no records on date');
r = q(ix([pr(), pr()])); ok(r.rows.length === 1 && !r.conflicts.length, 'exact duplicate deduplicated');
r = q(ix([pr(), pr({ units: 301 })])); t = P.summarize(r.rows); ok(r.rows.every(x => x.conflict) && r.conflicts.length === 1 && t.units === 0 && t.excluded === 2, 'same ids, different units => conflict, excluded from totals');
ok(P.summarize(r.rows, true).units === 601, 'conflict rows can be included on request');
r = q(ix([pr(), pr({ variance: 9 })])); ok(r.rows.length === 1 && !r.rows[0].conflict && r.softDups.length === 1 && P.summarize(r.rows).units === 300, 'same activity, only variance differs => counted once and noted');
r = q(ix([pr({ start: '', end: '', hours: '' })])); ok(r.rows[0].start === null && r.rows[0].hours === null && P.summarize(r.rows).missingHours === 1, 'missing times: nothing invented');
r = q(ix([pr({ start: '07:30', end: '08:30', hours: '' })])); ok(r.rows[0].hours === 1 && /calculated/.test(r.rows[0].from.hours), 'missing duration calculated from times and labelled');
r = q(ix([pr({ prod: '' })])); ok(r.rows[0].prod === null, 'missing productivity stays empty');
ok(P.parseDate('2026-09-01').iso === '2026-09-01' && P.parseDate('01/09/2026').iso === '2026-09-01' && P.parseDate(46266).iso === '2026-09-01' && P.parseDate('31/02/2026') === null, 'date formats (ISO, DD/MM, serial, invalid)');
ok(P.parseDate('2026-01-09', 'isoDayMonthSwapped').iso === '2026-09-01' && P.parseDate('2026-01-09', 'isoDayMonthSwapped').repaired, 'swapped ISO date repaired + flagged');
ok(P.parseDate('8/1/2026', 'mdy').iso === '2026-08-01' && P.parseDate('7/31', 'mdy', 2026).iso === '2026-07-31' && P.parseDate('7/31', 'mdy') === null && P.parseDate('8/1/2026 Total', 'mdy') === null, 'US dates, dates without a year, "Total" rows ignored');
ok(P.parseDuration('1:15:00') === 1.25 && P.parseDuration('25:30') === 25.5 && P.parseDuration('1.5') === 1.5 && P.parseDuration('-0:10:00') === null && P.parseDuration('12/30/1899') === null, 'durations: clock, decimal, negative/invalid');
r = q(ix([], [fr()]), '2026-09-14'); ok(r.mode === 'prs' && r.rows.length === 1, 'PRS sheet used only when Detailed Productivity has no record');
r = q(ix([pr({ date: '2026-09-14' })], [fr()]), '2026-09-14'); ok(r.mode === 'primary' && r.rows[0].target === 240 && r.rows[0].variance === 25 && r.rows[0].from.target === 'Sept. PRS', 'target+variance filled from PRS (labelled)');
r = q(ix([pr({ date: '2026-09-14' })], []), '2026-09-14'); ok(r.rows[0].target === null && r.rows[0].variance === null, 'no PRS match: variance with target 0 hidden, not invented');
r = q(ix([pr({ date: '2026-09-14', subtask: 'AV2' })], [fr()]), '2026-09-14'); ok(r.rows[0].target === null, 'AV2 not treated as Alert Validation (no alias confirmed)');
// ---- all months, all sheets ----
const jul = [FH, fr({ date: '03/07/2026', units: 80 }), fr({ date: '04/07/2026', units: 90 })];
let I = ix([pr()], [fr()], { [JULY.sheet]: jul }); r = q(I, '2026-07-03'); ok(r.mode === 'prs' && r.sources[0] === 'July PRS' && r.rows[0].units === 80, 'July PRS read for July days');
I = ix([], [], { [AV2.sheet]: [AVH, av()], [C.idLookup.sheet]: LK }); r = q(I, '2026-08-05'); ok(r.mode === 'logs' && r.rows[0].subtask === 'AV2' && r.rows[0].hours === 1.5 && r.rows[0].prod === 80 && /calculated/.test(r.rows[0].from.prod), 'workflow log used when no PRS row; productivity calculated and labelled');
I = ix([], [], { [SVG.sheet]: [['', '', '', '', ''], SVH, sv(), ['', '8/6/2026 Total', '', 40, '0:45:00']], [C.idLookup.sheet]: LK }); r = q(I, '2026-08-06');
ok(r.mode === 'logs' && r.rows.length === 1 && r.rows[0].empId === '1' && r.rows[0].hours === 0.75, 'header on row 2, missing ID filled from email (case-insensitive), Total row skipped');
I = ix([pr()], [fr({ date: '01/09/2026', units: 999 })], { [AV2.sheet]: [AVH, av({ date: '9/1', units: 5 })] }); r = q(I);
ok(r.mode === 'primary' && P.summarize(r.rows).units === 300, 'same day in several tiers: only the most detailed one counts (no double counting)');
I = ix([pr()], [fr()], { [JULY.sheet]: jul, [AV2.sheet]: [AVH, av()], [C.idLookup.sheet]: LK }); const days = P.daily(I, '1'), T = P.totals(days);
ok(days.map(d => d.date).join() === '2026-07-03,2026-07-04,2026-08-05,2026-09-01,2026-09-14' && T.units === 80 + 90 + 120 + 300 + 300 && T.months.length === 3, 'all days overview across July, August and September');
ok(P.daily(I, '1', '2026-08-01', '2026-08-31').length === 1, 'overview can be limited to a date range');
ok(P.resolveSheet(['Realogram Correction HOURS from'], { sheet: 'Realogram Correction HOURS from 1st August' }) === 'Realogram Correction HOURS from'
  && P.resolveSheet(['SVm-RC Hours'], { sheet: 'SVm- RC Hours' }) === null, 'tab names: 31-character XLSX names match, different tabs do not');
console.log(fail ? `\n${fail} FAILED` : '\nall passed'); process.exit(fail ? 1 : 0);
