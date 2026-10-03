// ALL source mappings + normalization rules live here. Change this file when the live sheet is confirmed.
// For each employee and day the app uses the most detailed tier that has data:
//   tier 1 = Detailed Productivity (has start/end times), tier 2 = PRS sheets, tier 3 = workflow logs (tool exports).
// Tiers are never added together for the same day, so the same work is not counted twice.
const PRS_COLUMNS = { project: 'Project', task: 'Task', subtask: 'SubTask', empId: 'Employee ID', empName: 'Employee Name', date: 'Completed Date',
  target: 'Target Productivity', units: 'Quantity', hours: 'Duration', prod: 'Productivity', variance: 'Percent Variance', remarks: 'Comment' };
const LOG_COLUMNS = { empId: 'Employee ID', date: 'DATE', email: 'USERNAME', units: 'Productivity', hours: 'Duration' };
const LOG2_COLUMNS = { empId: 'Employee ID', date: 'Date', email: 'Email Address', units: 'Productivity', hours: 'Duration' };

const PRS_CONFIG = {
  google: {
    spreadsheetId: '1qvUz1gkE8675_y-44DemQaEtSiUF-hpzwNHL4PFJk5Q', // the long ID in the sheet URL: docs.google.com/spreadsheets/d/<ID>/edit
    // spreadsheetId: '11Psd8-mcHOwv9uVUVd8ZnpWj_qMQToLK9HeaJDShd3E',
    clientId: '328807678456-0o4omh8c2icjnl0vt39qo59fi8fd2knv.apps.googleusercontent.com', // OAuth 2.0 *Web application* client ID (see SETUP.md). Not a secret; never put a client SECRET here.
    scope: 'https://www.googleapis.com/auth/spreadsheets.readonly' // read-only: the app can never change the sheet
  },
  // dateRule: 'dmy' = DD/MM/YYYY, 'mdy' = M/D/YYYY (US display), 'isoDayMonthSwapped' = YYYY-MM-DD stored with day/month swapped.
  // ISO (YYYY-MM-DD) is always accepted. year: used when a date is shown without a year (e.g. "7/1").
  // Tab names: long names match the 31-character names in XLSX exports automatically. Add other spellings to aliases.
  sources: [
    { tier: 1, sheet: 'Detailed Productivity', range: 'A:P', dateRule: 'dmy',
      columns: { project: 'Project Name', task: 'Task Name', subtask: 'Subtask Name', empId: 'Employee ID', empName: 'Employee Name',
        date: 'Date', start: 'Start Time', end: 'End Time', target: 'Expected Target (Units/Hour)', units: 'Total Units Completed',
        hours: 'Total Duration (Hours)', prod: 'Average Units/Hour Attained', variance: 'Percentage Variance', remarks: 'Remark Notes' } },
    { tier: 2, sheet: 'Sept. PRS', range: 'A:N', dateRule: 'isoDayMonthSwapped', columns: PRS_COLUMNS },
    { tier: 2, sheet: 'July PRS', range: 'A:N', dateRule: 'dmy', columns: PRS_COLUMNS },
    // Workflow logs: one row per person per day per tool. No target or variance; productivity per hour is calculated and labelled.
    { tier: 3, sheet: 'Alert Validation PRODUCTIVITY & HOURS', aliases: ['Alert Validation PRODUCTIVITY &'], label: 'AV2 log', workflow: 'AV2', task: 'Workflow log',
      range: 'A:I', dateRule: 'mdy', year: 2026, computeProd: true, columns: LOG_COLUMNS },
    { tier: 3, sheet: 'AVIP', label: 'AVIP log', workflow: 'AVIP', task: 'Workflow log', range: 'A:I', dateRule: 'mdy', year: 2026, computeProd: true, columns: LOG_COLUMNS },
    { tier: 3, sheet: 'LC-GC Hours', label: 'LC-GAP log', workflow: 'LC-GAP', task: 'Workflow log', range: 'A:I', dateRule: 'mdy', year: 2026, computeProd: true, columns: LOG_COLUMNS },
    { tier: 3, sheet: 'Realogram Correction HOURS from 1st August', label: 'LC-RC log', workflow: 'LC-RC', task: 'Workflow log', range: 'A:E', dateRule: 'mdy', computeProd: true, columns: LOG2_COLUMNS },
    { tier: 3, sheet: 'SVm- RC Hours', label: 'SVm-RC log', workflow: 'SVm-RC', task: 'Workflow log', range: 'A:E', dateRule: 'mdy', computeProd: true, columns: LOG2_COLUMNS },
    { tier: 3, sheet: 'SVm-Gap Hours', label: 'SVm-GAP log', workflow: 'SVm-GAP', task: 'Workflow log', range: 'A:E', dateRule: 'mdy', computeProd: true, columns: LOG2_COLUMNS }
  ],
  // Fills in the Employee ID for workflow-log rows where the ID cell is empty or #N/A (matched by email).
  idLookup: { sheet: 'Judith Analysis-New', range: 'H:I', columns: { email: 'Employee Email Address', empId: 'Employee ID' } },
  labelAliases: {} // e.g. { 'av2': 'alert validation' }. EMPTY on purpose: nothing confirmed as an alias yet.
};
if (typeof module !== 'undefined') module.exports = PRS_CONFIG;
