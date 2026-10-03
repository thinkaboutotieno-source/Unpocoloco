(function () {
  const $ = id => document.getElementById(id), cfg = PRS_CONFIG; let ix = null, last = null, view = null;
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const out = h => { $('out').innerHTML = h; };
  const alertBox = (cls, h) => `<div class="alert ${cls}">${h}</div>`;
  const wanted = () => [...cfg.sources, ...(cfg.idLookup ? [cfg.idLookup] : [])];
  const WD = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'], wd = iso => WD[new Date(iso + 'T00:00:00Z').getUTCDay()];

  function setData(sheets, label, missing) {
    ix = PRS.buildIndex(cfg, sheets); last = null; view = null;
    $('empList').innerHTML = [...ix.emps].sort((a, b) => a[1].localeCompare(b[1])).map(([id, n]) => `<option value="${esc(n)} (${id})">`).join('');
    const on = !ix.emps.size; $('find').disabled = on; $('all').disabled = on;
    const all = [...ix.tiers.values()].flatMap(m => [...m.keys()].map(k => k.split('|')[1])).sort();
    $('srcStatus').innerHTML = `${esc(label)}: ${ix.emps.size} employees` + (all.length ? `, ${PRS.fmtD(all[0])} – ${PRS.fmtD(all[all.length - 1])}` : '')
      + ` <button class="link" id="srcMore" type="button">Sources</button>`;
    $('srcMore').onclick = () => { const el = $('srcList'); el.hidden = !el.hidden; };
    $('srcList').innerHTML = `<table class="mini"><thead><tr><th>Tab</th><th>Used as</th><th class="n">Rows used</th><th class="n">Rows skipped</th></tr></thead><tbody>`
      + ix.sources.map(s => `<tr><td>${esc(s.sheet)}</td><td>${['', 'Main source', 'PRS sheet', 'Workflow log'][s.tier]}</td><td class="n">${s.loaded ? s.records.toLocaleString() : 'not found'}</td><td class="n">${s.loaded ? s.skipped.toLocaleString() : '—'}</td></tr>`).join('')
      + `</tbody></table><p class="muted">Skipped rows have no readable date or Employee ID, such as daily "Total" rows. For each person and day, the most detailed source with data is used, and sources are never added together.</p>`;
    $('srcList').hidden = true;
    const msgs = [...ix.errors, ...(missing || []).map(n => `Tab "${n}" was not found, so its data is not included.`)];
    out(msgs.length ? alertBox(ix.emps.size ? 'i' : 'b', msgs.map(esc).join('<br>')) : '<div class="card muted">Pick an employee, then choose a day or show all days.</div>');
  }
  const toRows = ws => XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
  $('fileIn').onchange = async e => { // demo/validation path: local XLSX copy, read in-browser, never uploaded
    const f = e.target.files[0]; if (!f) return; out('<div class="card muted">Loading workbook…</div>');
    try {
      const buf = await f.arrayBuffer(), titles = XLSX.read(buf, { type: 'array', bookSheets: true }).SheetNames;
      const map = wanted().map(s => ({ s, t: PRS.resolveSheet(titles, s) }));
      const wb = XLSX.read(buf, { type: 'array', sheets: map.filter(x => x.t).map(x => x.t) }), sheets = {};
      map.forEach(x => { sheets[x.s.sheet] = x.t && wb.Sheets[x.t] ? toRows(wb.Sheets[x.t]) : []; });
      setData(sheets, 'XLSX demo copy', map.filter(x => !x.t && x.s !== cfg.idLookup).map(x => x.s.sheet)); $('btnRefresh').hidden = !G.token;
    } catch (err) { out(alertBox('b', '<b>Could not read the file.</b> ' + esc(err.message))); }
    e.target.value = '';
  };
  // ---------- Google Sheets: live, read-only, in-browser (token stays in memory only) ----------
  const G = { client: null, token: null, expires: 0, pending: null };
  const showErr = (title, msg) => out(alertBox('b', `<b>${title}</b> ${esc(msg)}`));
  const run = fn => Promise.resolve().then(fn).catch(err => showErr('Could not load the Google Sheet.', err.message));

  function googleProblem() { // checks we can do before opening any popup
    if (location.protocol === 'file:') return 'Google sign-in does not work when index.html is opened as a file. Serve the folder (e.g. "python3 -m http.server 8000") and open http://localhost:8000 — see SETUP.md.';
    if (!cfg.google.clientId) return 'Set google.clientId in config.js to your OAuth Web client ID (see SETUP.md).';
    if (!cfg.google.spreadsheetId) return 'Set google.spreadsheetId in config.js.';
    if (!(window.google && google.accounts && google.accounts.oauth2)) return 'The Google sign-in script has not loaded (still loading, offline, or blocked by an ad/privacy blocker). Wait a moment and try again.';
    return null;
  }
  function apiError(status, body) {
    const msg = (body && body.error && body.error.message) || '';
    if (status === 401) return 'Your Google session expired. Click "Refresh data" to sign in again.';
    if (status === 403) return /not been used|disabled|SERVICE_DISABLED/i.test(msg)
      ? 'The Google Sheets API is not enabled in your Cloud project (APIs & Services → Library → Google Sheets API → Enable).'
      : 'Your Google account cannot open this spreadsheet (403). Ask the owner to share it with you as Viewer.';
    if (status === 404) return 'Spreadsheet not found (404). Check google.spreadsheetId in config.js.';
    if (status === 400) return `The sheet rejected the request (400): ${msg}. Check the tab names and ranges in config.js.`;
    if (status === 429) return 'Too many requests to Google (429). Wait a minute and refresh.';
    return `Sheets API error ${status}${msg ? ': ' + msg : ''}`;
  }
  async function api(url) {
    let r;
    try { r = await fetch(url, { headers: { Authorization: 'Bearer ' + G.token } }); }
    catch (e) { throw new Error('Network error reaching sheets.googleapis.com. Check your connection.'); }
    const body = await r.json().catch(() => null);
    if (!r.ok) { if (r.status === 401) G.token = null; throw new Error(apiError(r.status, body)); }
    return body;
  }
  async function loadFromSheets() {
    const base = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(cfg.google.spreadsheetId)}`;
    out('<div class="card muted">Reading Google Sheet…</div>');
    const meta = await api(base + '?fields=sheets.properties.title'), titles = ((meta && meta.sheets) || []).map(s => s.properties.title);
    const map = wanted().map(s => ({ s, t: PRS.resolveSheet(titles, s) })), have = map.filter(x => x.t), sheets = {};
    if (have.length) {
      const q = have.map(x => 'ranges=' + encodeURIComponent(`'${x.t.replace(/'/g, "''")}'!${x.s.range}`)).join('&');
      const body = await api(`${base}/values:batchGet?${q}&valueRenderOption=FORMATTED_VALUE&majorDimension=ROWS`), vr = (body && body.valueRanges) || [];
      have.forEach((x, i) => { sheets[x.s.sheet] = (vr[i] && vr[i].values) || []; });
    }
    setData(sheets, `Google Sheet (live, read-only, loaded ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })})`,
      map.filter(x => !x.t && x.s !== cfg.idLookup).map(x => x.s.sheet));
    $('btnGoogle').hidden = true; $('btnRefresh').hidden = false; $('btnSignOut').hidden = false;
  }
  function withToken(fn) {
    if (G.token && Date.now() < G.expires - 60e3) return run(fn); // still valid: no popup
    if (!G.client) G.client = google.accounts.oauth2.initTokenClient({
      client_id: cfg.google.clientId, scope: cfg.google.scope,
      callback: t => {
        if (t.error) return showErr('Authentication failed:', t.error_description || t.error);
        if (!google.accounts.oauth2.hasGrantedAllScopes(t, cfg.google.scope)) return showErr('Permission not granted.', 'Allow "See your Google Sheets spreadsheets" when signing in.');
        G.token = t.access_token; G.expires = Date.now() + (+t.expires_in || 3600) * 1000;
        const f = G.pending; G.pending = null; if (f) run(f);
      },
      error_callback: e => { G.pending = null; showErr('Sign-in did not complete.',
        e.type === 'popup_closed' ? 'The Google window was closed before finishing.' :
        e.type === 'popup_failed_to_open' ? 'The Google pop-up was blocked. Allow pop-ups for this page and try again.' : String(e.type || e.message || e)); }
    });
    G.pending = fn; out('<div class="card muted">Waiting for Google sign-in…</div>');
    G.client.requestAccessToken({ prompt: G.everSignedIn ? '' : 'consent' }); G.everSignedIn = true;
  }
  const connect = () => { const p = googleProblem(); if (p) return showErr('Google Sheet not available.', p); withToken(loadFromSheets); };
  $('btnGoogle').onclick = connect; $('btnRefresh').onclick = connect;
  $('btnSignOut').onclick = () => {
    const done = () => { G.token = null; G.expires = 0; ix = null; last = null; view = null; $('empList').innerHTML = ''; $('find').disabled = true; $('all').disabled = true; $('hint').textContent = '';
      $('btnGoogle').hidden = false; $('btnRefresh').hidden = true; $('btnSignOut').hidden = true; $('srcList').hidden = true;
      $('srcStatus').textContent = 'Disconnected from Google. No data loaded.'; out('<div class="card muted">Load the Google Sheet or the XLSX demo copy to begin.</div>'); };
    if (G.token && window.google && google.accounts) google.accounts.oauth2.revoke(G.token, done); else done();
  };
  function pickEmp(v) {
    v = v.trim().toLowerCase(); const m = v.match(/\((\d+)\)\s*$/); if (m && ix.emps.has(m[1])) return m[1];
    if (ix.emps.has(v)) return v; const hit = [...ix.emps].filter(([, n]) => n.toLowerCase() === v); return hit.length === 1 ? hit[0][0] : null;
  }
  $('emp').oninput = () => { const id = ix && pickEmp($('emp').value), d = id ? PRS.datesFor(ix, id) : [];
    $('hint').textContent = d.length ? `Data on ${d.length} day${d.length > 1 ? 's' : ''}, ${PRS.fmtD(d[0])} – ${PRS.fmtD(d[d.length - 1])}` : ''; };
  const needEmp = () => { const id = pickEmp($('emp').value); if (!id) out(alertBox('b', '<b>Pick an employee.</b> Choose a name from the list or enter an exact ID.')); return id; };
  $('find').onclick = () => {
    const id = needEmp(); if (!id) return; const d = $('date').value;
    if (!d) return out(alertBox('b', '<b>Pick a date,</b> or use "Show all days".'));
    last = PRS.query(ix, id, d); render(false);
  };
  $('all').onclick = () => { const id = needEmp(); if (!id) return; view = { id }; renderAll(); };

  const vtxt = v => v == null ? '—' : `<span class="${v >= 0 ? 'pos' : 'neg'}">${v >= 0 ? '+' : ''}${v}%</span>`;
  const cell = (r, f, txt) => txt + (r.from[f] && !/^calculated/.test(r.from[f]) ? `<span class="tag">${esc(r.from[f])}</span>` : r.from[f] ? `<span class="tag">calculated</span>` : '');
  const SRC_NOTE = {
    primary: s => `From ${s}.`,
    prs: s => `From ${s}. Start and end times aren't recorded in this sheet.`,
    logs: s => `From the workflow logs (${s}). Target and variance aren't recorded there; productivity per hour is calculated from units and duration.`
  };
  const fmtU = n => (Math.round(n * 100) / 100).toLocaleString();

  function renderAll() {
    const id = view.id, name = ix.emps.get(id) || '', days = PRS.daily(ix, id), t = PRS.totals(days);
    if (!days.length) return out(alertBox('i', `<b>No data found</b> for ${esc(name)} in any source.`));
    const byMonth = new Map(); days.forEach(d => { const k = d.date.slice(0, 7); (byMonth.get(k) || byMonth.set(k, []).get(k)).push(d); });
    const body = t.months.map(m => `<tr class="mh"><th colspan="3">${PRS.fmtMonth(m.month)}</th><th class="n">${m.days} day${m.days > 1 ? 's' : ''}</th><th class="n">${fmtU(m.units)}</th><th class="n">${PRS.fmtH(m.hours)}</th></tr>`
      + byMonth.get(m.month).map(d => `<tr class="day" data-d="${d.date}" tabindex="0"><td><b>${PRS.fmtD(d.date)}</b> <span class="muted">${wd(d.date)}</span></td>
        <td>${d.sources.map(s => `<span class="tag">${esc(s)}</span>`).join('')}</td><td>${esc(d.workflows.join(', '))}${d.conflicts ? '<span class="tag b">conflict</span>' : ''}</td>
        <td class="n">${d.n}</td><td class="n">${fmtU(d.units)}</td><td class="n">${PRS.fmtH(d.hours)}</td></tr>`).join('')).join('');
    out(`<div class="card"><div class="meta"><div><span>Employee</span><b>${esc(name)}</b></div><div><span>Employee ID</span><b>${esc(id)}</b></div>
      <div><span>Period</span><b>${PRS.fmtD(days[0].date)} – ${PRS.fmtD(days[days.length - 1].date)}</b></div></div>
      <p class="muted note">Each day uses its most detailed source: Detailed Productivity first, then the PRS sheets, then the workflow logs. Rows in data-quality conflict are left out of totals. Select a day to see its records.</p></div>
      <div class="cards"><div class="stat"><span>Total units</span><b>${fmtU(t.units)}</b></div><div class="stat"><span>Logged time</span><b>${PRS.fmtH(t.hours)}</b></div>
      <div class="stat"><span>Days with data</span><b>${t.days}</b><small>${t.months.length} month${t.months.length > 1 ? 's' : ''}</small></div><div class="stat"><span>Workflows</span><b>${t.workflows}</b></div></div>
      <div class="card tw"><table><thead><tr><th>Date</th><th>Source</th><th>Workflows</th><th class="n">Records</th><th class="n">Units</th><th class="n">Logged time</th></tr></thead><tbody>${body}</tbody></table></div>`);
    $('out').querySelectorAll('tr.day').forEach(tr => { const go = () => { $('date').value = tr.dataset.d; last = PRS.query(ix, id, tr.dataset.d); render(false, true); };
      tr.onclick = go; tr.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); go(); } }; });
  }
  function render(inc, fromAll) {
    const r = last, back = view && view.id === r.empId ? '<button class="link" id="back" type="button">← All days</button>' : '';
    if (r.mode === 'none') { const d = PRS.datesFor(ix, r.empId);
      out(back + alertBox('i', `<b>No data found</b> for ${esc(r.empName)} on ${PRS.fmtD(r.date)}.` + (d.length ? ` Records exist on ${d.length} day${d.length > 1 ? 's' : ''} from ${PRS.fmtD(d[0])} to ${PRS.fmtD(d[d.length - 1])}. Use "Show all days" to see them.` : '')));
      return wireBack(); }
    const t = PRS.summarize(r.rows, inc), A = [];
    if (r.conflicts.length) A.push(alertBox('b', `<b>Data-quality conflict:</b> ${r.conflicts.map(c => `${esc(c.label)} (rows ${c.rows.join(', ')}; differ in ${c.diff.join(', ')})`).join('; ')}. Same identifiers, different values; ${inc ? 'currently <b>included</b> in totals' : `<b>${t.excluded} row(s) excluded</b> from totals`}. <label><input type="checkbox" id="inc" ${inc ? 'checked' : ''}> include in totals</label>`));
    if (r.softDups.length) A.push(alertBox('i', `Duplicate rows with identical activity but differing ${r.softDups.map(c => c.diff.join('/')).join(', ')} were counted once.`));
    if (t.missingUnits || t.missingHours) A.push(alertBox('', `Some rows have no ${t.missingUnits ? 'units' : ''}${t.missingUnits && t.missingHours ? ' / ' : ''}${t.missingHours ? 'duration' : ''}; totals only include recorded values.`));
    const clock = t.mergedMin != null && Math.abs(t.mergedMin / 60 - t.hours) > 0.02 ? `<small>clock time ${PRS.fmtH(t.mergedMin / 60)} (overlaps merged)</small>` : t.mergedMin != null ? '<small>matches clock time</small>' : '';
    const rows = r.rows.map(x => `<tr class="${x.conflict ? 'cf' : ''}"><td>${esc(x.subtask)}${x.conflict ? '<span class="tag b">conflict</span>' : ''}<div class="muted">${esc(x.task)}</div></td>
      <td>${x.start != null && x.end != null ? PRS.fmtT(x.start) + '–' + PRS.fmtT(x.end) : '—'}</td><td class="n">${x.hours != null ? cell(x, 'hours', PRS.fmtH(x.hours)) : '—'}</td>
      <td class="n">${x.units != null ? cell(x, 'units', x.units.toLocaleString()) : '—'}</td><td class="n">${x.target != null ? cell(x, 'target', x.target) : '—'}</td>
      <td class="n">${x.prod != null ? cell(x, 'prod', x.prod) : '—'}</td><td class="n">${x.variance != null ? cell(x, 'variance', vtxt(x.variance)) : '—'}</td>
      <td>${esc(x.remarks) || '—'}${x.notes.map(n => `<div class="muted">${esc(n)}</div>`).join('')}</td></tr>`).join('');
    const sm = PRS.aiSummary(r, t), note = SRC_NOTE[r.mode](r.sources.map(esc).join(', ')) + (r.repaired ? ' Dates in this sheet are stored with day and month swapped, and were read accordingly.' : '');
    out(`${back}<div class="card"><div class="meta"><div><span>Employee</span><b>${esc(r.empName)}</b></div><div><span>Employee ID</span><b>${esc(r.empId)}</b></div><div><span>Date</span><b>${PRS.fmtD(r.date)}</b></div><div><span>Project</span><b>${esc(r.rows[0].project) || '—'}</b></div></div>
      <p class="muted note">${note}</p></div>
      ${A.join('')}<div class="cards"><div class="stat"><span>Total tasks (units)</span><b>${t.units.toLocaleString()}</b></div><div class="stat"><span>Logged time</span><b>${PRS.fmtH(t.hours)}</b>${clock}</div>
      <div class="stat"><span>Workflows</span><b>${t.workflows}</b></div><div class="stat"><span>Records</span><b>${t.n}</b><small>${t.excluded} excluded</small></div></div>
      ${sm ? `<div class="card sum"><b>Summary</b> <span class="tag">generated from the rows below only</span><p>${esc(sm)}</p></div>` : ''}
      <div class="card tw"><table><thead><tr><th>Queue / workflow</th><th>Time</th><th class="n">Duration</th><th class="n">Units</th><th class="n">Target/h</th><th class="n">Productivity/h</th><th class="n">Variance</th><th>Remarks</th></tr></thead><tbody>${rows}</tbody></table></div>`);
    const c = $('inc'); if (c) c.onchange = () => render(c.checked); wireBack();
  }
  function wireBack() { const b = $('back'); if (b) b.onclick = () => renderAll(); }
  out('<div class="card muted">Load the Google Sheet or the XLSX demo copy to begin.</div>');
})();
