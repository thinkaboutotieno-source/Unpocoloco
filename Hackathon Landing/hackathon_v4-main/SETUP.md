# Connecting PRS Quick View to Google Sheets

The app signs each user in with their own Google account and reads the sheet **read-only**, directly from the browser. There is no server and no client secret. People can only see the sheet if it is already shared with their Google account.

## 1. Create the Google Cloud credentials (one time, ~10 minutes)

1. Go to https://console.cloud.google.com and create a project (or pick an existing one).
2. **Enable the API:** APIs & Services → Library → search "Google Sheets API" → **Enable**.
3. **Consent screen:** APIs & Services → OAuth consent screen (called "Google Auth Platform" in newer consoles).
   - If everyone uses your company Google Workspace accounts, choose **Internal**. No review needed.
   - Otherwise choose **External** and, while in "Testing", add each user's email under **Test users** (up to 100).
   - Under Data access / Scopes, add `https://www.googleapis.com/auth/spreadsheets.readonly`.
4. **Client ID:** APIs & Services → Credentials → Create credentials → **OAuth client ID** → Application type **Web application**.
   - Under **Authorized JavaScript origins** add every address the app will be opened from, exactly (scheme + host + port, no path, no trailing slash):
     - `http://localhost:8000` for local testing
     - your hosted address, e.g. `https://prs.yourcompany.com`
   - Redirect URIs: leave empty (not used).
5. Copy the **Client ID** (ends in `.apps.googleusercontent.com`) into `config.js`:

```js
google: {
  spreadsheetId: '11Psd8-mcHOwv9uVUVd8ZnpWj_qMQToLK9HeaJDShd3E',
  clientId: '1234567890-abc123.apps.googleusercontent.com',
  ...
}
```

Do **not** copy the client secret anywhere. This app does not need it.

## 2. Share the sheet

Each person who will use the app needs at least **Viewer** access to the spreadsheet with the same Google account they sign in with.

## 3. Run the app

Google sign-in refuses to run from `file://`, so serve the folder instead of double-clicking `index.html`:

```bash
cd path/to/prs-quick-view
python3 -m http.server 8000      # or: npx serve -l 8000
```

Open http://localhost:8000, click **Connect Google Sheet**, sign in, approve read access. Use **Refresh data** to pull the latest values (it re-uses the session for about an hour, then asks again) and **Disconnect** to revoke access.

For team use, put the same files on any static host (Google Sites embed, Firebase Hosting, Netlify, an internal web server) and add that origin to step 4.

## 4. Which tabs are read

All tabs are listed in `sources` in `config.js`. For each person and day the app uses the most detailed tab that has data, and never adds tabs together, so the same work isn't counted twice:

1. **Detailed Productivity**: start/end times, targets.
2. **PRS sheets**: `Sept. PRS`, `July PRS`. Add a new month by copying the July line and changing `sheet`.
3. **Workflow logs**: `Alert Validation PRODUCTIVITY & HOURS` (AV2), `AVIP`, `LC-GC Hours`, `Realogram Correction HOURS from 1st August`, `SVm- RC Hours`, `SVm-Gap Hours`. These have no target or variance; productivity per hour is calculated and labelled. Rows whose Employee ID is empty or `#N/A` get it from `Judith Analysis-New` (columns H:I) by email.

A tab that isn't found is skipped with a note, and the rest still load. Click **Sources** next to the status line to see how many rows each tab contributed. **Show all days** lists every day with data for the chosen employee, grouped by month with subtotals; click a day to open it.

## 5. Check date formatting in the sheet (important)

The app reads values **as displayed** in the sheet. Each tab has a `dateRule` in `config.js`: `dmy` for `DD/MM/YYYY`, `mdy` for US-style `8/1/2026`, and `YYYY-MM-DD` is always accepted. The workflow logs are set to `mdy` because they display US dates; the AV2 log shows dates without a year (`7/1`), so it also has `year: 2026`. If a tab's dates show a different way in your sheet, change its `dateRule`. Dates that can't be read are skipped, never guessed.

## Troubleshooting

| Message / symptom | Fix |
|---|---|
| `Error 400: redirect_uri_mismatch` or `origin_mismatch` in the Google pop-up | The page's address isn't in Authorized JavaScript origins. Add it exactly (e.g. `http://localhost:8000`, not `127.0.0.1`). Changes can take a few minutes. |
| `Error 403: access_denied` in the pop-up | App is "External / Testing" and the user isn't a test user. Add them. |
| "Google Sheets API is not enabled" | Step 1.2. |
| "cannot open this spreadsheet (403)" | Share the sheet with that account (step 2). |
| "rejected the request (400)" | A range in `config.js` is invalid. Check `range` for each source. |
| "Tab … was not found" | That tab name in `config.js` doesn't match the spreadsheet. Fix `sheet`, or add the real name to `aliases`. |
| "column … not found" | A header in row 1 changed. Update `columns` in `config.js`. |
| Pop-up blocked / nothing happens | Allow pop-ups for the site; disable ad blockers that block `accounts.google.com`. |
| "Google app isn't verified" warning | Normal for External apps in testing. Use Internal (Workspace) or submit for verification before wide rollout. |
