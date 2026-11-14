/**
 * Backend for the /f2026-taur-poll/ page. Saves votes to the "Votes" tab of the Google
 * Sheet this script is attached to.
 *
 * This folder starts with "_", so Jekyll does not publish it. It is kept here
 * as the source of truth for what is pasted into Apps Script.
 *
 * Setup
 *   1. Create a Google Sheet. Open Extensions -> Apps Script.
 *   2. Replace the contents of Code.gs with this file and save.
 *   3. Deploy -> New deployment -> type "Web app".
 *        Execute as: Me
 *        Who has access: Anyone
 *      Authorize when prompted. ("Google hasn't verified this app" is expected
 *      for your own script: Advanced -> Go to ... (unsafe).)
 *   4. Copy the Web app URL (ends in /exec) into `poll_endpoint` in
 *      _pages/f2026-taur-poll.html.
 *
 * After editing this script, publish the change with
 *   Deploy -> Manage deployments -> pencil icon -> Version: New version -> Deploy
 * which keeps the same URL. "New deployment" would create a new URL.
 *
 * Sheet layout (one row per person; resubmitting with the same name,
 * ignoring case, overwrites that row):
 *   Name | Ranked? | Updated | <one column per option>
 * An option cell holds the person's rank (1 = favorite) if they ranked,
 * a check mark if they voted without ranking, and is blank otherwise.
 * COUNTA() over an option column gives its vote count.
 */

const SHEET_NAME = "Votes";
const BASE_HEADERS = ["Name", "Ranked?", "Updated"];
const MAX_NAME_LENGTH = 60;
const MAX_OPTION_LENGTH = 100;
const MAX_OPTIONS = 50;

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return json_({ ok: false, error: "The poll is busy. Please try again in a moment." });
  }

  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    const name = tidy_(data.name, MAX_NAME_LENGTH);
    if (!name) return json_({ ok: false, error: "Please enter your name." });

    const votes = list_(data.votes).filter((v) => BASE_HEADERS.indexOf(v) === -1);
    if (!votes.length) return json_({ ok: false, error: "Please pick at least one option." });

    const ranking = list_(data.ranking).filter((v) => votes.indexOf(v) !== -1);

    const sheet = getSheet_();
    const headers = ensureHeaders_(sheet, votes);

    // Build a full-width row so an overwrite also clears options no longer picked.
    const row = headers.map(() => "");
    row[0] = safe_(name);
    row[1] = ranking.length ? "Yes" : "No";
    row[2] = new Date();
    votes.forEach((v) => {
      const rank = ranking.indexOf(v);
      row[headers.indexOf(v)] = rank === -1 ? "✓" : rank + 1;
    });

    const existing = findRow_(sheet, name);
    const target = existing || sheet.getLastRow() + 1;
    sheet.getRange(target, 1, 1, row.length).setValues([row]);
    // Commit the write before replying, so "saved" on the page means it's in the sheet.
    SpreadsheetApp.flush();

    return json_({ ok: true, replaced: Boolean(existing) });
  } catch (err) {
    return json_({ ok: false, error: "Server error: " + err.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * ?action=lookup&name=X returns that person's saved vote so the page can
 * restore it on any device. Visiting the bare URL confirms the deployment is live.
 */
function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.action !== "lookup") return json_({ ok: true, message: "Poll endpoint is running." });

  try {
    const name = tidy_(params.name, MAX_NAME_LENGTH);
    if (!name) return json_({ ok: false, error: "Missing name." });

    const sheet = getSheet_();
    const row = findRow_(sheet, name);
    if (!row) return json_({ ok: true, found: false });

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

    const votes = [];
    const ranked = [];
    for (let c = BASE_HEADERS.length; c < lastCol; c++) {
      const cell = values[c];
      if (cell === "" || cell == null) continue;
      votes.push(headers[c]);
      if (typeof cell === "number") ranked.push([cell, headers[c]]);
    }
    ranked.sort((a, b) => a[0] - b[0]);

    return json_({
      ok: true,
      found: true,
      votes: votes,
      ranking: ranked.map((r) => r[1]),
      updated: values[2] instanceof Date ? values[2].toISOString() : null,
    });
  } catch (err) {
    return json_({ ok: false, error: "Server error: " + err.message });
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

/** Make sure row 1 has the base headers plus a column for every option voted on. */
function ensureHeaders_(sheet, options) {
  const lastCol = sheet.getLastColumn();
  let headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0] : [];
  if (!headers.length || headers[0] === "") headers = BASE_HEADERS.slice();

  const missing = options.filter((o) => headers.indexOf(o) === -1);
  if (missing.length || lastCol === 0) {
    headers = headers.concat(missing);
    sheet
      .getRange(1, 1, 1, headers.length)
      .setValues([headers.map(safe_)])
      .setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return headers;
}

/** Row number of an existing vote with this name (case-insensitive), or 0. */
function findRow_(sheet, name) {
  const count = sheet.getLastRow() - 1;
  if (count < 1) return 0;
  const key = name.toLowerCase();
  const names = sheet.getRange(2, 1, count, 1).getDisplayValues();
  for (let i = 0; i < count; i++) {
    if (tidy_(names[i][0], MAX_NAME_LENGTH).toLowerCase() === key) return i + 2;
  }
  return 0;
}

function tidy_(value, maxLength) {
  return (value == null ? "" : String(value)).trim().replace(/\s+/g, " ").slice(0, maxLength);
}

function list_(value) {
  if (!Array.isArray(value)) return [];
  const out = [];
  value.forEach((v) => {
    const s = tidy_(v, MAX_OPTION_LENGTH);
    if (s && out.indexOf(s) === -1) out.push(s);
  });
  return out.slice(0, MAX_OPTIONS);
}

/** Stop text that looks like a formula from being evaluated by Sheets. */
function safe_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
