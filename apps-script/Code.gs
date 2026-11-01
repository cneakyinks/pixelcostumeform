/**
 * AMJ26 pixel form — Google Sheets backend.
 *
 * Paste into the sheet's Extensions → Apps Script editor, then:
 *   1. Project Settings → Script Properties → add ADMIN_KEY = <your private key, 12+ characters>.
 *   2. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 *   3. Put the web app URL into SHEET_URL in index.html.
 *
 * Anyone can submit a piece (POST action "submit").
 * Reading entries (GET) and featuring (POST action "feature") need ADMIN_KEY,
 * so contact details are never readable from the public page.
 * Never commit the real ADMIN_KEY to the repo.
 */

const SHEET_NAME = "Entries";
const HEAD = ["submitted_at", "id", "name", "contact", "type", "title", "pixels", "consent", "featured"];
const CATS = ["hair", "top", "bottom", "accessory"];
const CELLS = 12 * 21;

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange("I2:I").insertCheckboxes();
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function authed_(k) {
  const real = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY") || "";
  return real.length >= 12 && typeof k === "string" && k === real;
}

function clean_(s, n) {
  return String(s == null ? "" : s).replace(/[\u0000-\u001f<>]/g, "").replace(/\s+/g, " ").trim().slice(0, n);
}

// A leading apostrophe makes Sheets store the value as plain text:
// stops formula injection (=, +, -, @) and stops ids/pixels being read as numbers.
function text_(s) { return "'" + s; }

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: "bad_request" }); }

  if (b.action === "submit") {
    const name = clean_(b.name, 16), title = clean_(b.title, 24), contact = clean_(b.contact, 60);
    const type = String(b.category), px = String(b.pixels || "");
    const id = String(b.id || "").replace(/[^a-z0-9]/g, "").slice(0, 24) || Utilities.getUuid();
    const filled = (px.match(/[0-9a-f]/g) || []).length;
    if (!name || !title || CATS.indexOf(type) < 0 || b.consent !== true ||
        px.length !== CELLS || !/^[0-9a-f.]+$/.test(px) || filled < 3) {
      return json_({ ok: false, error: "invalid" });
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = sheet_();
      sh.appendRow([new Date(), text_(id), text_(name), text_(contact), type, text_(title), text_(px), true, false]);
      sh.getRange(sh.getLastRow(), 9).insertCheckboxes();
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true, id: id });
  }

  if (b.action === "feature") {
    if (!authed_(b.key)) return json_({ ok: false, error: "unauthorized" });
    const sh = sheet_(), n = sh.getLastRow() - 1;
    if (n < 1) return json_({ ok: false, error: "not_found" });
    const ids = sh.getRange(2, 2, n, 1).getValues();
    for (let i = 0; i < n; i++) {
      if (String(ids[i][0]) === String(b.id)) {
        sh.getRange(i + 2, 9).setValue(b.featured === true);
        return json_({ ok: true });
      }
    }
    return json_({ ok: false, error: "not_found" });
  }

  return json_({ ok: false, error: "unknown_action" });
}

function doGet(e) {
  if (!authed_(e.parameter.key)) return json_({ ok: false, error: "unauthorized" });
  const sh = sheet_(), n = sh.getLastRow() - 1;
  if (n < 1) return json_({ ok: true, entries: [] });
  const rows = sh.getRange(2, 1, n, HEAD.length).getValues();
  const entries = rows.filter(r => r[1]).map(r => ({
    createdAt: r[0] instanceof Date ? r[0].getTime() : (Date.parse(r[0]) || 0),
    id: String(r[1]),
    name: String(r[2]),
    contact: String(r[3]),
    category: String(r[4]),
    title: String(r[5]),
    pixels: String(r[6]),
    featured: r[8] === true || String(r[8]).toUpperCase() === "TRUE",
  }));
  return json_({ ok: true, entries: entries });
}
