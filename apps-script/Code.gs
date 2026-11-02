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
 *
 * Live @ Arts & Music Jam (public, no key, never returns contacts):
 *   GET  ?action=pieces          featured pieces for the live stage
 *   GET  ?action=fans&since=ms   station characters (+ ids removed since)
 *   POST {action:"fan"}          add a station character -> {id, token}
 *   POST {action:"removeFan"}    remove one with its token (or ADMIN_KEY)
 * After editing, redeploy: Deploy -> Manage deployments -> Edit -> New version
 * (keeps the same /exec URL).
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

const FAN_SHEET = "Fans";
const FAN_HEAD = ["created_at", "id", "name", "look_json", "token", "removed_at"];
const LOOK_KEYS = ["skin", "hair", "shirt", "pants", "acc", "hairStyle", "outfit", "bottom", "gear"];
const MAX_FANS = 2000;

function fanSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(FAN_SHEET);
  if (!sh) {
    sh = ss.insertSheet(FAN_SHEET);
    sh.appendRow(FAN_HEAD);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Keep only the expected keys, as small ints or short id strings. The stage re-checks everything.
function cleanLook_(l) {
  if (!l || typeof l !== "object") return null;
  const out = {};
  for (let i = 0; i < LOOK_KEYS.length; i++) {
    const k = LOOK_KEYS[i], v = l[k];
    if (typeof v === "number" && v >= 0 && v < 64 && Math.floor(v) === v) out[k] = v;
    else if (typeof v === "string" && /^[a-z0-9:]{1,30}$/i.test(v)) out[k] = v;
  }
  return out;
}

function ms_(v) { return v instanceof Date ? v.getTime() : (Number(v) || Date.parse(v) || 0); }

function publicPieces_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get("pieces");
  if (hit) return hit;
  const sh = sheet_(), n = sh.getLastRow() - 1;
  const pieces = [];
  if (n >= 1) {
    sh.getRange(2, 1, n, HEAD.length).getValues().forEach(r => {
      const on = r[8] === true || String(r[8]).toUpperCase() === "TRUE";
      if (!on || !r[1]) return;
      pieces.push({ id: String(r[1]), title: String(r[5]), designer: String(r[2]), category: String(r[4]),
        pixels: String(r[6]), createdAt: ms_(r[0]) });
    });
  }
  const out = JSON.stringify({ ok: true, pieces: pieces });
  try { cache.put("pieces", out, 10); } catch (x) {}
  return out;
}

function publicFans_(since) {
  const now = Date.now(); // taken before reading so rows added mid-read are caught next poll
  const sh = fanSheet_(), n = sh.getLastRow() - 1;
  const fans = [], removed = [];
  if (n >= 1) {
    sh.getRange(2, 1, n, FAN_HEAD.length).getValues().forEach(r => {
      if (!r[1]) return;
      const id = String(r[1]), created = ms_(r[0]), gone = ms_(r[5]);
      if (gone) { if (gone > since) removed.push(id); return; }
      if (created <= since) return;
      let look = {};
      try { look = JSON.parse(String(r[3])); } catch (x) {}
      fans.push({ id: id, name: String(r[2]), look: look, createdAt: created });
    });
  }
  return { ok: true, now: now, fans: fans, removed: removed };
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

  if (b.action === "fan") {
    const name = String(b.name || "").toUpperCase().replace(/[^A-Z0-9 ]/g, "").replace(/\s+/g, " ").trim().slice(0, 12);
    const look = cleanLook_(b.look);
    if (!name || !look) return json_({ ok: false, error: "invalid" });
    const id = Utilities.getUuid().replace(/-/g, "").slice(0, 16);
    const token = Utilities.getUuid().replace(/-/g, "");
    const now = Date.now();
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = fanSheet_();
      if (sh.getLastRow() - 1 >= MAX_FANS) return json_({ ok: false, error: "quota_exceeded" });
      sh.appendRow([now, text_(id), text_(name), text_(JSON.stringify(look)), text_(token), ""]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true, id: id, token: token, createdAt: now });
  }

  if (b.action === "removeFan") {
    const sh = fanSheet_(), n = sh.getLastRow() - 1;
    if (n < 1) return json_({ ok: false, error: "not_found" });
    const rows = sh.getRange(2, 2, n, 4).getValues(); // id, name, look, token
    for (let i = 0; i < n; i++) {
      if (String(rows[i][0]) !== String(b.id)) continue;
      const tokenOk = typeof b.token === "string" && b.token.length >= 16 && b.token === String(rows[i][3]);
      if (!tokenOk && !authed_(b.key)) return json_({ ok: false, error: "unauthorized" });
      sh.getRange(i + 2, 6).setValue(Date.now());
      return json_({ ok: true });
    }
    return json_({ ok: false, error: "not_found" });
  }

  if (b.action === "feature") {
    if (!authed_(b.key)) return json_({ ok: false, error: "unauthorized" });
    const sh = sheet_(), n = sh.getLastRow() - 1;
    if (n < 1) return json_({ ok: false, error: "not_found" });
    const ids = sh.getRange(2, 2, n, 1).getValues();
    for (let i = 0; i < n; i++) {
      if (String(ids[i][0]) === String(b.id)) {
        sh.getRange(i + 2, 9).setValue(b.featured === true);
        try { CacheService.getScriptCache().remove("pieces"); } catch (x) {}
        return json_({ ok: true });
      }
    }
    return json_({ ok: false, error: "not_found" });
  }

  return json_({ ok: false, error: "unknown_action" });
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === "pieces") return ContentService.createTextOutput(publicPieces_()).setMimeType(ContentService.MimeType.JSON);
  if (action === "fans") return json_(publicFans_(Number(e.parameter.since) || 0));
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
