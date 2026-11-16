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
 *   GET  ?action=fans&since=ms   APPROVED station characters (+ ids removed since)
 *   GET  ?action=mine&id&token   a guest's own character: pending / approved / rejected / removed
 *   POST {action:"fan"}          add a station character (+ optional message) -> {id, token, status}
 *   POST {action:"removeFan"}    remove one with its token (or ADMIN_KEY)
 * Crew review queue (ADMIN_KEY):
 *   GET  ?action=queue&key       characters waiting for approval + the review setting
 *   POST {action:"review"}       {id, approve:true|false}
 *   POST {action:"setReview"}    {on:true|false}  off = characters without a message skip the queue
 *                                (messages always wait for a person). Stored as Script Property REVIEW.
 * Guest profiles (token = the private key the phone got when it made the character):
 *   GET  ?action=mine&id&token   now also returns name, look, tagline, fan number, outfit change, moments
 *   POST {action:"fan", email}   email is stored privately; it is how the guest gets back in later
 *   POST {action:"login"}        {email} -> that email's characters {id, token, name, status} (no emails are sent)
 *   POST {action:"editFan"}      {id, token, look}: outfit change, waits for the crew if already on stage
 *   POST {action:"moment"}       {fanId, token, image (base64 jpeg), thumb (data URL)}: up to 3 photos
 *   POST {action:"removeMoment"} {id, fanId, token} (or ADMIN_KEY)
 *   GET  ?action=moments         approved moments for the stage banner (Drive image links)
 * Crew (ADMIN_KEY): queue also returns outfit changes + moments waiting;
 *   POST {action:"reviewChange"} / {action:"reviewMoment"}  {id, approve}
 *   GET  ?action=momentsAdmin&key   approved moments with thumbnails
 * Photos go to a private Drive folder "AMJ26 Moments" and are only shared (anyone with the link) once approved.
 * Moments need Drive access: run setupPermissions once from the editor, then redeploy.
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
// status: pending | approved | rejected. Empty status (rows from before the review queue) counts as approved.
// shown_at: when it went on stage (approval time), so the stage feed picks up late approvals.
// email: private, only used so the guest can log back in. pending_look/change_status: an outfit change waiting for the crew.
// updated_at: when an approved outfit change went live, so the stage redraws that character.
const FAN_HEAD = ["created_at", "id", "name", "look_json", "token", "removed_at", "message", "status", "shown_at",
  "email", "pending_look", "change_status", "updated_at"];
const F = { created: 0, id: 1, name: 2, look: 3, token: 4, removed: 5, message: 6, status: 7, shown: 8,
  email: 9, pending: 10, change: 11, updated: 12 };
const LOOK_KEYS = ["skin", "hair", "shirt", "pants", "acc", "hairStyle", "outfit", "bottom", "gear"];
const MAX_FANS = 2000;
const MAX_PENDING = 300;
const MSG_MAX = 48;

function fanSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(FAN_SHEET);
  if (!sh) {
    sh = ss.insertSheet(FAN_SHEET);
    sh.appendRow(FAN_HEAD);
    sh.setFrozenRows(1);
  } else if (!CacheService.getScriptCache().get("fanhead2")) {
    // older sheets: add the new column headers once
    const head = sh.getRange(1, 1, 1, FAN_HEAD.length).getValues()[0];
    if (String(head[F.updated]) !== "updated_at") sh.getRange(1, 1, 1, FAN_HEAD.length).setValues([FAN_HEAD]);
    try { CacheService.getScriptCache().put("fanhead2", "1", 21600); } catch (x) {}
  }
  return sh;
}

function fanRows_(sh) {
  const n = sh.getLastRow() - 1;
  return n >= 1 ? sh.getRange(2, 1, n, FAN_HEAD.length).getValues() : [];
}
function statusOf_(r) { return String(r[F.status] || "approved"); }
function reviewOn_() { return PropertiesService.getScriptProperties().getProperty("REVIEW") !== "off"; }

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

function fanOut_(r) {
  let look = {};
  try { look = JSON.parse(String(r[F.look])); } catch (x) {}
  const created = ms_(r[F.created]);
  return { id: String(r[F.id]), name: String(r[F.name]), look: look, message: String(r[F.message] || ""),
    createdAt: created, shownAt: ms_(r[F.shown]) || created, updatedAt: ms_(r[F.updated]) };
}

// Only approved characters are public; pending and rejected ones never leave the sheet.
function publicFans_(since) {
  const now = Date.now(); // taken before reading so rows added mid-read are caught next poll
  const fans = [], removed = [];
  fanRows_(fanSheet_()).forEach(r => {
    if (!r[F.id] || statusOf_(r) !== "approved") return;
    const gone = ms_(r[F.removed]);
    if (gone) { if (gone > since) removed.push(String(r[F.id])); return; }
    const f = fanOut_(r);
    if (f.shownAt > since || f.updatedAt > since) fans.push(f);
  });
  return { ok: true, now: now, fans: fans, removed: removed };
}

function findFan_(sh, id) {
  const rows = fanRows_(sh);
  for (let i = 0; i < rows.length; i++) if (String(rows[i][F.id]) === String(id)) return { row: i + 2, r: rows[i] };
  return null;
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

/* Run this once from the editor (pick "setupPermissions" in the toolbar, then Run) to allow Drive access for moments. */
function setupPermissions() {
  momFolder_();
  Logger.log("Drive access OK. Moments folder: " + momFolder_().getUrl());
}

/* ===== Guest profiles: email login, outfit changes, moments ===== */
const MOM_SHEET = "Moments";
const MOM_HEAD = ["created_at", "id", "fan_id", "file_id", "thumb", "status", "removed_at", "shown_at"];
const M = { created: 0, id: 1, fan: 2, file: 3, thumb: 4, status: 5, removed: 6, shown: 7 };
const MOM_PER_FAN = 3, MAX_MOMENTS = 1500;

function cleanEmail_(s) {
  const e = String(s == null ? "" : s).trim().toLowerCase().slice(0, 80);
  return /^[^\s@<>"',;]+@[^\s@<>"',;]+\.[a-z]{2,}$/.test(e) ? e : "";
}
function maskEmail_(e) { e = String(e || ""); const i = e.indexOf("@"); return i > 0 ? e[0] + "•••" + e.slice(i) : ""; }
function tokenOk_(r, token) { return typeof token === "string" && token.length >= 16 && token === String(r[F.token]); }


function momSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MOM_SHEET);
  if (!sh) { sh = ss.insertSheet(MOM_SHEET); sh.appendRow(MOM_HEAD); sh.setFrozenRows(1); }
  return sh;
}
function momRows_(sh) {
  const n = sh.getLastRow() - 1;
  return n >= 1 ? sh.getRange(2, 1, n, MOM_HEAD.length).getValues() : [];
}
function momFolder_() {
  const props = PropertiesService.getScriptProperties(), id = props.getProperty("MOMENTS_FOLDER");
  if (id) { try { return DriveApp.getFolderById(id); } catch (x) {} }
  const f = DriveApp.createFolder("AMJ26 Moments");
  props.setProperty("MOMENTS_FOLDER", f.getId());
  return f;
}
function momUrl_(fileId) { return "https://drive.google.com/thumbnail?id=" + encodeURIComponent(fileId) + "&sz=w1600"; }
function trash_(fileId) { try { DriveApp.getFileById(String(fileId)).setTrashed(true); } catch (x) {} }

// approved, not removed, and the character itself is on stage
function publicMoments_() {
  const cache = CacheService.getScriptCache(), hit = cache.get("moments");
  if (hit) return hit;
  const live = {};
  fanRows_(fanSheet_()).forEach(r => { if (r[F.id] && statusOf_(r) === "approved" && !ms_(r[F.removed])) live[String(r[F.id])] = true; });
  const list = momRows_(momSheet_()).filter(r => r[M.id] && r[M.status] === "approved" && !ms_(r[M.removed]) && live[String(r[M.fan])])
    .map(r => ({ id: String(r[M.id]), fanId: String(r[M.fan]), url: momUrl_(String(r[M.file])), shownAt: ms_(r[M.shown]) }));
  const out = JSON.stringify({ ok: true, moments: list });
  try { cache.put("moments", out, 10); } catch (x) {}
  return out;
}
function momOut_(r, names) {
  return { id: String(r[M.id]), fanId: String(r[M.fan]), name: names ? (names[String(r[M.fan])] || "") : undefined,
    thumb: String(r[M.thumb] || ""), status: String(r[M.status]), createdAt: ms_(r[M.created]) };
}
function fanNames_() { const o = {}; fanRows_(fanSheet_()).forEach(r => { if (r[F.id]) o[String(r[F.id])] = String(r[F.name]); }); return o; }

// a guest's own profile
function mine_(id, token) {
  const rows = fanRows_(fanSheet_());
  let r = null;
  for (let i = 0; i < rows.length; i++) if (String(rows[i][F.id]) === String(id)) { r = rows[i]; break; }
  if (!r || !tokenOk_(r, token)) return { ok: false, error: "not_found" };
  const status = ms_(r[F.removed]) ? "removed" : statusOf_(r);
  let look = {}, pendingLook = null;
  try { look = JSON.parse(String(r[F.look])); } catch (x) {}
  if (r[F.pending]) { try { pendingLook = JSON.parse(String(r[F.pending])); } catch (x) {} }
  let fanNo = 0;
  if (status === "approved") {
    // same order as the stage: on-stage characters by creation time, then id
    const c = ms_(r[F.created]), me = String(r[F.id]);
    rows.forEach(x => {
      if (!x[F.id] || statusOf_(x) !== "approved" || ms_(x[F.removed])) return;
      const xc = ms_(x[F.created]);
      if (xc < c || (xc === c && String(x[F.id]) <= me)) fanNo++;
    });
  }
  const moments = momRows_(momSheet_()).filter(m => String(m[M.fan]) === me_(r) && !ms_(m[M.removed])).map(m => momOut_(m));
  return { ok: true, status: status, name: String(r[F.name]), message: String(r[F.message] || ""), look: look,
    fanNo: fanNo, email: maskEmail_(r[F.email]), change: String(r[F.change] || ""), pendingLook: pendingLook,
    updatedAt: ms_(r[F.updated]), moments: moments };
}
function me_(r) { return String(r[F.id]); }

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
    const message = clean_(b.message, MSG_MAX);
    const email = cleanEmail_(b.email);
    if (!name || !look || (b.email && !email)) return json_({ ok: false, error: "invalid" });
    const id = Utilities.getUuid().replace(/-/g, "").slice(0, 16);
    const token = Utilities.getUuid().replace(/-/g, "");
    const now = Date.now();
    const status = reviewOn_() || message ? "pending" : "approved";
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = fanSheet_();
      if (sh.getLastRow() - 1 >= MAX_FANS) return json_({ ok: false, error: "quota_exceeded" });
      if (status === "pending") {
        const n = sh.getLastRow() - 1;
        const waiting = n >= 1 ? sh.getRange(2, F.status + 1, n, 1).getValues().filter(v => v[0] === "pending").length : 0;
        if (waiting >= MAX_PENDING) return json_({ ok: false, error: "busy" });
      }
      sh.appendRow([now, text_(id), text_(name), text_(JSON.stringify(look)), text_(token), "",
        message ? text_(message) : "", status, status === "approved" ? now : "", email ? text_(email) : "", "", "", ""]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true, id: id, token: token, createdAt: now, status: status });
  }

  if (b.action === "login") {
    const email = cleanEmail_(b.email);
    if (!email) return json_({ ok: false, error: "invalid" });
    const list = fanRows_(fanSheet_())
      .filter(r => r[F.id] && cleanEmail_(r[F.email]) === email && !ms_(r[F.removed]) && statusOf_(r) !== "rejected")
      .map(r => ({ id: String(r[F.id]), token: String(r[F.token]), name: String(r[F.name]), status: statusOf_(r), createdAt: ms_(r[F.created]) }))
      .sort((a, b) => b.createdAt - a.createdAt);
    if (!list.length) return json_({ ok: false, error: "no_match" });
    return json_({ ok: true, fans: list });
  }

  if (b.action === "editFan") {
    const look = cleanLook_(b.look);
    if (!look) return json_({ ok: false, error: "invalid" });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = fanSheet_(), hit = findFan_(sh, b.id);
      if (!hit || !tokenOk_(hit.r, b.token)) return json_({ ok: false, error: "not_found" });
      const st = ms_(hit.r[F.removed]) ? "removed" : statusOf_(hit.r);
      if (st === "pending") {            // not on stage yet: just swap the look in the queue
        sh.getRange(hit.row, F.look + 1).setValue(text_(JSON.stringify(look)));
        return json_({ ok: true, mode: "direct" });
      }
      if (st !== "approved") return json_({ ok: false, error: "not_allowed" });
      sh.getRange(hit.row, F.pending + 1, 1, 2).setValues([[text_(JSON.stringify(look)), "pending"]]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true, mode: "pending" });
  }

  if (b.action === "reviewChange") {
    if (!authed_(b.key)) return json_({ ok: false, error: "unauthorized" });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = fanSheet_(), hit = findFan_(sh, b.id);
      if (!hit || String(hit.r[F.change]) !== "pending") return json_({ ok: false, error: "not_found" });
      if (b.approve === true) {
        sh.getRange(hit.row, F.look + 1).setValue(text_(String(hit.r[F.pending])));
        sh.getRange(hit.row, F.pending + 1, 1, 3).setValues([["", "", Date.now()]]);
      } else {
        sh.getRange(hit.row, F.pending + 1, 1, 2).setValues([["", "rejected"]]);
      }
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  }

  if (b.action === "moment") {
    const thumb = String(b.thumb || ""), img = String(b.image || "");
    if (thumb.indexOf("data:image/jpeg;base64,") !== 0 || thumb.length > 48000 || img.length < 2000 || img.length > 4000000)
      return json_({ ok: false, error: "invalid" });
    let bytes;
    try { bytes = Utilities.base64Decode(img); } catch (x) { return json_({ ok: false, error: "invalid" }); }
    const fan = findFan_(fanSheet_(), b.fanId);
    if (!fan || !tokenOk_(fan.r, b.token)) return json_({ ok: false, error: "not_found" });
    const st = ms_(fan.r[F.removed]) ? "removed" : statusOf_(fan.r);
    if (st === "removed" || st === "rejected") return json_({ ok: false, error: "not_allowed" });
    const id = Utilities.getUuid().replace(/-/g, "").slice(0, 16);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = momSheet_(), rows = momRows_(sh);
      if (rows.length >= MAX_MOMENTS) return json_({ ok: false, error: "quota_exceeded" });
      if (rows.filter(r => String(r[M.fan]) === String(b.fanId) && !ms_(r[M.removed])).length >= MOM_PER_FAN)
        return json_({ ok: false, error: "full" });
      const file = momFolder_().createFile(Utilities.newBlob(bytes, "image/jpeg", "moment-" + id + ".jpg"));
      sh.appendRow([Date.now(), text_(id), text_(String(b.fanId)), text_(file.getId()), text_(thumb), "pending", "", ""]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true, id: id });
  }

  if (b.action === "removeMoment" || b.action === "reviewMoment") {
    const admin = authed_(b.key);
    if (b.action === "reviewMoment" && !admin) return json_({ ok: false, error: "unauthorized" });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = momSheet_(), rows = momRows_(sh);
      let i = -1;
      for (let k = 0; k < rows.length; k++) if (String(rows[k][M.id]) === String(b.id)) { i = k; break; }
      if (i < 0 || ms_(rows[i][M.removed])) return json_({ ok: false, error: "not_found" });
      const r = rows[i], row = i + 2;
      if (!admin) {
        const fan = findFan_(fanSheet_(), r[M.fan]);
        if (!fan || !tokenOk_(fan.r, b.token)) return json_({ ok: false, error: "unauthorized" });
      }
      if (b.action === "removeMoment") {
        sh.getRange(row, M.removed + 1).setValue(Date.now());
        trash_(r[M.file]);
      } else if (b.approve === true) {
        if (r[M.status] !== "pending") return json_({ ok: false, error: "not_found" });
        DriveApp.getFileById(String(r[M.file])).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        sh.getRange(row, M.status + 1).setValue("approved");
        sh.getRange(row, M.shown + 1).setValue(Date.now());
      } else {
        if (r[M.status] !== "pending") return json_({ ok: false, error: "not_found" });
        sh.getRange(row, M.status + 1).setValue("rejected");
        trash_(r[M.file]);
      }
    } finally {
      lock.releaseLock();
    }
    try { CacheService.getScriptCache().remove("moments"); } catch (x) {}
    return json_({ ok: true });
  }

  if (b.action === "review") {
    if (!authed_(b.key)) return json_({ ok: false, error: "unauthorized" });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = fanSheet_(), hit = findFan_(sh, b.id);
      if (!hit || statusOf_(hit.r) !== "pending") return json_({ ok: false, error: "not_found" });
      const ok = b.approve === true;
      sh.getRange(hit.row, F.status + 1, 1, 2).setValues([[ok ? "approved" : "rejected", ok ? Date.now() : ""]]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  }

  if (b.action === "setReview") {
    if (!authed_(b.key)) return json_({ ok: false, error: "unauthorized" });
    PropertiesService.getScriptProperties().setProperty("REVIEW", b.on === false ? "off" : "on");
    return json_({ ok: true, review: b.on !== false });
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
  if (action === "mine") return json_(mine_(e.parameter.id, String(e.parameter.token || "")));
  if (action === "moments") return ContentService.createTextOutput(publicMoments_()).setMimeType(ContentService.MimeType.JSON);
  if (action === "queue") {
    if (!authed_(e.parameter.key)) return json_({ ok: false, error: "unauthorized" });
    const rows = fanRows_(fanSheet_()), names = {};
    rows.forEach(r => { if (r[F.id]) names[String(r[F.id])] = String(r[F.name]); });
    const pending = rows.filter(r => r[F.id] && statusOf_(r) === "pending" && !ms_(r[F.removed])).map(fanOut_);
    const changes = rows.filter(r => r[F.id] && String(r[F.change]) === "pending" && !ms_(r[F.removed])).map(r => {
      const f = fanOut_(r); let pl = {};
      try { pl = JSON.parse(String(r[F.pending])); } catch (x) {}
      f.pendingLook = pl; return f;
    });
    const moments = momRows_(momSheet_()).filter(m => m[M.id] && m[M.status] === "pending" && !ms_(m[M.removed])).map(m => momOut_(m, names));
    return json_({ ok: true, review: reviewOn_(), pending: pending, changes: changes, moments: moments });
  }
  if (action === "momentsAdmin") {
    if (!authed_(e.parameter.key)) return json_({ ok: false, error: "unauthorized" });
    const names = fanNames_();
    const list = momRows_(momSheet_()).filter(m => m[M.id] && m[M.status] === "approved" && !ms_(m[M.removed]))
      .map(m => momOut_(m, names)).sort((a, b) => b.createdAt - a.createdAt).slice(0, 80);
    return json_({ ok: true, moments: list });
  }
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
