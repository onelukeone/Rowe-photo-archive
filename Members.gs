// ===== Members: personal access codes, a shareable test code, and the photo list =====
// Add this as a NEW Script file named "Members". It uses the UPLOAD settings in your UploadServer file.
//
// How it works
//  - Each person gets their own access code. Codes live in a private "Members" tab (NAME, EMAIL, CODE, STATUS, SENT, LAST USED).
//  - To add people: type their NAME and EMAIL in the Members tab, then run sendMemberCodes. Each person is emailed a code and a sign-in link.
//  - People can also ask for access on the sign-in screen. Their request lands in the Members tab with STATUS PENDING and you get an email.
//    Type APPROVED in their STATUS cell, and the next sendMemberCodes run emails them a code. (Nothing is sent to them before you approve.)
//    To have that happen by itself, add a time trigger for sendMemberCodes: Triggers (clock icon), Add trigger, Time-driven, every 15 minutes.
//    With that trigger on, anyone you type in with an email and a blank STATUS is emailed within 15 minutes. Type HOLD in STATUS to hold someone back.
//  - To cancel someone: type REVOKED in their STATUS cell. It takes effect within a couple of minutes.
//  - To make someone an editor (can run the fingerprint tool): type ADMIN in their STATUS cell.
//  - TEST_CODE below is a shared sign-in you can hand to testers while the site is being worked on. Switch it off with TEST_ENABLED.
//  - The website no longer reads a published spreadsheet. It asks this script for the photo list, and the script only answers
//    people with a valid code. Emails and anything else not listed in MEM_COLUMNS never leave the spreadsheet.

const MEMBERS = {
  SHEET: "Members",                            // private tab. Never publish it.
  SITE_URL: "PASTE_YOUR_SITE_ADDRESS_HERE",    // your website address, like https://yourname.github.io/Rowe-photo-archive/
  SITE_NAME: "Rowe Camp Photo Archive",
  TEST_ENABLED: true,                          // the shared test sign-in. Set to false to switch it off.
  TEST_CODE: "ROWE-TEST-K7Q2",                 // change it whenever you like, and share it with testers
  TEST_EXPIRES: "",                            // optional last day for the test code, like "2026-12-31". Blank = no end date.
  ALLOW_OLD_PASSWORD: false,                   // true = the old shared password (UPLOAD.PASSWORD) also works, as an editor sign-in
  LIST_PER_HOUR: 40,                           // how often one personal code may load the archive per hour (a normal visit is 1 to 3)
  TEST_LIST_PER_HOUR: 200,                     // the same limit for the shared test code
  MAX_WRONG_TRIES: 30,                         // wrong codes allowed per ten minutes (across everyone) before the sign-in pauses
  REQUESTS_PER_HOUR: 10                        // access requests accepted per hour (across everyone), to keep spam out of your inbox
};

// The only Sheet1 columns the website is ever sent (heading names, capitals don't matter)
const MEM_COLUMNS = {
  archive: ["TITLE", "NAME", "DATE", "ERA", "PHOTOGRAPHER", "PEOPLE", "PROGRAM", "TAGS", "DESCRIPTION", "COMMENTS", "NOTES", "SOURCE", "FILE NAME", "PHOTO LINK", "STATUS"],
  upload: ["TITLE", "NAME", "TAGS", "FILE NAME", "PHOTO LINK", "STATUS", "HASH"]
};
const MEM_HEADERS = ["NAME", "EMAIL", "CODE", "STATUS", "SENT", "LAST USED", "REQUESTED", "NOTE"];
// STATUS values that count as an active member. Anything else (PENDING, REVOKED, HOLD, ...) can't sign in and isn't emailed.
function memStatusOk_(status) { return !status || status === "ACTIVE" || status === "ADMIN" || status === "APPROVED"; }
const MEM_ALPHABET = "ABCDEFGHJKMNPQRSTUVWXYZ23456789";   // no 0, O, 1, I or L, so codes are easy to read and type

// ---------- the check every request goes through ----------

function memErr_(message, kind) { const e = new Error(message); e.kind = kind; return e; }
function memNorm_(s) { return String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, ""); }

function memTestExpired_() {
  const m = String(MEMBERS.TEST_EXPIRES || "").match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  return !!m && new Date() > new Date(+m[1], +m[2] - 1, +m[3], 23, 59, 59);
}

// Returns { name, key, admin, test, row }. Throws if the code isn't valid.
function memCheck_(code) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get("fails") || 0);
  if (fails >= MEMBERS.MAX_WRONG_TRIES) throw memErr_("Too many wrong codes. Please try again in ten minutes.", "busy");
  const n = memNorm_(code);
  let m = null;
  if (n) {
    if (MEMBERS.TEST_ENABLED && memNorm_(MEMBERS.TEST_CODE) === n && !memTestExpired_()) m = { name: "Test visitor", key: "test", test: true, admin: false };
    else if (MEMBERS.ALLOW_OLD_PASSWORD && UPLOAD.PASSWORD && String(code) === UPLOAD.PASSWORD) m = { name: "Editor", key: "legacy", admin: true };
    else {
      const t = memTable_().byCode[n];
      if (t) m = { name: t.name, key: n, admin: !!t.admin, row: t.row };
    }
  }
  if (!m) {
    cache.put("fails", String(fails + 1), 600);
    throw memErr_("That access code isn't right.", "auth");
  }
  memSeen_(m);
  return m;
}

// For the editor-only tools (the fingerprint tool)
function upAdmin_(password) {
  const m = upCheck_(password);
  if (!m.admin) throw memErr_("That tool is only for the archive's editors.", "auth");
  return m;
}

// The Members tab as a lookup table, kept for two minutes so every request doesn't have to read the sheet
function memTable_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get("mem_tbl");
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const t = { byCode: {}, lastUsedCol: 0 };
  const sh = SpreadsheetApp.openById(UPLOAD.SHEET_ID).getSheetByName(MEMBERS.SHEET);
  if (sh && sh.getLastRow() >= 2) {
    const d = sh.getDataRange().getValues();
    const h = d[0].map(x => String(x).trim().toUpperCase());
    const cName = h.indexOf("NAME"), cCode = h.indexOf("CODE"), cStat = h.indexOf("STATUS"), cUsed = h.indexOf("LAST USED");
    if (cUsed >= 0) t.lastUsedCol = cUsed + 1;
    if (cCode >= 0) d.slice(1).forEach((r, i) => {
      const n = memNorm_(r[cCode]);
      const status = cStat >= 0 ? String(r[cStat]).trim().toUpperCase() : "";
      if (!n || !memStatusOk_(status)) return;
      t.byCode[n] = { name: cName >= 0 ? String(r[cName] || "").trim() : "", admin: status === "ADMIN", row: i + 2 };
    });
  }
  try { cache.put("mem_tbl", JSON.stringify(t), 120); } catch (e) {}
  return t;
}

// Notes "LAST USED" in the Members tab, at most once every six hours per person
function memSeen_(m) {
  if (!m.row) return;
  try {
    const cache = CacheService.getScriptCache(), k = "seen_" + m.key;
    if (cache.get(k)) return;
    cache.put(k, "1", 21600);
    const col = memTable_().lastUsedCol;
    if (col) SpreadsheetApp.openById(UPLOAD.SHEET_ID).getSheetByName(MEMBERS.SHEET).getRange(m.row, col).setValue(new Date());
  } catch (e) {}
}

// ---------- the photo list ----------

// Called by the website after sign-in. view = "archive" (the main page) or "upload" (the upload page).
function archiveList(password, view) {
  const m = upCheck_(password);
  view = view === "upload" ? "upload" : "archive";
  memThrottle_(m, "list", m.test ? MEMBERS.TEST_LIST_PER_HOUR : MEMBERS.LIST_PER_HOUR);
  const cache = CacheService.getScriptCache();
  const hit = memCacheGet_(cache, "list_" + view);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const data = memBuildList_(view);
  memCachePut_(cache, "list_" + view, JSON.stringify(data), 300);
  return data;
}

function memBuildList_(view) {
  const sh = SpreadsheetApp.openById(UPLOAD.SHEET_ID).getSheetByName(UPLOAD.MAIN_SHEET);
  const last = sh.getLastRow(), cols = sh.getLastColumn();
  if (last < 2) return { head: [], rows: [] };
  const vals = sh.getRange(1, 1, last, cols).getDisplayValues();   // text exactly as it shows in the sheet
  const head = vals[0].map(h => String(h).trim()), up = head.map(h => h.toUpperCase());
  const keep = [];
  MEM_COLUMNS[view].forEach(name => { const i = up.indexOf(name); if (i >= 0) keep.push(i); });
  const cLink = up.indexOf("PHOTO LINK"), cStat = up.indexOf("STATUS");
  const rows = [];
  vals.slice(1).forEach(r => {
    if (!/[-\w]{25,}/.test(String(r[cLink] || ""))) return;
    if (view === "archive" && cStat >= 0 && String(r[cStat]).trim().toLowerCase() !== "approved") return;   // unapproved photos are never sent
    rows.push(keep.map(i => r[i]));
  });
  return { head: keep.map(i => head[i]), rows: rows };
}

// A person (or a script pretending to be one) can only load the whole archive so many times an hour
function memThrottle_(m, what, limit) {
  const cache = CacheService.getScriptCache();
  const k = what + "_" + m.key + "_" + Utilities.formatDate(new Date(), "UTC", "yyyyMMddHH");
  const n = Number(cache.get(k) || 0) + 1;
  cache.put(k, String(n), 3600);
  if (n > limit) {
    if (n === limit + 1) {
      try {
        MailApp.sendEmail(UPLOAD.NOTIFY_EMAIL || Session.getEffectiveUser().getEmail(), "Unusual activity on the photo archive",
          (m.test ? "The shared test code" : "The access code for " + (m.name || "a member")) + " loaded the whole archive more than " + limit +
          " times in one hour, so it has been paused until the hour is up.\n\nIf that isn't expected, type REVOKED in that person's STATUS cell in the Members tab.");
      } catch (e) {}
    }
    throw memErr_("This code has loaded the archive a lot in the last hour. Please wait a while and try again.", "busy");
  }
}

// The cache holds 100 KB per entry, so the list is compressed and split into pieces
function memCachePut_(cache, key, text, ttl) {
  try {
    const gz = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(text, "application/octet-stream", "d")).getBytes());
    const SIZE = 90000, parts = Math.ceil(gz.length / SIZE), obj = {};
    for (let i = 0; i < parts; i++) obj[key + "_" + i] = gz.substr(i * SIZE, SIZE);
    obj[key + "_n"] = String(parts);
    cache.putAll(obj, ttl);
  } catch (e) {}
}
function memCacheGet_(cache, key) {
  try {
    const n = Number(cache.get(key + "_n") || 0);
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(key + "_" + i);
    const got = cache.getAll(keys);
    let s = "";
    for (let i = 0; i < n; i++) { const p = got[key + "_" + i]; if (p === null || p === undefined) return null; s += p; }
    return Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(s), "application/x-gzip")).getDataAsString();
  } catch (e) { return null; }
}

// ---------- a check you can run when something doesn't work ----------

// Run this, then open View > Logs (or the Execution log at the bottom) to see what is and isn't in place.
function checkMembersSetup() {
  const out = [];
  const say = (ok, text) => out.push((ok ? "OK:       " : "PROBLEM:  ") + text);
  let src = "";
  try { src = String(upCheck_); } catch (e) {}
  say(/memCheck_/.test(src), /memCheck_/.test(src)
    ? "upCheck_ (in UploadServer) uses the new access codes."
    : "upCheck_ in UploadServer is still the OLD one. Replace the whole function with the new one, which just says: return memCheck_(password);");
  let post = "";
  try { post = String(doPost); } catch (e) {}
  const routes = ["list", "requestAccess", "checkAdmin"].filter(a => post.indexOf('"' + a + '"') < 0);
  say(!routes.length, !routes.length ? "doPost has the new routes." : "doPost is missing the routes: " + routes.join(", ") + ". Replace the whole doPost function with the newest one.");
  try {
    const m = memCheck_(MEMBERS.TEST_CODE);
    say(true, "The test code (" + MEMBERS.TEST_CODE + ") is accepted.");
  } catch (e) {
    if (!MEMBERS.TEST_ENABLED) say(true, "The test code is switched off (TEST_ENABLED is false).");
    else say(false, "The test code is not accepted: " + e.message + (memTestExpired_() ? " Its end date (TEST_EXPIRES) has passed." : ""));
  }
  const sh = SpreadsheetApp.openById(UPLOAD.SHEET_ID).getSheetByName(MEMBERS.SHEET);
  if (!sh) say(false, "There is no Members tab yet. Run setupMembers.");
  else {
    const t = memTable_();
    say(true, "The Members tab has " + Object.keys(t.byCode).length + " active code" + (Object.keys(t.byCode).length === 1 ? "" : "s") + ".");
  }
  say(!/PASTE_/.test(MEMBERS.SITE_URL), /PASTE_/.test(MEMBERS.SITE_URL) ? "SITE_URL at the top of the Members file hasn't been set yet (needed before sendMemberCodes)." : "SITE_URL is set to " + MEMBERS.SITE_URL);
  out.push("");
  out.push("One thing this can't see: after any change, you must also choose Deploy > Manage deployments > pencil > Version: New version > Deploy. Saving alone isn't enough.");
  const text = out.join("\n");
  Logger.log(text);
  return text;
}

// ---------- asking for access ----------

// Called by the sign-in screen (no code needed). Adds a PENDING row to the Members tab and tells you by email.
// The reply is the same whether or not the address is already known, so nobody can use the form to find out who is a member.
function requestAccess(r) {
  r = r || {};
  const t = (v, n) => String(v || "").trim().slice(0, n);
  if (t(r.website, 20)) return { received: true };   // a hidden box that only automated form-fillers complete
  const name = t(r.name, 100).replace(/\s+/g, " "), email = t(r.email, 200), note = t(r.note, 500);
  if (!name) throw new Error("Please enter your name.");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) throw new Error("Please enter a valid email address.");

  const cache = CacheService.getScriptCache();
  const k = "req_" + Utilities.formatDate(new Date(), "UTC", "yyyyMMddHH");
  const n = Number(cache.get(k) || 0);
  if (n >= MEMBERS.REQUESTS_PER_HOUR) throw memErr_("A lot of requests are coming in right now. Please try again in an hour.", "busy");
  cache.put(k, String(n + 1), 3600);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = memSheet_();
    const d = sh.getDataRange().getValues();
    const h = d[0].map(x => String(x).trim().toUpperCase());
    const cMail = h.indexOf("EMAIL");
    if (d.slice(1).some(row => String(row[cMail] || "").trim().toLowerCase() === email.toLowerCase())) return { received: true };   // already on the list
    const row = new Array(h.length).fill("");
    const put = (col, v) => { const i = h.indexOf(col); if (i >= 0) row[i] = v; };
    put("NAME", upSafe_(name)); put("EMAIL", upSafe_(email)); put("STATUS", "PENDING"); put("REQUESTED", new Date()); put("NOTE", upSafe_(note));
    sh.appendRow(row);
    try {
      MailApp.sendEmail(UPLOAD.NOTIFY_EMAIL || Session.getEffectiveUser().getEmail(), "Access request from " + name,
        name + " (" + email + ") asked for access to the " + MEMBERS.SITE_NAME + "." + (note ? "\n\nThey wrote: " + note : "") + "\n\n" +
        "To approve, type APPROVED in their STATUS cell in the Members tab, then run sendMemberCodes (or wait for the automatic run, if you set one up):\n" +
        "https://docs.google.com/spreadsheets/d/" + UPLOAD.SHEET_ID + "/edit");
    } catch (err) {}
    return { received: true };
  } finally {
    lock.releaseLock();
  }
}

// ---------- creating and emailing codes ----------

// Creates the Members tab if it isn't there yet. Run it once, or let sendMemberCodes do it.
function setupMembers() {
  memSheet_();
  Logger.log("The Members tab is ready. Type each person's NAME and EMAIL, then run sendMemberCodes.");
}

function memSheet_() {
  const ss = SpreadsheetApp.openById(UPLOAD.SHEET_ID);
  let sh = ss.getSheetByName(MEMBERS.SHEET);
  if (!sh) {
    sh = ss.insertSheet(MEMBERS.SHEET);
    sh.appendRow(MEM_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange("C:C").setNumberFormat("@");
  }
  const h = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(x => String(x).trim().toUpperCase());
  MEM_HEADERS.forEach(name => { if (h.indexOf(name) < 0) { h.push(name); sh.getRange(1, h.length).setValue(name); } });
  return sh;
}

// A random code like K7QM-4TXD-9HPA (uses Google's secure random numbers)
function memNewCode_(taken) {
  for (let tries = 0; tries < 50; tries++) {
    const hex = Utilities.getUuid().replace(/-/g, "");
    let s = "";
    for (let i = 0; i < 12; i++) s += MEM_ALPHABET[parseInt(hex.substr(i * 2, 2), 16) % MEM_ALPHABET.length];
    if (!taken[s] && s !== memNorm_(MEMBERS.TEST_CODE)) return s;
  }
  throw new Error("Couldn't make a unique code. Please try again.");
}
function memShow_(n) { return n.slice(0, 4) + "-" + n.slice(4, 8) + "-" + n.slice(8); }

// Gives a code to everyone in the Members tab who has an EMAIL and hasn't been sent one yet, and emails it.
// To send someone their code again, clear the SENT cell in their row and run this again.
function sendMemberCodes() {
  if (/PASTE_/.test(MEMBERS.SITE_URL)) throw new Error("Please set SITE_URL at the top of the Members file to your website address first.");
  const site = String(MEMBERS.SITE_URL).trim().replace(/#.*$/, "");
  const siteBase = /\.html?$/i.test(site) || /\/$/.test(site) ? site : site + "/";

  const sh = memSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = sh.getDataRange().getValues();
    const h = d[0].map(x => String(x).trim().toUpperCase());
    const c = n => h.indexOf(n);
    const cName = c("NAME"), cMail = c("EMAIL"), cCode = c("CODE"), cStat = c("STATUS"), cSent = c("SENT");
    const taken = {};
    d.slice(1).forEach(r => { const n = memNorm_(r[cCode]); if (n) taken[n] = true; });

    let quota = MailApp.getRemainingDailyQuota(), sent = 0, left = 0, bad = 0;
    for (let i = 1; i < d.length; i++) {
      const r = d[i], email = String(r[cMail] || "").trim(), status = String(r[cStat] || "").trim().toUpperCase();
      if (!email || String(r[cSent] || "").trim()) continue;
      if (!memStatusOk_(status)) continue;
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) { sh.getRange(i + 1, cSent + 1).setValue("not a valid email"); bad++; continue; }
      if (quota <= 0) { left++; continue; }

      let code = memNorm_(r[cCode]);
      if (!code) { code = memNewCode_(taken); taken[code] = true; }
      const shown = memShow_(code);
      sh.getRange(i + 1, cCode + 1).setNumberFormat("@").setValue(shown);
      const first = String(r[cName] || "").trim().split(/\s+/)[0];
      try {
        MailApp.sendEmail(email, "Your access to the " + MEMBERS.SITE_NAME,
          "Hello" + (first ? " " + first : "") + ",\n\n" +
          "You now have access to the " + MEMBERS.SITE_NAME + ".\n\n" +
          "Your personal access code:  " + shown + "\n\n" +
          "To open the archive and sign in on this device, use this link:\n" + siteBase + "#k=" + shown + "\n\n" +
          "Or go to " + siteBase + " and type the code in when asked.\n\n" +
          "Please keep the code to yourself. It is yours alone, and can be switched off if it is shared.\n" +
          "If anything doesn't work, just reply to this email.\n",
          { name: MEMBERS.SITE_NAME });
        sh.getRange(i + 1, cSent + 1).setValue(new Date());
        sent++; quota--;
      } catch (err) {
        sh.getRange(i + 1, cSent + 1).setValue("failed: " + String((err && err.message) || err).slice(0, 80));
        bad++;
      }
    }
    CacheService.getScriptCache().remove("mem_tbl");
    const msg = "Sent " + sent + " code" + (sent === 1 ? "" : "s") + "." +
      (left ? " " + left + " more are waiting because today's email limit was reached. Run sendMemberCodes again tomorrow." : "") +
      (bad ? " " + bad + " row" + (bad === 1 ? "" : "s") + " had a problem: see the SENT column." : "");
    Logger.log(msg);
    return msg;
  } finally {
    lock.releaseLock();
  }
}
