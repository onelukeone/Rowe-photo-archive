// ===== Discussion: a message board for each photo =====
// Add this as a new Script file (for example named "Discussion").
// It uses the settings and helpers already in your UploadServer file.
// Messages are stored in a private "Discussion" tab. The website reads only the APPROVED ones,
// through your web app, so the tab (and the email addresses in it) never needs to be published.

const DISCUSS = {
  SHEET: "Discussion",
  AUTO_APPROVE: true,    // true = new messages appear on the site immediately, without your review
  MAX_TEXT: 2000
};
const DISCUSS_HEADERS = ["POST ID", "TIMESTAMP", "PHOTO ID", "FILE NAME", "NAME", "EMAIL", "MESSAGE", "STATUS", "REPLY TO"];

// Called by the website: returns every approved message
function discussionList() {
  const sh = SpreadsheetApp.openById(UPLOAD.SHEET_ID).getSheetByName(DISCUSS.SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const d = sh.getDataRange().getValues();
  const h = d[0].map(x => String(x).trim().toUpperCase());
  const c = n => h.indexOf(n);
  if ([c("POST ID"), c("TIMESTAMP"), c("PHOTO ID"), c("NAME"), c("MESSAGE"), c("STATUS")].some(i => i < 0)) return [];
  const idOf = v => (String(v).match(/[-\w]{25,}/) || [""])[0];
  return d.slice(1)
    .filter(r => String(r[c("STATUS")]).trim().toUpperCase() === "APPROVED" && String(r[c("MESSAGE")]).trim() && idOf(r[c("PHOTO ID")]))
    .map(r => ({
      id: String(r[c("POST ID")]),
      photo: idOf(r[c("PHOTO ID")]),
      name: String(r[c("NAME")] || "").trim() || "Anonymous",
      text: String(r[c("MESSAGE")]).trim(),
      replyTo: c("REPLY TO") >= 0 ? String(r[c("REPLY TO")] || "").trim() : "",
      when: r[c("TIMESTAMP")] instanceof Date ? r[c("TIMESTAMP")].toISOString() : String(r[c("TIMESTAMP")])
    }))
    .sort((a, b) => a.when < b.when ? -1 : a.when > b.when ? 1 : 0);
}

// Called by the website when a visitor posts a message
function discussionPost(password, d) {
  upCheck_(password);
  d = d || {};
  const t = (v, n) => String(v || "").trim().slice(0, n);
  const name = t(d.name, 100), email = t(d.email, 200), text = t(d.text, DISCUSS.MAX_TEXT);
  if (!name) throw new Error("Please enter your name.");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) throw new Error("Please enter a valid email address.");
  if (!text) throw new Error("Please write a message first.");
  const id = (String(d.photoId || "").match(/[-\w]{25,}/) || [""])[0];
  if (!id) throw new Error("That photo couldn't be identified.");

  // Gentle brake against floods of messages
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get("discposts") || 0);
  if (n >= 40) throw new Error("A lot of messages are being sent right now. Please try again in a few minutes.");
  cache.put("discposts", String(n + 1), 600);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(UPLOAD.SHEET_ID);
    const data = ss.getSheetByName(UPLOAD.MAIN_SHEET).getDataRange().getValues();
    const head = data[0].map(x => String(x).trim().toUpperCase());
    const cLink = head.indexOf("PHOTO LINK"), cFile = head.indexOf("FILE NAME");
    const row = data.slice(1).find(r => (String(r[cLink]).match(/[-\w]{25,}/) || [""])[0] === id);
    if (!row) throw new Error("That photo is no longer in the archive.");
    const fileName = cFile >= 0 ? String(row[cFile] || "") : "";

    let sh = ss.getSheetByName(DISCUSS.SHEET);
    if (!sh) {
      sh = ss.insertSheet(DISCUSS.SHEET);
      sh.appendRow(DISCUSS_HEADERS);
      sh.setFrozenRows(1);
    }
    // Replies: the parent must be an approved message on this same photo. A reply to a reply is filed under the original message.
    let replyTo = "";
    const wanted = String(d.replyTo || "").trim().slice(0, 40);
    if (wanted) {
      const all = sh.getDataRange().getValues();
      const hh = all[0].map(x => String(x).trim().toUpperCase());
      const iId = hh.indexOf("POST ID"), iPh = hh.indexOf("PHOTO ID"), iSt = hh.indexOf("STATUS"), iRe = hh.indexOf("REPLY TO");
      const parent = all.slice(1).find(r => String(r[iId]) === wanted);
      if (!parent || String(parent[iSt]).trim().toUpperCase() !== "APPROVED" ||
          (String(parent[iPh]).match(/[-\w]{25,}/) || [""])[0] !== id) throw new Error("The message you're replying to isn't available.");
      replyTo = (iRe >= 0 && String(parent[iRe] || "").trim()) || wanted;
    }
    // Older Discussion tabs have no REPLY TO column yet: add its heading
    if (!String(sh.getRange(1, 9).getValue()).trim()) sh.getRange(1, 9).setValue("REPLY TO");
    const status = DISCUSS.AUTO_APPROVE ? "APPROVED" : "";
    const postId = Utilities.getUuid().slice(0, 8);
    const now = new Date();
    // Store the text columns as plain text, so a message starting with = + - or @ is never read as a formula
    const next = sh.getLastRow() + 1;
    sh.getRange(next, 4, 1, 4).setNumberFormat("@");
    sh.getRange(next, 9).setNumberFormat("@");
    sh.getRange(next, 1, 1, 9).setValues([[postId, now, id, fileName, name, email, text, status, replyTo]]);

    if (!DISCUSS.AUTO_APPROVE) {
      try {
        MailApp.sendEmail(UPLOAD.NOTIFY_EMAIL || Session.getEffectiveUser().getEmail(),
          (replyTo ? "New discussion reply from " : "New discussion message from ") + name,
          name + " (" + email + ") wrote on \"" + (fileName || id) + "\"" + (replyTo ? " (a reply to message " + replyTo + ")" : "") + ":\n\n" + text + "\n\n" +
          "To publish it, type APPROVED in the STATUS cell of the Discussion tab:\n" +
          "https://docs.google.com/spreadsheets/d/" + UPLOAD.SHEET_ID + "/edit");
      } catch (err) {}
    }
    return { approved: !!status, post: status ? { id: postId, photo: id, name: name, text: text, when: now.toISOString(), replyTo: replyTo } : null };
  } finally {
    lock.releaseLock();
  }
}
