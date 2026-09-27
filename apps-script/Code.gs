/**
 * Faculty Computer & Tablet Order — Google Sheets backend.
 *
 * Paste this whole file into Extensions → Apps Script of the order spreadsheet,
 * run setup() once, then deploy as a Web app (see SETUP.md).
 *
 * Each person (identified by email) has ONE row in total: a new submission
 * replaces their earlier row, and choosing a computer removes their tablet row
 * (and the other way round).
 */

const CONFIG = {
  // Only emails ending in these domains may submit. Leave empty [] to allow any email.
  ALLOWED_DOMAINS: ["mahidol.ac.th", "mahidol.edu"],
  COMPUTER_ALLOWANCE: 50000,
  TABLET_ALLOWANCE: 40000,
  // Optional: an address that gets an email for every submission. "" = off.
  NOTIFY_EMAIL: "",
  // Optional: stop accepting submissions after this date (YYYY-MM-DD). "" = no deadline.
  CLOSE_DATE: ""
};

/* Prices are the VAT-inclusive amounts on the vendor quotations. Keep in sync with the website. */
const COMPUTERS = {
  "mba13-m5-8g-16-512":     { platform: "Mac", name: "MacBook Air 13″ (M5 10-core CPU / 8-core GPU, 16GB, 512GB, Thai keyboard)", base: 43900,    applecare: 6848, vendor: "COM7", quote: "01QTS/26091123" },
  "mba13-m5-10g-24-1tb":    { platform: "Mac", name: "MacBook Air 13″ (M5 10-core CPU / 10-core GPU, 24GB, 1TB, Thai keyboard)", base: 60700,   applecare: 6848, vendor: "COM7", quote: "01QTS/26091119" },
  "mba13-m5-10g-16-2tb-us": { platform: "Mac", name: "MacBook Air 13″ (M5 10-core CPU / 10-core GPU, 16GB, 2TB, US keyboard)", base: 69496.50, applecare: 6848, vendor: "COM7", quote: "01QTS/26091396" },
  "mba15-m5-10g-32-1tb":    { platform: "Mac", name: "MacBook Air 15″ (M5 10-core CPU / 10-core GPU, 32GB, 1TB, Thai keyboard)", base: 77000,   applecare: 7811, vendor: "COM7", quote: "01QTS/26091115" },
  "asus-expertbook-p5":     { platform: "PC",  name: "ASUS ExpertBook P5 P5405CSA (Core Ultra 7 258V, 32GB, 1TB)", base: 49969, vendor: "Inforgen Data System", quote: "260617-09" },
  "lenovo-legion5":         { platform: "PC",  name: "Lenovo Legion 5 15IAX11 (Core Ultra 9 290HX Plus, RTX 5070, 16GB, 512GB)", base: 79715, vendor: "COM7", quote: "01QTS/26090061" }
};
const MAC_COLOURS = ["Sky Blue", "Silver", "Starlight", "Midnight"];

/* Tablets: estimates, not quotations. iPad: COM7 iPad price list (education price, VAT incl.).
   Samsung: samsung.com/th regular price (VAT incl.), checked 27 Sep 2026. */
const TABLETS = {
  ipad: { name: "iPad", chip: "A16", sizes: [11], storage: [128, 256, 512],
    price: { wifi: { 11: [16300, 20300, 27300] }, cell: { 11: [22300, 26300, 33300] } },
    colours: ["Silver", "Blue", "Pink", "Yellow"] },
  air: { name: "iPad Air", chip: "M4", sizes: [11, 13], storage: [128, 256, 512, 1024],
    price: { wifi: { 11: [26700, 30700, 37700, 48700], 13: [33500, 37500, 44500, 55500] },
             cell: { 11: [32700, 36700, 43700, 54700], 13: [39500, 43500, 50500, 61500] } },
    colours: ["Space Grey", "Starlight", "Purple", "Blue"] },
  pro: { name: "iPad Pro", chip: "M5", sizes: [11, 13], storage: [256, 512, 1024, 2048],
    price: { wifi: { 11: [40900, 47900, 62900, 81900], 13: [52900, 59900, 74900, 93900] },
             cell: { 11: [47900, 54900, 69900, 88900], 13: [59900, 66900, 81900, 100900] } },
    nano: { wifi: { 11: { 1024: 66900, 2048: 85900 }, 13: { 1024: 78900, 2048: 97900 } },
            cell: { 11: { 1024: 73900, 2048: 92900 }, 13: { 1024: 85900, 2048: 104900 } } },
    colours: ["Silver", "Space Black"] },
  s11: { name: "Samsung Galaxy Tab S11", chip: "", cellLabel: "Wi-Fi + 5G", sizes: [11], storage: [128, 256],
    price: { wifi: { 11: [28900, 32900] }, cell: { 11: [33900, 37900] } },
    colours: ["Gray", "Silver"] }
};

const SHEETS = {
  computer: { name: "Computers", headers: ["Updated", "First submitted", "Email", "Name", "Platform", "Computer", "Colour", "AppleCare+", "Price (THB)", "Top-up (THB)", "Budget source", "Grant code", "Vendor", "Quote", "Notes", "Item ID"] },
  tablet:   { name: "Tablets",   headers: ["Updated", "First submitted", "Email", "Name", "Tablet", "Size", "Connectivity", "Storage", "Nano-texture", "Colour", "Est. price (THB)", "Est. top-up (THB)", "Budget source", "Grant code", "Accessories", "Status", "Notes", "Model ID"] }
};
const FIXED = ["Updated", "First submitted", "Email", "Name"]; // first four headers of both sheets

/**
 * Run once from the Apps Script editor to create the two sheets with headers.
 * Safe to run again: it only adds headers that are missing and never moves your data.
 * Values are written by header name, so you may reorder columns, delete ones you
 * don't need, or add your own (e.g. "PO number"); your own columns are kept on updates.
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(k => {
    const def = SHEETS[k];
    const sh = ss.getSheetByName(def.name) || ss.insertSheet(def.name);
    const have = headersOf(sh);
    if (!have.some(Boolean)) {
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
    } else {
      const missing = def.headers.filter(h => have.indexOf(h) < 0);
      if (missing.length) sh.getRange(1, have.length + 1, 1, missing.length).setValues([missing]);
    }
    const n = headersOf(sh).length;
    sh.getRange(1, 1, 1, n).setFontWeight("bold").setBackground("#e3edfb");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, n);
  });
}

function doGet() {
  return json({ ok: true, message: "Order endpoint is running." });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (CONFIG.CLOSE_DATE && new Date() > new Date(CONFIG.CLOSE_DATE + "T23:59:59")) throw new Error("Ordering is closed.");

    const email = String(data.email || "").trim().toLowerCase();
    const name = clean(data.fullName, 120);
    if (!name) throw new Error("Please fill in your full name.");
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error("Please enter a valid email.");
    if (CONFIG.ALLOWED_DOMAINS.length && !CONFIG.ALLOWED_DOMAINS.some(d => email.endsWith("@" + d))) {
      throw new Error("Please use your university email (" + CONFIG.ALLOWED_DOMAINS.map(d => "@" + d).join(" or ") + ").");
    }

    const kind = data.kind === "tablet" ? "tablet" : "computer";
    const rec = kind === "computer" ? computerRecord(data) : tabletRecord(data);
    const now = new Date();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const target = ss.getSheetByName(SHEETS[kind].name);
    const other = ss.getSheetByName(SHEETS[kind === "computer" ? "tablet" : "computer"].name);
    if (!target || !other) throw new Error("The order sheet is not set up yet. Run setup() in Apps Script.");

    // Keep the first submission time, then remove this person's row from the other sheet.
    let first = now;
    const inOther = findRow(other, email);
    if (inOther) { first = cellByHeader(other, inOther, "First submitted") || now; other.deleteRow(inOther); }
    const existing = findRow(target, email);
    if (existing) first = cellByHeader(target, existing, "First submitted") || first;

    const fields = { "Updated": now, "First submitted": first, "Email": email, "Name": name };
    SHEETS[kind].headers.slice(FIXED.length).forEach((h, i) => { fields[h] = rec.cells[i]; });
    writeRow(target, existing, fields);

    if (CONFIG.NOTIFY_EMAIL) {
      MailApp.sendEmail(CONFIG.NOTIFY_EMAIL, "Order: " + name + " — " + rec.summary,
        name + " (" + email + ") chose: " + rec.summary + "\n" + rec.budgetLine + "\n\nSee the spreadsheet: " + ss.getUrl());
    }
    return json(Object.assign({ ok: true, kind: kind, updatedAt: now.toISOString(), submittedAt: new Date(first).toISOString() }, rec.reply));
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function computerRecord(d) {
  const it = COMPUTERS[d.itemId];
  if (!it) throw new Error("Unknown computer. Reload the page and choose again.");
  const applecare = !!(it.applecare && d.applecare);
  const colour = it.platform === "Mac" ? (MAC_COLOURS.indexOf(d.color) >= 0 ? d.color : "Silver") : "";
  const total = round2(it.base + (applecare ? it.applecare : 0));
  const topUp = Math.max(0, round2(total - CONFIG.COMPUTER_ALLOWANCE));
  const source = topUp > 0 ? clean(d.budgetSource, 500) : "";
  if (topUp > 0 && !source) throw new Error("Please fill in the budget source for the top-up.");
  const grant = topUp > 0 ? clean(d.grantCode, 100) : "";
  return {
    cells: [it.platform, it.name, colour, it.platform === "Mac" ? (applecare ? "Yes" : "No") : "", total, topUp, source, grant, it.vendor, it.quote, clean(d.notes, 1000), d.itemId],
    summary: it.name + (applecare ? " + AppleCare+" : ""),
    budgetLine: topUp > 0 ? "Top-up ฿" + topUp + " from: " + source : "Within budget",
    reply: { total: total, topUp: topUp }
  };
}

function tabletRecord(d) {
  const fam = TABLETS[d.family];
  if (!fam) throw new Error("Unknown tablet. Reload the page and choose again.");
  const size = Number(d.size), storage = Number(d.storageGB);
  const conn = d.connectivity === "cell" ? "cell" : "wifi";
  if (fam.sizes.indexOf(size) < 0 || fam.storage.indexOf(storage) < 0) throw new Error("That tablet configuration is not available.");
  const nano = !!(d.nanoTexture && fam.nano && storage >= 1024);
  const price = nano ? fam.nano[conn][size][storage] : fam.price[conn][size][fam.storage.indexOf(storage)];
  const topUp = Math.max(0, price - CONFIG.TABLET_ALLOWANCE);
  const colour = fam.colours.indexOf(d.color) >= 0 ? d.color : fam.colours[0];
  const source = topUp > 0 ? clean(d.budgetSource, 500) : "";
  if (topUp > 0 && !source) throw new Error("Please fill in the budget source in case of a top-up.");
  const grant = topUp > 0 ? clean(d.grantCode, 100) : "";
  const gb = storage >= 1024 ? (storage / 1024) + "TB" : storage + "GB";
  const title = fam.chip ? fam.name + " (" + fam.chip + ")" : fam.name;
  const connName = conn === "cell" ? (fam.cellLabel || "Wi-Fi + Cellular") : "Wi-Fi";
  return {
    cells: [title, size + "-inch", connName, gb, nano ? "Yes" : "No", colour, price, topUp, source, grant, clean(d.accessories, 300), "Awaiting quotation", clean(d.notes, 1000), d.family],
    summary: title + " " + size + "″ " + connName + " " + gb + (nano ? " nano-texture" : ""),
    budgetLine: topUp > 0 ? "May need top-up (catalog estimate) from: " + source : "Likely within budget (catalog estimate)",
    reply: { estTopUp: topUp }
  };
}

function headersOf(sheet) {
  const cols = sheet.getLastColumn();
  return cols ? sheet.getRange(1, 1, 1, cols).getValues()[0].map(h => String(h).trim()) : [];
}
function colOf(sheet, header) {
  const i = headersOf(sheet).indexOf(header);
  if (i < 0) throw new Error("The sheet \"" + sheet.getName() + "\" has no \"" + header + "\" column. Run setup() in Apps Script.");
  return i + 1;
}
function cellByHeader(sheet, row, header) {
  const i = headersOf(sheet).indexOf(header);
  return i < 0 ? "" : sheet.getRange(row, i + 1).getValue();
}
// Write values into the columns whose header matches; other columns keep what they had.
function writeRow(sheet, row, fields) {
  const headers = headersOf(sheet);
  if (headers.indexOf("Email") < 0) throw new Error("The order sheet is not set up yet. Run setup() in Apps Script.");
  const current = row ? sheet.getRange(row, 1, 1, headers.length).getValues()[0] : headers.map(() => "");
  const values = headers.map((h, i) => Object.prototype.hasOwnProperty.call(fields, h) ? safeCell(fields[h]) : current[i]);
  if (row) sheet.getRange(row, 1, 1, headers.length).setValues([values]);
  else sheet.appendRow(values);
}
function findRow(sheet, email) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const emails = sheet.getRange(2, colOf(sheet, "Email"), last - 1, 1).getValues();
  for (let i = 0; i < emails.length; i++) {
    if (String(emails[i][0]).trim().toLowerCase() === email) return i + 2;
  }
  return 0;
}
function clean(v, max) { return String(v == null ? "" : v).replace(/\s+/g, " ").trim().slice(0, max || 500); }
// Stop typed text from being read as a spreadsheet formula.
function safeCell(v) { return typeof v === "string" && /^[=+\-@]/.test(v) ? "'" + v : v; }
function round2(n) { return Math.round(n * 100) / 100; }
function json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
