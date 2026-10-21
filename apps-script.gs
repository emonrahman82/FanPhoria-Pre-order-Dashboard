// FanPhoria backend: paste this whole file into Google Apps Script (Extensions > Apps Script)
// 1) Your Google Sheet (text data). Replace with YOUR sheet link:
const SHEET_URL  = "https://docs.google.com/spreadsheets/d/1xjgvfZgRUrVpil9rXIoCcXGYic3arGFlV-mVDvYZ6Og/edit";
const TOKEN     = "change-this-secret";   // must match APPS_SCRIPT_TOKEN in index.html

const HEAD = ["Order ID","Date","Status","Customer ID","Name","Phone","Address","Customer notes","Product","Anime","Category","Price","Image link","Total","Paid","Due","Order notes"];

function sheet_(name, header) {
  const ss = SpreadsheetApp.openByUrl(SHEET_URL);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    if (name === "Orders") sh.getRange("F:F").setNumberFormat("@"); // keep phone numbers as text
  }
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function removeRows_(id) {
  ["Orders", "Data"].forEach(n => {
    const sh = sheet_(n, n === "Orders" ? HEAD : ["Order ID", "JSON"]);
    const last = sh.getLastRow();
    if (last < 2) return;
    const ids = sh.getRange(1, 1, last, 1).getValues().flat();
    for (let i = ids.length; i >= 2; i--) if (ids[i - 1] === id) sh.deleteRow(i);
  });
}

function doGet(e) {
  const token = e && e.parameter ? e.parameter.token : "";   // "e" is empty when run from the editor
  if (token !== TOKEN) return out_({ error: "unauthorized" });
  const sh = sheet_("Data", ["Order ID", "JSON"]);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues() : [];
  return out_({ ok: true, items: rows.filter(r => r[1]).map(r => JSON.parse(r[1])) });
}

function doPost(e) {
  if (!e || !e.postData) return out_({ error: "no data (this runs from the dashboard, not the editor)" });
  const d = JSON.parse(e.postData.contents);
  if (d.token !== TOKEN) return out_({ error: "unauthorized" });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (d.action === "saveOrder") {
      const o = d.order, c = d.customer || {};
      removeRows_(o.id);
      const orders = sheet_("Orders", HEAD);
      (o.products || []).forEach(p => orders.appendRow([o.id, o.date, o.status, o.cid, c.name, c.phone, c.address, c.notes, p.name, p.anime, p.cat, p.price, p.img || "", o.total, o.paid || 0, (o.total || 0) - (o.paid || 0), o.notes]));
      sheet_("Data", ["Order ID", "JSON"]).appendRow([o.id, JSON.stringify({ order: o, customer: c })]);
      return out_({ ok: true });
    }
    if (d.action === "deleteOrder") { removeRows_(d.id); return out_({ ok: true }); }
    return out_({ error: "unknown action" });
  } finally { lock.releaseLock(); }
}

// Run THIS function once from the editor (select "setup" > Run) to grant permissions
// and create the tabs. Do not run doGet/doPost directly; the dashboard calls them.
function setup() {
  sheet_("Orders", HEAD);
  sheet_("Data", ["Order ID", "JSON"]);
  Logger.log("Setup OK: Sheet is reachable.");
}
