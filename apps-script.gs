/* =========================================================
   كود Google Sheet لطلبات سفر الجاذرينج
   الصقه في: الشيت ← الإضافات ← Apps Script
   ========================================================= */

// كلمة سر شاشة الكاشير — غيّرها لكلمة خاصة فيك، ونفسها تكتبها في شاشة الكاشير
const SECRET = "CHANGE-ME-1234";

const SHEET_NAME = "الطلبات";
const HEADERS = ["رقم الطلب", "وقت الطلب", "الحالة", "السفرة", "الأصناف", "الإجمالي",
                 "الاسم", "الجوال", "طريقة التواصل", "تاريخ الجاذرينج", "ملاحظات", "لغة العميل"];
const STATUSES = ["جديد", "مؤكد", "جاهز", "تم التسليم", "ملغي"];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sh.setRightToLeft(true);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clean_(v, max) {
  return String(v == null ? "" : v).slice(0, max || 500).replace(/^[=+\-@]/, "'$&"); // يمنع تنفيذ معادلات
}

// استقبال طلب جديد، أو تغيير حالة طلب من الكاشير
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: "bad_json" }); }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();

    if (d.action === "status") {
      if (d.key !== SECRET) return json_({ ok: false, error: "key" });
      if (STATUSES.indexOf(d.status) < 0) return json_({ ok: false, error: "status" });
      const last = sh.getLastRow();
      if (last < 2) return json_({ ok: false, error: "not_found" });
      const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]));
      const i = ids.indexOf(String(d.id));
      if (i < 0) return json_({ ok: false, error: "not_found" });
      sh.getRange(i + 2, 3).setValue(d.status);
      return json_({ ok: true });
    }

    const o = d.order || {};
    if (!o.name || !o.phone || !o.sufra) return json_({ ok: false, error: "missing" });
    sh.appendRow([
      clean_(o.id, 20), new Date(), "جديد",
      clean_(o.sufra), clean_(o.items, 3000), clean_(o.total, 200),
      clean_(o.name, 120), "'" + String(o.phone).replace(/[^0-9+\s]/g, "").slice(0, 20), clean_(o.contact, 20),
      clean_(o.date, 20), clean_(o.notes, 1000), clean_(o.lang, 20)
    ]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// شاشة الكاشير تقرأ آخر الطلبات (تحتاج كلمة السر)
function doGet(e) {
  if ((e.parameter.key || "") !== SECRET) return json_({ ok: false, error: "key" });
  const v = sheet_().getDataRange().getValues();
  v.shift();
  const tz = "Asia/Riyadh";
  const fmtDate = x => x instanceof Date ? Utilities.formatDate(x, tz, "yyyy-MM-dd") : String(x);
  const orders = v.slice(-300).map(r => ({
    id: String(r[0]),
    time: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
    status: r[2], sufra: r[3], items: r[4], total: r[5],
    name: r[6], phone: String(r[7]), contact: r[8],
    date: fmtDate(r[9]), notes: r[10], lang: r[11]
  })).reverse();
  return json_({ ok: true, orders: orders });
}
