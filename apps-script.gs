/**
 * استقبال طلبات صفحة الهبوط في Google Sheets
 * انسخ الكود ده كله في Apps Script (الخطوات في دليل-التشغيل.md)
 */

const SHEET_NAME = "الطلبات";
const HEADERS = ["التاريخ", "المنتج", "كود تاجر", "الكمية", "الإجمالي", "الاسم", "موبايل", "موبايل 2",
                 "المحافظة", "العنوان", "المصدر", "الحالة", "اتسجل في تاجر؟", "ملاحظات"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const p = e.parameter || {};
    sheet.appendRow([
      p.time || new Date(), p.product, p.taager_code, p.qty, p.total,
      p.name, "'" + (p.phone || ""), p.phone2 ? "'" + p.phone2 : "",
      p.gov, p.address, p.source, "جديد", "", ""
    ]);
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#0b4f6c").setFontColor("#ffffff");
    const status = SpreadsheetApp.newDataValidation()
      .requireValueInList(["جديد", "اتأكد", "مارديش", "اتلغى", "اتسلّم", "رفض الاستلام"]).build();
    sheet.getRange("L2:L").setDataValidation(status);
    const done = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange("M2:M").setDataValidation(done);
  }
  return sheet;
}

/** شغّل الدالة دي مرة واحدة يدوي عشان تعمل الشيت وتدي الصلاحيات */
function setup() {
  getSheet_();
}
