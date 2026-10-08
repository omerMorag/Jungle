/**
 * Election Party: registrations → Google Sheet
 *
 * Setup (one time):
 * 1. Create a new Google Sheet in your Drive (e.g. "Election Party – registrations").
 * 2. In the Sheet: Extensions → Apps Script. Delete what's there and paste this whole file. Save.
 * 3. Deploy → New deployment → type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 *    Approve the permissions, then copy the Web app URL.
 * 4. Paste that URL into the site's settings line:  const SHEET_URL = "https://script.google.com/macros/s/.../exec";
 *
 * Each person in an order becomes one row. The buyer's row carries the phone number.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const order = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['זמן', 'מספר הזמנה', 'כרטיס', 'מחיר', 'כמות בהזמנה', 'שם פרטי', 'שם משפחה', 'מין', 'טלפון', 'תפקיד']);
      sheet.setFrozenRows(1);
    }
    (order.people || []).forEach(function (p, i) {
      sheet.appendRow([
        new Date(order.createdAt), order.orderId, order.ticket, order.price, order.qty,
        p.first, p.last, p.gender,
        i === 0 ? "'" + order.phone : '',
        i === 0 ? 'מזמין/ה' : 'חבר/ה'
      ]);
    });
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
