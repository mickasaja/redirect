// ════════════════════════════════════════════
// Wedding RSVP Apps Script — FINAL VERSION
// 1. Paste in Extensions → Apps Script
// 2. Deploy → New deployment → Web App
//    Execute as: Me, Access: Anyone
// 3. Copy URL and send to me
// ════════════════════════════════════════════

function doPost(e) {
  return ContentService.createTextOutput(JSON.stringify(handlePost(e)))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({status:'ok'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function handlePost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    var phone = String(data.phone || '').trim();
    var lastRow = Math.max(sheet.getLastRow(), 2);
    var guestRow = -1;
    
    for (var i = 2; i <= lastRow; i++) {
      var cc = String(sheet.getRange(i, 6).getValue() || '').trim();
      var ph = String(sheet.getRange(i, 7).getValue() || '').trim();
      if ('+' + cc + ph === phone) { guestRow = i; break; }
    }
    
    if (guestRow === -1) {
      return {success: false, error: 'Guest not found: ' + phone};
    }
    
    sheet.getRange(guestRow, 9).setValue(data.attending === 'yes' ? 'Attending' : 'Decline');
    sheet.getRange(guestRow, 10).setValue(data.first_name || '');
    sheet.getRange(guestRow, 11).setValue(data.last_name || '');
    sheet.getRange(guestRow, 12).setValue(data.email || '');
    sheet.getRange(guestRow, 13).setValue(data.total_guests || 0);
    sheet.getRange(guestRow, 14).setValue(data.guest1 || '');
    sheet.getRange(guestRow, 15).setValue(data.guest2 || '');
    sheet.getRange(guestRow, 16).setValue(data.guest3 || '');
    sheet.getRange(guestRow, 17).setValue(data.all_days ? 'Yes' : 'No');
    sheet.getRange(guestRow, 18).setValue(data.friday ? 'Yes' : 'No');
    sheet.getRange(guestRow, 19).setValue(data.saturday ? 'Yes' : 'No');
    sheet.getRange(guestRow, 20).setValue(data.sunday ? 'Yes' : 'No');
    sheet.getRange(guestRow, 21).setValue(data.dietary_restrictions || '');
    sheet.getRange(guestRow, 22).setValue(data.message || '');
    sheet.getRange(guestRow, 23).setValue(new Date().toISOString());
    
    return {success: true, guest: data.first_name, row: guestRow};
  } catch(err) {
    return {success: false, error: String(err)};
  }
}
