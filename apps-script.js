// Wedding RSVP — Google Apps Script
// Re-deploy: Deploy → Manage deployments → Edit → New version → Deploy

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);
  var phone = (data.phone || '').toString();
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var guestRow = -1;
  
  for (var i = 2; i <= lastRow; i++) {
    var cc = String(sheet.getRange(i, 6).getValue() || '');
    var ph = String(sheet.getRange(i, 7).getValue() || '');
    if ('+' + cc.trim() + ph.trim() === phone.trim()) { guestRow = i; break; }
  }
  
  if (guestRow === -1) return json({success: false, error: 'Guest not found'});
  
  // I = Attending status
  sheet.getRange(guestRow, 9).setValue(data.attending === 'yes' ? 'Attending' : (data.attending === 'no' ? 'Decline' : ''));
  // J = First name
  sheet.getRange(guestRow, 10).setValue(data.first_name || '');
  // K = Last name
  sheet.getRange(guestRow, 11).setValue(data.last_name || '');
  // L = Email
  sheet.getRange(guestRow, 12).setValue(data.email || '');
  // M = Total guests confirmed
  sheet.getRange(guestRow, 13).setValue(data.total_guests || 0);
  // N = Guest 1 name
  sheet.getRange(guestRow, 14).setValue(data.guest1 || '');
  // O = Guest 2 name
  sheet.getRange(guestRow, 15).setValue(data.guest2 || '');
  // P = Guest 3 name
  sheet.getRange(guestRow, 16).setValue(data.guest3 || '');
  // Q = Attend 3 days
  sheet.getRange(guestRow, 17).setValue(data.all_days ? 'Yes' : 'No');
  // R = Friday
  sheet.getRange(guestRow, 18).setValue(data.friday ? 'Yes' : 'No');
  // S = Saturday
  sheet.getRange(guestRow, 19).setValue(data.saturday ? 'Yes' : 'No');
  // T = Sunday
  sheet.getRange(guestRow, 20).setValue(data.sunday ? 'Yes' : 'No');
  // U = Dietary
  sheet.getRange(guestRow, 21).setValue(data.dietary_restrictions || '');
  // V = Message
  sheet.getRange(guestRow, 22).setValue(data.message || '');
  // W = Submitted at
  sheet.getRange(guestRow, 23).setValue(new Date().toISOString());
  
  return json({success: true, guest: data.first_name, row: guestRow});
}

function doGet(e) {
  return json({status: 'ok'});
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
