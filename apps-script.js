// ═══════════════════════════════════════════════════
// Wedding RSVP — Google Apps Script  
// Re-deploy: Deploy → Manage deployments → Edit (pencil icon) → New version → Deploy
// ═══════════════════════════════════════════════════

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Parse incoming data (handles both text/plain and application/json)
  var contents = e.postData.contents;
  var data = typeof contents === 'string' ? JSON.parse(contents) : contents;
  
  // Find guest row by phone
  var phone = (data.phone || '').toString();
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var guestRow = -1;
  
  for (var i = 2; i <= lastRow; i++) {
    var cc = String(sheet.getRange(i, 6).getValue() || '');
    var ph = String(sheet.getRange(i, 7).getValue() || '');
    if ('+' + cc.trim() + ph.trim() === phone.trim()) {
      guestRow = i;
      break;
    }
  }
  
  if (guestRow === -1) {
    return json({success: false, error: 'Guest not found: ' + phone});
  }
  
  // Write to columns I–V (column 9–17)
  // I=First name, J=Last name, K=Email, L=Attending, M=Total guests,
  // N=Guest details, O=Days, P=Dietary, Q=Message, R=Submitted at
  sheet.getRange(guestRow, 9).setValue(data.first_name || '');
  sheet.getRange(guestRow, 10).setValue(data.last_name || '');
  sheet.getRange(guestRow, 11).setValue(data.email || '');
  sheet.getRange(guestRow, 12).setValue(data.attending || '');
  sheet.getRange(guestRow, 13).setValue(data.total_guests || 0);
  sheet.getRange(guestRow, 14).setValue(data.guests_detail || '');
  sheet.getRange(guestRow, 15).setValue(data.days || '');
  sheet.getRange(guestRow, 16).setValue(data.dietary_restrictions || '');
  sheet.getRange(guestRow, 17).setValue(data.message || '');
  sheet.getRange(guestRow, 18).setValue(new Date().toISOString());
  
  return json({success: true, guest: data.first_name, row: guestRow});
}

function doGet(e) {
  return json({status: 'ok', message: 'Wedding RSVP API is running'});
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
