// Wedding RSVP — OPTIMIZED (fast batch read)
// Deploy as Web App → Execute as Me → Anyone → NEW deployment (not edit)

function doGet(e) {
  var p = e.parameter;
  if (!p.phone) return out({status:'ok'});
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues(); // Read ALL at once
  var phone = String(p.phone).trim();
  var guestRow = -1;
  
  for (var i = 1; i < data.length; i++) {
    var cc = String(data[i][5] || '').trim();  // Column F (index 5)
    var ph = String(data[i][6] || '').trim();  // Column G (index 6)
    if ('+' + cc + ph === phone) { guestRow = i + 1; break; }
  }
  
  if (guestRow === -1) return out({success: false, error: 'Not found'});
  
  sheet.getRange(guestRow, 9).setValue(p.attending === 'yes' ? 'Attending' : 'Decline');
  sheet.getRange(guestRow, 10).setValue(p.first_name || '');
  sheet.getRange(guestRow, 11).setValue(p.last_name || '');
  sheet.getRange(guestRow, 12).setValue(p.email || '');
  sheet.getRange(guestRow, 13).setValue(p.total_guests || 0);
  sheet.getRange(guestRow, 14).setValue(p.guest1 || '');
  sheet.getRange(guestRow, 15).setValue(p.guest2 || '');
  sheet.getRange(guestRow, 16).setValue(p.guest3 || '');
  sheet.getRange(guestRow, 17).setValue(p.all_days === 'true' ? 'Yes' : 'No');
  sheet.getRange(guestRow, 18).setValue(p.friday === 'true' ? 'Yes' : 'No');
  sheet.getRange(guestRow, 19).setValue(p.saturday === 'true' ? 'Yes' : 'No');
  sheet.getRange(guestRow, 20).setValue(p.sunday === 'true' ? 'Yes' : 'No');
  sheet.getRange(guestRow, 21).setValue(p.dietary || '');
  sheet.getRange(guestRow, 22).setValue(p.message || '');
  sheet.getRange(guestRow, 23).setValue(new Date().toISOString());
  
  return out({success: true});
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
