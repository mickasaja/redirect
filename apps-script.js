// ═══════════════════════════════════════════════════
// Wedding RSVP — Google Apps Script
// Paste this into: Extensions → Apps Script
// Deploy as Web App → Execute as "Me" → Access "Anyone"
// ═══════════════════════════════════════════════════

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);
  
  // Find the guest row by phone number (column F=country code + G=phone)
  var phone = data.phone || '';
  var lastRow = sheet.getLastRow();
  var guestRow = -1;
  
  for (var i = 2; i <= lastRow; i++) {
    var cc = sheet.getRange(i, 6).getValue().toString();  // Column F
    var ph = sheet.getRange(i, 7).getValue().toString();  // Column G
    var fullPhone = '+' + cc + ph;
    if (fullPhone === phone) {
      guestRow = i;
      break;
    }
  }
  
  if (guestRow === -1) {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: 'Guest not found'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Write answers to columns I–V
  var values = [
    data.first_name || '',        // I: First Name
    data.last_name || '',         // J: Last Name
    data.email || '',             // K: Email
    data.attending || '',         // L: Attending (yes/no)
    data.total_guests || '',      // M: Total Guests
    data.guests_detail || '',     // N: Guest Details
    data.days || '',              // O: Days attending
    data.dietary_restrictions || '', // P: Dietary
    data.message || '',           // Q: Message
    new Date().toISOString()      // R: Submitted at
  ];
  
  for (var j = 0; j < values.length; j++) {
    sheet.getRange(guestRow, 9 + j).setValue(values[j]); // Column I = 9
  }
  
  return ContentService.createTextOutput(JSON.stringify({success: true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput('Wedding RSVP API is running')
    .setMimeType(ContentService.MimeType.TEXT);
}
