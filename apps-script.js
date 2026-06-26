function doGet(e) {
  var p = e.parameter;
  var d = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues();
  var out = {
    received_phone: p.phone,
    received_cc: p.cc,
    received_phone_num: p.phone_num,
    row1_F: String(d[1][5]),
    row1_G: String(d[1][6]),
    combined: '+' + String(d[1][5]).trim() + String(d[1][6]).trim()
  };
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
