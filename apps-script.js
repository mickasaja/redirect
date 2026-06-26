function doGet(e) {
  var p = e.parameter;
  if (!p.phone) return out({status:'ok'});
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var phone = String(p.phone).trim().replace(/\s/g,'+'); // fix +→space
  if (phone.charAt(0)!=='+') phone = '+' + phone;
  var r = -1;
  for (var i = 1; i < data.length; i++) {
    var cc = String(data[i][5]||'').trim();
    var ph = String(data[i][6]||'').trim();
    if ('+' + cc + ph === phone) { r = i + 1; break; }
  }
  if (r === -1) return out({success:false,phone:phone});
  sheet.getRange(r,9).setValue(p.attending==='yes'?'Attending':'Decline');
  sheet.getRange(r,10).setValue(p.first_name||'');
  sheet.getRange(r,11).setValue(p.last_name||'');
  sheet.getRange(r,12).setValue(p.email||'');
  sheet.getRange(r,13).setValue(p.total_guests||0);
  sheet.getRange(r,14).setValue(p.guest1||'');
  sheet.getRange(r,15).setValue(p.guest2||'');
  sheet.getRange(r,16).setValue(p.guest3||'');
  sheet.getRange(r,17).setValue(p.all_days==='true'?'Yes':'No');
  sheet.getRange(r,18).setValue(p.friday==='true'?'Yes':'No');
  sheet.getRange(r,19).setValue(p.saturday==='true'?'Yes':'No');
  sheet.getRange(r,20).setValue(p.sunday==='true'?'Yes':'No');
  sheet.getRange(r,21).setValue(p.dietary||'');
  sheet.getRange(r,22).setValue(p.message||'');
  sheet.getRange(r,23).setValue(new Date().toISOString());
  return out({success:true,r:r});
}
function out(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
