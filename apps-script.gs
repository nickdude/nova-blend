// Nova Blend — Google Apps Script collector
// Deploy as a web app: Execute as "Me", access "Anyone"
// Paste the deployment URL into NB_CONFIG.COLLECT_URL in index.html

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (data.type === 'lead') appendLead(ss, data.data);
    else if (data.type === 'event') appendEvent(ss, data.data);
  } catch (err) {}
  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}

function appendLead(ss, d) {
  var sheet = ss.getSheetByName('Leads') || ss.insertSheet('Leads');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'timestamp', 'email', 'phone', 'whatsapp_ok',
      'reason', 'reason_other', 'from_checkout',
      'cart', 'total (₹)', 'city', 'state', 'pay_method',
      'variant', 'utm_source', 'utm_medium', 'utm_campaign', 'visitor_id'
    ]);
    sheet.getRange(1, 1, 1, 17).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  var last = d.last || {};
  sheet.appendRow([
    d.t, d.email, d.phone, d.whatsapp_ok ? 'yes' : 'no',
    d.reason, d.reason_other, d.from_checkout ? 'yes' : 'no',
    d.cart, d.total, d.city, d.state, d.pay,
    d.variant, last.utm_source || '', last.utm_medium || '', last.utm_campaign || '', d.vid
  ]);
}

function appendEvent(ss, d) {
  var sheet = ss.getSheetByName('Events') || ss.insertSheet('Events');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['timestamp', 'event', 'visitor_id', 'session_id', 'variant', 'props', 'utm_source', 'screen_width']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  var props = Object.assign({}, d);
  ['event', 't', 'n', 'vid', 'sid', 'variant', 'w', 'first', 'last'].forEach(function (k) { delete props[k]; });
  var last = d.last || {};
  sheet.appendRow([d.t, d.event, d.vid, d.sid, d.variant, JSON.stringify(props), last.utm_source || '', d.w]);
}
