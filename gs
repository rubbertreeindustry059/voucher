// =============================================
// VOUCHER SYSTEM - Code.gs
// =============================================
// แก้ค่าด้านล่างนี้ก่อนใช้งาน
const SHEET_ID   = '1F81J-iXhEMokeLtNc77r9qYGHtdhtugiubXn2X2h8HA';   // ← Sheet ID ของคุณ
const STAFF_PIN  = '1234';            // ← PIN Staff
const ADMIN_PIN  = '9999';            // ← PIN Admin
// =============================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

// ฟังก์ชันนี้ถูกเรียกจาก index.html โดยตรง (google.script.run)
function processRequest(bodyStr) {
  try {
    var data   = JSON.parse(bodyStr);
    var action = data.action;
    if (action === 'login')          return login(data.pin);
    if (action === 'scanVoucher')    return scanVoucher(data.code);
    if (action === 'getHistory')     return getHistory();
    if (action === 'getAll')         return getAll();
    if (action === 'createVouchers') return createVouchers(data.prefix, data.count);
    return { ok: false, msg: 'Unknown action' };
  } catch(err) {
    return { ok: false, msg: err.toString() };
  }
}

function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var action = data.action;
    if (action === 'login')          return json(login(data.pin));
    if (action === 'scanVoucher')    return json(scanVoucher(data.code));
    if (action === 'getHistory')     return json(getHistory());
    if (action === 'getAll')         return json(getAll());
    if (action === 'createVouchers') return json(createVouchers(data.prefix, data.count));
    return json({ ok: false, msg: 'Unknown action' });
  } catch(err) {
    return json({ ok: false, msg: err.toString() });
  }
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function login(pin) {
  if (pin == ADMIN_PIN) return { ok: true, role: 'admin' };
  if (pin == STAFF_PIN) return { ok: true, role: 'staff' };
  return { ok: false, msg: 'PIN ไม่ถูกต้อง' };
}

function scanVoucher(code) {
  var sheet = getSheet('Vouchers');
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(code).trim()) {
      if (data[i][2] === 'ใช้แล้ว') {
        return { ok: false, msg: 'Voucher นี้ถูกใช้ไปแล้ว' };
      }
      var now     = new Date();
      var timeStr = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
      sheet.getRange(i + 1, 3).setValue('ใช้แล้ว');
      sheet.getRange(i + 1, 4).setValue(timeStr);
      getSheet('UsageLog').appendRow([code, timeStr, '']);
      return { ok: true, code: code, time: timeStr };
    }
  }
  return { ok: false, msg: 'ไม่พบ Voucher นี้ในระบบ' };
}

function getHistory() {
  var sheet = getSheet('UsageLog');
  var data  = sheet.getDataRange().getValues();
  var rows  = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) rows.push({ code: data[i][0], time: data[i][1] });
  }
  rows.reverse();
  return { ok: true, rows: rows };
}

function getAll() {
  var sheet = getSheet('Vouchers');
  var data  = sheet.getDataRange().getValues();
  var rows  = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      rows.push({
        code:   data[i][0],
        status: data[i][2],
        time:   data[i][3] || ''
      });
    }
  }
  var used = rows.filter(function(r){ return r.status === 'ใช้แล้ว'; }).length;
  return { ok: true, rows: rows, total: rows.length, used: used, unused: rows.length - used };
}

function createVouchers(prefix, count) {
  var sheet   = getSheet('Vouchers');
  var created = [];
  var p       = (prefix || 'VC').toUpperCase();
  
  for (var i = 0; i < count; i++) {
    var code  = p + '-' + Math.random().toString(36).toUpperCase().slice(2, 6);
    
    // สร้างสูตร IMAGE โดยดึงค่าจาก Column A ในแถวนั้นๆ มาสร้าง QR Code
    // ใช้สูตร: =IMAGE("https://quickchart.io/qr?text=" & A[row] & "&size=200")
    // ใน Google Apps Script เราจะใช้คำสั่งสูตรแบบ R1C1 หรือจะเขียนแบบ String ต่อกันก็ได้
    // แต่เพื่อให้ "ก๊อปวาง" แล้วใช้ได้ทันที ผมจะใช้สูตรที่อ้างอิงรหัสโดยตรงครับ
    
    var qrFormula = '=IMAGE("https://quickchart.io/qr?text=' + code + '&size=200")';
    
    sheet.appendRow([code, qrFormula, 'ยังไม่ใช้', '', '']);
    created.push(code);
  }
  
  return { ok: true, created: created };
}
