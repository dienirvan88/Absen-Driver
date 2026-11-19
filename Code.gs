// Tempelkan kode ini di Google Sheets > Ekstensi > Apps Script
const SHEET_NAME = 'Absensi';

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'driverId', 'driverName', 'plate', 'date', 'clockIn', 'clockOut', 'note', 'createdAt']);
  }
  return sheet;
}

function doGet(e) {
  const sheet = getOrCreateSheet();
  const rows = sheet.getDataRange().getValues();
  const logs = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    logs.push({
      id: String(r[0]),
      driverId: String(r[1] || 'drv-1'),
      date: formatDateCell(r[4]),
      clockIn: formatTimeCell(r[5]),
      clockOut: formatTimeCell(r[6]),
      note: String(r[7] || '')
    });
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', logs: logs }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const sheet = getOrCreateSheet();
  const payload = JSON.parse(e.postData.contents);

  if (payload.action === 'add' && payload.log) {
    const l = payload.log;
    sheet.appendRow([
      l.id,
      l.driverId || 'drv-1',
      payload.driverName || 'Arifin',
      payload.plate || 'B 1712 WMW',
      l.date,
      l.clockIn,
      l.clockOut,
      l.note || '',
      new Date().toISOString()
    ]);
  } else if (payload.action === 'delete' && payload.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === String(payload.id)) {
        sheet.deleteRow(i + 1);
      }
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDateCell(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val || '').slice(0, 10);
}

function formatTimeCell(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
  }
  return String(val || '').slice(0, 5);
}
