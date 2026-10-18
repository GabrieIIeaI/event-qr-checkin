/**
 * Event QR Check-in — Google Apps Script backend (free).
 * Sheet columns (row 1 = headers, created automatically):
 *   A Name | B Email | C Code | D QR | E Status | F Checked-in at | G Email sent
 */

const SHEET_NAME = 'Guests';
const HEADERS = ['Name', 'Email', 'Code', 'QR', 'Status', 'Checked-in at', 'Email sent'];
const COL = { NAME: 1, EMAIL: 2, CODE: 3, QR: 4, STATUS: 5, TIME: 6, SENT: 7 };
const EVENT_NAME = 'My Event'; // <-- change to your event name (used in emails)

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Event QR')
    .addItem('1) Setup sheet', 'setupSheet')
    .addItem('2) Generate QR codes', 'generateCodes')
    .addItem('3) Email QR codes to guests', 'sendEmails')
    .addSeparator()
    .addItem('Set staff PIN', 'setStaffPin')
    .addItem('Reset ALL check-ins', 'resetCheckins')
    .addToUi();
}

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function setupSheet() {
  const sh = getSheet_();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(COL.QR, 160);
  SpreadsheetApp.getUi().alert('Sheet ready. Fill in Name and Email (rows 2+), then use "Generate QR codes".');
}

function newCode_() {
  // 12 random chars, no ambiguous characters. Unguessable enough for events.
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let out = '';
  const bytes = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  for (let i = 0; i < 12; i++) out += chars[parseInt(bytes.substr(i * 2, 2), 16) % chars.length];
  return out;
}

function qrUrl_(code, size) {
  return 'https://quickchart.io/qr?size=' + (size || 300) + '&margin=2&text=' + encodeURIComponent(code);
}

function generateCodes() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return SpreadsheetApp.getUi().alert('No guests yet. Add names starting on row 2.');
  const range = sh.getRange(2, 1, last - 1, HEADERS.length);
  const rows = range.getValues();
  const existing = new Set(rows.map(r => r[COL.CODE - 1]).filter(String));
  let created = 0;
  rows.forEach((r, i) => {
    if (!r[COL.NAME - 1] || r[COL.CODE - 1]) return;
    let code;
    do { code = newCode_(); } while (existing.has(code));
    existing.add(code);
    const row = i + 2;
    sh.getRange(row, COL.CODE).setValue(code);
    sh.getRange(row, COL.QR).setFormula('=IMAGE("' + qrUrl_(code, 150) + '")');
    sh.getRange(row, COL.STATUS).setValue('NOT USED');
    sh.setRowHeight(row, 150);
    created++;
  });
  SpreadsheetApp.getUi().alert(created + ' new QR code(s) generated.');
}

function sendEmails() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  let sent = 0, skipped = 0;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const name = r[COL.NAME - 1], email = r[COL.EMAIL - 1], code = r[COL.CODE - 1];
    if (!email || !code || r[COL.SENT - 1]) continue;
    if (MailApp.getRemainingDailyQuota() < 1) { skipped = rows.length - i; break; }
    const img = UrlFetchApp.fetch(qrUrl_(code, 400)).getBlob().setName('ticket-' + code + '.png');
    MailApp.sendEmail({
      to: email,
      subject: 'Your ticket for ' + EVENT_NAME,
      htmlBody:
        '<p>Hi ' + name + ',</p>' +
        '<p>This is your personal ticket for <b>' + EVENT_NAME + '</b>. ' +
        'Show this QR code at the entrance. It can be used only once.</p>' +
        '<p><img src="cid:qr" width="300" height="300"></p>' +
        '<p>Code: <b>' + code + '</b></p>',
      inlineImages: { qr: img },
      attachments: [img],
    });
    sh.getRange(i + 2, COL.SENT).setValue(new Date());
    sent++;
  }
  SpreadsheetApp.getUi().alert(sent + ' email(s) sent.' +
    (skipped ? ' Daily Gmail quota reached; run again tomorrow for the remaining ' + skipped + '.' : ''));
}

function setStaffPin() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Staff PIN', 'Choose a PIN/password the scanner phones will use:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK || !res.getResponseText().trim()) return;
  PropertiesService.getScriptProperties().setProperty('STAFF_PIN', res.getResponseText().trim());
  ui.alert('Staff PIN saved.');
}

function resetCheckins() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Mark ALL codes as NOT USED again?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const rows = sh.getRange(2, COL.CODE, last - 1, 1).getValues();
  rows.forEach((r, i) => {
    if (r[0]) sh.getRange(i + 2, COL.STATUS, 1, 2).setValues([['NOT USED', '']]);
  });
}

/** Web API used by the scanner page: GET ?action=checkin&code=XXX&pin=YYY */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const pin = PropertiesService.getScriptProperties().getProperty('STAFF_PIN');
  if (!pin) return json_({ status: 'error', message: 'Staff PIN not set. Use menu Event QR > Set staff PIN.' });
  if (p.pin !== pin) return json_({ status: 'unauthorized', message: 'Wrong staff PIN' });
  if (p.action === 'ping') return json_({ status: 'ok' });
  if (p.action !== 'checkin') return json_({ status: 'error', message: 'Unknown action' });

  const code = String(p.code || '').trim().toUpperCase();
  if (!code) return json_({ status: 'invalid' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // prevents two phones accepting the same code at the same time
  try {
    const sh = getSheet_();
    const last = sh.getLastRow();
    if (last < 2) return json_({ status: 'invalid', code: code });
    const found = sh.getRange(2, COL.CODE, last - 1, 1)
      .createTextFinder(code).matchEntireCell(true).findNext();
    if (!found) return json_({ status: 'invalid', code: code });
    const row = found.getRow();
    const vals = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
    const name = vals[COL.NAME - 1];
    if (vals[COL.STATUS - 1] === 'USED') {
      return json_({ status: 'used', name: name, code: code, usedAt: fmt_(vals[COL.TIME - 1]) });
    }
    const now = new Date();
    sh.getRange(row, COL.STATUS, 1, 2).setValues([['USED', now]]);
    SpreadsheetApp.flush();
    return json_({ status: 'ok', name: name, code: code, usedAt: fmt_(now) });
  } finally {
    lock.releaseLock();
  }
}

function fmt_(d) {
  if (!(d instanceof Date)) return String(d || '');
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
