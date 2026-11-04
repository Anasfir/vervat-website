/**
 * Vervat & de Jong — form inbox
 * Receives the website's contact, service and newsletter forms,
 * stores them in this Google Sheet, saves uploaded photos to Google Drive,
 * and lets the admin panel read and update them.
 *
 * Setup: see "Form submissions" in the admin panel.
 */

// 1. Choose a long secret key. Paste the SAME key in the admin panel (Settings).
const ADMIN_KEY = 'CHANGE-ME-to-a-long-secret';

// 2. Optional: email address that gets a message for every new submission ('' = off).
const NOTIFY_EMAIL = '';

const SHEET_NAME = 'Submissions';
const FOLDER_NAME = 'Vervat website uploads';
const HEADERS = ['Date', 'ID', 'Type', 'Name', 'Email', 'Phone', 'Details', 'Files', 'Status', 'Language', 'Page'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clean_(v, max) {
  return String(v == null ? '' : v).slice(0, max || 500);
}

// Website forms post here.
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const type = ['contact', 'service', 'newsletter'].indexOf(d.type) >= 0 ? d.type : 'other';
    const id = Utilities.getUuid().slice(0, 8);
    const links = [];
    const files = Array.isArray(d.files) ? d.files.slice(0, 3) : [];
    if (files.length) {
      const folder = folder_();
      files.forEach(function (f, i) {
        if (!/^(image\/(jpeg|png)|application\/pdf)$/.test(f.type)) return;
        const bytes = Utilities.base64Decode(f.data);
        if (bytes.length > 5 * 1024 * 1024) return;
        const name = id + '-' + (i + 1) + '-' + clean_(f.name, 80).replace(/[^\w.\-]+/g, '_');
        const file = folder.createFile(Utilities.newBlob(bytes, f.type, name));
        links.push(file.getUrl());
      });
    }
    const fields = {};
    Object.keys(d.fields || {}).forEach(function (k) { fields[clean_(k, 40)] = clean_(d.fields[k], 4000); });
    sheet_().appendRow([new Date(), id, type, clean_(d.name, 200), clean_(d.email, 200), clean_(d.phone, 60),
      JSON.stringify(fields), links.join('\n'), 'new', clean_(d.lang, 5), clean_(d.page, 300)]);
    if (NOTIFY_EMAIL) {
      const lines = ['Type: ' + type, 'Name: ' + clean_(d.name), 'Email: ' + clean_(d.email), 'Phone: ' + clean_(d.phone)];
      Object.keys(fields).forEach(function (k) { lines.push(k + ': ' + fields[k]); });
      if (links.length) lines.push('Files:\n' + links.join('\n'));
      MailApp.sendEmail(NOTIFY_EMAIL, 'Website: new ' + type + ' request', lines.join('\n'));
    }
    return json_({ ok: true, id: id });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// The admin panel reads and updates submissions here.
function doGet(e) {
  const p = e.parameter || {};
  if (!ADMIN_KEY || ADMIN_KEY.indexOf('CHANGE-ME') === 0 || p.key !== ADMIN_KEY) {
    return json_({ ok: false, error: 'unauthorized' });
  }
  const sh = sheet_();
  const values = sh.getDataRange().getValues();
  if (p.action === 'status' || p.action === 'delete') {
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][1]) === p.id) {
        if (p.action === 'delete') sh.deleteRow(r + 1);
        else sh.getRange(r + 1, 9).setValue(clean_(p.status, 20));
        return json_({ ok: true });
      }
    }
    return json_({ ok: false, error: 'not found' });
  }
  const items = values.slice(1).map(function (row) {
    let details = {};
    try { details = JSON.parse(row[6] || '{}'); } catch (x) {}
    return {
      date: row[0] instanceof Date ? row[0].toISOString() : String(row[0]), id: String(row[1]), type: row[2],
      name: row[3], email: row[4], phone: row[5], details: details,
      files: String(row[7] || '').split('\n').filter(String), status: row[8] || 'new', lang: row[9], page: row[10]
    };
  }).reverse();
  return json_({ ok: true, items: items });
}
