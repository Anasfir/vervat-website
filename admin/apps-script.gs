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

// 2. Email address that receives every submission (with uploaded photos attached). '' = off.
const NOTIFY_EMAIL = 'Intissar@vervat-breukelen.nl';

// 3. WhatsApp copy via CallMeBot (free). '' = off. See the admin panel (Instellingen) for the 1-minute setup.
//    WHATSAPP_PHONE: international format with +, e.g. '+31620623248'.
const WHATSAPP_PHONE = '+31620623248';
const CALLMEBOT_APIKEY = '';

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
    const blobs = [];
    const files = Array.isArray(d.files) ? d.files.slice(0, 3) : [];
    if (files.length) {
      const folder = folder_();
      files.forEach(function (f, i) {
        if (!/^(image\/(jpeg|png)|application\/pdf)$/.test(f.type)) return;
        const bytes = Utilities.base64Decode(f.data);
        if (bytes.length > 5 * 1024 * 1024) return;
        const name = id + '-' + (i + 1) + '-' + clean_(f.name, 80).replace(/[^\w.\-]+/g, '_');
        const blob = Utilities.newBlob(bytes, f.type, name);
        const file = folder.createFile(blob);
        links.push(file.getUrl());
        blobs.push(blob);
      });
    }
    const fields = {};
    Object.keys(d.fields || {}).forEach(function (k) { fields[clean_(k, 40)] = clean_(d.fields[k], 4000); });
    sheet_().appendRow([new Date(), id, type, clean_(d.name, 200), clean_(d.email, 200), clean_(d.phone, 60),
      JSON.stringify(fields), links.join('\n'), 'new', clean_(d.lang, 5), clean_(d.page, 300)]);
    notify_(type, id, d, fields, links, blobs);
    return json_({ ok: true, id: id });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

const TYPE_NL = { service: 'Serviceverzoek', contact: 'Contactbericht', newsletter: 'Aanmelding nieuwsbrief', other: 'Bericht' };
const LABEL_NL = { company: 'Bedrijf', street: 'Straat en huisnummer', postcode: 'Postcode', city: 'Plaats', product: 'Product',
  year: 'Installatiejaar', suppliedByUs: 'Door ons geleverd', description: 'Omschrijving', subject: 'Onderwerp', message: 'Bericht' };

// Email + WhatsApp notification. Failures here never block saving the submission.
function notify_(type, id, d, fields, links, blobs) {
  const rows = [['Naam', clean_(d.name)], ['E-mail', clean_(d.email)], ['Telefoon', clean_(d.phone)]];
  Object.keys(fields).forEach(function (k) { if (fields[k]) rows.push([LABEL_NL[k] || k, fields[k]]); });
  const title = TYPE_NL[type] || 'Bericht';
  const subject = 'Website: ' + title + (d.name ? ' van ' + clean_(d.name, 80) : '') + (fields.product ? ' - ' + fields.product : '');
  if (NOTIFY_EMAIL) {
    try {
      const h = function (s) { return String(s).replace(/[&<>]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]; }).replace(/\n/g, '<br>'); };
      const html = '<h2 style="font-family:Arial">' + h(title) + '</h2><table cellpadding="6" style="font-family:Arial;font-size:14px;border-collapse:collapse">' +
        rows.filter(function (r) { return r[1]; }).map(function (r) { return '<tr><td style="color:#666;vertical-align:top">' + h(r[0]) + '</td><td>' + h(r[1]) + '</td></tr>'; }).join('') +
        (links.length ? '<tr><td style="color:#666">Bestanden</td><td>' + links.map(function (l, i) { return '<a href="' + l + '">Bestand ' + (i + 1) + '</a>'; }).join('<br>') + '</td></tr>' : '') +
        '</table><p style="font-family:Arial;color:#888;font-size:12px">Referentie ' + id + '. Beantwoorden gaat direct naar de klant.</p>';
      const opt = { htmlBody: html, name: 'Vervat & de Jong website' };
      if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(d.email || ''))) opt.replyTo = String(d.email);
      if (blobs.length) opt.attachments = blobs;
      MailApp.sendEmail(NOTIFY_EMAIL, subject, rows.map(function (r) { return r[0] + ': ' + r[1]; }).join('\n'), opt);
    } catch (err) { console.error('email failed', err); }
  }
  if (WHATSAPP_PHONE && CALLMEBOT_APIKEY) {
    try {
      let text = '*' + subject + '*\n' + rows.filter(function (r) { return r[1]; }).map(function (r) { return r[0] + ': ' + String(r[1]).slice(0, 400); }).join('\n');
      if (links.length) text += '\nBestanden: ' + links.join(' ');
      UrlFetchApp.fetch('https://api.callmebot.com/whatsapp.php?phone=' + encodeURIComponent(WHATSAPP_PHONE) +
        '&text=' + encodeURIComponent(text.slice(0, 1500)) + '&apikey=' + encodeURIComponent(CALLMEBOT_APIKEY), { muteHttpExceptions: true });
    } catch (err) { console.error('whatsapp failed', err); }
  }
}

// Run once from the editor (choose "testNotify" and click Run) to check email and WhatsApp.
function testNotify() {
  notify_('contact', 'test', { name: 'Test', email: 'test@example.com', phone: '0612345678' },
    { subject: 'Test', message: 'Dit is een testbericht van het websiteformulier.' }, [], []);
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
