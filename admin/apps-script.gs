/**
 * Vervat & de Jong — form inbox and newsletter
 * Receives the website's contact, service and newsletter forms,
 * stores them in this Google Sheet, saves uploaded photos to Google Drive,
 * lets the admin panel read and update them, and sends the weekly newsletter.
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

// 4. Website address (used for links and images in newsletters).
const SITE_URL = 'https://anasfir.github.io/vervat-website/';

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
    if (d.action) {
      if (!ADMIN_KEY || ADMIN_KEY.indexOf('CHANGE-ME') === 0 || d.key !== ADMIN_KEY) return json_({ ok: false, error: 'unauthorized' });
      try { return json_(adminAction_(d)); } catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
    }
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
    if (type === 'newsletter') {
      const st = subscribe_(d.email, d.lang, 'website');
      return json_({ ok: st !== 'invalid', id: id, status: st });
    }
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
  if (p.action === 'confirm' || p.action === 'unsub') return json_(publicAction_(p));
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

/* =====================================================================
   NEWSLETTER
   Subscribers (double opt-in), weekly automatic sending, unsubscribe.
   The admin panel writes the newsletter and sends it here.
   ===================================================================== */
const SUBS_SHEET = 'Subscribers';
const SUBS_HEADERS = ['Email', 'Language', 'Status', 'Signed up', 'Confirmed', 'Unsubscribed', 'Source'];
const NL_SHEET = 'Newsletters';
const NL_HEADERS = ['ID', 'Status', 'Subject NL', 'Subject EN', 'HTML NL', 'HTML EN', 'Source', 'Created', 'Sent at', 'Recipients'];
const TZ = 'Europe/Amsterdam';
const WEEKDAYS = ['MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY'];

function sheetOf_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}
const props_ = () => PropertiesService.getScriptProperties();
function getProp_(k, d) { try { const v = props_().getProperty(k); return v == null ? d : JSON.parse(v); } catch (x) { return d; } }
function setProp_(k, v) { props_().setProperty(k, JSON.stringify(v)); }
const isEmail_ = s => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(s || ''));
const iso_ = v => v instanceof Date ? v.toISOString() : String(v || '');

function token_(email) {
  const sig = Utilities.computeHmacSha256Signature(String(email).toLowerCase().trim(), ADMIN_KEY + '|newsletter');
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '').slice(0, 22);
}
function link_(action, email) {
  const base = SITE_URL.replace(/[?#].*$/, '').replace(/\/?$/, '/');
  return base + '?nl=' + action + '&e=' + encodeURIComponent(email) + '&t=' + token_(email);
}

function findSub_(sh, email) {
  const vals = sh.getDataRange().getValues();
  const e = String(email).toLowerCase().trim();
  for (let r = 1; r < vals.length; r++) if (String(vals[r][0]).toLowerCase().trim() === e) return { row: r + 1, v: vals[r] };
  return null;
}

// Called for every newsletter sign-up from the website.
function subscribe_(email, lang, source) {
  email = String(email || '').trim();
  if (!isEmail_(email)) return 'invalid';
  lang = lang === 'en' ? 'en' : 'nl';
  const sh = sheetOf_(SUBS_SHEET, SUBS_HEADERS);
  const f = findSub_(sh, email);
  if (f && f.v[2] === 'active') return 'already';
  if (f) { sh.getRange(f.row, 2, 1, 6).setValues([[lang, 'pending', new Date(), '', '', source || 'website']]); }
  else sh.appendRow([email, lang, 'pending', new Date(), '', '', source || 'website']);
  sendConfirm_(email, lang);
  return 'pending';
}

function sendConfirm_(email, lang) {
  const nl = lang !== 'en';
  const url = link_('confirm', email);
  const title = nl ? 'Bevestig uw aanmelding' : 'Please confirm your subscription';
  const text = nl ? 'Bedankt voor uw aanmelding voor de nieuwsbrief van Vervat &amp; de Jong. Klik op de knop om uw aanmelding te bevestigen.'
    : 'Thank you for subscribing to the Vervat &amp; de Jong newsletter. Click the button to confirm your subscription.';
  const btn = nl ? 'Ja, ik wil de nieuwsbrief' : 'Yes, subscribe me';
  const foot = nl ? 'Heeft u zich niet aangemeld? Dan kunt u deze e-mail negeren.' : 'Did not sign up? You can ignore this email.';
  const html = '<div style="background:#f3f3f0;padding:30px 12px;font-family:Arial,Helvetica,sans-serif"><table role="presentation" width="100%" style="max-width:560px;margin:0 auto;background:#fff;border-radius:12px" cellpadding="0" cellspacing="0">' +
    '<tr><td style="background:#232322;padding:22px 28px;border-radius:12px 12px 0 0"><img src="' + SITE_URL.replace(/\/?$/, '/') + 'assets/email/logo.png" width="150" alt="Vervat &amp; de Jong" style="display:block"></td></tr>' +
    '<tr><td style="padding:30px 28px;color:#1b1b1a;font-size:15px;line-height:1.6"><h1 style="font-size:22px;margin:0 0 14px">' + title + '</h1><p style="margin:0 0 24px">' + text + '</p>' +
    '<a href="' + url + '" style="display:inline-block;background:#F9C41E;color:#1b1b1a;text-decoration:none;font-weight:bold;padding:14px 26px;border-radius:99px">' + btn + '</a>' +
    '<p style="margin:26px 0 0;color:#888;font-size:12px">' + foot + '</p></td></tr></table></div>';
  MailApp.sendEmail({ to: email, subject: (nl ? 'Bevestig uw aanmelding – Vervat & de Jong' : 'Confirm your subscription – Vervat & de Jong'), htmlBody: html, name: 'Vervat & de Jong', replyTo: NOTIFY_EMAIL || undefined });
}

// Public: links in emails open the website, which calls this.
function publicAction_(p) {
  const email = String(p.e || '').trim();
  if (!isEmail_(email) || p.t !== token_(email)) return { ok: false, error: 'badlink' };
  const sh = sheetOf_(SUBS_SHEET, SUBS_HEADERS);
  const f = findSub_(sh, email);
  if (p.action === 'confirm') {
    if (!f) sh.appendRow([email, 'nl', 'active', new Date(), new Date(), '', 'link']);
    else sh.getRange(f.row, 3, 1, 4).setValues([['active', f.v[3] || new Date(), new Date(), '']]);
    return { ok: true, status: 'active' };
  }
  if (p.action === 'unsub') {
    if (f) sh.getRange(f.row, 3, 1, 4).setValues([['unsubscribed', f.v[3], f.v[4], new Date()]]);
    return { ok: true, status: 'unsubscribed' };
  }
  return { ok: false, error: 'unknown' };
}

function subscribers_() {
  const vals = sheetOf_(SUBS_SHEET, SUBS_HEADERS).getDataRange().getValues().slice(1);
  return vals.filter(r => r[0]).map(r => ({ email: String(r[0]), lang: r[1] || 'nl', status: r[2] || 'pending', signedUp: iso_(r[3]), confirmed: iso_(r[4]), unsubscribed: iso_(r[5]), source: r[6] || '' }));
}

function issues_() {
  const vals = sheetOf_(NL_SHEET, NL_HEADERS).getDataRange().getValues();
  const out = [];
  for (let r = 1; r < vals.length; r++) {
    const v = vals[r]; if (!v[0]) continue;
    out.push({ row: r + 1, id: String(v[0]), status: v[1], subject: { nl: v[2], en: v[3] }, html: { nl: v[4], en: v[5] }, source: v[6], created: iso_(v[7]), sentAt: iso_(v[8]), recipients: v[9] || 0 });
  }
  return out;
}
const openIssue_ = () => issues_().filter(i => i.status === 'draft' || i.status === 'scheduled' || i.status === 'sending').pop() || null;

function saveIssue_(d, status) {
  const sh = sheetOf_(NL_SHEET, NL_HEADERS);
  const cur = openIssue_();
  if (cur && cur.status === 'sending') throw new Error('Er wordt op dit moment een nieuwsbrief verstuurd. Probeer het straks opnieuw.');
  const row = [cur ? cur.id : Utilities.getUuid().slice(0, 8), status, clean_(d.subject.nl, 200), clean_(d.subject.en, 200),
    String(d.html.nl || '').slice(0, 49000), String(d.html.en || '').slice(0, 49000), String(d.source || '').slice(0, 49000), cur ? cur.created : new Date(), '', ''];
  if (cur) sh.getRange(cur.row, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  return row[0];
}

function setSchedule_(s) {
  const sched = { enabled: !!s.enabled, day: WEEKDAYS.indexOf(s.day) >= 0 ? s.day : 'MONDAY', hour: Math.max(0, Math.min(23, parseInt(s.hour, 10) || 9)) };
  setProp_('schedule', sched);
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'weeklySend').forEach(t => ScriptApp.deleteTrigger(t));
  if (sched.enabled) ScriptApp.newTrigger('weeklySend').timeBased().onWeekDay(ScriptApp.WeekDay[sched.day]).atHour(sched.hour).inTimezone(TZ).create();
  return sched;
}

// Weekly trigger: sends the prepared ("klaar om te versturen") newsletter, if there is one.
function weeklySend() {
  const cur = openIssue_();
  if (!cur || cur.status !== 'scheduled') { setProp_('lastRun', { at: new Date().toISOString(), result: 'skipped' }); return; }
  startSend_(cur.id);
}

function startSend_(id) {
  const sh = sheetOf_(NL_SHEET, NL_HEADERS);
  const iss = issues_().filter(i => i.id === id)[0];
  if (!iss) return;
  sh.getRange(iss.row, 2).setValue('sending');
  setProp_('sendState', { id: id, done: [] });
  continueSend();
}

// Sends in batches within Gmail's daily limit; continues automatically (next minute or next day).
function continueSend() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'continueSend').forEach(t => ScriptApp.deleteTrigger(t));
  const st = getProp_('sendState', null); if (!st) return;
  const iss = issues_().filter(i => i.id === st.id)[0];
  if (!iss) { props_().deleteProperty('sendState'); return; }
  const done = {}; st.done.forEach(e => done[e] = 1);
  const list = subscribers_().filter(s => s.status === 'active' && !done[s.email.toLowerCase()]);
  const t0 = Date.now(); let more = false;
  for (let i = 0; i < list.length; i++) {
    if (MailApp.getRemainingDailyQuota() <= 5) { more = 'quota'; break; }
    if (Date.now() - t0 > 4.5 * 60 * 1000) { more = 'time'; break; }
    const s = list[i]; const lang = s.lang === 'en' && iss.html.en ? 'en' : 'nl';
    const html = String(iss.html[lang]).split('{{UNSUBSCRIBE_URL}}').join(link_('unsub', s.email));
    try { MailApp.sendEmail({ to: s.email, subject: iss.subject[lang] || iss.subject.nl, htmlBody: html, name: 'Vervat & de Jong', replyTo: NOTIFY_EMAIL || undefined }); }
    catch (err) { console.error('send failed', s.email, err); }
    st.done.push(s.email.toLowerCase());
    if (i % 20 === 0) setProp_('sendState', st);
  }
  setProp_('sendState', st);
  const sh = sheetOf_(NL_SHEET, NL_HEADERS);
  if (more) {
    sh.getRange(iss.row, 10).setValue(st.done.length);
    ScriptApp.newTrigger('continueSend').timeBased().after(more === 'quota' ? 24 * 3600 * 1000 : 60 * 1000).create();
    return;
  }
  sh.getRange(iss.row, 2, 1, 1).setValue('sent');
  sh.getRange(iss.row, 9, 1, 2).setValues([[new Date(), st.done.length]]);
  props_().deleteProperty('sendState');
  setProp_('lastRun', { at: new Date().toISOString(), result: 'sent', recipients: st.done.length });
}

function nextRun_(sched) {
  if (!sched || !sched.enabled) return '';
  const now = new Date();
  const local = new Date(Utilities.formatDate(now, TZ, "yyyy-MM-dd'T'HH:mm:ss"));
  const want = (WEEKDAYS.indexOf(sched.day) + 1) % 7; // JS: 0 = Sunday
  const d = new Date(local); d.setHours(sched.hour, 0, 0, 0);
  let add = (want - local.getDay() + 7) % 7; if (add === 0 && d <= local) add = 7;
  d.setDate(d.getDate() + add);
  return Utilities.formatDate(d, 'UTC', "yyyy-MM-dd'T'HH:mm:ss") + '|' + TZ;
}

// Admin panel (POST with key) — newsletter actions.
function adminAction_(d) {
  const a = d.action;
  if (a === 'nl.state') {
    const subs = subscribers_(); const cur = openIssue_(); const sched = getProp_('schedule', { enabled: false, day: 'MONDAY', hour: 9 });
    return { ok: true, schedule: sched, nextRun: nextRun_(sched), lastRun: getProp_('lastRun', null), sending: getProp_('sendState', null) ? getProp_('sendState', null).done.length : null,
      issue: cur ? { id: cur.id, status: cur.status, source: cur.source } : null,
      history: issues_().filter(i => i.status === 'sent').reverse().slice(0, 30).map(i => ({ id: i.id, subject: i.subject, sentAt: i.sentAt, recipients: i.recipients, source: i.source })),
      counts: { active: subs.filter(s => s.status === 'active').length, pending: subs.filter(s => s.status === 'pending').length, unsubscribed: subs.filter(s => s.status === 'unsubscribed').length },
      quota: MailApp.getRemainingDailyQuota() };
  }
  if (a === 'nl.save') { const id = saveIssue_(d, d.ready ? 'scheduled' : 'draft'); if (d.schedule) setSchedule_(d.schedule); return { ok: true, id: id }; }
  if (a === 'nl.schedule') return { ok: true, schedule: setSchedule_(d.schedule) };
  if (a === 'nl.test') {
    if (!isEmail_(d.to)) return { ok: false, error: 'Ongeldig e-mailadres' };
    const lang = d.lang === 'en' ? 'en' : 'nl';
    MailApp.sendEmail({ to: d.to, subject: '[TEST] ' + (d.subject[lang] || ''), htmlBody: String(d.html[lang]).split('{{UNSUBSCRIBE_URL}}').join(link_('unsub', d.to)), name: 'Vervat & de Jong' });
    return { ok: true };
  }
  if (a === 'nl.sendNow') { const id = saveIssue_(d, 'scheduled'); startSend_(id); return { ok: true, id: id }; }
  if (a === 'nl.cancel') { const cur = openIssue_(); if (cur && cur.status === 'scheduled') sheetOf_(NL_SHEET, NL_HEADERS).getRange(cur.row, 2).setValue('draft'); return { ok: true }; }
  if (a === 'subs.list') return { ok: true, items: subscribers_() };
  if (a === 'subs.add') {
    if (!isEmail_(d.email)) return { ok: false, error: 'Ongeldig e-mailadres' };
    const sh = sheetOf_(SUBS_SHEET, SUBS_HEADERS); const f = findSub_(sh, d.email); const lang = d.lang === 'en' ? 'en' : 'nl';
    if (d.confirm) { if (f) sh.getRange(f.row, 2, 1, 5).setValues([[lang, 'active', f.v[3] || new Date(), new Date(), '']]); else sh.appendRow([String(d.email).trim(), lang, 'active', new Date(), new Date(), '', 'admin']); }
    else subscribe_(d.email, lang, 'admin');
    return { ok: true };
  }
  if (a === 'subs.set' || a === 'subs.delete') {
    const sh = sheetOf_(SUBS_SHEET, SUBS_HEADERS); const f = findSub_(sh, d.email); if (!f) return { ok: false, error: 'not found' };
    if (a === 'subs.delete') sh.deleteRow(f.row);
    else sh.getRange(f.row, 3).setValue(['active', 'pending', 'unsubscribed'].indexOf(d.status) >= 0 ? d.status : f.v[2]);
    return { ok: true };
  }
  return { ok: false, error: 'unknown action' };
}

// Run once from the editor to check that a test newsletter can be sent to NOTIFY_EMAIL.
function testNewsletter() {
  const html = '<p style="font-family:Arial">Testnieuwsbrief van Vervat &amp; de Jong. <a href="{{UNSUBSCRIBE_URL}}">Afmelden</a></p>';
  adminAction_({ action: 'nl.test', to: NOTIFY_EMAIL, lang: 'nl', subject: { nl: 'Nieuwsbrief test' }, html: { nl: html } });
}
