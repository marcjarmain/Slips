/**
 * Slips: shoot a slip, add a note, save to Drive.
 * Google Apps Script web app. No model, no server.
 *
 * Drive layout it creates on first use:
 *   Slips/
 *     Slip Log            (Google Sheet, one row per slip)
 *     2026/10/2026-10-25_mountainview_141502.jpg
 */

const CONFIG = {
  FOLDER: 'Slips',
  SHEET: 'Slip Log',
  TZ: 'Africa/Johannesburg',
  // Quick "what is it for" buttons. Edit freely.
  TAGS: [
    'Rondebosch Vista',
    'Mountainview Villa',
  ],
};

const HEADERS = [
  'Logged', 'Slip date', 'Amount', 'Type', 'For',
  'Note', 'Photos', 'Photo IDs', 'Status', 'Slip ID',
];

/* ---------- web endpoints for the hosted page ---------- */

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function keyOk_(k) {
  const want = props_().getProperty('KEY');
  return !!want && String(k || '') === want;
}

/** Run once from the editor: sets the app key. */
function setupKey() {
  const key = Utilities.getUuid().replace(/-/g, '')
    .slice(0, 12);
  props_().setProperty('KEY', key);
  Logger.log('Your Slips key: ' + key);
  return key;
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (!keyOk_(body.key)) {
      return json_({ ok: false, error: 'Wrong key' });
    }
    if (body.action === 'recent') {
      return json_({ ok: true, rows: getRecent() });
    }
    if (body.action === 'folder') {
      return json_({ ok: true, url: getFolderUrl() });
    }
    return json_(saveSlip(body.slip || {}));
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.tagsJson = JSON.stringify(CONFIG.TAGS);
  return t.evaluate()
    .setTitle('Slips')
    .addMetaTag('viewport',
      'width=device-width, initial-scale=1, ' +
      'maximum-scale=1, viewport-fit=cover');
}

/* ---------- storage helpers ---------- */

function props_() {
  return PropertiesService.getScriptProperties();
}

function root_() {
  const p = props_();
  const id = p.getProperty('ROOT');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) {}
  }
  const it = DriveApp.getFoldersByName(CONFIG.FOLDER);
  const f = it.hasNext()
    ? it.next()
    : DriveApp.createFolder(CONFIG.FOLDER);
  p.setProperty('ROOT', f.getId());
  return f;
}

function sub_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function sheet_() {
  const p = props_();
  const id = p.getProperty('SHEET');
  let ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) {}
  }
  if (!ss) {
    ss = SpreadsheetApp.create(CONFIG.SHEET);
    DriveApp.getFileById(ss.getId()).moveTo(root_());
    const sh = ss.getSheets()[0];
    sh.setName('Slips');
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold');
    p.setProperty('SHEET', ss.getId());
  }
  return ss.getSheetByName('Slips') || ss.getSheets()[0];
}

/* ---------- called from the phone ---------- */

/**
 * p = {
 *   date: 'yyyy-MM-dd', amount: '123,45',
 *   type: 'Business' | 'Personal' | 'Split',
 *   tag: 'Mountainview', note: 'text',
 *   images: [base64 jpeg, ...]
 * }
 */
function saveSlip(p) {
  p = p || {};
  const images = p.images || [];
  const note = String(p.note || '').trim();
  if (!images.length && !note) {
    throw new Error('Add a photo or a note.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const now = new Date();
    const date = /^\d{4}-\d{2}-\d{2}$/.test(p.date || '')
      ? p.date
      : Utilities.formatDate(now, CONFIG.TZ, 'yyyy-MM-dd');
    const parts = date.split('-');
    const folder = sub_(sub_(root_(), parts[0]), parts[1]);

    const slipId =
      Utilities.formatDate(now, CONFIG.TZ, 'yyyyMMdd-HHmmss');
    const slug = String(p.tag || p.type || 'slip')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-|-$/g, '') || 'slip';

    const desc = [
      p.type, p.tag, p.amount, note,
    ].filter(Boolean).join(' | ');

    const urls = [];
    const ids = [];
    images.forEach(function (b64, i) {
      const suffix = images.length > 1 ? '_' + (i + 1) : '';
      const name = date + '_' + slug + '_' +
        slipId.slice(-6) + suffix + '.jpg';
      const blob = Utilities.newBlob(
        Utilities.base64Decode(b64), 'image/jpeg', name);
      const file = folder.createFile(blob);
      file.setDescription(desc);
      urls.push(file.getUrl());
      ids.push(file.getId());
    });

    const raw = String(p.amount || '')
      .replace(/\s/g, '').replace(',', '.');
    const amount = raw && !isNaN(Number(raw))
      ? Number(raw) : '';

    sheet_().appendRow([
      Utilities.formatDate(now, CONFIG.TZ,
        'yyyy-MM-dd HH:mm'),
      date,
      amount,
      p.type || '',
      p.tag || '',
      note,
      urls.join('\n'),
      ids.join(','),
      'new',
      slipId,
    ]);

    return { ok: true, id: slipId, photos: ids.length };
  } finally {
    lock.releaseLock();
  }
}

/** Last few slips for the list on the phone. */
function getRecent() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(8, last - 1);
  const rows = sh
    .getRange(last - n + 1, 1, n, HEADERS.length)
    .getDisplayValues();
  return rows.reverse().map(function (r) {
    return {
      date: r[1], amount: r[2], type: r[3], tag: r[4],
      note: r[5], url: (r[6] || '').split('\n')[0],
      photos: r[7] ? r[7].split(',').length : 0,
    };
  });
}

/** Link to the Drive folder, shown in the app. */
function getFolderUrl() {
  sheet_();
  return root_().getUrl();
}
