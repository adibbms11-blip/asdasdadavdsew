/**
 * GALLERY INHISTORIA — Google Apps Script backend
 * Setup:
 * 1) Create a standalone Apps Script project and paste this file as Code.gs.
 * 2) Add Index.html (included in the ZIP).
 * 3) Run initializeGallery() once and authorize.
 * 4) Copy the ADMIN KEY from the execution log. Keep it private.
 * 5) Deploy as Web app: Execute as Me; Who has access: Anyone.
 */

const CONFIG = {
  PROJECTS_SHEET: 'Projects',
  SELECTIONS_SHEET: 'Selections',
  SETTINGS_SHEET: 'Settings',
  APP_TITLE: 'Gallery InHistoria',
  MAX_PHOTOS_PER_FOLDER: 800
};

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.projectToken = (e && e.parameter && e.parameter.p) ? String(e.parameter.p) : '';
  template.appTitle = CONFIG.APP_TITLE;
  return template.evaluate()
    .setTitle(CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Run once from the Apps Script editor. Creates a spreadsheet database
 * and a random admin key. Copy the key from Execution log.
 */
function initializeGallery() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty('GALLERY_DB_ID');
  let ss;
  if (spreadsheetId) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } else {
    ss = SpreadsheetApp.create('Gallery InHistoria — Database');
    props.setProperty('GALLERY_DB_ID', ss.getId());
  }

  ensureSheet_(ss, CONFIG.PROJECTS_SHEET, [
    'token', 'clientName', 'folderId', 'folderUrl', 'limit', 'createdAt',
    'status', 'submittedAt', 'selectedCodesJson', 'photoCount'
  ]);
  ensureSheet_(ss, CONFIG.SELECTIONS_SHEET, [
    'token', 'clientName', 'submittedAt', 'selectedCodesJson', 'selectedCount'
  ]);
  ensureSheet_(ss, CONFIG.SETTINGS_SHEET, ['key', 'value']);

  let adminKey = props.getProperty('GALLERY_ADMIN_KEY');
  if (!adminKey) {
    adminKey = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty('GALLERY_ADMIN_KEY', adminKey);
  }
  Logger.log('DATABASE URL: ' + ss.getUrl());
  Logger.log('ADMIN KEY (keep private): ' + adminKey);
  Logger.log('Setup complete. Deploy this project as a Web app after testing.');
  return { databaseUrl: ss.getUrl(), adminKey: adminKey };
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function getAdminKey_() {
  return PropertiesService.getScriptProperties().getProperty('GALLERY_ADMIN_KEY') || '';
}
function getDb_() {
  const id = PropertiesService.getScriptProperties().getProperty('GALLERY_DB_ID');
  if (!id) throw new Error('Database belum dibuat. Jalankan initializeGallery() dari editor Apps Script.');
  return SpreadsheetApp.openById(id);
}
function requireAdmin_(key) {
  if (!key || String(key) !== getAdminKey_()) throw new Error('Kunci admin salah. Periksa ADMIN KEY dari proses setup.');
}
function getSheetRows_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}
function findProject_(token) {
  const sheet = getDb_().getSheetByName(CONFIG.PROJECTS_SHEET);
  const rows = getSheetRows_(sheet);
  return rows.find(r => String(r.token) === String(token)) || null;
}
function findProjectRow_(token) {
  const sheet = getDb_().getSheetByName(CONFIG.PROJECTS_SHEET);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(token)) return i + 1;
  }
  return -1;
}
function parseFolderId_(value) {
  const s = String(value || '').trim();
  const match = s.match(/[-\w]{25,}/);
  if (!match) throw new Error('Link atau ID folder Google Drive tidak valid.');
  return match[0];
}
function safeJson_(value) {
  try { return JSON.parse(value || '[]'); } catch (err) { return []; }
}
function getWebAppUrl_() {
  return ScriptApp.getService().getUrl() || '';
}

/** Admin dashboard initial data */
function adminGetData(adminKey) {
  requireAdmin_(adminKey);
  const rows = getSheetRows_(getDb_().getSheetByName(CONFIG.PROJECTS_SHEET));
  return {
    appUrl: getWebAppUrl_(),
    projects: rows.map(r => ({
      token: String(r.token),
      clientName: String(r.clientName),
      folderUrl: String(r.folderUrl),
      limit: Number(r.limit) || 0,
      createdAt: r.createdAt ? new Date(r.createdAt).toISOString() : '',
      status: String(r.status || 'Belum dikirim'),
      submittedAt: r.submittedAt ? new Date(r.submittedAt).toISOString() : '',
      photoCount: Number(r.photoCount) || 0,
      selectedCodes: safeJson_(r.selectedCodesJson)
    })).reverse()
  };
}

/** Create a new client gallery. Folder must be accessible to the script owner. */
function adminCreateProject(payload, adminKey) {
  requireAdmin_(adminKey);
  payload = payload || {};
  const clientName = String(payload.clientName || '').trim();
  const folderUrl = String(payload.folderUrl || '').trim();
  const limit = Number(payload.limit);
  if (!clientName) throw new Error('Nama klien wajib diisi.');
  if (!folderUrl) throw new Error('Link folder Google Drive wajib diisi.');
  if (!Number.isInteger(limit) || limit < 1 || limit > 500) throw new Error('Batas pilihan harus angka bulat antara 1–500.');

  const folderId = parseFolderId_(folderUrl);
  let folder;
  try { folder = DriveApp.getFolderById(folderId); }
  catch (err) { throw new Error('Folder tidak dapat diakses. Pastikan folder benar dan akun pemilik script memiliki akses.'); }

  const photos = listPhotosInFolder_(folder);
  if (!photos.length) throw new Error('Folder ini belum berisi foto yang didukung (JPG, PNG, GIF, WebP).');
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const sheet = getDb_().getSheetByName(CONFIG.PROJECTS_SHEET);
  sheet.appendRow([token, clientName, folderId, folderUrl, limit, new Date(), 'Belum dikirim', '', '[]', photos.length]);
  return {
    token: token,
    clientName: clientName,
    limit: limit,
    photoCount: photos.length,
    url: getWebAppUrl_() + '?p=' + encodeURIComponent(token)
  };
}

/** List direct child image files; names are the photo codes (without extension). */
function listPhotosInFolder_(folder) {
  const supported = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];
  const files = folder.getFiles();
  const photos = [];
  while (files.hasNext() && photos.length < CONFIG.MAX_PHOTOS_PER_FOLDER) {
    const file = files.next();
    if (supported.indexOf(file.getMimeType()) === -1) continue;
    const name = file.getName();
    const code = name.replace(/\.[^.]+$/, '');
    photos.push({
      id: file.getId(),
      code: code,
      name: name,
      mimeType: file.getMimeType(),
      // Drive thumbnail endpoint; client must have viewing access to the image.
      thumbUrl: 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(file.getId()) + '&sz=w700',
      previewUrl: 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(file.getId()) + '&sz=w1800'
    });
  }
  photos.sort((a, b) => a.name.localeCompare(b.name, undefined, {numeric: true, sensitivity: 'base'}));
  return photos;
}

/** Client gallery data. Token is a long, unguessable project link. */
function clientGetGallery(token) {
  token = String(token || '').trim();
  if (!token || token.length < 40) throw new Error('Link galeri tidak valid.');
  const project = findProject_(token);
  if (!project) throw new Error('Proyek tidak ditemukan atau link sudah tidak berlaku.');
  const folder = DriveApp.getFolderById(String(project.folderId));
  const photos = listPhotosInFolder_(folder);
  return {
    clientName: String(project.clientName),
    limit: Number(project.limit) || 0,
    photos: photos,
    status: String(project.status || 'Belum dikirim'),
    submittedAt: project.submittedAt ? new Date(project.submittedAt).toISOString() : '',
    selectedCodes: safeJson_(project.selectedCodesJson)
  };
}

/** Server-side validates the token, file codes, quota and submission status. */
function clientSubmitSelection(token, selectedCodes) {
  token = String(token || '').trim();
  if (!token || token.length < 40) throw new Error('Link galeri tidak valid.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const project = findProject_(token);
    if (!project) throw new Error('Proyek tidak ditemukan.');
    if (String(project.status) === 'Terkirim') throw new Error('Pilihan sudah dikirim. Hubungi admin jika perlu revisi.');
    if (!Array.isArray(selectedCodes)) throw new Error('Daftar pilihan tidak valid.');

    const limit = Number(project.limit) || 0;
    const folder = DriveApp.getFolderById(String(project.folderId));
    const photos = listPhotosInFolder_(folder);
    const allowed = {};
    photos.forEach(p => allowed[p.code] = true);
    const uniqueCodes = Array.from(new Set(selectedCodes.map(String)));
    if (uniqueCodes.length > limit) throw new Error('Pilihan melebihi batas maksimal ' + limit + ' foto.');
    for (const code of uniqueCodes) {
      if (!allowed[code]) throw new Error('Kode foto tidak ditemukan di folder: ' + code);
    }

    const submittedAt = new Date();
    const row = findProjectRow_(token);
    if (row < 2) throw new Error('Proyek tidak ditemukan.');
    const sheet = getDb_().getSheetByName(CONFIG.PROJECTS_SHEET);
    sheet.getRange(row, 7, 1, 4).setValues([['Terkirim', submittedAt, JSON.stringify(uniqueCodes), photos.length]]);
    getDb_().getSheetByName(CONFIG.SELECTIONS_SHEET)
      .appendRow([token, String(project.clientName), submittedAt, JSON.stringify(uniqueCodes), uniqueCodes.length]);
    return { ok: true, submittedAt: submittedAt.toISOString(), selectedCodes: uniqueCodes, count: uniqueCodes.length };
  } finally {
    lock.releaseLock();
  }
}

/** Admin may reopen a submitted project for changes. */
function adminReopenProject(token, adminKey) {
  requireAdmin_(adminKey);
  const row = findProjectRow_(String(token));
  if (row < 2) throw new Error('Proyek tidak ditemukan.');
  getDb_().getSheetByName(CONFIG.PROJECTS_SHEET).getRange(row, 7, 1, 3)
    .setValues([['Belum dikirim', '', '[]']]);
  return { ok: true };
}
