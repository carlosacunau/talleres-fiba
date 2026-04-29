/**
 * Fiba Labs · Talleres · Form intake endpoint (with audio uploads)
 *
 * Two actions handled in doPost:
 *   1. action: "upload-audio"  → save audio blob to Drive, keyed by submission_id + filename
 *   2. action: "submit-form"   → write text answers to Sheet, resolving audio filenames to Drive URLs
 *
 * Setup:
 *   1. Open the target Spreadsheet (named "Talleres - Aplicaciones").
 *   2. Extensions > Apps Script. Paste this file (replaces the previous version).
 *   3. Set SPREADSHEET_ID below to the Sheet's ID.
 *   4. Set ROOT_DRIVE_FOLDER_ID below to a Drive folder where audios should land
 *      (or leave as '' to auto-create one named "Talleres - Audios" in My Drive).
 *   5. Set NOTIFY_EMAIL.
 *   6. Deploy > New deployment > Web app:
 *      - Execute as: Me
 *      - Who has access: Anyone
 *      Use "Manage deployments" to create a NEW version (don't edit the existing one,
 *      or copy the new /exec URL into chile-1/index.html if it changes).
 *
 * The form posts JSON via no-cors (browsers can't read the response).
 * For audio uploads we use base64 over JSON to keep the contract simple.
 */

const SPREADSHEET_ID = 'REPLACE_WITH_SPREADSHEET_ID';
const ROOT_DRIVE_FOLDER_ID = ''; // empty → auto-create "Talleres - Audios" folder in root
const NOTIFY_EMAIL = 'carlos@fibalabs.com';

// ---------- ROUTING ----------

function doPost(e) {
  dlog('doPost ENTRY · contentLength=' + (e.postData ? e.postData.contents.length : 'no-postData'));
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || 'submit-form';
    dlog('doPost ROUTING · action=' + action + ' · cohort=' + payload.cohort + ' · submission_id=' + payload.submission_id);

    if (action === 'upload-audio') return handleUploadAudio(payload);
    if (action === 'submit-form')  return handleSubmitForm(payload);
    if (action === 'delete-audio') return handleDeleteAudio(payload);

    dlog('doPost UNKNOWN_ACTION · ' + action);
    return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
  } catch (err) {
    dlog('doPost ERROR · ' + err + ' · stack=' + (err.stack || 'n/a'));
    return jsonResponse({ status: 'error', message: String(err) });
  }
}

function doGet(e) {
  // Debug endpoint: ?debug=logs returns the last 50 log lines we cached.
  // ?debug=clear wipes the cache.
  const params = (e && e.parameter) || {};
  if (params.debug === 'logs') {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty('DEBUG_LOG') || '[]';
    let logs;
    try { logs = JSON.parse(raw); } catch (e) { logs = []; }
    return ContentService
      .createTextOutput(logs.join('\n') || '(empty)')
      .setMimeType(ContentService.MimeType.TEXT);
  }
  if (params.debug === 'clear') {
    PropertiesService.getScriptProperties().deleteProperty('DEBUG_LOG');
    return jsonResponse({ status: 'ok', cleared: true });
  }
  return jsonResponse({ status: 'ok', service: 'talleres-fiba intake (v2 with audio)' });
}

/**
 * Persistent log: appends to ScriptProperties so we can read it via ?debug=logs.
 * Logger.log only goes to the in-editor view which is unreliable for some accounts.
 */
function dlog(msg) {
  Logger.log(msg);
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty('DEBUG_LOG') || '[]';
    let logs;
    try { logs = JSON.parse(raw); } catch (e) { logs = []; }
    const stamp = new Date().toISOString();
    logs.push(`[${stamp}] ${msg}`);
    if (logs.length > 50) logs.splice(0, logs.length - 50);
    props.setProperty('DEBUG_LOG', JSON.stringify(logs));
  } catch (err) {
    Logger.log('dlog failed: ' + err);
  }
}

// ---------- HANDLERS ----------

/**
 * Save an audio blob to Drive.
 * Drive layout: <root>/<cohort>/<submission_id>/<filename>
 */
function handleUploadAudio(payload) {
  const cohort = payload.cohort || 'unknown';
  const submissionId = payload.submission_id || 'anon';
  const filename = payload.filename || ('audio_' + Date.now() + '.webm');
  const mime = payload.mime || 'audio/webm';
  const base64 = payload.data_base64;

  if (!base64) return jsonResponse({ status: 'error', message: 'Missing data_base64' });

  const submissionFolder = getOrCreateSubmissionFolder_(cohort, submissionId);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mime, filename);
  const file = submissionFolder.createFile(blob);
  // Make it accessible by anyone with the link (so Carlos can play from Sheet)
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    dlog('setSharing failed: ' + e);
  }

  return jsonResponse({
    status: 'ok',
    file_id: file.getId(),
    url: file.getUrl(),
    filename,
  });
}

/**
 * Delete a previously uploaded audio file by filename within the submission folder.
 */
function handleDeleteAudio(payload) {
  const cohort = payload.cohort || 'unknown';
  const submissionId = payload.submission_id || '';
  const filename = payload.filename || '';
  if (!submissionId || !filename) {
    return jsonResponse({ status: 'error', message: 'Missing submission_id or filename' });
  }
  const folder = getOrCreateSubmissionFolder_(cohort, submissionId);
  const it = folder.getFilesByName(filename);
  let deleted = 0;
  while (it.hasNext()) {
    const f = it.next();
    f.setTrashed(true);
    deleted++;
  }
  dlog('handleDeleteAudio · cohort=' + cohort + ' · sub=' + submissionId + ' · file=' + filename + ' · deleted=' + deleted);
  return jsonResponse({ status: 'ok', deleted });
}

/**
 * Append a row to the cohort's sheet tab.
 * Resolves <qid>__audio_filenames into a comma-separated list of Drive URLs by looking
 * up files we just saved in the submission folder.
 */
function handleSubmitForm(payload) {
  dlog('handleSubmitForm START');
  // Coerce cohort to a single string (defensive: if duplicated in payload it can come as "chile-1,chile-1")
  let cohort = payload.cohort || 'unknown';
  if (Array.isArray(cohort)) cohort = cohort[0];
  if (typeof cohort === 'string' && cohort.indexOf(',') !== -1) cohort = cohort.split(',')[0];
  const submissionId = payload.submission_id || '';
  dlog('handleSubmitForm OPENING_SHEET · ' + SPREADSHEET_ID);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  dlog('handleSubmitForm SHEET_OPEN · name=' + ss.getName());

  let sheet = ss.getSheetByName(cohort);
  const isNewSheet = !sheet;
  dlog('handleSubmitForm TAB_LOOKUP · cohort=' + cohort + ' · existed=' + !isNewSheet);
  if (isNewSheet) sheet = ss.insertSheet(cohort);

  const knownOrder = [
    'submitted_at', 'cohort', 'submission_id',
    'contact_name', 'contact_email',
    'q1_empresa', 'q1_web', 'q1_industria',
    'q2_equipo', 'q3_cliente_modelo',
    'q5_herramientas', 'q5b_fuentes_datos', 'q6_info_critica', 'q7_correo', 'q8_computador',
    'q9_trigger', 'q9_inputs', 'q9_proceso', 'q9_output', 'q9_tiempo',
    'q10_trigger', 'q10_inputs', 'q10_proceso', 'q10_output', 'q10_tiempo',
    'q11_trigger', 'q11_inputs', 'q11_proceso', 'q11_output', 'q11_tiempo',
    'q12_reportes', 'q13_dashboard',
    'q14_conceptos', 'q14_dudas', 'q15_extra',
    // Audio Drive URL columns (parallel to long-form questions)
    'q3_cliente_modelo_audio', 'q5_herramientas_audio', 'q5b_fuentes_datos_audio',
    'q6_info_critica_audio', 'q9_proceso_audio', 'q10_proceso_audio',
    'q11_proceso_audio', 'q12_reportes_audio', 'q13_dashboard_audio', 'q15_extra_audio',
  ];

  // Ensure header row matches the current schema. Replace it if missing or stale.
  const lastCol = sheet.getLastColumn();
  let needsHeader = isNewSheet || sheet.getLastRow() === 0;
  if (!needsHeader) {
    const currentHeader = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
    const matches = knownOrder.every((k, i) => currentHeader[i] === k);
    if (!matches) {
      // Insert a fresh header row above existing data
      sheet.insertRowBefore(1);
      needsHeader = true;
    }
  }
  if (needsHeader) {
    sheet.getRange(1, 1, 1, knownOrder.length).setValues([knownOrder]);
    sheet.getRange(1, 1, 1, knownOrder.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Build the row, resolving audio filenames to Drive URLs
  const submissionFolder = submissionId
    ? getOrCreateSubmissionFolder_(cohort, submissionId)
    : null;

  const filesByName = {};
  if (submissionFolder) {
    const it = submissionFolder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      filesByName[f.getName()] = f.getUrl();
    }
  }

  const row = knownOrder.map(key => {
    if (key.endsWith('_audio')) {
      const baseQid = key.replace(/_audio$/, '');
      const filenamesField = payload[`${baseQid}__audio_filenames`];
      if (!filenamesField) return '';
      return filenamesField.split('|')
        .map(fn => filesByName[fn] || `(missing: ${fn})`)
        .join(' , ');
    }
    const val = payload[key];
    if (val === undefined || val === null) return '';
    if (Array.isArray(val)) return val.join(', ');
    return String(val);
  });
  dlog('handleSubmitForm APPENDING_ROW · cells=' + row.length + ' · firstFew=' + JSON.stringify(row.slice(0, 5)));
  sheet.appendRow(row);
  dlog('handleSubmitForm ROW_APPENDED · lastRow now=' + sheet.getLastRow());

  // Email notification
  try {
    const audioCount = Object.keys(filesByName).length;
    const subject = `Nueva aplicación · ${cohort} · ${payload.contact_name || 'sin nombre'}`;
    const body = [
      `Cohort: ${cohort}`,
      `Submission: ${submissionId}`,
      `Nombre: ${payload.contact_name || ''}`,
      `Correo: ${payload.contact_email || ''}`,
      `Empresa: ${payload.q1_empresa || ''} (${payload.q1_industria || ''})`,
      `Equipo: ${payload.q2_equipo || ''}`,
      `Audios adjuntos: ${audioCount}`,
      '',
      `Sheet: https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit`,
      submissionFolder ? `Drive: ${submissionFolder.getUrl()}` : '',
    ].filter(Boolean).join('\n');
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  } catch (mailErr) {
    dlog('Mail failed: ' + mailErr);
  }

  return jsonResponse({ status: 'ok' });
}

// ---------- DRIVE HELPERS ----------

function getOrCreateRootFolder_() {
  if (ROOT_DRIVE_FOLDER_ID) return DriveApp.getFolderById(ROOT_DRIVE_FOLDER_ID);
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty('AUTO_ROOT_FOLDER_ID');
  if (cached) {
    try { return DriveApp.getFolderById(cached); } catch (e) { /* fall through */ }
  }
  const folder = DriveApp.createFolder('Talleres - Audios');
  props.setProperty('AUTO_ROOT_FOLDER_ID', folder.getId());
  return folder;
}

function getOrCreateChild_(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

function getOrCreateSubmissionFolder_(cohort, submissionId) {
  // LockService prevents the parallel-upload race that would otherwise create
  // multiple folders with the same name (Drive doesn't enforce uniqueness).
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // up to 20s
  try {
    const root = getOrCreateRootFolder_();
    const cohortFolder = getOrCreateChild_(root, cohort);
    return getOrCreateChild_(cohortFolder, submissionId);
  } finally {
    lock.releaseLock();
  }
}

// ---------- UTIL ----------

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
