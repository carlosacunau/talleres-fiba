/**
 * Fiba Labs · Talleres · Form intake endpoint
 *
 * Receives JSON applications from talleres.fibalabs.com forms,
 * appends them to a Google Sheet (one tab per cohort), and emails Carlos.
 *
 * Setup:
 *   1. Open the target Spreadsheet (create one named "Talleres - Aplicaciones").
 *   2. Extensions > Apps Script. Paste this file.
 *   3. Replace SPREADSHEET_ID below with the sheet's ID (from its URL).
 *   4. Replace NOTIFY_EMAIL with where you want notifications.
 *   5. Deploy > New deployment > type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 *   6. Copy the /exec URL and paste into chile-1/index.html (ENDPOINT const).
 *
 * Adding future cohorts: no code changes — the script auto-creates a tab
 * per `cohort` value sent by the form (e.g. "chile-1", "chile-2", "colombia-1").
 */

const SPREADSHEET_ID = 'REPLACE_WITH_SPREADSHEET_ID';
const NOTIFY_EMAIL = 'carlos@fibalabs.com';

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const cohort = payload.cohort || 'unknown';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    let sheet = ss.getSheetByName(cohort);
    const isNewSheet = !sheet;
    if (isNewSheet) {
      sheet = ss.insertSheet(cohort);
    }

    // Build column list from payload keys (stable order: known keys first, then any extras)
    const knownOrder = [
      'submitted_at', 'cohort', 'contact_name', 'contact_email',
      'q1_empresa', 'q1_web', 'q1_industria',
      'q2_equipo', 'q3_cliente_modelo',
      'q4_canales', 'q4_otro',
      'q5_herramientas', 'q6_info_critica', 'q7_correo', 'q8_computador',
      'q9_trigger', 'q9_inputs', 'q9_proceso', 'q9_output', 'q9_tiempo',
      'q10_trigger', 'q10_inputs', 'q10_proceso', 'q10_output', 'q10_tiempo',
      'q11_trigger', 'q11_inputs', 'q11_proceso', 'q11_output', 'q11_tiempo',
      'q12_reportes', 'q13_dashboard',
      'q14_conceptos', 'q14_dudas', 'q15_extra',
    ];

    if (isNewSheet || sheet.getLastRow() === 0) {
      sheet.appendRow(knownOrder);
      sheet.getRange(1, 1, 1, knownOrder.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const row = knownOrder.map(key => {
      const val = payload[key];
      if (val === undefined || val === null) return '';
      if (Array.isArray(val)) return val.join(', ');
      return String(val);
    });
    sheet.appendRow(row);

    // Email notification
    const subject = `Nueva aplicación · ${cohort} · ${payload.contact_name || 'sin nombre'}`;
    const body = [
      `Cohort: ${cohort}`,
      `Nombre: ${payload.contact_name || ''}`,
      `Correo: ${payload.contact_email || ''}`,
      `Empresa: ${payload.q1_empresa || ''} (${payload.q1_industria || ''})`,
      `Equipo: ${payload.q2_equipo || ''}`,
      '',
      `Sheet: https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit`,
    ].join('\n');

    try {
      MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
    } catch (mailErr) {
      // Don't fail the submission if mail quota is exhausted
      Logger.log('Mail failed: ' + mailErr);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log(err);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', service: 'talleres-fiba intake' }))
    .setMimeType(ContentService.MimeType.JSON);
}
