// ══════════════════════════════════════════════════════════════════
//  مستشفيات الحمادي — بوابة مستندات الأطباء
//  Google Apps Script — Code.gs
//  Doctor Documents Portal Backend
// ══════════════════════════════════════════════════════════════════

// ─── ⚙️  إعدادات — عدّل هذا القسم فقط ─────────────────────────────
const SHEET_ID    = 'YOUR_GOOGLE_SHEET_ID_HERE';
const FOLDER_ID   = 'YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE';
const TIMEZONE    = 'Asia/Riyadh';
// ─────────────────────────────────────────────────────────────────

// Sheet names
const SH_DOCTORS   = 'Doctors';
const SH_UPLOADS   = 'Uploads';
const SH_LOGS      = 'Logs';
const SH_DASHBOARD = 'Dashboard_Summary';


// ══════════════════════════════════════════════════════════════════
//  doGet — للاختبار
// ══════════════════════════════════════════════════════════════════
function doGet(e) {
  return jsonResponse({ status: 'ok', message: 'Al Hammadi Doctor Portal API — Active ✓' });
}


// ══════════════════════════════════════════════════════════════════
//  doPost — نقطة الدخول الرئيسية
// ══════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const action = e.parameter.action || '';

    switch(action) {
      case 'submit':    return handleSubmit(e);
      case 'verifyAdmin': return handleAdminLogin(e);
      case 'log':       return handleLog(e);
      case 'dashboard': return handleDashboard(e);
      default:
        return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
    }
  } catch(err) {
    Logger.log('doPost error: ' + err.message + '\n' + err.stack);
    return jsonResponse({ status: 'error', message: err.message });
  }
}


// ══════════════════════════════════════════════════════════════════
//  ADMIN VERIFICATION
// ══════════════════════════════════════════════════════════════════
function handleAdminLogin(e) {
  const code = (e.parameter.code || '').trim();

  // Get admin code from Script Properties (never hardcoded in JS/HTML)
  const stored = PropertiesService.getScriptProperties().getProperty('ADMIN_CODE');

  if (!stored) {
    return jsonResponse({ status: 'error', message: 'Admin code not configured. Run setupProperties().' });
  }

  if (code !== stored) {
    logToSheet('ADMIN', 'FAILED_LOGIN', 'كود خاطئ');
    return jsonResponse({ status: 'ok', authorized: false });
  }

  // Return dashboard data
  const data = buildDashboardData();
  logToSheet('ADMIN', 'LOGIN', 'دخول ناجح');
  return jsonResponse({ status: 'ok', authorized: true, data });
}


// ══════════════════════════════════════════════════════════════════
//  SUBMIT HANDLER
// ══════════════════════════════════════════════════════════════════
function handleSubmit(e) {
  const params     = e.parameter;
  const doctorId   = (params.doctorId || '').trim();
  const doctorName = (params.doctorName || '').trim();
  const dept       = (params.dept || '').trim();
  const spec       = (params.spec || '').trim();
  const timestamp  = params.timestamp || new Date().toISOString();
  const rowsRaw    = params.data || '[]';
  const rows       = JSON.parse(rowsRaw);

  if (!doctorId) return jsonResponse({ status: 'error', message: 'Missing doctorId' });

  // ── 1. Get/create doctor folder in Drive ──────────────────────
  const folderName  = doctorId + ' — ' + doctorName;
  const doctorFolder = getOrCreateFolder(folderName);

  // ── 2. Save base64 files ─────────────────────────────────────
  const fileLinks = {};
  const fileNames = {};
  const fileDates = {};

  for (const key in params) {
    if (!key.startsWith('file_b64_')) continue;
    try {
      const docLabel  = decodeURIComponent(key.replace('file_b64_', ''));
      const parts     = params[key].split('||');
      if (parts.length < 3) continue;

      const origName  = parts[0];
      const mimeType  = parts[1];
      const b64Data   = parts[2];

      // Create subfolder per document
      const docSubFolder = getOrCreateFolder(docLabel, doctorFolder);

      // Add timestamp to filename (no delete of old files)
      const ts = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd_HHmmss');
      const ext = origName.split('.').pop();
      const baseName = origName.replace(/\.[^.]+$/, '');
      const newName = baseName + '_' + ts + '.' + ext;

      const blob = Utilities.newBlob(
        Utilities.base64Decode(b64Data), mimeType, newName
      );
      const savedFile = docSubFolder.createFile(blob);
      savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      fileLinks[docLabel] = savedFile.getUrl();
      fileNames[docLabel] = newName;
      fileDates[docLabel] = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    } catch(ferr) {
      Logger.log('File save error: ' + ferr.message);
    }
  }

  // ── 3. Write to Uploads sheet ─────────────────────────────────
  const uploadsSheet = getOrCreateSheet(SH_UPLOADS, [
    'Timestamp','Employee ID','Doctor Name','Department','Specialty',
    'Document Type','File Name','File URL','Note','Upload Status','Uploaded Date'
  ]);

  const now = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  rows.forEach(row => {
    uploadsSheet.appendRow([
      now,
      doctorId,
      doctorName,
      dept,
      spec,
      row.doc,
      fileNames[row.doc] || '—',
      fileLinks[row.doc] || '—',
      row.note || '—',
      fileLinks[row.doc] ? 'Uploaded' : (row.note ? 'Note Provided' : 'Missing'),
      fileDates[row.doc] || '—'
    ]);
  });

  // ── 4. Update Doctors sheet ───────────────────────────────────
  updateDoctorsSheet(doctorId, doctorName, dept, spec, rows, fileLinks);

  // ── 5. Log ───────────────────────────────────────────────────
  const submissionId = doctorId + '-' + Date.now().toString(36).toUpperCase();
  logToSheet(doctorId, 'SUBMIT', `رفع ${rows.filter(r=>r.hasFile).length} ملف و ${rows.filter(r=>!r.hasFile&&r.note).length} ملاحظة | ID: ${submissionId}`);

  // ── 6. Update Dashboard ───────────────────────────────────────
  try { updateDashboard(); } catch(e) { Logger.log('Dashboard update error: '+e.message); }

  return jsonResponse({ status: 'ok', submissionId, message: 'تم الحفظ بنجاح' });
}


// ══════════════════════════════════════════════════════════════════
//  DOCTORS SHEET UPDATE
// ══════════════════════════════════════════════════════════════════
function updateDoctorsSheet(doctorId, doctorName, dept, spec, rows, fileLinks) {
  const sheet = getOrCreateSheet(SH_DOCTORS, [
    'Employee ID','Doctor Name','Mobile','Department','Specialty',
    'National Address','Qualifications','CV','SCFHS Classification',
    'MOH License','National ID','IBAN','Malpractice Insurance',
    'Photo','Required Docs','Uploaded Count','Notes Count',
    'Completion %','Final Status','Last Updated'
  ]);

  const data   = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(doctorId)) {
      rowIndex = i + 1; break;
    }
  }

  const total    = rows.length;
  const uploaded = rows.filter(r => r.hasFile).length;
  const noted    = rows.filter(r => !r.hasFile && r.note && r.note.trim()).length;
  const done     = uploaded + noted;
  const pct      = total > 0 ? Math.round(done / total * 100) : 0;
  const status   = pct >= 100 ? 'Complete' : pct > 0 ? 'Partial' : 'Not Started';
  const now      = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 16).setValue(uploaded);
    sheet.getRange(rowIndex, 17).setValue(noted);
    sheet.getRange(rowIndex, 18).setValue(pct + '%');
    sheet.getRange(rowIndex, 19).setValue(status);
    sheet.getRange(rowIndex, 20).setValue(now);
  } else {
    sheet.appendRow([
      doctorId, doctorName, '—', dept, spec,
      '','','','','','','','','',
      total, uploaded, noted, pct + '%', status, now
    ]);
  }
}


// ══════════════════════════════════════════════════════════════════
//  DASHBOARD DATA
// ══════════════════════════════════════════════════════════════════
function buildDashboardData() {
  const ss     = SpreadsheetApp.openById(SHEET_ID);
  const sh     = ss.getSheetByName(SH_DOCTORS);
  const upSh   = ss.getSheetByName(SH_UPLOADS);

  if (!sh) return { total:0, complete:0, partial:0, notStarted:0, doctors:[] };

  const data   = sh.getDataRange().getValues();
  const headers= data[0];

  let total=0, complete=0, partial=0, notStarted=0;
  let totalFiles=0, totalNotes=0;
  const doctors = [];

  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    if (!row[0]) continue;
    total++;

    const id     = String(row[0]);
    const name   = row[1];
    const mobile = row[2];
    const dept   = row[3];
    const spec   = row[4];
    const reqd   = Number(row[14]) || 0;
    const upl    = Number(row[15]) || 0;
    const notes  = Number(row[16]) || 0;
    const pctStr = String(row[17] || '0%');
    const pct    = parseInt(pctStr) || 0;
    const status = row[18] || 'Not Started';
    const last   = row[19] || '—';

    if (status === 'Complete')    complete++;
    else if (status === 'Partial') partial++;
    else                           notStarted++;

    totalFiles += upl;
    totalNotes += notes;

    // Get document details from Uploads sheet
    const docs = getDocDetailsForDoctor(upSh, id);

    doctors.push({ id, name, mobile, dept, spec,
      required:reqd, uploaded:upl, notes, completionPct:pct,
      finalStatus:status, lastUpdated:last, documents:docs });
  }

  const avg = total > 0 ? Math.round(
    doctors.reduce((sum,d)=>sum+(d.completionPct||0),0) / total
  ) : 0;

  return { total, complete, partial, notStarted, avgCompletion:avg,
    totalFiles, totalNotes, doctors };
}

function getDocDetailsForDoctor(upSheet, doctorId) {
  if (!upSheet) return [];
  const data = upSheet.getDataRange().getValues();
  const result = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[1]) !== String(doctorId)) continue;
    const docName  = row[5];
    const fileName = row[6];
    const fileUrl  = row[7];
    const note     = row[8];
    const date     = row[10];

    if (!result[docName]) {
      result[docName] = { name:docName, fileName, fileUrl, note, uploadDate:date };
    }
  }
  return Object.values(result);
}


// ══════════════════════════════════════════════════════════════════
//  DASHBOARD SHEET UPDATE
// ══════════════════════════════════════════════════════════════════
function updateDashboard() {
  const data = buildDashboardData();
  const sh   = getOrCreateSheet(SH_DASHBOARD, [
    'Metric','Value','Last Updated'
  ]);

  const now = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  sh.clearContents();
  sh.appendRow(['Metric','Value','Last Updated']);

  [
    ['Total Doctors', data.total],
    ['Complete',      data.complete],
    ['Partial',       data.partial],
    ['Not Started',   data.notStarted],
    ['Avg Completion (%)', data.avgCompletion],
    ['Total Files Uploaded', data.totalFiles],
    ['Total Notes',   data.totalNotes],
  ].forEach(r => sh.appendRow([r[0], r[1], now]));
}

function handleDashboard(e) {
  return jsonResponse({ status:'ok', data: buildDashboardData() });
}


// ══════════════════════════════════════════════════════════════════
//  LOG
// ══════════════════════════════════════════════════════════════════
function handleLog(e) {
  const params = e.parameter;
  logToSheet(params.doctorId||'—', params.actionType||'ACTION', params.details||'');
  return jsonResponse({ status:'ok' });
}

function logToSheet(doctorId, action, details) {
  const sh  = getOrCreateSheet(SH_LOGS, ['Timestamp','Employee ID','Action','Details']);
  const now = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  sh.appendRow([now, doctorId, action, details]);
}


// ══════════════════════════════════════════════════════════════════
//  DRIVE HELPERS
// ══════════════════════════════════════════════════════════════════
function getOrCreateFolder(name, parentFolder) {
  const parent = parentFolder || DriveApp.getFolderById(FOLDER_ID);
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}


// ══════════════════════════════════════════════════════════════════
//  SHEET HELPERS
// ══════════════════════════════════════════════════════════════════
function getOrCreateSheet(sheetName, headers) {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  let sheet  = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    const hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setBackground('#002654')
          .setFontColor('#FFFFFF')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
    // Auto-resize
    headers.forEach((_, i) => sheet.setColumnWidth(i+1, 160));
  }

  return sheet;
}


// ══════════════════════════════════════════════════════════════════
//  JSON RESPONSE (with CORS)
// ══════════════════════════════════════════════════════════════════
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ══════════════════════════════════════════════════════════════════
//  SETUP — شغّلها مرة واحدة
// ══════════════════════════════════════════════════════════════════
function setupProperties() {
  // غيّر الكود أدناه قبل التشغيل
  const ADMIN_CODE = 'AlHammadi@Admin2025';  // ← عدّل هذا

  PropertiesService.getScriptProperties().setProperties({
    'ADMIN_CODE': ADMIN_CODE,
  });

  Logger.log('✓ Script Properties set successfully');
  Logger.log('  ADMIN_CODE: [HIDDEN]');
}

function setupAllSheets() {
  getOrCreateSheet(SH_DOCTORS,   ['Employee ID','Doctor Name','Mobile','Department','Specialty','National Address','Qualifications','CV','SCFHS Classification','MOH License','National ID','IBAN','Malpractice Insurance','Photo','Required Docs','Uploaded Count','Notes Count','Completion %','Final Status','Last Updated']);
  getOrCreateSheet(SH_UPLOADS,   ['Timestamp','Employee ID','Doctor Name','Department','Specialty','Document Type','File Name','File URL','Note','Upload Status','Uploaded Date']);
  getOrCreateSheet(SH_LOGS,      ['Timestamp','Employee ID','Action','Details']);
  getOrCreateSheet(SH_DASHBOARD, ['Metric','Value','Last Updated']);
  Logger.log('✓ All sheets created successfully');
}
