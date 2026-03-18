/***********************
 * CONFIG
 ***********************/
const SPREADSHEET_ID = '1lCqBcbGObaSu4yguvk1JyHvGAsqa4b7Q9mwDbDtwGL0';
const ADMIN_PIN = '1522'; // Stored securely on server side

const STUDENT_SHEET_NAME = 'Students';
const LOG_SHEET_NAME = 'BreakfastLog';
const LOG_ARCHIVE_SHEET_NAME = 'BreakfastLog_Archive';

const STUDENT_HEADERS = {
  STUDENT_ID: 1,
  FIRST_NAME: 2,
  LAST_NAME: 3,
  TUTOR: 4,
  CARD_ID: 5,
  VISITS: 6
};

// Performance tuning: how many recent log rows to scan for "last action today"
const LOG_LOOKBACK_ROWS = 800;

/***********************
 * WEB APP
 ***********************/
function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Breakfast Club Sign-in')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Breakfast Club')
    .addItem('Archive old logs', 'archiveOldLogs')
    .addToUi();
}

/***********************
 * SECURITY
 ***********************/
function validateAdminPin(inputPin) {
  return inputPin === ADMIN_PIN;
}

/***********************
 * HELPERS
 ***********************/

/**
 * Opens the spreadsheet once and returns references to all relevant sheets.
 * Avoids repeated openById calls within a single execution context.
 */
function getSheets_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    student: ss.getSheetByName(STUDENT_SHEET_NAME),
    log: ss.getSheetByName(LOG_SHEET_NAME),
    ss: ss
  };
}

function getArchiveSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(LOG_ARCHIVE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_ARCHIVE_SHEET_NAME);
    sheet.getRange(1, 1, 1, 5).setValues([['Date','Time','StudentID','Action','Source']]);
  }
  return sheet;
}

/***********************
 * CORE LOGIC
 ***********************/
function getInitialData() {
  const sheets = getSheets_();
  const studentSheet = sheets.student;
  const lastRow = studentSheet.getLastRow();
  const students = [];

  // Get today's status map (pass log sheet to avoid re-opening)
  const statusMap = getTodayStatusMap_(sheets.log);

  if (lastRow >= 2) {
    const data = studentSheet.getRange(2, 1, lastRow - 1, 6).getValues();

    data.forEach(r => {
      const id = r[STUDENT_HEADERS.STUDENT_ID - 1];
      if (!id) return;

      const first = r[STUDENT_HEADERS.FIRST_NAME - 1] || '';
      const last = r[STUDENT_HEADERS.LAST_NAME - 1] || '';

      students.push({
        studentId: String(id),
        displayName: `${first} ${last}`.trim() || String(id),
        tutor: r[STUDENT_HEADERS.TUTOR - 1] || '',
        status: statusMap[String(id)] || null
      });
    });
  }

  return { students };
}

/**
 * Scans log backwards for today's entries. Breaks early once it hits
 * a date before today (assumes chronological append order).
 */
function getTodayStatusMap_(logSheet) {
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return {};

  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const startRow = Math.max(2, lastRow - LOG_LOOKBACK_ROWS + 1);
  const numRows = lastRow - startRow + 1;
  const values = logSheet.getRange(startRow, 1, numRows, 4).getValues();

  const map = {};

  // Scan backwards — break early once we pass today
  for (let i = values.length - 1; i >= 0; i--) {
    const r = values[i];
    if (!r[0] || !r[2]) continue;

    const d = Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd');
    if (d < today) break; // All remaining rows are older — stop

    if (d === today) {
      const sid = String(r[2]);
      // Only keep the most recent action per student (first hit scanning backwards)
      if (!map[sid]) {
        map[sid] = r[3]; // 'IN' or 'OUT'
      }
    }
  }

  return map;
}

function logSignAction(studentId, newStatus) {
  if (!studentId || !newStatus) return { success: false, message: 'Missing data' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    // Open spreadsheet once — pass sheets to all helpers
    const sheets = getSheets_();
    const logSheet = sheets.log;
    const studentSheet = sheets.student;
    const now = new Date();
    const tz = Session.getScriptTimeZone();

    // 1. Check for duplicates (single log scan)
    const currentStatus = getLastActionToday_(logSheet, studentId, now);
    if (currentStatus === newStatus) {
      return { success: true, skippedDuplicate: true, studentId: String(studentId), newStatus };
    }

    // 2. Increment visits if this is the FIRST IN of the day
    //    (uses the result we already fetched — no second scan)
    if (newStatus === 'IN' && currentStatus !== 'IN') {
      incrementVisitCount_(studentSheet, studentId);
    }

    // 3. Log it
    const dateOnly = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');

    const nextRow = logSheet.getLastRow() + 1;
    logSheet.getRange(nextRow, 1, 1, 5).setValues([[dateOnly, timeStr, studentId, newStatus, 'WebApp']]);

    return { success: true, studentId: String(studentId), newStatus };

  } catch (e) {
    console.error('logSignAction error:', e);
    return { success: false, message: 'Server busy, try again.' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Scans recent log rows backwards to find the student's last action today.
 * Accepts a sheet reference to avoid re-opening the spreadsheet.
 */
function getLastActionToday_(logSheet, studentId, now) {
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return null;

  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

  const startRow = Math.max(2, lastRow - LOG_LOOKBACK_ROWS + 1);
  const numRows = lastRow - startRow + 1;
  const values = logSheet.getRange(startRow, 1, numRows, 4).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const r = values[i];
    if (String(r[2]) === String(studentId)) {
      const d = Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd');
      if (d === today) return r[3];
      if (d < today) return null;
    }
  }
  return null;
}

/**
 * Increments the visit count for a student. Called only when we already
 * know this is the first IN of the day (no redundant log check).
 */
function incrementVisitCount_(studentSheet, studentId) {
  const data = studentSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][STUDENT_HEADERS.STUDENT_ID - 1]) === String(studentId)) {
      const current = Number(data[i][STUDENT_HEADERS.VISITS - 1]) || 0;
      studentSheet.getRange(i + 1, STUDENT_HEADERS.VISITS).setValue(current + 1);
      break;
    }
  }
}

function archiveOldLogs() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOG_SHEET_NAME);
  const archive = getArchiveSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const tz = Session.getScriptTimeZone();
  const cutoff = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  const cutoffStr = Utilities.formatDate(cutoff, tz, 'yyyy-MM-dd');

  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const keep = [];
  const move = [];

  values.forEach(r => {
    const d = r[0] ? Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd') : '';
    if (d && d < cutoffStr) move.push(r);
    else keep.push(r);
  });

  if (move.length) {
    archive.getRange(archive.getLastRow() + 1, 1, move.length, 5).setValues(move);
  }

  sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  if (keep.length) {
    sheet.getRange(2, 1, keep.length, 5).setValues(keep);
  }
}