/***********************
 * CONFIG
 ***********************/
// Defaults — override via PropertiesService (run setupProperties_ once)
const DEFAULT_SPREADSHEET_ID = '1lCqBcbGObaSu4yguvk1JyHvGAsqa4b7Q9mwDbDtwGL0';
const DEFAULT_ADMIN_PIN = '1522';

function getSpreadsheetId_() {
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID;
}

function getAdminPin_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || DEFAULT_ADMIN_PIN;
}

/**
 * Run once from the script editor to store secrets in PropertiesService.
 * After running, the DEFAULT_ constants above can be removed from source.
 */
function setupProperties_() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', DEFAULT_SPREADSHEET_ID);
  props.setProperty('ADMIN_PIN', DEFAULT_ADMIN_PIN);
}

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
const LOG_LOOKBACK_ROWS = 1500;

/***********************
 * WEB APP
 ***********************/
function doGet(e) {
  const view = (e && e.parameter && e.parameter.view) ? String(e.parameter.view).toLowerCase() : '';

  let webAppUrl = '';
  try { webAppUrl = ScriptApp.getService().getUrl() || ''; } catch (err) { /* preview/editor context */ }

  // Explicit mobile request (or a phone auto-redirected here from the kiosk).
  if (view === 'mobile') {
    return HtmlService
      .createTemplateFromFile('Mobile')
      .evaluate()
      .setTitle('Rise and Shine — Mobile')
      // Apps Script's addMetaTag only whitelists a few names — viewport and
      // apple-mobile-web-app-capable are allowed; status-bar-style/theme-color
      // are rejected ("meta tag not allowed in this context"), so omit them.
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover')
      .addMetaTag('apple-mobile-web-app-capable', 'yes')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Default (no param) and ?view=kiosk both serve the iPad kiosk. When no view
  // is forced, the kiosk page auto-redirects phone-sized devices to the mobile
  // view (see the detection script at the top of Index.html). The iPad sees no
  // change — detection returns false, so it loads the kiosk directly.
  const t = HtmlService.createTemplateFromFile('Index');
  t.webAppUrl = webAppUrl;
  t.forcedView = (view === 'kiosk') ? 'kiosk' : '';
  return t
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
  return inputPin === getAdminPin_();
}

/***********************
 * HELPERS
 ***********************/

/**
 * Opens the spreadsheet and returns references to all relevant sheets.
 * Call once per server function to avoid repeated openById calls.
 */
function getSheets_() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  return {
    student: ss.getSheetByName(STUDENT_SHEET_NAME),
    log: ss.getSheetByName(LOG_SHEET_NAME),
    ss: ss
  };
}

function getArchiveSheet_() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
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

  // Use cached status map when available (reduces Sheets API load from polling)
  const statusMap = getCachedStatusMap_(sheets.log);

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
        visits: Number(r[STUDENT_HEADERS.VISITS - 1]) || 0,
        status: statusMap[String(id)] || null
      });
    });
  }

  return { students };
}

/**
 * Returns today's status map from CacheService if available,
 * otherwise builds it from the log sheet and caches for 20 seconds.
 */
function getCachedStatusMap_(logSheet) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('todayStatusMap');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* fall through */ }
  }

  const map = getTodayStatusMap_(logSheet);
  try { cache.put('todayStatusMap', JSON.stringify(map), 20); } catch (e) { /* cache write failure is non-fatal */ }
  return map;
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

  try {
    const sheets = getSheets_();
    const logSheet = sheets.log;
    const studentSheet = sheets.student;
    const now = new Date();
    const tz = Session.getScriptTimeZone();

    // 1. Duplicate check (lockless read — rare race is acceptable)
    const currentStatus = getLastActionToday_(logSheet, studentId, now);
    if (currentStatus === newStatus) {
      return { success: true, skippedDuplicate: true, studentId: String(studentId), newStatus };
    }

    // 2. Append log row (lockless — Sheets handles concurrent appends safely)
    const dateOnly = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
    const nextRow = logSheet.getLastRow() + 1;
    logSheet.getRange(nextRow, 1, 1, 5).setValues([[dateOnly, timeStr, studentId, newStatus, 'WebApp']]);

    // 3. Invalidate cached status map so next poll picks up this change
    try { CacheService.getScriptCache().remove('todayStatusMap'); } catch (e) {}

    // 4. Increment visits outside the main write path (brief targeted lock)
    //    Failure here is non-fatal — the sign-in log is already saved
    if (newStatus === 'IN' && currentStatus !== 'IN') {
      try {
        const lock = LockService.getScriptLock();
        lock.waitLock(10000);
        try {
          incrementVisitCount_(studentSheet, studentId);
        } finally {
          try { lock.releaseLock(); } catch (e) {}
        }
      } catch (e) {
        console.error('incrementVisitCount_ lock timeout (non-fatal):', e);
      }
    }

    return { success: true, studentId: String(studentId), newStatus };

  } catch (e) {
    console.error('logSignAction error:', e);
    return { success: false, message: 'Server busy, try again.' };
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
  const lastRow = studentSheet.getLastRow();
  if (lastRow < 2) return;

  // Read only column 1 (student IDs) to find the row — avoids reading entire sheet
  const ids = studentSheet.getRange(2, STUDENT_HEADERS.STUDENT_ID, lastRow - 1, 1).getValues();
  const sid = String(studentId);

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === sid) {
      const row = i + 2;
      const current = Number(studentSheet.getRange(row, STUDENT_HEADERS.VISITS).getValue()) || 0;
      studentSheet.getRange(row, STUDENT_HEADERS.VISITS).setValue(current + 1);
      break;
    }
  }
}

function archiveOldLogs() {
  const sheets = getSheets_();
  const sheet = sheets.log;
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

/***********************
 * AUTO SIGN-OUT
 ***********************/

/**
 * Signs out all students still marked IN for today.
 * Set up as a time-driven trigger (e.g., daily at 09:30).
 * Run setupAutoSignOutTrigger() once from the script editor to install.
 */
function autoSignOutEndOfDay() {
  const sheets = getSheets_();
  const logSheet = sheets.log;
  const statusMap = getTodayStatusMap_(logSheet);

  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const dateOnly = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');

  const stillIn = Object.entries(statusMap).filter(([_, status]) => status === 'IN');
  if (stillIn.length === 0) return;

  const rows = stillIn.map(([sid]) => [dateOnly, timeStr, sid, 'OUT', 'AutoSignOut']);
  const nextRow = logSheet.getLastRow() + 1;
  logSheet.getRange(nextRow, 1, rows.length, 5).setValues(rows);

  try { CacheService.getScriptCache().remove('todayStatusMap'); } catch (e) {}
}

/**
 * Installs a daily trigger for autoSignOutEndOfDay at 09:30.
 * Run once from the script editor. Safe to re-run (removes duplicates).
 */
function setupAutoSignOutTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'autoSignOutEndOfDay') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('autoSignOutEndOfDay')
    .timeBased()
    .atHour(9)
    .nearMinute(30)
    .everyDays(1)
    .create();
}