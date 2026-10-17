/**
 * Attendance upload endpoint.
 * Receives parsed attendance JSON from index.html and writes it into this spreadsheet.
 * Re-uploading the same period overwrites matching rows instead of duplicating them.
 */

// Leave empty if this script is bound to the spreadsheet (Extensions > Apps Script).
// Otherwise paste the spreadsheet ID from its URL: docs.google.com/spreadsheets/d/<ID>/edit
const SPREADSHEET_ID = '1Zu1TopmawLXq6XwVtHwKqX_aTMagYyGktNbnus3gOgU';

const DAILY_SHEET = 'Daily Attendance';
const SUMMARY_SHEET = 'Summary';

const DAILY_HEADERS = [
  'ID', 'Name', 'Department', 'Date', 'Day', 'In', 'Out', 'All Punches',
  'Late (Min)', 'Left Early (Min)', 'Absent (Min)', 'Status',
];

const SUMMARY_HEADERS = [
  'Period From', 'Period To', 'ID', 'Name', 'Department', 'Normal Hours', 'Real Hours',
  'Late Times', 'Late (Min)', 'Left Early Times', 'Left Early (Min)',
  'OT Workday', 'OT Holiday', 'Scheduled Days', 'Present Days', 'Out Days',
  'Absent Days', 'Leave Days',
];

const STATUS_COLORS = {
  'Present': '#d9ead3',
  'Late': '#fff2cc',
  'Left Early': '#fff2cc',
  'Late, Left Early': '#fce5cd',
  'Missed Punch': '#fce5cd',
  'Absent': '#f4cccc',
  'Holiday': '#efefef',
  'Present (Holiday)': '#d9ead3',
  'Leave': '#cfe2f3',
  'Out': '#cfe2f3',
};

function doGet() {
  return json({ ok: true, message: 'Attendance endpoint is running.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (!data.daily || !data.summary || !data.period) throw new Error('Invalid payload.');

    // Accept either the bare ID or a full sheet URL.
    const sheetId = (SPREADSHEET_ID.match(/\/d\/([a-zA-Z0-9_-]+)/) || [, SPREADSHEET_ID.trim()])[1];
    const ss = sheetId ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.getActiveSpreadsheet();

    const dailyRows = data.daily.map((r) => [
      r.id, r.name, r.dept, r.date, r.day, r.inTime, r.outTime, r.punches,
      r.lateMin, r.earlyMin, r.absenceMin, r.status,
    ]);
    const daily = upsert(ss, DAILY_SHEET, DAILY_HEADERS, dailyRows, (row) => row[0] + '|' + row[3]);
    colorStatus(ss.getSheetByName(DAILY_SHEET));

    const summaryRows = data.summary.map((r) => [
      data.period.from, data.period.to, r.id, r.name, r.dept, r.normalHours, r.realHours,
      r.lateTimes, r.lateMin, r.earlyTimes, r.earlyMin, r.overtimeWorkday, r.overtimeHoliday,
      r.scheduledDays, r.presentDays, r.outDays, r.absentDays, r.leaveDays,
    ]);
    const summary = upsert(ss, SUMMARY_SHEET, SUMMARY_HEADERS, summaryRows, (row) => row[0] + '|' + row[1] + '|' + row[2]);

    return json({ ok: true, period: data.period, daily, summary, sheetUrl: ss.getUrl() });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Writes rows into a sheet, replacing rows whose key already exists and appending the rest.
 * Returns { added, updated }.
 */
function upsert(ss, name, headers, rows, keyOf) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  const width = headers.length;
  const lastRow = sheet.getLastRow();
  const existing = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, width).getDisplayValues() : [];

  const index = {};
  existing.forEach((row, i) => (index[keyOf(row.map(String))] = i));

  let added = 0;
  let updated = 0;
  rows.forEach((row) => {
    const key = keyOf(row.map(String));
    if (key in index) {
      existing[index[key]] = row;
      updated++;
    } else {
      index[key] = existing.length;
      existing.push(row);
      added++;
    }
  });

  if (existing.length) {
    // Plain text keeps dates and times exactly as in the report (no locale conversion)
    const range = sheet.getRange(2, 1, existing.length, width);
    range.setNumberFormat('@');
    range.setValues(existing);
  }
  sheet.autoResizeColumns(1, width);
  return { added, updated };
}

function colorStatus(sheet) {
  const n = sheet.getLastRow() - 1;
  if (n < 1) return;
  const col = DAILY_HEADERS.indexOf('Status') + 1;
  const range = sheet.getRange(2, col, n, 1);
  const colors = range.getValues().map((r) => [STATUS_COLORS[r[0]] || null]);
  range.setBackgrounds(colors);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
