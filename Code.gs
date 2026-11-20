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
  'Late (Min)', 'Left Early (Min)', 'Absent (Min)', 'Status', 'Extra Work (Min)', 'Holiday',
];

const SUMMARY_HEADERS = [
  'Period From', 'Period To', 'ID', 'Name', 'Department',
  'Late Times', 'Late (Min)', 'Left Early Times', 'Left Early (Min)',
  'OT Workday', 'OT Holiday', 'Scheduled Days', 'Present Days', 'Out Days',
  'Absent Days', 'Leave Days', 'Extra Work Days', 'Extra Work (Min)',
  'Half Days', 'Late Relief', 'Charged Lates', 'Late Fine (Rs)', 'Missed Punches', 'Worked on Holiday',
  'Late Dates', 'Half Day Dates', 'Extra Work Dates',
];

const PAYROLL_SHEET = 'Payroll';
const PAYROLL_HEADERS = [
  'Period From', 'Period To', 'ID', 'Name', 'Department', 'Monthly Salary', 'Per Day',
  'Working Days', 'Present Days', 'Absent Days', 'Leave Days', 'Paid Leaves Used', 'Unpaid Days',
  'Unpaid Days Deduction', 'Half Days', 'Half Day Deduction', 'Lates', 'Lates Charged', 'Late Fine',
  'Overtime (Hours)', 'Overtime Pay', 'Total Deductions', 'Net Pay',
];

// Statuses combine flags ("Late, Left Early", "Half Day, Missed Punch"); first match wins
function statusColor(status) {
  const flags = String(status).split(', ');
  if (status === 'Absent') return '#f4cccc';
  if (flags.includes('Half Day')) return '#f9cb9c';
  if (flags.includes('Missed Punch')) return '#fce5cd';
  if (flags.includes('Late') || flags.includes('Left Early')) return '#fff2cc';
  if (status === 'Leave' || status === 'Out') return '#cfe2f3';
  if (status === 'Holiday') return '#efefef';
  if (status === 'Present (Holiday)') return '#d9ead3';
  if (status === 'Present' || status === 'Present (Holiday)') return '#d9ead3';
  return null;
}

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
      r.lateMin, r.earlyMin, r.absenceMin, r.status, r.extraMin, r.holiday || '',
    ]);
    const daily = upsert(ss, DAILY_SHEET, DAILY_HEADERS, dailyRows, (row) => row[0] + '|' + row[3]);
    colorStatus(ss.getSheetByName(DAILY_SHEET));

    const summaryRows = data.summary.map((r) => [
      data.period.from, data.period.to, r.id, r.name, r.dept,
      r.lateTimes, r.lateMin, r.earlyTimes, r.earlyMin, r.overtimeWorkday, r.overtimeHoliday,
      r.scheduledDays, r.presentDays, r.outDays, r.absentDays, r.leaveDays, r.extraDays, r.extraMin,
      r.halfDays, r.lateRelief, r.chargedLates, r.lateFine, r.missedPunches, r.holidayWorkDays,
      r.lateDates || '', r.halfDayDates || '', r.extraWorkDates || '',
    ]);
    const periodKey = (row) => row[0] + '|' + row[1] + '|' + row[2];
    const summary = upsert(ss, SUMMARY_SHEET, SUMMARY_HEADERS, summaryRows, periodKey);

    let payroll = null;
    if (Array.isArray(data.payroll)) {
      const payrollRows = data.payroll.map((r) => [
        data.period.from, data.period.to, r.id, r.name, r.dept, r.salary, r.dayRate,
        r.workingDays, r.presentDays, r.absentDays, r.leaveDays, r.paidLeave, r.unpaidDays,
        r.absentDed, r.halfDays, r.halfDed, r.lateTimes, r.chargedLates, r.lateFine,
        r.otHours, r.otPay, r.deductions, r.net,
      ]);
      payroll = upsert(ss, PAYROLL_SHEET, PAYROLL_HEADERS, payrollRows, periodKey);
    }

    return json({ ok: true, period: data.period, daily, summary, payroll, sheetUrl: ss.getUrl() });
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

  // Rewritten on every upload so existing sheets pick up new columns
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

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
  const colors = range.getValues().map((r) => [statusColor(r[0])]);
  range.setBackgrounds(colors);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
