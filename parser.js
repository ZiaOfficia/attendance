// Parses the biometric "StandardReport.xls" export into daily and summary records.
// Works in the browser (window.AttendanceParser) and in Node (module.exports).
(function (root) {
  'use strict';

  const DAY_NAMES = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'];

  function rows(XLSX, wb, name) {
    const ws = wb.Sheets[name];
    if (!ws) throw new Error('Sheet "' + name + '" not found in the file.');
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  }

  function str(v) {
    return String(v == null ? '' : v).trim();
  }

  function num(v) {
    const n = parseFloat(v);
    return isNaN(n) ? 0 : n;
  }

  // "2026-10-01 ~ 2026-10-15" -> list of {iso, dayOfMonth, dayName}
  function parsePeriod(text) {
    const m = str(text).match(/(\d{4})-(\d{2})-(\d{2})\s*~\s*(\d{4})-(\d{2})-(\d{2})/);
    if (!m) throw new Error('Could not read the report period (Stat.Date).');
    const start = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    const end = new Date(Date.UTC(+m[4], +m[5] - 1, +m[6]));
    const days = [];
    for (let d = new Date(start); d <= end; d.setUTCDate(d.getUTCDate() + 1)) {
      days.push({
        iso: d.toISOString().slice(0, 10),
        dayOfMonth: d.getUTCDate(),
        dayName: DAY_NAMES[d.getUTCDay()],
      });
    }
    return { from: days[0].iso, to: days[days.length - 1].iso, days };
  }

  // "10:3918:31" -> ["10:39", "18:31"]
  function splitPunches(cell) {
    return str(cell).match(/\d{1,2}:\d{2}/g) || [];
  }

  function parse(XLSX, wb) {
    const sched = rows(XLSX, wb, 'Schedule Infor.');
    const stat = rows(XLSX, wb, 'Att. Stat.');
    const log = rows(XLSX, wb, 'Att.log report');
    const exc = rows(XLSX, wb, 'Exception Stat.');

    const period = parsePeriod(sched[1][1] || stat[1][1]);

    // Schedule: shift code per employee per day-of-month column
    const schedHeader = sched[2];
    const schedule = {}; // id -> {dayOfMonth: code}
    for (let r = 4; r < sched.length; r++) {
      const id = str(sched[r][0]);
      if (!id) continue;
      schedule[id] = {};
      for (let c = 3; c < schedHeader.length; c++) {
        schedule[id][num(schedHeader[c])] = str(sched[r][c]);
      }
    }

    // Exceptions: late / early / absence minutes per id+date
    const exceptions = {};
    for (let r = 4; r < exc.length; r++) {
      const id = str(exc[r][0]);
      const date = str(exc[r][3]);
      if (!id || !/^\d{4}-\d{2}-\d{2}$/.test(date)) continue;
      exceptions[id + '|' + date] = {
        late: num(exc[r][8]),
        early: num(exc[r][9]),
        absence: num(exc[r][10]),
      };
    }

    // Punch log: "ID:" row followed by a row of punches, columns = day-of-month order
    const dayCols = log[3].map(num);
    const employees = [];
    const punches = {}; // id -> {dayOfMonth: [times]}
    for (let r = 4; r < log.length; r++) {
      if (str(log[r][0]) !== 'ID:') continue;
      const id = str(log[r][2]);
      employees.push({ id, name: str(log[r][10]), dept: str(log[r][20]) });
      punches[id] = {};
      const p = log[r + 1] || [];
      dayCols.forEach((dom, c) => {
        if (dom) punches[id][dom] = splitPunches(p[c]);
      });
    }

    const daily = [];
    employees.forEach((emp) => {
      period.days.forEach((day) => {
        const times = (punches[emp.id] || {})[day.dayOfMonth] || [];
        const code = (schedule[emp.id] || {})[day.dayOfMonth];
        const ex = exceptions[emp.id + '|' + day.iso] || { late: 0, early: 0, absence: 0 };
        // A lone punch in the afternoon is a check-out with a missed check-in
        const loneOut = times.length === 1 && parseInt(times[0], 10) >= 14;
        const inTime = loneOut ? '' : times[0] || '';
        const outTime = times.length > 1 ? times[times.length - 1] : loneOut ? times[0] : '';

        let status;
        if (code === '25') status = 'Leave';
        else if (code === '26') status = 'Out';
        else if (code === '' || code === undefined) status = times.length ? 'Present (Holiday)' : 'Holiday';
        else if (!times.length) status = 'Absent';
        else if (times.length === 1) status = 'Missed Punch';
        else {
          const flags = [];
          if (ex.late > 0) flags.push('Late');
          if (ex.early > 0) flags.push('Left Early');
          status = flags.length ? flags.join(', ') : 'Present';
        }

        daily.push({
          id: emp.id,
          name: emp.name,
          dept: emp.dept,
          date: day.iso,
          day: day.dayName,
          inTime,
          outTime,
          punches: times.join(' '),
          lateMin: ex.late,
          earlyMin: ex.early,
          absenceMin: ex.absence,
          status,
        });
      });
    });

    const summary = [];
    for (let r = 4; r < stat.length; r++) {
      const s = stat[r];
      const id = str(s[0]);
      if (!id) continue;
      const days = str(s[11]).split('/');
      summary.push({
        id,
        name: str(s[1]),
        dept: str(s[2]),
        lateTimes: num(s[5]),
        lateMin: num(s[6]),
        earlyTimes: num(s[7]),
        earlyMin: num(s[8]),
        overtimeWorkday: str(s[9]),
        overtimeHoliday: str(s[10]),
        scheduledDays: num(days[0]),
        presentDays: num(days[1]),
        outDays: num(s[12]),
        absentDays: num(s[13]),
        leaveDays: num(s[14]),
      });
    }

    return { period: { from: period.from, to: period.to }, daily, summary };
  }

  const api = { parse };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.AttendanceParser = api;
})(this);
