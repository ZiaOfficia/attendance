// Parses the biometric "StandardReport.xls" export into daily and summary records.
// Works in the browser (window.AttendanceParser) and in Node (module.exports).
(function (root) {
  'use strict';

  const DAY_NAMES = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'];
  // Office rules (minutes since midnight)
  const SHIFT_START = 10 * 60 + 45; // 10:45 AM entry, every working day
  const HALF_DAY_AFTER = 11 * 60 + 30; // Punching in after 11:30 AM is a half day
  // Exit rules. Leaving before halfDayBefore = half day; before lateBefore = a late;
  // lateBefore up to end = grace; after end = extra work (overtime).
  const WEEKDAY_EXIT = { halfDayBefore: 17 * 60 + 30, lateBefore: 18 * 60 + 15, end: 18 * 60 + 30 }; // 5:30, 6:15, 6:30 PM
  const SATURDAY_EXIT = { halfDayBefore: 15 * 60 + 30, lateBefore: 16 * 60 + 15, end: 16 * 60 + 30 }; // 3:30, 4:15, 4:30 PM
  const DEFAULTS = {
    lateFine: 100, // Rs per chargeable late
    freeLatesPerMonth: 2, // First lates each month are not charged
    holidays: {}, // 'YYYY-MM-DD' -> name; days off entered by hand, on top of the machine's schedule
  };

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

  // "18:42" -> 1122
  function minutes(time) {
    const [h, m] = time.split(':').map(Number);
    return h * 60 + m;
  }

  function parse(XLSX, wb, options) {
    const opt = Object.assign({}, DEFAULTS, options);
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

        const exit = day.dayName === 'SAT' ? SATURDAY_EXIT : WEEKDAY_EXIT;
        const shiftEnd = exit.end;
        const inMin = inTime ? minutes(inTime) : null;
        const outMin = times.length > 1 || loneOut ? minutes(outTime) : null;
        let lateMin = 0;
        let earlyMin = 0;
        let lateCount = 0; // Lates charged against this day: late entry and late-counted exit each count once

        const holiday = opt.holidays[day.iso] || '';

        let status;
        if (holiday || code === '' || code === undefined) status = times.length ? 'Present (Holiday)' : 'Holiday';
        else if (code === '25') status = 'Leave';
        else if (code === '26') status = 'Out';
        else if (!times.length) status = 'Absent';
        else {
          const flags = [];
          if (inMin !== null && inMin > HALF_DAY_AFTER) flags.push('Half Day');
          else if (inMin !== null && inMin > SHIFT_START) {
            flags.push('Late');
            lateMin = inMin - SHIFT_START;
            lateCount++;
          }
          if (times.length > 1 && outMin < shiftEnd) {
            earlyMin = shiftEnd - outMin;
            if (outMin < exit.halfDayBefore) {
              if (!flags.includes('Half Day')) flags.push('Half Day');
            } else if (outMin < exit.lateBefore) {
              flags.push('Left Early');
              lateCount++;
            } else earlyMin = 0; // Grace window (6:15-6:29 PM, Sat 4:15-4:29 PM): not flagged or counted
          }
          if (times.length === 1) flags.push('Missed Punch');
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
          lateMin,
          earlyMin,
          lateCount,
          absenceMin: ex.absence,
          extraMin: outMin !== null ? Math.max(0, outMin - shiftEnd) : 0,
          status,
          holiday,
        });
      });
    });

    const summary = [];
    for (let r = 4; r < stat.length; r++) {
      const s = stat[r];
      const id = str(s[0]);
      if (!id) continue;
      const mine = daily.filter((d) => d.id === id);
      const countOf = (fn) => mine.filter((d) => fn(d.status)).length;
      const extra = mine.filter((d) => d.extraMin > 0);
      const lates = mine.filter((d) => d.lateCount > 0);
      const lateTimes = lates.reduce((total, d) => total + d.lateCount, 0);
      const early = mine.filter((d) => d.earlyMin > 0);

      // Relief is per calendar month: the first FREE_LATES_PER_MONTH lates of each month are not charged
      const latesByMonth = {};
      lates.forEach((d) => (latesByMonth[d.date.slice(0, 7)] = (latesByMonth[d.date.slice(0, 7)] || 0) + d.lateCount));
      const chargedLates = Object.values(latesByMonth)
        .reduce((total, n) => total + Math.max(0, n - opt.freeLatesPerMonth), 0);

      summary.push({
        id,
        name: str(s[1]),
        dept: str(s[2]),
        lateTimes,
        lateMin: lates.reduce((total, d) => total + d.lateMin, 0),
        lateRelief: lateTimes - chargedLates,
        chargedLates,
        lateFine: chargedLates * opt.lateFine,
        halfDays: mine.filter((d) => d.status.includes('Half Day')).length,
        earlyTimes: early.length,
        earlyMin: early.reduce((total, d) => total + d.earlyMin, 0),
        overtimeWorkday: str(s[9]),
        overtimeHoliday: str(s[10]),
        // Day counts come from the daily records so hand-entered holidays are reflected
        scheduledDays: countOf((st) => st !== 'Holiday' && st !== 'Present (Holiday)'),
        presentDays: countOf((st) => !['Absent', 'Holiday', 'Leave', 'Out', 'Present (Holiday)'].includes(st)),
        holidayWorkDays: countOf((st) => st === 'Present (Holiday)'),
        outDays: countOf((st) => st === 'Out'),
        absentDays: countOf((st) => st === 'Absent'),
        leaveDays: countOf((st) => st === 'Leave'),
        missedPunches: countOf((st) => st.split(', ').includes('Missed Punch')),
        extraDays: extra.length,
        extraMin: extra.reduce((total, d) => total + d.extraMin, 0),
      });
    }

    return { period: { from: period.from, to: period.to }, daily, summary };
  }

  const api = { parse, DEFAULTS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.AttendanceParser = api;
})(this);
