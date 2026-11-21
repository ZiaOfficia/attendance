// Attendance & Payroll app: reads the machine report with parser.js, applies holidays and salary rules,
// and renders the dashboard, calendar, payroll and Google Sheet sync.
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const rs = (n) => 'Rs ' + Math.round(n).toLocaleString('en-US');
  const int = (n) => Math.round(n).toLocaleString('en-US');
  const hours = (min) => (min / 60).toFixed(1).replace(/\.0$/, '');
  const initials = (name) => String(name).split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0]).join('').toUpperCase() || '?';
  const cssVar = (name) => getComputedStyle(document.documentElement).getPropertyValue(name).trim();
  const WEEKDAY_SHIFT_HOURS = 7.75; // 10:45 AM to 6:30 PM, used for the overtime hourly rate

  // ---------- Saved preferences (this browser only) ----------
  const STORE_KEY = 'attendance.v1';
  const DEFAULT_SETTINGS = { basis: '30', half: '0.5', fine: 100, free: 2, ot: 'none', otRate: 0, paidLeaves: 2, ignoreOff: false, sundayOff: true };
  const saved = (() => { try { return JSON.parse(localStorage.getItem(STORE_KEY)) || {}; } catch (e) { return {}; } })();
  const prefs = {
    salaries: saved.salaries || {},
    holidays: saved.holidays || {},
    noExtra: saved.noExtra || {}, // 'YYYY-MM-DD' -> {name, after}: party / event days, extra work not counted
    settings: Object.assign({}, DEFAULT_SETTINGS, saved.settings),
    url: saved.url || '',
    theme: saved.theme || '',
  };
  function persist() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(prefs)); } catch (e) { /* storage unavailable: keep working in memory */ }
  }

  const state = { wb: null, fileName: '', data: null, pay: [], tab: 'dashboard', dashEmp: '' };
  const charts = {};

  // ---------- Status groups (stack order validated for colour separation) ----------
  const GROUPS = [
    { key: 'ok', label: 'On time', cls: 'c-ok', color: '--st-ok', pill: 't-good' },
    { key: 'late', label: 'Late / late exit', cls: 'c-late', color: '--st-late', pill: 't-warn' },
    { key: 'leave', label: 'Leave / out', cls: 'c-leave', color: '--st-leave', pill: 't-info' },
    { key: 'half', label: 'Half day', cls: 'c-half', color: '--st-half', pill: 't-half' },
    { key: 'miss', label: 'Missed punch', cls: 'c-miss', color: '--st-miss', pill: 't-miss' },
    { key: 'absent', label: 'Absent', cls: 'c-absent', color: '--st-absent', pill: 't-bad' },
  ];
  const G = Object.fromEntries(GROUPS.map((g) => [g.key, g]));
  const has = (status, flag) => status.split(', ').includes(flag);
  function groupOf(status) {
    if (status === 'Absent') return G.absent;
    if (has(status, 'Half Day')) return G.half;
    if (has(status, 'Missed Punch')) return G.miss;
    if (has(status, 'Late') || has(status, 'Left Early')) return G.late;
    if (status === 'Present' || status === 'Present (Holiday)') return G.ok;
    if (status === 'Leave' || status === 'Out') return G.leave;
    return null; // Holiday
  }
  const LETTER = { 'Late': 'L', 'Half Day': 'H', 'Left Early': 'E', 'Missed Punch': 'M' };
  const SHORT = { 'Present': 'P', 'Present (Holiday)': 'P', 'Absent': 'A', 'Holiday': '', 'Leave': 'LV', 'Out': 'O' };
  const shortOf = (s) => (s in SHORT ? SHORT[s] : s.split(', ').map((f) => LETTER[f] || '').join(''));
  const legendHtml = (items) => items.map((g) => '<span><i class="' + g.cls + '"></i>' + esc(g.label) + '</span>').join('');
  function pill(status) {
    const g = groupOf(status);
    const label = status === 'Present (Holiday)' ? 'Worked on holiday' : status.replace('Left Early', 'Late Exit');
    return '<span class="pill ' + (g ? g.pill : 't-off') + '"><i class="' + (g ? g.cls : 'c-off') + '"></i>' + esc(label) + '</span>';
  }

  // ---------- Theme ----------
  function applyTheme() {
    if (prefs.theme) document.documentElement.dataset.theme = prefs.theme;
    else delete document.documentElement.dataset.theme;
  }
  const isDark = () => prefs.theme ? prefs.theme === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
  applyTheme();
  $('themeBtn').addEventListener('click', () => {
    prefs.theme = isDark() ? 'light' : 'dark';
    persist(); applyTheme(); if (state.data) renderCharts();
  });
  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => { if (state.data && !prefs.theme) renderCharts(); });

  // ---------- Toast ----------
  let toastTimer;
  function toast(msg) {
    const t = $('toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), 2400);
  }

  // ---------- File loading ----------
  const drop = $('drop');
  drop.addEventListener('click', () => $('file').click());
  drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('file').click(); } });
  ['dragover', 'dragenter'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  document.addEventListener('dragleave', (e) => { if (!e.relatedTarget) drop.classList.remove('over'); });
  document.addEventListener('drop', (e) => {
    e.preventDefault(); drop.classList.remove('over');
    if (e.dataTransfer.files[0]) load(e.dataTransfer.files[0]);
  });
  $('file').addEventListener('change', (e) => { if (e.target.files[0]) load(e.target.files[0]); e.target.value = ''; });
  $('replaceBtn').addEventListener('click', () => $('file').click());

  async function load(file) {
    $('parseMsg').classList.add('hidden');
    try {
      state.wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      state.fileName = file.name;
      reparse();
    } catch (err) {
      if (state.data) { toast('Could not read ' + file.name + ': ' + err.message); return; }
      $('parseMsg').innerHTML = '<b>Could not read this file.</b>&nbsp;' + esc(err.message);
      $('parseMsg').classList.remove('hidden');
      return;
    }
    $('empty').classList.add('hidden');
    ['nav', 'replaceBtn', 'periodBadge'].forEach((id) => $(id).classList.remove('hidden'));
    showTab(state.tab);
    toast('Loaded ' + file.name);
  }

  function reparse() {
    const s = prefs.settings;
    state.data = AttendanceParser.parse(XLSX, state.wb, {
      holidays: prefs.holidays,
      noExtraDays: prefs.noExtra,
      ignoreScheduleOff: s.ignoreOff === true,
      sundaysOff: s.sundayOff !== false,
      lateFine: Math.max(0, Number(s.fine) || 0),
      freeLatesPerMonth: Math.max(0, Math.floor(Number(s.free) || 0)),
    });
    computePay();
    renderAll();
  }

  // ---------- Payroll ----------
  function daysInMonth(iso) {
    const [y, m] = iso.split('-').map(Number);
    return new Date(Date.UTC(y, m, 0)).getUTCDate();
  }
  function isFullMonth(p) {
    return p.from.slice(8) === '01' && p.from.slice(0, 7) === p.to.slice(0, 7) && Number(p.to.slice(8)) === daysInMonth(p.to);
  }

  function computePay() {
    const s = prefs.settings;
    const basis = s.basis === 'calendar' ? daysInMonth(state.data.period.from) : Number(s.basis);
    const allowance = Math.max(0, Math.floor(Number(s.paidLeaves) || 0));

    // Days off (leave + absent) per employee per calendar month; each month's first `allowance` days are paid
    const offByMonth = {};
    state.data.daily.forEach((d) => {
      if (d.status !== 'Absent' && d.status !== 'Leave') return;
      const m = (offByMonth[d.id] = offByMonth[d.id] || {});
      m[d.date.slice(0, 7)] = (m[d.date.slice(0, 7)] || 0) + 1;
    });

    state.pay = state.data.summary.map((e) => {
      const salary = Math.max(0, Number(prefs.salaries[e.id]) || 0);
      const dayRate = salary / basis;
      const offDays = e.absentDays + e.leaveDays;
      const paidLeave = Object.values(offByMonth[e.id] || {}).reduce((t, n) => t + Math.min(allowance, n), 0);
      const unpaidDays = offDays - paidLeave;
      const absentDed = Math.round(unpaidDays * dayRate); // Unpaid days off: absences and leave beyond the paid allowance
      const halfDed = Math.round(e.halfDays * Number(s.half) * dayRate);
      const lateFine = e.lateFine;
      const otRate = s.ot === 'auto' ? dayRate / WEEKDAY_SHIFT_HOURS : s.ot === 'fixed' ? Math.max(0, Number(s.otRate) || 0) : 0;
      const otPay = Math.round((e.extraMin / 60) * otRate);
      const deductions = absentDed + halfDed + lateFine;
      return {
        id: e.id, name: e.name, dept: e.dept, salary, hasSalary: salary > 0, dayRate,
        workingDays: e.scheduledDays, presentDays: e.presentDays,
        absentDays: e.absentDays, offDays, paidLeave, unpaidDays, absentDed, halfDays: e.halfDays, halfDed,
        lateTimes: e.lateTimes, lateRelief: e.lateRelief, chargedLates: e.chargedLates, lateFine,
        leaveDays: e.leaveDays, otMin: e.extraMin, otPay, deductions,
        net: salary ? Math.max(0, salary - deductions + otPay) : 0,
      };
    });
  }
  const payOf = (id) => state.pay.find((p) => p.id === id);

  // ---------- Tabs ----------
  document.querySelectorAll('[data-tab]').forEach((b) => b.addEventListener('click', () => showTab(b.dataset.tab)));
  function showTab(tab) {
    state.tab = tab;
    document.querySelectorAll('[data-tab]').forEach((b) => b.classList.toggle('on', b.dataset.tab === tab));
    ['dashboard', 'attendance', 'payroll', 'settings', 'sync'].forEach((t) => $('tab-' + t).classList.toggle('hidden', t !== tab));
    if (tab === 'dashboard') renderCharts();
    window.scrollTo({ top: 0 });
  }

  function renderAll() {
    const p = state.data.period;
    $('periodBadge').innerHTML = '<svg class="i" viewBox="0 0 24 24" style="width:14px;height:14px"><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/></svg>' +
      esc(fmtDate(p.from) + ' – ' + fmtDate(p.to));
    $('dashSub').textContent = state.fileName + ' · ' + state.data.summary.length + ' employees';
    renderKpis();
    renderDashPicker();
    renderCalendar();
    renderEmpFilter();
    renderDailyTable();
    renderPayroll();
    renderSettings();
    if (state.tab === 'dashboard') renderCharts();
    if (openId) openDetail(openId);
  }

  function fmtDate(iso, withYear = true) {
    const d = new Date(iso + 'T00:00:00Z');
    return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: withYear ? 'numeric' : undefined, timeZone: 'UTC' });
  }

  // ---------- KPI tiles ----------
  const ICONS = {
    users: '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.9M16 3.1a4 4 0 0 1 0 7.8"/>',
    check: '<circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>',
    clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    half: '<circle cx="12" cy="12" r="9"/><path d="M12 3v18" /><path d="M12 3a9 9 0 0 1 0 18Z" fill="currentColor" stroke="none" opacity=".35"/>',
    x: '<circle cx="12" cy="12" r="9"/><path d="m9 9 6 6m0-6-6 6"/>',
    coin: '<circle cx="12" cy="12" r="9"/><path d="M15 9.5c-.5-1-1.6-1.5-3-1.5-1.7 0-3 .8-3 2s1.3 1.7 3 2 3 .8 3 2-1.3 2-3 2c-1.4 0-2.5-.5-3-1.5M12 6v2m0 8v2"/>',
    wallet: '<rect x="2" y="6" width="20" height="13" rx="2"/><path d="M16 12.5h2M2 10h20"/>',
    plus: '<path d="M12 5v14M5 12h14"/>',
    alert: '<path d="M12 9v4m0 4h.01"/><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/>',
    info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v5m0-8h.01"/>',
    close: '<path d="M18 6 6 18M6 6l12 12"/>',
  };
  const icon = (n) => '<svg class="i" viewBox="0 0 24 24">' + ICONS[n] + '</svg>';
  const kpi = (label, value, ic, tone, foot) => '<div class="kpi"><div class="top"><span class="ic ' + tone + '">' + icon(ic) + '</span>' +
    esc(label) + '</div><div class="val">' + value + '</div>' + (foot ? '<div class="foot">' + foot + '</div>' : '') + '</div>';

  function renderKpis() {
    const sum = state.data.summary;
    const tot = (k) => sum.reduce((t, s) => t + s[k], 0);
    const sched = tot('scheduledDays');
    const rate = sched ? Math.round((tot('presentDays') / sched) * 100) : 0;
    const withSalary = state.pay.filter((p) => p.hasSalary);
    const net = withSalary.reduce((t, p) => t + p.net, 0);
    $('kpis').innerHTML = [
      kpi('Attendance rate', rate + '<small>%</small>', 'check', 't-good', int(tot('presentDays')) + ' of ' + int(sched) + ' working days'),
      kpi('Lates', tot('lateTimes'), 'clock', 't-warn', tot('chargedLates') + ' charged · ' + tot('lateRelief') + ' excused'),
      kpi('Half days', tot('halfDays'), 'half', 't-half', 'Late arrival or early exit'),
      kpi('Absent days', tot('absentDays'), 'x', 't-bad', tot('missedPunches') + ' missed punches to check'),
      kpi('Late fines', rs(tot('lateFine')), 'coin', 't-warn', 'Rs ' + prefs.settings.fine + ' per charged late'),
      kpi('Overtime', hours(tot('extraMin')) + '<small>h</small>', 'plus', 't-info',
        tot('eventMin') ? '+' + hours(tot('eventMin')) + ' h party / event, not counted' : 'No party days marked'),
      kpi('Net payroll', withSalary.length ? rs(net) : '—', 'wallet', 't-info',
        withSalary.length ? withSalary.length + ' of ' + sum.length + ' salaries entered' : 'Enter salaries in Payroll'),
    ].join('');
  }

  // ---------- Charts (Chart.js) ----------
  function baseOptions() {
    const text2 = cssVar('--text-2'), muted = cssVar('--muted');
    Chart.defaults.font.family = '"Inter", system-ui, -apple-system, "Segoe UI", sans-serif';
    Chart.defaults.color = muted;
    return {
      responsive: true, maintainAspectRatio: false, animation: { duration: 350 },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: cssVar('--surface'), titleColor: cssVar('--text'), bodyColor: text2, footerColor: muted,
          borderColor: cssVar('--line-strong'), borderWidth: 1, padding: 10, cornerRadius: 10, boxPadding: 5,
          usePointStyle: true, titleFont: { weight: '600' }, footerFont: { weight: '400' },
        },
      },
    };
  }
  const axis = (extra) => Object.assign({
    grid: { color: cssVar('--grid'), drawTicks: false }, border: { color: cssVar('--axis') },
    ticks: { color: cssVar('--muted'), padding: 6, font: { size: 11 } },
  }, extra);

  // Writes each bar's value just past its end (horizontal bar charts)
  const endLabels = {
    id: 'endLabels',
    afterDatasetsDraw(chart, args, opts) {
      if (!opts || !opts.format) return;
      const { ctx } = chart;
      const meta = chart.getDatasetMeta(chart.data.datasets.length - 1);
      ctx.save();
      ctx.font = '600 11.5px ' + Chart.defaults.font.family;
      ctx.fillStyle = cssVar('--text-2');
      ctx.textBaseline = 'middle';
      meta.data.forEach((bar, i) => {
        const v = opts.totals ? opts.totals[i] : chart.data.datasets[0].data[i];
        ctx.fillText(opts.format(v, i), bar.x + 6, bar.y);
      });
      ctx.restore();
    },
  };

  function makeChart(key, boxId, config) {
    if (charts[key]) charts[key].destroy();
    const box = $(boxId);
    box.innerHTML = '<canvas></canvas>';
    charts[key] = new Chart(box.firstChild, config);
  }
  function emptyBox(key, boxId, msg) {
    if (charts[key]) { charts[key].destroy(); delete charts[key]; }
    $(boxId).style.height = 'auto';
    $(boxId).innerHTML = '<div class="empty">' + msg + '</div>';
  }

  // ---------- Dashboard employee picker ----------
  function renderDashPicker() {
    const sel = $('dashEmp');
    const sum = state.data.summary;
    if (!sum.some((s) => s.id === state.dashEmp)) state.dashEmp = '';
    sel.innerHTML = '<option value="">All employees</option>' +
      sum.map((s) => '<option value="' + esc(s.id) + '">' + esc(s.name) + ' (ID ' + esc(s.id) + ')</option>').join('');
    sel.value = state.dashEmp;
  }
  $('dashEmp').addEventListener('change', (e) => {
    state.dashEmp = e.target.value;
    renderCharts();
    window.scrollTo({ top: 0 });
  });

  function renderCharts() {
    if (!state.data || state.tab !== 'dashboard') return;
    const one = state.data.summary.find((s) => s.id === state.dashEmp);
    $('dashAll').classList.toggle('hidden', !!one);
    $('dashOne').classList.toggle('hidden', !one);
    $('dashTitle').textContent = one ? one.name : 'Dashboard';
    $('dashSub').textContent = one
      ? 'ID ' + one.id + (one.dept ? ' · ' + one.dept : '') + ' · ' + fmtDate(state.data.period.from) + ' – ' + fmtDate(state.data.period.to)
      : state.fileName + ' · ' + state.data.summary.length + ' employees';
    if (one) { ['daily', 'rate', 'fine', 'ot', 'pay'].forEach((k) => { if (charts[k]) { charts[k].destroy(); delete charts[k]; } }); renderPerson(one); return; }
    ['pTimes', 'pExtra', 'pStatus'].forEach((k) => { if (charts[k]) { charts[k].destroy(); delete charts[k]; } });
    renderDailyChart();
    const sum = state.data.summary;

    const rate = sum.map((s) => ({ s, v: s.scheduledDays ? (s.presentDays / s.scheduledDays) * 100 : 0 })).sort((a, b) => b.v - a.v);
    hbar('rate', 'rateBox', rate.map((r) => r.s.name), rate.map((r) => Math.round(r.v)), {
      max: 100, fmt: (v) => v + '%',
      tip: (i) => [rate[i].s.presentDays + ' of ' + rate[i].s.scheduledDays + ' working days',
        rate[i].s.absentDays + ' absent · ' + rate[i].s.leaveDays + ' leave'],
    });

    $('fineSub').textContent = 'Rs ' + prefs.settings.fine + ' per late, first ' + prefs.settings.free + ' each month excused';
    const fined = sum.filter((s) => s.lateFine > 0).sort((a, b) => b.lateFine - a.lateFine);
    if (!fined.length) emptyBox('fine', 'fineBox', 'No late fines in this period.');
    else hbar('fine', 'fineBox', fined.map((s) => s.name), fined.map((s) => s.lateFine), {
      fmt: (v) => rs(v),
      tip: (i) => ['Late ' + fined[i].lateTimes + ' times · ' + fined[i].lateRelief + ' excused',
        fined[i].chargedLates + ' × Rs ' + prefs.settings.fine],
    });

    renderOtChart(sum);
    renderEventBox();

    renderPayChart();
  }

  // Party / event days: totals per day and per person, kept apart from overtime
  const minText = (m) => (m >= 60 ? hours(m) + ' h' : m + ' min');
  function renderEventBox() {
    const daily = state.data.daily;
    const dates = Object.keys(prefs.noExtra).filter((d) => daily.some((x) => x.date === d)).sort();
    if (!dates.length) {
      // Suggest the days when the most people stayed late
      const stayed = {};
      daily.forEach((d) => { if (d.extraRawMin > 0) stayed[d.date] = (stayed[d.date] || 0) + 1; });
      const top = Object.keys(stayed).sort((a, b) => stayed[b] - stayed[a]).slice(0, 3);
      $('eventBox').innerHTML = '<p class="note" style="margin:0">No party or event days are marked, so all late stays count as overtime.' +
        (top.length ? ' Days when the most people stayed late: ' + top.map((d) => '<b>' + fmtDate(d, false) + '</b> (' + stayed[d] + ' people)').join(', ') + '.' : '') + '</p>';
      return;
    }
    const rows = dates.map((date) => {
      const ppl = daily.filter((d) => d.date === date && d.extraRawMin > d.extraMin);
      const min = ppl.reduce((t, d) => t + d.extraRawMin - d.extraMin, 0);
      const ev = prefs.noExtra[date];
      return '<tr><td><b>' + esc(ev.name) + '</b><small class="sub">' + first(date).day.slice(0, 1) + first(date).day.slice(1, 3).toLowerCase() + ', ' + fmtDate(date) +
        (ev.after ? ' · counted after ' + esc(ev.after) : '') + '</small></td>' +
        '<td class="r">' + ppl.length + '</td><td class="r">' + (min ? minText(min) : '—') + '</td>' +
        '<td>' + (ppl.length ? ppl.sort((a, b) => (b.extraRawMin - b.extraMin) - (a.extraRawMin - a.extraMin))
          .map((d) => esc(d.name) + ' <span class="dim">' + minText(d.extraRawMin - d.extraMin) + '</span>').join(', ') : '<span class="dim">Nobody stayed late</span>') + '</td></tr>';
    });
    const total = state.data.summary.reduce((t, s) => t + s.eventMin, 0);
    $('eventBox').innerHTML = '<div class="table-wrap" style="max-height:none"><table><thead><tr><th>Event</th><th class="r">People</th>' +
      '<th class="r">Not counted</th><th>Who stayed</th></tr></thead><tbody>' + rows.join('') + '</tbody>' +
      '<tfoot><tr><td>Total</td><td></td><td class="r">' + minText(total) + '</td><td></td></tr></tfoot></table></div>';
  }
  const first = (date) => state.data.daily.find((d) => d.date === date);
  $('goEvents').addEventListener('click', () => showTab('settings'));

  // Overtime per employee: counted overtime, with party / event time stacked separately in grey
  function renderOtChart(sum) {
    const ot = sum.filter((s) => s.extraMin > 0 || s.eventMin > 0).sort((a, b) => b.extraMin - a.extraMin || b.eventMin - a.eventMin);
    const anyEvent = ot.some((s) => s.eventMin > 0);
    $('otLegend').innerHTML = anyEvent
      ? '<span><i style="background:var(--s1)"></i>Overtime (counted)</span><span><i style="background:var(--axis)"></i>Party / event (not counted)</span>' : '';
    if (!ot.length) { emptyBox('ot', 'otBox', 'Nobody worked overtime in this period.'); return; }
    const h = (min) => +(min / 60).toFixed(1);
    const box = $('otBox');
    box.style.height = Math.max(120, ot.length * 32 + 36) + 'px';
    const base = baseOptions();
    const surface = cssVar('--surface');
    makeChart('ot', 'otBox', {
      type: 'bar',
      data: {
        labels: ot.map((s) => s.name),
        datasets: [
          { label: 'Overtime (counted)', data: ot.map((s) => h(s.extraMin)), backgroundColor: cssVar('--s1') },
          { label: 'Party / event (not counted)', data: ot.map((s) => h(s.eventMin)), backgroundColor: cssVar('--axis') },
        ].map((d) => Object.assign(d, { borderColor: surface, borderWidth: { right: 2 }, borderSkipped: false, borderRadius: 3, maxBarThickness: 16 })),
      },
      options: Object.assign(base, {
        indexAxis: 'y',
        layout: { padding: { right: 110 } },
        interaction: { mode: 'index', intersect: false, axis: 'y' },
        scales: {
          x: axis({ stacked: true, beginAtZero: true, ticks: { display: false } }),
          y: axis({ stacked: true, grid: { display: false }, ticks: { color: cssVar('--text'), font: { size: 12.5 }, autoSkip: false } }),
        },
        plugins: Object.assign(base.plugins, {
          // Label shows counted hours, plus party hours in brackets when there are any
          endLabels: { format: (v, i) => h(ot[i].extraMin) + ' h' + (ot[i].eventMin ? ' (+' + h(ot[i].eventMin) + ' party)' : ''), totals: ot.map((s) => s.extraMin) },
          tooltip: Object.assign(base.plugins.tooltip, {
            filter: (it) => it.raw > 0,
            callbacks: {
              label: (it) => ' ' + it.dataset.label + ': ' + it.raw + ' h',
              footer: (items) => {
                const s = ot[items[0].dataIndex];
                const pay = payOf(s.id).otPay;
                return [s.extraDays + ' overtime days · ' + s.extraMin + ' min',
                  s.eventMin ? s.eventDays + ' party / event days · ' + s.eventMin + ' min' : '',
                  pay ? 'Overtime pay ' + rs(pay) : ''].filter(Boolean);
              },
            },
          }),
        }),
      }),
      plugins: [endLabels],
    });
  }

  function renderDailyChart() {
    const dates = [...new Set(state.data.daily.map((d) => d.date))];
    const byDate = {};
    state.data.daily.forEach((d) => (byDate[d.date] = byDate[d.date] || []).push(d));
    const surface = cssVar('--surface');
    $('dailyLegend').innerHTML = legendHtml(GROUPS);
    makeChart('daily', 'dailyBox', {
      type: 'bar',
      data: {
        labels: dates.map((d) => [String(Number(d.slice(8))), byDate[d][0].day.slice(0, 2)]),
        datasets: GROUPS.map((g) => ({
          label: g.label,
          data: dates.map((d) => byDate[d].filter((r) => groupOf(r.status) === g).length),
          backgroundColor: cssVar(g.color), borderColor: surface, borderWidth: { top: 2 }, borderSkipped: false,
          borderRadius: 3, maxBarThickness: 26, categoryPercentage: 0.8, barPercentage: 0.9,
        })),
      },
      options: Object.assign(baseOptions(), {
        interaction: { mode: 'index', intersect: false },
        scales: {
          x: axis({ stacked: true, grid: { display: false } }),
          y: axis({ stacked: true, beginAtZero: true, suggestedMax: state.data.summary.length, ticks: { precision: 0, color: cssVar('--muted'), font: { size: 11 } } }),
        },
        plugins: Object.assign(baseOptions().plugins, {
          tooltip: Object.assign(baseOptions().plugins.tooltip, {
            filter: (item) => item.raw > 0,
            callbacks: {
              title: (items) => { const d = dates[items[0].dataIndex]; return byDate[d][0].day + ', ' + fmtDate(d); },
              footer: (items) => {
                const d = dates[items[0].dataIndex];
                const hol = byDate[d][0].holiday;
                return hol ? 'Holiday: ' + hol : items.length ? '' : 'Holiday';
              },
            },
          }),
        }),
      }),
    });
  }

  function hbar(key, boxId, labels, values, o) {
    const box = $(boxId);
    box.style.height = Math.max(120, labels.length * 32 + 36) + 'px';
    const base = baseOptions();
    makeChart(key, boxId, {
      type: 'bar',
      data: { labels, datasets: [{ data: values, backgroundColor: cssVar('--s1'), borderRadius: 4, borderSkipped: 'start', maxBarThickness: 16 }] },
      options: Object.assign(base, {
        indexAxis: 'y',
        layout: { padding: { right: 64 } },
        scales: {
          x: axis({ beginAtZero: true, max: o.max, ticks: { display: false }, grid: { color: cssVar('--grid'), drawTicks: false } }),
          y: axis({ grid: { display: false }, ticks: { color: cssVar('--text'), font: { size: 12.5 }, autoSkip: false } }),
        },
        plugins: Object.assign(base.plugins, {
          endLabels: { format: o.fmt },
          tooltip: Object.assign(base.plugins.tooltip, {
            displayColors: false,
            callbacks: { label: (it) => o.fmt(it.raw), afterLabel: (it) => (o.tip ? o.tip(it.dataIndex).filter(Boolean) : []) },
          }),
        }),
      }),
      plugins: [endLabels],
    });
  }

  function renderPayChart() {
    const rows = state.pay.filter((p) => p.hasSalary).sort((a, b) => b.salary - a.salary);
    const series = [
      { label: 'Net pay', color: '--s1', cls: 's1', get: (p) => p.net - p.otPay },
      { label: 'Unpaid days off', color: '--s2', cls: 's2', get: (p) => p.absentDed },
      { label: 'Half days', color: '--s3', cls: 's3', get: (p) => p.halfDed },
      { label: 'Late fines', color: '--s4', cls: 's4', get: (p) => p.lateFine },
    ];
    $('payLegend').innerHTML = rows.length
      ? series.map((s) => '<span><i style="background:var(' + s.color + ')"></i>' + s.label + '</span>').join('') : '';
    if (!rows.length) {
      emptyBox('pay', 'payBox', 'Enter salaries in the <b>Payroll</b> tab to see each person\'s pay here.');
      return;
    }
    const box = $('payBox');
    box.style.height = Math.max(120, rows.length * 32 + 36) + 'px';
    const base = baseOptions();
    const surface = cssVar('--surface');
    makeChart('pay', 'payBox', {
      type: 'bar',
      data: {
        labels: rows.map((p) => p.name),
        datasets: series.map((s) => ({
          label: s.label, data: rows.map((p) => Math.max(0, s.get(p))), backgroundColor: cssVar(s.color),
          borderColor: surface, borderWidth: { right: 2 }, borderSkipped: false, borderRadius: 3, maxBarThickness: 16,
        })),
      },
      options: Object.assign(base, {
        indexAxis: 'y',
        layout: { padding: { right: 84 } },
        interaction: { mode: 'index', intersect: false, axis: 'y' },
        scales: {
          x: axis({ stacked: true, beginAtZero: true, ticks: { display: false } }),
          y: axis({ stacked: true, grid: { display: false }, ticks: { color: cssVar('--text'), font: { size: 12.5 }, autoSkip: false } }),
        },
        plugins: Object.assign(base.plugins, {
          endLabels: { format: (v) => rs(v), totals: rows.map((p) => p.net) },
          tooltip: Object.assign(base.plugins.tooltip, {
            filter: (it) => it.raw > 0,
            callbacks: {
              label: (it) => ' ' + it.dataset.label + ': ' + rs(it.raw),
              footer: (items) => { const p = rows[items[0].dataIndex]; return 'Salary ' + rs(p.salary) + (p.otPay ? ' · Overtime +' + rs(p.otPay) : '') + ' · Net ' + rs(p.net); },
            },
          }),
        }),
      }),
      plugins: [endLabels],
    });
  }

  // ---------- One employee's dashboard ----------
  const clock = (h) => { const hh = Math.floor(h), mm = Math.round((h - hh) * 60); const p = hh >= 12 ? 'PM' : 'AM'; return ((hh + 11) % 12 + 1) + ':' + String(mm).padStart(2, '0') + ' ' + p; };
  const toHours = (t) => { const [h, m] = t.split(':').map(Number); return h + m / 60; };

  // Dashed reference lines across a chart at given y values
  const refLines = {
    id: 'refLines',
    afterDatasetsDraw(chart, args, opts) {
      if (!opts || !opts.lines) return;
      const { ctx, chartArea: a, scales: { y } } = chart;
      ctx.save();
      ctx.font = '500 11px ' + Chart.defaults.font.family;
      opts.lines.forEach((l) => {
        const py = y.getPixelForValue(l.v);
        ctx.strokeStyle = cssVar('--muted'); ctx.setLineDash([4, 4]); ctx.lineWidth = 1;
        ctx.beginPath(); ctx.moveTo(a.left, py); ctx.lineTo(a.right, py); ctx.stroke();
        ctx.setLineDash([]); ctx.fillStyle = cssVar('--muted'); ctx.textAlign = 'right';
        ctx.fillText(l.label, a.right - 4, py - 5);
      });
      ctx.restore();
    },
  };

  function renderPerson(s) {
    const p = payOf(s.id);
    const days = state.data.daily.filter((d) => d.id === s.id);
    const notCounted = days.reduce((t, d) => t + d.extraRawMin - d.extraMin, 0);
    const rate = s.scheduledDays ? Math.round((s.presentDays / s.scheduledDays) * 100) : 0;
    const card = (title, sub, body) => '<div class="card"><div class="card-head"><div><h2>' + title + '</h2>' + (sub ? '<p>' + sub + '</p>' : '') + '</div></div>' + body + '</div>';

    $('dashOne').innerHTML =
      '<div class="grid kpis">' +
        kpi('Attendance', rate + '<small>%</small>', 'check', 't-good', s.presentDays + ' of ' + s.scheduledDays + ' working days') +
        kpi('Lates', s.lateTimes, 'clock', 't-warn', s.chargedLates + ' charged · ' + rs(s.lateFine)) +
        kpi('Half days', s.halfDays, 'half', 't-half', p.halfDed ? '−' + rs(p.halfDed) : '') +
        kpi('Absent / leave', s.absentDays + s.leaveDays, 'x', 't-bad', s.absentDays + ' absent · ' + s.leaveDays + ' leave') +
        kpi('Extra work', hours(s.extraMin) + '<small>h</small>', 'plus', 't-info', s.extraDays + ' days' + (notCounted ? ' · ' + hours(notCounted) + ' h not counted' : '')) +
        kpi('Net pay', p.hasSalary ? rs(p.net) : '—', 'wallet', 't-info', p.hasSalary ? 'Salary ' + rs(p.salary) : 'Enter salary in Payroll') +
      '</div>' +
      card('In and out times', 'Each bar runs from check-in to check-out, coloured by status',
        '<div class="legend">' + legendHtml(GROUPS) + '</div><div class="chart-box" id="pTimesBox" style="height:300px"></div>') +
      '<div class="grid two">' +
        card('Extra work by day', 'Minutes after shift end' + (notCounted ? ' · grey = party / event, not counted' : ''),
          '<div class="chart-box" id="pExtraBox" style="height:240px"></div>') +
        card('Days by status', 'How each working day went', '<div class="chart-box" id="pStatusBox"></div>') +
      '</div>' +
      '<div class="grid two">' +
        card('Salary', 'For this period', breakdownHtml(p)) +
        card('Which days', 'When each late, half day and extra work happened', whichHtml(days)) +
      '</div>' +
      card('Day by day', '', dayTableHtml(days));

    // In / out floating bars
    const labels = days.map((d) => [String(Number(d.date.slice(8))), d.day.slice(0, 2)]);
    const span = days.map((d) => {
      if (!d.inTime && !d.outTime) return null;
      const a = d.inTime ? toHours(d.inTime) : toHours(d.outTime) - 0.1;
      const b = d.outTime && d.outTime !== d.inTime ? toHours(d.outTime) : a + 0.1;
      return [a, b];
    });
    const all = span.filter(Boolean).flat();
    const base = baseOptions();
    makeChart('pTimes', 'pTimesBox', {
      type: 'bar',
      data: { labels, datasets: [{
        data: span,
        backgroundColor: days.map((d) => { const g = groupOf(d.status); return cssVar(g ? g.color : '--axis'); }),
        borderRadius: 4, borderSkipped: false, maxBarThickness: 22,
      }] },
      options: Object.assign(base, {
        scales: {
          x: axis({ grid: { display: false } }),
          y: axis({
            min: Math.floor(Math.min(10, ...all)), max: Math.ceil(Math.max(19, ...all)),
            ticks: { stepSize: 1, color: cssVar('--muted'), font: { size: 11 }, callback: (v) => clock(v) },
          }),
        },
        plugins: Object.assign(base.plugins, {
          refLines: { lines: [{ v: 10.75, label: 'Entry 10:45 AM' }, { v: 18.5, label: 'Exit 6:30 PM (Sat 4:30)' }] },
          tooltip: Object.assign(base.plugins.tooltip, {
            displayColors: false,
            callbacks: {
              title: (it) => { const d = days[it[0].dataIndex]; return d.day + ', ' + fmtDate(d.date); },
              label: (it) => { const d = days[it.dataIndex]; return [d.inTime && 'In ' + d.inTime, d.outTime && 'Out ' + d.outTime].filter(Boolean).join(' · '); },
              afterLabel: (it) => { const d = days[it.dataIndex]; return [d.status, d.lateMin && d.lateMin + ' min late', d.extraMin && d.extraMin + ' min extra',
                d.extraRawMin > d.extraMin && d.extraNote + ': ' + (d.extraRawMin - d.extraMin) + ' min not counted'].filter(Boolean); },
            },
          }),
        }),
      }),
      plugins: [refLines],
    });

    // Extra work: counted vs not counted (party / event days)
    const b2 = baseOptions();
    makeChart('pExtra', 'pExtraBox', {
      type: 'bar',
      data: { labels, datasets: [
        { label: 'Counted', data: days.map((d) => d.extraMin), backgroundColor: cssVar('--s1'), borderRadius: 3, maxBarThickness: 18 },
        { label: 'Not counted (event)', data: days.map((d) => d.extraRawMin - d.extraMin), backgroundColor: cssVar('--axis'), borderRadius: 3, maxBarThickness: 18 },
      ] },
      options: Object.assign(b2, {
        interaction: { mode: 'index', intersect: false },
        scales: { x: axis({ stacked: true, grid: { display: false } }), y: axis({ stacked: true, beginAtZero: true, ticks: { precision: 0, color: cssVar('--muted'), font: { size: 11 }, callback: (v) => v + 'm' } }) },
        plugins: Object.assign(b2.plugins, {
          tooltip: Object.assign(b2.plugins.tooltip, {
            filter: (it) => it.raw > 0,
            callbacks: {
              title: (it) => { const d = days[it[0].dataIndex]; return d.day + ', ' + fmtDate(d.date) + (d.outTime ? ' · out ' + d.outTime : ''); },
              label: (it) => ' ' + it.dataset.label + ': ' + it.raw + ' min' + (it.datasetIndex === 1 ? ' (' + days[it.dataIndex].extraNote + ')' : ''),
            },
          }),
        }),
      }),
    });

    // Days by status
    const counts = GROUPS.map((g) => days.filter((d) => groupOf(d.status) === g).length);
    const shown = GROUPS.map((g, i) => ({ g, n: counts[i] })).filter((x) => x.n > 0);
    if (!shown.length) { emptyBox('pStatus', 'pStatusBox', 'No working days in this period.'); return; }
    hbar('pStatus', 'pStatusBox', shown.map((x) => x.g.label), shown.map((x) => x.n), { fmt: (v) => plural(v, 'day') });
    charts.pStatus.data.datasets[0].backgroundColor = shown.map((x) => cssVar(x.g.color));
    charts.pStatus.update('none');
  }

  // ---------- Calendar heatmap ----------
  function renderCalendar() {
    const dates = [...new Set(state.data.daily.map((d) => d.date))];
    const byKey = {};
    state.data.daily.forEach((d) => (byKey[d.id + '|' + d.date] = d));
    const first = {};
    state.data.daily.forEach((d) => { if (!first[d.date]) first[d.date] = d; });
    const cal = $('calendar');
    cal.style.gridTemplateColumns = '130px repeat(' + dates.length + ', minmax(28px, 1fr))';
    let html = '<span></span>' + dates.map((date) => {
      const hol = holidayOf(date);
      const hint = autoHoliday(date) && !prefs.holidays[date] ? 'Every Sunday is a holiday' : hol ? 'Click to remove' : 'Click to mark as holiday';
      return '<button class="hd' + (hol ? ' hol' : '') + '" data-date="' + date + '" data-tip="' +
        esc('<b>' + first[date].day + ', ' + fmtDate(date) + '</b><br>' + (hol ? 'Holiday: ' + esc(hol) + '<br>' : '') + hint) + '">' +
        Number(date.slice(8)) + '<small>' + first[date].day.slice(0, 2) + '</small></button>';
    }).join('');
    state.data.summary.forEach((s) => {
      html += '<button class="nm" data-emp="' + esc(s.id) + '" title="' + esc(s.name) + '">' + esc(s.name) + '</button>';
      dates.forEach((date) => {
        const d = byKey[s.id + '|' + date];
        if (!d) { html += '<span class="c c-off"></span>'; return; }
        const g = groupOf(d.status);
        const times = [d.inTime && 'In ' + d.inTime, d.outTime && 'Out ' + d.outTime].filter(Boolean).join(' · ');
        const mins = [d.lateMin && d.lateMin + ' min late', d.earlyMin && d.earlyMin + ' min before shift end',
          d.extraMin && d.extraMin + ' min overtime',
          d.extraNote && d.extraRawMin > d.extraMin && esc(d.extraNote) + ': ' + (d.extraRawMin - d.extraMin) + ' min not counted'].filter(Boolean).join('<br>');
        const tip = '<b>' + esc(s.name) + '</b><br>' + d.day + ', ' + fmtDate(date) + '<br>' +
          (g ? '<i class="' + g.cls + '"></i>' : '') + esc(d.holiday ? d.status + ' (' + d.holiday + ')' : d.status) +
          (times ? '<br>' + times : '') + (mins ? '<br>' + mins : '');
        html += '<span class="c ' + (g ? g.cls : 'c-off') + '" data-tip="' + esc(tip) + '">' + shortOf(d.status) + '</span>';
      });
    });
    cal.innerHTML = html;
    $('calLegend').innerHTML = legendHtml(GROUPS.concat({ label: 'Holiday / off', cls: 'c-off' }));
  }

  $('calendar').addEventListener('click', (e) => {
    const hd = e.target.closest('[data-date]');
    if (hd) { toggleHoliday(hd.dataset.date); return; }
    const nm = e.target.closest('[data-emp]');
    if (nm) openDetail(nm.dataset.emp);
  });

  // Sundays count as holidays automatically while the Sunday setting is on
  const isSunday = (date) => new Date(date + 'T00:00:00').getDay() === 0;
  const autoHoliday = (date) => (prefs.settings.sundayOff !== false && isSunday(date) ? 'Sunday' : '');
  const holidayOf = (date) => prefs.holidays[date] || autoHoliday(date);

  function toggleHoliday(date) {
    if (!prefs.holidays[date] && autoHoliday(date)) {
      toast('Sundays are holidays automatically. Turn this off in Holidays & Rules.');
      return;
    }
    if (prefs.holidays[date]) {
      delete prefs.holidays[date];
      toast(fmtDate(date) + ' is a working day again');
    } else {
      prefs.holidays[date] = 'Holiday';
      toast(fmtDate(date) + ' marked as holiday');
    }
    persist();
    reparse();
  }

  // ---------- Daily table ----------
  function renderEmpFilter() {
    const sel = $('empFilter');
    const cur = sel.value;
    sel.innerHTML = '<option value="">All employees</option>' +
      state.data.summary.map((s) => '<option value="' + esc(s.id) + '">' + esc(s.name) + '</option>').join('');
    sel.value = state.data.summary.some((s) => s.id === cur) ? cur : '';
  }
  ['q', 'empFilter', 'statusFilter'].forEach((id) => $(id).addEventListener('input', renderDailyTable));

  function renderDailyTable() {
    const q = $('q').value.trim().toLowerCase();
    const emp = $('empFilter').value;
    const st = $('statusFilter').value;
    const rows = state.data.daily.filter((d) => {
      if (emp && d.id !== emp) return false;
      if (q && !(d.name.toLowerCase().includes(q) || d.id.includes(q))) return false;
      if (!st) return true;
      if (st === 'ok') return d.status === 'Present' || d.status === 'Present (Holiday)';
      if (st === 'leave') return d.status === 'Leave' || d.status === 'Out';
      if (st === 'overtime') return d.extraMin > 0;
      return has(d.status, st);
    });
    $('dailyCount').textContent = rows.length + ' of ' + state.data.daily.length + ' records';
    const dash = '<span style="color:var(--muted)">—</span>';
    $('dailyTable').innerHTML = '<thead><tr><th>Employee</th><th>Date</th><th>In</th><th>Out</th><th>Status</th>' +
      '<th class="r">Late (min)</th><th class="r">Before end (min)</th><th class="r">Overtime (min)</th></tr></thead><tbody>' +
      (rows.length ? rows.map((d) => '<tr><td><button class="link" data-emp="' + esc(d.id) + '">' + esc(d.name) + '</button></td>' +
        '<td>' + d.day.slice(0, 1) + d.day.slice(1).toLowerCase() + ', ' + fmtDate(d.date, false) + '</td>' +
        '<td class="num">' + (d.inTime || dash) + '</td><td class="num">' + (d.outTime || dash) + '</td>' +
        '<td>' + pill(d.status) + (d.holiday ? ' <span class="pill t-off">' + esc(d.holiday) + '</span>' : '') + '</td>' +
        '<td class="r">' + (d.lateMin || dash) + '</td><td class="r">' + (d.earlyMin || dash) + '</td>' +
        '<td class="r">' + (d.extraMin || dash) +
          (d.extraNote && d.extraRawMin > d.extraMin ? '<small class="sub">' + esc(d.extraNote) + ', ' + (d.extraRawMin - d.extraMin) + ' min not counted</small>' : '') + '</td></tr>').join('')
        : '<tr><td colspan="8" class="empty">No records match these filters.</td></tr>') + '</tbody>';
  }
  $('dailyTable').addEventListener('click', (e) => { const b = e.target.closest('[data-emp]'); if (b) openDetail(b.dataset.emp); });

  // ---------- Payroll table ----------
  // Amount on top, what it is for underneath (one cell per deduction keeps Net pay on screen)
  function amt(n, sign, note) {
    const cls = !n ? 'dim' : sign === '+' ? 'pos' : 'neg';
    return '<td class="r"><div class="' + cls + '">' + (n ? sign + int(n) : '—') + '</div>' + (note ? '<small class="sub">' + note + '</small>' : '') + '</td>';
  }
  const plural = (n, w) => n + ' ' + w + (n === 1 ? '' : 's');

  function payCells(p) {
    return [
      '<td class="r dim">' + (p.hasSalary ? int(p.dayRate) : '—') + '</td>',
      '<td class="r">' + p.presentDays + ' / ' + p.workingDays + '</td>',
      '<td class="r">' + (p.offDays ? plural(p.offDays, 'day') : '<span class="dim">—</span>') +
        (p.offDays ? '<small class="sub">' + p.paidLeave + ' paid · ' + p.unpaidDays + ' unpaid</small>' : '') + '</td>',
      amt(p.absentDed, '−', p.unpaidDays ? plural(p.unpaidDays, 'day') : ''),
      amt(p.halfDed, '−', p.halfDays ? plural(p.halfDays, 'half day') : ''),
      amt(p.lateFine, '−', p.lateTimes ? p.chargedLates + ' of ' + plural(p.lateTimes, 'late') : ''),
      amt(p.otPay, '+', p.otMin ? hours(p.otMin) + ' h' : ''),
      '<td class="r net">' + (p.hasSalary ? rs(p.net) : '<span style="color:var(--muted);font-weight:500">Enter salary</span>') + '</td>',
    ];
  }

  function payFoot() {
    const t = (k) => state.pay.reduce((a, p) => a + p[k], 0);
    const salaried = state.pay.filter((p) => p.hasSalary);
    return '<tr><td>Total</td><td class="r">' + rs(t('salary')) + '</td><td></td><td></td>' +
      '<td class="r">' + plural(t('offDays'), 'day') + '<small class="sub">' + t('paidLeave') + ' paid</small></td>' +
      amt(t('absentDed'), '−', plural(t('unpaidDays'), 'day')) + amt(t('halfDed'), '−', plural(t('halfDays'), 'half day')) +
      amt(t('lateFine'), '−', t('chargedLates') + ' charged') + amt(t('otPay'), '+', hours(t('otMin')) + ' h') +
      '<td class="r net">' + rs(salaried.reduce((a, p) => a + p.net, 0)) + '</td></tr>';
  }

  function renderPayKpis() {
    const t = (k) => state.pay.reduce((a, p) => a + p[k], 0);
    const salaried = state.pay.filter((p) => p.hasSalary);
    $('payKpis').innerHTML = [
      kpi('Total salaries', rs(t('salary')), 'wallet', 't-info', salaried.length + ' of ' + state.pay.length + ' entered'),
      kpi('Deductions', rs(t('absentDed') + t('halfDed')), 'x', 't-bad', t('unpaidDays') + ' unpaid days off, ' + t('halfDays') + ' half days'),
      kpi('Late fines', rs(t('lateFine')), 'clock', 't-warn', t('chargedLates') + ' charged lates'),
      kpi('Overtime pay', rs(t('otPay')), 'plus', 't-good', hours(t('otMin')) + ' hours worked'),
      kpi('Net payable', rs(salaried.reduce((a, p) => a + p.net, 0)), 'coin', 't-info', 'For this period'),
    ].join('');
  }

  function renderPayroll() {
    const p = state.data.period;
    const s = prefs.settings;
    $('payBanner').innerHTML = isFullMonth(p) ? '' :
      '<div class="banner warn">' + icon('alert') + '<div>This report covers <b>' + fmtDate(p.from) + ' – ' + fmtDate(p.to) +
      '</b>, not a full month. Net pay below is the full monthly salary minus this period\'s deductions only.</div></div>';
    const basisText = s.basis === 'calendar' ? 'days in the month' : s.basis + ' days';
    $('payRuleNote').textContent = 'Per day = salary ÷ ' + basisText + ' · half day = ' +
      (s.half === '0.5' ? 'half a day' : s.half === '1' ? 'a full day' : 'no deduction') +
      ' · late = Rs ' + s.fine + ' after ' + s.free + ' free each month · ' + s.paidLeaves + ' paid days off each month';
    $('payTable').innerHTML = '<thead><tr><th>Employee</th><th class="r">Monthly salary (Rs)</th><th class="r">Per day</th>' +
      '<th class="r">Present</th><th class="r">Days off</th><th class="r">Unpaid days</th><th class="r">Half days</th>' +
      '<th class="r">Late fines</th><th class="r">Overtime</th><th class="r">Net pay</th></tr></thead><tbody>' +
      state.pay.map((p) => '<tr data-id="' + esc(p.id) + '"><td><div class="who"><span class="avatar">' + esc(initials(p.name)) + '</span>' +
        '<div><button class="link" data-emp="' + esc(p.id) + '">' + esc(p.name) + '</button><small>ID ' + esc(p.id) + (p.dept ? ' · ' + esc(p.dept) : '') + '</small></div></div></td>' +
        '<td class="r"><input class="money-in" inputmode="numeric" data-salary="' + esc(p.id) + '" value="' + (p.salary || '') + '" placeholder="0" aria-label="Monthly salary for ' + esc(p.name) + '"></td>' +
        payCells(p).join('') + '</tr>').join('') +
      '</tbody><tfoot>' + payFoot() + '</tfoot>';
    renderPayKpis();
  }

  function setSalary(id, raw) {
    const v = Math.max(0, Number(String(raw).replace(/[^\d.]/g, '')) || 0);
    if (v) prefs.salaries[id] = v; else delete prefs.salaries[id];
    persist();
    computePay();
  }

  $('payTable').addEventListener('input', (e) => {
    const inp = e.target.closest('[data-salary]');
    if (!inp) return;
    setSalary(inp.dataset.salary, inp.value);
    const row = inp.closest('tr');
    const tmp = document.createElement('tbody');
    tmp.innerHTML = '<tr><td></td><td></td>' + payCells(payOf(inp.dataset.salary)).join('') + '</tr>';
    const fresh = tmp.firstChild.cells;
    for (let i = row.cells.length - 1; i >= 2; i--) row.replaceChild(fresh[i], row.cells[i]);
    $('payTable').tFoot.innerHTML = payFoot();
    renderPayKpis();
    renderKpis();
  });
  $('payTable').addEventListener('click', (e) => { const b = e.target.closest('[data-emp]'); if (b) openDetail(b.dataset.emp); });
  $('printBtn').addEventListener('click', () => window.print());

  // ---------- Excel export ----------
  $('exportBtn').addEventListener('click', () => {
    const p = state.data.period;
    const wb = XLSX.utils.book_new();
    const payroll = state.pay.map((x) => ({
      'ID': x.id, 'Name': x.name, 'Department': x.dept, 'Monthly Salary': x.salary, 'Per Day': Math.round(x.dayRate),
      'Working Days': x.workingDays, 'Present Days': x.presentDays, 'Absent Days': x.absentDays, 'Leave Days': x.leaveDays,
      'Paid Leaves Used': x.paidLeave, 'Unpaid Days': x.unpaidDays, 'Unpaid Days Deduction': x.absentDed,
      'Half Days': x.halfDays, 'Half Day Deduction': x.halfDed, 'Lates': x.lateTimes, 'Lates Excused': x.lateRelief,
      'Lates Charged': x.chargedLates, 'Late Fine': x.lateFine,
      'Overtime (Hours)': +(x.otMin / 60).toFixed(2), 'Overtime Pay': x.otPay, 'Total Deductions': x.deductions, 'Net Pay': x.net,
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(payroll), 'Payroll');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.data.daily.map((d) => ({
      'ID': d.id, 'Name': d.name, 'Date': d.date, 'Day': d.day, 'In': d.inTime, 'Out': d.outTime, 'All Punches': d.punches,
      'Status': d.status, 'Holiday': d.holiday, 'Late (Min)': d.lateMin, 'Before End (Min)': d.earlyMin, 'Overtime (Min)': d.extraMin,
      'Event (Extra Not Counted)': d.extraNote,
    }))), 'Daily Attendance');
    const ev = Object.keys(prefs.noExtra).sort().map((d) => ({ 'Date': d, 'Event': prefs.noExtra[d].name, 'Count Extra After': prefs.noExtra[d].after || 'Not counted' }));
    if (ev.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ev), 'Event Days');
    const hol = Object.keys(prefs.holidays).sort().map((d) => ({ 'Date': d, 'Name': prefs.holidays[d] }));
    if (hol.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(hol), 'Holidays');
    XLSX.writeFile(wb, 'Payroll ' + p.from + ' to ' + p.to + '.xlsx');
    toast('Excel file downloaded');
  });

  // ---------- Holidays & rules ----------
  function renderSettings() {
    const daily = state.data.daily;
    const dates = [...new Set(daily.map((d) => d.date))];
    const info = {};
    daily.forEach((d) => {
      const i = info[d.date] || (info[d.date] = { day: d.day, machineOff: true });
      if (!(d.status === 'Holiday' || d.status === 'Present (Holiday)') || d.holiday) i.machineOff = false;
    });
    const WD = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN'];
    const lead = WD.indexOf(info[dates[0]].day);
    let html = WD.map((w) => '<div class="wd">' + w.slice(0, 1) + w.slice(1).toLowerCase() + '</div>').join('');
    for (let i = 0; i < lead; i++) html += '<div class="day blank"></div>';
    dates.forEach((date) => {
      const hol = holidayOf(date);
      const auto = !prefs.holidays[date] && autoHoliday(date);
      const off = info[date].machineOff;
      html += '<button class="day' + (hol ? ' hol' : off ? ' sched-off' : '') + '" data-date="' + date + '" title="' +
        (auto ? 'Every Sunday is a holiday' : hol ? 'Holiday: ' + esc(hol) + ' (click to remove)' : off ? 'Off in machine schedule' : 'Click to mark as holiday') + '">' +
        '<b>' + Number(date.slice(8)) + '</b><small>' + (hol ? esc(hol) : off ? 'Off' : '') + '</small></button>';
    });
    $('dayPicker').innerHTML = html;

    const list = Object.keys(prefs.holidays).sort();
    $('holList').innerHTML = list.length
      ? list.map((d) => '<div class="hol-row"><span class="date">' + fmtDate(d) + '</span>' +
        '<input class="input" data-hol="' + d + '" value="' + esc(prefs.holidays[d]) + '" aria-label="Holiday name for ' + fmtDate(d) + '">' +
        '<button class="btn ghost icon sm" data-remove="' + d + '" title="Remove" aria-label="Remove holiday">' + icon('close') + '</button></div>').join('')
      : '<p class="note" style="margin:0">No holidays added yet.</p>';

    // Party / event days: staying late is not counted as extra work
    let ev = WD.map((w) => '<div class="wd">' + w.slice(0, 1) + w.slice(1).toLowerCase() + '</div>').join('');
    for (let i = 0; i < lead; i++) ev += '<div class="day blank"></div>';
    dates.forEach((date) => {
      const e = prefs.noExtra[date];
      const stayed = daily.filter((d) => d.date === date && d.extraRawMin > 0).length;
      ev += '<button class="day' + (e ? ' evt' : '') + '" data-evt="' + date + '" title="' +
        (e ? esc(e.name) + ' (click to remove)' : 'Click to mark as a party / event day' + (stayed ? ' · ' + stayed + ' stayed late' : '')) + '">' +
        '<b>' + Number(date.slice(8)) + '</b><small>' + (e ? esc(e.name) : stayed ? '+' + stayed : '') + '</small></button>';
    });
    $('evtPicker').innerHTML = ev;
    const evList = Object.keys(prefs.noExtra).filter((d) => dates.includes(d)).sort();
    $('evtList').innerHTML = evList.length
      ? evList.map((d) => {
        const removed = daily.filter((x) => x.date === d).reduce((t, x) => t + x.extraRawMin - x.extraMin, 0);
        return '<div class="hol-row evt-row"><span class="date">' + fmtDate(d) + '<small class="sub">' + (removed ? hours(removed) + ' h not counted' : 'Nothing removed') + '</small></span>' +
          '<input class="input" data-evt-name="' + d + '" value="' + esc(prefs.noExtra[d].name) + '" aria-label="Event name for ' + fmtDate(d) + '">' +
          '<label class="evt-after"><span>Count extra after</span><input class="input" type="time" data-evt-after="' + d + '" value="' + esc(prefs.noExtra[d].after || '') + '" aria-label="Count extra work after this time on ' + fmtDate(d) + '"></label>' +
          '<button class="btn ghost icon sm" data-evt-remove="' + d + '" title="Remove" aria-label="Remove event day">' + icon('close') + '</button></div>';
      }).join('')
      : '<p class="note" style="margin:0">No event days in this period.</p>';

    const s = prefs.settings;
    $('r-basis').value = s.basis; $('r-half').value = s.half; $('r-fine').value = s.fine; $('r-free').value = s.free;
    $('r-ot').value = s.ot; $('r-otRate').value = s.otRate; $('r-paidLeaves').value = s.paidLeaves;
    $('r-ignoreOff').checked = s.ignoreOff === true;
    $('r-sundayOff').checked = s.sundayOff !== false;
    $('r-otRateWrap').classList.toggle('hidden', s.ot !== 'fixed');
  }

  $('dayPicker').addEventListener('click', (e) => { const b = e.target.closest('[data-date]'); if (b) toggleHoliday(b.dataset.date); });
  $('holList').addEventListener('change', (e) => {
    const inp = e.target.closest('[data-hol]');
    if (!inp) return;
    prefs.holidays[inp.dataset.hol] = inp.value.trim() || 'Holiday';
    persist(); reparse();
  });
  $('holList').addEventListener('click', (e) => {
    const b = e.target.closest('[data-remove]');
    if (b) toggleHoliday(b.dataset.remove);
  });

  function toggleEvent(date) {
    if (prefs.noExtra[date]) {
      delete prefs.noExtra[date];
      toast('Extra work on ' + fmtDate(date) + ' counts again');
    } else {
      prefs.noExtra[date] = { name: 'Party', after: '' };
      toast('Extra work on ' + fmtDate(date) + ' will not be counted');
    }
    persist(); reparse();
  }
  $('evtPicker').addEventListener('click', (e) => { const b = e.target.closest('[data-evt]'); if (b) toggleEvent(b.dataset.evt); });
  $('evtList').addEventListener('click', (e) => { const b = e.target.closest('[data-evt-remove]'); if (b) toggleEvent(b.dataset.evtRemove); });
  $('evtList').addEventListener('change', (e) => {
    const n = e.target.closest('[data-evt-name]');
    const a = e.target.closest('[data-evt-after]');
    if (n) prefs.noExtra[n.dataset.evtName].name = n.value.trim() || 'Party';
    else if (a) prefs.noExtra[a.dataset.evtAfter].after = a.value;
    else return;
    persist(); reparse();
    if (a) toast(a.value ? 'Only extra work after ' + a.value + ' counts' : 'No extra work counted that day');
  });

  const RULES = { 'r-basis': 'basis', 'r-half': 'half', 'r-fine': 'fine', 'r-free': 'free', 'r-ot': 'ot', 'r-otRate': 'otRate', 'r-paidLeaves': 'paidLeaves', 'r-ignoreOff': 'ignoreOff', 'r-sundayOff': 'sundayOff' };
  Object.keys(RULES).forEach((id) => $(id).addEventListener('change', (e) => {
    prefs.settings[RULES[id]] = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
    persist();
    reparse();
    toast('Rules updated');
  }));

  // ---------- Employee detail drawer ----------
  let openId = null;
  function openDetail(id) {
    const s = state.data.summary.find((x) => x.id === id);
    if (!s) { closeDetail(); return; }
    openId = id;
    const p = payOf(id);
    const days = state.data.daily.filter((d) => d.id === id);
    const WD = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN'];
    const lead = days.length ? WD.indexOf(days[0].day) : 0;
    const strip = WD.map((w) => '<div class="wd" style="font-size:11px;color:var(--muted);text-align:center">' + w.slice(0, 2) + '</div>').join('') +
      '<div></div>'.repeat(lead) + days.map((d) => {
        const g = groupOf(d.status);
        return '<div class="c ' + (g ? g.cls : 'c-off') + '" title="' + esc(fmtDate(d.date) + ': ' + d.status + (d.inTime ? ' · in ' + d.inTime : '') + (d.outTime ? ' · out ' + d.outTime : '')) + '">' +
          '<b>' + Number(d.date.slice(8)) + '</b>' + (shortOf(d.status) || (d.holiday ? 'Off' : '')) + '</div>';
      }).join('');
    $('drawer').innerHTML =
      '<div class="drawer-head"><span class="avatar">' + esc(initials(s.name)) + '</span><div style="flex:1;min-width:0"><h2>' + esc(s.name) + '</h2>' +
      '<p>ID ' + esc(s.id) + (s.dept ? ' · ' + esc(s.dept) : '') + '</p></div>' +
      '<button class="btn ghost sm" id="drawerDash">Charts</button>' +
      '<button class="btn ghost icon" id="closeDrawer" aria-label="Close">' + icon('close') + '</button></div>' +
      '<div class="drawer-body">' +
      '<div class="grid kpis">' +
        kpi('Present', s.presentDays + '<small>/' + s.scheduledDays + '</small>', 'check', 't-good') +
        kpi('Lates', s.lateTimes, 'clock', 't-warn', s.chargedLates + ' charged') +
        kpi('Half days', s.halfDays, 'half', 't-half') +
        kpi('Absent', s.absentDays, 'x', 't-bad') +
        kpi('Overtime', hours(s.extraMin) + '<small>h</small>', 'plus', 't-info', s.extraDays + ' days') +
        kpi('Missed punch', s.missedPunches, 'alert', 't-miss') +
      '</div>' +
      '<div class="card"><div class="card-head"><div><h2>Salary</h2><p>For ' + fmtDate(state.data.period.from) + ' – ' + fmtDate(state.data.period.to) + '</p></div>' +
        '<label class="field"><span>Monthly salary (Rs)</span><input class="money-in" inputmode="numeric" id="drawerSalary" value="' + (p.salary || '') + '" placeholder="0"></label></div>' +
        breakdownHtml(p) +
      '</div>' +
      '<div class="card"><div class="card-head"><div><h2>This period</h2></div></div>' +
        '<div class="legend">' + legendHtml(GROUPS) + '</div><div class="strip">' + strip + '</div></div>' +
      '<div class="card"><div class="card-head"><div><h2>Which days</h2><p>When each late, half day and extra work happened</p></div></div>' +
        whichHtml(days) + '</div>' +
      '<div class="card"><div class="card-head"><div><h2>Day by day</h2></div></div>' + dayTableHtml(days) + '</div></div>';

    $('closeDrawer').addEventListener('click', closeDetail);
    $('drawerDash').addEventListener('click', () => {
      state.dashEmp = id; $('dashEmp').value = id;
      closeDetail(); showTab('dashboard');
    });
    $('drawerSalary').addEventListener('change', (e) => {
      setSalary(id, e.target.value);
      renderPayroll(); renderKpis(); if (state.tab === 'dashboard') renderCharts();
      openDetail(id);
    });
    const scrim = $('scrim');
    scrim.classList.remove('hidden');
    requestAnimationFrame(() => { scrim.classList.add('open'); $('drawer').classList.add('open'); });
    $('drawer').setAttribute('aria-hidden', 'false');
  }

  const line = (label, value, cls) => '<div><span>' + label + '</span><b' + (cls ? ' style="color:var(' + cls + ')"' : '') + '>' + value + '</b></div>';
  function breakdownHtml(p) {
    return p.hasSalary ? '<div class="breakdown">' +
      line('Monthly salary', rs(p.salary)) +
      line('Days off · ' + plural(p.offDays, 'day') + ', ' + p.paidLeave + ' paid', p.absentDed ? '−' + rs(p.absentDed) : rs(0), p.absentDed ? '--bad-text' : '') +
      line('Half days · ' + p.halfDays, p.halfDed ? '−' + rs(p.halfDed) : rs(0), p.halfDed ? '--bad-text' : '') +
      line('Late fines · ' + p.chargedLates + ' of ' + p.lateTimes + ' charged', p.lateFine ? '−' + rs(p.lateFine) : rs(0), p.lateFine ? '--bad-text' : '') +
      line('Overtime · ' + hours(p.otMin) + ' h', p.otPay ? '+' + rs(p.otPay) : rs(0), p.otPay ? '--good-text' : '') +
      '<div class="total"><span>Net pay</span><b>' + rs(p.net) + '</b></div></div>'
      : '<p class="note" style="margin:0">Enter a monthly salary to see the deductions and net pay.</p>';
  }

  function dayTableHtml(days) {
    return '<div class="table-wrap" style="max-height:none"><table>' +
      '<thead><tr><th>Date</th><th>In</th><th>Out</th><th>Status</th><th class="r">Extra (min)</th></tr></thead><tbody>' +
      days.map((d) => '<tr><td>' + d.day.slice(0, 1) + d.day.slice(1, 3).toLowerCase() + ', ' + fmtDate(d.date, false) + '</td><td class="num">' + (d.inTime || '—') +
        '</td><td class="num">' + (d.outTime || '—') + '</td><td>' + pill(d.status) + (d.holiday ? ' <span class="pill t-off">' + esc(d.holiday) + '</span>' : '') + '</td>' +
        '<td class="r">' + (d.extraMin || '—') + (d.extraNote && d.extraRawMin > d.extraMin ? '<small class="sub">' + esc(d.extraNote) + ', ' + (d.extraRawMin - d.extraMin) + ' not counted</small>' : '') + '</td></tr>').join('') +
      '</tbody></table></div>';
  }

  // Which days each issue happened on, e.g. "Late: Mon 3 Nov · 12 min"
  function whichHtml(days) {
    const has = (d, flag) => d.status.split(', ').includes(flag);
    const dayName = (d) => d.day.slice(0, 1) + d.day.slice(1, 3).toLowerCase() + ' ' + fmtDate(d.date, false);
    const whichGroups = [
      { label: 'Late', tone: 't-warn', pick: (d) => has(d, 'Late'), note: (d) => 'in ' + d.inTime + ' · ' + d.lateMin + ' min late' },
      { label: 'Late exit', tone: 't-warn', pick: (d) => has(d, 'Left Early'), note: (d) => 'out ' + d.outTime + ' · ' + d.earlyMin + ' min early' },
      { label: 'Half day', tone: 't-half', pick: (d) => has(d, 'Half Day'), note: (d) => [d.inTime && 'in ' + d.inTime, d.outTime && 'out ' + d.outTime].filter(Boolean).join(' · ') },
      { label: 'Extra work', tone: 't-info', pick: (d) => d.extraMin > 0, note: (d) => 'out ' + d.outTime + ' · ' + (d.extraMin >= 60 ? hours(d.extraMin) + ' h' : d.extraMin + ' min') + (d.extraNote ? ' (after ' + d.extraNote + ')' : '') },
      { label: 'Not counted (event)', tone: 't-off', pick: (d) => d.extraNote && d.extraRawMin > d.extraMin, note: (d) => d.extraNote + ' · out ' + d.outTime + ' · ' + (d.extraRawMin - d.extraMin) + ' min not counted' },
      { label: 'Absent', tone: 't-bad', pick: (d) => d.status === 'Absent', note: () => '' },
      { label: 'Missed punch', tone: 't-miss', pick: (d) => has(d, 'Missed Punch'), note: (d) => d.punches ? 'only ' + d.punches : '' },
    ];
    const which = whichGroups.map((g) => {
      const list = days.filter(g.pick);
      return '<div class="which"><div class="which-head"><span class="pill ' + g.tone + '">' + g.label + '</span><small>' + plural(list.length, 'day') + '</small></div>' +
        (list.length
          ? '<ul>' + list.map((d) => '<li><b>' + dayName(d) + '</b>' + (g.note(d) ? '<span>' + esc(g.note(d)) + '</span>' : '') + '</li>').join('') + '</ul>'
          : '<p class="dim">None</p>') + '</div>';
    }).join('');
    return '<div class="which-grid">' + which + '</div>';
  }
  function closeDetail() {
    openId = null;
    $('drawer').classList.remove('open');
    $('drawer').setAttribute('aria-hidden', 'true');
    $('scrim').classList.remove('open');
    setTimeout(() => { if (!openId) $('scrim').classList.add('hidden'); }, 250);
  }
  $('scrim').addEventListener('click', closeDetail);
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && openId) closeDetail(); });

  // ---------- Hover tooltip for calendar cells ----------
  const tip = $('tip');
  function placeTip(x, y) {
    const w = tip.offsetWidth, h = tip.offsetHeight;
    let left = x + 14, top = y - h - 10;
    if (left + w > innerWidth - 8) left = x - w - 14;
    if (top < 8) top = y + 18;
    tip.style.left = Math.max(8, left) + 'px';
    tip.style.top = top + 'px';
  }
  function showTip(t) {
    if (!t) { tip.hidden = true; return false; }
    tip.innerHTML = t.dataset.tip;
    tip.hidden = false;
    return true;
  }
  document.addEventListener('mouseover', (e) => { if (showTip(e.target.closest('[data-tip]'))) placeTip(e.clientX, e.clientY); });
  document.addEventListener('mousemove', (e) => { if (!tip.hidden) placeTip(e.clientX, e.clientY); });
  document.addEventListener('focusin', (e) => {
    const t = e.target.closest('[data-tip]');
    if (showTip(t)) { const r = t.getBoundingClientRect(); placeTip(r.left + r.width / 2, r.top); }
  });
  document.addEventListener('focusout', () => (tip.hidden = true));
  addEventListener('scroll', () => (tip.hidden = true), { passive: true });

  // ---------- Google Sheet ----------
  if (prefs.url) $('url').value = prefs.url;
  function sendMsg(kind, html) {
    $('sendMsg').innerHTML = '<div class="banner ' + kind + '" style="margin:0">' + icon(kind === 'ok' ? 'check' : kind === 'bad' ? 'alert' : 'info') + '<div>' + html + '</div></div>';
  }
  $('sendBtn').addEventListener('click', async () => {
    const url = $('url').value.trim();
    if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(url)) {
      sendMsg('bad', 'Enter a valid Apps Script Web App URL ending in <b>/exec</b>.');
      return;
    }
    prefs.url = url; persist();
    const btn = $('sendBtn');
    btn.disabled = true;
    sendMsg('info', 'Writing ' + state.data.daily.length + ' daily rows, ' + state.data.summary.length + ' summary rows and the payroll…');
    try {
      const payload = Object.assign({}, state.data, {
        payroll: state.pay.map((x) => Object.assign({}, x, { dayRate: Math.round(x.dayRate), otHours: +(x.otMin / 60).toFixed(2) })),
        holidays: Object.keys(prefs.holidays).sort().map((d) => ({ date: d, name: prefs.holidays[d] })),
      });
      // text/plain avoids a CORS preflight, which Apps Script does not support
      const res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(payload) });
      const text = await res.text();
      let out;
      try { out = JSON.parse(text); } catch (e) {
        throw new Error(/doPost/.test(text)
          ? 'The script has no doPost function. Paste Code.gs into Apps Script and deploy a new version.'
          : 'Unexpected response from Apps Script. Check the deployment access is set to "Anyone".');
      }
      if (!out.ok) throw new Error(out.error);
      sendMsg('ok', 'Done. Daily: ' + out.daily.added + ' added, ' + out.daily.updated + ' updated · Summary: ' + out.summary.added + ' added, ' +
        out.summary.updated + ' updated' + (out.payroll ? ' · Payroll: ' + out.payroll.added + ' added, ' + out.payroll.updated + ' updated' :
        '. <b>Payroll was not saved</b>: deploy the latest Code.gs to enable it') + '.' +
        (out.sheetUrl ? ' <a href="' + esc(out.sheetUrl) + '" target="_blank" rel="noopener">Open the sheet</a>' : ''));
    } catch (err) {
      sendMsg('bad', 'Upload failed: ' + esc(err.message));
    } finally {
      btn.disabled = false;
    }
  });
})();
