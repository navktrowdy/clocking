/**
 * MaaMadurai Clocking: staff clock in / clock out
 * Google Apps Script web app, bound to a Google Sheet.
 *
 * Paste this file into Extensions → Apps Script → Code.gs of your Google Sheet,
 * add Index.html, then Deploy → New deployment → Web app.
 * See README.md for step-by-step setup.
 */

// ---------------------------------------------------------------------------
// Sheet layout
// ---------------------------------------------------------------------------

const SHEET = {
  SETTINGS: 'Settings',
  STAFF: 'Staff',
  SHIFTS: 'Shifts',
  EVENTS: 'Events',
  TIMESHEETS: 'Timesheets',
  AUDIT: 'Audit log',
};

const HEADERS = {
  'Settings': ['Setting', 'Value', 'Notes'],
  'Staff': ['Staff ID', 'Name', 'Email', 'Hourly rate', 'Role', 'Active', 'Sign-in token', 'Added'],
  'Shifts': ['Shift ID', 'Staff ID', 'Name', 'Email', 'Date', 'Week start', 'Clock in', 'Clock out',
    'Break minutes', 'On break since', 'Worked hours', 'Hourly rate', 'Pay', 'In location',
    'In distance (m)', 'Out location', 'Out distance (m)', 'Status', 'Approval', 'Approval reason',
    'Staff note', 'Manager note', 'Week approved', 'Updated'],
  'Events': ['Time', 'Staff ID', 'Name', 'Action', 'Latitude', 'Longitude', 'Accuracy (m)',
    'Distance (m)', 'Location', 'Note', 'Shift ID'],
  'Timesheets': ['Week start', 'Staff ID', 'Name', 'Email', 'Shifts', 'Hours', 'Pay',
    'Approved by', 'Approved at', 'Pay fortnight start'],
  'Audit log': ['Time', 'By', 'Shift ID', 'Change', 'Old value', 'New value'],
};

const SETTING = {
  NAME: 'Business name',
  LAT: 'Restaurant latitude',
  LNG: 'Restaurant longitude',
  RADIUS: 'Allowed radius (metres)',
  CURRENCY: 'Currency',
  FORTNIGHT: 'Pay fortnight starts',
  MISSING_HOUR: 'Missing clock-out hour',
};

const STATUS = { OPEN: 'Open', BREAK: 'On break', DONE: 'Complete', MISSING: 'Missing clock-out' };
const APPROVAL = { AUTO: 'Auto-approved', NEEDS: 'Needs approval', OK: 'Approved', REJECTED: 'Rejected' };
const LOC = { ON: 'On site', OFF: 'Off site', NONE: 'No location', UNSET: 'Restaurant location not set', MANUAL: 'Manual entry' };
const MAX_SHIFT_HOURS = 16;

function idx_(name) {
  const m = {};
  HEADERS[name].forEach(function (h, i) { m[h] = i; });
  return m;
}
const ST = idx_('Staff');
const SH = idx_('Shifts');

// ---------------------------------------------------------------------------
// Web app entry point
// ---------------------------------------------------------------------------

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!isSetUp_()) {
    const owner = setup();
    return page_('Setup complete',
      'Your manager sign-in link has been emailed to <b>' + esc_(owner) + '</b>. Open that email on your phone to continue.');
  }
  const token = /^[a-f0-9]{32}$/.test(p.t || '') ? p.t : '';
  const user = token ? findStaffByToken_(token) : null;

  const t = HtmlService.createTemplateFromFile('Index');
  t.token = user ? token : '';
  t.isManager = user && user.role === 'Manager' ? '1' : '';
  t.badToken = p.t && !user ? '1' : '';
  t.page = /^(clock|week|staff|settings)$/.test(p.page || '') ? p.page : '';
  t.week = /^\d{4}-\d{2}-\d{2}$/.test(p.week || '') ? p.week : '';
  t.appName = settings_()[SETTING.NAME] || 'MaaMadurai';
  return t.evaluate()
    .setTitle(t.appName + ' Clock')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function page_(title, html) {
  return HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui,sans-serif;max-width:480px;margin:40px auto;padding:0 16px">' +
    '<h2>' + esc_(title) + '</h2><p style="font-size:17px;line-height:1.5">' + html + '</p></div>')
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------------------------------------------------------------------------
// One-time setup (runs automatically the first time the web app is opened)
// ---------------------------------------------------------------------------

function setup() {
  const ss = ss_();
  const tz = Session.getScriptTimeZone() || 'Australia/Sydney';
  ss.setSpreadsheetTimeZone(tz);

  Object.keys(HEADERS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });

  // Text columns (stop Sheets turning dates / tokens into other types)
  sheet_(SHEET.SHIFTS).getRange('A:F').setNumberFormat('@');
  sheet_(SHEET.SHIFTS).getRange('G:H').setNumberFormat('yyyy-mm-dd hh:mm');
  sheet_(SHEET.SHIFTS).getRange('J:J').setNumberFormat('yyyy-mm-dd hh:mm');
  sheet_(SHEET.STAFF).getRange('A:C').setNumberFormat('@');
  sheet_(SHEET.STAFF).getRange('E:G').setNumberFormat('@');
  sheet_(SHEET.TIMESHEETS).getRange('A:D').setNumberFormat('@');
  sheet_(SHEET.TIMESHEETS).getRange('J:J').setNumberFormat('@');

  // Default settings
  const set = sheet_(SHEET.SETTINGS);
  if (set.getLastRow() < 2) {
    const thisMonday = weekStartOf_(ymd_(new Date()));
    set.getRange('B:B').setNumberFormat('@');
    set.getRange(2, 1, 7, 3).setValues([
      [SETTING.NAME, 'MaaMadurai', ''],
      [SETTING.LAT, '', 'Set from the app: Manager → Settings → Use my current location (stand inside the restaurant)'],
      [SETTING.LNG, '', ''],
      [SETTING.RADIUS, '150', 'Clock-ins further than this from the restaurant need manager approval'],
      [SETTING.CURRENCY, 'AUD', ''],
      [SETTING.FORTNIGHT, thisMonday, 'Any Monday that starts a pay fortnight (yyyy-mm-dd)'],
      [SETTING.MISSING_HOUR, '4', 'Shifts still open after this hour the next morning are flagged "Missing clock-out"'],
    ]);
    set.autoResizeColumns(1, 3);
  }

  // Owner becomes the first manager
  const owner = Session.getEffectiveUser().getEmail();
  if (!readStaff_().some(function (s) { return s.email === owner.toLowerCase(); })) {
    appendRow_(SHEET.STAFF, [newId_('S'), 'Owner', owner, 0, 'Manager', 'Yes', newToken_(), new Date()]);
  }
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);

  installTriggers_();
  PropertiesService.getScriptProperties().setProperty('SETUP_DONE', 'yes');

  const me = readStaff_().filter(function (s) { return s.email === owner.toLowerCase(); })[0];
  if (me) sendSignInLink_(me);
  return owner;
}

function isSetUp_() {
  return PropertiesService.getScriptProperties().getProperty('SETUP_DONE') === 'yes' && !!sheet_(SHEET.STAFF);
}

function installTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['dailyCheck', 'weeklyReminder'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyCheck').timeBased().everyDays(1).atHour(4).create();
  ScriptApp.newTrigger('weeklyReminder').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
}

/** Run from the Apps Script editor if you lose your manager link. */
function resendMyManagerLink() {
  const owner = Session.getEffectiveUser().getEmail().toLowerCase();
  const me = readStaff_().filter(function (s) { return s.email === owner; })[0];
  if (!me) throw new Error('Owner is not in the Staff sheet');
  sendSignInLink_(me);
}

// ---------------------------------------------------------------------------
// Staff API (called from Index.html via google.script.run)
// ---------------------------------------------------------------------------

/** Email a sign-in link to a registered staff email. Never reveals whether the email exists. */
function requestLink(email) {
  email = String(email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Please enter a valid email address.');
  const cache = CacheService.getScriptCache();
  if (cache.get('link:' + email)) return 'If that email is registered, a sign-in link has been sent. Check your inbox.';
  cache.put('link:' + email, '1', 120);
  const s = readStaff_().filter(function (x) { return x.email === email && x.active; })[0];
  if (s) sendSignInLink_(s);
  return 'If that email is registered, a sign-in link has been sent. Check your inbox.';
}

function getState(token) {
  const user = auth_(token);
  const s = settings_();
  const shifts = readShifts_();
  const open = openShiftFor_(shifts, user.id);
  const week = weekStartOf_(ymd_(new Date()));
  const mine = shifts.filter(function (x) { return x.staffId === user.id && x.weekStart === week; });
  let hours = 0, pay = 0;
  mine.forEach(function (x) { if (x.approval !== APPROVAL.REJECTED) { hours += num_(x.hours); pay += num_(x.pay); } });
  return {
    name: user.name,
    role: user.role,
    currency: s[SETTING.CURRENCY] || 'AUD',
    locationReady: !isNaN(parseFloat(s[SETTING.LAT])),
    open: open ? {
      status: open.status,
      clockIn: hm_(open.clockIn),
      clockInDay: niceDate_(open.date),
      breakMin: num_(open.breakMin),
      onBreakSince: hm_(open.onBreakSince),
      inLocation: open.inLocation,
    } : null,
    week: {
      label: weekLabel_(week),
      shifts: mine.map(shiftView_),
      hours: round2_(hours),
      pay: round2_(pay),
    },
  };
}

/** action: IN | OUT | BREAK_START | BREAK_END; loc: {lat,lng,acc} or {error} or null */
function punch(token, action, loc, note) {
  const user = auth_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    markMissing_();
    const now = new Date();
    const s = settings_();
    const where = evalLocation_(loc, s);
    note = String(note || '').slice(0, 300);
    const shifts = readShifts_();
    let open = openShiftFor_(shifts, user.id);
    let msg = '';

    if (action === 'IN') {
      if (open) throw new Error('You are already clocked in (since ' + hm_(open.clockIn) + ').');
      const date = ymd_(now);
      open = {
        id: newId_('SH'), staffId: user.id, name: user.name, email: user.email, date: date,
        weekStart: weekStartOf_(date), clockIn: now, clockOut: '', breakMin: 0, onBreakSince: '',
        hours: '', rate: num_(user.rate), pay: '', inLocation: where.status, inDistance: where.distance,
        outLocation: '', outDistance: '', status: STATUS.OPEN, approval: '', reason: '',
        staffNote: note, managerNote: '', weekApproved: '', updated: now,
      };
      if (where.status !== LOC.ON) { open.approval = APPROVAL.NEEDS; open.reason = 'Clock-in: ' + where.status; }
      appendRow_(SHEET.SHIFTS, shiftToRow_(open));
      msg = 'Clocked in at ' + hm_(now) + '.';
    } else {
      if (!open) throw new Error('You are not clocked in.');
      if (action === 'BREAK_START') {
        if (open.status === STATUS.BREAK) throw new Error('You are already on a break.');
        open.onBreakSince = now;
        open.status = STATUS.BREAK;
        msg = 'Break started at ' + hm_(now) + '.';
      } else if (action === 'BREAK_END') {
        if (open.status !== STATUS.BREAK) throw new Error('You are not on a break.');
        endBreak_(open, now);
        msg = 'Break ended. Total break today: ' + open.breakMin + ' min.';
      } else if (action === 'OUT') {
        if (open.status === STATUS.BREAK) endBreak_(open, now);
        open.clockOut = now;
        open.outLocation = where.status;
        open.outDistance = where.distance;
        open.status = STATUS.DONE;
        const reasons = [];
        if (open.inLocation !== LOC.ON) reasons.push('Clock-in: ' + open.inLocation);
        if (where.status !== LOC.ON) reasons.push('Clock-out: ' + where.status);
        open.approval = reasons.length ? APPROVAL.NEEDS : APPROVAL.AUTO;
        open.reason = reasons.join('; ');
        recompute_(open);
        msg = 'Clocked out at ' + hm_(now) + '. ' + open.hours + ' hours worked.';
      } else {
        throw new Error('Unknown action');
      }
      if (note) open.staffNote = open.staffNote ? open.staffNote + ' | ' + note : note;
      open.updated = now;
      writeShift_(open);
    }

    appendRow_(SHEET.EVENTS, [now, user.id, user.name, action,
      loc && isFinite(loc.lat) ? Number(loc.lat) : '', loc && isFinite(loc.lng) ? Number(loc.lng) : '',
      loc && isFinite(loc.acc) ? Math.round(loc.acc) : '', where.distance, where.status,
      note || (loc && loc.error ? 'GPS error: ' + String(loc.error).slice(0, 100) : ''), open.id]);

    if (where.status !== LOC.ON && action !== 'BREAK_START' && action !== 'BREAK_END') {
      msg += ' Your location was "' + where.status + '", so this shift will be checked by the manager.';
    }
    return { message: msg, state: getState(token) };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Manager API
// ---------------------------------------------------------------------------

function mgrGetWeek(token, weekStart) {
  auth_(token, true);
  markMissing_();
  const s = settings_();
  const thisWeek = weekStartOf_(ymd_(new Date()));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(weekStart || '')) weekStart = addDays_(thisWeek, -7);
  weekStart = weekStartOf_(weekStart);

  const staff = readStaff_();
  const shifts = readShifts_();
  const inWeek = shifts.filter(function (x) { return x.weekStart === weekStart; });
  const byStaff = {};
  staff.forEach(function (p) { byStaff[p.id] = { id: p.id, name: p.name, active: p.active, shifts: [], hours: 0, pay: 0 }; });
  inWeek.forEach(function (x) {
    if (!byStaff[x.staffId]) byStaff[x.staffId] = { id: x.staffId, name: x.name, active: false, shifts: [], hours: 0, pay: 0 };
    const b = byStaff[x.staffId];
    b.shifts.push(shiftView_(x));
    if (x.approval !== APPROVAL.REJECTED) { b.hours += num_(x.hours); b.pay += num_(x.pay); }
  });
  const list = Object.keys(byStaff).map(function (k) { return byStaff[k]; })
    .filter(function (b) { return b.active || b.shifts.length; })
    .map(function (b) {
      b.shifts.sort(function (a, c) { return a.sortKey - c.sortKey; });
      b.hours = round2_(b.hours); b.pay = round2_(b.pay);
      return b;
    });

  const blocking = inWeek.filter(isBlocking_).length;
  const approvedRows = rows_(SHEET.TIMESHEETS).filter(function (r) { return String(r[0]) === weekStart; });
  const otherWeeks = {};
  shifts.forEach(function (x) {
    if (x.weekStart !== weekStart && isBlocking_(x) && x.weekStart < thisWeek) otherWeeks[x.weekStart] = true;
  });

  return {
    weekStart: weekStart,
    label: weekLabel_(weekStart),
    isCurrentOrFuture: weekStart >= thisWeek,
    fortnight: fortnightInfo_(weekStart, s),
    staff: list,
    totalHours: round2_(list.reduce(function (a, b) { return a + b.hours; }, 0)),
    totalPay: round2_(list.reduce(function (a, b) { return a + b.pay; }, 0)),
    blocking: blocking,
    approvedAt: approvedRows.length ? fmt_(approvedRows[0][8], 'EEE d MMM HH:mm') : '',
    approvedBy: approvedRows.length ? String(approvedRows[0][7]) : '',
    otherPendingWeeks: Object.keys(otherWeeks).sort().map(function (w) { return { weekStart: w, label: weekLabel_(w) }; }),
    currency: s[SETTING.CURRENCY] || 'AUD',
  };
}

/**
 * Edit a shift. data: {date:'yyyy-MM-dd', clockIn:'HH:mm', clockOut:'HH:mm'|'', breakMin, managerNote, decision:'approve'|'reject'|'save'}
 * If clock-out time is earlier than clock-in time the shift is treated as finishing the next day.
 */
function mgrUpdateShift(token, shiftId, data) {
  const mgr = auth_(token, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const shifts = readShifts_();
    const sh = shifts.filter(function (x) { return x.id === shiftId; })[0];
    if (!sh) throw new Error('Shift not found.');
    const before = { 'Clock in': fmtDT_(sh.clockIn), 'Clock out': fmtDT_(sh.clockOut), 'Break minutes': String(num_(sh.breakMin)), 'Approval': sh.approval, 'Manager note': sh.managerNote };

    applyTimes_(sh, data);
    sh.managerNote = String(data.managerNote || '').slice(0, 300);
    if (sh.clockOut instanceof Date && (sh.status === STATUS.MISSING || sh.status === STATUS.OPEN || sh.status === STATUS.BREAK)) {
      sh.status = STATUS.DONE;
      sh.onBreakSince = '';
    }
    if (data.decision === 'approve') {
      if (!(sh.clockOut instanceof Date)) throw new Error('Enter a clock-out time before approving.');
      sh.approval = APPROVAL.OK;
    } else if (data.decision === 'reject') {
      sh.approval = APPROVAL.REJECTED;
      sh.status = STATUS.DONE; // closes open / missing shifts so they no longer block
      sh.onBreakSince = '';
    }
    recompute_(sh);
    sh.updated = new Date();
    writeShift_(sh);

    const after = { 'Clock in': fmtDT_(sh.clockIn), 'Clock out': fmtDT_(sh.clockOut), 'Break minutes': String(num_(sh.breakMin)), 'Approval': sh.approval, 'Manager note': sh.managerNote };
    Object.keys(before).forEach(function (k) {
      if (before[k] !== after[k]) appendRow_(SHEET.AUDIT, [new Date(), mgr.email, sh.id, k, before[k], after[k]]);
    });
    return mgrGetWeek(token, sh.weekStart);
  } finally {
    lock.releaseLock();
  }
}

/** Manager adds a shift that the staff member forgot to record. */
function mgrAddShift(token, staffId, data) {
  const mgr = auth_(token, true);
  const p = readStaff_().filter(function (x) { return x.id === staffId; })[0];
  if (!p) throw new Error('Staff member not found.');
  if (!data.clockOut) throw new Error('Enter both clock-in and clock-out times.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = {
      id: newId_('SH'), staffId: p.id, name: p.name, email: p.email, date: '', weekStart: '',
      clockIn: '', clockOut: '', breakMin: 0, onBreakSince: '', hours: '', rate: num_(p.rate), pay: '',
      inLocation: LOC.MANUAL, inDistance: '', outLocation: LOC.MANUAL, outDistance: '',
      status: STATUS.DONE, approval: APPROVAL.OK, reason: 'Added by manager',
      staffNote: '', managerNote: String(data.managerNote || '').slice(0, 300), weekApproved: '', updated: new Date(),
    };
    applyTimes_(sh, data);
    recompute_(sh);
    appendRow_(SHEET.SHIFTS, shiftToRow_(sh));
    appendRow_(SHEET.AUDIT, [new Date(), mgr.email, sh.id, 'Shift added', '', fmtDT_(sh.clockIn) + ' → ' + fmtDT_(sh.clockOut)]);
    return mgrGetWeek(token, sh.weekStart);
  } finally {
    lock.releaseLock();
  }
}

/** Approve the week, save totals to the Timesheets tab and email everyone. */
function mgrApproveWeek(token, weekStart) {
  const mgr = auth_(token, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let result;
  try {
    markMissing_();
    weekStart = weekStartOf_(weekStart);
    const shifts = readShifts_();
    const inWeek = shifts.filter(function (x) { return x.weekStart === weekStart; });
    const blocking = inWeek.filter(isBlocking_);
    if (blocking.length) throw new Error(blocking.length + ' shift(s) still need approval or a clock-out time.');
    if (!inWeek.length) throw new Error('There are no shifts in this week.');

    const s = settings_();
    const fn = fortnightInfo_(weekStart, s);
    const now = new Date();
    const staff = readStaff_();
    const summary = summarise_(shifts, weekStart, fn);

    // Upsert Timesheets rows for this week
    const tsSheet = sheet_(SHEET.TIMESHEETS);
    const existing = rows_(SHEET.TIMESHEETS);
    for (let i = existing.length - 1; i >= 0; i--) {
      if (String(existing[i][0]) === weekStart) tsSheet.deleteRow(i + 2);
    }
    summary.people.forEach(function (p) {
      appendRow_(SHEET.TIMESHEETS, [weekStart, p.id, p.name, p.email, p.count, p.hours, p.pay, mgr.email, now, fn.start]);
    });
    inWeek.forEach(function (x) { x.weekApproved = 'Yes'; writeShift_(x); });
    result = { summary: summary, staff: staff, fn: fn, s: s };
  } finally {
    lock.releaseLock();
  }

  // Emails (outside the lock)
  const cur = result.s[SETTING.CURRENCY] || 'AUD';
  const biz = result.s[SETTING.NAME] || 'MaaMadurai';
  let sent = 0;
  result.summary.people.forEach(function (p) {
    if (!p.email) return;
    MailApp.sendEmail({
      to: p.email,
      subject: biz + ': your hours for ' + weekLabel_(weekStart) + ' (' + p.hours + ' h, ' + money_(p.pay, cur) + ')',
      htmlBody: staffEmailHtml_(p, weekStart, result.fn, cur, biz),
      name: biz + ' Clocking',
    });
    sent++;
  });
  managers_(result.staff).forEach(function (m) {
    MailApp.sendEmail({
      to: m.email,
      subject: biz + ': approved timesheet ' + weekLabel_(weekStart) + ' (' + result.summary.hours + ' h, ' + money_(result.summary.pay, cur) + ')',
      htmlBody: managerEmailHtml_(result.summary, weekStart, result.fn, cur, biz, link_(m.token)),
      name: biz + ' Clocking',
    });
  });
  const week = mgrGetWeek(token, weekStart);
  week.message = 'Week approved. Summary emailed to ' + sent + ' staff and the manager.';
  return week;
}

function mgrListStaff(token) {
  auth_(token, true);
  return readStaff_().map(function (p) {
    return { id: p.id, name: p.name, email: p.email, rate: num_(p.rate), role: p.role, active: p.active };
  });
}

/** data: {id?, name, email, rate, role, active} */
function mgrSaveStaff(token, data) {
  const mgr = auth_(token, true);
  const name = String(data.name || '').trim();
  const email = String(data.email || '').trim().toLowerCase();
  const rate = Number(data.rate);
  if (!name) throw new Error('Enter a name.');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Enter a valid email.');
  if (!(rate >= 0)) throw new Error('Enter an hourly rate (e.g. 28.50).');
  const role = data.role === 'Manager' ? 'Manager' : 'Staff';
  const active = data.active === false ? 'No' : 'Yes';

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const staff = readStaff_();
    if (staff.some(function (p) { return p.email === email && p.id !== data.id; })) throw new Error('That email is already used by another staff member.');
    let p;
    if (data.id) {
      p = staff.filter(function (x) { return x.id === data.id; })[0];
      if (!p) throw new Error('Staff member not found.');
      if (p.email === mgr.email && (role !== 'Manager' || active !== 'Yes')) throw new Error('You cannot remove your own manager access.');
      sheet_(SHEET.STAFF).getRange(p.row, 2, 1, 5).setValues([[name, email, rate, role, active]]);
      return { staff: mgrListStaff(token), message: 'Saved ' + name + '.' };
    }
    p = { id: newId_('S'), name: name, email: email, rate: rate, role: role, active: true, token: newToken_() };
    appendRow_(SHEET.STAFF, [p.id, name, email, rate, role, active, p.token, new Date()]);
    sendSignInLink_(p);
    return { staff: mgrListStaff(token), message: 'Added ' + name + ' and emailed their sign-in link to ' + email + '.' };
  } finally {
    lock.releaseLock();
  }
}

/** Email a staff member their sign-in link. reset=true creates a new link (old one stops working). */
function mgrSendLink(token, staffId, reset) {
  auth_(token, true);
  const p = readStaff_().filter(function (x) { return x.id === staffId; })[0];
  if (!p) throw new Error('Staff member not found.');
  if (reset) {
    p.token = newToken_();
    sheet_(SHEET.STAFF).getRange(p.row, ST['Sign-in token'] + 1).setValue(p.token);
  }
  sendSignInLink_(p);
  return (reset ? 'New link created and emailed to ' : 'Sign-in link emailed to ') + p.email + '.';
}

function mgrGetSettings(token) {
  auth_(token, true);
  const s = settings_();
  return {
    name: s[SETTING.NAME], lat: s[SETTING.LAT], lng: s[SETTING.LNG], radius: s[SETTING.RADIUS],
    currency: s[SETTING.CURRENCY], fortnight: s[SETTING.FORTNIGHT], sheetUrl: ss_().getUrl(),
  };
}

/** data: {lat, lng, radius, fortnight, name} */
function mgrSaveSettings(token, data) {
  auth_(token, true);
  const upd = {};
  if (data.lat !== undefined && data.lat !== '') {
    const lat = Number(data.lat), lng = Number(data.lng);
    if (!(lat >= -90 && lat <= 90 && lng >= -180 && lng <= 180)) throw new Error('Invalid location.');
    upd[SETTING.LAT] = lat.toFixed(6);
    upd[SETTING.LNG] = lng.toFixed(6);
  }
  if (data.radius !== undefined) {
    const r = Number(data.radius);
    if (!(r >= 20 && r <= 5000)) throw new Error('Radius must be between 20 and 5000 metres.');
    upd[SETTING.RADIUS] = String(Math.round(r));
  }
  if (data.fortnight) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(data.fortnight)) throw new Error('Fortnight start must be a date.');
    upd[SETTING.FORTNIGHT] = weekStartOf_(data.fortnight);
  }
  if (data.name) upd[SETTING.NAME] = String(data.name).slice(0, 60);
  const sh = sheet_(SHEET.SETTINGS);
  const rows = rows_(SHEET.SETTINGS);
  Object.keys(upd).forEach(function (k) {
    const i = rows.findIndex(function (r) { return r[0] === k; });
    if (i >= 0) sh.getRange(i + 2, 2).setValue(upd[k]);
    else appendRow_(SHEET.SETTINGS, [k, upd[k], '']);
  });
  return mgrGetSettings(token);
}

// ---------------------------------------------------------------------------
// Scheduled jobs
// ---------------------------------------------------------------------------

/** Every day ~4am: flag shifts where staff forgot to clock out. */
function dailyCheck() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { markMissing_(); } finally { lock.releaseLock(); }
}

/** Every Monday ~7am: send the manager last week's timesheet to approve. */
function weeklyReminder() {
  dailyCheck();
  const s = settings_();
  const biz = s[SETTING.NAME] || 'MaaMadurai';
  const cur = s[SETTING.CURRENCY] || 'AUD';
  const week = addDays_(weekStartOf_(ymd_(new Date())), -7);
  const shifts = readShifts_();
  const inWeek = shifts.filter(function (x) { return x.weekStart === week; });
  if (!inWeek.length) return;
  const blocking = inWeek.filter(isBlocking_).length;
  const summary = summarise_(shifts, week, fortnightInfo_(week, s));
  managers_(readStaff_()).forEach(function (m) {
    const url = link_(m.token) + '&page=week&week=' + week;
    MailApp.sendEmail({
      to: m.email,
      subject: biz + ': timesheet ready to approve (' + weekLabel_(week) + ')' + (blocking ? ', ' + blocking + ' need attention' : ''),
      htmlBody:
        '<div style="font-family:Arial,sans-serif;font-size:15px">' +
        '<p>The timesheet for <b>' + weekLabel_(week) + '</b> is ready.</p>' +
        (blocking ? '<p style="color:#b45309"><b>' + blocking + ' shift(s) need your approval</b> (off-site clock-ins or missing clock-outs).</p>' : '<p>All shifts were recorded on site.</p>') +
        summaryTable_(summary.people, cur, false) +
        '<p><a href="' + url + '" style="background:#b91c1c;color:#fff;padding:12px 18px;border-radius:8px;text-decoration:none;display:inline-block">Review &amp; approve</a></p>' +
        '<p style="color:#666;font-size:13px">After you approve, each staff member is emailed their hours and pay.</p></div>',
      name: biz + ' Clocking',
    });
  });
}

function markMissing_() {
  const s = settings_();
  const cutoff = parseInt(s[SETTING.MISSING_HOUR], 10);
  const now = new Date();
  const today = ymd_(now);
  const hour = parseInt(fmt_(now, 'H'), 10);
  readShifts_().forEach(function (x) {
    if (x.status !== STATUS.OPEN && x.status !== STATUS.BREAK) return;
    const ageH = x.clockIn instanceof Date ? (now - x.clockIn) / 3600000 : 99;
    const pastCutoff = x.date < today && hour >= (isNaN(cutoff) ? 4 : cutoff);
    if (ageH > MAX_SHIFT_HOURS || pastCutoff) {
      x.status = STATUS.MISSING;
      x.approval = APPROVAL.NEEDS;
      x.reason = x.reason ? x.reason + '; No clock-out' : 'No clock-out';
      x.updated = now;
      writeShift_(x);
    }
  });
}

// ---------------------------------------------------------------------------
// Emails
// ---------------------------------------------------------------------------

function sendSignInLink_(p) {
  const s = settings_();
  const biz = s[SETTING.NAME] || 'MaaMadurai';
  const url = link_(p.token);
  if (!url) throw new Error('Deploy the web app first (Deploy → New deployment → Web app).');
  MailApp.sendEmail({
    to: p.email,
    subject: biz + ': your clock in / clock out link',
    htmlBody:
      '<div style="font-family:Arial,sans-serif;font-size:15px">' +
      '<p>Hi ' + esc_(p.name) + ',</p>' +
      '<p>This is your personal link for clocking in and out at ' + esc_(biz) + (p.role === 'Manager' ? ' (manager access)' : '') + '.</p>' +
      '<p><a href="' + url + '" style="background:#b91c1c;color:#fff;padding:12px 18px;border-radius:8px;text-decoration:none;display:inline-block">Open my clock</a></p>' +
      '<p><b>Tip:</b> open the link on your phone, then use <i>Share → Add to Home Screen</i> (iPhone) or <i>⋮ → Add to Home screen</i> (Android) so it works like an app.</p>' +
      '<p style="color:#666;font-size:13px">Do not share this link. It signs in as you. If you lose it, open the app and request a new link with this email address.</p></div>',
    name: biz + ' Clocking',
  });
}

function summarise_(shifts, weekStart, fn) {
  const people = {};
  let hours = 0, pay = 0;
  shifts.forEach(function (x) {
    if (x.approval === APPROVAL.REJECTED) return;
    const inThisWeek = x.weekStart === weekStart;
    const inFortnight = fn.isSecondWeek && (x.weekStart === fn.start || x.weekStart === weekStart);
    if (!inThisWeek && !inFortnight) return;
    if (!people[x.staffId]) people[x.staffId] = { id: x.staffId, name: x.name, email: x.email, count: 0, hours: 0, pay: 0, fnHours: 0, fnPay: 0, shifts: [] };
    const p = people[x.staffId];
    if (inThisWeek) {
      p.count++; p.hours += num_(x.hours); p.pay += num_(x.pay); p.shifts.push(x);
      hours += num_(x.hours); pay += num_(x.pay);
    }
    if (inFortnight) { p.fnHours += num_(x.hours); p.fnPay += num_(x.pay); }
  });
  const list = Object.keys(people).map(function (k) {
    const p = people[k];
    p.hours = round2_(p.hours); p.pay = round2_(p.pay); p.fnHours = round2_(p.fnHours); p.fnPay = round2_(p.fnPay);
    p.shifts.sort(function (a, b) { return a.clockIn - b.clockIn; });
    return p;
  }).filter(function (p) { return p.count > 0; }).sort(function (a, b) { return a.name < b.name ? -1 : 1; });
  return { people: list, hours: round2_(hours), pay: round2_(pay), fn: fn };
}

function summaryTable_(people, cur, withFortnight) {
  const th = 'style="text-align:left;padding:6px 10px;border-bottom:2px solid #ddd"';
  const td = 'style="padding:6px 10px;border-bottom:1px solid #eee"';
  let h = '<table style="border-collapse:collapse;font-size:14px"><tr><th ' + th + '>Staff</th><th ' + th + '>Shifts</th><th ' + th + '>Hours</th><th ' + th + '>Pay</th>' +
    (withFortnight ? '<th ' + th + '>Fortnight hours</th><th ' + th + '>Fortnight pay</th>' : '') + '</tr>';
  let th_ = 0, tp = 0, fh = 0, fp = 0;
  people.forEach(function (p) {
    th_ += p.hours; tp += p.pay; fh += p.fnHours; fp += p.fnPay;
    h += '<tr><td ' + td + '>' + esc_(p.name) + '</td><td ' + td + '>' + p.count + '</td><td ' + td + '>' + p.hours + '</td><td ' + td + '>' + money_(p.pay, cur) + '</td>' +
      (withFortnight ? '<td ' + td + '>' + p.fnHours + '</td><td ' + td + '>' + money_(p.fnPay, cur) + '</td>' : '') + '</tr>';
  });
  h += '<tr><td ' + td + '><b>Total</b></td><td ' + td + '></td><td ' + td + '><b>' + round2_(th_) + '</b></td><td ' + td + '><b>' + money_(tp, cur) + '</b></td>' +
    (withFortnight ? '<td ' + td + '><b>' + round2_(fh) + '</b></td><td ' + td + '><b>' + money_(fp, cur) + '</b></td>' : '') + '</tr></table>';
  return h;
}

function staffEmailHtml_(p, weekStart, fn, cur, biz) {
  const td = 'style="padding:6px 10px;border-bottom:1px solid #eee"';
  const th = 'style="text-align:left;padding:6px 10px;border-bottom:2px solid #ddd"';
  let rows = '';
  p.shifts.forEach(function (x) {
    rows += '<tr><td ' + td + '>' + niceDate_(x.date) + '</td><td ' + td + '>' + hm_(x.clockIn) + '–' + hm_(x.clockOut) + '</td><td ' + td + '>' + num_(x.breakMin) + ' min</td><td ' + td + '>' + x.hours + '</td><td ' + td + '>' + money_(x.pay, cur) + '</td></tr>';
  });
  return '<div style="font-family:Arial,sans-serif;font-size:15px">' +
    '<p>Hi ' + esc_(p.name) + ',</p>' +
    '<p>Your timesheet for <b>' + weekLabel_(weekStart) + '</b> has been approved.</p>' +
    '<table style="border-collapse:collapse;font-size:14px"><tr><th ' + th + '>Day</th><th ' + th + '>Time</th><th ' + th + '>Break</th><th ' + th + '>Hours</th><th ' + th + '>Pay</th></tr>' + rows +
    '<tr><td ' + td + '><b>Total</b></td><td ' + td + '></td><td ' + td + '></td><td ' + td + '><b>' + p.hours + '</b></td><td ' + td + '><b>' + money_(p.pay, cur) + '</b></td></tr></table>' +
    (fn.isSecondWeek
      ? '<p><b>Pay fortnight ' + fortnightLabel_(fn) + ':</b> ' + p.fnHours + ' hours, ' + money_(p.fnPay, cur) + '</p>'
      : '<p style="color:#666">This is week 1 of the pay fortnight ' + fortnightLabel_(fn) + '.</p>') +
    '<p style="color:#666;font-size:13px">Pay shown is gross (before tax and super). If anything looks wrong, please speak to your manager.</p>' +
    '<p>Thanks,<br>' + esc_(biz) + '</p></div>';
}

function managerEmailHtml_(summary, weekStart, fn, cur, biz, url) {
  return '<div style="font-family:Arial,sans-serif;font-size:15px">' +
    '<p>Timesheet for <b>' + weekLabel_(weekStart) + '</b> approved.</p>' +
    summaryTable_(summary.people, cur, fn.isSecondWeek) +
    (fn.isSecondWeek ? '<p><b>This completes the pay fortnight ' + fortnightLabel_(fn) + '.</b> Use the fortnight columns for payroll.</p>'
      : '<p style="color:#666">Week 1 of pay fortnight ' + fortnightLabel_(fn) + '.</p>') +
    '<p style="color:#666;font-size:13px">Gross pay before tax and super. Details are in the Timesheets and Shifts tabs of the Google Sheet.' +
    (url ? ' <a href="' + url + '">Open the app</a>.' : '') + '</p></div>';
}

// ---------------------------------------------------------------------------
// Data access
// ---------------------------------------------------------------------------

function ss_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('This script must be created from the Google Sheet (Extensions → Apps Script).');
  return SpreadsheetApp.openById(id);
}
function sheet_(name) { return ss_().getSheetByName(name); }

function rows_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS[name].length).getValues();
}

function appendRow_(name, row) {
  const sh = sheet_(name);
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

function settings_() {
  const o = {};
  rows_(SHEET.SETTINGS).forEach(function (r) { if (r[0]) o[String(r[0])] = r[1] instanceof Date ? ymd_(r[1]) : String(r[1]); });
  return o;
}

/** Reads staff and fills in any missing ID / token / defaults (so the manager can add staff straight into the sheet). */
function readStaff_() {
  const sh = sheet_(SHEET.STAFF);
  const rows = rows_(SHEET.STAFF);
  const out = [];
  rows.forEach(function (r, i) {
    if (!r[ST['Email']] && !r[ST['Name']]) return;
    let dirty = false;
    if (!r[ST['Staff ID']]) { r[ST['Staff ID']] = newId_('S'); dirty = true; }
    if (!r[ST['Sign-in token']]) { r[ST['Sign-in token']] = newToken_(); dirty = true; }
    if (!r[ST['Role']]) { r[ST['Role']] = 'Staff'; dirty = true; }
    if (!r[ST['Active']]) { r[ST['Active']] = 'Yes'; dirty = true; }
    if (!r[ST['Added']]) { r[ST['Added']] = new Date(); dirty = true; }
    if (dirty) sh.getRange(i + 2, 1, 1, r.length).setValues([r]);
    out.push({
      row: i + 2,
      id: String(r[ST['Staff ID']]),
      name: String(r[ST['Name']]).trim(),
      email: String(r[ST['Email']]).trim().toLowerCase(),
      rate: num_(r[ST['Hourly rate']]),
      role: /^manager$/i.test(String(r[ST['Role']]).trim()) ? 'Manager' : 'Staff',
      active: !/^(no|n|false|0)$/i.test(String(r[ST['Active']]).trim()),
      token: String(r[ST['Sign-in token']]),
    });
  });
  return out;
}

function findStaffByToken_(token) {
  return readStaff_().filter(function (s) { return s.token === token && s.active; })[0] || null;
}

function auth_(token, managerOnly) {
  const u = /^[a-f0-9]{32}$/.test(token || '') ? findStaffByToken_(token) : null;
  if (!u) throw new Error('Your sign-in link is not valid any more. Please request a new one.');
  if (managerOnly && u.role !== 'Manager') throw new Error('Manager access only.');
  return u;
}

function managers_(staff) {
  return staff.filter(function (s) { return s.role === 'Manager' && s.active && s.email; });
}

function readShifts_() {
  return rows_(SHEET.SHIFTS).map(function (r, i) {
    return {
      row: i + 2,
      id: String(r[SH['Shift ID']]),
      staffId: String(r[SH['Staff ID']]),
      name: String(r[SH['Name']]),
      email: String(r[SH['Email']]),
      date: r[SH['Date']] instanceof Date ? ymd_(r[SH['Date']]) : String(r[SH['Date']]),
      weekStart: r[SH['Week start']] instanceof Date ? ymd_(r[SH['Week start']]) : String(r[SH['Week start']]),
      clockIn: r[SH['Clock in']],
      clockOut: r[SH['Clock out']],
      breakMin: r[SH['Break minutes']],
      onBreakSince: r[SH['On break since']],
      hours: r[SH['Worked hours']],
      rate: r[SH['Hourly rate']],
      pay: r[SH['Pay']],
      inLocation: String(r[SH['In location']]),
      inDistance: r[SH['In distance (m)']],
      outLocation: String(r[SH['Out location']]),
      outDistance: r[SH['Out distance (m)']],
      status: String(r[SH['Status']]),
      approval: String(r[SH['Approval']]),
      reason: String(r[SH['Approval reason']]),
      staffNote: String(r[SH['Staff note']]),
      managerNote: String(r[SH['Manager note']]),
      weekApproved: String(r[SH['Week approved']]),
      updated: r[SH['Updated']],
    };
  }).filter(function (x) { return x.id; });
}

function shiftToRow_(x) {
  return [x.id, x.staffId, x.name, x.email, x.date, x.weekStart, x.clockIn, x.clockOut, x.breakMin,
    x.onBreakSince, x.hours, x.rate, x.pay, x.inLocation, x.inDistance, x.outLocation, x.outDistance,
    x.status, x.approval, x.reason, x.staffNote, x.managerNote, x.weekApproved, x.updated];
}

function writeShift_(x) {
  if (!x.row) throw new Error('Shift has no row');
  const row = shiftToRow_(x);
  sheet_(SHEET.SHIFTS).getRange(x.row, 1, 1, row.length).setValues([row]);
}

function openShiftFor_(shifts, staffId) {
  return shifts.filter(function (x) {
    return x.staffId === staffId && (x.status === STATUS.OPEN || x.status === STATUS.BREAK);
  })[0] || null;
}

function isBlocking_(x) {
  if (x.approval === APPROVAL.REJECTED) return false;
  return x.approval === APPROVAL.NEEDS || x.status === STATUS.MISSING ||
    x.status === STATUS.OPEN || x.status === STATUS.BREAK;
}

// ---------------------------------------------------------------------------
// Calculations
// ---------------------------------------------------------------------------

function endBreak_(sh, now) {
  if (sh.onBreakSince instanceof Date) {
    sh.breakMin = num_(sh.breakMin) + Math.max(0, Math.round((now - sh.onBreakSince) / 60000));
  }
  sh.onBreakSince = '';
  sh.status = STATUS.OPEN;
}

function recompute_(sh) {
  if (sh.clockIn instanceof Date && sh.clockOut instanceof Date) {
    const mins = Math.max(0, (sh.clockOut - sh.clockIn) / 60000 - num_(sh.breakMin));
    sh.hours = round2_(mins / 60);
    sh.pay = round2_(sh.hours * num_(sh.rate));
  } else {
    sh.hours = '';
    sh.pay = '';
  }
}

/** Apply manager-entered date / times to a shift. */
function applyTimes_(sh, data) {
  const date = /^\d{4}-\d{2}-\d{2}$/.test(data.date || '') ? data.date : sh.date;
  if (!date) throw new Error('Enter a date.');
  const hmRe = /^([01]?\d|2[0-3]):[0-5]\d$/;
  const inT = data.clockIn || (sh.clockIn instanceof Date ? hm_(sh.clockIn) : '');
  if (!hmRe.test(inT)) throw new Error('Enter a clock-in time like 09:30.');
  const inD = parseLocal_(date, inT);
  let outD = '';
  if (data.clockOut) {
    if (!hmRe.test(data.clockOut)) throw new Error('Enter a clock-out time like 17:00.');
    outD = parseLocal_(date, data.clockOut);
    if (outD <= inD) outD = parseLocal_(addDays_(date, 1), data.clockOut); // overnight
    if ((outD - inD) / 3600000 > 24) throw new Error('Shift cannot be longer than 24 hours.');
  }
  const br = Number(data.breakMin === undefined || data.breakMin === '' ? num_(sh.breakMin) : data.breakMin);
  if (!(br >= 0 && br <= 600)) throw new Error('Break minutes must be between 0 and 600.');
  sh.date = date;
  sh.weekStart = weekStartOf_(date);
  sh.clockIn = inD;
  sh.clockOut = outD;
  sh.breakMin = Math.round(br);
}

function evalLocation_(loc, s) {
  if (!loc || !isFinite(loc.lat) || !isFinite(loc.lng)) return { status: LOC.NONE, distance: '' };
  const lat = parseFloat(s[SETTING.LAT]), lng = parseFloat(s[SETTING.LNG]);
  if (isNaN(lat) || isNaN(lng)) return { status: LOC.UNSET, distance: '' };
  const radius = parseFloat(s[SETTING.RADIUS]) || 150;
  const d = Math.round(distanceM_(lat, lng, Number(loc.lat), Number(loc.lng)));
  const slack = Math.min(Math.max(Number(loc.acc) || 0, 0), 50); // allow for GPS accuracy, up to 50 m
  return { status: d <= radius + slack ? LOC.ON : LOC.OFF, distance: d };
}

function distanceM_(lat1, lng1, lat2, lng2) {
  const R = 6371000, rad = Math.PI / 180;
  const dLat = (lat2 - lat1) * rad, dLng = (lng2 - lng1) * rad;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return 2 * R * Math.asin(Math.sqrt(a));
}

function fortnightInfo_(weekStart, s) {
  const anchor = /^\d{4}-\d{2}-\d{2}$/.test(s[SETTING.FORTNIGHT] || '') ? weekStartOf_(s[SETTING.FORTNIGHT]) : weekStart;
  const weeks = Math.round(daysBetween_(anchor, weekStart) / 7);
  const isSecondWeek = ((weeks % 2) + 2) % 2 === 1;
  const start = isSecondWeek ? addDays_(weekStart, -7) : weekStart;
  return { start: start, end: addDays_(start, 13), isSecondWeek: isSecondWeek };
}

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

function tz_() { return Session.getScriptTimeZone() || 'Australia/Sydney'; }
function fmt_(d, f) { return d instanceof Date ? Utilities.formatDate(d, tz_(), f) : ''; }
function ymd_(d) { return fmt_(d, 'yyyy-MM-dd'); }
function hm_(d) { return fmt_(d, 'HH:mm'); }
function fmtDT_(d) { return fmt_(d, 'yyyy-MM-dd HH:mm'); }
function parseLocal_(ymd, hm) { return Utilities.parseDate(ymd + ' ' + hm, tz_(), 'yyyy-MM-dd HH:mm'); }

function ymdParts_(ymd) { const p = String(ymd).split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2]); }
function addDays_(ymd, n) { return new Date(ymdParts_(ymd) + n * 86400000).toISOString().slice(0, 10); }
function daysBetween_(a, b) { return Math.round((ymdParts_(b) - ymdParts_(a)) / 86400000); }
function weekStartOf_(ymd) {
  const dow = new Date(ymdParts_(ymd)).getUTCDay(); // 0 = Sunday
  return addDays_(ymd, -((dow + 6) % 7));
}

const MONTHS_ = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const DAYS_ = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
function niceDate_(ymd) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd)) return String(ymd);
  const d = new Date(ymdParts_(ymd));
  return DAYS_[d.getUTCDay()] + ' ' + d.getUTCDate() + ' ' + MONTHS_[d.getUTCMonth()];
}
function shortDate_(ymd) { const d = new Date(ymdParts_(ymd)); return d.getUTCDate() + ' ' + MONTHS_[d.getUTCMonth()]; }
function weekLabel_(ws) { return 'Mon ' + shortDate_(ws) + ' – Sun ' + shortDate_(addDays_(ws, 6)) + ' ' + addDays_(ws, 6).slice(0, 4); }
function fortnightLabel_(fn) { return shortDate_(fn.start) + ' – ' + shortDate_(fn.end); }

function shiftView_(x) {
  return {
    id: x.id, staffId: x.staffId, name: x.name, date: x.date, day: niceDate_(x.date),
    clockIn: hm_(x.clockIn), clockOut: hm_(x.clockOut),
    nextDay: x.clockIn instanceof Date && x.clockOut instanceof Date && ymd_(x.clockOut) !== ymd_(x.clockIn),
    breakMin: num_(x.breakMin), hours: x.hours === '' ? '' : num_(x.hours), pay: x.pay === '' ? '' : num_(x.pay),
    rate: num_(x.rate), status: x.status, approval: x.approval, reason: x.reason,
    inLocation: x.inLocation, inDistance: x.inDistance === '' ? '' : num_(x.inDistance),
    outLocation: x.outLocation, outDistance: x.outDistance === '' ? '' : num_(x.outDistance),
    staffNote: x.staffNote, managerNote: x.managerNote,
    blocking: isBlocking_(x),
    sortKey: x.clockIn instanceof Date ? x.clockIn.getTime() : 0,
  };
}

function link_(token) {
  const url = ScriptApp.getService().getUrl();
  return url ? url + '?t=' + token : '';
}
function newToken_() { return Utilities.getUuid().replace(/-/g, '').toLowerCase(); }
function newId_(prefix) { return prefix + '-' + Utilities.getUuid().slice(0, 8).toUpperCase(); }
function num_(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }
function round2_(n) { return Math.round((Number(n) + Number.EPSILON) * 100) / 100; }
function money_(n, cur) { return (cur === 'AUD' || !cur ? '$' : cur + ' ') + round2_(n).toFixed(2); }
function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

/** Lets Index.html include other files if needed. */
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }
