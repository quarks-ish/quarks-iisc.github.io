/**
 * QUARKS: SHEET -> GOOGLE CALENDAR SYNC  (Google Apps Script)
 *
 * Turns each row of the events spreadsheet into an event on the shared clubs
 * calendar, and keeps the two in step: edit a row and its event updates, clear
 * a row's Title and its event is deleted. Nothing is ever duplicated: the
 * script writes each event's ID back into a "Calendar Event ID" column and
 * uses it on the next run.
 *
 * SETUP (once)
 *  1. Open the events spreadsheet -> Extensions -> Apps Script, paste this file.
 *  2. Set CALENDAR_ID below (Calendar settings -> "Integrate calendar"). The
 *     Google account that runs the script needs "Make changes to events" on it.
 *  3. Project settings -> tick "Show appsscript.json", set "timeZone" to
 *     "Asia/Kolkata".
 *  4. Run installTrigger() once and approve the permissions. After that every
 *     edit to the sheet syncs automatically. (Or use the Calendar menu in the
 *     sheet -> "Sync now" any time.)
 *
 * COLUMNS (matched by header name, any order, case-insensitive)
 *   Title        required ("Announcement Title" also works). Cleared = delete the event.
 *   Date         required (or "Deadline", as in the announcements sheet: that
 *                makes an all-day event titled "Deadline: <Title>").
 *   Start Time   optional ("Time" also works). Blank = all-day event.
 *   End Time     optional. Blank = one hour after Start Time.
 *   End Date     optional. Blank = same as Date (use it for multi-day events).
 *   Location     optional ("Venue" also works).
 *   Category     optional ("Club name" also works). The club name. Shown on the website as a coloured
 *                tag: the event title becomes "[Category] Title".
 *   Link         optional. Added to the event description.
 *   Description  optional.
 */

var CONFIG = {
  CALENDAR_ID: 'PASTE_CALENDAR_ID_HERE@group.calendar.google.com',
  SHEET_NAME: '', // leave blank to use the first tab (form-response tabs have odd names)
  DEFAULT_DURATION_MIN: 60
};

var ID_HEADER = 'Calendar Event ID';

/** Adds a Calendar menu to the sheet. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Calendar')
    .addItem('Sync now', 'syncAll')
    .addToUi();
}

/** Run once: makes the sheet sync on every change. Safe to re-run. */
function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncAll') ScriptApp.deleteTrigger(t);
  });
  var ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('syncAll').forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('syncAll').forSpreadsheet(ss).onFormSubmit().create();
}

function syncAll() {
  // Edits can arrive in bursts; make runs take turns so rows are never doubled.
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    run_();
  } finally {
    lock.releaseLock();
  }
}

function run_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = (CONFIG.SHEET_NAME && ss.getSheetByName(CONFIG.SHEET_NAME)) || ss.getSheets()[0];
  var cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!cal) throw new Error('Calendar not found / no access: ' + CONFIG.CALENDAR_ID);

  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h).trim().toLowerCase();
  });
  // First matching header wins, so the same script works for the form-response
  // sheet ("Announcement Title", "Club name", "Venue", ...) and plain sheets.
  var col = function (names) {
    names = [].concat(names);
    for (var i = 0; i < names.length; i++) {
      var k = headers.indexOf(names[i].toLowerCase());
      if (k >= 0) return k;
    }
    return -1;
  };

  // Create the ID column the first time.
  var idCol = col(ID_HEADER);
  if (idCol < 0) {
    idCol = lastCol;
    sheet.getRange(1, idCol + 1).setValue(ID_HEADER);
    sheet.getRange(1, idCol + 1, 1, 1).setFontColor('#888888');
    lastCol++;
  }

  var c = {
    title: col(['Title', 'Announcement Title']),
    date: col(['Date', 'Event Date', 'Deadline']),
    isDeadline: col(['Date', 'Event Date']) < 0,
    start: col(['Start Time', 'Time']), end: col('End Time'), endDate: col('End Date'),
    loc: col(['Location', 'Venue']), cat: col(['Category', 'Club name', 'Club']),
    link: col('Link'), desc: col('Description')
  };
  if (c.title < 0 || c.date < 0) throw new Error('Sheet needs a title column ("Title"/"Announcement Title") and a date column ("Date"/"Deadline")');

  var tz = ss.getSpreadsheetTimeZone();
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var ids = [];

  values.forEach(function (row, i) {
    var id = String(row[idCol] || '');
    var title = String(row[c.title] || '').trim();
    var ev = id ? getEvent_(cal, id) : null;

    // Cleared title -> no event; remove one if it exists.
    if (!title) {
      if (ev) ev.deleteEvent();
      ids.push(['']);
      return;
    }

    var dateStr = dateStr_(row[c.date], tz);
    if (!dateStr) { ids.push([id]); return; } // no usable date yet: leave alone

    var startStr = c.start >= 0 ? timeStr_(row[c.start], tz) : '';
    var endStr = c.end >= 0 ? timeStr_(row[c.end], tz) : '';
    var endDateStr = (c.endDate >= 0 && dateStr_(row[c.endDate], tz)) || dateStr;

    var cat = c.cat >= 0 ? String(row[c.cat] || '').trim() : '';
    // The announcements sheet only has a Deadline column: those rows are
    // registration cut-offs, so label them as such on the calendar.
    if (c.isDeadline) title = 'Deadline: ' + title;
    var fullTitle = cat ? '[' + cat + '] ' + title : title;
    var loc = c.loc >= 0 ? String(row[c.loc] || '').trim() : '';
    var link = c.link >= 0 ? String(row[c.link] || '').trim() : '';
    if (!/^https?:\/\//i.test(link)) link = ''; // ignore notes typed into the Link box
    var desc = c.desc >= 0 ? String(row[c.desc] || '').trim() : '';
    var body = [desc, link ? 'More info: ' + link : ''].filter(String).join('\n\n');
    var opts = { location: loc, description: body };

    var when = times_(dateStr, startStr, endDateStr, endStr);

    if (ev) {
      ev.setTitle(fullTitle);
      ev.setLocation(loc);
      ev.setDescription(body);
      if (when.allDay) ev.setAllDayDates(when.start, when.end);
      else ev.setTime(when.start, when.end);
    } else {
      ev = when.allDay
        ? cal.createAllDayEvent(fullTitle, when.start, when.end, opts)
        : cal.createEvent(fullTitle, when.start, when.end, opts);
      id = ev.getId();
    }
    ids.push([id]);
  });

  sheet.getRange(2, idCol + 1, ids.length, 1).setValues(ids);
}

// ---- helpers ---------------------------------------------------------------

function getEvent_(cal, id) {
  try { return cal.getEventById(id); } catch (e) { return null; }
}

/** Cell -> "yyyy-MM-dd", or '' if it isn't a date. */
function dateStr_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = String(v || '').trim();
  if (!s) return '';
  var d = new Date(s);
  return isNaN(d) ? '' : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

/** Cell -> "HH:mm", or '' if blank. Accepts time cells and text like "6:30 PM". */
function timeStr_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'HH:mm');
  var s = String(v || '').trim();
  if (!s) return '';
  var m = s.match(/^(\d{1,2})(?::(\d{2}))?\s*([ap]m)?$/i);
  if (!m) return '';
  var h = parseInt(m[1], 10), min = m[2] ? parseInt(m[2], 10) : 0, ap = (m[3] || '').toLowerCase();
  if (ap === 'pm' && h < 12) h += 12;
  if (ap === 'am' && h === 12) h = 0;
  return (h < 10 ? '0' : '') + h + ':' + (min < 10 ? '0' : '') + min;
}

/** Builds {start, end, allDay}. Dates are parsed in the script's time zone. */
function times_(dateStr, startStr, endDateStr, endStr) {
  var tz = Session.getScriptTimeZone();
  var parse = function (d, t) { return Utilities.parseDate(d + ' ' + t, tz, 'yyyy-MM-dd HH:mm'); };
  if (!startStr) {
    // All-day: Calendar wants the end to be the day AFTER the last day.
    var start = parse(dateStr, '00:00'), end = parse(endDateStr, '00:00');
    end.setDate(end.getDate() + 1);
    return { start: start, end: end, allDay: true };
  }
  var s = parse(dateStr, startStr);
  var e = endStr ? parse(endDateStr, endStr) : new Date(s.getTime() + CONFIG.DEFAULT_DURATION_MIN * 60000);
  if (e <= s) e = new Date(s.getTime() + CONFIG.DEFAULT_DURATION_MIN * 60000);
  return { start: s, end: e, allDay: false };
}
