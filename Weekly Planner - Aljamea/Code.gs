/**
 * ==================================================================
 * JADWAL -> DATED WEEKLY GRID -> GOOGLE CALENDAR   (v4)
 * ------------------------------------------------------------------
 * The grid is DATED. Its header reads "Monday / 03 Aug", and the week
 * it represents is stored in the hidden state sheet. Nothing is inferred
 * from "is this day past or future" any more — a cell's date is simply
 * whatever the header says.
 *
 *   Set Jadwal URL     store your own MyJadwal link (no code editing)
 *   Jadwal → Calendar  toggle: scraped periods become events too
 *   Set this week      pin the grid to the current Sunday-Saturday week
 *   Sync to Calendar   every filled cell -> an event on its header date
 *   Roll over week     archive -> wipe the grid -> advance the dates
 *
 * Archive is a flat log in one sheet:
 *   Date | Day | Start | End | Source | Entry | Archived at
 *
 * Rolling over does NOT delete already-created calendar events — the week
 * happened, so it stays in your calendar as history.
 * ==================================================================
 */

const CONFIG = Object.freeze({
  TIMEZONE: 'Asia/Kolkata',
  JADWAL_URL: 'https://jameasaifiyah.org/MyJadwal.aspx?ID=MTg5Ng%3d%3d', // default; override from the menu
  SYNC_JADWAL_TO_CALENDAR: false,   // default for the menu toggle, first run only
  GRID_SHEET_NAME: '',              // blank = auto-detect the sheet holding the grid
  STATE_SHEET_NAME: '_GridState',   // hidden bookkeeping, do not edit by hand
  ARCHIVE_SHEET_NAME: 'Archive',
  SKIP_CANCELLED: true,             // cancelled Jadwal periods leave the slot empty
  MARK_UNMATCHED_AS_NOTE: true,
  KEEP_TIME_LABEL: true,            // leave "HH:MM - HH:MM" in cleared cells
  ARCHIVE_JADWAL_ROWS: true,        // archive scraped periods too, not just yours
  DATE_FORMAT: 'dd MMM',            // shown under each day name
  EVENT_PREFIX: '',                 // e.g. '[Grid] ' to tag grid events in Calendar
  TOAST_ON_EDIT: true,              // pop a corner toast so you can see the edit synced
  NAG_ON_STALE_WEEK: true,          // remind you when the grid week has drifted
  NAG_EMAIL: '',                    // blank = whoever owns the script
  NAG_EMAIL_MAX_PER_DAY: 1          // 0 = toast/log only, never email
});

const DAY_NAMES = Object.freeze(
  ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']);
const SLOT_RE = /^(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})\s*$/;
const ARCHIVE_HEADERS = ['Date', 'Day', 'Start', 'End', 'Source', 'Entry', 'Archived at'];
// A Jadwal-rendered block always starts "P<n> · <darajah>". Used as a
// content-based ownership fallback so a lost state row can never make the
// scraper's output look like something you typed.
const JADWAL_BLOCK_RE = /^P\d+\s*·/;

const KIND_MANUAL = 'MANUAL';
const KIND_JADWAL = 'JADWAL';

/** =================================================================
 * User settings that live in the document, not in the code.
 * CONFIG supplies the defaults; the menu overrides them.
 * ================================================================= */
class Settings {
  constructor(config) {
    this.config = config;
    this.store = PropertiesService.getDocumentProperties() ||
      PropertiesService.getScriptProperties();
  }

  static get KEY_URL() { return 'JADWAL_URL'; }
  static get KEY_SYNC() { return 'SYNC_JADWAL_TO_CALENDAR'; }

  _read(key) {
    try { return this.store.getProperty(key); } catch (e) { return null; }
  }

  _write(key, value) {
    if (value === null) this.store.deleteProperty(key);
    else this.store.setProperty(key, String(value));
  }

  get jadwalUrl() {
    const stored = this._read(Settings.KEY_URL);
    return stored ? stored : this.config.JADWAL_URL;
  }

  get jadwalUrlIsCustom() { return !!this._read(Settings.KEY_URL); }

  set jadwalUrl(url) {
    const clean = String(url == null ? '' : url).trim();
    if (!/^https?:\/\/\S+$/i.test(clean)) {
      throw new Error('That does not look like a URL. It must start with http:// or https://');
    }
    this._write(Settings.KEY_URL, clean);
  }

  resetJadwalUrl() { this._write(Settings.KEY_URL, null); }

  get syncJadwalToCalendar() {
    const stored = this._read(Settings.KEY_SYNC);
    return stored === null ? !!this.config.SYNC_JADWAL_TO_CALENDAR : stored === 'true';
  }

  set syncJadwalToCalendar(on) { this._write(Settings.KEY_SYNC, !!on); }

  /** Flip the toggle and hand back the new value. */
  toggleJadwalCalendar() {
    const next = !this.syncJadwalToCalendar;
    this.syncJadwalToCalendar = next;
    return next;
  }

  get jadwalCalendarLabel() {
    return 'Jadwal → Calendar: ' + (this.syncJadwalToCalendar ? 'ON' : 'OFF');
  }
}

/** =================================================================
 * Time helpers
 * ================================================================= */
class TimeKey {
  static pad(h, m) { return ('0' + h).slice(-2) + ':' + m; }

  static parseRange(text) {
    const m = String(text == null ? '' : text).trim().match(SLOT_RE);
    return m ? { start: TimeKey.pad(m[1], m[2]), end: TimeKey.pad(m[3], m[4]) } : null;
  }

  /** Drop a leading "HH:MM - HH:MM" line if the cell happens to carry one. */
  static stripLabel(text) {
    const lines = String(text == null ? '' : text).split('\n');
    if (lines.length && SLOT_RE.test(lines[0].trim())) lines.shift();
    return lines.join('\n').trim();
  }

  static label(slot) { return slot.start + ' - ' + slot.end; }

  /** Split a cell body into logical blocks: a Jadwal block is its header
   *  line plus the subject line beneath it; anything else is one line. */
  static blocks(body) {
    const lines = String(body == null ? '' : body).split('\n');
    const out = [];
    for (let i = 0; i < lines.length; i++) {
      if (!lines[i].trim()) continue;
      if (JADWAL_BLOCK_RE.test(lines[i].trim()) && i + 1 < lines.length) {
        out.push(lines[i].trim() + '\n' + lines[i + 1].trim());
        i++;
      } else {
        out.push(lines[i].trim());
      }
    }
    return out;
  }

  /** Drop repeated identical blocks, keeping first occurrence order. */
  static dedupe(body) {
    const seen = {};
    return TimeKey.blocks(body)
      .filter(b => (seen[b] ? false : (seen[b] = true)))
      .join('\n');
  }

  static looksJadwal(body) {
    return TimeKey.blocks(body).some(b => JADWAL_BLOCK_RE.test(b));
  }
}

/** =================================================================
 * The dated week the grid currently represents.
 * ================================================================= */
class WeekContext {
  constructor(config, state) {
    this.tz = config.TIMEZONE;
    this.fmt = config.DATE_FORMAT;
    this.state = state;
    const stored = state.get('META', 'WEEK', 'START');
    if (stored && stored.content) {
      this.weekStart = this._parseIso(String(stored.content));
    } else {
      // First use: pin to the current week straight away, so an on-edit sync
      // works without the user having run "Set this week" first.
      this.weekStart = this.currentSunday();
      state.put('META', 'WEEK', 'START', this.key, '', '');
    }
  }

  _parseIso(value) {
    if (value instanceof Date) value = Utilities.formatDate(value, this.tz, 'yyyy-MM-dd');
    return Utilities.parseDate(String(value).trim() + ' 00:00', this.tz, 'yyyy-MM-dd HH:mm');
  }

  ymd(date) { return Utilities.formatDate(date, this.tz, 'yyyy-MM-dd'); }
  pretty(date) { return Utilities.formatDate(date, this.tz, this.fmt); }

  currentSunday() {
    const now = new Date();
    const idx = DAY_NAMES.indexOf(Utilities.formatDate(now, this.tz, 'EEEE'));
    const d = new Date(now.getTime());
    d.setDate(d.getDate() - idx);
    return this._parseIso(this.ymd(d));
  }

  dateFor(dayName) {
    const idx = DAY_NAMES.indexOf(dayName);
    if (idx < 0) return null;
    const d = new Date(this.weekStart.getTime());
    d.setDate(d.getDate() + idx);
    return d;
  }

  get key() { return this.ymd(this.weekStart); }

  get rangeLabel() {
    return this.pretty(this.weekStart) + ' – ' +
      Utilities.formatDate(this.dateFor('Saturday'), this.tz, 'dd MMM yyyy');
  }

  at(date, hhmm) {
    return Utilities.parseDate(this.ymd(date) + ' ' + hhmm, this.tz, 'yyyy-MM-dd HH:mm');
  }

  /** Point the grid at a specific Sunday. */
  setWeekStart(date) {
    this.weekStart = this._parseIso(this.ymd(date));
    this.state.put('META', 'WEEK', 'START', this.key, '', '');
  }

  /** Advance a week — or jump to the real current week if further ahead. */
  advance() {
    const next = new Date(this.weekStart.getTime());
    next.setDate(next.getDate() + 7);
    const current = this.currentSunday();
    this.setWeekStart(next.getTime() > current.getTime() ? next : current);
  }
}

/** =================================================================
 * Hidden state sheet.
 * Columns: Kind | Day | Start | Content | EventId | WeekStart | Updated
 * Kinds: META (week pointer) · JADWAL (scraper-owned cells) · MANUAL (your cells)
 * ================================================================= */
class StateStore {
  constructor(spreadsheet, config) {
    this.tz = config.TIMEZONE;
    const name = config.STATE_SHEET_NAME;
    this.sheet = spreadsheet.getSheetByName(name);
    if (!this.sheet) {
      this.sheet = spreadsheet.insertSheet(name);
      this.sheet.appendRow(['Kind', 'Day', 'Start', 'Content', 'EventId', 'WeekStart', 'Updated']);
      this.sheet.hideSheet();
    }
    // Sheets silently coerces "2026-07-26" to a Date and "16:00" to a time.
    // Force plain text so keys round-trip as the strings we wrote.
    try {
      this.sheet.getRange(1, 1, Math.max(this.sheet.getMaxRows(), 1), 6).setNumberFormat('@');
    } catch (e) { /* older sheets / restricted scope — _hydrate still normalises */ }
    this.rows = this.sheet.getLastRow() > 1
      ? this.sheet.getRange(2, 1, this.sheet.getLastRow() - 1, 7).getValues()
      : [];
    this.dirty = false;
  }

  /** Coerce a cell back to the string form we wrote, whatever Sheets did to it. */
  _text(value, dateFormat) {
    if (value instanceof Date) return Utilities.formatDate(value, this.tz, dateFormat);
    return value == null ? '' : String(value).trim();
  }

  static key(kind, day, start) { return kind + '|' + day + '|' + start; }

  _indexOf(kind, day, start) {
    const k = StateStore.key(kind, day, start);
    for (let i = 0; i < this.rows.length; i++) {
      const h = this._hydrate(this.rows[i]);
      if (StateStore.key(h.kind, h.day, h.start) === k) return i;
    }
    return -1;
  }

  _hydrate(r) {
    return {
      kind: this._text(r[0], 'yyyy-MM-dd'),
      day: this._text(r[1], 'yyyy-MM-dd'),
      start: this._text(r[2], 'HH:mm'),
      content: this._text(r[3], 'yyyy-MM-dd'),
      eventId: this._text(r[4], 'yyyy-MM-dd'),
      weekStart: this._text(r[5], 'yyyy-MM-dd')
    };
  }

  get(kind, day, start) {
    const i = this._indexOf(kind, day, start);
    return i < 0 ? null : this._hydrate(this.rows[i]);
  }

  all(kind) {
    return this.rows.map(r => this._hydrate(r)).filter(h => h.kind === kind);
  }

  put(kind, day, start, content, eventId, weekStart) {
    const row = [kind, day, start, content, eventId || '', weekStart || '', new Date()];
    const i = this._indexOf(kind, day, start);
    if (i < 0) this.rows.push(row); else this.rows[i] = row;
    this.dirty = true;
  }

  remove(kind, day, start) {
    const i = this._indexOf(kind, day, start);
    if (i >= 0) { this.rows.splice(i, 1); this.dirty = true; }
  }

  removeKind(kind) {
    const before = this.rows.length;
    this.rows = this.rows.filter(r => this._text(r[0], 'yyyy-MM-dd') !== kind);
    if (this.rows.length !== before) this.dirty = true;
  }

  removeAll(kind, day) {
    const before = this.rows.length;
    this.rows = this.rows.filter(r => !(this._text(r[0], 'yyyy-MM-dd') === kind &&
                                        this._text(r[1], 'yyyy-MM-dd') === day));
    if (this.rows.length !== before) this.dirty = true;
  }

  flush() {
    if (!this.dirty) return;
    const last = this.sheet.getLastRow();
    if (last > 1) this.sheet.getRange(2, 1, last - 1, 7).clearContent();
    if (this.rows.length) this.sheet.getRange(2, 1, this.rows.length, 7).setValues(this.rows);
    this.dirty = false;
  }
}

/** =================================================================
 * Flat-log archive.
 * ================================================================= */
class Archive {
  constructor(spreadsheet, name) {
    this.sheet = spreadsheet.getSheetByName(name);
    if (!this.sheet) {
      this.sheet = spreadsheet.insertSheet(name);
      this.sheet.appendRow(ARCHIVE_HEADERS);
      this.sheet.setFrozenRows(1);
    }
    // Start/End are clock strings, not times — keep "09:25" from becoming 9:25.
    try {
      this.sheet.getRange(1, 3, Math.max(this.sheet.getMaxRows(), 1), 2).setNumberFormat('@');
    } catch (e) { /* non-fatal */ }
  }

  /** rows: [Date, Day, Start, End, Source, Entry] — Archived at is stamped here. */
  append(rows) {
    if (!rows.length) return 0;
    const stamp = new Date();
    const payload = rows.map(r => r.concat([stamp]));
    this.sheet.getRange(this.sheet.getLastRow() + 1, 1, payload.length, ARCHIVE_HEADERS.length)
      .setValues(payload);
    return payload.length;
  }
}

/** =================================================================
 * One period scraped from the site.
 * ================================================================= */
class Period {
  constructor(cols) {
    this.period = cols[0] || '';
    this.darajah = cols[1] || '';
    this.subject = cols[2] || '';
    this.time = cols[3] || '';
    this.type = cols[4] || '';
    this.status = cols[5] || '';
  }

  static fromColumns(cols) {
    return cols.length >= 4 && String(cols[3]).indexOf('-') !== -1 ? new Period(cols) : null;
  }

  get startTime() {
    const p = TimeKey.parseRange(this.time);
    return p ? p.start : null;
  }

  get isCancelled() { return this.status.toLowerCase().indexOf('cancel') !== -1; }

  render() {
    const head = ['P' + this.period, this.darajah].filter(String).join(' · ');
    return head + '\n' + this.subject + (this.isCancelled ? ' (CANCELLED)' : '');
  }

  toString() { return 'P' + this.period + ' ' + this.time + ' ' + this.subject; }
}

/** =================================================================
 * Scrapes the Jadwal page.
 * ================================================================= */
class JadwalPage {
  constructor(html) { this.html = html; }

  static fetch(url) {
    if (!url) throw new Error('No Jadwal URL set. Use 📅 Jadwal Grid → Set Jadwal URL.');
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() !== 200) {
      throw new Error('Jadwal fetch failed with HTTP ' + res.getResponseCode());
    }
    return new JadwalPage(res.getContentText());
  }

  _spanById(id) {
    const m = this.html.match(new RegExp('<span[^>]*id="' + id + '"[^>]*>([\\s\\S]*?)<\\/span>', 'i'));
    return m ? m[1].replace(/<[^>]+>/g, '').trim() : null;
  }

  _tableById(id) {
    const m = this.html.match(new RegExp('<table[^>]*id="' + id + '"[\\s\\S]*?<\\/table>', 'i'));
    return m ? m[0] : null;
  }

  _periodsFrom(tableHtml) {
    if (!tableHtml) return [];
    const out = [];
    const rowRe = /<tr[\s\S]*?>([\s\S]*?)<\/tr>/gi;
    let row, isHeader = true;
    while ((row = rowRe.exec(tableHtml)) !== null) {
      if (isHeader) { isHeader = false; continue; }
      const colRe = /<t[dh][\s\S]*?>([\s\S]*?)<\/t[dh]>/gi;
      const cols = [];
      let c;
      while ((c = colRe.exec(row[1])) !== null) {
        cols.push(c[1].replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim());
      }
      const p = Period.fromColumns(cols);
      if (p) out.push(p);
    }
    return out;
  }

  get todayDayName() { return this._spanById('litDayName'); }
  get nextDayName() { return this._spanById('litNextDayName'); }
  get todayPeriods() { return this._periodsFrom(this._tableById('gvTodaysPeriods')); }
  get nextPeriods() { return this._periodsFrom(this._tableById('gvNextPeriods')); }
}

/** =================================================================
 * The weekly grid.
 *
 * Slot geometry is structural: column A is a fine-grained time axis, and a
 * day cell merged over rows r0..r1 covers axis[r0].start -> axis[r1].end.
 * Cell text is therefore free-form — type anything into a cell.
 * ================================================================= */
class ScheduleGrid {
  constructor(sheet) {
    this.sheet = sheet;
    this.values = sheet.getDataRange().getDisplayValues();
    this.headerRow = this._findHeaderRow();
    this.dayColumns = this._findDayColumns();
    this.axis = this._buildAxis();
    if (!this.axis.filter(Boolean).length) {
      throw new Error('no "HH:MM - HH:MM" time axis in column A');
    }
    this.mergeAnchors = this._buildMergeMap();
    this.slotIndex = this._buildSlotIndex();
  }

  static locate(spreadsheet, preferredName) {
    if (preferredName) {
      const named = spreadsheet.getSheetByName(preferredName);
      if (named) return new ScheduleGrid(named);
    }
    const problems = [];
    const sheets = spreadsheet.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      try {
        const grid = new ScheduleGrid(sheets[i]);
        if (Object.keys(grid.dayColumns).length >= 5) return grid;
        problems.push(sheets[i].getName() + ': only ' +
          Object.keys(grid.dayColumns).length + ' day columns');
      } catch (e) {
        problems.push(sheets[i].getName() + ': ' + e.message);
      }
    }
    throw new Error('Could not find the weekly grid.\n' + problems.join('\n'));
  }

  static dayNameIn(cellText) {
    const first = String(cellText == null ? '' : cellText).split('\n')[0].trim();
    return DAY_NAMES.indexOf(first) !== -1 ? first : null;
  }

  _findHeaderRow() {
    for (let r = 0; r < Math.min(this.values.length, 12); r++) {
      if (this.values[r].some(v => ScheduleGrid.dayNameIn(v))) return r;
    }
    throw new Error('no header row with day names in the first 12 rows');
  }

  _findDayColumns() {
    const map = {};
    this.values[this.headerRow].forEach((v, c) => {
      const name = ScheduleGrid.dayNameIn(v);
      if (name) map[name] = c;
    });
    return map;
  }

  _buildAxis() {
    const axis = [];
    for (let r = this.headerRow + 1; r < this.values.length; r++) {
      const parsed = TimeKey.parseRange(this.values[r][0]);
      if (parsed) axis[r] = parsed;
    }
    return axis;
  }

  _buildMergeMap() {
    const map = {};
    let ranges = [];
    try { ranges = this.sheet.getDataRange().getMergedRanges(); } catch (e) { ranges = []; }
    ranges.forEach(rg => {
      const top = rg.getRow() - 1, bottom = top + rg.getNumRows() - 1;
      const left = rg.getColumn() - 1, right = left + rg.getNumColumns() - 1;
      for (let r = top; r <= bottom; r++) {
        for (let c = left; c <= right; c++) map[r + ',' + c] = { top: top, bottom: bottom };
      }
    });
    return map;
  }

  _buildSlotIndex() {
    const index = {};
    Object.keys(this.dayColumns).forEach(day => {
      const col = this.dayColumns[day];
      const slots = {};
      this.axis.forEach((interval, r) => {
        if (!interval) return;
        const merge = this.mergeAnchors[r + ',' + col];
        if (merge && merge.top !== r) return;
        const bottom = merge ? merge.bottom : r;
        const raw = this.values[r][col];
        slots[interval.start] = {
          row: r, col: col, spanTo: bottom,
          start: interval.start, end: (this.axis[bottom] || interval).end,
          raw: String(raw == null ? '' : raw),
          body: TimeKey.stripLabel(raw)
        };
      });
      index[day] = slots;
    });
    return index;
  }

  hasDay(day) { return Object.prototype.hasOwnProperty.call(this.dayColumns, day); }
  slot(day, start) { return this.hasDay(day) ? (this.slotIndex[day][start] || null) : null; }
  slots(day) { const s = this.slotIndex[day] || {}; return Object.keys(s).map(k => s[k]); }
  days() { return DAY_NAMES.filter(d => this.hasDay(d)); }
  dayAtColumn(col0) { return DAY_NAMES.find(d => this.dayColumns[d] === col0) || null; }
  slotAtRow(day, row0) {
    return this.slots(day).find(s => row0 >= s.row && row0 <= s.spanTo) || null;
  }
  range(slot) { return this.sheet.getRange(slot.row + 1, slot.col + 1); }

  /** Stamp "Monday / 03 Aug" into each day header, and the range into A1. */
  applyDates(week) {
    this.days().forEach(day => {
      this.sheet.getRange(this.headerRow + 1, this.dayColumns[day] + 1)
        .setValue(day + '\n' + week.pretty(week.dateFor(day)))
        .setWrap(true);
    });
    const title = this.sheet.getRange(1, 1);
    if (String(title.getDisplayValue()).toUpperCase().indexOf('WEEKLY SCHEDULE') === 0) {
      title.setValue('WEEKLY SCHEDULE  ·  ' + week.rangeLabel);
    }
  }

  clearSlots(day, startTimes, keepLabel) {
    startTimes.forEach(start => {
      const slot = this.slot(day, start);
      if (!slot || !slot.body) return;
      this.range(slot).setValue(keepLabel ? TimeKey.label(slot) : '');
      slot.body = '';
    });
  }

  clearAll(keepLabel) {
    let n = 0;
    this.days().forEach(day => {
      this.slots(day).forEach(slot => {
        if (!slot.body) return;
        this.range(slot).setValue(keepLabel ? TimeKey.label(slot) : '');
        slot.body = '';
        n++;
      });
    });
    return n;
  }

  /** Overwrite a slot's body outright. Never appends — appending was how a
   *  lost state row turned one period into three. */
  setSlot(slot, body, keepLabel) {
    slot.body = TimeKey.dedupe(body);
    const cell = this.range(slot);
    cell.setValue((keepLabel ? TimeKey.label(slot) + '\n' : '') +
      (slot.body ? slot.body : '')).setWrap(true).setVerticalAlignment('middle');
  }

  /** Write every period that shares a start time, as one fresh block. */
  placeAll(day, start, periods, keepLabel) {
    const slot = this.slot(day, start);
    if (!slot) return false;
    this.setSlot(slot, periods.map(p => p.render()).join('\n'), keepLabel);
    return true;
  }

  /** Slots whose current content was clearly written by the scraper. */
  jadwalLookingStarts(day) {
    return this.slots(day).filter(s => TimeKey.looksJadwal(s.body)).map(s => s.start);
  }

  noteOnDay(day, text) {
    if (!this.hasDay(day)) return;
    this.sheet.getRange(this.headerRow + 1, this.dayColumns[day] + 1).setNote(text || '');
  }
}

/** =================================================================
 * Calendar operations for one grid cell.
 * ================================================================= */
class GridCalendar {
  constructor(week, prefix) {
    this.calendar = CalendarApp.getDefaultCalendar();
    this.week = week;
    this.prefix = prefix || '';
  }

  /** Title from the most meaningful line. For a Jadwal block ("P3 · Darajah"
   *  then the subject) the subject makes the better event title. */
  _split(body) {
    const lines = body.split('\n').map(l => l.trim()).filter(String);
    if (lines.length > 1 && JADWAL_BLOCK_RE.test(lines[0])) {
      return {
        title: this.prefix + lines[1],
        rest: [lines[0]].concat(lines.slice(2)).join('\n')
      };
    }
    return { title: this.prefix + (lines[0] || 'Untitled'), rest: lines.slice(1).join('\n') };
  }

  fetch(eventId) {
    if (!eventId) return null;
    try { return this.calendar.getEventById(eventId); } catch (e) { return null; }
  }

  upsert(kind, day, slot, body, date, existingEventId) {
    const start = this.week.at(date, slot.start);
    const end = this.week.at(date, slot.end);
    const parts = this._split(body);
    const description = (parts.rest ? parts.rest + '\n\n' : '') +
      'From weekly schedule grid · ' + day + ' ' + this.week.ymd(date) + ' ' + TimeKey.label(slot);

    const event = this.fetch(existingEventId);
    if (event) {
      if (event.getTitle() !== parts.title) event.setTitle(parts.title);
      if (event.getStartTime().getTime() !== start.getTime() ||
          event.getEndTime().getTime() !== end.getTime()) {
        event.setTime(start, end);
      }
      event.setDescription(description);
      return event.getId();
    }
    const created = this.calendar.createEvent(parts.title, start, end, { description: description });
    created.setTag('gridKey', StateStore.key(kind, day, slot.start));
    return created.getId();
  }

  remove(eventId) {
    const event = this.fetch(eventId);
    if (event) { event.deleteEvent(); return true; }
    return false;
  }
}

/** =================================================================
 * Binds grid slots of ONE ownership kind to calendar events, keeping the
 * state sheet as the record of which event belongs to which slot.
 * Shared by MANUAL (what you type) and JADWAL (what the scraper writes).
 * ================================================================= */
class CalendarBinder {
  constructor(kind, state, week, calendar) {
    this.kind = kind;
    this.state = state;
    this.week = week;
    this.cal = calendar;
  }

  key(day, start) { return StateStore.key(this.kind, day, start); }

  /** Create or update the event for a filled slot. Returns a log line, or
   *  null when nothing needed doing. */
  apply(day, slot) {
    const date = this.week.dateFor(day);
    const rec = this.state.get(this.kind, day, slot.start);
    const sameWeek = rec && rec.weekStart === this.week.key;

    if (sameWeek && rec.content === slot.body && this.cal.fetch(rec.eventId)) return null;
    if (rec && !sameWeek) this.cal.remove(rec.eventId);

    const id = this.cal.upsert(this.kind, day, slot, slot.body, date,
      sameWeek ? rec.eventId : null);
    this.state.put(this.kind, day, slot.start, slot.body, id, this.week.key);
    return (rec && sameWeek ? 'Updated' : 'Created') + ': ' + day + ' ' +
      this.week.ymd(date) + ' ' + TimeKey.label(slot) + ' — ' + slot.body.split('\n')[0];
  }

  /** Record ownership WITHOUT an event, deleting one if it exists. Used when
   *  the Jadwal → Calendar toggle is off. */
  store(day, slot) {
    const rec = this.state.get(this.kind, day, slot.start);
    const had = rec && rec.eventId ? this.cal.remove(rec.eventId) : false;
    this.state.put(this.kind, day, slot.start, slot.body, '', this.week.key);
    return had ? 'Removed event: ' + day + ' ' + slot.start : null;
  }

  /** Slot is gone: delete its event and forget it. */
  release(day, start) {
    const rec = this.state.get(this.kind, day, start);
    if (!rec) return null;
    const gone = this.cal.remove(rec.eventId);
    this.state.remove(this.kind, day, start);
    return gone ? 'Deleted event: ' + day + ' ' + start : null;
  }

  /** Drop every remembered slot of this kind that no longer exists in the grid. */
  prune(seenKeys) {
    const lines = [];
    this.state.all(this.kind).forEach(rec => {
      if (seenKeys[this.key(rec.day, rec.start)]) return;
      const line = this.release(rec.day, rec.start);
      if (line) lines.push(line);
    });
    return lines;
  }

  /** Delete every event of this kind but keep the ownership rows. */
  detachAll() {
    const lines = [];
    this.state.all(this.kind).forEach(rec => {
      if (!rec.eventId) return;
      if (this.cal.remove(rec.eventId)) {
        lines.push('Removed event: ' + rec.day + ' ' + rec.start);
      }
      this.state.put(this.kind, rec.day, rec.start, rec.content, '', rec.weekStart);
    });
    return lines;
  }
}

/** =================================================================
 * Shared plumbing: open the spreadsheet, grid, state, week and calendar
 * once, for whichever ownership kind the caller cares about.
 * ================================================================= */
class SyncSession {
  constructor(config, kind) {
    this.config = config;
    this.spreadsheet = SpreadsheetApp.getActive();
    this.settings = new Settings(config);
    this.grid = ScheduleGrid.locate(this.spreadsheet, config.GRID_SHEET_NAME);
    this.state = new StateStore(this.spreadsheet, config);
    this.week = new WeekContext(config, this.state);
    this.calendar = new GridCalendar(this.week, config.EVENT_PREFIX);
    this.binder = new CalendarBinder(kind, this.state, this.week, this.calendar);
  }

  /** True when this slot belongs to the scraper — by state row, or, if that
   *  row was lost, by the shape of its content. */
  isJadwalOwned(day, slot) {
    return !!this.state.get(KIND_JADWAL, day, slot.start) || TimeKey.looksJadwal(slot.body);
  }

  finish() { this.state.flush(); }
}

/** =================================================================
 * Part A — Jadwal page into the grid (and, if enabled, into Calendar).
 * ================================================================= */
class JadwalSync {
  constructor(config) { this.config = config; this.log = []; }

  _dateAt(offset) {
    const d = new Date();
    d.setDate(d.getDate() + offset);
    return d;
  }

  _dayName(offset) {
    return Utilities.formatDate(this._dateAt(offset), this.config.TIMEZONE, 'EEEE');
  }

  run() {
    const session = new SyncSession(this.config, KIND_JADWAL);
    const grid = session.grid;
    const state = session.state;
    const week = session.week;
    const binder = session.binder;
    const keepLabel = this.config.KEEP_TIME_LABEL;
    const toCalendar = session.settings.syncJadwalToCalendar;
    const page = JadwalPage.fetch(session.settings.jadwalUrl);

    [
      { webDay: page.todayDayName, sysDay: this._dayName(0),
        realDate: week.ymd(this._dateAt(0)), periods: page.todayPeriods, tag: 'today' },
      { webDay: page.nextDayName, sysDay: this._dayName(1),
        realDate: week.ymd(this._dateAt(1)), periods: page.nextPeriods, tag: 'tomorrow' }
    ].forEach(t => {
      if (!t.webDay) { this.log.push('Skipped ' + t.tag + ': no day name on the page.'); return; }
      if (t.webDay !== t.sysDay) {
        this.log.push('Skipped ' + t.tag + ': page says ' + t.webDay + ', system says ' + t.sysDay + '.');
        return;
      }
      if (!grid.hasDay(t.webDay)) { this.log.push('No "' + t.webDay + '" column.'); return; }

      // The grid is dated, so a matching day NAME is not enough. On a Saturday
      // the page's "tomorrow" is Sunday — but this grid's Sunday column is six
      // days in the PAST, so writing there would backdate next week's classes.
      const columnDate = week.ymd(week.dateFor(t.webDay));
      if (columnDate !== t.realDate) {
        this.log.push('Skipped ' + t.webDay + ' (' + t.tag + '): grid column is ' +
          columnDate + ' but that day is ' + t.realDate +
          '. Roll over the week to catch up.');
        return;
      }

      // Clear the cells this scraper owns: those the state remembers, PLUS any
      // whose content still looks like scraper output (self-heals lost state).
      // State ROWS are kept for now so their event ids survive the rewrite;
      // rows whose slot does not come back are released further down.
      const owned = state.all(KIND_JADWAL).filter(r => r.day === t.webDay).map(r => r.start);
      const stale = grid.jadwalLookingStarts(t.webDay);
      const toClear = owned.concat(stale.filter(x => owned.indexOf(x) === -1));
      grid.clearSlots(t.webDay, toClear, keepLabel);

      // Group by start time so two periods in one slot are written together,
      // in a single overwrite rather than successive appends.
      const bySlot = {};
      const unmatched = [];
      t.periods.forEach(p => {
        if (this.config.SKIP_CANCELLED && p.isCancelled) {
          this.log.push('Cancelled, left blank: ' + t.webDay + ' ' + p);
          return;
        }
        const key = p.startTime;
        if (!key) { unmatched.push(p.toString()); return; }
        (bySlot[key] = bySlot[key] || []).push(p);
      });

      const placed = {};
      Object.keys(bySlot).forEach(start => {
        const group = bySlot[start];
        if (grid.placeAll(t.webDay, start, group, keepLabel)) {
          placed[start] = true;
          const slot = grid.slot(t.webDay, start);
          const line = toCalendar ? binder.apply(t.webDay, slot) : binder.store(t.webDay, slot);
          if (line) this.log.push(line);
        } else {
          group.forEach(p => unmatched.push(p.toString()));
        }
      });

      // Anything the scraper used to own on this day but no longer does.
      owned.forEach(start => {
        if (placed[start]) return;
        const line = binder.release(t.webDay, start);
        if (line) this.log.push(line);
      });

      if (this.config.MARK_UNMATCHED_AS_NOTE) {
        grid.noteOnDay(t.webDay, unmatched.length
          ? 'No matching slot in the grid:\n' + unmatched.join('\n') : '');
      }
      unmatched.forEach(u => this.log.push('UNMATCHED ' + t.webDay + ': ' + u));
      this.log.push(t.webDay + ' (' + t.tag + '): ' +
        (t.periods.length - unmatched.length) + '/' + t.periods.length + ' placed.');
    });

    if (!toCalendar) this.log.push('Jadwal → Calendar is OFF — periods stayed in the grid only.');

    session.finish();
    SpreadsheetApp.flush();
    this.log.forEach(l => console.log(l));
    return { log: this.log };
  }
}

/** =================================================================
 * Part B — your grid entries into Google Calendar, on the header dates.
 * ================================================================= */
class ManualSync {
  constructor(config) { this.config = config; this.log = []; }

  runAll() {
    const session = new SyncSession(this.config, KIND_MANUAL);
    const grid = session.grid;
    const binder = session.binder;

    const seen = {};
    grid.days().forEach(day => {
      grid.slots(day).forEach(slot => {
        if (!slot.body) return;
        if (session.isJadwalOwned(day, slot)) return;   // handled by JadwalCalendarSync
        seen[binder.key(day, slot.start)] = true;
        const line = binder.apply(day, slot);
        if (line) this.log.push(line);
      });
    });

    binder.prune(seen).forEach(l => this.log.push(l));

    session.finish();
    if (!this.log.length) this.log.push('Nothing to change — grid and calendar already agree.');
    this.log.forEach(l => console.log(l));
    return { log: this.log };
  }

  runCell(sheet, row1, col1) {
    const session = new SyncSession(this.config, KIND_MANUAL);
    const grid = session.grid;
    if (grid.sheet.getSheetId() !== sheet.getSheetId()) return { log: [] };

    const day = grid.dayAtColumn(col1 - 1);
    if (!day) return { log: [] };
    const slot = grid.slotAtRow(day, row1 - 1);
    if (!slot) return { log: [] };

    if (session.isJadwalOwned(day, slot)) {
      // A Jadwal cell you edited by hand: keep it in sync too, when enabled.
      const jadwal = new CalendarBinder(KIND_JADWAL, session.state, session.week, session.calendar);
      if (!session.settings.syncJadwalToCalendar) {
        this.log.push('Jadwal-owned slot, Jadwal → Calendar is OFF: ' + day + ' ' + slot.start);
      } else if (!slot.body) {
        const line = jadwal.release(day, slot.start);
        this.log.push(line || ('Cleared Jadwal slot: ' + day + ' ' + slot.start));
      } else {
        const line = jadwal.apply(day, slot);
        if (line) this.log.push(line);
      }
    } else if (!slot.body) {
      const line = session.binder.release(day, slot.start);
      if (line) this.log.push(line);
    } else {
      const line = session.binder.apply(day, slot);
      if (line) this.log.push(line);
    }

    session.finish();
    this.log.forEach(l => console.log(l));
    return { log: this.log };
  }
}

/** =================================================================
 * Part B2 — scraped Jadwal entries into Google Calendar. Catch-up pass for
 * every Jadwal-owned cell in the grid, not just today's and tomorrow's.
 * ================================================================= */
class JadwalCalendarSync {
  constructor(config) { this.config = config; this.log = []; }

  runAll() {
    const session = new SyncSession(this.config, KIND_JADWAL);
    if (!session.settings.syncJadwalToCalendar) {
      this.log.push('Jadwal → Calendar is OFF — skipped ' +
        session.state.all(KIND_JADWAL).length + ' scraped slot(s).');
      this.log.forEach(l => console.log(l));
      return { log: this.log };
    }

    const binder = session.binder;
    const seen = {};
    session.grid.days().forEach(day => {
      session.grid.slots(day).forEach(slot => {
        if (!slot.body || !session.isJadwalOwned(day, slot)) return;
        seen[binder.key(day, slot.start)] = true;
        const line = binder.apply(day, slot);
        if (line) this.log.push(line);
      });
    });

    binder.prune(seen).forEach(l => this.log.push(l));

    session.finish();
    if (!this.log.length) this.log.push('Jadwal entries already in the calendar.');
    this.log.forEach(l => console.log(l));
    return { log: this.log };
  }

  /** Toggle turned off: pull every scraped event back out of the calendar. */
  detachAll() {
    const session = new SyncSession(this.config, KIND_JADWAL);
    const lines = session.binder.detachAll();
    session.finish();
    this.log = lines.length
      ? lines.concat(['Removed ' + lines.length + ' Jadwal event(s) from the calendar.'])
      : ['No Jadwal events were in the calendar.'];
    this.log.forEach(l => console.log(l));
    return { log: this.log };
  }
}

/** =================================================================
 * Weekly rollover: archive -> wipe -> advance the dates.
 * ================================================================= */
class WeekRollover {
  constructor(config) { this.config = config; this.log = []; }

  /** Collect every filled slot as a flat archive row. */
  _harvest(grid, week, state) {
    const rows = [];
    grid.days().forEach(day => {
      const date = week.ymd(week.dateFor(day));
      grid.slots(day).forEach(slot => {
        if (!slot.body) return;
        const source = (state.get(KIND_JADWAL, day, slot.start) || TimeKey.looksJadwal(slot.body))
          ? 'Jadwal' : 'Manual';
        if (source === 'Jadwal' && !this.config.ARCHIVE_JADWAL_ROWS) return;
        rows.push([date, day, slot.start, slot.end, source, TimeKey.dedupe(slot.body)]);
      });
    });
    return rows.sort((a, b) => (a[0] + a[2]).localeCompare(b[0] + b[2]));
  }

  run() {
    const ss = SpreadsheetApp.getActive();
    const grid = ScheduleGrid.locate(ss, this.config.GRID_SHEET_NAME);
    const state = new StateStore(ss, this.config);
    const week = new WeekContext(this.config, state);
    const archive = new Archive(ss, this.config.ARCHIVE_SHEET_NAME);

    const rows = this._harvest(grid, week, state);
    const archived = archive.append(rows);
    this.log.push('Archived ' + archived + ' entr' + (archived === 1 ? 'y' : 'ies') +
      ' for week of ' + week.key + '.');

    const cleared = grid.clearAll(this.config.KEEP_TIME_LABEL);
    this.log.push('Cleared ' + cleared + ' cell' + (cleared === 1 ? '' : 's') + '.');

    // Calendar events for the archived week stay put — that week happened.
    state.removeKind(KIND_MANUAL);
    state.removeKind(KIND_JADWAL);
    grid.days().forEach(d => grid.noteOnDay(d, ''));

    const previous = week.key;
    week.advance();
    grid.applyDates(week);
    state.flush();
    SpreadsheetApp.flush();
    this.log.push('Grid re-dated: ' + previous + ' -> ' + week.key + ' (' + week.rangeLabel + ').');
    this.log.forEach(l => console.log(l));
    return { log: this.log, archived: archived };
  }
}

/** =================================================================
 * Stale-week reminder. Rollover is manual by choice, so the grid can drift
 * out of the real week — and once it does, the Jadwal auto-fill correctly
 * refuses to write (it would be backdating). This makes that visible.
 * ================================================================= */
class StaleWeekNag {
  constructor(config, state, week) {
    this.config = config;
    this.state = state;
    this.week = week;
  }

  /** How many weeks behind the real week the grid is. */
  get weeksBehind() {
    const gridMs = this.week.weekStart.getTime();
    const realMs = this.week.currentSunday().getTime();
    return Math.round((realMs - gridMs) / (7 * 24 * 60 * 60 * 1000));
  }

  get isStale() { return this.weeksBehind > 0; }

  get message() {
    const n = this.weeksBehind;
    return 'Your schedule grid is still dated ' + this.week.key + ' — ' + n +
      ' week' + (n === 1 ? '' : 's') + ' behind. The daily Jadwal fill is paused ' +
      'until you roll over, because writing to it would backdate your classes.\n\n' +
      'Fix: open the sheet and choose 📅 Jadwal Grid → Roll over week.';
  }

  _alreadyEmailedToday() {
    const rec = this.state.get('META', 'NAG', 'SENT');
    const today = this.week.ymd(new Date());
    return !!(rec && rec.content === today);
  }

  _recipient() {
    if (this.config.NAG_EMAIL) return this.config.NAG_EMAIL;
    try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; }
  }

  /** Returns a log line if it nagged, null otherwise. */
  run() {
    if (!this.config.NAG_ON_STALE_WEEK || !this.isStale) return null;

    const lines = ['STALE WEEK: grid is dated ' + this.week.key + ', ' +
      this.weeksBehind + ' week(s) behind. Roll over week to resume auto-fill.'];
    _notify(this.message, 'Roll over week', 15);

    if (this.config.NAG_EMAIL_MAX_PER_DAY > 0 && !this._alreadyEmailedToday()) {
      const to = this._recipient();
      if (to) {
        try {
          MailApp.sendEmail(to, 'Schedule grid needs rolling over',
            this.message + '\n\n' + SpreadsheetApp.getActive().getUrl());
          this.state.put('META', 'NAG', 'SENT', this.week.ymd(new Date()), '', '');
          lines.push('Reminder emailed to ' + to + '.');
        } catch (e) {
          lines.push('Could not email reminder: ' + e);
        }
      }
    }
    lines.forEach(l => console.log(l));
    return lines;
  }
}

function checkStaleWeek() {
  const ss = SpreadsheetApp.getActive();
  const state = new StateStore(ss, CONFIG);
  const week = new WeekContext(CONFIG, state);
  const nag = new StaleWeekNag(CONFIG, state, week);
  const lines = nag.run();
  state.flush();
  return { log: lines || ['Grid week is current (' + week.key + ').'], stale: nag.isStale };
}

/** =================================================================
 * Entry points
 * ================================================================= */
function setThisWeek() {
  const ss = SpreadsheetApp.getActive();
  const grid = ScheduleGrid.locate(ss, CONFIG.GRID_SHEET_NAME);
  const state = new StateStore(ss, CONFIG);
  const week = new WeekContext(CONFIG, state);
  week.setWeekStart(week.currentSunday());
  grid.applyDates(week);
  state.flush();
  return { log: ['Grid pinned to week of ' + week.key + ' (' + week.rangeLabel + ').'] };
}

/** De-duplicate repeated blocks in every grid cell, and re-register any
 *  scraper-looking cell as Jadwal-owned. Fixes grids damaged by the old
 *  append-instead-of-replace behaviour. */
function repairGrid() {
  const ss = SpreadsheetApp.getActive();
  const grid = ScheduleGrid.locate(ss, CONFIG.GRID_SHEET_NAME);
  const state = new StateStore(ss, CONFIG);
  const week = new WeekContext(CONFIG, state);
  const log = [];
  let fixed = 0, reclaimed = 0;

  grid.days().forEach(day => {
    grid.slots(day).forEach(slot => {
      if (!slot.body) return;
      const clean = TimeKey.dedupe(slot.body);
      if (clean !== slot.body) {
        const before = TimeKey.blocks(slot.body).length;
        grid.setSlot(slot, clean, CONFIG.KEEP_TIME_LABEL);
        fixed++;
        log.push(day + ' ' + TimeKey.label(slot) + ': ' + before + ' blocks -> ' +
          TimeKey.blocks(clean).length);
      }
      if (TimeKey.looksJadwal(clean) && !state.get(KIND_JADWAL, day, slot.start)) {
        state.put(KIND_JADWAL, day, slot.start, clean, '', week.key);
        state.remove(KIND_MANUAL, day, slot.start);
        reclaimed++;
      }
    });
  });

  state.flush();
  SpreadsheetApp.flush();
  log.unshift('De-duplicated ' + fixed + ' cell' + (fixed === 1 ? '' : 's') +
    '; re-tagged ' + reclaimed + ' as Jadwal-owned.');
  log.forEach(l => console.log(l));
  return { log: log };
}

function syncJadwalToGrid() { return new JadwalSync(CONFIG).run(); }

/** Catch-up pass: your cells, then the scraped ones (if the toggle is on). */
function syncGridToCalendar() {
  const manual = new ManualSync(CONFIG).runAll();
  const jadwal = new JadwalCalendarSync(CONFIG).runAll();
  return { log: manual.log.concat(jadwal.log) };
}

function rollOverWeek() { return new WeekRollover(CONFIG).run(); }

function syncEverything() {
  const nag = checkStaleWeek();
  const a = syncJadwalToGrid();
  const b = syncGridToCalendar();
  return { log: (nag.stale ? nag.log : []).concat(a.log, b.log) };
}

function _notify(message, title, seconds) {
  try { SpreadsheetApp.getActive().toast(message, title || 'Calendar', seconds || 5); } catch (e) {}
}

function _toast(message, title) {
  if (!CONFIG.TOAST_ON_EDIT) return;
  _notify(message, title);
}

/**
 * Installable onEdit trigger — instant per-cell calendar sync.
 * Fires on every cell you type into, including deletes and pastes.
 * Script-driven writes do NOT fire it, so there is no feedback loop.
 */
function onGridEdit(e) {
  if (!e || !e.range) return;
  try {
    const range = e.range;
    const multi = range.getNumRows() * range.getNumColumns() > 1;
    // A paste or block delete covers many cells; one full pass is cheaper
    // than re-reading the grid once per cell, and it's idempotent anyway.
    const res = multi
      ? syncGridToCalendar()
      : new ManualSync(CONFIG).runCell(range.getSheet(), range.getRow(), range.getColumn());
    if (res.log.length) _toast(res.log.slice(0, 6).join('\n'));
  } catch (err) {
    console.error('onGridEdit failed: ' + err + (err.stack ? '\n' + err.stack : ''));
    _toast('Sync failed: ' + err, 'Error');
  }
}

function diagnoseGrid() {
  const out = [];
  const ss = SpreadsheetApp.getActive();
  try {
    const settings = new Settings(CONFIG);
    const grid = ScheduleGrid.locate(ss, CONFIG.GRID_SHEET_NAME);
    const state = new StateStore(ss, CONFIG);
    const week = new WeekContext(CONFIG, state);
    out.push('Grid sheet: "' + grid.sheet.getName() + '", header on row ' + (grid.headerRow + 1));
    out.push('Jadwal URL: ' + settings.jadwalUrl +
      (settings.jadwalUrlIsCustom ? '  (custom)' : '  (built-in default)'));
    out.push(settings.jadwalCalendarLabel);
    out.push('Grid week: ' + week.key + '  (' + week.rangeLabel + ')');
    out.push('Real current week: ' + week.ymd(week.currentSunday()) +
      (week.ymd(week.currentSunday()) === week.key ? '  — in sync' : '  — ROLL OVER NEEDED'));
    out.push('Time axis rows in column A: ' + grid.axis.filter(Boolean).length);
    out.push('Jadwal slots tracked: ' + state.all(KIND_JADWAL).length +
      ', with an event: ' + state.all(KIND_JADWAL).filter(r => r.eventId).length);
    out.push('Manual slots tracked: ' + state.all(KIND_MANUAL).length +
      ', with an event: ' + state.all(KIND_MANUAL).filter(r => r.eventId).length);
    out.push('');
    grid.days().forEach(day => {
      const filled = grid.slots(day).filter(s => s.body);
      out.push(day + ' ' + week.pretty(week.dateFor(day)) + ': ' + grid.slots(day).length +
        ' slots, ' + filled.length + ' filled' + (filled.length ? ' -> ' +
        filled.map(s => TimeKey.label(s) + ' "' + s.body.split('\n')[0].slice(0, 22) + '"').join('; ') : ''));
    });
    out.push('');
    out.push('Default calendar: ' + CalendarApp.getDefaultCalendar().getName());
    const arch = ss.getSheetByName(CONFIG.ARCHIVE_SHEET_NAME);
    out.push('Archive rows: ' + (arch ? Math.max(0, arch.getLastRow() - 1) : 'sheet not created yet'));
    const triggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
    out.push('Sync on edit: ' + (triggers.indexOf('onGridEdit') !== -1
      ? 'ON' : 'OFF — run "Install triggers"'));
    out.push('Triggers: ' + (triggers.length ? triggers.join(', ') : 'none'));
  } catch (e) {
    out.push('ERROR: ' + e.message);
  }
  const text = out.join('\n');
  console.log(text);
  try { SpreadsheetApp.getUi().alert('Diagnostics', text, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
  return text;
}

function _alert(title, fn) {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = fn();
    ui.alert(title, res.log.join('\n') || 'Nothing to do.', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert(title + ' failed', String(err), ui.ButtonSet.OK);
  }
}

function menuSetWeek() { _alert('Set this week', setThisWeek); }
function menuSyncJadwal() { _alert('Jadwal → Grid', syncJadwalToGrid); }
function menuSyncCalendar() { _alert('Grid → Calendar', syncGridToCalendar); }
function menuSyncBoth() { _alert('Full sync', syncEverything); }
function menuRepair() { _alert('Repair grid', repairGrid); }

/** Ask for the MyJadwal link and remember it on this spreadsheet. */
function menuSetJadwalUrl() {
  const ui = SpreadsheetApp.getUi();
  const settings = new Settings(CONFIG);
  const res = ui.prompt('Set Jadwal URL',
    'Current URL:\n' + settings.jadwalUrl +
    (settings.jadwalUrlIsCustom ? '\n(custom)' : '\n(built-in default)') +
    '\n\nPaste your MyJadwal link, or type "default" to restore the built-in one.\n' +
    'Leave blank and press OK to keep the current URL.',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const input = String(res.getResponseText() || '').trim();
  if (!input) { ui.alert('Set Jadwal URL', 'Unchanged:\n' + settings.jadwalUrl, ui.ButtonSet.OK); return; }

  try {
    if (input.toLowerCase() === 'default') {
      settings.resetJadwalUrl();
      ui.alert('Set Jadwal URL', 'Restored the built-in URL:\n' + settings.jadwalUrl, ui.ButtonSet.OK);
    } else {
      settings.jadwalUrl = input;
      ui.alert('Set Jadwal URL', 'Saved:\n' + settings.jadwalUrl +
        '\n\nRun "Fill grid from Jadwal" to test it.', ui.ButtonSet.OK);
    }
  } catch (err) {
    ui.alert('Set Jadwal URL failed', String(err.message || err), ui.ButtonSet.OK);
  }
}

/** Turn scraped-period calendar events on or off, applying the change now. */
function menuToggleJadwalCalendar() {
  const ui = SpreadsheetApp.getUi();
  const settings = new Settings(CONFIG);
  const turningOn = !settings.syncJadwalToCalendar;

  const question = turningOn
    ? 'Turn Jadwal → Calendar ON?\n\nEvery scraped period in the grid becomes an event ' +
      'on your default calendar, on the date in its column header.\n\nContinue?'
    : 'Turn Jadwal → Calendar OFF?\n\nEvents already created from scraped periods will be ' +
      'DELETED from your calendar. Your own typed entries are untouched.\n\nContinue?';
  if (ui.alert('Jadwal → Calendar', question, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  try {
    settings.toggleJadwalCalendar();
    const res = turningOn
      ? new JadwalCalendarSync(CONFIG).runAll()
      : new JadwalCalendarSync(CONFIG).detachAll();
    ui.alert('Jadwal → Calendar is now ' + (turningOn ? 'ON' : 'OFF'),
      res.log.join('\n') + '\n\n(The menu label updates next time the sheet is opened.)',
      ui.ButtonSet.OK);
  } catch (err) {
    settings.syncJadwalToCalendar = !turningOn;   // roll the flag back
    ui.alert('Jadwal → Calendar failed', String(err), ui.ButtonSet.OK);
  }
}

function menuRollOver() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  let summary = '';
  try {
    const grid = ScheduleGrid.locate(ss, CONFIG.GRID_SHEET_NAME);
    const state = new StateStore(ss, CONFIG);
    const week = new WeekContext(CONFIG, state);
    const filled = grid.days().reduce((n, d) => n + grid.slots(d).filter(s => s.body).length, 0);
    summary = 'Archive ' + filled + ' entries from the week of ' + week.key +
      ', then wipe the grid and re-date it.\n\nCalendar events already created are kept.\n\nContinue?';
  } catch (e) {
    ui.alert('Roll over week failed', String(e), ui.ButtonSet.OK);
    return;
  }
  if (ui.alert('Roll over week', summary, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  _alert('Roll over week', rollOverWeek);
}

function installTriggers() {
  const mine = ['syncJadwalToGrid', 'syncGridToCalendar', 'syncEverything', 'onGridEdit'];
  ScriptApp.getProjectTriggers()
    .filter(t => mine.indexOf(t.getHandlerFunction()) !== -1)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('syncEverything')
    .timeBased().everyDays(1).atHour(6).nearMinute(0)
    .inTimezone(CONFIG.TIMEZONE).create();

  ScriptApp.newTrigger('onGridEdit')
    .forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();

  SpreadsheetApp.getUi().alert('Triggers installed',
    'Sync on edit: ON — type in any grid cell and the event is created, ' +
    'updated, or deleted immediately. A corner toast confirms it.\n\n' +
    'Daily 6 AM: Jadwal scrape + a catch-up calendar sync.\n\n' +
    'Rollover stays manual — use "Roll over week".',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function onOpen() {
  let toggleLabel = 'Jadwal → Calendar: OFF';
  try {
    const ss = SpreadsheetApp.getActive();
    toggleLabel = new Settings(CONFIG).jadwalCalendarLabel;
    const state = new StateStore(ss, CONFIG);
    const week = new WeekContext(CONFIG, state);
    const nag = new StaleWeekNag(CONFIG, state, week);
    if (nag.isStale) _notify(nag.message, 'Roll over week', 15);
  } catch (e) { /* never block the menu */ }

  SpreadsheetApp.getUi().createMenu('📅 Jadwal Grid')
    .addItem('Turn on sync-on-edit', 'installTriggers')
    .addItem('Set Jadwal URL', 'menuSetJadwalUrl')
    .addItem(toggleLabel + ' (click to toggle)', 'menuToggleJadwalCalendar')
    .addItem('Set this week', 'menuSetWeek')
    .addSeparator()
    .addItem('Catch-up sync to Calendar', 'menuSyncCalendar')
    .addItem('Fill grid from Jadwal', 'menuSyncJadwal')
    .addItem('Do both', 'menuSyncBoth')
    .addSeparator()
    .addItem('Roll over week', 'menuRollOver')
    .addSeparator()
    .addItem('Repair grid', 'menuRepair')
    .addItem('Diagnose', 'diagnoseGrid')
    .addToUi();
}
