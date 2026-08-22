/**
 * ==================================================================
 * JADWAL -> DATED WEEKLY GRID -> GOOGLE CALENDAR   (v4)
 * ------------------------------------------------------------------
 * The grid is DATED. Its header reads "Monday / 03 Aug", and the week it
 * represents is stored in the hidden state sheet. Nothing is inferred from
 * "is this day past or future" — a cell's date is whatever the header says.
 *
 *   Set this week      pin the grid to the current Sunday-Saturday week
 *   Sync to Calendar   every cell YOU own -> an event on its header date
 *   Roll over week     archive -> wipe the grid -> advance the dates
 *
 * ------------------------------------------------------------------
 * NEW IN v4 — OWNERSHIP
 * ------------------------------------------------------------------
 * Two things write into this grid: the Jadwal scraper, and you. v4 tells
 * them apart and treats them differently:
 *
 *   · cells the scraper wrote      -> stay OUT of Google Calendar
 *   · a cell YOU typed into        -> becomes yours, and is created,
 *                                     updated or deleted in Google
 *                                     Calendar immediately — even if the
 *                                     scraper had filled it first
 *
 * The signal is AUTHORSHIP, not content. Google's onEdit trigger fires
 * only for edits a person makes through the Sheets UI; script-driven
 * writes never fire it. So an edit that reaches onGridEdit is, by
 * definition, yours: it records an OVERRIDE row and the slot changes hands.
 *
 * When a later fetch finds the site disagreeing with a slot you edited,
 * nothing is overwritten silently. Your version is held and the clash is
 * queued for a decision:
 *
 *   · fetch you started from the menu -> a dialog asks you there and then
 *   · the 6 AM fetch (no UI exists in -> a note on the cell + one email,
 *     an unattended trigger)             waiting in "Review conflicts"
 *
 * "Keep mine" is remembered, so the same clash never nags you twice.
 * "Take Jadwal's" restores the site's text, hands the slot back to the
 * scraper and removes the calendar event your edit had created.
 *
 * Ownership lasts a week. "Roll over week" clears every override, dismissal
 * and conflict; "Set this week" carries them to the new dates instead, since
 * it re-dates the grid without wiping it. "Hand my edits back to Jadwal"
 * releases them on demand.
 *
 * Two deliberate exceptions to "an edit makes it yours":
 *   · clearing several slots at once is a wipe, not a decision about each —
 *     it resets ownership rather than taking it
 *   · a slot already holding text that is not the scraper's is protected
 *     from the fetch even without a claim, so nothing you typed is
 *     overwritten silently when sync-on-edit is off
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
  JADWAL_URL: 'https://jameasaifiyah.org/MyJadwal.aspx?ID=MTg5Ng%3d%3d',
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
  NAG_EMAIL_MAX_PER_DAY: 1,         // 0 = toast/log only, never email

  // --- Ownership of a slot the scraper filled ---------------------------
  SYNC_MY_JADWAL_EDITS: true,       // master switch. ON: a Jadwal cell you edit
                                    // by hand becomes yours and syncs to
                                    // Calendar. OFF: v3 behaviour, Jadwal-shaped
                                    // cells are ignored by the calendar sync.
  ASK_BEFORE_OVERWRITING_MY_EDITS: true,  // show the review dialog when the fetch
                                          // was started by you from the menu
  CONFLICT_EMAIL: true,             // email when an unattended fetch holds an edit
  CONFLICT_NOTE_PREFIX: '⚠ Jadwal differs'  // only notes starting with this are
                                            // ever cleared by the script
});

const DAY_NAMES = Object.freeze(
  ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']);
const SLOT_RE = /^(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})\s*$/;
const ARCHIVE_HEADERS = ['Date', 'Day', 'Start', 'End', 'Source', 'Entry', 'Archived at'];
// The first line of a Jadwal-rendered block is a period label: "P2 · <darajah>",
// or "P2" when the darajah column came back blank, or "P · <darajah>" when the
// period number did. Deliberately narrow: it takes a period NUMBER, or else the
// separator, so ordinary two-line human cells — "Prep / Balaghat revision",
// "Pray / Zohr salah", "Page 42", "Practice · Nahw" — are not caught by it.
// This is the content-based ownership fallback that stops a lost state row from
// making the scraper's output look like something you typed. See
// TimeKey.looksJadwal, which also requires the subject line beneath the label.
const JADWAL_HEAD_RE = /^P\d+[A-Za-z]?\s*(·|$)|^P\s*·/;

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
      if (JADWAL_HEAD_RE.test(lines[i].trim()) && i + 1 < lines.length) {
        out.push(lines[i].trim() + '\n' + lines[i + 1].trim());
        i++;
      } else {
        out.push(lines[i].trim());
      }
    }
    return out;
  }

  /** Drop repeated identical blocks, keeping first occurrence order.
   *  A null-prototype map, or a cell line reading "constructor" vanishes. */
  static dedupe(body) {
    const seen = Object.create(null);
    return TimeKey.blocks(body)
      .filter(b => (seen[b] ? false : (seen[b] = true)))
      .join('\n');
  }

  /** Scraper output is a period label PLUS the subject line under it. The
   *  label alone is not enough: "P3" on its own is something a person types,
   *  and reading it as the scraper's would strip the cell of protection. */
  static looksJadwal(body) {
    return TimeKey.blocks(body).some(b => {
      const lines = b.split('\n');
      return lines.length > 1 && JADWAL_HEAD_RE.test(lines[0].trim());
    });
  }

  /** Do two cell bodies say the same thing? Whitespace and block order
   *  aside, this is what "the site disagrees with you" is measured on. */
  static same(a, b) { return TimeKey.dedupe(a) === TimeKey.dedupe(b); }
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

  /** "Monday 03 Aug" — for logs, notes and the review dialog. */
  stamp(dayName) {
    const d = this.dateFor(dayName);
    return d ? dayName + ' ' + this.pretty(d) : dayName;
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
 *
 * Kinds:
 *   META       week pointer, e-mail throttles
 *   JADWAL     cells the scraper owns
 *   MANUAL     cells of yours that have a calendar event (EventId)
 *   OVERRIDE   cells YOU authored this week — content = what you wrote
 *              ('' when you deliberately cleared the slot)
 *   DISMISSED  a site version you rejected — content = that version, so the
 *              same clash is never queued twice
 *   CONFLICT   a decision still owed — content = the site's incoming text
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
    const row = [kind, day, start, content == null ? '' : content,
                 eventId || '', weekStart || '', new Date()];
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

  /** Move every row of a kind to another week — used when the grid is
   *  re-dated with its contents intact, so its bookkeeping moves too. */
  restamp(kind, weekStart) {
    let n = 0;
    this.rows.forEach(r => {
      if (this._text(r[0], 'yyyy-MM-dd') !== kind) return;
      if (this._text(r[5], 'yyyy-MM-dd') === weekStart) return;
      r[5] = weekStart;
      r[6] = new Date();
      n++;
      this.dirty = true;
    });
    return n;
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

  /** Unchanged since v3 — grids in the wild are full of these strings, and
   *  JADWAL_HEAD_RE has to keep recognising every shape this can produce. */
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

  /** Every (day, slot) the given 1-based rectangle touches, by row overlap. */
  slotsIn(row1, col1, numRows, numCols) {
    const top = row1 - 1, bottom = row1 - 1 + Math.max(1, numRows) - 1;
    const out = [];
    for (let c = col1; c < col1 + Math.max(1, numCols); c++) {
      const day = this.dayAtColumn(c - 1);
      if (!day) continue;
      this.slots(day).forEach(slot => {
        if (slot.spanTo < top || slot.row > bottom) return;
        out.push({ day: day, slot: slot });
      });
    }
    return out;
  }

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

  slotNote(slot) {
    try { return String(this.range(slot).getNote() || ''); } catch (e) { return ''; }
  }

  setSlotNote(slot, text) {
    try { this.range(slot).setNote(text || ''); } catch (e) { /* non-fatal */ }
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

  /** First line is the title — except for a Jadwal-shaped block, where the
   *  first line is "P2 · Darajah" and the subject beneath it is the real
   *  title. Matters now that an edited Jadwal cell becomes an event. */
  _split(body) {
    const lines = String(body == null ? '' : body).split('\n').map(l => l.trim()).filter(String);
    if (lines.length > 1 && JADWAL_HEAD_RE.test(lines[0])) {
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

  upsert(day, slot, body, date, existingEventId) {
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
    created.setTag('gridKey', StateStore.key('MANUAL', day, slot.start));
    return created.getId();
  }

  remove(eventId) {
    const event = this.fetch(eventId);
    if (event) { event.deleteEvent(); return true; }
    return false;
  }
}

/** =================================================================
 * Who owns a slot — the scraper, or you?
 *
 * Authorship, not content, decides. The onEdit trigger fires only for
 * edits a person makes in the Sheets UI (script writes never fire it), so
 * anything arriving through it is yours: claim() records an OVERRIDE row
 * and the slot changes hands for the rest of the week.
 *
 * Resolution order: your override wins, then the scraper's own state row,
 * then the "P<n> ·" content shape as a last-resort fallback.
 * ================================================================= */
class SlotOwnership {
  static get MINE() { return 'Manual'; }
  static get JADWAL() { return 'Jadwal'; }
  static get HANDED_BACK() { return '(handed back)'; }

  constructor(config, state, week) {
    this.enabled = config.SYNC_MY_JADWAL_EDITS !== false;
    this.state = state;
    this.week = week;
  }

  /** A row only counts for the week the grid currently shows. */
  _live(kind, day, start) {
    const rec = this.state.get(kind, day, start);
    return rec && rec.weekStart === this.week.key ? rec : null;
  }

  override(day, start) { return this.enabled ? this._live('OVERRIDE', day, start) : null; }
  isMineByHand(day, start) { return !!this.override(day, start); }

  /** A person just wrote this slot. body '' means they deliberately cleared it. */
  claim(day, start, body) {
    if (!this.enabled) return false;
    const had = this.override(day, start);
    this.state.put('OVERRIDE', day, start, body || '', '', this.week.key);
    this.state.remove('JADWAL', day, start);
    return !had;
  }

  /** Forget that you authored this slot. Enough on its own only for an empty
   *  cell: one still holding your text would be read as yours again by the
   *  content fallback, so use release() to hand a filled slot over. */
  forget(day, start) {
    this.state.remove('OVERRIDE', day, start);
    this.state.remove('DISMISSED', day, start);
  }

  /** The scraper's own text is now in the cell — a plain handover. */
  takeOver(day, start, body) {
    this.forget(day, start);
    this.state.put('JADWAL', day, start,
      String(body || '').split('\n')[0], '', this.week.key);
  }

  /** Hand a slot back although YOUR text is still sitting in it. Marked, so
   *  the Archive still credits the entry to you if the scraper never refills
   *  the slot — the archive is the only record of authorship that outlives
   *  the week. */
  release(day, start, body) {
    this.forget(day, start);
    this.state.put('JADWAL', day, start,
      SlotOwnership.HANDED_BACK + (body ? ' ' + String(body).split('\n')[0] : ''),
      '', this.week.key);
  }

  wasHandedBack(day, start) {
    const rec = this._live('JADWAL', day, start);
    return !!(rec && rec.content.indexOf(SlotOwnership.HANDED_BACK) === 0);
  }

  /** Who to credit for what the cell holds now. Like ownerOf, except that a
   *  slot you handed back and the scraper never refilled is still your work. */
  creditFor(day, slot) {
    if (this.ownerOf(day, slot) === SlotOwnership.MINE) return SlotOwnership.MINE;
    return this.wasHandedBack(day, slot.start) ? SlotOwnership.MINE : SlotOwnership.JADWAL;
  }

  /** Every slot you authored, handed back at once. Returns how many. */
  handBackAll(grid) {
    const mine = this.state.all('OVERRIDE').filter(r => r.weekStart === this.week.key);
    mine.forEach(r => {
      const slot = grid ? grid.slot(r.day, r.start) : null;
      this.release(r.day, r.start, slot ? slot.body : '');
    });
    this.state.removeKind('OVERRIDE');
    this.state.removeKind('DISMISSED');
    this.state.removeKind('CONFLICT');
    return mine.length;
  }

  /** Drop the bookkeeping outright — for a wipe, where nothing is left to own. */
  forgetAll() {
    ['OVERRIDE', 'DISMISSED', 'CONFLICT'].forEach(k => this.state.removeKind(k));
  }

  ownerOf(day, slot) {
    if (this.isMineByHand(day, slot.start)) return SlotOwnership.MINE;
    if (this.state.get('JADWAL', day, slot.start)) return SlotOwnership.JADWAL;
    return TimeKey.looksJadwal(slot.body) ? SlotOwnership.JADWAL : SlotOwnership.MINE;
  }

  isJadwal(day, slot) { return this.ownerOf(day, slot) === SlotOwnership.JADWAL; }
  isMine(day, slot) { return !this.isJadwal(day, slot); }

  /** Start times on this day that you authored this week. */
  myStarts(day) {
    if (!this.enabled) return [];
    return this.state.all('OVERRIDE')
      .filter(r => r.day === day && r.weekStart === this.week.key)
      .map(r => r.start);
  }

  myCount() { return this.enabled ? this.state.all('OVERRIDE')
    .filter(r => r.weekStart === this.week.key).length : 0; }

  /** A site version you already rejected for this slot. */
  dismissed(day, start) { return this._live('DISMISSED', day, start); }
  dismiss(day, start, incoming) {
    this.state.put('DISMISSED', day, start, incoming || '', '', this.week.key);
  }
}

/** =================================================================
 * Decisions still owed: "the site says X, you wrote Y".
 * ================================================================= */
class ConflictQueue {
  constructor(state, week) { this.state = state; this.week = week; }

  all() {
    return this.state.all('CONFLICT')
      .filter(r => r.weekStart === this.week.key)
      .map(r => ({ day: r.day, start: r.start, incoming: r.content }))
      .sort((a, b) => (DAY_NAMES.indexOf(a.day) + a.start)
        .localeCompare(DAY_NAMES.indexOf(b.day) + b.start));
  }

  get count() { return this.all().length; }

  get(day, start) {
    const r = this.state.get('CONFLICT', day, start);
    return r && r.weekStart === this.week.key ? r : null;
  }

  record(day, start, incoming) {
    this.state.put('CONFLICT', day, start, incoming, '', this.week.key);
  }

  drop(day, start) { this.state.remove('CONFLICT', day, start); }
  dropDay(day) { this.state.removeAll('CONFLICT', day); }
  clear() { this.state.removeKind('CONFLICT'); }
}

/** =================================================================
 * How a held-back clash reaches you: a note on the cell always, plus one
 * email when the fetch ran unattended and no dialog could be shown.
 * ================================================================= */
class ConflictReporter {
  constructor(config, state, week, grid) {
    this.config = config;
    this.state = state;
    this.week = week;
    this.grid = grid;
  }

  get prefix() { return this.config.CONFLICT_NOTE_PREFIX || '⚠ Jadwal differs'; }

  _isOurs(note) { return String(note || '').indexOf(this.prefix) === 0; }

  /** Only ever clears a note the script itself wrote. */
  clearNote(slot) {
    if (this._isOurs(this.grid.slotNote(slot))) this.grid.setSlotNote(slot, '');
  }

  markNote(day, slot, incoming) {
    this.grid.setSlotNote(slot, this.prefix + ' — it says:\n' + incoming +
      '\n\nYour version is being kept. Decide with\n📅 Jadwal Grid → Review conflicts.');
  }

  /** Put notes on the clashing slots of one day, clear ours from the rest. */
  refreshNotes(day, conflicts) {
    const pending = {};
    conflicts.filter(c => c.day === day).forEach(c => { pending[c.start] = c.incoming; });
    this.grid.slots(day).forEach(slot => {
      if (Object.prototype.hasOwnProperty.call(pending, slot.start)) {
        this.markNote(day, slot, pending[slot.start]);
      } else {
        this.clearNote(slot);
      }
    });
  }

  clearAllNotes() {
    this.grid.days().forEach(day => this.grid.slots(day).forEach(s => this.clearNote(s)));
  }

  static hash(text) {
    let h = 0;
    const s = String(text == null ? '' : text);
    for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
    return String(h);
  }

  _recipient() {
    if (this.config.NAG_EMAIL) return this.config.NAG_EMAIL;
    try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; }
  }

  describe(conflicts) {
    return conflicts.map(c => '• ' + this.week.stamp(c.day) + ' ' + c.start +
      '\n    Jadwal: ' + String(c.incoming).replace(/\n/g, ' / ')).join('\n');
  }

  /** One email per distinct set of clashes per day. Returns a log line or null. */
  email(conflicts) {
    if (!this.config.CONFLICT_EMAIL || !conflicts.length) return null;
    const signature = this.week.ymd(new Date()) + '#' + ConflictReporter.hash(
      conflicts.map(c => c.day + '|' + c.start + '|' + c.incoming).join(';;'));
    const seen = this.state.get('META', 'CONFLICT', 'MAILED');
    if (seen && seen.content === signature) return null;

    const to = this._recipient();
    if (!to) return 'No address to email about the held edits.';
    const body = 'The Jadwal site now says something different for ' + conflicts.length +
      ' slot' + (conflicts.length === 1 ? '' : 's') + ' you edited by hand. Your version ' +
      'has been kept and nothing was overwritten.\n\n' + this.describe(conflicts) +
      '\n\nDecide each one with 📅 Jadwal Grid → Review conflicts.\n\n' +
      SpreadsheetApp.getActive().getUrl();
    try {
      MailApp.sendEmail(to, 'Jadwal differs from your edits (' + conflicts.length + ')', body);
      this.state.put('META', 'CONFLICT', 'MAILED', signature, '', '');
      return 'Emailed ' + to + ' about ' + conflicts.length + ' held edit(s).';
    } catch (e) {
      return 'Could not email about the held edits: ' + e;
    }
  }
}

/** =================================================================
 * The review dialog. Built as an inline HTML string so the whole project
 * stays one file. Shown modally; the Apply button posts the decisions back
 * through applyJadwalConflictDecisions().
 * ================================================================= */
class ConflictDialog {
  constructor(conflicts, week) {
    this.conflicts = conflicts;
    this.week = week;
  }

  static escape(text) {
    return String(text == null ? '' : text)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  static multiline(text) { return ConflictDialog.escape(text).replace(/\n/g, '<br>'); }

  _card(c, i) {
    const yours = c.yours
      ? ConflictDialog.multiline(c.yours)
      : '<em>you cleared this slot</em>';
    const n = 'c' + i;
    return '<div class="card">' +
      '<div class="when">' + ConflictDialog.escape(this.week.stamp(c.day)) +
        ' &middot; ' + ConflictDialog.escape(c.start) + '</div>' +
      '<div class="pair">' +
        '<label class="opt"><input type="radio" name="' + n + '" value="mine" checked>' +
          '<span><b>Keep mine</b><div class="txt">' + yours + '</div></span></label>' +
        '<label class="opt"><input type="radio" name="' + n + '" value="jadwal">' +
          '<span><b>Take Jadwal&#39;s</b><div class="txt">' +
            ConflictDialog.multiline(c.incoming) + '</div></span></label>' +
      '</div></div>';
  }

  html() {
    const payload = JSON.stringify(this.conflicts.map(c =>
      ({ day: c.day, start: c.start }))).replace(/</g, '\\u003c');
    return '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">' +
      '<style>' +
      'body{font:13px/1.45 Roboto,Arial,sans-serif;margin:0;padding:14px;color:#202124}' +
      'p.lead{margin:0 0 12px;color:#5f6368}' +
      '.card{border:1px solid #dadce0;border-radius:8px;padding:10px;margin-bottom:10px}' +
      '.when{font-weight:600;margin-bottom:8px}' +
      '.pair{display:flex;gap:10px}' +
      '.opt{flex:1;display:flex;gap:7px;align-items:flex-start;border:1px solid #e8eaed;' +
        'border-radius:6px;padding:8px;cursor:pointer}' +
      '.opt:hover{background:#f8f9fa}' +
      '.txt{color:#5f6368;margin-top:3px;white-space:normal}' +
      '.bar{position:sticky;bottom:0;background:#fff;padding-top:10px;' +
        'border-top:1px solid #eee;display:flex;gap:8px;align-items:center}' +
      'button{font:13px Roboto,Arial,sans-serif;padding:7px 14px;border-radius:4px;' +
        'border:1px solid #dadce0;background:#fff;cursor:pointer}' +
      'button.primary{background:#1a73e8;color:#fff;border-color:#1a73e8}' +
      'button[disabled]{opacity:.5;cursor:default}' +
      '#status{color:#5f6368;margin-left:auto}' +
      '.quick{margin:0 0 10px;display:flex;gap:8px}' +
      '</style></head><body>' +
      '<p class="lead">The site now says something different for ' + this.conflicts.length +
        ' slot' + (this.conflicts.length === 1 ? '' : 's') + ' you edited. ' +
        'Nothing has been overwritten. Choose per slot:</p>' +
      '<div class="quick"><button onclick="setAll(\'mine\')">Keep all mine</button>' +
        '<button onclick="setAll(\'jadwal\')">Take all Jadwal&#39;s</button></div>' +
      this.conflicts.map((c, i) => this._card(c, i)).join('') +
      '<div class="bar"><button class="primary" id="apply" onclick="apply()">Apply</button>' +
        '<button onclick="google.script.host.close()">Cancel</button>' +
        '<span id="status"></span></div>' +
      '<script>' +
      'var CONFLICTS=' + payload + ';' +
      'function setAll(v){for(var i=0;i<CONFLICTS.length;i++){' +
        'var el=document.querySelector(\'input[name="c\'+i+\'"][value="\'+v+\'"]\');' +
        'if(el)el.checked=true;}}' +
      'function apply(){var out=CONFLICTS.map(function(c,i){' +
        'var sel=document.querySelector(\'input[name="c\'+i+\'"]:checked\');' +
        'return{day:c.day,start:c.start,choice:sel?sel.value:"mine"};});' +
        'document.getElementById("apply").disabled=true;' +
        'document.getElementById("status").textContent="Applying\\u2026";' +
        'google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
        '.withFailureHandler(function(e){' +
          'document.getElementById("status").textContent="Failed: "+e.message;' +
          'document.getElementById("apply").disabled=false;})' +
        '.applyJadwalConflictDecisions(out);}' +
      '</script></body></html>';
  }

  /** Returns true if the dialog was actually shown. */
  show() {
    if (!this.conflicts.length) return false;
    try {
      const height = Math.min(560, 190 + this.conflicts.length * 118);
      SpreadsheetApp.getUi().showModalDialog(
        HtmlService.createHtmlOutput(this.html()).setWidth(620).setHeight(height),
        'Jadwal differs from your edits');
      return true;
    } catch (e) {
      // No UI in this context (time-driven trigger) — caller falls back to
      // the cell note and the email.
      console.log('Conflict dialog unavailable: ' + e);
      return false;
    }
  }
}

/** =================================================================
 * One read of everything the syncs need, shared by all of them so a run
 * reads the grid and the state sheet once. Calendar is lazy: onOpen and
 * other low-authorisation contexts must not touch CalendarApp.
 * ================================================================= */
class SyncSession {
  constructor(config, spreadsheet) {
    this.config = config;
    this.ss = spreadsheet || SpreadsheetApp.getActive();
    this.grid = ScheduleGrid.locate(this.ss, config.GRID_SHEET_NAME);
    this.state = new StateStore(this.ss, config);
    this.week = new WeekContext(config, this.state);
    this.ownership = new SlotOwnership(config, this.state, this.week);
    this.conflicts = new ConflictQueue(this.state, this.week);
    this.reporter = new ConflictReporter(config, this.state, this.week, this.grid);
    this._calendar = null;
  }

  get calendar() {
    if (!this._calendar) this._calendar = new GridCalendar(this.week, this.config.EVENT_PREFIX);
    return this._calendar;
  }

  /** Pending clashes, each carrying the text currently in the cell. */
  pendingConflicts() {
    return this.conflicts.all().map(c => {
      const slot = this.grid.slot(c.day, c.start);
      return { day: c.day, start: c.start, incoming: c.incoming,
               yours: slot ? slot.body : '' };
    });
  }

  flush() { this.state.flush(); }
}

/** =================================================================
 * Part A — Jadwal page into the grid.
 *
 * Slots you authored are never cleared and never overwritten. If the site
 * disagrees with one, the disagreement is queued instead.
 * ================================================================= */
class JadwalSync {
  constructor(config, options) {
    this.config = config;
    this.interactive = !!(options && options.interactive);
    this.log = [];
  }

  _dateAt(offset) {
    const d = new Date();
    d.setDate(d.getDate() + offset);
    return d;
  }

  _dayName(offset) {
    return Utilities.formatDate(this._dateAt(offset), this.config.TIMEZONE, 'EEEE');
  }

  _targets(page, week) {
    return [
      { webDay: page.todayDayName, sysDay: this._dayName(0),
        realDate: week.ymd(this._dateAt(0)), periods: page.todayPeriods, tag: 'today' },
      { webDay: page.nextDayName, sysDay: this._dayName(1),
        realDate: week.ymd(this._dateAt(1)), periods: page.nextPeriods, tag: 'tomorrow' }
    ];
  }

  /** All the reasons a target day is not safe to write to. */
  _eligible(session, t) {
    if (!t.webDay) { this.log.push('Skipped ' + t.tag + ': no day name on the page.'); return false; }
    if (t.webDay !== t.sysDay) {
      this.log.push('Skipped ' + t.tag + ': page says ' + t.webDay +
        ', system says ' + t.sysDay + '.');
      return false;
    }
    if (!session.grid.hasDay(t.webDay)) {
      this.log.push('No "' + t.webDay + '" column.');
      return false;
    }
    // The grid is dated, so a matching day NAME is not enough. On a Saturday
    // the page's "tomorrow" is Sunday — but this grid's Sunday column is six
    // days in the PAST, so writing there would backdate next week's classes.
    const columnDate = session.week.ymd(session.week.dateFor(t.webDay));
    if (columnDate !== t.realDate) {
      this.log.push('Skipped ' + t.webDay + ' (' + t.tag + '): grid column is ' +
        columnDate + ' but that day is ' + t.realDate + '. Roll over the week to catch up.');
      return false;
    }
    return true;
  }

  /** You own this slot. Compare, and queue a decision if the site differs. */
  _hold(session, day, start, incoming) {
    const slot = session.grid.slot(day, start);
    const override = session.ownership.override(day, start);
    const yours = slot ? slot.body : (override ? override.content : '');

    if (TimeKey.same(yours, incoming)) {
      session.conflicts.drop(day, start);
      return;
    }
    const dismissed = session.ownership.dismissed(day, start);
    if (dismissed && TimeKey.same(dismissed.content, incoming)) {
      this.log.push('Kept yours, already decided: ' + day + ' ' + start);
      return;
    }
    session.conflicts.record(day, start, incoming);
    this.log.push('Held your version, needs a decision: ' + day + ' ' + start +
      ' (Jadwal: ' + incoming.split('\n').join(' / ') + ')');
  }

  /** The scraper is taking this slot back: your event for it must go. */
  _dropMyEvent(session, day, start) {
    const rec = session.state.get('MANUAL', day, start);
    if (!rec) return;
    if (session.calendar.remove(rec.eventId)) {
      this.log.push('Deleted your event, Jadwal reclaimed the slot: ' + day + ' ' + start);
    }
    session.state.remove('MANUAL', day, start);
  }

  /** Start times on this day the scraper must not write over: ones you
   *  claimed by hand (even cleared ones), plus any slot already holding text
   *  that is not the scraper's. The second half matters when sync-on-edit is
   *  off, or when a claim was lost — the calendar sync would treat that text
   *  as yours, so the fetch must not silently delete it either. */
  _protectedStarts(session, day) {
    const out = Object.create(null);
    session.ownership.myStarts(day).forEach(start => { out[start] = true; });
    session.grid.slots(day).forEach(slot => {
      if (slot.body && session.ownership.isMine(day, slot)) out[slot.start] = true;
    });
    return out;
  }

  _fill(session, t) {
    const day = t.webDay;
    const keepLabel = this.config.KEEP_TIME_LABEL;
    const mine = this._protectedStarts(session, day);

    // Re-derive this day's clashes from scratch — the site is the fact, and
    // it may have changed since the last fetch.
    session.conflicts.dropDay(day);

    // Clear the slots this scraper owns: those the state remembers, PLUS any
    // whose content still looks like scraper output (self-heals lost state).
    // Slots you authored are excluded from both.
    const owned = session.state.all('JADWAL').filter(r => r.day === day).map(r => r.start);
    const stale = session.grid.jadwalLookingStarts(day);
    const toClear = owned.concat(stale.filter(x => owned.indexOf(x) === -1))
      .filter(x => !mine[x]);
    session.grid.clearSlots(day, toClear, keepLabel);
    session.state.removeAll('JADWAL', day);

    // Group by start time so two periods in one slot are written together,
    // in a single overwrite rather than successive appends.
    const bySlot = {};
    const unmatched = [];
    t.periods.forEach(p => {
      if (this.config.SKIP_CANCELLED && p.isCancelled) {
        this.log.push('Cancelled, left blank: ' + day + ' ' + p);
        return;
      }
      const key = p.startTime;
      if (!key) { unmatched.push(p.toString()); return; }
      (bySlot[key] = bySlot[key] || []).push(p);
    });

    let held = 0;
    Object.keys(bySlot).forEach(start => {
      const group = bySlot[start];
      if (mine[start]) {
        this._hold(session, day, start, group.map(p => p.render()).join('\n'));
        held++;
        return;
      }
      if (session.grid.placeAll(day, start, group, keepLabel)) {
        this._dropMyEvent(session, day, start);
        session.state.put('JADWAL', day, start,
          group.map(p => p.toString()).join(' + '), '', session.week.key);
      } else {
        group.forEach(p => unmatched.push(p.toString()));
      }
    });

    if (this.config.MARK_UNMATCHED_AS_NOTE) {
      session.grid.noteOnDay(day, unmatched.length
        ? 'No matching slot in the grid:\n' + unmatched.join('\n') : '');
    }
    unmatched.forEach(u => this.log.push('UNMATCHED ' + day + ': ' + u));
    this.log.push(day + ' (' + t.tag + '): ' + (t.periods.length - unmatched.length) +
      '/' + t.periods.length + ' placed' + (held ? ', ' + held + ' left to you' : '') + '.');
    return day;
  }

  /** Notes always; then a dialog if you started this run, an email if not. */
  _surface(session, touchedDays) {
    const conflicts = session.pendingConflicts();
    touchedDays.forEach(day => session.reporter.refreshNotes(day, conflicts));
    // Commit the notes before the modal opens: showModalDialog does not block,
    // and its Apply handler is a separate execution that clears the note it
    // finds. An uncommitted note would survive the decision that resolved it.
    SpreadsheetApp.flush();
    if (!conflicts.length) return false;

    if (this.interactive && this.config.ASK_BEFORE_OVERWRITING_MY_EDITS) {
      if (new ConflictDialog(conflicts, session.week).show()) {
        this.log.push('Asking you about ' + conflicts.length + ' held edit(s).');
        return true;
      }
    }
    const line = session.reporter.email(conflicts);
    if (line) this.log.push(line);
    this.log.push(conflicts.length + ' held edit(s) waiting in "Review conflicts".');
    _notify(conflicts.length + ' slot' + (conflicts.length === 1 ? '' : 's') +
      ' you edited now differ from Jadwal. Your version was kept — ' +
      'use 📅 Jadwal Grid → Review conflicts.', 'Jadwal differs', 12);
    return false;
  }

  run() {
    const page = JadwalPage.fetch(this.config.JADWAL_URL);
    const session = new SyncSession(this.config);
    const touched = [];

    this._targets(page, session.week).forEach(t => {
      if (this._eligible(session, t)) touched.push(this._fill(session, t));
    });

    // Everything is written and flushed BEFORE the dialog opens: a modal
    // dialog does not block this run, and its Apply handler is a separate
    // execution that re-reads the state sheet.
    session.flush();
    SpreadsheetApp.flush();

    const dialogShown = this._surface(session, touched);
    session.flush();

    this.log.forEach(l => console.log(l));
    return { log: this.log, dialogShown: dialogShown };
  }
}

/** =================================================================
 * Part B — the cells YOU own, into Google Calendar, on the header dates.
 * ================================================================= */
class ManualSync {
  constructor(config) { this.config = config; this.log = []; }

  /** Full pass: menu item and the daily catch-up. Claims nothing — it
   *  cannot tell who typed what, it only trusts what state already says. */
  runAll() {
    const session = new SyncSession(this.config);
    session.grid.days().forEach(day => {
      session.grid.slots(day).forEach(slot => {
        if (!slot.body) return;
        if (session.ownership.isJadwal(day, slot)) return;
        this._apply(session, day, slot);
      });
    });
    this._reconcile(session);
    session.flush();
    if (!this.log.length) this.log.push('Nothing to change — grid and calendar already agree.');
    this.log.forEach(l => console.log(l));
    return { log: this.log };
  }

  /** The onEdit path. Only human edits reach here, so every slot the edit
   *  touched becomes yours — even one the scraper had filled. */
  runEdit(sheet, row1, col1, numRows, numCols) {
    const session = new SyncSession(this.config);
    if (session.grid.sheet.getSheetId() !== sheet.getSheetId()) return { log: [] };

    const touched = session.grid.slotsIn(row1, col1, numRows, numCols);
    if (!touched.length) return { log: [] };

    // Emptying several slots at once is a wipe, not a decision about each one.
    // Claiming there would lock a whole week to you on one Delete keypress, so
    // a block clear resets ownership instead of taking it.
    const wipe = touched.length > 1 && touched.every(hit => !hit.slot.body);
    if (wipe) {
      this.log.push('Block clear of ' + touched.length +
        ' slots — ownership reset, Jadwal may refill them.');
    }

    touched.forEach(hit => {
      const day = hit.day, slot = hit.slot;
      if (wipe) {
        session.ownership.forget(day, slot.start);   // the cell is empty
        session.conflicts.drop(day, slot.start);
        session.reporter.clearNote(slot);
        return;
      }
      const wasJadwal = session.ownership.isJadwal(day, slot);
      if (session.ownership.claim(day, slot.start, slot.body) && wasJadwal) {
        this.log.push('Slot is yours now, Jadwal will leave it alone: ' +
          day + ' ' + TimeKey.label(slot));
      }
      // Typing the site's own text back in settles any pending clash.
      const conflict = session.conflicts.get(day, slot.start);
      if (conflict && TimeKey.same(conflict.content, slot.body)) {
        session.conflicts.drop(day, slot.start);
        session.reporter.clearNote(slot);
      }
    });

    // Ownership is persisted BEFORE any calendar call. A CalendarApp failure
    // must not lose the fact that this slot is now yours — otherwise the next
    // fetch would quietly overwrite it.
    session.flush();

    touched.forEach(hit => {
      if (wipe) return;
      const day = hit.day, slot = hit.slot;
      // Only reachable with SYNC_MY_JADWAL_EDITS off, where claiming is
      // disabled and the scraper's cells stay out of Calendar as in v3.
      if (session.ownership.isJadwal(day, slot)) {
        this.log.push('Jadwal-owned slot, ignored: ' + day + ' ' + slot.start);
        return;
      }
      if (slot.body) this._apply(session, day, slot);
    });

    this._reconcile(session);
    session.flush();
    this.log.forEach(l => console.log(l));
    return { log: this.log };
  }

  /** Events whose slot is no longer yours-with-content are deleted. */
  _reconcile(session) {
    const live = {};
    session.grid.days().forEach(day => {
      session.grid.slots(day).forEach(slot => {
        if (slot.body && session.ownership.isMine(day, slot)) {
          live[StateStore.key('MANUAL', day, slot.start)] = true;
        }
      });
    });
    session.state.all('MANUAL').forEach(rec => {
      if (live[StateStore.key('MANUAL', rec.day, rec.start)]) return;
      if (session.calendar.remove(rec.eventId)) {
        this.log.push('Deleted event: ' + rec.day + ' ' + rec.start);
      }
      session.state.remove('MANUAL', rec.day, rec.start);
    });
  }

  _apply(session, day, slot) {
    const week = session.week;
    const date = week.dateFor(day);
    const rec = session.state.get('MANUAL', day, slot.start);
    const sameWeek = rec && rec.weekStart === week.key;
    const event = sameWeek ? session.calendar.fetch(rec.eventId) : null;

    // Already correct only if the text AND the date still match — re-dating
    // the grid moves the date while leaving the text untouched, and an event
    // deleted in Calendar has to be recreated.
    if (event && rec.content === slot.body &&
        event.getStartTime().getTime() === week.at(date, slot.start).getTime()) return;
    if (rec && !sameWeek) session.calendar.remove(rec.eventId);

    const id = session.calendar.upsert(day, slot, slot.body, date,
      sameWeek ? rec.eventId : null);
    session.state.put('MANUAL', day, slot.start, slot.body, id, week.key);
    // Commit the id now. If a later slot in this pass throws, the events
    // already created stay reachable instead of being orphaned and remade.
    session.flush();
    this.log.push((event ? 'Updated' : 'Created') + ': ' + day + ' ' +
      week.ymd(date) + ' ' + TimeKey.label(slot) + ' — ' + slot.body.split('\n')[0]);
  }
}

/** =================================================================
 * Applying the review decisions. Runs as its own execution when the
 * dialog's Apply button posts back, and from "Review conflicts".
 * ================================================================= */
class ConflictResolver {
  constructor(config) { this.config = config; this.log = []; }

  apply(decisions) {
    const session = new SyncSession(this.config);
    let mine = 0, theirs = 0;

    (decisions || []).forEach(d => {
      const conflict = session.conflicts.get(d.day, d.start);
      if (!conflict) return;
      const slot = session.grid.slot(d.day, d.start);
      if (!slot) { session.conflicts.drop(d.day, d.start); return; }

      if (d.choice === 'jadwal') { this._takeJadwal(session, d.day, slot, conflict.content); theirs++; }
      else { this._keepMine(session, d.day, slot, conflict.content); mine++; }

      session.conflicts.drop(d.day, d.start);
      session.reporter.clearNote(slot);
    });

    session.flush();
    SpreadsheetApp.flush();
    this.log.unshift('Kept yours: ' + mine + ' · restored from Jadwal: ' + theirs + '.');
    this.log.forEach(l => console.log(l));
    _notify(this.log.join('\n'), 'Conflicts resolved', 8);
    return { log: this.log };
  }

  /** Put the site's text back and hand the slot to the scraper. */
  _takeJadwal(session, day, slot, incoming) {
    session.grid.setSlot(slot, incoming, this.config.KEEP_TIME_LABEL);
    session.ownership.takeOver(day, slot.start, incoming);
    const rec = session.state.get('MANUAL', day, slot.start);
    if (rec) {
      session.calendar.remove(rec.eventId);
      session.state.remove('MANUAL', day, slot.start);
    }
    this.log.push('Restored Jadwal: ' + session.week.stamp(day) + ' ' + TimeKey.label(slot));
  }

  /** Remember the rejected version so this clash never asks twice. */
  _keepMine(session, day, slot, incoming) {
    session.ownership.dismiss(day, slot.start, incoming);
    this.log.push('Kept yours: ' + session.week.stamp(day) + ' ' + TimeKey.label(slot));
  }
}

/** =================================================================
 * Weekly rollover: archive -> wipe -> advance the dates.
 * ================================================================= */
class WeekRollover {
  constructor(config) { this.config = config; this.log = []; }

  /** Collect every filled slot as a flat archive row. */
  _harvest(session) {
    const rows = [];
    session.grid.days().forEach(day => {
      const date = session.week.ymd(session.week.dateFor(day));
      session.grid.slots(day).forEach(slot => {
        if (!slot.body) return;
        const source = session.ownership.creditFor(day, slot);
        if (source === SlotOwnership.JADWAL && !this.config.ARCHIVE_JADWAL_ROWS) return;
        rows.push([date, day, slot.start, slot.end, source, TimeKey.dedupe(slot.body)]);
      });
    });
    return rows.sort((a, b) => (a[0] + a[2]).localeCompare(b[0] + b[2]));
  }

  run() {
    const session = new SyncSession(this.config);
    const archive = new Archive(session.ss, this.config.ARCHIVE_SHEET_NAME);

    const rows = this._harvest(session);
    const archived = archive.append(rows);
    this.log.push('Archived ' + archived + ' entr' + (archived === 1 ? 'y' : 'ies') +
      ' for week of ' + session.week.key + '.');

    const cleared = session.grid.clearAll(this.config.KEEP_TIME_LABEL);
    this.log.push('Cleared ' + cleared + ' cell' + (cleared === 1 ? '' : 's') + '.');

    // Calendar events for the archived week stay put — that week happened.
    session.state.removeKind('MANUAL');
    session.state.removeKind('JADWAL');
    session.ownership.forgetAll();           // overrides, dismissals, conflicts
    session.grid.days().forEach(d => session.grid.noteOnDay(d, ''));
    session.reporter.clearAllNotes();
    this.log.push('Ownership reset: the new week starts with every slot Jadwal\'s.');

    const previous = session.week.key;
    session.week.advance();
    session.grid.applyDates(session.week);
    session.flush();
    SpreadsheetApp.flush();
    this.log.push('Grid re-dated: ' + previous + ' -> ' + session.week.key +
      ' (' + session.week.rangeLabel + ').');
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
  return _withLock(() => {
    const ss = SpreadsheetApp.getActive();
    const state = new StateStore(ss, CONFIG);
    const week = new WeekContext(CONFIG, state);
    const nag = new StaleWeekNag(CONFIG, state, week);
    const lines = nag.run();
    state.flush();
    return { log: lines || ['Grid week is current (' + week.key + ').'], stale: nag.isStale };
  });
}

/** =================================================================
 * Entry points
 * ================================================================= */

/**
 * StateStore.flush() rewrites the whole hidden sheet from an in-memory
 * snapshot, so two overlapping executions would clobber each other's rows —
 * and there are several: the onEdit trigger, the 6 AM trigger, a menu run,
 * and the review dialog's Apply handler, which is its own execution and can
 * fire while a menu run is still going. Every mutating entry point therefore
 * takes the document lock, which makes each run's read-modify-flush atomic.
 * Re-entrant by design: a locked function may call another one.
 */
let _lockDepth = 0;

function _withLock(fn, timeoutMs) {
  if (_lockDepth > 0) return fn();
  let lock = null;
  try { lock = LockService.getDocumentLock(); } catch (e) { return fn(); }
  if (!lock.tryLock(timeoutMs == null ? 25000 : timeoutMs)) {
    throw new Error('Another schedule sync is still running. Try again in a moment.');
  }
  _lockDepth++;
  try { return fn(); } finally { _lockDepth--; lock.releaseLock(); }
}
function setThisWeek() {
  return _withLock(() => {
    const session = new SyncSession(CONFIG);
    const previous = session.week.key;
    const target = session.week.ymd(session.week.currentSunday());
    const log = [];

    // Every ownership row and event id is stamped with the week it belongs to.
    // Re-dating keeps the grid's contents, so the bookkeeping MOVES with them:
    // your slots stay yours and their calendar events shift to the new dates on
    // the next sync. Leaving the rows behind would strand them — invisible to
    // every lookup, and their events unreachable and duplicated.
    // (To archive the old week and start clean, use "Roll over week".)
    if (target !== previous) {
      const moved = session.ownership.myCount();
      const pending = session.conflicts.count;
      session.state.restamp('OVERRIDE', target);
      session.state.restamp('MANUAL', target);
      session.state.restamp('JADWAL', target);
      // A "Keep mine" pairs with the text and the override that move with the
      // grid, so it moves too — otherwise the same clash asks a second time.
      session.state.restamp('DISMISSED', target);
      // Conflicts are the exception: they are re-derived by the next fetch.
      session.state.removeKind('CONFLICT');
      session.reporter.clearAllNotes();
      log.push('Moved the grid\'s contents from ' + previous + ' to ' + target +
        ': ' + moved + ' edited slot(s) still yours, ' + pending +
        ' pending conflict(s) dropped. Calendar events follow the new dates on ' +
        'the next sync — use "Roll over week" instead to archive and start clean.');
    }

    session.week.setWeekStart(session.week.currentSunday());
    session.grid.applyDates(session.week);
    session.flush();
    log.push('Grid pinned to week of ' + session.week.key +
      ' (' + session.week.rangeLabel + ').');
    return { log: log };
  });
}

/** De-duplicate repeated blocks in every grid cell, and re-register any
 *  scraper-looking cell as Jadwal-owned. Fixes grids damaged by the old
 *  append-instead-of-replace behaviour. Cells you authored are left alone. */
function repairGrid() { return _withLock(_repairGrid); }

function _repairGrid() {
  const session = new SyncSession(CONFIG);
  const log = [];
  let fixed = 0, reclaimed = 0, skipped = 0;

  session.grid.days().forEach(day => {
    session.grid.slots(day).forEach(slot => {
      if (!slot.body) return;
      const clean = TimeKey.dedupe(slot.body);
      if (clean !== slot.body) {
        const before = TimeKey.blocks(slot.body).length;
        session.grid.setSlot(slot, clean, CONFIG.KEEP_TIME_LABEL);
        fixed++;
        log.push(day + ' ' + TimeKey.label(slot) + ': ' + before + ' blocks -> ' +
          TimeKey.blocks(clean).length);
      }
      if (session.ownership.isMineByHand(day, slot.start)) { skipped++; return; }
      if (TimeKey.looksJadwal(clean) && !session.state.get('JADWAL', day, slot.start)) {
        session.state.put('JADWAL', day, slot.start, clean.split('\n')[0], '', session.week.key);
        // The slot becomes the scraper's, so any event of yours for it must
        // go with it — dropping the row alone would orphan the event beyond
        // the reach of every later pass.
        const rec = session.state.get('MANUAL', day, slot.start);
        if (rec) {
          if (session.calendar.remove(rec.eventId)) {
            log.push('Deleted orphan event: ' + day + ' ' + TimeKey.label(slot));
          }
          session.state.remove('MANUAL', day, slot.start);
        }
        reclaimed++;
      }
    });
  });

  session.flush();
  SpreadsheetApp.flush();
  log.unshift('De-duplicated ' + fixed + ' cell' + (fixed === 1 ? '' : 's') +
    '; re-tagged ' + reclaimed + ' as Jadwal-owned; left ' + skipped + ' of your edits alone.');
  log.forEach(l => console.log(l));
  return { log: log };
}

/** Give every slot you edited back to the scraper. The next fetch fills
 *  them, and the calendar events they created are removed. */
function releaseMyEdits() { return _withLock(_releaseMyEdits); }

function _releaseMyEdits() {
  const session = new SyncSession(CONFIG);
  const n = session.ownership.handBackAll(session.grid);
  session.reporter.clearAllNotes();
  session.flush();
  SpreadsheetApp.flush();
  // The handover makes those slots the scraper's, so this pass removes the
  // calendar events they had while they were yours.
  const after = new ManualSync(CONFIG).runAll();
  return { log: ['Handed ' + n + ' edited slot' + (n === 1 ? '' : 's') +
    ' back to Jadwal. Run "Fill grid from Jadwal" to refill them.'].concat(after.log) };
}

function reviewConflicts() {
  const session = new SyncSession(CONFIG);
  const conflicts = session.pendingConflicts();
  if (!conflicts.length) {
    return { log: ['No pending conflicts — Jadwal and your edits agree.'] };
  }
  const shown = new ConflictDialog(conflicts, session.week).show();
  return {
    log: [shown
      ? 'Reviewing ' + conflicts.length + ' conflict(s).'
      : 'Could not open the dialog. Pending:\n' + session.reporter.describe(conflicts)],
    dialogShown: shown
  };
}

/** Called by the review dialog's Apply button — its own execution, which is
 *  why it must queue behind whatever run opened the dialog. */
function applyJadwalConflictDecisions(decisions) {
  return _withLock(() => new ConflictResolver(CONFIG).apply(decisions), 60000);
}

function syncJadwalToGrid(options) {
  return _withLock(() => new JadwalSync(CONFIG, options).run());
}
function syncGridToCalendar() { return _withLock(() => new ManualSync(CONFIG).runAll()); }
function rollOverWeek() { return _withLock(() => new WeekRollover(CONFIG).run()); }

function syncEverything(options) {
  const nag = checkStaleWeek();
  const a = syncJadwalToGrid(options);
  const b = syncGridToCalendar();
  return {
    log: (nag.stale ? nag.log : []).concat(a.log, b.log),
    dialogShown: !!a.dialogShown
  };
}

function _notify(message, title, seconds) {
  try { SpreadsheetApp.getActive().toast(message, title || 'Calendar', seconds || 5); } catch (e) {}
}

function _toast(message, title) {
  if (!CONFIG.TOAST_ON_EDIT) return;
  _notify(message, title);
}

/**
 * Installable onEdit trigger — instant per-cell calendar sync, and the one
 * place authorship is known. Google fires this only for edits a person
 * makes in the Sheets UI; writes made by this script never fire it, so
 * there is no feedback loop and no way for the scraper to masquerade as you.
 */
function onGridEdit(e) {
  if (!e || !e.range) return;
  try {
    const r = e.range;
    const res = _withLock(() => new ManualSync(CONFIG)
      .runEdit(r.getSheet(), r.getRow(), r.getColumn(), r.getNumRows(), r.getNumColumns()),
      20000);
    if (res.log.length) _toast(res.log.slice(0, 6).join('\n'));
  } catch (err) {
    console.error('onGridEdit failed: ' + err + (err.stack ? '\n' + err.stack : ''));
    _toast('Sync failed: ' + err, 'Error');
  }
}

function diagnoseGrid() {
  const out = [];
  try {
    const session = new SyncSession(CONFIG);
    const week = session.week;
    out.push('Grid sheet: "' + session.grid.sheet.getName() + '", header on row ' +
      (session.grid.headerRow + 1));
    out.push('Grid week: ' + week.key + '  (' + week.rangeLabel + ')');
    out.push('Real current week: ' + week.ymd(week.currentSunday()) +
      (week.ymd(week.currentSunday()) === week.key ? '  — in sync' : '  — ROLL OVER NEEDED'));
    out.push('Time axis rows in column A: ' + session.grid.axis.filter(Boolean).length);
    out.push('Edits of yours this week: ' + session.ownership.myCount() +
      (CONFIG.SYNC_MY_JADWAL_EDITS ? '' : '   (SYNC_MY_JADWAL_EDITS is OFF)'));
    const conflicts = session.conflicts.all();
    out.push('Pending conflicts: ' + conflicts.length +
      (conflicts.length ? '  -> ' + conflicts.map(c => c.day + ' ' + c.start).join(', ') : ''));
    out.push('');
    session.grid.days().forEach(day => {
      const filled = session.grid.slots(day).filter(s => s.body);
      out.push(day + ' ' + week.pretty(week.dateFor(day)) + ': ' +
        session.grid.slots(day).length + ' slots, ' + filled.length + ' filled' +
        (filled.length ? ' -> ' + filled.map(s => TimeKey.label(s) +
          (session.ownership.isMineByHand(day, s.start) ? ' [yours]' : '') +
          ' "' + s.body.split('\n')[0].slice(0, 22) + '"').join('; ') : ''));
    });
    out.push('');
    out.push('Default calendar: ' + CalendarApp.getDefaultCalendar().getName());
    const arch = session.ss.getSheetByName(CONFIG.ARCHIVE_SHEET_NAME);
    out.push('Archive rows: ' + (arch ? Math.max(0, arch.getLastRow() - 1) : 'sheet not created yet'));
    const triggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
    out.push('Sync on edit: ' + (triggers.indexOf('onGridEdit') !== -1
      ? 'ON' : 'OFF — run "Turn on sync-on-edit"'));
    out.push('Triggers: ' + (triggers.length ? triggers.join(', ') : 'none'));
  } catch (e) {
    out.push('ERROR: ' + e.message);
  }
  const text = out.join('\n');
  console.log(text);
  try { SpreadsheetApp.getUi().alert('Diagnostics', text, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
  return text;
}

/** Runs fn and reports it — unless it opened a modal dialog of its own, in
 *  which case a second modal would fight with it and a toast is used. */
function _alert(title, fn) {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = fn() || {};
    const text = (res.log || []).join('\n') || 'Nothing to do.';
    if (res.dialogShown) _notify(text, title, 8);
    else ui.alert(title, text, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert(title + ' failed', String(err), ui.ButtonSet.OK);
  }
}

function menuSetWeek() {
  const ui = SpreadsheetApp.getUi();
  try {
    const session = new SyncSession(CONFIG);
    const target = session.week.ymd(session.week.currentSunday());
    const n = session.ownership.myCount(), pending = session.conflicts.count;
    if (target !== session.week.key && (n || pending) &&
      ui.alert('Set this week',
        'Re-date the grid from ' + session.week.key + ' to ' + target +
        ', keeping everything in it?\n\nThe entries move with it: your ' + n +
        ' edited slot(s) stay yours and their calendar events shift to the new ' +
        'dates. ' + pending + ' pending conflict(s) are dropped.\n\n' +
        'Nothing is archived — to file this week away and start clean, use ' +
        '"Roll over week" instead.\n\nContinue?',
        ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  } catch (e) { /* fall through — setThisWeek reports its own errors */ }
  _alert('Set this week', setThisWeek);
}
function menuSyncJadwal() { _alert('Jadwal → Grid', () => syncJadwalToGrid({ interactive: true })); }
function menuSyncCalendar() { _alert('Grid → Calendar', syncGridToCalendar); }
function menuSyncBoth() { _alert('Full sync', () => syncEverything({ interactive: true })); }
function menuReviewConflicts() { _alert('Review conflicts', reviewConflicts); }
function menuRepair() { _alert('Repair grid', repairGrid); }

function menuReleaseMyEdits() {
  const ui = SpreadsheetApp.getUi();
  let n = 0;
  try {
    n = new SyncSession(CONFIG).ownership.myCount();
  } catch (e) {
    ui.alert('Hand my edits back failed', String(e), ui.ButtonSet.OK);
    return;
  }
  if (!n) { ui.alert('Hand my edits back', 'You have no edited slots this week.', ui.ButtonSet.OK); return; }
  if (ui.alert('Hand my edits back',
    'Give ' + n + ' slot(s) you edited back to Jadwal, and forget every pending ' +
    'conflict?\n\nThe calendar events your edits created are deleted, and the next ' +
    'fetch is free to overwrite those cells. Whatever you wrote stays on screen ' +
    'until it does, and is still archived as yours.\n\nContinue?',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  _alert('Hand my edits back', releaseMyEdits);
}

function menuRollOver() {
  const ui = SpreadsheetApp.getUi();
  let summary = '';
  try {
    const session = new SyncSession(CONFIG);
    const filled = session.grid.days()
      .reduce((n, d) => n + session.grid.slots(d).filter(s => s.body).length, 0);
    summary = 'Archive ' + filled + ' entries from the week of ' + session.week.key +
      ', then wipe the grid and re-date it.\n\nCalendar events already created are kept. ' +
      'Your ' + session.ownership.myCount() + ' edited slot(s) and any pending conflicts ' +
      'are reset — the new week starts with every slot Jadwal\'s.\n\nContinue?';
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
    'updated, or deleted immediately. That includes a cell Jadwal filled: ' +
    'editing it makes the slot yours.\n\n' +
    'Daily 6 AM: Jadwal scrape + a catch-up calendar sync. If the site then ' +
    'disagrees with something you edited, your version is kept and the clash ' +
    'waits in "Review conflicts".\n\n' +
    'Rollover stays manual — use "Roll over week".',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function onOpen() {
  let pending = 0;
  try {
    // Deliberately not a SyncSession: onOpen must not scan every sheet
    // looking for the grid, nor touch CalendarApp.
    const state = new StateStore(SpreadsheetApp.getActive(), CONFIG);
    const week = new WeekContext(CONFIG, state);
    pending = new ConflictQueue(state, week).count;
    const nag = new StaleWeekNag(CONFIG, state, week);
    if (nag.isStale) _notify(nag.message, 'Roll over week', 15);
    else if (pending) {
      _notify(pending + ' slot' + (pending === 1 ? '' : 's') + ' you edited now differ ' +
        'from Jadwal. Your version is being kept — Review conflicts to decide.',
        'Jadwal differs', 12);
    }
  } catch (e) { /* never block the menu */ }

  SpreadsheetApp.getUi().createMenu('📅 Jadwal Grid')
    .addItem('Turn on sync-on-edit', 'installTriggers')
    .addItem('Set this week', 'menuSetWeek')
    .addSeparator()
    .addItem('Catch-up sync to Calendar', 'menuSyncCalendar')
    .addItem('Fill grid from Jadwal', 'menuSyncJadwal')
    .addItem('Do both', 'menuSyncBoth')
    .addSeparator()
    .addItem('Review conflicts' + (pending ? ' (' + pending + ')' : ''), 'menuReviewConflicts')
    .addItem('Hand my edits back to Jadwal', 'menuReleaseMyEdits')
    .addSeparator()
    .addItem('Roll over week', 'menuRollOver')
    .addSeparator()
    .addItem('Repair grid', 'menuRepair')
    .addItem('Diagnose', 'diagnoseGrid')
    .addToUi();
}
