/**
 * ============================================================================
 * PeriodCount.gs  —  Hijri period counter for Google Sheets
 * ============================================================================
 *
 *   =PeriodCount("Mon,Tue,Wed", "[1/1-15/1,20/4]", "1/2/1447", "29/12/1447")
 *
 * Counts how many times the given weekdays occur between Start and End
 * (both inclusive), skipping any day that falls inside an exclusion.
 *
 * Built on the HijriDate implementation from hijri_date.js (Fatimid / Bohra
 * tabular calendar: Kabisa remainders 2,5,8,10,13,16,19,21,24,27,29).
 * Day-of-week uses the same convention as hijri_calendar.js:
 * (AJD + 1.5) % 7, where 0 = Sunday.
 *
 * INSTALL: Extensions > Apps Script > paste this file > Save.
 * ============================================================================
 */


/* ==========================================================================
 * 1. SETTINGS
 * ========================================================================== */

/**
 * Shift the whole Hijri calendar by a whole number of days.
 * 0 = exactly as hijri_date.js computes it.
 * Use -1 or +1 if your local calendar starts months a day earlier/later.
 */
var HIJRI_DAY_ADJUSTMENT = 0;

/**
 * How to read Start/End when they are typed as TEXT ("1/2/1447").
 * "hijri" or "gregorian". Real date cells are always read as Gregorian.
 */
var DEFAULT_DATE_SYSTEM = "hijri";

/** Safety valve: refuse windows longer than this many days. */
var MAX_WINDOW_DAYS = 40000;


/* ==========================================================================
 * 2. HijriDate  — port of hijri_date.js (dependency-free, no Lazy.js)
 * ========================================================================== */

var HijriDate = (function () {
  "use strict";

  var KABISA_YEAR_REMAINDERS = [2, 5, 8, 10, 13, 16, 19, 21, 24, 27, 29];

  var DAYS_IN_YEAR = [30, 59, 89, 118, 148, 177, 207, 236, 266, 295, 325];

  var DAYS_IN_30_YEARS = [
     354,  708, 1063, 1417, 1771, 2126, 2480, 2834,  3189,  3543,
    3898, 4252, 4606, 4961, 5315, 5669, 6024, 6378,  6732,  7087,
    7441, 7796, 8150, 8504, 8859, 9213, 9567, 9922, 10276, 10631
  ];

  var MONTH_NAMES = {
    long: [
      "Moharram al-Haraam", "Safar al-Muzaffar", "Rabi al-Awwal",
      "Rabi al-Aakhar", "Jumada al-Ula", "Jumada al-Ukhra",
      "Rajab al-Asab", "Shabaan al-Karim", "Ramadaan al-Moazzam",
      "Shawwal al-Mukarram", "Zilqadah al-Haraam", "Zilhaj al-Haraam"
    ],
    short: [
      "Moharram", "Safar", "Rabi I", "Rabi II", "Jumada I", "Jumada II",
      "Rajab", "Shabaan", "Ramadaan", "Shawwal", "Zilqadah", "Zilhaj"
    ]
  };

  var hijriDate = function (year, month, day) {
    this.year = year;
    this.month = month;
    this.day = day;
  };

  hijriDate.prototype.getYear  = function () { return this.year;  };
  hijriDate.prototype.getMonth = function () { return this.month; };
  hijriDate.prototype.getDate  = function () { return this.day;   };

  hijriDate.getMonthName      = function (m) { return MONTH_NAMES.long[m];  };
  hijriDate.getShortMonthName = function (m) { return MONTH_NAMES.short[m]; };

  hijriDate.isJulian = function (date) {
    if (date.getFullYear() < 1582) return true;
    if (date.getFullYear() === 1582) {
      if (date.getMonth() < 9) return true;
      if (date.getMonth() === 9 && date.getDate() < 5) return true;
    }
    return false;
  };

  hijriDate.gregorianToAJD = function (date) {
    var a, b,
        year  = date.getFullYear(),
        month = date.getMonth() + 1,
        day   = date.getDate()
              + date.getHours()   / 24
              + date.getMinutes() / 1440
              + date.getSeconds() / 86400
              + date.getMilliseconds() / 86400000;
    if (month < 3) { year--; month += 12; }
    if (hijriDate.isJulian(date)) {
      b = 0;
    } else {
      a = Math.floor(year / 100);
      b = 2 - a + Math.floor(a / 4);
    }
    return Math.floor(365.25 * (year + 4716)) +
           Math.floor(30.6001 * (month + 1)) + day + b - 1524.5;
  };

  hijriDate.ajdToGregorian = function (ajd) {
    var a, b, c, d, e, f, z, alpha, year, month, day, hrs, min, sec, msc;
    z = Math.floor(ajd + 0.5);
    f = (ajd + 0.5 - z);
    if (z < 2299161) {
      a = z;
    } else {
      alpha = Math.floor((z - 1867216.25) / 36524.25);
      a = z + 1 + alpha - Math.floor(0.25 * alpha);
    }
    b = a + 1524;
    c = Math.floor((b - 122.1) / 365.25);
    d = Math.floor(365.25 * c);
    e = Math.floor((b - d) / 30.6001);

    day   = b - d - Math.floor(30.6001 * e) + f;
    hrs   = (day - Math.floor(day)) * 24;
    min   = (hrs - Math.floor(hrs)) * 60;
    sec   = (min - Math.floor(min)) * 60;
    msc   = (sec - Math.floor(sec)) * 1000;
    month = (e < 14) ? (e - 2) : (e - 14);
    year  = (month < 2) ? (c - 4715) : (c - 4716);
    return new Date(year, month, day, hrs, min, sec, msc);
  };

  hijriDate.isKabisa = function (year) {
    for (var i = 0; i < KABISA_YEAR_REMAINDERS.length; i++) {
      if (year % 30 === KABISA_YEAR_REMAINDERS[i]) return true;
    }
    return false;
  };

  hijriDate.daysInMonth = function (year, month) {
    return ((month === 11 && hijriDate.isKabisa(year)) || (month % 2 === 0)) ? 30 : 29;
  };

  hijriDate.prototype.dayOfYear = function () {
    return (this.month === 0) ? this.day : (DAYS_IN_YEAR[this.month - 1] + this.day);
  };

  hijriDate.fromAJD = function (ajd) {
    var year, month, date, i = 0,
        left = Math.floor(ajd - 1948083.5),
        y30  = Math.floor(left / 10631.0);

    left -= y30 * 10631;
    while (left > DAYS_IN_30_YEARS[i]) i += 1;

    year = Math.round(y30 * 30.0 + i);
    if (i > 0) left -= DAYS_IN_30_YEARS[i - 1];

    i = 0;
    while (left > DAYS_IN_YEAR[i]) i += 1;
    month = Math.round(i);
    date  = (i > 0) ? Math.round(left - DAYS_IN_YEAR[i - 1]) : Math.round(left);

    // --- BOUNDARY FIX (not in the original hijri_date.js) -------------------
    // The original returns date 0 for the last day of a year whose remainder
    // mod 30 is 29 (e.g. 30 Zilhaj 1409, 1439, 1469). Roll back to the real
    // last day of the previous month.
    if (date < 1) {
      month -= 1;
      if (month < 0) { month = 11; year -= 1; }
      date = hijriDate.daysInMonth(year, month);
    }
    // -----------------------------------------------------------------------

    return new hijriDate(year, month, date);
  };

  hijriDate.prototype.toAJD = function () {
    var y30 = Math.floor(this.year / 30.0),
        ajd = 1948083.5 + y30 * 10631 + this.dayOfYear();
    if (this.year % 30 !== 0) ajd += DAYS_IN_30_YEARS[this.year - y30 * 30 - 1];
    return ajd;
  };

  hijriDate.fromGregorian = function (date) {
    return hijriDate.fromAJD(hijriDate.gregorianToAJD(date));
  };

  hijriDate.prototype.toGregorian = function () {
    return hijriDate.ajdToGregorian(this.toAJD());
  };

  return hijriDate;
})();


/* ==========================================================================
 * 3. NAME TABLES
 * ========================================================================== */

/** 0 = Sunday, matching (AJD + 1.5) % 7 in hijri_calendar.js */
var DAY_ALIASES_ = {
  sun: 0, sunday: 0, ahad: 0, alahad: 0, yaumalahad: 0,
  mon: 1, monday: 1, ithnain: 1, isnain: 1, alithnain: 1,
  tue: 2, tues: 2, tuesday: 2, thulatha: 2, salasa: 2, sulasa: 2,
  wed: 3, weds: 3, wednesday: 3, arbaa: 3, arbia: 3, alarbaa: 3,
  thu: 4, thur: 4, thurs: 4, thursday: 4, khamis: 4, alkhamis: 4,
  fri: 5, friday: 5, jumua: 5, jumuah: 5, juma: 5, jummah: 5,
  sat: 6, saturday: 6, sabt: 6, alsabt: 6
};

/** Month aliases, already normalised (lowercase, letters/digits only). */
var MONTH_ALIASES_ = [
  ["moharram", "moharam", "muharram", "muharam", "moharramalharaam", "moh", "muh"],
  ["safar", "saphar", "safaralmuzaffar", "saf"],
  ["rabii", "rabi1", "rabialawwal", "rabiulawwal", "rabialawal", "rabiulawal", "rabiawwal"],
  ["rabiii", "rabi2", "rabialaakhar", "rabiulaakhar", "rabiulakhar", "rabialakhir",
   "rabialthani", "rabiulthani", "rabiussani"],
  ["jumadai", "jumada1", "jumadaalula", "jumadaulula", "jumadaula",
   "jumadaalawwal", "jamadiulawwal", "jamadaulula"],
  ["jumadaii", "jumada2", "jumadaalukhra", "jumadaulukhra", "jumadaukhra",
   "jumadaalakhirah", "jamadiulakhar", "jumadaalthani"],
  ["rajab", "rajabalasab", "raj"],
  ["shabaan", "shaban", "shabaanalkarim", "shabaanalkareem", "shaa"],
  ["ramadaan", "ramadan", "ramzan", "ramadhan", "ramadaanalmoazzam", "ram"],
  ["shawwal", "shawal", "shawwalalmukarram", "shaw"],
  ["zilqadah", "zilqada", "zilqad", "zilqaad", "zulqadah", "dhulqadah", "dhualqadah"],
  ["zilhaj", "zilhajj", "zilhajjah", "zilhijjah", "zulhaj", "dhulhijjah", "dhulhijja"]
];


/* ==========================================================================
 * 4. CUSTOM FUNCTIONS
 * ========================================================================== */

/**
 * Counts how many periods fall between two Hijri dates.
 *
 * @param {"Mon,Tue,Wed"} days Weekdays the subject has periods on.
 * @param {"[1/1-15/1,20/4]"} exclusions Hijri dates/ranges to skip (blank for none).
 * @param {"1/2/1447"} start First day of the window, inclusive.
 * @param {"29/12/1447"} end Last day of the window, inclusive.
 * @param {"hijri"} dateSystem Optional. How to read TEXT start/end: "hijri" (default) or "gregorian".
 * @return {number} The number of periods.
 * @customfunction
 */
function PeriodCount(days, exclusions, start, end, dateSystem) {
  var result = periodScan_(days, exclusions, start, end, dateSystem);
  return result.count;
}

/**
 * Same as PeriodCount but lists every counted day, for checking your setup.
 *
 * @param {"Mon,Tue,Wed"} days Weekdays the subject has periods on.
 * @param {"[1/1-15/1,20/4]"} exclusions Hijri dates/ranges to skip (blank for none).
 * @param {"1/2/1447"} start First day of the window, inclusive.
 * @param {"29/12/1447"} end Last day of the window, inclusive.
 * @param {"hijri"} dateSystem Optional. How to read TEXT start/end: "hijri" (default) or "gregorian".
 * @return {Array} Hijri date, weekday and Gregorian date for each period counted.
 * @customfunction
 */
function PeriodDates(days, exclusions, start, end, dateSystem) {
  var result = periodScan_(days, exclusions, start, end, dateSystem);
  if (!result.dates.length) return [["No periods found"]];

  var rows = [["Hijri", "Month", "Weekday", "Gregorian"]];
  for (var i = 0; i < result.dates.length; i++) {
    var ajd = result.dates[i],
        h   = ajdToHijri_(ajd),
        g   = HijriDate.ajdToGregorian(ajd);
    rows.push([
      h.getDate() + "/" + (h.getMonth() + 1) + "/" + h.getYear(),
      HijriDate.getShortMonthName(h.getMonth()),
      ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"][weekdayOfAJD_(ajd)],
      g
    ]);
  }
  return rows;
}

/**
 * Converts a Hijri date to a Gregorian date.
 *
 * @param {"1/1/1447"} hijri A Hijri date as d/m/yyyy, or "1 Moharram 1447".
 * @return {Date} The matching Gregorian date.
 * @customfunction
 */
function HijriToGregorian(hijri) {
  var p = parseDatePart_(String(firstValue_(hijri)), null);
  if (p.year === null) throw new Error("HijriToGregorian: include a year, e.g. 1/1/1447");
  return HijriDate.ajdToGregorian(hijriToAJD_(p.year, p.month, p.day));
}

/**
 * Converts a Gregorian date to a Hijri date string.
 *
 * @param {DATE(2025,6,26)} gregorian A Gregorian date cell.
 * @return {string} The matching Hijri date, e.g. "1 Moharram 1447".
 * @customfunction
 */
function GregorianToHijri(gregorian) {
  var v = firstValue_(gregorian);
  if (!(v instanceof Date)) throw new Error("GregorianToHijri: needs a real date cell.");
  var h = ajdToHijri_(HijriDate.gregorianToAJD(v));
  return h.getDate() + " " + HijriDate.getShortMonthName(h.getMonth()) + " " + h.getYear();
}


/* ==========================================================================
 * 5. CORE ENGINE
 * ========================================================================== */

function periodScan_(days, exclusions, start, end, dateSystem) {
  var system = String(firstValue_(dateSystem) || DEFAULT_DATE_SYSTEM)
                 .toLowerCase().indexOf("greg") === 0 ? "gregorian" : "hijri";

  var wanted = parseDays_(days);
  if (!wanted.length) throw new Error('PeriodCount: no weekdays given, e.g. "Mon,Tue,Wed".');

  var startAJD = resolveEndpoint_(start, system, "Start"),
      endAJD   = resolveEndpoint_(end,   system, "End");

  if (startAJD > endAJD) { var t = startAJD; startAJD = endAJD; endAJD = t; }

  var span = Math.round(endAJD - startAJD) + 1;
  if (span > MAX_WINDOW_DAYS) {
    throw new Error("PeriodCount: window is " + span + " days; max is " + MAX_WINDOW_DAYS + ".");
  }

  var blocks = buildExclusions_(exclusions, startAJD, endAJD);

  var wantedSet = {};
  for (var w = 0; w < wanted.length; w++) wantedSet[wanted[w]] = true;

  var hits = [];
  for (var ajd = startAJD; ajd <= endAJD; ajd += 1) {
    if (!wantedSet[weekdayOfAJD_(ajd)]) continue;
    if (isBlocked_(ajd, blocks)) continue;
    hits.push(ajd);
  }

  return { count: hits.length, dates: hits };
}

/** Weekday of an AJD. 0 = Sunday, same convention as hijri_calendar.js. */
function weekdayOfAJD_(ajd) {
  return ((Math.floor(ajd + 0.5) % 7) + 8) % 7;
}

/** Hijri y/m/d -> real AJD, including the global day adjustment. */
function hijriToAJD_(year, month, day) {
  return new HijriDate(year, month, day).toAJD() + HIJRI_DAY_ADJUSTMENT;
}

/** Real AJD -> Hijri date, undoing the global day adjustment. */
function ajdToHijri_(ajd) {
  return HijriDate.fromAJD(ajd - HIJRI_DAY_ADJUSTMENT);
}

function isBlocked_(ajd, blocks) {
  for (var i = 0; i < blocks.length; i++) {
    if (ajd >= blocks[i][0] && ajd <= blocks[i][1]) return true;
  }
  return false;
}


/* ==========================================================================
 * 6. PARSING
 * ========================================================================== */

function firstValue_(v) {
  while (Array.isArray(v)) {
    if (!v.length) return "";
    v = v[0];
  }
  return (v === null || v === undefined) ? "" : v;
}

function normalise_(s) {
  return String(s).toLowerCase().replace(/[^a-z0-9]/g, "");
}

/** "Mon,Tue,Wed" -> [1,2,3] */
function parseDays_(days) {
  var raw = firstValue_(days);
  if (raw instanceof Date) throw new Error("PeriodCount: Days should be text like \"Mon,Tue,Wed\".");

  var parts = String(raw).split(/[,;/|+&\s]+/),
      out = [], seen = {};

  for (var i = 0; i < parts.length; i++) {
    var key = normalise_(parts[i]);
    if (!key) continue;
    if (!(key in DAY_ALIASES_)) {
      throw new Error('PeriodCount: "' + parts[i] + '" is not a weekday. Use Mon, Tue, Wed, Thu, Fri, Sat, Sun.');
    }
    var d = DAY_ALIASES_[key];
    if (!seen[d]) { seen[d] = true; out.push(d); }
  }
  return out;
}

/** Start/End -> real AJD. */
function resolveEndpoint_(value, system, label) {
  var v = firstValue_(value);

  if (v instanceof Date) return HijriDate.gregorianToAJD(v);

  var text = String(v).trim();
  if (!text) throw new Error("PeriodCount: " + label + " is empty.");

  if (system === "gregorian") {
    var g = parseGregorianText_(text);
    if (!g) throw new Error("PeriodCount: could not read " + label + ' "' + text + '" as a Gregorian date.');
    return HijriDate.gregorianToAJD(g);
  }

  var p = parseDatePart_(text, null);
  if (p.year === null) {
    throw new Error("PeriodCount: " + label + ' "' + text + '" needs a year, e.g. 1/2/1447.');
  }
  return hijriToAJD_(p.year, p.month, p.day);
}

function parseGregorianText_(text) {
  var m = text.match(/^(\d{1,2})\s*[\/\-.]\s*(\d{1,2})\s*[\/\-.]\s*(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  var d = new Date(text);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Reads one Hijri date such as:
 *   "20/4"   "20/4/1447"   "20 Rabi II"   "20 Rabi al-Aakhar 1447"   "Ramadaan"
 * Returns {day, month, year, wholeMonth}. day/year may be null.
 */
function parseDatePart_(text, contextYear) {
  var s = String(text).trim();
  if (!s) throw new Error("PeriodCount: empty date in Exclusions.");

  // numeric: d/m or d/m/yyyy
  var m = s.match(/^(\d{1,2})\s*\/\s*(\d{1,2})(?:\s*\/\s*(\d{1,4}))?$/);
  if (m) {
    return makePart_(Number(m[1]), Number(m[2]) - 1,
                     m[3] ? Number(m[3]) : contextYear, false, s);
  }

  // month name only -> the whole month
  var monthOnly = monthFromName_(s);
  if (monthOnly !== -1) {
    return makePart_(null, monthOnly, contextYear, true, s);
  }

  // "d MonthName [yyyy]"  /  "MonthName d[, yyyy]"
  var named = s.match(/^(\d{1,2})\s+(.+?)(?:\s+(\d{3,4}))?$/) ||
              s.match(/^(.+?)\s+(\d{1,2})(?:\s*,?\s*(\d{3,4}))?$/);
  if (named) {
    var a = named[1], b = named[2], yr = named[3] ? Number(named[3]) : contextYear;
    var day, monthIdx;
    if (/^\d+$/.test(a)) { day = Number(a); monthIdx = monthFromName_(b); }
    else                 { day = Number(b); monthIdx = monthFromName_(a); }
    if (monthIdx !== -1) return makePart_(day, monthIdx, yr, false, s);
  }

  // trailing year on a month name: "Ramadaan 1447"
  var mn = s.match(/^(.+?)\s+(\d{3,4})$/);
  if (mn) {
    var idx = monthFromName_(mn[1]);
    if (idx !== -1) return makePart_(null, idx, Number(mn[2]), true, s);
  }

  throw new Error('PeriodCount: could not read "' + s + '" in Exclusions. Use d/m, d/m/yyyy, ' +
                  '"20 Rabi II", or a month name like "Ramadaan".');
}

function makePart_(day, month, year, wholeMonth, original) {
  if (month < 0 || month > 11 || isNaN(month)) {
    throw new Error('PeriodCount: month out of range in "' + original + '". Use 1-12.');
  }
  if (day !== null && (isNaN(day) || day < 1 || day > 30)) {
    throw new Error('PeriodCount: day out of range in "' + original + '". Use 1-30.');
  }
  return { day: day, month: month, year: (year === undefined ? null : year), wholeMonth: !!wholeMonth };
}

function monthFromName_(text) {
  var key = normalise_(text);
  if (!key) return -1;

  for (var i = 0; i < MONTH_ALIASES_.length; i++) {
    for (var j = 0; j < MONTH_ALIASES_[i].length; j++) {
      if (MONTH_ALIASES_[i][j] === key) return i;
    }
  }
  // unique prefix match, e.g. "rama" -> Ramadaan
  if (key.length >= 3) {
    var found = -1;
    for (var a = 0; a < MONTH_ALIASES_.length; a++) {
      for (var b = 0; b < MONTH_ALIASES_[a].length; b++) {
        if (MONTH_ALIASES_[a][b].indexOf(key) === 0) {
          if (found !== -1 && found !== a) return -1;
          found = a;
        }
      }
    }
    return found;
  }
  return -1;
}


/* ==========================================================================
 * 7. EXCLUSIONS
 * ========================================================================== */

/**
 * Turns "[1/1-15/1,20/4]" into a list of [startAJD, endAJD] blocks.
 * Entries without a year repeat in every Hijri year the window touches.
 * Entries that fall outside the window are simply dropped.
 */
function buildExclusions_(exclusions, startAJD, endAJD) {
  var raw = firstValue_(exclusions);
  if (raw instanceof Date) {
    var one = HijriDate.gregorianToAJD(raw);
    return [[one, one]];
  }

  var text = String(raw).trim();
  if (!text) return [];

  // strip wrapping brackets/quotes
  text = text.replace(/^[\s\[\({'"]+/, "").replace(/[\s\]\)}'"]+$/, "").trim();
  if (!text || text === "-") return [];

  // protect hyphens inside month names ("Rabi al-Awwal") before splitting ranges
  text = text.replace(/([a-zA-Z])\s*-\s*([a-zA-Z])/g, "$1 $2");
  // normalise range separators to a single dash
  text = text.replace(/\s*(?:–|—|\.\.|\bto\b|\bthru\b|\bthrough\b|\btill\b|\buntil\b)\s*/gi, "-");

  var firstYear = ajdToHijri_(startAJD).getYear(),
      lastYear  = ajdToHijri_(endAJD).getYear(),
      blocks    = [],
      tokens    = text.split(/[,;\n]+/);

  for (var t = 0; t < tokens.length; t++) {
    var token = tokens[t].trim();
    if (!token) continue;

    var halves = token.split("-"),
        left   = halves[0].trim(),
        right  = (halves.length > 1) ? halves.slice(1).join("-").trim() : null;

    if (halves.length > 2 && right) {
      throw new Error('PeriodCount: too many dashes in "' + token + '". Use one range per entry.');
    }

    var lp = parseDatePart_(left, null),
        rp = right ? parseDatePart_(right, null) : null;

    var anchored = (lp.year !== null) || (rp && rp.year !== null);

    // A dated entry happens once; an undated one repeats every year in range.
    var years = anchored
      ? [lp.year !== null ? lp.year : rp.year]
      : yearSpan_(firstYear - 1, lastYear + 1);

    for (var y = 0; y < years.length; y++) {
      var block = makeBlock_(lp, rp, years[y]);
      if (block && block[1] >= startAJD && block[0] <= endAJD) blocks.push(block);
    }
  }
  return blocks;
}

function yearSpan_(from, to) {
  var out = [];
  for (var y = from; y <= to; y++) out.push(y);
  return out;
}

/** Builds one [startAJD, endAJD] block for a given Hijri year. */
function makeBlock_(lp, rp, year) {
  var startYear = (lp.year !== null) ? lp.year : year;

  var fromDay = (lp.day === null) ? 1 : clampDay_(lp.day, startYear, lp.month);
  var fromAJD = hijriToAJD_(startYear, lp.month, fromDay);

  var toAJD;
  if (!rp) {
    // single date, or a whole month
    var lastDay = lp.wholeMonth
      ? HijriDate.daysInMonth(startYear, lp.month)
      : fromDay;
    toAJD = hijriToAJD_(startYear, lp.month, lastDay);
  } else {
    var endYear = (rp.year !== null) ? rp.year : startYear;
    // range that wraps past Zilhaj, e.g. 25/12 - 5/1
    if (rp.year === null &&
        (rp.month < lp.month || (rp.month === lp.month && rp.day !== null &&
         lp.day !== null && rp.day < lp.day))) {
      endYear = startYear + 1;
    }
    var endDay = (rp.day === null || rp.wholeMonth)
      ? HijriDate.daysInMonth(endYear, rp.month)
      : clampDay_(rp.day, endYear, rp.month);
    toAJD = hijriToAJD_(endYear, rp.month, endDay);
  }

  if (toAJD < fromAJD) return null;
  return [fromAJD, toAJD];
}

/** 30 Zilhaj in a non-Kabisa year becomes 29, rather than spilling over. */
function clampDay_(day, year, month) {
  var max = HijriDate.daysInMonth(year, month);
  return (day > max) ? max : day;
}
