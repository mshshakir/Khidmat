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

/** Gregorian month names for the Gregorian exclusion list. */
var GREGORIAN_MONTHS_ = [
  ["jan", "january"], ["feb", "february"], ["mar", "march"], ["apr", "april"],
  ["may"], ["jun", "june"], ["jul", "july"], ["aug", "august"],
  ["sep", "sept", "september"], ["oct", "october"], ["nov", "november"],
  ["dec", "december"]
];

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
 * @param {"Tue,Wed,Thu+1,Fri,Sat"} schedule Days with periods. Add "+1" for a double period, "x2" for an explicit count, or a range like "Mon-Thu".
 * @param {"[1/1-15/1,20/4]"} hijriExclusions Hijri dates/ranges to skip, e.g. miqaat. Blank for none.
 * @param {"[25/12,15/8]"} gregorianExclusions Gregorian dates/ranges to skip, written d/m or d/m/yyyy. Blank for none.
 * @param {"1/10/1448"} start First day of the window, inclusive.
 * @param {"29/6/1449"} end Last day of the window, inclusive.
 * @param {5} buffer Optional. Periods to hold back, as a number (5) or a percentage ("10%").
 * @param {"hijri"} dateSystem Optional. How to read TEXT start/end: "hijri" (default) or "gregorian".
 * @return {number} The number of periods.
 * @customfunction
 */
function PeriodCount(schedule, hijriExclusions, gregorianExclusions, start, end, buffer, dateSystem) {
  var result = periodScan_(schedule, hijriExclusions, gregorianExclusions, start, end, dateSystem);
  return applyBuffer_(result.count, buffer);
}

/**
 * Counts periods before any buffer is taken off. Useful next to PeriodCount
 * to show what the buffer costs.
 *
 * @param {"Tue,Wed,Thu+1"} schedule Days with periods.
 * @param {"[1/1-15/1]"} hijriExclusions Hijri dates/ranges to skip (blank for none).
 * @param {"[25/12]"} gregorianExclusions Gregorian dates/ranges to skip (blank for none).
 * @param {"1/10/1448"} start First day of the window, inclusive.
 * @param {"29/6/1449"} end Last day of the window, inclusive.
 * @param {"hijri"} dateSystem Optional. "hijri" (default) or "gregorian".
 * @return {number} Periods before the buffer.
 * @customfunction
 */
function PeriodCountGross(schedule, hijriExclusions, gregorianExclusions, start, end, dateSystem) {
  return periodScan_(schedule, hijriExclusions, gregorianExclusions, start, end, dateSystem).count;
}

/**
 * Counts teaching DAYS rather than periods, ignoring double periods.
 *
 * @param {"Tue,Wed,Thu+1"} schedule Days with periods.
 * @param {"[1/1-15/1]"} hijriExclusions Hijri dates/ranges to skip (blank for none).
 * @param {"[25/12]"} gregorianExclusions Gregorian dates/ranges to skip (blank for none).
 * @param {"1/10/1448"} start First day of the window, inclusive.
 * @param {"29/6/1449"} end Last day of the window, inclusive.
 * @param {"hijri"} dateSystem Optional. "hijri" (default) or "gregorian".
 * @return {number} Number of days on which the subject meets.
 * @customfunction
 */
function PeriodDayCount(schedule, hijriExclusions, gregorianExclusions, start, end, dateSystem) {
  return periodScan_(schedule, hijriExclusions, gregorianExclusions, start, end, dateSystem).days;
}

/**
 * Lists every day counted, for checking your setup.
 *
 * @param {"Tue,Wed,Thu+1"} schedule Days with periods.
 * @param {"[1/1-15/1]"} hijriExclusions Hijri dates/ranges to skip (blank for none).
 * @param {"[25/12]"} gregorianExclusions Gregorian dates/ranges to skip (blank for none).
 * @param {"1/10/1448"} start First day of the window, inclusive.
 * @param {"29/6/1449"} end Last day of the window, inclusive.
 * @param {"hijri"} dateSystem Optional. "hijri" (default) or "gregorian".
 * @return {Array} Hijri date, month, weekday, Gregorian date and periods for each day counted.
 * @customfunction
 */
function PeriodDates(schedule, hijriExclusions, gregorianExclusions, start, end, dateSystem) {
  var result = periodScan_(schedule, hijriExclusions, gregorianExclusions, start, end, dateSystem);
  if (!result.dates.length) return [["No periods found"]];

  var rows = [["Hijri", "Month", "Weekday", "Gregorian", "Periods"]];
  for (var i = 0; i < result.dates.length; i++) {
    var ajd = result.dates[i].ajd,
        h   = ajdToHijri_(ajd);
    rows.push([
      h.getDate() + "/" + (h.getMonth() + 1) + "/" + h.getYear(),
      HijriDate.getShortMonthName(h.getMonth()),
      ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"][weekdayOfAJD_(ajd)],
      HijriDate.ajdToGregorian(ajd),
      result.dates[i].weight
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
  var p = parseHijriPart_(String(firstValue_(hijri)));
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
  return formatHijri_(ajdToHijri_(HijriDate.gregorianToAJD(v)));
}


/* ==========================================================================
 * 5. CORE ENGINE
 * ========================================================================== */

function periodScan_(schedule, hijriExclusions, gregExclusions, start, end, dateSystem) {
  // Argument order changed: the two exclusion lists now sit together, before
  // start and end. An older four-argument formula lands here with End empty.
  if (isBlank_(end) && !isBlank_(start)) {
    throw new Error('PeriodCount: End is empty. The argument order is ' +
      '(schedule, hijriExclusions, gregorianExclusions, start, end, [buffer], [dateSystem]). ' +
      'If this formula was written for an earlier version, insert "" for the Gregorian list ' +
      'before the start date.');
  }

  var system = String(firstValue_(dateSystem) || DEFAULT_DATE_SYSTEM)
                 .toLowerCase().indexOf("greg") === 0 ? "gregorian" : "hijri";

  var weights = parseSchedule_(schedule);

  var startAJD = resolveEndpoint_(start, system, "Start"),
      endAJD   = resolveEndpoint_(end,   system, "End");

  if (startAJD > endAJD) {
    var sh = ajdToHijri_(startAJD), eh = ajdToHijri_(endAJD), hint = "";
    // Most common cause: an academic year that crosses the Hijri new year,
    // written with the same year at both ends.
    if (eh.getYear() === sh.getYear() && eh.getMonth() < sh.getMonth()) {
      hint = " If your year runs across the Hijri new year, End should be " +
             eh.getDate() + "/" + (eh.getMonth() + 1) + "/" + (eh.getYear() + 1) + ".";
    }
    throw new Error("PeriodCount: Start (" + formatHijri_(sh) + ") is after End (" +
                    formatHijri_(eh) + ")." + hint);
  }

  var span = Math.round(endAJD - startAJD) + 1;
  if (span > MAX_WINDOW_DAYS) {
    throw new Error("PeriodCount: window is " + span + " days; max is " + MAX_WINDOW_DAYS + ".");
  }

  var blocks = buildExclusions_(hijriExclusions, startAJD, endAJD, "hijri")
        .concat(buildExclusions_(gregExclusions, startAJD, endAJD, "gregorian"));

  var hits = [], total = 0;
  for (var ajd = startAJD; ajd <= endAJD; ajd += 1) {
    var w = weights[weekdayOfAJD_(ajd)];
    if (!w) continue;
    if (isBlocked_(ajd, blocks)) continue;
    hits.push({ ajd: ajd, weight: w });
    total += w;
  }

  return { count: total, days: hits.length, dates: hits };
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

/** "1 Shawwal 1448" - used in error messages. */
function formatHijri_(h) {
  return h.getDate() + " " + HijriDate.getShortMonthName(h.getMonth()) + " " + h.getYear();
}

function isBlocked_(ajd, blocks) {
  for (var i = 0; i < blocks.length; i++) {
    if (ajd >= blocks[i][0] && ajd <= blocks[i][1]) return true;
  }
  return false;
}

/**
 * Takes the buffer off a raw count.
 * Accepts a plain number of periods, or a percentage such as "10%".
 * A negative buffer adds periods. The result never drops below zero.
 */
function applyBuffer_(count, buffer) {
  var raw = firstValue_(buffer);
  if (raw === "" || raw === null || raw === undefined) return count;

  var text = String(raw).trim();
  if (!text) return count;

  var out;
  if (/%$/.test(text)) {
    var pct = Number(text.replace(/%$/, "").trim());
    if (isNaN(pct)) throw new Error('PeriodCount: buffer "' + text + '" is not a percentage.');
    out = Math.floor(count * (1 - pct / 100));
  } else {
    var n = Number(text);
    // Sheets may hand a percent-formatted cell over as a fraction, e.g. 0.1
    if (isNaN(n)) throw new Error('PeriodCount: buffer "' + text + '" is not a number or percentage.');
    if (typeof raw === "number" && raw > 0 && raw < 1) out = Math.floor(count * (1 - raw));
    else out = count - n;
  }
  return (out < 0) ? 0 : out;
}


/* ==========================================================================
 * 6. PARSING
 * ========================================================================== */

function isBlank_(v) {
  var x = firstValue_(v);
  return x === "" || x === null || x === undefined;
}

function firstValue_(v) {
  while (Array.isArray(v)) {
    if (!v.length) return "";
    v = v[0];
  }
  return (v === null || v === undefined) ? "" : v;
}

/** Flattens a cell range into one comma-separated string. */
function flattenToText_(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return v;
  if (!Array.isArray(v)) return String(v);
  var parts = [];
  (function walk(x) {
    if (Array.isArray(x)) { for (var i = 0; i < x.length; i++) walk(x[i]); return; }
    if (x === null || x === undefined) return;
    var t = String(x).trim();
    if (t) parts.push(t);
  })(v);
  return parts.join(",");
}

function normalise_(s) {
  return String(s).toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Reads the schedule and returns periods-per-weekday, indexed 0 = Sunday.
 *
 *   "Mon,Tue,Wed"        -> Mon 1, Tue 1, Wed 1
 *   "Mon+1,Tue"          -> Mon 2, Tue 1        (+N adds N extra periods)
 *   "Mon x2, Tue"        -> Mon 2, Tue 1        (xN sets the count outright)
 *   "Mon-Thu"            -> Mon..Thu, 1 each
 *   "Mon-Wed+1"          -> Mon..Wed, 2 each
 *   "Thurs, THU"         -> summed, so Thu 2
 *
 * Case, spacing and punctuation are all ignored.
 */
function parseSchedule_(schedule) {
  var raw = firstValue_(schedule);
  if (raw instanceof Date) {
    throw new Error('PeriodCount: Days should be text like "Mon,Tue,Wed".');
  }

  // Normalise the multiplier notations to a single form before tokenising.
  var text = String(raw)
        .replace(/\u00D7/g, "x")
        .replace(/([A-Za-z])\s*[x*]\s*(\d+)/gi, "$1*$2")   // "Mon x2" / "Mon * 2"
        .replace(/([A-Za-z])\s*\+\s*(\d+)/g, "$1+$2")      // "Mon + 1"
        .replace(/([A-Za-z])\s*(\d+)/g, "$1*$2");          // "Mon 2" / "Mon2"

  var re = /([A-Za-z][A-Za-z'\u2019\-]*)(?:([+*])(\d+))?/g,
      weights = {}, spans = [], m, found = false;

  while ((m = re.exec(text)) !== null) {
    spans.push([m.index, m.index + m[0].length]);

    var days = daysFromToken_(m[1]),
        n    = m[3] ? Number(m[3]) : null,
        w;

    if (m[2] === "+")      w = 1 + n;
    else if (m[2] === "*") w = n;
    else                   w = 1;

    if (w < 0 || w > 50) {
      throw new Error('PeriodCount: "' + m[0] + '" gives ' + w + ' periods a day, which looks wrong.');
    }

    for (var i = 0; i < days.length; i++) {
      weights[days[i]] = (weights[days[i]] || 0) + w;
      found = true;
    }
  }

  // Anything left between the matches must be a separator, not a stray word.
  var rest = "", pos = 0;
  for (var k = 0; k < spans.length; k++) {
    rest += text.slice(pos, spans[k][0]);
    pos = spans[k][1];
  }
  rest += text.slice(pos);
  if (/[^,;|&\/\s]/.test(rest)) {
    throw new Error('PeriodCount: could not read "' + rest.trim() + '" in Days.');
  }

  if (!found) throw new Error('PeriodCount: no weekdays given, e.g. "Mon,Tue,Wed".');
  return weights;
}

/** One schedule token -> a list of weekday numbers. Handles "Mon-Thu" ranges. */
function daysFromToken_(token) {
  var key = normalise_(token);
  if (key in DAY_ALIASES_) return [DAY_ALIASES_[key]];

  // day range, e.g. "Mon-Thu" or "Sat-Mon"
  if (token.indexOf("-") !== -1) {
    var halves = token.split("-");
    if (halves.length === 2) {
      var a = normalise_(halves[0]), b = normalise_(halves[1]);
      if (a in DAY_ALIASES_ && b in DAY_ALIASES_) {
        var from = DAY_ALIASES_[a], to = DAY_ALIASES_[b], out = [], d = from;
        for (var guard = 0; guard < 7; guard++) {
          out.push(d);
          if (d === to) break;
          d = (d + 1) % 7;
        }
        return out;
      }
    }
  }

  throw new Error('PeriodCount: "' + token + '" is not a weekday. ' +
                  'Use Mon, Tue, Wed, Thu, Fri, Sat, Sun.');
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

  var p = parseHijriPart_(text);
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
function parseHijriPart_(text) {
  return parseDatePart_(text, monthFromName_, 1, 1000, 1700, "Hijri");
}

/** Same, for Gregorian entries: "25/12", "25/12/2027", "December". */
function parseGregorianPart_(text) {
  return parseDatePart_(text, gregorianMonthFromName_, 4, 1900, 2200, "Gregorian");
}

function parseDatePart_(text, monthLookup, yearDigits, minYear, maxYear, calName) {
  var s = String(text).trim();
  if (!s) throw new Error("PeriodCount: empty date in Exclusions.");

  var yearPattern = (yearDigits === 4) ? "\\d{4}" : "\\d{1,4}";

  // numeric: d/m or d/m/yyyy
  var m = s.match(new RegExp("^(\\d{1,2})\\s*\\/\\s*(\\d{1,2})(?:\\s*\\/\\s*(" + yearPattern + "))?$"));
  if (m) {
    return makePart_(Number(m[1]), Number(m[2]) - 1, m[3] ? Number(m[3]) : null, false, s, minYear, maxYear, calName);
  }

  // month name only -> the whole month
  var monthOnly = monthLookup(s);
  if (monthOnly !== -1) return makePart_(null, monthOnly, null, true, s, minYear, maxYear, calName);

  // "d MonthName [yyyy]"  /  "MonthName d[, yyyy]"
  var named = s.match(new RegExp("^(\\d{1,2})\\s+(.+?)(?:\\s+(" + yearPattern + "))?$")) ||
              s.match(new RegExp("^(.+?)\\s+(\\d{1,2})(?:\\s*,?\\s*(" + yearPattern + "))?$"));
  if (named) {
    var a = named[1], b = named[2], yr = named[3] ? Number(named[3]) : null, day, monthIdx;
    if (/^\d+$/.test(a)) { day = Number(a); monthIdx = monthLookup(b); }
    else                 { day = Number(b); monthIdx = monthLookup(a); }
    if (monthIdx !== -1) return makePart_(day, monthIdx, yr, false, s, minYear, maxYear, calName);
  }

  // trailing year on a month name: "Ramadaan 1447", "December 2027"
  var mn = s.match(new RegExp("^(.+?)\\s+(" + yearPattern + ")$"));
  if (mn) {
    var idx = monthLookup(mn[1]);
    if (idx !== -1) return makePart_(null, idx, Number(mn[2]), true, s, minYear, maxYear, calName);
  }

  throw new Error('PeriodCount: could not read "' + s + '" in Exclusions. Use d/m, d/m/yyyy, ' +
                  '"20 Rabi II", or a month name like "Ramadaan".');
}

function makePart_(day, month, year, wholeMonth, original, minYear, maxYear, calName) {
  if (month < 0 || month > 11 || isNaN(month)) {
    throw new Error('PeriodCount: month out of range in "' + original + '". Use 1-12.');
  }
  if (day !== null && (isNaN(day) || day < 1 || day > 31)) {
    throw new Error('PeriodCount: day out of range in "' + original + '". Use 1-31.');
  }
  if (year !== null && (year < minYear || year > maxYear)) {
    var swap = (calName === "Gregorian")
      ? ' Did you mean to put it in the Hijri list?'
      : ' Did you mean to put it in the Gregorian list?';
    throw new Error('PeriodCount: "' + original + '" has year ' + year +
                    ', which is not a ' + calName + ' year (expected ' +
                    minYear + '-' + maxYear + ').' + swap);
  }
  return { day: day, month: month, year: year, wholeMonth: !!wholeMonth };
}

function lookupMonth_(text, table) {
  var key = normalise_(text);
  if (!key) return -1;

  for (var i = 0; i < table.length; i++) {
    for (var j = 0; j < table[i].length; j++) {
      if (table[i][j] === key) return i;
    }
  }
  if (key.length >= 3) {
    var found = -1;
    for (var a = 0; a < table.length; a++) {
      for (var b = 0; b < table[a].length; b++) {
        if (table[a][b].indexOf(key) === 0) {
          if (found !== -1 && found !== a) return -1;
          found = a;
        }
      }
    }
    return found;
  }
  return -1;
}

function monthFromName_(text)          { return lookupMonth_(text, MONTH_ALIASES_); }
function gregorianMonthFromName_(text) { return lookupMonth_(text, GREGORIAN_MONTHS_); }


/* ==========================================================================
 * 7. EXCLUSIONS
 * ========================================================================== */

/** The two calendars share one exclusion engine through these adapters. */
var HIJRI_CAL_ = {
  name: "Hijri",
  parse: parseHijriPart_,
  daysInMonth: function (y, m) { return HijriDate.daysInMonth(y, m); },
  toAJD: function (y, m, d) { return hijriToAJD_(y, m, d); },
  yearOf: function (ajd) { return ajdToHijri_(ajd).getYear(); }
};

var GREGORIAN_CAL_ = {
  name: "Gregorian",
  parse: parseGregorianPart_,
  daysInMonth: function (y, m) { return new Date(y, m + 1, 0).getDate(); },
  toAJD: function (y, m, d) { return HijriDate.gregorianToAJD(new Date(y, m, d)); },
  yearOf: function (ajd) { return HijriDate.ajdToGregorian(ajd).getFullYear(); }
};

/**
 * Turns an exclusion list into [startAJD, endAJD] blocks.
 *
 * Entries without a year repeat in every year the window touches.
 * Entries outside the window are dropped, never fatal.
 * A "G:" tag switches to Gregorian for the rest of that cell, "H:" back to Hijri.
 */
function buildExclusions_(exclusions, startAJD, endAJD, defaultMode) {
  var raw = firstValue_(exclusions);
  if (raw instanceof Date) {
    var one = HijriDate.gregorianToAJD(raw);
    return [[one, one]];
  }

  var text = flattenToText_(exclusions);
  if (text instanceof Date) {
    var d = HijriDate.gregorianToAJD(text);
    return [[d, d]];
  }
  text = String(text).trim();
  if (!text) return [];

  text = text.replace(/^[\s\[\({'"]+/, "").replace(/[\s\]\)}'"]+$/, "").trim();
  if (!text || text === "-") return [];

  // protect hyphens inside month names ("Rabi al-Awwal") before splitting ranges
  text = text.replace(/([a-zA-Z])\s*-\s*([a-zA-Z])/g, "$1 $2");
  text = text.replace(/\s*(?:\u2013|\u2014|\.\.|\bto\b|\bthru\b|\bthrough\b|\btill\b|\buntil\b)\s*/gi, "-");

  var mode   = (defaultMode === "gregorian") ? GREGORIAN_CAL_ : HIJRI_CAL_,
      blocks = [],
      tokens = text.split(/[,;\n]+/);

  for (var t = 0; t < tokens.length; t++) {
    var token = tokens[t].trim();
    if (!token) continue;

    // inline calendar tag, e.g. "G: 25/12/2027"
    var tag = token.match(/^([a-zA-Z]+)\s*:\s*(.*)$/);
    if (tag) {
      var word = tag[1].toLowerCase();
      if (word === "g" || word === "greg" || word === "gregorian") { mode = GREGORIAN_CAL_; token = tag[2].trim(); }
      else if (word === "h" || word === "hijri" || word === "misri") { mode = HIJRI_CAL_; token = tag[2].trim(); }
      if (!token) continue;
    }

    var halves = token.split("-"),
        left   = halves[0].trim(),
        right  = (halves.length > 1) ? halves.slice(1).join("-").trim() : null;

    if (halves.length > 2 && right) {
      throw new Error('PeriodCount: too many dashes in "' + token + '". Use one range per entry.');
    }

    var lp = mode.parse(left),
        rp = right ? mode.parse(right) : null;

    var anchored = (lp.year !== null) || (rp && rp.year !== null),
        years;

    if (anchored) {
      years = [lp.year !== null ? lp.year : rp.year];
    } else {
      var first = mode.yearOf(startAJD), last = mode.yearOf(endAJD);
      years = [];
      for (var y = first - 1; y <= last + 1; y++) years.push(y);
    }

    for (var i = 0; i < years.length; i++) {
      var block = makeBlock_(lp, rp, years[i], mode);
      if (block && block[1] >= startAJD && block[0] <= endAJD) blocks.push(block);
    }
  }
  return blocks;
}

/** Builds one [startAJD, endAJD] block for a given year. */
function makeBlock_(lp, rp, year, cal) {
  var startYear = (lp.year !== null) ? lp.year : year,
      startMax  = cal.daysInMonth(startYear, lp.month);

  // A single date that does not exist this year (30 Zilhaj in a short year,
  // 29 February in a common year) simply does not happen. Skip it.
  if (!rp && !lp.wholeMonth && lp.day !== null && lp.day > startMax) return null;

  var fromDay = (lp.day === null) ? 1 : Math.min(lp.day, startMax),
      fromAJD = cal.toAJD(startYear, lp.month, fromDay),
      toAJD;

  if (!rp) {
    var lastDay = lp.wholeMonth ? startMax : fromDay;
    toAJD = cal.toAJD(startYear, lp.month, lastDay);
  } else {
    var endYear = (rp.year !== null) ? rp.year : startYear;
    // range that wraps past the end of the year, e.g. 25/12 - 5/1
    if (rp.year === null &&
        (rp.month < lp.month ||
         (rp.month === lp.month && rp.day !== null && lp.day !== null && rp.day < lp.day))) {
      endYear = startYear + 1;
    }
    var endMax = cal.daysInMonth(endYear, rp.month),
        endDay = (rp.day === null || rp.wholeMonth) ? endMax : Math.min(rp.day, endMax);
    toAJD = cal.toAJD(endYear, rp.month, endDay);
  }

  if (toAJD < fromAJD) return null;
  return [fromAJD, toAJD];
}
