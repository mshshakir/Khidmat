# PeriodCount — Hijri period counter for Google Sheets

A Google Apps Script custom function that counts how many times given weekdays occur between two Hijri dates, skipping holidays and breaks you define in Hijri terms.

**Its main use is working out how many periods a subject actually gets in an academic year**, once miqaat and other custom Hijri dates are taken out. Set the term dates once, list the miqaat once, and every subject's period count follows from its weekly timetable.

*"Physics meets on Mon, Tue and Wed. How many periods is that this academic year, given that Ashara Mubaraka and 20 Rabi al-Aakhar are off?"*

```
=PeriodCount("Mon,Tue,Wed", "[1/1-15/1,20/4]", "1/2/1447", "29/12/1447")
```

→ `138`

---

## Contents

- [Install](#install)
- [Syntax](#syntax)
- [Exclusion syntax](#exclusion-syntax)
- [Examples](#examples)
- [Helper functions](#helper-functions)
- [Configuration](#configuration)
- [Accepted names](#accepted-names)
- [The calendar](#the-calendar)
- [Errors](#errors)
- [Testing](#testing)
- [Credits](#credits)

---

## Install

1. Open your spreadsheet.
2. **Extensions → Apps Script**.
3. Delete the placeholder `myFunction`, paste the contents of `PeriodCount.gs`.
4. **Save** (💾). Close the editor tab.
5. Type `=PeriodCount(` in any cell — autocomplete should offer it.

No libraries, no manifest changes, no permissions prompt. The script is self-contained and does not touch the network or your data.

---

## Syntax

```
=PeriodCount(days, exclusions, start, end, [dateSystem])
```

| Argument | Required | Description |
|---|---|---|
| `days` | yes | Weekdays the subject has periods on, e.g. `"Mon,Tue,Wed"` |
| `exclusions` | yes (may be blank) | Hijri dates and ranges to skip, e.g. `"[1/1-15/1,20/4]"`. Pass `""` for none |
| `start` | yes | First day of the window, **inclusive** |
| `end` | yes | Last day of the window, **inclusive** |
| `dateSystem` | no | How to read *text* `start`/`end`: `"hijri"` (default) or `"gregorian"` |

**Returns:** a plain number.

### Dates in, dates out

- **Text** dates (`"1/2/1447"`) are read as **Hijri** `d/m/yyyy` by default.
- **Real date cells** are always read as **Gregorian** and converted internally.
- Month names work anywhere a date is accepted: `"1 Moharram 1447"`, `"20 Rabi II 1447"`.
- If `start` is later than `end`, the two are swapped rather than returning an error.

Cell references work everywhere, so a timetable can be driven entirely from cells:

```
=PeriodCount(B2, $F$1, $B$1, $C$1)
```

---

## Exclusion syntax

Exclusions are your miqaat and any other custom Hijri dates the school doesn't teach on — single days, multi-day ranges, or whole months. They all go in one string. Wrapping brackets are optional; entries are comma-separated.

Because miqaat fall on fixed Hijri dates but drift through the Gregorian week, an entry written once keeps working in every year — you never have to restate it.

| Form | Meaning |
|---|---|
| `20/4` | 20 Rabi al-Aakhar, **every year** in the window |
| `1/1-15/1` | 1 to 15 Moharram inclusive, **every year** |
| `20/4/1447` | 20 Rabi al-Aakhar 1447 only, **once** |
| `1/1/1447-15/1/1447` | That range in 1447 only, **once** |
| `Ramadaan` | The **whole month**, every year |
| `Ramadaan 1447` | The whole month of Ramadaan 1447, once |
| `1 Moharram - 15 Moharram` | Same as `1/1-15/1` |
| `25/12-5/1` | Wraps across the year end into the next Moharram |

### Rules

**Undated entries repeat annually.** `1/1-15/1` blocks Ashara in every Hijri year the window touches, so one formula covers a multi-year window. Add a year to pin an entry to a single occurrence.

**Out-of-range entries are ignored, never fatal.** If the window starts in Safar and an exclusion names Moharram, the entry is silently dropped for that year — it does not throw, and it does not shift anything. It will still apply if the window is long enough to reach the *following* Moharram.

**Day counts are clamped.** `30/12` in a non-Kabisa year (where Zilhaj has 29 days) resolves to 29 rather than spilling into Moharram.

**Range separators.** Use `-` between two dates. Also accepted: `–`, `—`, `..`, `to`, `thru`, `through`, `till`, `until`. Hyphens inside month names (`Rabi al-Awwal`) are protected, so they don't split a range by accident.

---

## Examples

All figures below are actual output.

| Formula | Result |
|---|---|
| `=PeriodCount("Mon,Tue,Wed","","1/1/1447","29/12/1447")` | `150` |
| `=PeriodCount("Mon,Tue,Wed","[1/1-15/1,20/4]","1/1/1447","29/12/1447")` | `144` |
| `=PeriodCount("Mon,Tue,Wed","[1/1-15/1,20/4]","1/2/1447","29/12/1447")` | `138` |
| `=PeriodCount("Mon,Tue,Wed","[1/1-15/1]","1/1/1446","29/12/1448")` | `437` |
| `=PeriodCount("Mon,Tue,Wed","[1/1/1447-15/1/1447]","1/1/1446","29/12/1448")` | `450` |
| `=PeriodCount("Mon,Tue,Wed,Thu","[25/12-5/1]","1/1/1446","29/12/1448")` | `589` |
| `=PeriodCount("Mon,Wed","Ramadaan","1/1/1447","29/12/1447")` | `91` |
| `=PeriodCount("Mon,Tue","Rabi al-Awwal","1/1/1447","29/12/1447")` | `91` |
| `=PeriodCount("Mon,Tue,Wed","1 Moharram - 15 Moharram","1/1/1447","29/12/1447")` | `144` |
| `=PeriodCount("Mon","","1/1/1447","29/12/1447")` | `50` |

Note rows 4 and 5: identical windows, identical dates, but the undated exclusion repeats across all three years (`437`) while the dated one fires once (`450`).

### A worked academic year

Put the year's boundaries and the miqaat list in fixed cells, then drag one formula down the subject column.

| | A | B | C |
|---|---|---|---|
| **1** | Year start | Year end | Miqaat list |
| **2** | `1/2/1447` | `29/12/1447` | `[1/1-15/1,20/4,Ramadaan]` |
| **3** | **Subject** | **Days** | **Periods** |
| **4** | Physics | `Mon,Tue,Wed` | `=PeriodCount(B4,$C$2,$A$2,$B$2)` |
| **5** | Chemistry | `Tue,Thu` | `=PeriodCount(B5,$C$2,$A$2,$B$2)` |
| **6** | Maths | `Mon,Wed,Fri` | `=PeriodCount(B6,$C$2,$A$2,$B$2)` |

Because `$C$2` is shared, adding a miqaat to that one cell re-costs every subject in the sheet at once. And because undated entries repeat annually, the same sheet keeps working when you roll the year start and end forward — no need to restate the recurring miqaat.

Two things this makes easy:

- **Syllabus planning** — compare periods available against periods required per subject, before the year starts.
- **Miqaat impact** — duplicate the miqaat cell, delete one entry, and the difference in the totals is exactly what that occasion costs each subject.

---

## Helper functions

### `PeriodDates(days, exclusions, start, end, [dateSystem])`

Same arguments as `PeriodCount`, but spills a table of every day counted — Hijri date, month name, weekday, and Gregorian date. The fastest way to check a formula before trusting the number.

```
=PeriodDates("Mon","Ramadaan","1/1/1447","29/1/1447")
```

| Hijri | Month | Weekday | Gregorian |
|---|---|---|---|
| 5/1/1447 | Moharram | Mon | 30/06/2025 |
| 12/1/1447 | Moharram | Mon | 07/07/2025 |
| … | | | |

### `HijriToGregorian(hijri)`

```
=HijriToGregorian("1/1/1447")       → 26 June 2025
=HijriToGregorian("1 Moharram 1447") → 26 June 2025
```

Returns a real date value; format the cell as a date.

### `GregorianToHijri(gregorian)`

```
=GregorianToHijri(DATE(2025,6,26))  → "1 Moharram 1447"
```

Takes a real date cell and returns a Hijri date string.

---

## Configuration

Three constants at the top of `PeriodCount.gs`:

```js
var HIJRI_DAY_ADJUSTMENT = 0;
var DEFAULT_DATE_SYSTEM  = "hijri";
var MAX_WINDOW_DAYS      = 40000;
```

| Constant | Purpose |
|---|---|
| `HIJRI_DAY_ADJUSTMENT` | Shifts the entire Hijri calendar by whole days. Set to `-1` or `1` if your local calendar begins months a day earlier or later than the arithmetic result. Affects which weekday each Hijri date lands on, so set it before trusting counts |
| `DEFAULT_DATE_SYSTEM` | Change to `"gregorian"` if you'd rather type Gregorian text dates without passing the 5th argument each time |
| `MAX_WINDOW_DAYS` | Guard against a typo producing a thousand-year loop. Raise it if you genuinely need a longer window |

---

## Accepted names

### Weekdays

Case-insensitive; punctuation and spacing are ignored. Separate with commas, spaces, semicolons, slashes, `+` or `&`.

| Day | Accepted |
|---|---|
| Sunday | `Sun`, `Sunday`, `Ahad`, `al-Ahad` |
| Monday | `Mon`, `Monday`, `Ithnain`, `Isnain` |
| Tuesday | `Tue`, `Tues`, `Tuesday`, `Thulatha`, `Salasa` |
| Wednesday | `Wed`, `Weds`, `Wednesday`, `Arbaa`, `Arbia` |
| Thursday | `Thu`, `Thur`, `Thurs`, `Thursday`, `Khamis` |
| Friday | `Fri`, `Friday`, `Jumua`, `Jumuah`, `Jummah` |
| Saturday | `Sat`, `Saturday`, `Sabt` |

### Months

| # | Name | Also accepted |
|---|---|---|
| 1 | Moharram | Muharram, Moharam, Moharram al-Haraam |
| 2 | Safar | Safar al-Muzaffar |
| 3 | Rabi I | Rabi al-Awwal, Rabi ul Awwal |
| 4 | Rabi II | Rabi al-Aakhar, Rabi ul Akhar, Rabi al-Thani |
| 5 | Jumada I | Jumada al-Ula, Jamadi ul Awwal |
| 6 | Jumada II | Jumada al-Ukhra, Jamadi ul Akhar |
| 7 | Rajab | Rajab al-Asab |
| 8 | Shabaan | Shaban, Shabaan al-Karim |
| 9 | Ramadaan | Ramadan, Ramzan, Ramadhan |
| 10 | Shawwal | Shawal, Shawwal al-Mukarram |
| 11 | Zilqadah | Zulqadah, Dhul Qadah |
| 12 | Zilhaj | Zilhajj, Zulhaj, Dhul Hijjah |

Unambiguous prefixes also resolve, so `Rama` → Ramadaan and `Zilh` → Zilhaj.

---

## The calendar

This is an **arithmetic (tabular)** Hijri calendar, not a sighting-based one. Every date is computed, so results are deterministic and reproducible — but they may differ by a day or two from a locally announced moon sighting. Use `HIJRI_DAY_ADJUSTMENT` to align if needed.

The implementation is a port of `hijri_date.js`, which uses:

- **Kabisa (leap) year remainders:** 2, 5, 8, 10, 13, 16, 19, 21, 24, 27, 29 in each 30-year cycle — the Fatimid / Dawoodi Bohra variant. This differs from the more common civil tabular rule (which leaps at 7, 18 and 26 instead of 8, 19 and 27), so dates will not always match a generic Hijri converter.
- **Epoch:** 1 Moharram 1 AH = AJD 1948438.5 = Thursday 15 July 622 CE.
- **Month lengths:** 30 and 29 alternating, with Zilhaj taking 30 days in a Kabisa year.
- **Weekday convention:** `(AJD + 1.5) % 7`, where `0` = Sunday — matching `dayOfWeek()` in `hijri_calendar.js`.

Reference point: **1 Moharram 1447 = Thursday 26 June 2025**.

### Upstream bug fix

`fromAJD()` in the original `hijri_date.js` returns day `0` of the following Moharram for the last day of any year whose remainder mod 30 is 29 — that is, 30 Zilhaj **1409**, **1439** and **1469**. The port carries a fix (marked `BOUNDARY FIX`) that rolls the result back to the correct last day of the previous month. If you use `hijri_date.js` elsewhere, apply the same fix there.

---

## Errors

Errors surface as `#ERROR!` with the message in the cell tooltip.

| Message | Cause |
|---|---|
| `"Funday" is not a weekday…` | Unrecognised day name in `days` |
| `no weekdays given…` | `days` was empty |
| `day out of range in "99/1". Use 1-30.` | Day outside 1–30 |
| `month out of range in "…". Use 1-12.` | Month outside 1–12 |
| `Start "1/1" needs a year…` | `start`/`end` must include a year |
| `could not read "…" in Exclusions.` | Unparseable exclusion entry |
| `too many dashes in "…"` | More than one range per comma-separated entry |
| `window is N days; max is 40000.` | Window exceeds `MAX_WINDOW_DAYS` |

A syntactically valid exclusion that falls outside the window is **not** an error — it is ignored.

---

## Testing

The logic was validated against an independent brute-force count built directly on the original `HijriDate` class:

- 15 targeted cases — full years, mid-year starts, multi-year windows, wrap-around ranges, whole-month names, dated vs undated exclusions, Gregorian input, reversed arguments, blank exclusions.
- 400 randomised configurations across Hijri years 1440–1460 with random windows and exclusion sets — **0 mismatches**.
- Round-trip check of `toAJD` ⇄ `fromAJD` over 21,616 consecutive dates (1400–1460 AH), which is what surfaced the day-zero boundary bug.
- Year lengths cross-checked against the Kabisa table for the same span — 0 mismatches.

---

## Repository layout

```
├── PeriodCount.gs      # the Apps Script file — this is what you paste
├── hijri_date.js       # original Hijri date maths (browser)
├── hijri_calendar.js   # original month-grid builder (browser, needs Lazy.js)
└── README.md
```

`PeriodCount.gs` is standalone — it embeds its own copy of the date maths and does not require the other two files or Lazy.js.

---

## Credits

Date arithmetic ported from `hijri_date.js`, with the month-grid conventions of `hijri_calendar.js` preserved.

## License

MIT — see `LICENSE`.
