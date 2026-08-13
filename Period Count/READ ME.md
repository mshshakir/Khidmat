# PeriodCount — Hijri period counter for Google Sheets

A Google Apps Script custom function that counts how many times given weekdays occur between two Hijri dates, skipping holidays and breaks you define in Hijri terms.

**Its main use is working out how many periods a subject actually gets in an academic year**, once miqaat and other custom Hijri dates are taken out. Set the term dates once, list the miqaat once, and every subject's period count follows from its weekly timetable.

*"Physics meets on Mon, Tue and Wed. How many periods is that this academic year, given that Ashara Mubaraka and 20 Rabi al-Aakhar are off?"*

```
=PeriodCount(schedule, hijriExclusions, gregorianExclusions, start, end, [buffer], [dateSystem])
```

```
=PeriodCount("Mon,Tue,Wed", "[1/1-15/1,20/4]", "", "1/2/1447", "29/12/1447")
```

→ `138`

---

## Contents

- [Install](#install)
- [Syntax](#syntax)
- [Exclusion syntax](#exclusion-syntax)
- [Double periods](#double-periods)
- [Buffer](#buffer)
- [Examples](#examples)
- [Helper functions](#helper-functions) — `PeriodCountGross`, `PeriodDayCount`, `PeriodDates`, `HijriToGregorian`, `GregorianToHijri`
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
=PeriodCount(schedule, hijriExclusions, gregorianExclusions, start, end, [buffer], [dateSystem])
```

| # | Argument | Required | Description |
|---|---|---|---|
| 1 | `schedule` | yes | Days with periods, e.g. `"Mon,Tue,Wed"`. Supports double periods and day ranges — see below |
| 2 | `hijriExclusions` | yes (may be blank) | **Hijri** dates and ranges to skip — your miqaat. Pass `""` for none |
| 3 | `gregorianExclusions` | yes (may be blank) | **Gregorian** dates and ranges to skip. Pass `""` for none |
| 4 | `start` | yes | First day of the window, **inclusive** |
| 5 | `end` | yes | Last day of the window, **inclusive** |
| 6 | `buffer` | no | Periods to hold back for emergencies — a number (`5`) or a percentage (`"10%"`) |
| 7 | `dateSystem` | no | How to read *text* `start`/`end`: `"hijri"` (default) or `"gregorian"` |

The two exclusion lists sit side by side so a formula reads left to right: *what days, what to skip, over what window.* Both are positional — if you only use miqaat, pass `""` for the Gregorian list.

### Double periods

If a subject meets twice on a given day, mark it in the schedule. Case, spacing and short forms all work — `THU`, `Thu`, `Thurs` and `thursday` are the same day.

| Written | Meaning |
|---|---|
| `Mon` | 1 period every Monday |
| `Mon+1` | **2** periods every Monday (`+N` adds N extra) |
| `Mon+2` | 3 periods every Monday |
| `Mon x2` | 2 periods (`xN` sets the count outright; `*` and `×` also work) |
| `Mon-Thu` | Mon, Tue, Wed and Thu, 1 period each |
| `Mon-Wed+1` | Mon, Tue and Wed, 2 periods each |
| `Thu,Thu` | Summed, so 2 periods |

`"Tue,Wed,Thu+1,Fri,Sat"` means five teaching days, six periods a week.

### Buffer

`buffer` takes periods off the total, for exam days, emergencies or slippage.

| Written | Effect on a count of 155 |
|---|---|
| *(blank)* or `0` | `155` |
| `5` | `150` |
| `"10%"` | `139` — 10% off, rounded down |
| a percent-formatted cell (`0.1`) | `139` |
| `-5` | `160` — a negative buffer adds |

The result never drops below zero. Use `PeriodCountGross()` alongside `PeriodCount()` to show what the buffer costs.

**Returns:** a plain number.

### Dates in, dates out

- **Text** dates (`"1/2/1447"`) are read as **Hijri** `d/m/yyyy` by default.
- **Real date cells** are always read as **Gregorian** and converted internally.
- Month names work anywhere a date is accepted: `"1 Moharram 1447"`, `"20 Rabi II 1447"`.
- Gregorian entries in the exclusion lists are also **d/m**, not m/d. `25/12` is Christmas, not 12 May.
- If `start` is later than `end`, you get an error rather than a silent guess. The usual cause is an academic year crossing the Hijri new year — `1/10/1448` to `29/6/1448` should be `29/6/1449`. The error message says so.

Cell references work everywhere, so a timetable can be driven entirely from cells:

```
=PeriodCount(B2, $C$2, $D$2, $A$2, $B$2)
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

**Dates that do not exist are skipped, not shifted.** A single `30/12` in a non-Kabisa year (Zilhaj has 29 days) simply does not occur that year, so nothing is excluded. Inside a *range*, an out-of-reach endpoint clamps to the last real day of the month rather than spilling into the next.

**Two calendars, two lists.** Argument 2 is read as Hijri, argument 3 as Gregorian. Miqaat go in the Hijri list because they recur on fixed Hijri dates; national holidays, exam boards and term breaks usually go in the Gregorian list because they recur on fixed Gregorian dates.

```
=PeriodCount("Tue,Wed,Thu", "[1/1-15/1,20/4]", "[25/12,15/8]", "1/10/1448", "29/6/1449")
```

Gregorian entries take the same forms — `25/12`, `25/12/2027`, `24/12-31/12`, `December`, `1/1/2028-5/1/2028` — and undated ones repeat every Gregorian year the window touches. **They are d/m, not m/d.** A single date that does not exist in a given year (`29/2` outside a leap year) is skipped for that year.

You can also mix both in one cell with a `G:` tag, which applies to the rest of that cell until an `H:` tag switches back:

```
[1/1-15/1, 20/4, G:25/12, G:1/1, H:Ramadaan]
```

The `G:` tag is mainly useful when a single cell has to carry both kinds of date. With two separate arguments you rarely need it.

**Range separators.** Use `-` between two dates. Also accepted: `–`, `—`, `..`, `to`, `thru`, `through`, `till`, `until`. Hyphens inside month names (`Rabi al-Awwal`) are protected, so they don't split a range by accident.

---

## Examples

All figures below are actual output.

| Formula | Result |
|---|---|
| `=PeriodCount("Mon,Tue,Wed","","","1/1/1447","29/12/1447")` | `150` |
| `=PeriodCount("Mon,Tue,Wed","[1/1-15/1,20/4]","","1/1/1447","29/12/1447")` | `144` |
| `=PeriodCount("Mon,Tue,Wed","[1/1-15/1,20/4]","","1/2/1447","29/12/1447")` | `138` |
| `=PeriodCount("Mon,Tue,Wed","[1/1-15/1]","","1/1/1446","29/12/1448")` | `437` |
| `=PeriodCount("Mon,Tue,Wed","[1/1/1447-15/1/1447]","","1/1/1446","29/12/1448")` | `450` |
| `=PeriodCount("Mon,Tue,Wed,Thu","[25/12-5/1]","","1/1/1446","29/12/1448")` | `589` |
| `=PeriodCount("Mon,Wed","Ramadaan","","1/1/1447","29/12/1447")` | `91` |
| `=PeriodCount("Mon,Tue","Rabi al-Awwal","","1/1/1447","29/12/1447")` | `91` |
| `=PeriodCount("Mon,Tue,Wed","1 Moharram - 15 Moharram","","1/1/1447","29/12/1447")` | `144` |
| `=PeriodCount("Mon","","","1/1/1447","29/12/1447")` | `50` |

Note rows 4 and 5: identical windows, identical dates, but the undated exclusion repeats across all three years (`437`) while the dated one fires once (`450`).

### A worked academic year

Put the year's boundaries and the miqaat list in fixed cells, then drag one formula down the subject column.

| | A | B | C | D |
|---|---|---|---|---|
| **1** | Year start | Year end | Miqaat (Hijri) | Holidays (Gregorian) |
| **2** | `1/10/1448` | `29/6/1449` | `[1/1-15/1,20/4]` | `[25/12,15/8]` |
| **3** | **Subject** | **Schedule** | **Periods** | **Usable** |
| **4** | Physics | `Tue,Wed,Thu+1` | `=PeriodCountGross(B4,$C$2,$D$2,$A$2,$B$2)` | `=PeriodCount(B4,$C$2,$D$2,$A$2,$B$2,"10%")` |
| **5** | Chemistry | `Tue,Thu` | `=PeriodCountGross(B5,$C$2,$D$2,$A$2,$B$2)` | `=PeriodCount(B5,$C$2,$D$2,$A$2,$B$2,"10%")` |
| **6** | Maths | `Mon-Wed` | `=PeriodCountGross(B6,$C$2,$D$2,$A$2,$B$2)` | `=PeriodCount(B6,$C$2,$D$2,$A$2,$B$2,"10%")` |

Column C is what the calendar gives you; column D is what you'd actually plan against, holding 10% back. Physics has a double period on Thursday, so it earns four periods a week from three teaching days.

Because `$C$2` and `$D$2` are shared, adding one miqaat re-costs every subject at once. And because undated entries repeat annually, the same sheet keeps working when you roll the year forward.

Two things this makes easy:

- **Syllabus planning** — compare periods available against periods required per subject, before the year starts.
- **Miqaat impact** — delete one entry from the miqaat cell and the change in the totals is exactly what that occasion costs each subject.

> **Watch the end year.** An academic year that starts in Shawwal ends in the *next* Hijri year. `1/10/1448` to `29/6/1448` is backwards; you want `1/10/1448` to `29/6/1449`. The function raises an error rather than guessing, and tells you which year to use.

---

## Helper functions

The script installs six functions in total. `PeriodCount` is the one you'll use most; the rest exist to check it, break it down, or convert dates on their own.

| Function | Returns | Use it for |
|---|---|---|
| `PeriodCount` | number | The headline figure — periods after exclusions and buffer |
| `PeriodCountGross` | number | The same count *before* the buffer |
| `PeriodDayCount` | number | Teaching **days**, ignoring double periods |
| `PeriodDates` | table | Every day counted, for checking your setup |
| `HijriToGregorian` | date | Convert a single Hijri date |
| `GregorianToHijri` | text | Convert a single Gregorian date |

The first four share the same argument list, so a formula can be copied between them by changing only the name.

---

### `PeriodCountGross(schedule, hijriExclusions, gregorianExclusions, start, end, [dateSystem])`

Identical to `PeriodCount` but with no `buffer` argument — it always returns the full count. Put it in the column beside `PeriodCount` so the difference between the two shows what the buffer is holding back.

```
=PeriodCountGross("Tue,Wed,Thu+1,Fri,Sat", miqaat, "", "1/10/1448", "29/6/1449")   → 186
=PeriodCount(     "Tue,Wed,Thu+1,Fri,Sat", miqaat, "", "1/10/1448", "29/6/1449", "10%")   → 167
```

---

### `PeriodDayCount(schedule, hijriExclusions, gregorianExclusions, start, end, [dateSystem])`

Counts the **days** a subject meets rather than the periods it gets. Double periods are ignored, so each qualifying day counts once.

```
=PeriodDayCount("Tue,Wed,Thu+1,Fri,Sat", miqaat, "", "1/10/1448", "29/6/1449")   → 155
=PeriodCount(   "Tue,Wed,Thu+1,Fri,Sat", miqaat, "", "1/10/1448", "29/6/1449")   → 186
```

155 days in the classroom, 186 periods taught, because Thursday is a double. Useful for attendance registers, room bookings, or anything counted per visit rather than per period.

---

### `PeriodDates(schedule, hijriExclusions, gregorianExclusions, start, end, [dateSystem])`

Spills a table of every day counted. **This is the function to reach for whenever a number looks wrong** — it shows exactly which days were included, so a missing miqaat or a mis-typed weekday is visible at a glance.

```
=PeriodDates("Tue,Thu+1", "Ramadaan", "", "1/10/1448", "20/10/1448")
```

| Hijri | Month | Weekday | Gregorian | Periods |
|---|---|---|---|---|
| 2/10/1448 | Shawwal | Tue | 09/03/2027 | 1 |
| 4/10/1448 | Shawwal | Thu | 11/03/2027 | 2 |
| 9/10/1448 | Shawwal | Tue | 16/03/2027 | 1 |
| 11/10/1448 | Shawwal | Thu | 18/03/2027 | 2 |
| 16/10/1448 | Shawwal | Tue | 23/03/2027 | 1 |
| 18/10/1448 | Shawwal | Thu | 25/03/2027 | 2 |

Notes:

- The `Periods` column reflects double periods, so it sums to the `PeriodCount` total. The row count equals `PeriodDayCount`.
- The `Gregorian` column contains real date values — format the column as a date if it shows as a serial number.
- Excluded days are simply absent; the table lists what counted, not what was skipped.
- `buffer` is not an argument here, since a buffer is a deduction from a total rather than a specific day.
- If nothing matches, it returns a single cell reading `No periods found`.
- Needs empty cells below and to the right to spill into, or Sheets shows `#REF!`.

Because it spills, keep it on a scratch sheet rather than inside your main timetable grid.

---

### `HijriToGregorian(hijri)`

Converts one Hijri date to a real Gregorian date value.

| Formula | Result |
|---|---|
| `=HijriToGregorian("1/1/1447")` | 26 June 2025 |
| `=HijriToGregorian("1 Moharram 1447")` | 26 June 2025 |
| `=HijriToGregorian("1/10/1448")` | 8 March 2027 |
| `=HijriToGregorian("29/6/1449")` | 28 November 2027 |
| `=HijriToGregorian("1 Ramadaan 1448")` | 6 February 2027 |

- Accepts `d/m/yyyy` or `d MonthName yyyy`, with the same month names and short forms as everywhere else.
- The **year is required** — `"1/1"` has no single answer and raises an error.
- Returns a genuine date value, not text, so it sorts correctly and works inside `TEXT()`, `WEEKDAY()` and date arithmetic. Format the cell as a date if it appears as a number.

Handy for printing a term calendar: put Hijri dates in one column and this formula beside them.

---

### `GregorianToHijri(gregorian)`

The reverse. Takes a real date cell and returns the Hijri date as text.

| Formula | Result |
|---|---|
| `=GregorianToHijri(DATE(2025,6,26))` | `1 Moharram 1447` |
| `=GregorianToHijri(DATE(2026,1,1))` | `13 Rajab 1447` |
| `=GregorianToHijri(DATE(2027,3,8))` | `1 Shawwal 1448` |
| `=GregorianToHijri(TODAY())` | today's Hijri date |

- The input must be a **real date value** — a date cell, `DATE()`, or `TODAY()`. Text like `"26/6/2025"` raises an error, because there is no reliable way to tell `3/4` apart as March 4th or April 3rd.
- Returns text in `d MonthName yyyy` form using the short month names. To get the parts separately, wrap in `SPLIT()`.
- Both converters honour `HIJRI_DAY_ADJUSTMENT`, so they always agree with whatever `PeriodCount` is counting.

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
| `Start (1 Shawwal 1448) is after End (29 Jumada II 1448)…` | End is before start. If the year crosses the Hijri new year, increment the end year |
| `"1/1/1448" has year 1448, which is not a Gregorian year…` | A Hijri date landed in the Gregorian list, or vice versa. The message says which list to move it to |
| `buffer "abc" is not a number or percentage.` | Buffer must be a number, a `"10%"` string, or blank |
| `End is empty. The argument order is…` | A formula written for an earlier version. Insert `""` for the Gregorian list before the start date |
| `"Monx" is not a weekday.` | Unreadable schedule token — check the `+N` / `xN` form |
| `could not read "…" in Days.` | Stray text in the schedule that isn't a day or a period count |

A syntactically valid exclusion that falls outside the window is **not** an error — it is ignored.

---

## Testing

The logic was validated against an independent brute-force count built directly on the original `HijriDate` class:

- 15 targeted cases — full years, mid-year starts, multi-year windows, wrap-around ranges, whole-month names, dated vs undated exclusions, Gregorian input, reversed arguments, blank exclusions.
- 400 randomised configurations across Hijri years 1440–1460 with random windows and exclusion sets — **0 mismatches**.
- A further 300 randomised cases with random double-period weights on random weekdays — **0 mismatches**.
- Every double-period spelling checked against a single canonical result: `THU+1`, `thu+1`, `Thurs+1`, `THURSDAY+1`, `Thu + 1`, `thu*2`, `THU X2`, `Thu2`, `Thu×2` and `  Thu  +  1  ` all agree.
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
