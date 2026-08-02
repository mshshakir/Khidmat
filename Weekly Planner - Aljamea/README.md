 # Weekly Planner → Google Calendar

A Google Sheet that keeps a dated weekly timetable, fills it with your Jadwal periods each morning, and turns every filled cell into an event on your Google Calendar.

**Template sheet:** https://docs.google.com/spreadsheets/d/17dTF57_MTLJ8i91mG7Uz5V2-LmL7FQIliFcRQMg7N34/edit

---

## Quick start (5 minutes)

1. **Make your own copy.** Open the template link above → **File → Make a copy**. The script travels with the copy. Do not work in the template itself; you almost certainly do not have edit rights to it, and even if you did, everyone would share one calendar.

2. **Reload the copy once.** A menu called **📅 Jadwal Grid** appears to the right of *Help*. If it is not there, wait ten seconds and refresh the page.

3. **Approve the script.** Click **📅 Jadwal Grid → Diagnose**. Google will ask for permission the first time:
   - *Authorization required* → **Continue**
   - Pick your Google account (use the one whose calendar you want)
   - *Google hasn't verified this app* → **Advanced** → **Go to … (unsafe)**
   - **Allow**

   This is normal for a script you copied yourself. You are the owner and the only user. What it asks for and why:

   | Permission | Why |
   |---|---|
   | See and manage this spreadsheet | Read and write the grid |
   | See and manage your calendars | Create the events |
   | Connect to an external service | Fetch your Jadwal page |
   | Send email as you | The "roll over your week" reminder |
   | Run when you are not present | The daily 6 AM fill |

4. **Set your Jadwal URL.** Log in to the Jamea portal, open your own *MyJadwal* page, and copy the whole address from the browser's address bar. Then in the sheet: **📅 Jadwal Grid → Set Jadwal URL**, paste, **OK**. The URL ends with a personal `ID=…` — this is what makes the timetable yours rather than someone else's.

5. **Pin the week.** **📅 Jadwal Grid → Set this week.** Each day header now reads e.g. `Monday / 03 Aug`.

6. **Turn on live syncing.** **📅 Jadwal Grid → Turn on sync-on-edit.** From now on, anything you type into a grid cell becomes a calendar event within a second or two, and a small toast in the bottom-right corner confirms it.

7. **Decide about Jadwal periods.** By default your *typed* entries go to the calendar but scraped *Jadwal* periods do not. To send those too: **📅 Jadwal Grid → Jadwal → Calendar: OFF (click to toggle)**. Confirm. Reload the sheet and the label reads **ON**.

8. **Fill it.** **📅 Jadwal Grid → Do both.** Today's and tomorrow's periods land in the grid; everything filled goes to your calendar.

That's the setup. After this the only thing you must do by hand is **Roll over week** each week.

---

## How it thinks

Three ideas explain nearly all the behaviour.

**The grid is dated, not generic.** The Sunday column is not "Sundays in general", it is one specific Sunday whose date is printed in the header. An event is always created on the date in the header — never on "the next Monday" or anything inferred.

**Every cell has an owner.** A cell written by the Jadwal scraper is *Jadwal-owned*; a cell you typed is *yours*. Ownership is recorded in a hidden sheet called `_GridState`, and can be recovered from the content itself if that record is lost (scraped blocks always begin `P3 · …`). Ownership decides which switch applies and stops the scraper from overwriting your notes.

**Weeks end deliberately.** The script never silently advances the dates. When the real world moves past the week printed in your grid, the daily fill *stops writing* rather than backdate your classes, and nags you until you roll over. Rolling over archives the week, wipes the grid, and moves the dates forward.

---

## Menu reference

### Turn on sync-on-edit
Installs the two triggers that make everything automatic:

- **on edit** — you type in a cell, the event appears; you clear the cell, the event is deleted; you change the text, the event is renamed.
- **daily at 6:00 AM (Asia/Kolkata)** — fetch the Jadwal page, fill today and tomorrow, then a catch-up calendar pass for anything the on-edit trigger missed.

Safe to run again; it removes its old triggers before adding new ones. Run it again if you ever restore an older copy of the script.

### Set Jadwal URL
Stores your personal MyJadwal link on this spreadsheet, so nobody has to open the code editor. Must start with `http://` or `https://`.

- Leave the box blank and press OK → keeps the current URL.
- Type `default` → restores the URL built into the script.

### Jadwal → Calendar: ON / OFF *(click to toggle)*
Controls whether **scraped periods** become calendar events. Your own typed entries are always synced and are not affected by this switch.

- **Turning ON** immediately creates events for every Jadwal cell currently in the grid.
- **Turning OFF** immediately deletes those events from your calendar. The periods stay in the grid; only the calendar copies go.

The menu label shows the current state, but it is built when the sheet opens — after toggling, reload the page to see the label change. The confirmation dialog always tells you what actually happened.

Event titles use the subject, e.g. *Fiqh*, with `P3 · Darajah` kept in the event description.

### Set this week
Pins the grid to the current Sunday–Saturday week and writes the date under each day name. Use it on first setup, or any time you want to jump the grid back to today's week without archiving.

### Catch-up sync to Calendar
Walks every filled cell in the grid and makes the calendar agree with it: creates what is missing, updates what changed, deletes events whose cells are now empty. Use it after pasting a block of entries, or if you edited cells while sync-on-edit was off.

### Fill grid from Jadwal
Fetches your Jadwal page now and writes **today's and tomorrow's** periods into their matching time slots. Cancelled periods are left blank on purpose. Periods whose time does not match any row in the grid are listed in a note on the day header (hover the little black triangle).

### Do both
*Fill grid from Jadwal* followed by *Catch-up sync to Calendar*, with a stale-week check first. This is exactly what the 6 AM trigger runs.

### Roll over week
The end-of-week ritual. It:

1. copies every filled cell into the **Archive** sheet — `Date | Day | Start | End | Source | Entry | Archived at`
2. clears the grid, leaving the `HH:MM - HH:MM` labels in place
3. moves every date forward one week (or straight to the real current week if you are further behind)

**Calendar events from the finished week are kept.** That week happened; it stays in your calendar as history. Only the grid is emptied.

You are shown a count and asked to confirm before anything is touched.

### Repair grid
For grids damaged by an older version of the script, which could append a period to a cell instead of replacing it. It removes repeated blocks and re-registers scraper-looking cells as Jadwal-owned. Harmless to run at any time.

### Diagnose
A read-only report: which sheet is being used as the grid, your Jadwal URL, whether Jadwal→Calendar is on, the grid week versus the real week, how many time slots and filled cells each day has, how many slots have events, the target calendar, archive size, and whether the triggers are installed. **Start here whenever something looks wrong** — paste its output when asking for help.

---

## The weekly routine

| When | What |
|---|---|
| Every morning, 6 AM | Automatic: Jadwal fill + calendar catch-up |
| As you plan | Type into any cell → event appears immediately |
| Start of a new week | **Roll over week** (the only manual step) |

If you forget to roll over, the sheet reminds you: a toast when you open it, and one email a day. The Jadwal fill deliberately pauses until you do.

---

## Troubleshooting

**The 📅 Jadwal Grid menu is missing.** Refresh the page. If it is still missing, you are probably looking at the template rather than your copy, or the copy was made without the script — copy it again with *File → Make a copy*.

**"Skipped Monday (today): grid column is 2026-07-27 but that day is 2026-08-03."** Your grid is dated to an old week. Run **Roll over week** (archives it first) or **Set this week** (does not archive).

**Nothing appears in my calendar.**
- Run **Diagnose** — check *Sync on edit: ON* and the *Default calendar* line.
- If only the Jadwal periods are missing, the **Jadwal → Calendar** toggle is OFF.
- Events go to your Google account's **default** calendar, not to a secondary one you may have made.

**"Jadwal fetch failed with HTTP 302 / 401 / 500."** The URL is wrong, expired, or the portal is asking for a login the script cannot perform. Open the URL in a private browser window: if it does not show your timetable without logging in, the script cannot read it either. Re-copy the link with **Set Jadwal URL**.

**A period appears two or three times in one cell.** Run **Repair grid**.

**I typed in a Jadwal cell.** That is allowed. With the toggle ON your edit updates that period's event; with it OFF the edit is kept in the grid only. The next Jadwal fill for that day will overwrite the cell with fresh data.

**Times are an hour out.** The script runs on `Asia/Kolkata`. If you are elsewhere, change `TIMEZONE` at the top of the script (**Extensions → Apps Script**) and also set the same zone in **File → Settings → Time zone**.

**I want to start clean.** Delete the rows (not the header) of the hidden `_GridState` sheet — right-click any sheet tab → *Show all sheets* to reach it — then run **Set this week**. This orphans any events already created; delete those from Calendar by hand.

---

## If you are building your own grid

The script finds the grid by shape, not by name, so you can restyle it freely as long as:

- **Column A** holds the time axis, each row exactly `HH:MM - HH:MM` (e.g. `09:25 - 10:15`).
- **One header row** in the first 12 rows contains the day names `Sunday` … `Saturday`, one per column. The date is added on a second line by the script.
- A lesson spanning several rows is a **merged cell** — its event runs from the first row's start to the last row's end.
- Cell text is free-form. The first line becomes the event title, the rest becomes the description.

Two sheets are managed for you and should not be edited by hand: `_GridState` (hidden bookkeeping) and `Archive` (the flat log of past weeks).
