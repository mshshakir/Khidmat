/**
 * ------------------------------------------------------------------
 * 1. SETUP & TRIGGER MANAGEMENT
 * ------------------------------------------------------------------
 */

function initializeJadwalTriggers() {
  stopAllAutomation();

  ScriptApp.newTrigger("dailyKickoff")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
    
  console.log("Daily trigger set for 8 AM. Automation started.");
}

function stopAllAutomation() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  console.log("All triggers deleted. Automation stopped.");
}

function dailyKickoff() {
  processAndScheduleNext();
}

function processAndScheduleNext() {
  const now = new Date();
  const currentHour = now.getHours();

  // SAFETY CHECK: Only run between 8:00 (8) and 16:00 (4 PM)
  if (currentHour >= 16) {
    console.log("Past 4 PM. Stopping loop for today.");
    return;
  }

  // --- EXECUTE SYNC ---
  try {
    syncJadwalDaily();
    console.log(`Synced at ${now.toLocaleTimeString()}`);
  } catch (e) {
    console.error("Error in syncJadwalDaily:", e);
  }

  // --- SCHEDULE NEXT RUN (CHAINING) ---
  const intervalMinutes = 35;
  const nextRunMs = intervalMinutes * 60 * 1000; 
  const nextRunTime = new Date(now.getTime() + nextRunMs);

  if (nextRunTime.getHours() >= 16) {
    console.log("Next run would be past 4 PM. Loop finished for today.");
  } else {
    ScriptApp.newTrigger("processAndScheduleNext")
      .timeBased()
      .after(nextRunMs)
      .create();
    console.log(`Next run scheduled for approx ${nextRunTime.toLocaleTimeString()}`);
  }
}

/**
 * ------------------------------------------------------------------
 * 2. MAIN SYNC LOGIC (DYNAMIC B1 LINK + STATUS CHECK)
 * ------------------------------------------------------------------
 */

function syncJadwalDaily() {
  const ss = SpreadsheetApp.getActive();
  const SHEET_NAME = "Jadwal";
  const calendar = CalendarApp.getDefaultCalendar();

  // 1️⃣ FETCH DYNAMIC LINK FROM B1 (First Sheet)
  // We assume the user puts the link in B1 of the first visible sheet
  const setupSheet = ss.getSheets()[0]; 
  const JADWAL_URL = setupSheet.getRange("B1").getValue();

  // Check if B1 is empty or not a string
  if (!JADWAL_URL || typeof JADWAL_URL !== 'string' || !JADWAL_URL.includes("http")) {
    console.error("❌ Error: No valid URL found in Cell B1 of the first sheet.");
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const dateKey = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

  // Fetch HTML
  let html;
  try {
    html = UrlFetchApp.fetch(JADWAL_URL, { muteHttpExceptions: true }).getContentText();
  } catch (e) {
    console.error("❌ Error fetching URL from B1: " + e.message);
    return;
  }

  // Extract table
  const tableMatch = html.match(/<table[\s\S]*?<\/table>/i);
  if (!tableMatch) return;

  const rowsHtml = tableMatch[0].match(/<tr[\s\S]*?<\/tr>/gi);
  if (!rowsHtml) return;

  // Prepare "Jadwal" Sheet (Target)
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Date", "Period", "Darajah",
      "Subject", "Start Time", "End Time", "Type", "Status", "EventKey"
    ]);
  }

  // Load sheet rows for today
  const data = sheet.getDataRange().getValues().slice(1);
  const sheetMap = {}; 

  data.forEach(r => {
    // EventKey is at index 8
    if (!r[8]) return;
    sheetMap[r[8]] = true;
  });

  const rowsToWrite = [];

  // Parse Jadwal rows
  for (let i = 1; i < rowsHtml.length; i++) {

    const cols = [];
    const colRegex = /<td[\s\S]*?>([\s\S]*?)<\/td>/gi;
    let m;

    while ((m = colRegex.exec(rowsHtml[i])) !== null) {
      cols.push(
        m[1].replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim()
      );
    }

    if (cols.length < 5) continue;

    const period = cols[0];
    const darajah = cols[1];
    const subject = cols[2];
    const time = cols[3];
    const type = cols[4];
    const status = cols.length > 5 ? cols[5] : ""; 

    if (!time.includes("-")) continue;

    const [startTime, endTime] = time.split("-").map(t => t.trim());
    const eventKey = `${dateKey}|${period}|${subject}|${startTime}`;

    const startDate = combine(today, startTime);
    const endDate = combine(today, endTime);
    const title = `${subject} (${darajah})`;

    // 🔍 Check calendar for this key
    const events = calendar.getEvents(startDate, endDate);
    let calendarEvent = null;

    for (let e = 0; e < events.length; e++) {
      if (events[e].getTag("eventKey") === eventKey) {
        calendarEvent = events[e];
        break;
      }
    }

    // 🛑 CANCELLATION CHECK 🛑
    if (status.toLowerCase().includes("cancel")) {
      if (calendarEvent) {
        calendarEvent.deleteEvent();
        console.log(`Deleted cancelled event: ${title}`);
      }
      continue; 
    }

    // 🛠 CASE 1: Sheet exists, calendar missing → recreate
    if (sheetMap[eventKey] && !calendarEvent) {
      const ev = calendar.createEvent(
        title,
        startDate,
        endDate,
        { description: `Period ${period}\nType: ${type}\nStatus: ${status}` }
      );
      ev.setTag("eventKey", eventKey);
      continue;
    }

    // 🆕 CASE 2: Completely new → create both
    if (!sheetMap[eventKey] && !calendarEvent) {
      const ev = calendar.createEvent(
        title,
        startDate,
        endDate,
        { description: `Period ${period}\nType: ${type}\nStatus: ${status}` }
      );
      ev.setTag("eventKey", eventKey);

      rowsToWrite.push([
        new Date(today),
        period,
        darajah,
        subject,
        startTime,
        endTime,
        type,
        status,
        eventKey
      ]);
    }
  }

  if (rowsToWrite.length > 0) {
    sheet.getRange(
      sheet.getLastRow() + 1,
      1,
      rowsToWrite.length,
      rowsToWrite[0].length
    ).setValues(rowsToWrite);
  }
}

function combine(date, timeStr) {
  const [h, m] = timeStr.split(":").map(Number);
  const d = new Date(date);
  d.setHours(h, m, 0, 0);
  return d;
}
