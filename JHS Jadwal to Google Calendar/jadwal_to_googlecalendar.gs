/**
 * ------------------------------------------------------------------
 * CONFIGURATION
 * ------------------------------------------------------------------
 */
const TIMEZONE = "Asia/Kolkata"; // Set to Mumbai/India Standard Time

/**
 * 1. SETUP & TRIGGER MANAGEMENT
 */
function initializeJadwalTriggers() {
  stopAllAutomation();

  ScriptApp.newTrigger("dailyKickoff")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .nearMinute(0) 
    .inTimezone(TIMEZONE) 
    .create();
    
  console.log(`Daily trigger set for 8 AM (${TIMEZONE}). Automation started.`);
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
  // Get current hour in Mumbai time
  const currentHour = parseInt(Utilities.formatDate(now, TIMEZONE, "H"));

  // SAFETY CHECK: Only run between 8:00 and 16:00 Mumbai Time
  if (currentHour >= 16) {
    console.log("Past 4 PM Mumbai time. Stopping loop for today.");
    return;
  }

  try {
    syncJadwalDaily();
    console.log(`Synced at ${Utilities.formatDate(now, TIMEZONE, "HH:mm:ss")}`);
  } catch (e) {
    console.error("Error in syncJadwalDaily:", e);
  }

  const intervalMinutes = 35;
  const nextRunMs = intervalMinutes * 60 * 1000; 
  const nextRunTime = new Date(now.getTime() + nextRunMs);
  const nextRunHour = parseInt(Utilities.formatDate(nextRunTime, TIMEZONE, "H"));

  if (nextRunHour >= 16) {
    console.log("Next run would be past 4 PM Mumbai time. Loop finished.");
  } else {
    ScriptApp.newTrigger("processAndScheduleNext")
      .timeBased()
      .after(nextRunMs)
      .create();
    console.log(`Next run scheduled for approx ${Utilities.formatDate(nextRunTime, TIMEZONE, "HH:mm")}`);
  }
}

/**
 * 2. MAIN SYNC LOGIC
 */
function syncJadwalDaily() {
  const ss = SpreadsheetApp.getActive();
  const SHEET_NAME = "Jadwal";
  const calendar = CalendarApp.getDefaultCalendar();

  const setupSheet = ss.getSheets()[0]; 
  const JADWAL_URL = setupSheet.getRange("B1").getValue();

  if (!JADWAL_URL || typeof JADWAL_URL !== 'string' || !JADWAL_URL.includes("http")) {
    console.error("❌ Error: No valid URL found in Cell B1.");
    return;
  }

  const today = new Date();
  const dateStr = Utilities.formatDate(today, TIMEZONE, "yyyy-MM-dd");

  let html;
  try {
    html = UrlFetchApp.fetch(JADWAL_URL, { muteHttpExceptions: true }).getContentText();
  } catch (e) {
    console.error("❌ Error fetching URL: " + e.message);
    return;
  }

  const tableMatch = html.match(/<table[\s\S]*?<\/table>/i);
  if (!tableMatch) return;

  const rowsHtml = tableMatch[0].match(/<tr[\s\S]*?<\/tr>/gi);
  if (!rowsHtml) return;

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Date", "Period", "Darajah", "Subject", "Start Time", "End Time", "Type", "Status", "EventKey"]);
  }

  const data = sheet.getDataRange().getValues().slice(1);
  const sheetMap = {}; 
  data.forEach(r => { if (r[8]) sheetMap[r[8]] = true; });

  const rowsToWrite = [];

  for (let i = 1; i < rowsHtml.length; i++) {
    const cols = [];
    const colRegex = /<td[\s\S]*?>([\s\S]*?)<\/td>/gi;
    let m;
    while ((m = colRegex.exec(rowsHtml[i])) !== null) {
      cols.push(m[1].replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim());
    }

    if (cols.length < 5) continue;

    const [period, darajah, subject, time, type] = cols;
    const status = cols.length > 5 ? cols[5] : ""; 

    if (!time.includes("-")) continue;

    const [startTime, endTime] = time.split("-").map(t => t.trim());
    const eventKey = `${dateStr}|${period}|${subject}|${startTime}`;

    // Convert parsed time to Mumbai-localized Date objects
    const startDate = parseTimeForZone(dateStr, startTime, TIMEZONE);
    const endDate = parseTimeForZone(dateStr, endTime, TIMEZONE);
    const title = `${subject} (${darajah})`;

    const events = calendar.getEvents(startDate, endDate);
    let calendarEvent = null;

    for (let e = 0; e < events.length; e++) {
      if (events[e].getTag("eventKey") === eventKey) {
        calendarEvent = events[e];
        break;
      }
    }

    // Cancellation logic
    if (status.toLowerCase().includes("cancel")) {
      if (calendarEvent) {
        calendarEvent.deleteEvent();
        console.log(`Deleted: ${title}`);
      }
      continue; 
    }

    // Event creation logic
    if (!calendarEvent) {
      const ev = calendar.createEvent(
        title,
        startDate,
        endDate,
        { description: `Period ${period}\nType: ${type}\nStatus: ${status}` }
      );
      ev.setTag("eventKey", eventKey);

      if (!sheetMap[eventKey]) {
        rowsToWrite.push([dateStr, period, darajah, subject, startTime, endTime, type, status, eventKey]);
      }
    }
  }

  if (rowsToWrite.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
  }
}

/**
 * Correctly parses a date and time string into a Date object for a specific timezone.
 */
function parseTimeForZone(dateStr, timeStr, tz) {
  // dateStr is "yyyy-MM-dd", timeStr is "HH:mm"
  // Constructing a string format that Utilities.parseDate understands
  const dateTimeStr = `${dateStr} ${timeStr}`;
  return Utilities.parseDate(dateTimeStr, tz, "yyyy-MM-dd HH:mm");
}
