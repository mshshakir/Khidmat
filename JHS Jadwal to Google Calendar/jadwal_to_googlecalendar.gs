/**
 * ------------------------------------------------------------------
 * CONFIGURATION & AUTOMATION
 * ------------------------------------------------------------------
 */
const TIMEZONE = "Asia/Kolkata"; 

function initializeJadwalTriggers() {
  stopAllAutomation();
  ScriptApp.newTrigger("dailyKickoff")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .nearMinute(0) 
    .inTimezone(TIMEZONE) 
    .create();
  console.log(`Daily trigger set for 8 AM (${TIMEZONE}).`);
}

function stopAllAutomation() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
}

function dailyKickoff() {
  processAndScheduleNext();
}

function processAndScheduleNext() {
  const now = new Date();
  const currentHour = parseInt(Utilities.formatDate(now, TIMEZONE, "H"));
  if (currentHour >= 18) return;
  try { syncJadwalDaily(); } catch (e) { console.error("Error:", e); }
  ScriptApp.newTrigger("processAndScheduleNext").timeBased().after(35 * 60 * 1000).create();
}

/**
 * MAIN SYNC FUNCTION
 */
function syncJadwalDaily() {
  const ss = SpreadsheetApp.getActive();
  const SHEET_NAME = "Jadwal";
  const calendar = CalendarApp.getDefaultCalendar();
  const setupSheet = ss.getSheets()[0]; 
  const JADWAL_URL = setupSheet.getRange("B1").getValue();

  if (!JADWAL_URL || !JADWAL_URL.includes("http")) return;

  const html = UrlFetchApp.fetch(JADWAL_URL, { muteHttpExceptions: true }).getContentText();
  const today = new Date();
  const tomorrow = new Date();
  tomorrow.setDate(today.getDate() + 1);

  const dateTodayStr = Utilities.formatDate(today, TIMEZONE, "yyyy-MM-dd");
  const dateTomorrowStr = Utilities.formatDate(tomorrow, TIMEZONE, "yyyy-MM-dd");

  const todayTable = getTableById(html, "gvTodaysPeriods");
  const tomorrowTable = getTableById(html, "gvNextPeriods");

  // 1. Get current keys from Web for comparison
  let webKeys = [];
  if (todayTable) webKeys = webKeys.concat(extractKeysFromTable(todayTable, dateTodayStr));
  if (tomorrowTable) webKeys = webKeys.concat(extractKeysFromTable(tomorrowTable, dateTomorrowStr));

  let sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Period", "Darajah", "Subject", "Start Time", "End Time", "Type", "Status", "EventKey"]);
    sheet.getRange("I:I").setNumberFormat("@");
  }

  const lastRow = sheet.getLastRow();

  // --- PART 1: SCAN SHEET FOR REMOVALS ---
  if (lastRow > 1) {
    const range = sheet.getRange(2, 1, lastRow - 1, 9);
    const values = range.getValues();
    let sheetUpdated = false;

    for (let i = 0; i < values.length; i++) {
      const rowDateRaw = values[i][0];
      const rowDate = (rowDateRaw instanceof Date) ? Utilities.formatDate(rowDateRaw, TIMEZONE, "yyyy-MM-dd") : String(rowDateRaw);
      const rowKey = String(values[i][8]).trim();
      const currentStatus = String(values[i][7]);

      // Focus on Today/Tomorrow and ignore rows already marked Removed/Cancelled
      if ((rowDate === dateTodayStr || rowDate === dateTomorrowStr) && 
          !currentStatus.includes("REMOVED") && 
          !currentStatus.includes("CANCEL")) {
        
        // If the unique key in our sheet is MISSING from the web keys list
        if (!webKeys.includes(rowKey)) {
           console.log("Match not found on web. Deleting: " + rowKey);
           values[i][7] = "REMOVED FROM WEB"; 
           sheetUpdated = true;
           deleteCalendarEventByKey(calendar, rowKey, rowDate);
        }
      }
    }
    if (sheetUpdated) range.setValues(values);
  }

  // --- PART 2: ADD NEW ROWS & HANDLE WEB-MARKED CANCELLATIONS ---
  const existingKeys = lastRow > 1 
    ? sheet.getRange(2, 9, lastRow - 1, 1).getValues().flat().map(k => String(k).trim()) 
    : [];
  
  const rowsToWrite = [];
  if (todayTable) processTableRows(todayTable, dateTodayStr, calendar, existingKeys, rowsToWrite);
  if (tomorrowTable) processTableRows(tomorrowTable, dateTomorrowStr, calendar, existingKeys, rowsToWrite);

  if (rowsToWrite.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
  }
}

/**
 * HELPER: EXTRACT ALL KEYS FROM WEB TABLES
 */
function extractKeysFromTable(tableHtml, dateStr) {
  const keys = [];
  const rowRegex = /<tr[\s\S]*?>([\s\S]*?)<\/tr>/gi;
  let rowMatch;
  let isHeader = true;
  while ((rowMatch = rowRegex.exec(tableHtml)) !== null) {
    if (isHeader) { isHeader = false; continue; }
    const colRegex = /<td[\s\S]*?>([\s\S]*?)<\/td>/gi;
    let colMatch;
    const cols = [];
    while ((colMatch = colRegex.exec(rowMatch[1])) !== null) {
      cols.push(colMatch[1].replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim());
    }
    if (cols.length < 4) continue;
    const [period, , subject, time] = cols;
    if (!time.includes("-")) continue;
    const startT = time.split("-")[0].trim();
    keys.push(`${dateStr}|${period}|${subject.substring(0, 15)}|${startT}`);
  }
  return keys;
}

/**
 * HELPER: PROCESS WEB ROWS (ADD/CANCEL)
 */
function processTableRows(tableHtml, dateStr, calendar, existingKeys, rowsToWrite) {
  const rowRegex = /<tr[\s\S]*?>([\s\S]*?)<\/tr>/gi;
  let rowMatch;
  let isHeader = true;

  while ((rowMatch = rowRegex.exec(tableHtml)) !== null) {
    if (isHeader) { isHeader = false; continue; }

    const colRegex = /<td[\s\S]*?>([\s\S]*?)<\/td>/gi;
    let colMatch;
    const cols = [];
    while ((colMatch = colRegex.exec(rowMatch[1])) !== null) {
      cols.push(colMatch[1].replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim());
    }

    if (cols.length < 5) continue;
    const [period, darajah, subject, time, type] = cols;
    const status = cols[5] || "";
    if (!time.includes("-")) continue;

    const [startT, endT] = time.split("-").map(t => t.trim());
    const eventKey = `${dateStr}|${period}|${subject.substring(0, 15)}|${startT}`;
    
    if (!existingKeys.includes(eventKey)) {
      rowsToWrite.push([dateStr, period, darajah, subject, startT, endT, type, status, eventKey]);
      existingKeys.push(eventKey);
    }

    const startDate = Utilities.parseDate(`${dateStr} ${startT}`, TIMEZONE, "yyyy-MM-dd HH:mm");
    const endDate = Utilities.parseDate(`${dateStr} ${endT}`, TIMEZONE, "yyyy-MM-dd HH:mm");
    const events = calendar.getEvents(startDate, endDate);
    let calendarEvent = events.find(e => e.getTag("eventKey") === eventKey);

    if (status.toLowerCase().includes("cancel")) {
      if (calendarEvent) calendarEvent.deleteEvent();
    } else if (!calendarEvent) {
      const ev = calendar.createEvent(`${subject} (${darajah})`, startDate, endDate, {
        description: `Period: ${period}\nType: ${type}\nStatus: ${status}`
      });
      ev.setTag("eventKey", eventKey);
    }
  }
}

/**
 * HELPER: DELETE BY KEY
 */
function deleteCalendarEventByKey(calendar, key, dateStr) {
  const startOfDay = Utilities.parseDate(`${dateStr} 00:00`, TIMEZONE, "yyyy-MM-dd HH:mm");
  const endOfDay = Utilities.parseDate(`${dateStr} 23:59`, TIMEZONE, "yyyy-MM-dd HH:mm");
  
  const events = calendar.getEvents(startOfDay, endOfDay);
  const ev = events.find(e => e.getTag("eventKey") === key);
  
  if (ev) {
    ev.deleteEvent();
    console.log("✅ Successfully deleted from Calendar: " + key);
  } else {
    // Fallback: search by partial title in case tag failed
    const titlePart = key.split('|')[2];
    const fallbackEv = events.find(e => e.getTitle().includes(titlePart));
    if (fallbackEv) {
      fallbackEv.deleteEvent();
      console.log("✅ Deleted via Fallback match.");
    }
  }
}

/**
 * HELPER: EXTRACT TABLE
 */
function getTableById(html, id) {
  const regex = new RegExp(`<table[^>]*id="${id}"[\\s\\S]*?<\\/table>`, "i");
  const match = html.match(regex);
  return match ? match[0] : null;
}
