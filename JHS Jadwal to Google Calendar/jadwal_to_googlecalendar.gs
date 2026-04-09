/**
 * ------------------------------------------------------------------
 * CONFIGURATION
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

  let sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Period", "Darajah", "Subject", "Start Time", "End Time", "Type", "Status", "EventKey"]);
    sheet.getRange("I:I").setNumberFormat("@"); // Force Column I to be Plain Text
  }

  const lastRow = sheet.getLastRow();
  const existingKeys = lastRow > 1 
    ? sheet.getRange(2, 9, lastRow - 1, 1).getValues().flat().map(k => String(k).trim()) 
    : [];
  
  const rowsToWrite = [];

  if (todayTable) processTableRows(todayTable, dateTodayStr, calendar, existingKeys, rowsToWrite);
  if (tomorrowTable) processTableRows(tomorrowTable, dateTomorrowStr, calendar, existingKeys, rowsToWrite);

  if (rowsToWrite.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
    console.log(`Added ${rowsToWrite.length} rows.`);
  }
}

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
    
    // THE KEY: Date + Period + Subject (first 15 chars) + Time
    const eventKey = `${dateStr}|${period}|${subject.substring(0, 15)}|${startT}`;
    
    // SHEET UPDATE
    if (!existingKeys.includes(eventKey)) {
      rowsToWrite.push([dateStr, period, darajah, subject, startT, endT, type, status, eventKey]);
      existingKeys.push(eventKey);
    }

    // CALENDAR UPDATE
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

function getTableById(html, id) {
  const regex = new RegExp(`<table[^>]*id="${id}"[\\s\\S]*?<\\/table>`, "i");
  const match = html.match(regex);
  return match ? match[0] : null;
}
