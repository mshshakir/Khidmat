/**
 * ------------------------------------------------------------------
 * CONFIGURATION & UPDATE SETTINGS
 * ------------------------------------------------------------------
 */
const TIMEZONE = "Asia/Kolkata"; 
const CURRENT_VERSION = "1.5.0"; // Increment this when you release a fix
const GITHUB_RAW_VERSION_URL = "https://raw.githubusercontent.com/mshshakir/Khidmat/refs/heads/main/JHS%20Jadwal%20to%20Google%20Calendar/version.txt";

/**
 * INSTALLATION & UI MENU
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // This creates the custom menu in the Google Sheet UI
  ui.createMenu('📅 Jadwal')
    .addItem('Sync Now', 'syncJadwalDaily')
    .addSeparator()
    .addItem('Check for Updates', 'checkForUpdates')
    .addToUi();

  // Runs the update check automatically when the sheet is opened
  checkForUpdates();
}

function initializeJadwalTriggers() {
  stopAllAutomation();
  
  // Primary daily trigger
  ScriptApp.newTrigger("dailyKickoff")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .nearMinute(0) 
    .inTimezone(TIMEZONE) 
    .create();
    
  console.log(`Daily trigger set for 8 AM (${TIMEZONE}).`);
  SpreadsheetApp.getActive().toast("Automation Initialized!", "Success");
}

function stopAllAutomation() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
}

/**
 * AUTOMATION LOGIC
 */
function dailyKickoff() {
  processAndScheduleNext();
}

function processAndScheduleNext() {
  const now = new Date();
  const currentHour = parseInt(Utilities.formatDate(now, TIMEZONE, "H"));
  
  // Stop processing after 6 PM
  if (currentHour >= 18) return;

  try { 
    syncJadwalDaily(); 
  } catch (e) { 
    console.error("Error during sync:", e); 
  }

  // Reschedule next run in 35 minutes
  ScriptApp.newTrigger("processAndScheduleNext")
    .timeBased()
    .after(35 * 60 * 1000)
    .create();
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
  
  // 1. Setup Dates
  const today = new Date();
  const tomorrow = new Date();
  tomorrow.setDate(today.getDate() + 1);

  const dateTodayStr = Utilities.formatDate(today, TIMEZONE, "yyyy-MM-dd");
  const dateTomorrowStr = Utilities.formatDate(tomorrow, TIMEZONE, "yyyy-MM-dd");
  
  const expectedDayToday = Utilities.formatDate(today, TIMEZONE, "EEEE");
  const expectedDayTomorrow = Utilities.formatDate(tomorrow, TIMEZONE, "EEEE");

  // 2. Extract and Validate Day Names from Web
  const webDayToday = getElementContentById(html, "litDayName");
  const webDayTomorrow = getElementContentById(html, "litNextDayName");

  const processToday = (webDayToday === expectedDayToday);
  const processTomorrow = (webDayTomorrow === expectedDayTomorrow);

  if (!processToday && !processTomorrow) {
    console.warn("Website days do not match current system dates. Skipping sync.");
    return;
  }

  // 3. Get Table HTML if validated
  const todayTable = processToday ? getTableById(html, "gvTodaysPeriods") : null;
  const tomorrowTable = processTomorrow ? getTableById(html, "gvNextPeriods") : null;

  // 4. Extract Keys for comparison
  let webKeys = [];
  if (todayTable) webKeys = webKeys.concat(extractKeysFromTable(todayTable, dateTodayStr));
  if (tomorrowTable) webKeys = webKeys.concat(extractKeysFromTable(tomorrowTable, dateTomorrowStr));

  let sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Period", "Darajah", "Subject", "Start Time", "End Time", "Type", "Status", "EventKey"]);
    sheet.getRange("I:I").setNumberFormat("@");
  }

  const lastRow = sheet.getLastRow();

  // --- PART 1: REMOVALS ---
  if (lastRow > 1) {
    const range = sheet.getRange(2, 1, lastRow - 1, 9);
    const values = range.getValues();
    let sheetUpdated = false;

    for (let i = 0; i < values.length; i++) {
      const rowDateRaw = values[i][0];
      const rowDate = (rowDateRaw instanceof Date) ? Utilities.formatDate(rowDateRaw, TIMEZONE, "yyyy-MM-dd") : String(rowDateRaw);
      const rowKey = String(values[i][8]).trim();
      const currentStatus = String(values[i][7]);

      if (((rowDate === dateTodayStr && processToday) || (rowDate === dateTomorrowStr && processTomorrow)) && 
          !currentStatus.includes("REMOVED") && !currentStatus.includes("CANCEL")) {
        
        if (!webKeys.includes(rowKey)) {
           values[i][7] = "REMOVED FROM WEB"; 
           sheetUpdated = true;
           deleteCalendarEventByKey(calendar, rowKey, rowDate);
        }
      }
    }
    if (sheetUpdated) range.setValues(values);
  }

  // --- PART 2: ADDITIONS ---
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
 * HELPERS: WEB PARSING
 */
function getElementContentById(html, id) {
  const regex = new RegExp(`<span[^>]*id="${id}"[^>]*>([\\s\\S]*?)<\\/span>`, "i");
  const match = html.match(regex);
  return match ? match[1].replace(/<[^>]+>/g, "").trim() : null;
}

function getTableById(html, id) {
  const regex = new RegExp(`<table[^>]*id="${id}"[\\s\\S]*?<\\/table>`, "i");
  const match = html.match(regex);
  return match ? match[0] : null;
}

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
 * HELPERS: CALENDAR & UPDATES
 */
function deleteCalendarEventByKey(calendar, key, dateStr) {
  const startOfDay = Utilities.parseDate(`${dateStr} 00:00`, TIMEZONE, "yyyy-MM-dd HH:mm");
  const endOfDay = Utilities.parseDate(`${dateStr} 23:59`, TIMEZONE, "yyyy-MM-dd HH:mm");
  const events = calendar.getEvents(startOfDay, endOfDay);
  const ev = events.find(e => e.getTag("eventKey") === key);
  if (ev) ev.deleteEvent();
}

function checkForUpdates() {
  try {
    const response = UrlFetchApp.fetch(GITHUB_RAW_VERSION_URL);
    const latestVersion = response.getContentText().trim();
    if (latestVersion !== CURRENT_VERSION) {
      const ui = SpreadsheetApp.getUi();
      ui.alert("Update Available", `A new version (${latestVersion}) is available on GitHub. Please update your code to fix bugs.`, ui.ButtonSet.OK);
    }
  } catch (e) {
    console.log("Update check skipped.");
  }
}
