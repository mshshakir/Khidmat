This Google Apps Script is designed to automate a schedule (Jadwal) synchronization. It scrapes an HTML table from a specific website URL, saves the data into a Google Sheet, and creates corresponding events in your Google Calendar.

It even features a **self-chaining loop** to check for updates every 35 minutes during work hours.

---

## ## 1. How the Code Works

### **A. The Automation Loop (Self-Chaining)**
Unlike a standard trigger that runs once an hour, this script uses "Chaining":
1.  **`initializeJadwalTriggers`**: Starts the cycle by setting a trigger to run every morning at 8:00 AM.
2.  **`processAndScheduleNext`**: This is the engine. It runs the sync, then looks at the clock. If it’s before 4:00 PM, it creates a **one-time trigger** to run itself again in 35 minutes. 
3.  **The Safety Valve**: Once the clock hits 4:00 PM, it stops creating new triggers, effectively "going to sleep" until the next morning.

### **B. The Sync Logic (`syncJadwalDaily`)**
1.  **Dynamic URL**: It looks at **Cell B1** of your very first sheet tab to find the website link.
2.  **Web Scraping**: It fetches the HTML from that link and uses "Regex" (Regular Expressions) to find a `<table>` and its rows (`<tr>`) and columns (`<td>`).
3.  **Google Sheet Logging**: It checks a sheet named "Jadwal." If an event (identified by a unique `eventKey`) isn't there, it adds it.
4.  **Calendar Integration**: 
    * It creates a Google Calendar event for each row.
    * **Smart Deletion**: If the status in the table contains the word "cancel," it automatically finds the event in your calendar and deletes it.
    * **Tags**: It hides a "fingerprint" (EventKey) inside the calendar event so it can recognize it later.

---

## ## 2. Instructions for Setup

Follow these steps to get the automation running:

### **Step 1: Prepare the Google Sheet**
1.  Open a new or existing Google Sheet.
2.  In the **first tab** (the leftmost one), type or paste your schedule URL into **Cell B1**. 
    * *Example: `https://example.com/daily-schedule`*

### **Step 2: Install the Script**
1.  In your Google Sheet, go to **Extensions** > **Apps Script**.
2.  Delete any code in the editor and paste the entire code block you provided.
3.  Click the **Save** icon (floppy disk) and name it "Jadwal Sync."

### **Step 3: Authorize & Run**
1.  In the toolbar at the top, select `initializeJadwalTriggers` from the function dropdown.
2.  Click **Run**.
3.  A "Permissions required" box will appear. Click **Review Permissions**.
    * Select your Google Account.
    * You will see a "Google hasn't verified this app" warning (this is normal for private scripts). Click **Advanced** > **Go to Jadwal Sync (unsafe)**.
    * Click **Allow**.

### **Step 4: Verification**
1.  Check the **Execution Log** at the bottom. You should see "Daily trigger set for 8 AM."
2.  To test the sync immediately without waiting for 8 AM, select the function `syncJadwalDaily` and click **Run**.
3.  Check your Google Calendar and the "Jadwal" sheet tab to see if the data populated.

---

## ## 3. Important Maintenance Notes
* **Time Window**: The script only syncs between **8:00 AM and 4:00 PM**. If you run it at 5:00 PM, it will log a message and stop.
* **URL Changes**: If the source website changes its URL, just update **Cell B1** in your sheet. The script will pick up the new link on its next 35-minute cycle.
* **Stopping the Script**: If you want to kill the automation entirely, run the `stopAllAutomation` function from the script editor.
