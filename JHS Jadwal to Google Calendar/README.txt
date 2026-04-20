This is a comprehensive guide designed for a **GitHub README.md**. It covers everything from the initial setup to how the update system works.

---

# 📅 JHS Jadwal to Google Calendar Sync

This script automatically syncs your "My Jadwal" portal with your Google Calendar. It includes automatic bug-fix notifications and smart date validation to ensure your calendar is always accurate.

## 🚀 Initial Setup

Follow these steps to get started. You only need to do this **once**.

### 1. Prepare your Spreadsheet
* Create a new **Google Sheet**.
* Rename the first tab (at the bottom) to `Jadwal`.
* In cell **B1** of the `Jadwal` sheet, paste your unique **Jadwal URL** (the link you use to view your schedule).

### 2. Add the Script
* In your Google Sheet, go to **Extensions** > **Apps Script**.
* Delete any code currently in the editor window.
* Copy the entire code from **`sync.gs`** in this repository.
* Paste the code into the Apps Script editor.
* Click the **Disk Icon (Save)** and name the project `JHS Jadwal Sync`.

### 3. Initialize & Authorize
This is the most important step to start the automation.
* In the toolbar at the top of the editor, find the dropdown menu and select **`initializeJadwalTriggers`**.
* Click **Run**.
* A "Permissions Required" box will appear. Click **Review Permissions**.
* Select your Google Account.
* You will see a "Google hasn't verified this app" screen. 
    * Click **Advanced**.
    * Click **Go to JHS Jadwal Sync (unsafe)** at the bottom.
* Click **Allow**.
* Once the execution log says `Automation Initialized!`, you can close the tab.

---

## 🛠️ How to Use

### Manual Sync
If you want to sync your calendar immediately:
1. Go back to your Google Sheet.
2. **Refresh the page** in your browser.
3. You will see a new menu at the top called **📅 Jadwal**.
4. Click **📅 Jadwal** > **Sync Now**.

### Automatic Sync
The script is programmed to run automatically:
* It starts every morning at **8:00 AM**.
* It checks for updates every **35 minutes** throughout the day.
* It stops checking after **6:00 PM** to save energy.

---

## 🔄 How Updates Work

I periodically fix bugs or add new features. I have built an "Auto-Notification" system into this script.

1. **The Notification:** If I release a new version on GitHub, a popup will appear in your Google Sheet the next time you open it, saying: *"Update Available (Version X.X.X)"*.
2. **How to Update:** * Go back to this GitHub repository.
    * Copy the new code.
    * Go to **Extensions** > **Apps Script** in your sheet.
    * Replace the old code with the new code and **Save**.
    * *No need to re-initialize unless the update instructions specifically say so!*

---

## ⚠️ Troubleshooting

| Issue | Solution |
| :--- | :--- |
| **No "📅 Jadwal" Menu** | Refresh your browser tab and wait 5 seconds. |
| **Wrong Date on Calendar** | The script checks the website's "Day Name." If the website hasn't updated to the current day yet, the script skips syncing to prevent putting the wrong data on your calendar. |
| **Update Check Skipped** | This happens if you haven't authorized the script to talk to GitHub yet. Run `checkForUpdates` manually in the script editor to trigger the permission prompt. |

---
*Created by [Your Name/mshshakir]*
