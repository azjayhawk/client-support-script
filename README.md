# client-support-script
Google Apps Script for managing monthly client docs
# Support Automation Script

This is a Google Apps Script project used to automate monthly rollover tasks for client support documentation.

## ✨ What It Does

This script:

- Identifies the previous month
- Iterates over all active clients in the **"Master Tracker"** Google Sheet
- Creates a support summary **Google Doc** for each client
- Logs those documents in the sheet (instead of emailing them)

## 🚀 How to Use (In Google Sheets)

1. Open the Google Sheet that this script is attached to
2. Go to **Extensions > Apps Script**
3. Run the function: `monthlyRolloverAndCreateDocs()`
4. Optionally enable the `Client Tools` menu to run it from the UI

> 🔒 You must have access to the connected Drive folders and spreadsheet for the script to function.

## 📁 Project Structure

| File                            | Purpose                                      |
|-------------------------------|----------------------------------------------|
| `monthlyRolloverAndCreateDocs.gs` | Main function to generate monthly docs      |
| `appsscript.json`              | Script manifest (controls file settings)     |
| `.clasp.json`                  | CLASP config (don’t upload this to GitHub)   |
| `README.md`                    | You're reading it! 😄                         |

## 💡 Powered By

- [Google Apps Script](https://developers.google.com/apps-script)
- [CLASP](https://developers.google.com/apps-script/guides/clasp) – command-line tool to sync code
- [GitHub](https://github.com) – version control & backup

---

## 🛠️ Developer Notes

- If you want to test without creating files, set `DRY_RUN = true` in the script.
- To update from your computer to Google Apps Script, run:

```bash
clasp push
