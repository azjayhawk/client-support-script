Client Support Tracker URL - https://docs.google.com/spreadsheets/d/1TnQ_FSgTGRbz0KqzqLPKHJw7oQu9sC4hLCbTVXkxXhM/edit?gid=0#gid=0

Here’s a detailed End-of-Month Checklist based on your current monthlyRolloverAndCreateDocs() script and overall system:

⸻

✅ End-of-Month Checklist

Purpose: Prepare your tracker, generate client support docs, and reset formulas to begin the new month.

⸻

🔹 BEFORE Running the Script

* Confirm All Hours Are Entered  
  * Double-check that all time entries for the month are logged in the Time Entry tab.  
  * Ensure client time is accurately reflected for each client.   
* Update Block Hours  
* In the Master Tracker, enter or update values in Column E (“Block Hours Available”) as needed.   
* Verify Client Statuses	In the Client Directory, confirm that:   
  * Active clients are marked as "Active" in Column D.   
  * Clients who should not receive reports are marked "Inactive" or "Transitioning".   
* Insert Any New Clients   
  * Use the menu: Client Tools \> ➕ Add Client and Sync to Master Tracker.   
  * This ensures new clients are added to both sheets and properly formatted.   
* Run Formula Reset (Optional but Recommended)   
  * Use the menu: Client Tools \> 🔁 Reset Calculated Formulas.   
  * This ensures that formulas in Columns F, H, I, and M–S are refreshed for accuracy.   
* (Optional) Unhide All Rows   
  * If you need to review all clients before processing:   
  * Use: Client Tools \> 🫣 Unhide All Client Rows

⸻

🔹 RUN Monthly Script

* Run: Client Tools > [Script Button] Or manually run: monthlyRolloverAndCreateDocs() from Apps Script.

Script Will:

* Calculate the previous month’s name (e.g., “July 2025”)
* Create Google Docs for each Active client
* Includes logo, support summary, block hours used, remaining, and overage
* Includes info from the directory (Domain Expire, Google Analytics access)
* Trash any duplicate documents from prior runs
* Move the new doc into the client’s Google Drive folder
* Insert a link in Column N (“Support Summary Link”) in the Master Tracker
* Log info in the Document Summary

⸻

🔹 AFTER Script Runs

* Review Output
* Check the Document Summary tab to confirm:
  * All expected clients are listed
  * Docs were successfully created
  * Spot check a few generated docs in client folders

* Hide Inactive Clients
  * Use: Client Tools > 🙈 Hide Inactive/Transitioning Rows
  * This will hide all rows in the Master Tracker for non-active clients

(Optional) Save/Archive the Month

* You may copy the Master Tracker tab or export a version for recordkeeping.

⸻

🧠 Notes
Dry Run Mode

* You can toggle const DRY_RUN = true in the script to simulate doc creation without actually generating files.

Client Folder Creation

* The script automatically creates a Google Drive folder for each client (if not already present) under the parent folder.

Hyperlink Placement

* Doc links are inserted into Column N.