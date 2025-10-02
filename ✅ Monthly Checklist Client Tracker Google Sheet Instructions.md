Client Support Tracker URL - https://docs.google.com/spreadsheets/d/1TnQ_FSgTGRbz0KqzqLPKHJw7oQu9sC4hLCbTVXkxXhM/edit?gid=0#gid=0

Here’s a detailed End-of-Month Checklist based on your current monthlyRolloverAndCreateDocsSafe() script and overall system:

⸻

✅ End-of-Month Checklist

Purpose: Prepare your tracker, generate client support docs using Safe Mode, and reset formulas to begin the new month.

⸻

🔹 BEFORE Running the Script

* Confirm All Hours Are Entered  
  * Double-check that all time entries for the month are logged in the Time Entry tab.  
  * Ensure client time is accurately reflected for each client.   
* Make a copy of the Master Tracker Sheet and save it as the month just past for future records. Use copy and paste and make sure to paste special values only. 

* Update Block Hours  
 * In the Master Tracker, enter or update values in Column E (“Block Hours Available”) Paste them from Column I and use "paste values only" 

* Verify Client Statuses	In the Client Directory, confirm that:   
  * Active clients are marked as "Active" in Column D.   
  * Clients who should not receive reports are marked "Inactive" or "Transitioning".

* Insert Any New Clients   
  * Use the menu: Client Tools \> ➕ Add Client and Sync to Master Tracker.   
  * This ensures new clients are added to both sheets and properly formatted.   

* (Optional) Unhide All Rows   
  * If you need to review all clients before processing:   
  * Use: Client Tools \> 🫣 Unhide All Client Rows

⸻

🔹 RUN Monthly Script

* Run: Client Tools > [Safe Mode Script Button] Or manually run: monthlyRolloverAndCreateDocsSafe() from Apps Script.  
* Safe Mode is the recommended process and maintains one rolling row per client to prevent duplicate document creation.

Script Will:

* Calculate the previous month’s name (e.g., “July 2025”)  
* Create or update a single Google Doc per Active client with rolling monthly updates  
* Include logo, support summary, block hours used, remaining, and overage  
* Include info from the directory (Domain Expire, Google Analytics access)  
* Hide the KEY and DOC_ID columns used for tracking internally  
* Insert document links in Columns R and S (“Support Summary Link” and related) in the Master Tracker  
* Prevent duplicate documents by updating existing ones rather than creating new files  

⸻

🔹 AFTER Script Runs

* Review Output  
* Check the Master Tracker to confirm:  
  * All expected clients have updated links in Columns R and S  
  * Docs were successfully created or updated  
  * Spot check a few generated docs in client folders

* Hide Inactive Clients  
  * Use: Client Tools > 🙈 Hide Inactive/Transitioning Rows  
  * This will hide all rows in the Master Tracker for non-active clients

(Optional) Save/Archive the Month

* You may copy the Master Tracker tab or export a version for recordkeeping.

⸻

🧠 Notes  

Client Folder Creation

* The script automatically creates a Google Drive folder for each client (if not already present) under the parent folder.

Hyperlink Placement

* Doc links are inserted into Columns R and S.