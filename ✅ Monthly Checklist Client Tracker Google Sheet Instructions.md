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

* Bill Clients - Add any uncovered overage to clients in Stripe. Help Doc - https://docs.radiateu.com/docs/support-and-or-one-time-payment-to-stripe/
* Note those purchases on the master tracker and in the block history tab. 

* Verify Client Statuses	In the Client Directory, confirm that:   
  * Active clients are marked as "Active" in Column D.   
  * Clients who should not receive reports are marked "Inactive" or "Transitioning".

* Insert Any New Clients   
  * Use the menu: ❤️ Client Tools > ➕ Add Client and Sync to Master Tracker.   
  * This ensures new clients are added to both sheets and properly formatted.   
* Do NOT touch Column E (Block Hours Available) yet — changing it now will recalculate last month’s numbers.

* (Optional) Unhide All Rows   
  * If you need to review all clients before processing:   
  * Use: ❤️ Client Tools > 🫣 Unhide All Client Rows

⸻

🔹 RUN Monthly Script (Safe Mode Only)

* Run: 🛡️ Client Tools (Safe) > [Safe Mode Script Button]  
* Or manually run: monthlyRolloverAndCreateDocsSafe() from Apps Script.  
* Safe Mode is the only supported method and maintains one rolling row per client to prevent duplicate document creation.

Script Will:

* Calculate the previous month’s name (e.g., “July 2025”)  
* Create or update a single Google Doc per Active client with rolling monthly updates  
* Include logo, support summary, block hours used, remaining, and overage  
* Include info from the directory (Domain Expire, Google Analytics access)  
* Hide the KEY and DOC_ID columns used for tracking internally  
* Insert document links in Columns R and S (“Support Summary Link” and related) in the Master Tracker  
* Prevent duplicate documents by updating existing ones rather than creating new files  
* Automatically create client folders in Google Drive if missing
* Include new “Hours Purchased This Month” value from the Master Tracker (Column L) in each client’s support summary.

⸻

🔹 AFTER Script Runs

**• Advance Block Balances (carryover) — ONLY AFTER running Safe Mode script**  
  - Why after: Column E (Block Hours Available) is used by formulas for Block Hours Used/Remaining. If you change it before running the script, it will inflate/alter last month’s numbers.  
  - Action: Copy **I → E** (values only) for Active clients.  
    1) Select **I2:I** on Master Tracker and copy.  
    2) Select **E2:E**, then **Edit → Paste special → Paste values only**.  
    3) Spot-check a few rows to confirm E now equals the prior month’s remaining (I).

* Review Output  
  * Check the Master Tracker to confirm:  
    * All expected clients have updated links in Columns R and S  
    * Docs were successfully created or updated  
    * Client folders exist in Google Drive for each active client  
    * Spot check a few generated docs in client folders to verify content
    * Generated documents now include: Hours Used, Hours Purchased This Month, Block Hours Remaining, and Overage Beyond Block.

* Hide Inactive Clients  
  * Use: ❤️ Client Tools > 🙈 Hide Inactive/Transitioning Rows  
  * This will hide all rows in the Master Tracker for non-active clients

(Optional) Save/Archive the Month

* You may copy the Master Tracker tab or export a version for recordkeeping.

⸻

🧠 Notes  

* All old scripts are deprecated. Please use only the Safe Mode script via 🛡️ Client Tools (Safe).  
* Client Folder Creation is automatic in Safe Mode for any missing folders.  
* Doc links are inserted into Columns R and S in the Master Tracker for easy access.
* “Hours Purchased This Month” is now pulled from Column L in the Master Tracker. This reflects additional support hours purchased during the billing period.