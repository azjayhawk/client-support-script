📊 Client Support Tracker – RadiateU

Google Apps Script system for managing monthly support reporting, client block hour tracking, and auto-generating Google Docs from a central tracker.

⸻

📋 Access to Live Sheet

👉 Monthly Support Tracker (Production)

⸻

✨ What It Does

This script automates the monthly rollover for RadiateU’s client support reporting process:
	•	Detects the previous month (e.g., “June 2025”)
	•	Iterates over all Active clients in the Master Tracker
	•	Generates branded Google Docs (support summaries) for each client
	•	Logs document links and folder paths in the tracker
	•	Resets calculated formulas (overage, block usage, deficit warnings)
	•	Keeps client folders organized in Google Drive
	•	Avoids duplicate creation with DRY_RUN preview mode

⸻

🚀 How to Use (in Google Sheets)
	1.	Open the sheet:
Go to Extensions > Apps Script
	2.	Enable dry-run preview:
Edit this line in the script to avoid file creation:

const DRY_RUN = true;


	3.	Run the function manually:
Run monthlyRolloverAndCreateDocs() from the script editor or
from the sheet menu:
🗂 Client Tools > Run Monthly Rollover & Docs
	4.	Verify output:
	•	Docs created in Drive
	•	Column N: support doc links
	•	Column T: client folder URLs
	•	Document Summary tab: audit trail
	5.	Set DRY_RUN = false when ready to create final docs

⸻

🧰 Client Tools Menu

Menu Item	Function
Run Monthly Rollover & Docs	Generates support summary Google Docs
Reset Master Tracker Formulas	Refreshes Columns G–J (Overage, Block Used, etc.)
Clear Doc & Folder Links	Clears Columns N (Doc Link) and T (Folder URL)
(Optional additions)	insertNewClientIntoDirectory() and insertAllMissingClients() available in full script


⸻

🗂 Sheet Architecture

Master Tracker Columns

Column	Description
A	Month (e.g., “July 2025”)
B	Client Name
C	Plan Type
D	Monthly Plan Hours
E	Block Hours Available (manually pasted at month end)
F	Hours Used (from Time Entry)
G	Overage Beyond Monthly Hours (hrs)
H	Block Hours Used
I	Block Hours Remaining
J	Block Deficit Warning (hrs) (shows if Block < 0 & no monthly plan)
K	Notes
M–T	Client Email, Doc URL, First Name, Status, GA access, Folder URL


⸻

📁 Project Structure

File	Purpose
monthlyRolloverAndCreateDocs.gs	Main logic – creates documents
resetFormulasInMasterTracker.gs	Utility – resets formulas in columns G–J
helpers.gs	(optional) Utility functions like folder creation, sorting
appsscript.json	Script manifest
README.md	Documentation


⸻

💡 Developer Notes
	•	Use DRY_RUN = true for test runs – no docs created
	•	Docs are replaced each time with the same filename
	•	Block Deficit Warning column (J) is only visual – not included in reports
	•	For full month-end checklist, see the Instructions Doc

⸻

🛠 Powered By
	•	Google Apps Script
	•	CLASP
	•	GitHub – version history & backups

⸻

🧪 CLASP Command Tips

clasp login         # authenticate once
clasp pull          # pull latest code from Apps Script
clasp push          # push local code to Apps Script
clasp open          # open the script editor in browser
