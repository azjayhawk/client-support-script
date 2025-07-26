✅ Monthly Checklist

RadiateU Client Reporting
Use this checklist to guide your monthly process for generating support summaries, verifying data accuracy, and delivering reports.

⸻

✅ 1. Before Running the Script

Prep your data and confirm accuracy.
  • Ensure DRY_RUN mode is ON
Open the Apps Script (Extensions > Apps Script) and confirm:

const DRY_RUN = true;


  • Sort the Master Tracker
From the 🗂 Client Tools menu → select Sort Master Tracker A–Z
  • Time Entry Tab:
  • Copy/paste your weekly time entries (format: hh:mm) into the appropriate columns
  • Let built-in formulas auto-calculate:
  • Column L → Decimal Hours
  • Column M → Total Duration
  • These totals feed automatically into Column F of the Master Tracker

❌ Do NOT paste block hours yet – you’ll do that in Step 5.

⸻

✅ 2. Run the Script (Dry Run Mode)

This will preview the results and confirm doc generation logic is working correctly.
  • Open the Apps Script Editor
  • Run:

monthlyRolloverAndCreateDocs()


  • Check the following:
  • ✅ Only clients with Status = Active are included
  • ✅ One Google Doc is created in each client’s folder
  • ✅ Column N (Support Summary Link) has working hyperlinks
  • ✅ Column T contains the correct client Folder URL
  • ✅ The Document Summary sheet includes a list of documents and timestamps

⸻

✅ 3. Enter Usage & Finalize Data

At this point, you’ll review hours used and validate calculations.
  • In the Master Tracker:
  • Column F = Hours Used (from Time Entry)
  • Columns G–I auto-calculate based on usage:
  • G = Overage Beyond Monthly Hours
  • H = Block Hours Used
  • I = Block Hours Remaining
  • Column J (Block Deficit Warning) flags only when a client has:
  • 0 Monthly Hours, and
  • A negative block balance

🧠 Tip: You may ignore Column J for billing/reporting. It’s a visual alert for internal use only.

  • Adjust rows manually if needed (e.g., new/transitioning clients)

⸻

✅ 4. Deliver Reports

You have two supported options:

Option A – YAMM (Yet Another Mail Merge)
  • Use Column M (Email) for sending
  • Use Column N (Support Summary Link) for document URLs

Option B – WP Umbrella Integration
  • Copy each client’s Folder URL from Column T in the Master Tracker
  • Paste into their recurring WP Umbrella reporting task for ongoing access

⸻

✅ 5. End-of-Month Finalization

This locks in block carryover values for the next cycle.
  • In the Master Tracker:
  • Copy Column I (Block Hours Remaining)
  • Paste values only into Column E (Block Hours Available)

This makes Column E the starting balance for the next month while preserving visibility during the current cycle.

  • Rename the Time Entry tab (e.g., “Time Entry – July 2025”)
  • Create a blank copy for the upcoming month
  • Optional:
  • Set DRY_RUN = false and re-run script for clean final copies
  • Re-run only if changes were made or documents need refreshing

⸻

📘 Overview

This Google Sheet powers your full client support system:
  • Tracks monthly hours and block hour usage
  • Calculates remaining balances and overages
  • Auto-generates branded Google Docs for each client
  • Supports optional report delivery via email (YAMM) or WP Umbrella
  • Reduces manual entry through linked formulas and scripts

⸻

✅ Master Tracker – Column Guide

Column  Description
A Month (e.g., “July 2025”) – inserted by script
B Client Name – dropdown from helper sheet (Active clients only)
C Plan Type – auto-filled from Client Directory
D Monthly Plan Hours – auto-filled from Plans tab
E Block Hours Available – manually pasted at end of month
F Hours Used – pulled from Time Entry
G Overage Beyond Monthly Hours (hrs) – shows how many hours exceeded monthly allotment
H Block Hours Used – how many of the overage hours came from the block
I Block Hours Remaining – what’s left after usage
J Block Deficit Warning (hrs) – only applies if Monthly Plan = 0 and Block goes negative
K Cost – calculated separately if needed
L Notes – freeform manual input
M Client Email – auto-filled from Client Directory
N Support Summary Link – script inserts link
O First Name – auto-filled
P Last Name – auto-filled
Q Status – must be “Active” to trigger report
R Domain Expire – auto-filled
S Access to GA – auto-filled
T Folder URL – script auto-generates this path


⸻

🛠 Script Functions Summary

Function  Purpose
monthlyRolloverAndCreateDocs()  Generates support summary Google Docs for all Active clients
resetFormulasInMasterTracker()  Refreshes formulas in Columns G–J to maintain accurate calculations
insertNewClientIntoDirectory()  Prompts user to add a client (domain + name) into the Client Directory
insertAllMissingClients() Inserts any missing clients from helper sheet into the Master Tracker
clearDocAndFolderLinks()  Clears Columns N and T for a clean doc/folder refresh
onOpen()  Loads the “🗂 Client Tools” menu automatically when sheet opens


⸻

📌 Final Tips
  • Don’t forget to set DRY_RUN = false before finalizing
  • Clients must be Active to receive a doc
  • No doc is created if Client Name is blank
  • You may delete docs manually or allow the script to overwrite when re-run
  • The “Support Summary Link” in Column N refreshes with each run