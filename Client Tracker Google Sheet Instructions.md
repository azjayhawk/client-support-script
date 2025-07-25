# ✅ **Monthly Checklist**

RadiateU Client Reporting Use this checklist to guide your monthly process for generating client support summaries, updating data, and delivering reports.

---

### ✅ 1\. Before Running the Script

**Prep your data and check for accuracy.**

- **Ensure DRY\_RUN mode is ON:** Open the script and confirm:  
  const DRY\_RUN \= true;  
- **Sort Master Tracker:** Use the 🗂 Client Tools menu → **Sort Master Tracker A–Z**  
- **Time Entry tab:**  
  - Copy/paste weekly time entries into appropriate columns (format: hh:mm)  
  - Let formulas auto-calculate Total (Column M) and Decimal Hours (Column L)  
- **Block Carryover NOT needed yet:** Skip Column E for now — you’ll fill that in **after** monthly rollover is complete (see Step 5).

---

### ✅ 2\. Run the Script (Dry Run Mode)

**Preview output and verify everything before sending anything to clients.**

- Go to the Apps Script Editor (Extensions \> Apps Script)  
- Run: `monthlyRolloverAndCreateDocs()`  
- Verify:  
  - ✅ Only Active clients are processed (Status \= Active)  
  - ✅ Support summaries appear in each client folder  
  - ✅ Master Tracker Column N (Support Summary Link) contains working doc links  
  - ✅ Column T contains the correct Folder URL  
  - ✅ Document Summary sheet shows accurate results

---

### ✅ 3\. Enter Usage & Finalize Data

**Use this step to verify monthly usage and block hour activity.**

- In the **Master Tracker**:  
  - Column F (Hours Used) is pulled from Time Entry tab  
  - Columns G–J are auto-calculated (Overage Beyond Monthly Hours, Block Used, Remaining Block, Block Deficit Warning (hrs))  
- **Check Column H (Block Used)** is accurate based on usage vs. monthly plan  
- Adjust any rows manually as needed (e.g., transition clients or edge cases)

---

### ✅ 4\. Deliver Reports

You now have two delivery options:

#### Option A – YAMM (Yet Another Mail Merge)

Use YAMM to send personalized emails with support summaries.

- Use the links in Column N (Support Summary Link)  
- Pull email addresses from Column M

#### Option B – WP Umbrella (Monthly Automation)

Add each client’s folder link (Column K from Client Directory) into your WP Umbrella recurring task so they can access support summaries monthly.

---

### ✅ 5\. End-of-Month Finalization

**These steps should be done at the end of the current month.**

- Open the **Master Tracker**:  
  - **Copy values from Column I (Remaining Block)**  
  - **Paste values only into Column E (Block Hours Available)**

  This locks in the carryover for next month while preserving visibility throughout the current month.

- Rename the **Time Entry** tab (e.g., "Time Entry – July 2025") and create a new blank one for the next month  
- Run the script again with `DRY_RUN = false` if you want to finalize and regenerate the docs cleanly  
- Delete and regenerate support documents only if needed

---

# 📘 **Overview** 

This Google Sheet manages:

- Monthly hours and support block usage  
- Remaining time balances and overage calculations  
- Report creation and delivery via Google Docs  
- Optional email delivery using external tools (YAMM or WP Umbrella)

---

### **✅ Master Tracker – Column Guide**

| Column | Description |
| ----- | :---- |
| A | **Month** (e.g., “July 2025”) – inserted by script |
| B | **Client Name** – dropdown from helper sheet (active clients) |
| C | **Plan Type** – auto-filled via formula from Client Directory |
| D | **Monthly Plan Hours** – VLOOKUP from Plans tab |
| E | **Block Hours Available** – *manually pasted at end of month* |
| F | **Hours Used** – VLOOKUP from Time Entry tab |
| G | **Overage Beyond Monthly Hours (hrs)** – formula-driven |
| H | **Block Hours Used** – formula-driven |
| I | **Block Hours Remaining** – formula-driven |
| J | **Block Deficit Warning (hrs)** – formula-driven |
| K | **Cost** – calculated if uncovered time exists |
| L | **Notes** – optional manual input |
| M | **Client Email** – auto-filled from Client Directory |
| N | **Support Summary Link** – script inserts doc link |
| O | **First Name** – formula-filled from Client Directory |
| P | **Last Name** – formula-filled from Client Directory |
| Q | **Status** – used to filter clients (Active only) |
| R | **Domain Expire** – formula-filled from Client Directory |
| S | **Access to GA** – formula-filled from Client Directory |
| T | **Folder URL** – auto-filled by script |

---

### 

### **🛠 Script Functions Summary**

| Function | Purpose |
| :---- | :---- |
| monthlyRolloverAndCreateDocs() | Generates support summary documents for each **Active** client |
| resetFormulasInMasterTracker() | Resets formulas in columns **G–J** (Overage, Block, etc.) for all clients |
| insertNewClientIntoDirectory() | Prompts user to add a client to the **Client Directory** |
| insertAllMissingClients() | Adds any clients from the **helper sheet** who are missing from the Master Tracker |
| clearDocAndFolderLinks() | Clears old links from **columns N and T** |
| onOpen() | Loads the **Client Tools** menu with all script options |

---

# 📌 **Final Notes**

- Use the **Client Directory** as your source of truth  
- Clients must be marked **Active** to receive reports  
- The script automatically skips clients with blank names or inactive statuses  
- The "Support Summary Link" column (N) is refreshed each time the script runs

