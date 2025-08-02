// === Main Monthly Automation Script ===
/**
 * monthlyRolloverAndCreateDocs
 *
 * Runs the end-of-month rollover process:
 * - Identifies the prior month
 * - Iterates over each client in "Master Tracker"
 * - Creates and logs a support summary Google Doc
 */
function monthlyRolloverAndCreateDocs() {
  const DRY_RUN = false;

  // === Spreadsheet and UI references ===
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // === Sheet references ===
  const masterSheet = ss.getSheetByName("Master Tracker");
  const directorySheet = ss.getSheetByName("Client Directory");

  // === Drive folder & time metadata ===
  const timeZone = ss.getSpreadsheetTimeZone();
  const parentFolderId = '1UI4zQ_YIEWWJT0kSP2x8EaQlue303Xl-';
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const today = new Date();
  today.setMonth(today.getMonth() - 1);
  const monthLabel = Utilities.formatDate(today, timeZone, "MMMM yyyy");
  const timestamp = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss");

  // === Prepare summary sheet ===
  const summarySheet = ss.getSheetByName("Document Summary") || ss.insertSheet("Document Summary");
  summarySheet.clear();
  summarySheet.appendRow(["Client Name", "Email", "Doc URL", "Timestamp"]);

  const data = masterSheet.getDataRange().getValues();
  const rows = data.slice(1); // Skip header
  let createdCount = 0;

  // === Loop through all clients ===
  rows.forEach((row, i) => {
    const rowNum = i + 2;
    const [ , clientName, , , , , , blockUsed, remainingBlock, uncoveredOverage, , , clientEmail, firstName, , status, domainExpire, accessToGA ] = row;
    const trimmedName = typeof clientName === "string" ? clientName.trim() : "";

    if (!trimmedName) {
      console.log(`⚠️ Skipping row ${rowNum}: No client name found.`);
      return;
    }

    const docName = `${monthLabel} - ${clientName}`;
    const clientFolder = getOrCreateClientFolder(parentFolder, clientName);

    const existingFiles = clientFolder.getFilesByName(docName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    if (DRY_RUN) {
      console.log(`🟡 Dry run - would have created doc for ${clientName}`);
      return;
    }

    // === Create summary document ===
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    const logoBlob = DriveApp.getFileById("1fW300SGxEFVFvndaLkkWz3_O7L3BOq84").getBlob();
    body.appendImage(logoBlob).setWidth(250);

    body.appendParagraph("\nHello,");
    body.appendParagraph(`Here’s your monthly support summary for ${clientName} – ${monthLabel}:\n`);
    body.appendParagraph(`Block Hours Applied: ${blockUsed || 0}`);
    body.appendParagraph(`Remaining Block Balance: ${remainingBlock || 0}`);
    body.appendParagraph(`Overage Hours (Uncovered): ${uncoveredOverage || 0}`);
    body.appendParagraph("\nIf you need additional support hours, visit https://radiateu.com/request-support-time.");
    body.appendParagraph("\nFor our clients on a monthly plan:");
    body.appendParagraph("🔐 Domain Expiration: " + (domainExpire || "N/A"));
    body.appendParagraph("📊 Access to Google Analytics: " + (accessToGA || "N/A"));
    body.appendParagraph("\nIf you have any questions, feel free to reply here or send a message to support@radiateu.com.");
    body.appendParagraph("\n*If you have trouble accessing your support summary, let us know and we’ll send you a PDF version.*");
    doc.saveAndClose();

    const file = DriveApp.getFileById(doc.getId());
    file.moveTo(clientFolder);
    const docUrl = doc.getUrl();
    const hyperlink = `=HYPERLINK("${docUrl}", "Open Doc")`;
    masterSheet.getRange(rowNum, 14).setFormula(hyperlink);
    summarySheet.appendRow([clientName, clientEmail || "N/A", docUrl, timestamp]);
    console.log(`✅ Created support doc for ${clientName}`);
    createdCount++;
  });

  ui.alert(`✅ ${createdCount} support summaries were created.`);
}

/**
 * getOrCreateClientFolder
 * Creates or retrieves a folder for a client under the designated parent folder.
 */
function getOrCreateClientFolder(parentFolder, clientName) {
  const folders = parentFolder.getFoldersByName(clientName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(clientName);
}

/**
 * updateMasterTrackerFormulas
 * Rewrites static formulas in F (Hours Used), H (Block Used), I (Remaining Block)
 * and M–S (Client Directory lookups), using FILTER-based VLOOKUP to ensure only Active clients are referenced.
 */
function updateMasterTrackerFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName("Master Tracker");
  const startRow = 2;
  const lastRow = trackerSheet.getLastRow();
  const clientNames = trackerSheet.getRange(startRow, 2, lastRow - 1, 1).getValues();
  if (clientNames.length === 0) return;

  // === F, H, I columns ===
  const colF = [], colH = [], colI = [];
  for (let i = 0; i < clientNames.length; i++) {
    const row = i + startRow;
    const b = clientNames[i][0];
    if (!b) {
      colF.push([""]); colH.push([""]); colI.push([""]); continue;
    }
    colF.push([`=IF(B${row}="", "", IFERROR(VLOOKUP(B${row}, 'Time Entry'!A:L, 12, FALSE), 0))`]);
    colH.push([`=IF(F${row}=0, 0, IF(F${row}<=D${row}, 0, IF(E${row}<=0, F${row}-D${row}, MIN(F${row}-D${row}, E${row}))))`]);
    colI.push([`=MAX(E${row} - H${row}, 0)`]);
  }
  trackerSheet.getRange(startRow, 6, colF.length, 1).setFormulas(colF); // F
  trackerSheet.getRange(startRow, 8, colH.length, 1).setFormulas(colH); // H
  trackerSheet.getRange(startRow, 9, colI.length, 1).setFormulas(colI); // I

  // === M–S using FILTER-based VLOOKUP ===
const mToS = {
  L: { col: 12, index: 3 },  // Support Summary Link
  M: { col: 13, index: 7 },  // First Name
  N: { col: 14, index: 8 },  // Last Name
  O: { col: 15, index: 4 },  // Status
  P: { col: 16, index: 9 },  // Domain Expire
  Q: { col: 17, index: 10 }, // Access to GA
};

  Object.entries(mToS).forEach(([_, cfg]) => {
    const output = [];
    for (let i = 0; i < clientNames.length; i++) {
      const row = i + startRow;
      const b = clientNames[i][0];
      if (!b) { output.push([""]); continue; }
      const formula = `=IF(B${row}="", "", IFERROR(VLOOKUP(TO_TEXT(B${row}), FILTER('Client Directory'!A:J, 'Client Directory'!D:D = "Active"), ${cfg.index}, FALSE), ""))`;
      output.push([formula]);
    }
    trackerSheet.getRange(startRow, cfg.col, output.length, 1).setFormulas(output);
  });

  SpreadsheetApp.getUi().alert("✅ Formulas updated in Columns F, H, I, and M–S.");
}

/**
 * PURPOSE:
 * Inserts all clients from the Client Directory into the Master Tracker
 * if they are not already present. Preserves row order and copies formulas
 * from the template row (Row 2) to ensure consistency.
 *
 * Assumptions:
 * - Client Directory: client name in Column A, plan type in B, monthly hours in C, status in D, email in E, first name in F, last name in G
 * - Master Tracker: client name in Column B (index 2)
 * - Template row is Row 2 in Master Tracker (row with correct formulas)
 */

function insertAllMissingClients() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const directorySheet = ss.getSheetByName('Client Directory');
  const masterSheet = ss.getSheetByName('Master Tracker');

  const dirData = directorySheet.getDataRange().getValues();
  const masterData = masterSheet.getRange(2, 2, masterSheet.getLastRow() - 1).getValues(); // Column B = Client Name

  const masterClientNames = masterData.map(row => row[0].toString().trim().toLowerCase());
  const newClients = [];

  for (let i = 1; i < dirData.length; i++) {
    const name = dirData[i][0];
    if (!name) continue;
    const lowerName = name.toString().trim().toLowerCase();
    if (!masterClientNames.includes(lowerName)) {
      newClients.push(dirData[i]);
    }
  }

  if (newClients.length === 0) {
    console.log('✅ No missing clients to insert.');
    return;
  }

  const TEMPLATE_ROW = 2;

  newClients.forEach(client => {
    const lastRow = masterSheet.getLastRow();
    masterSheet.insertRowAfter(lastRow);

    // Copy formulas from Row 2
    const templateRange = masterSheet.getRange(TEMPLATE_ROW, 1, 1, masterSheet.getLastColumn());
    const newRowRange = masterSheet.getRange(lastRow + 1, 1, 1, masterSheet.getLastColumn());
    templateRange.copyTo(newRowRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

    // Fill in values from Client Directory
    masterSheet.getRange(lastRow + 1, 2).setValue(client[0]); // Client Name
    masterSheet.getRange(lastRow + 1, 3).setValue(client[1]); // Plan Type
    masterSheet.getRange(lastRow + 1, 4).setValue(client[2]); // Monthly Hours
    masterSheet.getRange(lastRow + 1, 15).setValue(client[3]); // Status (column O)
    masterSheet.getRange(lastRow + 1, 11).setValue(client[4]); // Email (column K)
    masterSheet.getRange(lastRow + 1, 13).setValue(client[5]); // First Name (M)
    masterSheet.getRange(lastRow + 1, 14).setValue(client[6]); // Last Name (N)
  });

  console.log(`✅ Inserted ${newClients.length} new client(s) into Master Tracker.`);
}

/**
 * 
 * PURPOSE:
 * This function hides rows in the "Master Tracker" sheet for clients marked as
 * "Inactive" or "Transitioning" in the Status column.
 * It is used for visual clarity without deleting or removing any data.
 * 
 * Assumes:
 * - Header is in Row 1
 * - Status is in Column O (Column 15)
 *
 * USAGE:
 * - Automatically runs from the "Client Tools" custom menu.
 * - Pair with `unhideAllClientRows()` to show all clients again.
 */

function hideInactiveAndTransitioningRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Master Tracker');

  const START_ROW = 2;
  const STATUS_COL = 15; // ✅ Column O — corrected from Column Q
  const numRows = sheet.getLastRow() - START_ROW + 1;
  const data = sheet.getRange(START_ROW, 1, numRows, sheet.getLastColumn()).getDisplayValues();

  // Unhide all rows before applying filter
  sheet.showRows(START_ROW, sheet.getMaxRows() - 1);

  let hiddenCount = 0;

  for (let i = 0; i < data.length; i++) {
    const statusRaw = data[i][STATUS_COL - 1];
    const status = statusRaw.toLowerCase().trim();

    console.log(`Row ${i + START_ROW}: Status = "${statusRaw}" → "${status}"`);

    if (status === 'inactive' || status === 'transitioning') {
      sheet.hideRows(i + START_ROW);
      console.log(`→ Hiding row ${i + START_ROW}`);
      hiddenCount++;
    }
  }

  console.log(`✅ Finished: ${hiddenCount} rows hidden.`);
}

/**
 * PURPOSE:
 * Utility function to unhide all client rows in the "Master Tracker" sheet.
 * 
 * USAGE:
 * - Manually run from the "Client Tools" menu if you want to view all clients again.
 */

function unhideAllClientRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master Tracker');
  sheet.showRows(2, sheet.getMaxRows() - 1); // Unhide all rows below header
}

/**
 * PURPOSE:
 * Inserts a new client into the "Client Directory" sheet.
 * Prompts the user for a client name, adds it to the next available row,
 * and sets default values like status ("Active") and a timestamp.
 *
 * ASSUMPTIONS:
 * - Column A: Client Name
 * - Column D: Status ("Active", "Inactive", etc.)
 * - Column L: Timestamp
 *
 * USAGE:
 * - Access via the "Client Tools" menu.
 * - Typically the first step in onboarding a new client.
 */
function insertNewClientIntoDirectory() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getSheetByName("Client Directory");
  const lastRow = sheet.getLastRow();

  // Prompt for client name
  const response = ui.prompt('Add New Client', 'Enter the client name (e.g., example.org):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const clientName = response.getResponseText().trim();
  if (!clientName) {
    ui.alert("Client name cannot be blank.");
    return;
  }

  // Check for duplicates
  const existingClients = sheet.getRange("A2:A" + lastRow).getValues().flat();
  if (existingClients.includes(clientName)) {
    ui.alert(`"${clientName}" already exists in the Client Directory.`);
    return;
  }

  // Find first empty row
  const targetRow = lastRow + 1;

  sheet.getRange(targetRow, 1).setValue(clientName);         // Column A: Client Name
  sheet.getRange(targetRow, 4).setValue("Active");           // Column D: Status
  sheet.getRange(targetRow, 12).setValue(new Date());        // Column L: Timestamp
}

/**
 * onOpen
 * Adds the Client Tools menu when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🗂 Client Tools')
    .addItem('Hide Inactive/Transitioning Rows', 'hideInactiveAndTransitioningRows')
    .addItem('Unhide All Client Rows', 'unhideAllClientRows')
    .addItem('Run Monthly Rollover & Docs', 'monthlyRolloverAndCreateDocs')
    .addItem('Fix Master Tracker Lookup Formulas (F, H, I, M–S)', 'updateMasterTrackerFormulas')
    .addItem('Insert New Client into Directory', 'insertNewClientIntoDirectory')
    .addItem('Insert All Missing Clients into Master Tracker', 'insertAllMissingClients')
    .addItem('Clear Doc & Folder Links', 'clearDocAndFolderLinks')
    .addItem('Sort Master Tracker A–Z', 'sortMasterTrackerAZ')
    .addToUi();
}