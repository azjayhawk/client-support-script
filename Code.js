// === Main Monthly Automation Script ===
/**
 * monthlyRolloverAndCreateDocs
 *
 * Runs the end-of-month rollover process:
 * - Identifies the prior month
 * - Iterates over each client in "Master Tracker"
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
    const [ , clientName, planType, , , , , blockUsed, remainingBlock, uncoveredOverage, clientEmail, , firstName, , status, domainExpire, accessToGA ] = row;

const trimmedName = typeof clientName === "string" ? clientName.trim() : "";
if (!trimmedName) {
  console.log(`⚠️ Skipping row ${rowNum}: No client name found.`);
  return;
}

// ✅ Skip clients with "Hosting" plan
if (planType && planType.toString().trim().toLowerCase() === "hosting") {
  console.log(`⏭️ Skipping ${clientName} (Hosting plan)`);
  return;
}

// ✅ Skip clients with "Inactive" or "Transitioning" status
if (status && ["inactive", "transitioning"].includes(status.toString().trim().toLowerCase())) {
  console.log(`⏭️ Skipping ${clientName} (${status} status)`);
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
    const file = clientFolder.createFile(docName, "");
    const doc = DocumentApp.openById(file.getId());
    const body = doc.getBody();
    const logoBlob = DriveApp.getFileById("1fW300SGxEFVFvndaLkkWz3_O7L3BOq84").getBlob();
    const image = body.appendImage(logoBlob);
    const aspectRatio = image.getHeight() / image.getWidth();
    image.setWidth(150).setHeight(150 * aspectRatio); // smaller logo

    body.appendParagraph(`Hello ${firstName || ""},\n`);
    body.appendParagraph(""); // blank line
    body.appendParagraph(`Here’s your monthly support summary for ${clientName} – ${monthLabel}:\n`);
    body.appendParagraph(`Block Hours Applied: ${blockUsed || 0}`);
    body.appendParagraph(`Remaining Block Balance: ${remainingBlock || 0}`);
    body.appendParagraph(`Overage Hours (Uncovered): ${uncoveredOverage || 0}`);
    body.appendParagraph(""); // blank line
    body.appendParagraph("\nIf you need additional support hours, visit https://radiateu.com/request-support-time.");
    body.appendParagraph(""); // blank line
    body.appendParagraph("\nFor our clients on a monthly plan:");

    const formattedDomainExpire = domainExpire instanceof Date
      ? Utilities.formatDate(domainExpire, timeZone, "MMM dd, yyyy")
      : (domainExpire || "N/A");

    body.appendParagraph("🔐 Domain Expiration: " + formattedDomainExpire);
    body.appendParagraph("📊 Access to Google Analytics: " + (accessToGA || "N/A"));
    body.appendParagraph("\nIf you have any questions, feel free to reply here or send a message to support@radiateu.com.");
    doc.saveAndClose();

    const docUrl = doc.getUrl();
    const hyperlink = `=HYPERLINK("${docUrl}", "Open Doc")`;
    masterSheet.getRange(rowNum, 19).setFormula(hyperlink); // ✅ Column S (the 19th column)
    summarySheet.appendRow([clientName, clientEmail || "N/A", docUrl, timestamp]);
    console.log(`✅ Created support doc for ${clientName}`);
    masterSheet.getRange(rowNum, 18).setFormula(`=HYPERLINK("${clientFolder.getUrl()}", "Open Folder")`);
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
 * monthlyRolloverAndCreateDocsSafe
 *
 * Safe mode version of the rollover script:
 * - Ensures only one rolling row per client
 * - Does NOT duplicate clients each month
 * - Places docs into the correct client folder
 */
function monthlyRolloverAndCreateDocsSafe() {
  const DRY_RUN = false; // set true if testing only

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const masterSheet = ss.getSheetByName("Master Tracker");
  const directorySheet = ss.getSheetByName("Client Directory");

  const timeZone = ss.getSpreadsheetTimeZone();
  const parentFolderId = '1UI4zQ_YIEWWJT0kSP2x8EaQlue303Xl-';
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const today = new Date();
  today.setMonth(today.getMonth() - 1);
  const monthLabel = Utilities.formatDate(today, timeZone, "MMMM yyyy");
  const timestamp = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss");

  const summarySheet = ss.getSheetByName("Document Summary") || ss.insertSheet("Document Summary");
  summarySheet.clear();
  summarySheet.appendRow(["Client Name", "Email", "Doc URL", "Timestamp"]);

  const data = masterSheet.getDataRange().getValues();
  const rows = data.slice(1); // skip header
  let createdCount = 0;

  rows.forEach((row, i) => {
    const rowNum = i + 2;
    const [, clientName, planType, , , , , blockUsed, remainingBlock, uncoveredOverage, clientEmail, , firstName, , status, domainExpire, accessToGA] = row;

    const trimmedName = typeof clientName === "string" ? clientName.trim() : "";
    if (!trimmedName) return;

    if (planType && planType.toString().trim().toLowerCase() === "hosting") return;
    if (status && ["inactive", "transitioning"].includes(status.toString().trim().toLowerCase())) return;

    const docName = `${monthLabel} - ${clientName}`;
    const clientFolder = getOrCreateClientFolder(parentFolder, clientName);

    // ✅ Safe mode: reuse active doc if it exists, otherwise create new
    let doc;
    let file;
    let activeFile = null;
    const existingFiles = clientFolder.getFilesByName(docName);
    while (existingFiles.hasNext()) {
      const candidate = existingFiles.next();
      if (!candidate.isTrashed()) {  // ignore trashed files
        activeFile = candidate;
        break;
      }
    }

    if (activeFile) {
      file = activeFile;
      doc = DocumentApp.openById(file.getId());
      console.log(`♻️ Reusing existing doc for ${clientName}`);
    } else {
      file = clientFolder.createFile(docName, "");
      doc = DocumentApp.openById(file.getId());
      console.log(`🆕 Created new doc for ${clientName}`);
    }

    if (DRY_RUN) {
      console.log(`🟡 Dry run - would have created doc for ${clientName}`);
      return;
    }

    // === Create or refresh document content ===
    const body = doc.getBody();
    body.clear();
    const logoBlob = DriveApp.getFileById("1fW300SGxEFVFvndaLkkWz3_O7L3BOq84").getBlob();
    const image = body.appendImage(logoBlob);
    const originalWidth = image.getWidth();
    const originalHeight = image.getHeight();
    const targetWidth = 200; // set consistent width
    const scaledHeight = Math.round((originalHeight / originalWidth) * targetWidth);
    image.setWidth(targetWidth).setHeight(scaledHeight); // preserve aspect ratio

    body.appendParagraph(`Hello ${firstName || ""},\n`);
    body.appendParagraph(`Here’s your monthly support summary for ${clientName} – ${monthLabel}:\n`);
    body.appendParagraph(`Block Hours Applied: ${blockUsed || 0}`);
    body.appendParagraph(`Remaining Block Balance: ${remainingBlock || 0}`);
    body.appendParagraph(`Overage Hours (Uncovered): ${uncoveredOverage || 0}`);
    body.appendParagraph("\nIf you need additional support hours, visit https://radiateu.com/request-support-time.");

    const formattedDomainExpire = domainExpire instanceof Date
      ? Utilities.formatDate(domainExpire, timeZone, "MMM dd, yyyy")
      : (domainExpire || "N/A");

    body.appendParagraph("🔐 Domain Expiration: " + formattedDomainExpire);
    body.appendParagraph("📊 Access to Google Analytics: " + (accessToGA || "N/A"));
    body.appendParagraph("\nIf you have any questions, feel free to reply here or send a message to support@radiateu.com.");
    doc.saveAndClose();

    // file has already been assigned and moved above
    const docUrl = doc.getUrl();
    masterSheet.getRange(rowNum, 19).setFormula(`=HYPERLINK("${docUrl}", "Open Doc")`);
    masterSheet.getRange(rowNum, 18).setFormula(`=HYPERLINK("${clientFolder.getUrl()}", "Open Folder")`);

    summarySheet.appendRow([clientName, clientEmail || "N/A", docUrl, timestamp]);
    console.log(`✅ Created support doc for ${clientName}`);
    createdCount++;
  });

  ui.alert(`✅ ${createdCount} support summaries were created.`);
}
/**
 * updateMasterTrackerFormulas
 * Rewrites static formulas in F (Hours Used), H (Block Used), I (Remaining Block)
 * and L–Q (Client Directory lookups), using FILTER-based VLOOKUP to ensure only Active clients are referenced.
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
    // Uses Time Entry column O (15th) rather than L (12th)
    colF.push([`=IF(B${row}="", "", IFERROR(VLOOKUP(B${row}, 'Time Entry'!A:O, 15, FALSE), 0))`]);
    colH.push([`=IF(F${row}=0, 0, IF(F${row}<=D${row}, 0, IF(E${row}<=0, F${row}-D${row}, MIN(F${row}-D${row}, E${row}))))`]);
    colI.push([`=MAX(E${row} - H${row}, 0)`]);
  }
  trackerSheet.getRange(startRow, 6, colF.length, 1).setFormulas(colF); // F
  trackerSheet.getRange(startRow, 8, colH.length, 1).setFormulas(colH); // H
  trackerSheet.getRange(startRow, 9, colI.length, 1).setFormulas(colI); // I

  // === L–Q using FILTER-based VLOOKUP (Active-only from Client Directory) ===
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

  SpreadsheetApp.getUi().alert("✅ Formulas updated in Columns F, H, I, and L–Q.");
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
  const ui = SpreadsheetApp.getUi();
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
    ui.alert('✅ No missing clients to insert.');
    console.log('✅ No missing clients to insert.');
    return;
  }

  const TEMPLATE_ROW = 2;

  newClients.forEach(client => {
    const lastRow = masterSheet.getLastRow();
    masterSheet.insertRowAfter(lastRow);

    // Copy formulas from template row
    const templateRange = masterSheet.getRange(TEMPLATE_ROW, 1, 1, masterSheet.getLastColumn());
    const newRowRange = masterSheet.getRange(lastRow + 1, 1, 1, masterSheet.getLastColumn());
    templateRange.copyTo(newRowRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

    // Fill in values from Client Directory
    masterSheet.getRange(lastRow + 1, 2).setValue(client[0]);  // Client Name
    masterSheet.getRange(lastRow + 1, 3).setValue(client[1]);  // Plan Type
    // ⚠️ Skip Column D (Monthly Hours) — preserve formula
    masterSheet.getRange(lastRow + 1, 15).setValue(client[3]); // Status
    masterSheet.getRange(lastRow + 1, 11).setValue(client[2]); // Email
    masterSheet.getRange(lastRow + 1, 13).setValue(client[5]); // First Name
    masterSheet.getRange(lastRow + 1, 14).setValue(client[6]); // Last Name

    // 🚫 Clear copied links/formulas from template row in Columns R & S
    masterSheet.getRange(lastRow + 1, 18).clearContent(); // Report Folder URL
    masterSheet.getRange(lastRow + 1, 19).clearContent(); // Report URL
  });

  const message = `✅ Inserted ${newClients.length} new client(s) into Master Tracker.`;
  console.log(message);
  ui.alert(message);
}

/**
 * PURPOSE:
 * Inserts a new client into the "Client Directory" sheet with default values
 * and copies formatting & data validation from the template row (Row 2).
 */
function insertNewClientIntoDirectory() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Client Directory');

  // Prompt user for input
  const clientNamePrompt = ui.prompt('New Client', 'Enter the client domain or name:', ui.ButtonSet.OK_CANCEL);
  if (clientNamePrompt.getSelectedButton() !== ui.Button.OK) return;

  const planPrompt = ui.prompt('Plan Type', 'Enter plan type:', ui.ButtonSet.OK_CANCEL);
  if (planPrompt.getSelectedButton() !== ui.Button.OK) return;

  const emailPrompt = ui.prompt('Email', 'Enter client email address:', ui.ButtonSet.OK_CANCEL);
  if (emailPrompt.getSelectedButton() !== ui.Button.OK) return;

  const partnerPrompt = ui.prompt('Client Partner', 'Enter Client Partner (if applicable):', ui.ButtonSet.OK_CANCEL);
  if (partnerPrompt.getSelectedButton() !== ui.Button.OK) return;

  // Determine insertion point
  const lastRow = sheet.getLastRow();
  const newRowIndex = lastRow + 1;

  // Insert blank row at bottom
  sheet.insertRowsAfter(lastRow, 1);

  // Copy formatting + data validation from Row 2
  const templateRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
  const newRowRange = sheet.getRange(newRowIndex, 1, 1, sheet.getLastColumn());
  templateRange.copyTo(newRowRange, { formatOnly: true });

  // Fill in data values
  const newRowValues = [];
  newRowValues[0] = clientNamePrompt.getResponseText().trim();  // A: Client Name
  newRowValues[1] = planPrompt.getResponseText().trim();        // B: Plan Type
  newRowValues[2] = emailPrompt.getResponseText().trim();       // C: Email
  newRowValues[3] = 'Active';                                   // D: Status (default)
  newRowValues[4] = partnerPrompt.getResponseText().trim();     // E: Client Partner

  sheet.getRange(newRowIndex, 1, 1, newRowValues.length).setValues([newRowValues]);

  ui.alert('✅ New client added to Client Directory.');
}

/**
 * Sorts the Master Tracker sheet A–Z by Client Name (Column B).
 * Header in Row 1 is preserved.
 */
function sortMasterTrackerAZ() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master Tracker');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // Unhide all client rows before sorting
  sheet.showRows(2, sheet.getMaxRows() - 1);

  // Sort by Column B (Client Name), ascending
  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  range.sort({ column: 2, ascending: true });

  // Re-hide Inactive and Transitioning rows
  hideInactiveAndTransitioningRows();

  console.log('✅ Master Tracker sorted A–Z by Client Name (including hidden rows).');
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'Client Directory') return;

  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();

  // Only run if a cell in Columns A–K (1–11) was edited, and not the header
  if (editedRow > 1 && editedCol >= 1 && editedCol <= 11) {
    const timestampCell = sheet.getRange(editedRow, 11); // Column K
    timestampCell.setValue(new Date());
  }
}

/**
 * PURPOSE:
 * Combined tool to:
 * 1. Prompt user to add a new client to the Client Directory
 * 2. Add any missing clients from the directory into the Master Tracker
 * 3. Sort Master Tracker alphabetically
 *
 * USAGE:
 * - Adds formatting & validation from template row in directory
 * - Adds missing rows to tracker with formulas from Row 2
 * - Ensures the tracker is sorted A–Z by client name
 */

function addNewClientToTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const directorySheet = ss.getSheetByName('Client Directory');
  const masterSheet = ss.getSheetByName('Master Tracker');

  // === Step 1: Prompt for Client Info ===
  const namePrompt = ui.prompt('New Client', 'Enter the client domain or name:', ui.ButtonSet.OK_CANCEL);
  if (namePrompt.getSelectedButton() !== ui.Button.OK) return;
  const clientName = namePrompt.getResponseText().trim();

  const planPrompt = ui.prompt('Plan Type', 'Enter plan type:', ui.ButtonSet.OK_CANCEL);
  if (planPrompt.getSelectedButton() !== ui.Button.OK) return;
  const planType = planPrompt.getResponseText().trim();

  const emailPrompt = ui.prompt('Email', 'Enter client email address:', ui.ButtonSet.OK_CANCEL);
  if (emailPrompt.getSelectedButton() !== ui.Button.OK) return;
  const email = emailPrompt.getResponseText().trim();

  const partnerPrompt = ui.prompt('Client Partner', 'Enter Client Partner (if applicable):', ui.ButtonSet.OK_CANCEL);
  if (partnerPrompt.getSelectedButton() !== ui.Button.OK) return;
  const partner = partnerPrompt.getResponseText().trim();

  const firstNamePrompt = ui.prompt('First Name', 'Enter the first name of the client:', ui.ButtonSet.OK_CANCEL);
  if (firstNamePrompt.getSelectedButton() !== ui.Button.OK) return;
  const firstName = firstNamePrompt.getResponseText().trim();

  const lastNamePrompt = ui.prompt('Last Name', 'Enter the last name of the client:', ui.ButtonSet.OK_CANCEL);
  if (lastNamePrompt.getSelectedButton() !== ui.Button.OK) return;
  const lastName = lastNamePrompt.getResponseText().trim();

  // === Step 2: Insert into Client Directory ===
  const lastDirRow = directorySheet.getLastRow();
  const newDirRow = lastDirRow + 1;
  directorySheet.insertRowsAfter(lastDirRow, 1);

  const templateDirRange = directorySheet.getRange(2, 1, 1, directorySheet.getLastColumn());
  const newDirRange = directorySheet.getRange(newDirRow, 1, 1, directorySheet.getLastColumn());
  templateDirRange.copyTo(newDirRange, { formatOnly: true });

  const dirValues = [];
  dirValues[0] = clientName;   // A - Client Name
  dirValues[1] = planType;     // B - Plan Type
  dirValues[2] = email;        // C - Email
  dirValues[3] = 'Active';     // D - Status
  dirValues[4] = partner;      // E - Client Partner
  directorySheet.getRange(newDirRow, 1, 1, dirValues.length).setValues([dirValues]);

  console.log("Client added to Client Directory: " + clientName);

  // === Step 3: Sync to Master Tracker if missing ===
  const dirData = directorySheet.getDataRange().getValues();
  const masterNames = masterSheet.getRange(2, 2, masterSheet.getLastRow() - 1).getValues()
    .map(row => row[0].toString().trim().toLowerCase());

  if (!masterNames.includes(clientName.toLowerCase())) {
    const clientRow = dirData.find(row => row[0].toString().trim().toLowerCase() === clientName.toLowerCase());
    if (!clientRow) {
      console.warn('⚠️ Could not find client in directory data after adding.');
      return;
    }

    const insertRow = masterSheet.getLastRow() + 1;
    masterSheet.insertRowAfter(insertRow);

    const templateRow = masterSheet.getRange(2, 1, 1, masterSheet.getLastColumn());
    const newTrackerRow = masterSheet.getRange(insertRow + 1, 1, 1, masterSheet.getLastColumn());
    templateRow.copyTo(newTrackerRow, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

    // Prevent auto-hyperlinking of Client Name
    masterSheet.getRange(insertRow + 1, 2).setNumberFormat('@STRING@');

    masterSheet.getRange(insertRow + 1, 2).setValue(clientRow[0]);  // Client Name
    masterSheet.getRange(insertRow + 1, 3).setValue(clientRow[1]);  // Plan Type
    masterSheet.getRange(insertRow + 1, 11).setValue(clientRow[2]); // Email
    masterSheet.getRange(insertRow + 1, 15).setValue(clientRow[3]); // Status
    masterSheet.getRange(insertRow + 1, 13).setValue(clientRow[5]); // First Name
    masterSheet.getRange(insertRow + 1, 14).setValue(clientRow[6]); // Last Name

    console.log(`Client added to Master Tracker: ${clientRow[0]}`);
  }

  // === Step 4: Sort Master Tracker A–Z by Client Name (Column B) ===
  const lastRow = masterSheet.getLastRow();
  if (lastRow > 2) {
    masterSheet.getRange(2, 1, lastRow - 1, masterSheet.getLastColumn())
      .sort({ column: 2, ascending: true });
    console.log('✅ Master Tracker sorted.');
  }

  // === Step 5: Ensure the client exists in Weekly Time Entry (permanent) ===
  try {
    ensureClientInTimeEntry_(clientName);
  } catch (e) {
    console.warn('Time Entry add skipped: ' + e);
  }

  ui.alert(`Client "${clientName}" added and synced to Master Tracker & Time Entry.`);
}
/**
 * Hides rows in the Master Tracker for clients marked as Inactive or Transitioning.
 * Assumes:
 * - "Master Tracker" has status in Column O (15)
 * - Header is in Row 1
 */
function hideInactiveAndTransitioningRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master Tracker');
  const START_ROW = 2;
  const STATUS_COL = 15;
  const numRows = sheet.getLastRow() - 1;

  const statuses = sheet.getRange(START_ROW, STATUS_COL, numRows).getValues();
  sheet.showRows(START_ROW, sheet.getMaxRows() - 1); // Unhide all first

  let hiddenCount = 0;
  for (let i = 0; i < statuses.length; i++) {
    const status = (statuses[i][0] || '').toString().toLowerCase().trim();
    if (status === 'inactive' || status === 'transitioning') {
      sheet.hideRows(START_ROW + i);
      hiddenCount++;
    }
  }

  console.log(`✅ ${hiddenCount} row(s) hidden based on status.`);
}

/**
 * Unhides all rows in the Master Tracker (below the header).
 */
function unhideAllClientRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master Tracker');
  sheet.showRows(2, sheet.getMaxRows() - 1);
  console.log('✅ All client rows unhidden.');
}

/**
 * onOpen
 * Adds the Client Tools menu when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('❤️ Client Tools')
    .addItem('📄 Run Monthly Rollover + Create Docs (Safe)', 'monthlyRolloverAndCreateDocsSafe')
    .addSeparator()
    .addItem('➕ Add Client and Sync to Master Tracker', 'addNewClientToTracker')
    .addSeparator()
    .addItem('📋 Insert All Missing Clients to Master Tracker', 'insertAllMissingClients')
    .addItem('🔤 Sort Master Tracker A–Z', 'sortMasterTrackerAZ')
    .addItem('🗂 Insert New Client into Directory', 'insertNewClientIntoDirectory')
    .addSeparator()
    .addItem('🙈 Hide Inactive/Transitioning Rows', 'hideInactiveAndTransitioningRows')
    .addItem('🫣 Unhide All Client Rows', 'unhideAllClientRows')
    .addToUi();
}
// Time Entry sheet settings (GLOBAL)
const TIME_ENTRY_SHEET = 'Time Entry';      // tab name
const TIME_ENTRY_TEMPLATE_ROW = 3;          // ⬅️ change this to 3

/**
 * === Time Entry Helpers ===
 * Ensures clients are also present in the "Time Entry" tab.
 * Guarantees a single row for the client in Time Entry:
 *  • Doesn’t bail out if the sheet only has headers
 *  • Copies the Row 2 template (formulas/validation)
 *  • Sets the client name in Column A as text (avoids auto-linking)
 *  • Idempotent (does nothing if the client already exists)
 */
function ensureClientInTimeEntry_(clientName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const te = ss.getSheetByName(TIME_ENTRY_SHEET);
  if (!te) throw new Error(`Sheet "${TIME_ENTRY_SHEET}" not found.`);

  const lastRow = te.getLastRow(); // may be 1 if only headers

  // SAFE: only read names if at least one data row exists
  let names = [];
  if (lastRow >= 2) {
    names = te.getRange(2, 1, lastRow - 1, 1)
      .getValues()
      .map(r => (r[0] || '').toString().trim().toLowerCase());
  }

  const target = (clientName || '').toString().trim().toLowerCase();
  if (!target) return false;
  if (names.indexOf(target) !== -1) return false; // already present

  // Insert a new row, copy template from Row 2
  const insertAt = lastRow + 1;                 // if only header, this becomes 2
  te.insertRowAfter(insertAt - 1);              // inserts after row 1 safely
  te.getRange(TIME_ENTRY_TEMPLATE_ROW, 1, 1, te.getLastColumn())
    .copyTo(te.getRange(insertAt, 1, 1, te.getLastColumn()), { contentsOnly: false });

  te.getRange(insertAt, 1).setValue(clientName).setNumberFormat('@STRING@'); // prevent auto-link
  console.log(`✅ Added "${clientName}" to Time Entry at row ${insertAt}`);
  return true;
}

/**
 * Utility to sync all clients from Master Tracker into Time Entry
 * (Run manually if needed).
 */
function syncAllClientsToTimeEntry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tracker = ss.getSheetByName('Master Tracker');
  const names = tracker.getRange(2, 2, tracker.getLastRow() - 1, 1).getValues();

  let added = 0;
  names.forEach(r => {
    const clientName = r[0];
    if (clientName && ensureClientInTimeEntry_(clientName)) {
      added++;
    }
  });
  SpreadsheetApp.getUi().alert(`✅ Sync complete. ${added} client(s) added to Time Entry.`);
}

