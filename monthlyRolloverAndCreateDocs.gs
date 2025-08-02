// === Main Monthly Automation Script ===
/**
 * monthlyRolloverAndCreateDocs
 *
 * Runs the end-of-month rollover process:
 * - Identifies the prior month
 * - Iterates over each Active client in "Master Tracker"
 * - Creates and logs a support summary Google Doc
 */
function monthlyRolloverAndCreateDocs() {
  const DRY_RUN = false;

  // === Spreadsheet and UI references ===
  const ss = SpreadsheetApp.getActiveSpreadsheet();  // Reference to spreadsheet
  const ui = SpreadsheetApp.getUi();  // UI for alerts

  // === Primary data sheets ===
  const masterSheet = ss.getSheetByName("Master Tracker");
  const directorySheet = ss.getSheetByName("Client Directory");

  // === Timezone and parent Drive folder for client subfolders ===
  const timeZone = ss.getSpreadsheetTimeZone();
  const parentFolderId = '1UI4zQ_YIEWWJT0kSP2x8EaQlue303Xl-';
  const parentFolder = DriveApp.getFolderById(parentFolderId);

  // === Date labeling and logging ===
  const today = new Date();
  today.setMonth(today.getMonth() - 1);
  const monthLabel = Utilities.formatDate(today, timeZone, "MMMM yyyy");
  const timestamp = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss");

  // === Summary sheet prep ===
  const summarySheet = ss.getSheetByName("Document Summary") || ss.insertSheet("Document Summary");
  summarySheet.clear();
  summarySheet.appendRow(["Client Name", "Email", "Doc URL", "Timestamp"]);

  const dataRange = masterSheet.getDataRange();
  const data = dataRange.getValues();
  const rows = data.slice(1);
  let createdCount = 0;

  // === Process each client row ===
  rows.forEach((row, i) => {
    const rowNum = i + 2;
    const [ , clientName, , , , , , blockUsed, remainingBlock, uncoveredOverage, , , clientEmail, firstName, , status, domainExpire, accessToGA ] = row;
    const normalizedStatus = typeof status === "string" ? status.trim().toLowerCase() : "";
    const trimmedName = typeof clientName === "string" ? clientName.trim() : "";

    if (!trimmedName || normalizedStatus !== "active") {
      console.log(`⚠️ Skipping row ${rowNum}: Name="${trimmedName}" | Status="${normalizedStatus}"`);
      return;
    }

    const docName = `${monthLabel} - ${clientName}`;
    const clientFolder = getOrCreateClientFolder(parentFolder, clientName);

    // Delete old doc if already exists
    const existingFiles = clientFolder.getFilesByName(docName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    if (DRY_RUN) {
      console.log(`🟡 Dry run - would have created doc for ${clientName}`);
      return;
    }

    // === Create and write to Google Doc ===
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
 *
 * Locates or creates a Google Drive folder for a client within the parent folder.
 */
function getOrCreateClientFolder(parentFolder, clientName) {
  const folders = parentFolder.getFoldersByName(clientName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(clientName);
}


function resetFormulasInMasterTracker() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Tracker");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  for (let i = 2; i <= lastRow; i++) {
    const client = sheet.getRange(i, 2).getValue();
    if (!client) continue;

    // G - Overage Beyond Monthly Hours (hrs)
    sheet.getRange(i, 7).setFormula(`=IF(F${i}=0, 0, MAX(F${i} - D${i}, 0))`);

    // H - Block Hours Used (robust fix including negative block balances)
    sheet.getRange(i, 8).setFormula(`=IF(F${i}=0, 0, IF(F${i}<=D${i}, 0, IF(E${i}<=0, F${i}-D${i}, MIN(F${i}-D${i}, E${i}))))`);

    // I - Block Hours Remaining
    sheet.getRange(i, 9).setFormula(`=IF(H${i}="", "", E${i}-H${i})`);

    // J - Block Deficit Warning (hrs) — formula removed as column is deprecated
    // sheet.getRange(i, 10).setFormula(`=IF(AND(D${i}=0, E${i}<0), ABS(E${i}), 0)`);
  }

  SpreadsheetApp.getUi().alert("✅ Formulas in columns G–I reset. Column J has been deprecated.");
}

/**
 * insertAllMissingClients
 *
 * Appends any missing clients to the bottom of the Master Tracker.
 * Preserves formula integrity by NOT inserting alphabetically.
 */
function insertAllMissingClients() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const helperSheet = ss.getSheetByName("Active Clients Sorted");
  const masterSheet = ss.getSheetByName("Master Tracker");

  const helperClients = helperSheet.getRange(2, 1, helperSheet.getLastRow() - 1).getValues().flat();
  const existingClients = masterSheet.getRange(2, 2, masterSheet.getLastRow() - 1).getValues().flat();

  const missingClients = helperClients.filter(name => name && !existingClients.includes(name));

  if (missingClients.length === 0) {
    SpreadsheetApp.getUi().alert("✅ All active clients are already in the Master Tracker.");
    return;
  }

  const monthLabel = masterSheet.getRange(2, 1).getValue(); // Use the current month from row 2

  missingClients.forEach(name => {
    masterSheet.appendRow([monthLabel, name]);
  });

  SpreadsheetApp.getUi().alert(`✅ ${missingClients.length} missing client(s) added to the bottom of the Master Tracker.`);
}

/**
 * insertNewClientIntoDirectory
 *
 * Prompts the user for a new client name and status,
 * then adds them to the Client Directory if not already present.
 * Status should be added to Column D.
 */
function insertNewClientIntoDirectory() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Client Directory");
  const lastRow = sheet.getLastRow();

  const clientPrompt = ui.prompt("New Client", "Enter the client's domain (e.g., example.com):", ui.ButtonSet.OK_CANCEL);
  if (clientPrompt.getSelectedButton() !== ui.Button.OK) return;

  const clientName = clientPrompt.getResponseText().trim();
  if (!clientName) {
    ui.alert("No client name provided.");
    return;
  }

  const statusPrompt = ui.prompt("Client Status", "Enter status (Active, Inactive, Transitioning):", ui.ButtonSet.OK_CANCEL);
  if (statusPrompt.getSelectedButton() !== ui.Button.OK) return;

  const status = statusPrompt.getResponseText().trim();
  if (!status) {
    ui.alert("No status provided.");
    return;
  }

  const existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  if (existing.includes(clientName)) {
    ui.alert(`⚠️ ${clientName} already exists in the Client Directory.`);
    return;
  }

  // Column A = Client Name
  // Column D = Status
  sheet.appendRow([clientName, "", "", status]);
  ui.alert(`✅ ${clientName} added to the Client Directory.`);
}


/**
 * clearDocAndFolderLinks
 *
 * Wipes any previous HYPERLINKs or folder URLs from columns N and T on the Master Tracker.
 */
function clearDocAndFolderLinks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Tracker");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.getRange(2, 14, lastRow - 1).clearContent(); // Column N
  sheet.getRange(2, 20, lastRow - 1).clearContent(); // Column T
  SpreadsheetApp.getUi().alert("🧹 Cleared support summary and folder links from columns N and T.");
}

/**
 * onOpen
 *
 * Loads the "Client Tools" custom menu with all utility options.
 * Updated to reflect correct order and labeling.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🗂 Client Tools')
    .addItem('Run Monthly Rollover & Docs', 'monthlyRolloverAndCreateDocs')
    .addItem('Reset Master Tracker Formulas', 'resetFormulasInMasterTracker')
    .addItem('Insert New Client into Directory', 'insertNewClientIntoDirectory')
    .addItem('Append Missing Clients to Master Tracker', 'insertAllMissingClients') // <- new name for clarity
    .addItem('Sort Master Tracker A–Z', 'sortMasterTrackerAZ')
    .addItem('Clear Doc & Folder Links', 'clearDocAndFolderLinks')
    .addToUi();
}

/**
 * onEdit
 *
 * Automatically adds a timestamp to Column L of Client Directory
 * whenever any cell in columns A–K is edited.
 */
function onEdit(e) {
  const sheet = e.source.getSheetByName("Client Directory");
  if (!sheet || sheet.getName() !== e.range.getSheet().getName()) return;

  const editedColumn = e.range.getColumn();
  const editedRow = e.range.getRow();

  // Only trigger for rows 2+ and columns A–K (1–11)
  if (editedRow < 2 || editedColumn > 11) return;

  const timeZone = e.source.getSpreadsheetTimeZone();
  const now = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss");
  sheet.getRange(editedRow, 12).setValue(now); // Column L
}

/**
 * onEdit
 *
 * Automatically adds a timestamp to Column L of Client Directory
 * whenever any cell in columns A–K is edited.
 */
function onEdit(e) {
  const sheet = e.source.getSheetByName("Client Directory");
  if (!sheet || sheet.getName() !== e.range.getSheet().getName()) return;

  const editedColumn = e.range.getColumn();
  const editedRow = e.range.getRow();

  // Only trigger for rows 2+ and columns A–K (1–11)
  if (editedRow < 2 || editedColumn > 11) return;

  const timeZone = e.source.getSpreadsheetTimeZone();
  const now = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss");
  sheet.getRange(editedRow, 12).setValue(now); // Column L
}