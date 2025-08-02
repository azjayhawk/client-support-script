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
 * resetFormulasInMasterTracker
 * Reapplies formulas in columns G, H, I.
 */
function resetFormulasInMasterTracker() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Tracker");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  for (let i = 2; i <= lastRow; i++) {
    const client = sheet.getRange(i, 2).getValue();
    if (!client) continue;

    sheet.getRange(i, 7).setFormula(`=IF(F${i}=0, 0, MAX(F${i} - D${i}, 0))`);
    sheet.getRange(i, 8).setFormula(`=IF(F${i}=0, 0, IF(F${i}<=D${i}, 0, IF(E${i}<=0, F${i}-D${i}, MIN(F${i}-D${i}, E${i}))))`);
    sheet.getRange(i, 9).setFormula(`=IF(H${i}="", "", E${i}-H${i})`);
  }

  SpreadsheetApp.getUi().alert("✅ Master Tracker formulas (G–I) have been reset.");
}

/**
 * insertAllMissingClients
 * Pulls clients from the Client Directory and appends any that don’t already exist in Master Tracker.
 */
function insertAllMissingClients() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const directorySheet = ss.getSheetByName("Client Directory");
  const masterSheet = ss.getSheetByName("Master Tracker");

  const dirData = directorySheet.getDataRange().getValues();
  const masterClients = masterSheet.getRange(2, 2, masterSheet.getLastRow() - 1).getValues().flat();
  const monthLabel = masterSheet.getRange(2, 1).getValue();

  const allClients = dirData.slice(1).map(row => row[0]).filter(name => !!name);
  const missingClients = allClients.filter(name => !masterClients.includes(name));

  if (missingClients.length === 0) {
    SpreadsheetApp.getUi().alert("✅ All clients already exist in the Master Tracker.");
    return;
  }

  missingClients.forEach(name => {
    masterSheet.appendRow([monthLabel, name]);
  });

  SpreadsheetApp.getUi().alert(`✅ ${missingClients.length} client(s) added to the Master Tracker.`);
}

/**
 * insertNewClientIntoDirectory
 * Prompts the user to add a new client to the Client Directory.
 */
function insertNewClientIntoDirectory() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Client Directory");
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

  sheet.appendRow([clientName, "", "", status]);
  ui.alert(`✅ ${clientName} added to the Client Directory.`);
}

/**
 * clearDocAndFolderLinks
 * Clears columns N and T in the Master Tracker.
 */
function clearDocAndFolderLinks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Tracker");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.getRange(2, 14, lastRow - 1).clearContent(); // Column N
  sheet.getRange(2, 20, lastRow - 1).clearContent(); // Column T
  SpreadsheetApp.getUi().alert("🧹 Cleared support summary and folder links.");
}

/**
 * sortMasterTrackerAZ
 * Sorts Master Tracker A–Z by client name.
 */
function sortMasterTrackerAZ() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Tracker");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const rangeToSort = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  rangeToSort.sort({ column: 2, ascending: true });

  SpreadsheetApp.getUi().alert("✅ Master Tracker sorted A–Z by client name.");
}

/**
 * onOpen
 * Adds the Client Tools menu when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🗂 Client Tools')
    .addItem('Run Monthly Rollover & Docs', 'monthlyRolloverAndCreateDocs')
    .addItem('Reset Master Tracker Formulas', 'resetFormulasInMasterTracker')
    .addItem('Insert New Client into Directory', 'insertNewClientIntoDirectory')
    .addItem('Insert All Missing Clients into Master Tracker', 'insertAllMissingClients')
    .addItem('Clear Doc & Folder Links', 'clearDocAndFolderLinks')
    .addItem('Sort Master Tracker A–Z', 'sortMasterTrackerAZ')
    .addToUi();
}

/**
 * onEdit
 * Tracks updates in Client Directory by adding timestamp to Column L
 */
function onEdit(e) {
  const sheet = e.source.getSheetByName("Client Directory");
  if (!sheet || sheet.getName() !== e.range.getSheet().getName()) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row < 2 || col > 11) return;

  const now = Utilities.formatDate(new Date(), e.source.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  sheet.getRange(row, 12).setValue(now); // Column L
}