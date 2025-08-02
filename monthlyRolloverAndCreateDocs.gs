// === RadiateU Client Support Automation Script ===
//
// This script automates the monthly support summary process for RadiateU clients.
// It performs client folder setup, generates support summary Google Docs,
// manages block hour tracking, and provides utility functions via a custom Client Tools menu.
//
// Each function is documented below with its purpose and usage.
//
// Author: [Your Name]
// Last Updated: 2025-08-01


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

  const data = masterSheet.getDataRange().getValues().slice(1);
  let createdCount = 0;

  data.forEach((row, i) => {
    const rowNum = i + 2;
    const [ , clientName, , , , , , blockUsed, remainingBlock, uncoveredOverage, , , clientEmail, firstName, , status, domainExpire, accessToGA ] = row;
    const normalizedStatus = (status || "").toLowerCase().trim();
    const trimmedName = (clientName || "").trim();

    if (!trimmedName || normalizedStatus !== "active") {
      console.log(`⚠️ Skipping row ${rowNum}: Name="${trimmedName}" | Status="${normalizedStatus}"`);
      return;
    }

    const docName = `${monthLabel} - ${clientName}`;
    const clientFolder = getOrCreateClientFolder(parentFolder, clientName);
    const existingFiles = clientFolder.getFilesByName(docName);
    while (existingFiles.hasNext()) existingFiles.next().setTrashed(true);

    if (DRY_RUN) {
      console.log(`🟡 Dry run - would have created doc for ${clientName}`);
      return;
    }

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
    masterSheet.getRange(rowNum, 14).setFormula(`=HYPERLINK("${doc.getUrl()}", "Open Doc")`);
    summarySheet.appendRow([clientName, clientEmail || "N/A", doc.getUrl(), timestamp]);
    console.log(`✅ Created support doc for ${clientName}`);
    createdCount++;
  });

  ui.alert(`✅ ${createdCount} support summaries were created.`);
}

/**
 * getOrCreateClientFolder
 *
 * Ensures a Drive folder exists for the given client.
 */
function getOrCreateClientFolder(parentFolder, clientName) {
  const folders = parentFolder.getFoldersByName(clientName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(clientName);
}

/**
 * resetFormulasInMasterTracker
 *
 * Re-applies formulas to columns G–J for all rows in the Master Tracker.
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
    sheet.getRange(i, 10).setFormula(`=IF(AND(D${i}=0, E${i}<0), ABS(E${i}), 0)`);
  }

  SpreadsheetApp.getUi().alert("✅ Formulas in columns G–J have been updated and patched.");
}

/**
 * insertNewClientIntoDirectory
 *
 * Prompts the user to enter a new client, adds to the Client Directory.
 */
function insertNewClientIntoDirectory() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Client Directory");

  const clientPrompt = ui.prompt("New Client", "Enter the client's domain (e.g., example.com):", ui.ButtonSet.OK_CANCEL);
  if (clientPrompt.getSelectedButton() !== ui.Button.OK) return;

  const clientName = clientPrompt.getResponseText().trim();
  if (!clientName) return ui.alert("No client name provided.");

  const statusPrompt = ui.prompt("Client Status", "Enter status (Active, Inactive, Transitioning):", ui.ButtonSet.OK_CANCEL);
  if (statusPrompt.getSelectedButton() !== ui.Button.OK) return;

  const status = statusPrompt.getResponseText().trim();
  if (!status) return ui.alert("No status provided.");

  const lastRow = sheet.getLastRow();
  const existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  if (existing.includes(clientName)) return ui.alert(`⚠️ ${clientName} already exists.`);

  sheet.appendRow([clientName, "", "", status]);
  ui.alert(`✅ ${clientName} added to the Client Directory.`);
}

/**
 * insertAllMissingClients
 *
 * Adds clients from Client Directory to the bottom of the Master Tracker if missing.
 */
function insertAllMissingClients() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const directorySheet = ss.getSheetByName("Client Directory");
  const masterSheet = ss.getSheetByName("Master Tracker");

  const directoryData = directorySheet.getRange(2, 1, directorySheet.getLastRow() - 1, 4).getValues();
  const activeClients = directoryData.filter(row => row[3]?.toLowerCase() === "active").map(row => row[0]);

  const existingClients = masterSheet.getRange(2, 2, masterSheet.getLastRow() - 1).getValues().flat();
  const missingClients = activeClients.filter(name => !existingClients.includes(name));
  const monthLabel = masterSheet.getRange(2, 1).getValue();

  missingClients.forEach(name => {
    masterSheet.appendRow([monthLabel, name]);
  });

  SpreadsheetApp.getUi().alert(`✅ ${missingClients.length} missing client(s) added to the bottom of the Master Tracker.`);
}

/**
 * clearDocAndFolderLinks
 *
 * Clears Columns N and T of the Master Tracker.
 */
function clearDocAndFolderLinks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Tracker");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.getRange(2, 14, lastRow - 1).clearContent(); // Column N
  sheet.getRange(2, 20, lastRow - 1).clearContent(); // Column T
  SpreadsheetApp.getUi().alert("🧹 Cleared doc and folder links from columns N and T.");
}

/**
 * onOpen
 *
 * Loads Client Tools menu in the Google Sheet UI.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🗂 Client Tools')
    .addItem('Run Monthly Rollover & Docs', 'monthlyRolloverAndCreateDocs')
    .addItem('Reset Master Tracker Formulas', 'resetFormulasInMasterTracker')
    .addItem('Insert New Client into Directory', 'insertNewClientIntoDirectory')
    .addItem('Insert All Missing Clients into Master Tracker', 'insertAllMissingClients')
    .addItem('Clear Doc & Folder Links', 'clearDocAndFolderLinks')
    .addToUi();
}