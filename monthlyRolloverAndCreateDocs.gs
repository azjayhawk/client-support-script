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
 * onOpen
 * Adds the Client Tools menu when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🗂 Client Tools')
    .addItem('Run Monthly Rollover & Docs', 'monthlyRolloverAndCreateDocs')
    .addItem('Fix Master Tracker Lookup Formulas (F, H, I, M–S)', 'updateMasterTrackerFormulas')
    .addItem('Insert New Client into Directory', 'insertNewClientIntoDirectory')
    .addItem('Insert All Missing Clients into Master Tracker', 'insertAllMissingClients')
    .addItem('Clear Doc & Folder Links', 'clearDocAndFolderLinks')
    .addItem('Sort Master Tracker A–Z', 'sortMasterTrackerAZ')
    .addToUi();
}