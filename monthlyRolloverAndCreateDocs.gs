// === Main Monthly Automation Script ===
/**
/**
 * monthlyRolloverAndCreateDocs
 *
 * Runs the end-of-month rollover process:
 * - Identifies the prior month
 * - Iterates over each Active client in "Master Tracker"
 * - Creates and logs a support summary Google Doc
 * - Dynamically calculates uncovered overage from Column I (Remaining Block)
 */
function monthlyRolloverAndCreateDocs() {
  // === DRY RUN mode: true = test only (no docs created), false = live run
  const DRY_RUN = false;

  // === Spreadsheet and UI references
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // === Primary data sheets
  const masterSheet = ss.getSheetByName("Master Tracker");
  const directorySheet = ss.getSheetByName("Client Directory");

  // === Timezone and parent folder setup for Google Docs
  const timeZone = ss.getSpreadsheetTimeZone();
  const parentFolderId = '1UI4zQ_YIEWWJT0kSP2x8EaQlue303Xl-'; // Root folder ID for client folders
  const parentFolder = DriveApp.getFolderById(parentFolderId);

  // === Determine previous month and timestamp
  const today = new Date();
  today.setMonth(today.getMonth() - 1);
  const monthLabel = Utilities.formatDate(today, timeZone, "MMMM yyyy");
  const timestamp = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss");

  // === Ensure Document Summary sheet exists and clear it
  const summarySheet = ss.getSheetByName("Document Summary") || ss.insertSheet("Document Summary");
  summarySheet.clear();
  summarySheet.appendRow(["Client Name", "Email", "Doc URL", "Timestamp"]);

  // === Load all rows from Master Tracker (excluding headers)
  const dataRange = masterSheet.getDataRange();
  const data = dataRange.getValues();
  const rows = data.slice(1);

  let createdCount = 0;

  // === Iterate through each row/client
  rows.forEach((row, i) => {
    const rowNum = i + 2;

    // === Destructure key columns for clarity
    const [
      , // A - Month
      clientName,       // B
      , ,               // C–D
      blockAvailable,   // E
      hoursUsed,        // F
      overageBeyondPlan,// G
      blockUsed,        // H
      remainingBlock,   // I
      , ,               // J–K
      clientEmail,      // L (was column M)
      firstName,        // M (was column O)
      ,                 // N (Last Name)
      status,           // O (was Q)
      domainExpire,     // P
      accessToGA        // Q
    ] = row;

    const normalizedStatus = typeof status === "string" ? status.trim().toLowerCase() : "";
    const trimmedName = typeof clientName === "string" ? clientName.trim() : "";

    // === Skip if client is not active or has no name
    if (!trimmedName || normalizedStatus !== "active") {
      console.log(`⚠️ Skipping row ${rowNum}: Name="${trimmedName}" | Status="${normalizedStatus}"`);
      return;
    }

    const docName = `${monthLabel} - ${clientName}`;
    const clientFolder = getOrCreateClientFolder(parentFolder, clientName);

    // === Trash any previous docs with same name
    const existingFiles = clientFolder.getFilesByName(docName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    // === DRY RUN mode: log and skip creation
    if (DRY_RUN) {
      console.log(`🟡 Dry run - would have created doc for ${clientName}`);
      return;
    }

    // === Create the Google Doc
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();

    // === Add RadiateU logo
    const logoBlob = DriveApp.getFileById("1fW300SGxEFVFvndaLkkWz3_O7L3BOq84").getBlob();
    body.appendImage(logoBlob).setWidth(250);

    // === Compose content
    body.appendParagraph("\nHello,");
    body.appendParagraph(`Here’s your monthly support summary for ${clientName} – ${monthLabel}:\n`);
    body.appendParagraph(`Block Hours Applied: ${blockUsed || 0}`);
    body.appendParagraph(`Remaining Block Balance: ${remainingBlock || 0}`);

    // Dynamically calculate uncovered overage if balance is negative
    const uncovered = typeof remainingBlock === "number" && remainingBlock < 0
      ? Math.abs(remainingBlock)
      : 0;
    body.appendParagraph(`Overage Hours (Uncovered): ${uncovered}`);

    body.appendParagraph("\nIf you need additional support hours, visit https://radiateu.com/request-support-time.");
    body.appendParagraph("\nFor our clients on a monthly plan:");
    body.appendParagraph("🔐 Domain Expiration: " + (domainExpire || "N/A"));
    body.appendParagraph("📊 Access to Google Analytics: " + (accessToGA || "N/A"));
    body.appendParagraph("\nIf you have any questions, feel free to reply here or send a message to support@radiateu.com.");
    body.appendParagraph("\n*If you have trouble accessing your support summary, let us know and we’ll send you a PDF version.*");

    doc.saveAndClose();

    // === Move to client folder and insert link
    const file = DriveApp.getFileById(doc.getId());
    file.moveTo(clientFolder);
    const docUrl = doc.getUrl();
    const hyperlink = `=HYPERLINK("${docUrl}", "Open Doc")`;
    masterSheet.getRange(rowNum, 14).setFormula(hyperlink);  // Column N

    // === Log to Document Summary tab
    summarySheet.appendRow([clientName, clientEmail || "N/A", docUrl, timestamp]);
    console.log(`✅ Created support doc for ${clientName}`);
    createdCount++;
  });

  // === Final alert
  ui.alert(`✅ ${createdCount} support summaries were created.`);
}

/**
 * getOrCreateClientFolder
 *
 * Returns an existing folder for a client or creates a new one if missing.
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

    // I - Block Remaining
    sheet.getRange(i, 9).setFormula(`=IF(H${i}="", "", E${i}-H${i})`);

    // J - Block Deficit Warning (hrs)
    sheet.getRange(i, 10).setFormula(`=IF(AND(D${i}=0, E${i}<0), ABS(E${i}), 0)`);
  }

  SpreadsheetApp.getUi().alert("✅ Formulas in columns G–J have been updated and patched.");
}

/**
 * insertAllMissingClients
 *
 * Inserts any clients from the Helper Sheet into the Master Tracker
 * if they’re missing from the current list.
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

  const monthLabel = masterSheet.getRange(2, 1).getValue(); // Use existing month

  missingClients.forEach(name => {
    masterSheet.appendRow([monthLabel, name]);
  });

  SpreadsheetApp.getUi().alert(`✅ ${missingClients.length} missing client(s) added to the Master Tracker.`);
}

/**
 * insertNewClientIntoDirectory
 *
 * Prompts the user for a new client name and status,
 * then adds them to the Client Directory if not already present.
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

  sheet.appendRow([clientName, "", "", "", status]);
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