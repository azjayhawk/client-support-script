/**************************************
 * SAFE MONTHLY MODULE – RadiateU
 * ------------------------------------
 * This module handles the generation
 * of monthly Google Docs for each client.
 *
 * It supports:
 *  - Document creation/updating
 *  - Folder management
 *  - Monthly summaries
 *  - Client deduplication
 *
 * Designed for non-coders: Comments
 * explain each function clearly.
 **************************************/

// --- CONFIGURATION CONSTANTS ---

// Sheet/tab names
const MT_SHEET = 'Master Tracker';

// Template row in the Master Tracker for formulas/formatting
const TEMPLATE_ROW_IDX = 2;
const FIRST_DATA_ROW_IDX = 2;

// Report columns in Master Tracker (where links to folder/docs are stored)
const REPORT_FOLDER_COL = 18; // Column R
const REPORT_DOC_COL = 19;    // Column S

// Hidden tracking columns (not visible to user but used by script)
const KEY_HIDDEN_COL_NAME = "KEY";
const DOC_ID_HIDDEN_COL_NAME = "DOC_ID";

// Branding resources
const LOGO_FILE_ID = '1fW300SGxEFVFvndaLkkWz3_O7L3BOq84';
const PARENT_FOLDER_ID = '1UI4zQ_YIEWWJT0kSP2x8EaQlue303Xl-'; // Master parent folder for all clients

// --- UTILITY FUNCTIONS ---

// Creates/ensures hidden columns for internal tracking (used to avoid duplicates)
function ensureHiddenColumns_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(MT_SHEET);
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  let keyIndex = header.indexOf(KEY_HIDDEN_COL_NAME) + 1;
  let docIdIndex = header.indexOf(DOC_ID_HIDDEN_COL_NAME) + 1;

  if (!keyIndex) {
    keyIndex = sh.getLastColumn() + 1;
    sh.getRange(1, keyIndex).setValue(KEY_HIDDEN_COL_NAME);
  }

  if (!docIdIndex) {
    docIdIndex = sh.getLastColumn() + 1;
    sh.getRange(1, docIdIndex).setValue(DOC_ID_HIDDEN_COL_NAME);
  }

  return [keyIndex, docIdIndex];
}

// Converts current date to a label like “August 2025”
function currentMonthLabel_(tz = "America/Phoenix") {
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1);
  return Utilities.formatDate(lastMonth, tz, "MMMM yyyy");
}

// Normalizes a string into a simple key (e.g., lowercased, no special characters)
function buildClientKey_(clientName) {
  return clientName.toLowerCase().replace(/[^\w]+/g, "");
}

// Build an index of first occurrence of each client (by key) in the Master Tracker
function buildClientIndex_(sh, keyCol) {
  const data = sh.getRange(FIRST_DATA_ROW_IDX, keyCol, sh.getLastRow() - 1).getValues();
  const index = {};
  data.forEach((row, i) => {
    const key = row[0];
    if (key && !(key in index)) {
      index[key] = i + FIRST_DATA_ROW_IDX;
    }
  });
  return index;
}

// Copy formulas/formatting from template row to a new client row
function copyTemplateRowTo_(sh, targetRow) {
  const template = sh.getRange(TEMPLATE_ROW_IDX, 1, 1, sh.getLastColumn());
  const dest = sh.getRange(targetRow, 1, 1, sh.getLastColumn());
  template.copyTo(dest, { contentsOnly: false });
}

// Makes sure a client folder exists in Drive (and creates it if not)
function getOrCreateClientFolder_(clientName) {
  const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
  const folders = parentFolder.getFoldersByName(clientName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(clientName);
}

// Tries to find a doc by name inside a folder
function findMonthlyDocInFolder_(folder, docName) {
  const files = folder.getFilesByName(docName);
  return files.hasNext() ? files.next() : null;
}

// Main function that fills out the Google Doc body
function ensureAndFillMonthlyDocFast_(doc, row, monthLabel, clientName) {
  const docBody = doc.getBody();
  docBody.clear();

  const HOURS_USED_COL = 5;
  const REMAINING_BLOCK_COL = 8;
  const OVERAGE_HOURS_COL = 9;
  const DOMAIN_EXPIRY_COL = 16;
  const GA_ACCESS_COL = 17;
  const CLIENT_FIRST_NAME_COL = 13;

  const hoursUsed = Number(row[HOURS_USED_COL] || 0).toFixed(1);
  const remainingBlock = Number(row[REMAINING_BLOCK_COL] || 0).toFixed(1);
  const overageUncovered = Number(row[OVERAGE_HOURS_COL] || 0).toFixed(1);
  const domainExpiration = row[DOMAIN_EXPIRY_COL] || "Not available";
  const accessToGA = (row[GA_ACCESS_COL] || "").toString().toLowerCase() === "yes" ? "Yes" : "No";
  const firstName = row[CLIENT_FIRST_NAME_COL] || "there";

  const logoImage = DriveApp.getFileById(LOGO_FILE_ID).getBlob();
  docBody.appendImage(logoImage).setWidth(120);

  docBody.appendParagraph(`Hello ${firstName},`)
    .setSpacingAfter(16)
    .setFontSize(11);

  docBody.appendParagraph(`Here’s your monthly website support summary for **${clientName}** – **${monthLabel}**:`)
    .setSpacingAfter(12)
    .setFontSize(11);

  docBody.appendParagraph(`📊 **Support Summary**`)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  docBody.appendParagraph(
    `• **Support Hours Logged:** ${hoursUsed} hrs\n` +
    `• **Remaining Block Hours:** ${remainingBlock} hrs\n` +
    `• **Overage Hours (Uncovered):** ${overageUncovered} hrs`
  ).setFontSize(11).setSpacingAfter(12);

  docBody.appendParagraph(`Need extra time? You can request additional support hours here:\n👉 https://radiateu.com/request-support-time`)
    .setFontSize(10)
    .setSpacingAfter(18);

  docBody.appendParagraph(`🔐 **Account Info**`)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  docBody.appendParagraph(
    `• 🔐 **Domain Expiration:** ${domainExpiration}\n` +
    `• 📊 **Access to Google Analytics:** ${accessToGA}`
  ).setFontSize(11).setSpacingAfter(16);

  docBody.appendParagraph(`If you have any questions, just reply to this message or contact us at support@radiateu.com. We’re happy to help!`)
    .setFontSize(10)
    .setSpacingAfter(10);

  docBody.appendParagraph(`*Having trouble accessing your support summary? Let us know and we’ll send you a PDF version.*`)
    .setFontSize(9)
    .setItalic(true);

  return doc;
}

// This function removes duplicate client rows in the Master Tracker (keeps the first)
function dedupeByClientKeepFirst_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(MT_SHEET);
  const [keyIndex] = ensureHiddenColumns_();
  const index = buildClientIndex_(sh, keyIndex);

  const rowsToDelete = [];
  const seen = new Set();

  Object.entries(index).forEach(([key, row]) => seen.add(row));
  for (let i = sh.getLastRow(); i >= FIRST_DATA_ROW_IDX; i--) {
    if (!seen.has(i)) rowsToDelete.push(i);
  }
  rowsToDelete.forEach(row => sh.deleteRow(row));
}

// Adds the 🛡️ Safe Tools menu to Sheets UI
function onOpen_AddSafeItems() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🛡️ Client Tools (Safe)")
    .addItem("📄 Run Monthly (Rolling rows)", "monthlyRolloverAndCreateDocsSafe")
    .addItem("🧹 Dedupe by Client (keep first)", "dedupeByClientKeepFirst_")
    .addToUi();
}

// MAIN FUNCTION – creates or updates the monthly summary docs
function monthlyRolloverAndCreateDocsSafe() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(MT_SHEET);
  const data = sheet.getRange(FIRST_DATA_ROW_IDX, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const [keyColIdx, docIdColIdx] = ensureHiddenColumns_();
  const monthLabel = currentMonthLabel_();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const planType = row[2];  // Column C
    const status = row[15];   // Column O
    const clientName = row[1]; // Column B

    // Skip Hosting, Inactive, and Transitioning clients
    if (!clientName || planType === "Hosting" || ["Inactive", "Transitioning"].includes(status)) continue;

    const key = buildClientKey_(clientName);
    const folder = getOrCreateClientFolder_(clientName);
    const docName = `Support Summary – ${monthLabel}`;
    let doc = null;

    const existingDocId = sheet.getRange(i + FIRST_DATA_ROW_IDX, docIdColIdx).getValue();
    if (existingDocId) {
      try {
        doc = DocumentApp.openById(existingDocId);
      } catch (e) {
        doc = null;
      }
    }

    if (!doc) {
      const existing = findMonthlyDocInFolder_(folder, docName);
      doc = existing
        ? DocumentApp.openById(existing.getId())
        : DocumentApp.create(docName);
    }

    ensureAndFillMonthlyDocFast_(doc, row, monthLabel, clientName);
    const docUrl = doc.getUrl();

    sheet.getRange(i + FIRST_DATA_ROW_IDX, REPORT_FOLDER_COL).setFormula(`=HYPERLINK("${folder.getUrl()}", "📂 Folder")`);
    sheet.getRange(i + FIRST_DATA_ROW_IDX, REPORT_DOC_COL).setFormula(`=HYPERLINK("${docUrl}", "📄 Summary")`);
    sheet.getRange(i + FIRST_DATA_ROW_IDX, keyColIdx).setValue(key);
    sheet.getRange(i + FIRST_DATA_ROW_IDX, docIdColIdx).setValue(doc.getId());
  }
}