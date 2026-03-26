/**
 * SOLOC Annual Contributions Report Generator
 *
 * Bound to the 2025 Individual Contributions spreadsheet.
 * Reads each donor row, generates a PDF from the Google Doc template,
 * and emails it to the donor from finance@soloc.net.
 *
 * Spreadsheet: 12qBk_Of1_x110T1O08812oXX8cBpUazBolX7KlKISOY
 * Template:    1n-1fjDPumhuRZ-tMbLxla7g03_PibW6ya7j55txTnik
 */

var ANNUAL_CONFIG = {
  spreadsheetId: "12qBk_Of1_x110T1O08812oXX8cBpUazBolX7KlKISOY",
  templateId: "1n-1fjDPumhuRZ-tMbLxla7g03_PibW6ya7j55txTnik",
  sheetName: null,  // uses first sheet
  headerRow: 4,     // row 4 has the column headers; data starts at row 5
  receiptPrefix: "SOL-2025-",
  receiptPad: 4,
  email: {
    fromAlias: "finance@soloc.net",
    fromName: "Seeds of Love Online Community, Inc."
  },
  columns: {
    name: 0,          // A
    january: 1,       // B
    december: 12,     // M
    annualTotal: 13,  // N
    email: 14,        // O
    receiptNo: 15,    // P
    dateIssued: 16,   // Q
    emailSentAt: 17   // R  (script adds this column)
  }
};

/**
 * Run once to add the Email Sent At tracking column if missing.
 * Email is already in column O, Receipt No. in P, Date Issued in Q.
 */
function setupAnnualReport() {
  var sheet = getContributionsSheet_();
  var headers = getContributionHeaders_(sheet);
  var lastCol = sheet.getLastColumn();

  // Check if Email Sent At column exists
  var sentIdx = findHeaderIndex_(headers, "Email Sent At");
  if (sentIdx === -1) {
    sheet.getRange(ANNUAL_CONFIG.headerRow, lastCol + 1).setValue("Email Sent At");
  }

  SpreadsheetApp.flush();
  Logger.log("Setup complete. Email Sent At column ready.");
}

/**
 * Generate and email annual contribution receipts for ALL donors
 * that have an email address, annual total > 0, and have not been sent yet.
 */
function sendAllAnnualReceipts() {
  var sheet = getContributionsSheet_();
  var headers = getContributionHeaders_(sheet);
  var data = getContributionData_(sheet);
  var cols = resolveColumns_(headers);
  var sent = 0;
  var skipped = 0;
  var errors = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var sheetRow = dataIndexToRow_(i);
    var name = String(row[cols.name] || "").trim();
    var email = String(row[cols.email] || "").trim();
    var annualTotal = Number(row[cols.annualTotal] || 0);
    var alreadySent = String(row[cols.emailSentAt] || "").trim();

    if (!name || !email || annualTotal <= 0) {
      skipped++;
      continue;
    }
    if (alreadySent) {
      skipped++;
      continue;
    }

    try {
      var receiptNo = String(row[cols.receiptNo] || "").trim();
      if (!receiptNo) {
        receiptNo = nextAnnualReceiptNo_(data, cols);
        sheet.getRange(sheetRow, cols.receiptNo + 1).setValue(receiptNo);
      }

      var dateIssued = String(row[cols.dateIssued] || "").trim();
      if (!dateIssued) {
        dateIssued = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
        sheet.getRange(sheetRow, cols.dateIssued + 1).setValue(dateIssued);
      }

      var mergeData = {
        RECEIPT_NO: receiptNo,
        DATE_ISSUED: dateIssued,
        DONOR_NAME: name,
        ANNUAL_TOTAL: formatUsd_(annualTotal)
      };

      var pdfBlob = createAnnualPdf_(mergeData, name);
      emailAnnualReceipt_(email, name, receiptNo, annualTotal, pdfBlob);

      sheet.getRange(sheetRow, cols.emailSentAt + 1).setValue(new Date());
      sent++;
      Logger.log("Sent receipt " + receiptNo + " to " + email);
    } catch (e) {
      errors.push(name + ": " + e.message);
      Logger.log("ERROR for " + name + ": " + e.message);
    }
  }

  SpreadsheetApp.flush();
  Logger.log("Done. Sent: " + sent + ", Skipped: " + skipped + ", Errors: " + errors.length);
  if (errors.length > 0) {
    Logger.log("Errors:\n" + errors.join("\n"));
  }
}

/**
 * Send receipt for the currently selected row only.
 * Useful for testing or resending a single donor.
 */
function sendActiveRowReceipt() {
  var sheet = getContributionsSheet_();
  var activeRow = sheet.getActiveRange().getRow();
  var firstDataRow = ANNUAL_CONFIG.headerRow + 1;
  if (activeRow < firstDataRow) {
    throw new Error("Select a donor row (row " + firstDataRow + " or later).");
  }

  var headers = getContributionHeaders_(sheet);
  var data = getContributionData_(sheet);
  var cols = resolveColumns_(headers);
  var dataIndex = activeRow - firstDataRow;
  var row = data[dataIndex];

  var name = String(row[cols.name] || "").trim();
  var email = String(row[cols.email] || "").trim();
  var annualTotal = Number(row[cols.annualTotal] || 0);

  if (!name) throw new Error("No donor name in row " + activeRow);
  if (!email) throw new Error("No email for " + name + " in row " + activeRow);
  if (annualTotal <= 0) throw new Error("Annual total is $0 for " + name);

  var receiptNo = String(row[cols.receiptNo] || "").trim();
  if (!receiptNo) {
    receiptNo = nextAnnualReceiptNo_(data, cols);
    sheet.getRange(activeRow, cols.receiptNo + 1).setValue(receiptNo);
  }

  var dateIssued = String(row[cols.dateIssued] || "").trim();
  if (!dateIssued) {
    dateIssued = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.getRange(activeRow, cols.dateIssued + 1).setValue(dateIssued);
  }

  var mergeData = {
    RECEIPT_NO: receiptNo,
    DATE_ISSUED: dateIssued,
    DONOR_NAME: name,
    ANNUAL_TOTAL: formatUsd_(annualTotal)
  };

  var pdfBlob = createAnnualPdf_(mergeData, name);
  emailAnnualReceipt_(email, name, receiptNo, annualTotal, pdfBlob);

  sheet.getRange(activeRow, cols.emailSentAt + 1).setValue(new Date());
  SpreadsheetApp.flush();
  Logger.log("Sent receipt " + receiptNo + " to " + email + " for " + name);
}

/**
 * Reset the "Email Sent At" column for the selected row so it can be resent.
 */
function resetActiveRowSentStatus() {
  var sheet = getContributionsSheet_();
  var activeRow = sheet.getActiveRange().getRow();
  var firstDataRow = ANNUAL_CONFIG.headerRow + 1;
  if (activeRow < firstDataRow) {
    throw new Error("Select a donor row (row " + firstDataRow + " or later).");
  }
  var headers = getContributionHeaders_(sheet);
  var cols = resolveColumns_(headers);
  sheet.getRange(activeRow, cols.emailSentAt + 1).setValue("");
  Logger.log("Reset sent status for row " + activeRow);
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

function getContributionsSheet_() {
  var ss = SpreadsheetApp.getActive() ||
           SpreadsheetApp.openById(ANNUAL_CONFIG.spreadsheetId);
  var sheet;
  if (ANNUAL_CONFIG.sheetName) {
    sheet = ss.getSheetByName(ANNUAL_CONFIG.sheetName);
  } else {
    sheet = ss.getSheets()[0];
  }
  return sheet;
}

/**
 * Returns headers array from the configured header row (row 4).
 */
function getContributionHeaders_(sheet) {
  return sheet.getRange(ANNUAL_CONFIG.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * Returns all data rows starting after the header row.
 * Each element is a row array; index 0 = first data row (row 5).
 */
function getContributionData_(sheet) {
  var startRow = ANNUAL_CONFIG.headerRow + 1;
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];
  var numRows = lastRow - startRow + 1;
  return sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
}

/**
 * Converts a data-array index (0-based) to the actual sheet row number.
 */
function dataIndexToRow_(dataIndex) {
  return ANNUAL_CONFIG.headerRow + 1 + dataIndex;
}

function resolveColumns_(headers) {
  var cols = {
    name: 0,
    annualTotal: -1,
    receiptNo: -1,
    dateIssued: -1,
    email: -1,
    emailSentAt: -1
  };

  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || "").toLowerCase().trim();
    if (h.indexOf("annual total") !== -1) cols.annualTotal = i;
    else if (h === "email" || h === "email address") cols.email = i;
    else if (h.indexOf("receipt no") !== -1) cols.receiptNo = i;
    else if (h.indexOf("date issued") !== -1) cols.dateIssued = i;
    else if (h.indexOf("email sent") !== -1) cols.emailSentAt = i;
  }

  if (cols.annualTotal === -1) throw new Error("Cannot find ANNUAL TOTAL column");
  if (cols.email === -1) throw new Error("Cannot find Email column. Run setupAnnualReport() first.");
  if (cols.emailSentAt === -1) throw new Error("Cannot find Email Sent At column. Run setupAnnualReport() first.");

  // Receipt No and Date Issued may not exist yet in some sheets; default to after annual total
  if (cols.receiptNo === -1) cols.receiptNo = cols.annualTotal + 1;
  if (cols.dateIssued === -1) cols.dateIssued = cols.annualTotal + 2;

  return cols;
}

function findHeaderIndex_(headers, target) {
  var needle = target.toLowerCase().trim();
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").toLowerCase().trim() === needle) {
      return i;
    }
  }
  return -1;
}

function nextAnnualReceiptNo_(data, cols) {
  var maxNum = 0;
  for (var i = 0; i < data.length; i++) {
    var existing = String(data[i][cols.receiptNo] || "");
    var match = existing.match(/(\d+)$/);
    if (match) {
      var num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  }
  var next = maxNum + 1;
  var padded = String(next);
  while (padded.length < ANNUAL_CONFIG.receiptPad) {
    padded = "0" + padded;
  }
  return ANNUAL_CONFIG.receiptPrefix + padded;
}

function createAnnualPdf_(mergeData, donorName) {
  var templateFile = DriveApp.getFileById(ANNUAL_CONFIG.templateId);
  var outputName = "Annual Contribution Receipt 2025 - " + donorName;
  var copyFile = templateFile.makeCopy(outputName);
  var doc = DocumentApp.openById(copyFile.getId());
  var body = doc.getBody();

  for (var key in mergeData) {
    if (mergeData.hasOwnProperty(key)) {
      body.replaceText("\\{\\{" + key + "\\}\\}", String(mergeData[key] || ""));
    }
  }

  doc.saveAndClose();
  Utilities.sleep(1000);

  var pdfBlob = copyFile.getAs(MimeType.PDF).setName(outputName + ".pdf");
  copyFile.setTrashed(true);
  return pdfBlob;
}

function emailAnnualReceipt_(toEmail, donorName, receiptNo, annualTotal, pdfBlob) {
  var subject = "SOLOC 2025 Annual Contribution Receipt - " + receiptNo;

  var plainBody =
    "Dear " + donorName + ",\n\n" +
    "Thank you for your generous contributions to Seeds of Love Online Community in 2025.\n\n" +
    "Attached is your Annual Contribution Acknowledgment for tax year 2025.\n\n" +
    "Receipt Number: " + receiptNo + "\n" +
    "Total Contributions: $" + formatUsd_(annualTotal) + "\n\n" +
    "Please retain this receipt for your tax records.\n\n" +
    "If you have any questions, please contact us at +1-623-2177823 or finance@soloc.net.\n\n" +
    "With gratitude,\n" +
    "Seeds of Love Online Community, Inc.\n";

  var htmlBody =
    "<p>Dear " + escapeHtml_(donorName) + ",</p>" +
    "<p>Thank you for your generous contributions to Seeds of Love Online Community in 2025.</p>" +
    "<p>Attached is your <strong>Annual Contribution Acknowledgment</strong> for tax year 2025.</p>" +
    "<p><strong>Receipt Number:</strong> " + escapeHtml_(receiptNo) + "<br>" +
    "<strong>Total Contributions:</strong> $" + escapeHtml_(formatUsd_(annualTotal)) + "</p>" +
    "<p>Please retain this receipt for your tax records.</p>" +
    "<p>If you have any questions, please contact us at +1-623-2177823 or " +
    "<a href=\"mailto:finance@soloc.net\">finance@soloc.net</a>.</p>" +
    "<p>With gratitude,<br>" +
    "Seeds of Love Online Community, Inc.</p>";

  var options = {
    name: ANNUAL_CONFIG.email.fromName,
    htmlBody: htmlBody,
    attachments: [pdfBlob]
  };

  if (ANNUAL_CONFIG.email.fromAlias) {
    options.from = ANNUAL_CONFIG.email.fromAlias;
    options.replyTo = ANNUAL_CONFIG.email.fromAlias;
  }

  GmailApp.sendEmail(toEmail, subject, plainBody, options);
}

function formatUsd_(amount) {
  var n = Number(amount || 0);
  return n.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function escapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
