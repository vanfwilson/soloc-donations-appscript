// === United States & Canada ===================================================
var ANNUAL_CONFIG_US = {
  spreadsheetId: "12qBk_Of1_x110T1O08812oXX8cBpUazBolX7KlKISOY",
  templateId: "1n-1fjDPumhuRZ-tMbLxla7g03_PibW6ya7j55txTnik",
  sheetName: "annual donations 2025",
  headerRow: 4,
  useFixedColumns: false,
  receiptPrefix: "SOL-2025-",
  receiptPad: 4,
  email: {
    fromAlias: "finance@soloc.net",
    fromName: "Seeds of Love Online Community, Inc.",
    bccArchive: "finance@soloc.net"
  },
  archiveFolderId: "1cJK5PWnIMy_gN_cO-5g-muSEJbdZXE3p",
  columns: {
    name: 0,
    annualTotal: 13,
    email: 14,
    receiptNo: 15,
    dateIssued: 16,
    emailSentAt: 17
  }
};

// === Philippines & Euro-Asia ==================================================
// Column layout (0-based): A=0 B=1 C=2 D=3 E=4 F=5 G=6 H=7 I=8
var ANNUAL_CONFIG_PH = {
  spreadsheetId: "12qBk_Of1_x110T1O08812oXX8cBpUazBolX7KlKISOY",
  templateId: "1n-1fjDPumhuRZ-tMbLxla7g03_PibW6ya7j55txTnik",
  sheetName: "phil & euroasia 2025",
  headerRow: 2,
  useFixedColumns: true,
  receiptPrefix: "SOL-2025-",
  receiptPad: 4,
  email: {
    fromAlias: "finance@soloc.net",
    fromName: "Seeds of Love Online Community, Inc.",
    bccArchive: "finance@soloc.net"
  },
  archiveFolderId: "1cJK5PWnIMy_gN_cO-5g-muSEJbdZXE3p",
  columns: {
    name: 0,
    email: 1,
    annualTotal: 6,
    receiptNo: 7,
    dateIssued: 8,
    emailSentAt: 9
  }
};

// === US & Canada entry points =================================================
function setupAnnualReport_US()        { setupAnnualReport_(ANNUAL_CONFIG_US); }
function sendAllAnnualReceipts_US()    { sendAllAnnualReceipts_(ANNUAL_CONFIG_US); }
function sendActiveRowReceipt_US()     { sendActiveRowReceipt_(ANNUAL_CONFIG_US); }
function resetActiveRowSentStatus_US() { resetActiveRowSentStatus_(ANNUAL_CONFIG_US); }

// === Philippines & Euro-Asia entry points =====================================
function setupAnnualReport_PH()        { setupAnnualReport_(ANNUAL_CONFIG_PH); }
function sendAllAnnualReceipts_PH()    { sendAllAnnualReceipts_(ANNUAL_CONFIG_PH); }
function sendActiveRowReceipt_PH()     { sendActiveRowReceipt_(ANNUAL_CONFIG_PH); }
function resetActiveRowSentStatus_PH() { resetActiveRowSentStatus_(ANNUAL_CONFIG_PH); }

// === Debug helpers ============================================================
function debugHeaders_US() { debugHeaders_(ANNUAL_CONFIG_US); }
function debugHeaders_PH() { debugHeaders_(ANNUAL_CONFIG_PH); }

function debugHeaders_(config) {
  var sheet = getContributionsSheet_(config);
  var headers = getContributionHeaders_(sheet, config);
  Logger.log("=== Headers in row " + config.headerRow + " of tab: " + sheet.getName() + " ===");
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").trim() !== "") {
      Logger.log("  Col " + (i + 1) + " [index " + i + "]: " + JSON.stringify(headers[i]));
    }
  }
}

// === Internal implementations =================================================

function setupAnnualReport_(config) {
  var sheet = getContributionsSheet_(config);
  if (config.useFixedColumns) {
    var cols = config.columns;
    if (cols.receiptNo >= 0)   sheet.getRange(config.headerRow, cols.receiptNo  + 1).setValue("Receipt No.");
    if (cols.dateIssued >= 0)  sheet.getRange(config.headerRow, cols.dateIssued + 1).setValue("Date Issued");
    if (cols.emailSentAt >= 0) sheet.getRange(config.headerRow, cols.emailSentAt + 1).setValue("Email Sent At");
  } else {
    var headers = getContributionHeaders_(sheet, config);
    var lastCol = sheet.getLastColumn();
    var sentIdx = findHeaderIndex_(headers, "Email Sent At");
    if (sentIdx === -1) {
      sheet.getRange(config.headerRow, lastCol + 1).setValue("Email Sent At");
    }
  }
  SpreadsheetApp.flush();
  Logger.log("Setup complete. Tracking columns ready.");
}

function sendAllAnnualReceipts_(config) {
  var sheet = getContributionsSheet_(config);
  var headers = getContributionHeaders_(sheet, config);
  var data = getContributionData_(sheet, config);
  var cols = resolveColumns_(headers, config);
  var sent = 0, skipped = 0, errors = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var sheetRow = dataIndexToRow_(i, config);
    var name = String(row[cols.name] || "").trim();
    var email = String(row[cols.email] || "").trim();
    var annualTotal = Number(row[cols.annualTotal] || 0);
    var alreadySent = String(row[cols.emailSentAt] || "").trim();

    if (!name || !email || annualTotal <= 0) { skipped++; continue; }
    if (alreadySent) { skipped++; continue; }

    try {
      var receiptNo = String(row[cols.receiptNo] || "").trim();
      if (!receiptNo) {
        receiptNo = nextAnnualReceiptNo_(data, cols, config);
        sheet.getRange(sheetRow, cols.receiptNo + 1).setValue(receiptNo);
      }
      var dateIssued = String(row[cols.dateIssued] || "").trim();
      if (!dateIssued) {
        dateIssued = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
        sheet.getRange(sheetRow, cols.dateIssued + 1).setValue(dateIssued);
      }
      var mergeData = { RECEIPT_NO: receiptNo, DATE_ISSUED: dateIssued, DONOR_NAME: name, ANNUAL_TOTAL: formatUsd_(annualTotal) };
      var pdfBlob = createAnnualPdf_(mergeData, name, config);
      emailAnnualReceipt_(email, name, receiptNo, annualTotal, pdfBlob, config);
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
  if (errors.length > 0) { Logger.log("Errors:\n" + errors.join("\n")); }
}

function sendActiveRowReceipt_(config) {
  var sheet = getContributionsSheet_(config);
  var activeRow = sheet.getActiveRange().getRow();
  var firstDataRow = config.headerRow + 1;
  if (activeRow < firstDataRow) {
    throw new Error("Select a donor row (row " + firstDataRow + " or later).");
  }
  var headers = getContributionHeaders_(sheet, config);
  var data = getContributionData_(sheet, config);
  var cols = resolveColumns_(headers, config);
  var row = data[activeRow - firstDataRow];

  var name = String(row[cols.name] || "").trim();
  var email = String(row[cols.email] || "").trim();
  var annualTotal = Number(row[cols.annualTotal] || 0);

  if (!name) throw new Error("No donor name in row " + activeRow);
  if (!email) throw new Error("No email for " + name + " in row " + activeRow);
  if (annualTotal <= 0) throw new Error("Annual total is $0 for " + name);

  var receiptNo = String(row[cols.receiptNo] || "").trim();
  if (!receiptNo) {
    receiptNo = nextAnnualReceiptNo_(data, cols, config);
    sheet.getRange(activeRow, cols.receiptNo + 1).setValue(receiptNo);
  }
  var dateIssued = String(row[cols.dateIssued] || "").trim();
  if (!dateIssued) {
    dateIssued = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.getRange(activeRow, cols.dateIssued + 1).setValue(dateIssued);
  }
  var mergeData = { RECEIPT_NO: receiptNo, DATE_ISSUED: dateIssued, DONOR_NAME: name, ANNUAL_TOTAL: formatUsd_(annualTotal) };
  var pdfBlob = createAnnualPdf_(mergeData, name, config);
  emailAnnualReceipt_(email, name, receiptNo, annualTotal, pdfBlob, config);
  sheet.getRange(activeRow, cols.emailSentAt + 1).setValue(new Date());
  SpreadsheetApp.flush();
  Logger.log("Sent receipt " + receiptNo + " to " + email + " for " + name);
}

function resetActiveRowSentStatus_(config) {
  var sheet = getContributionsSheet_(config);
  var activeRow = sheet.getActiveRange().getRow();
  var firstDataRow = config.headerRow + 1;
  if (activeRow < firstDataRow) {
    throw new Error("Select a donor row (row " + firstDataRow + " or later).");
  }
  var headers = getContributionHeaders_(sheet, config);
  var cols = resolveColumns_(headers, config);
  sheet.getRange(activeRow, cols.emailSentAt + 1).setValue("");
  Logger.log("Reset sent status for row " + activeRow);
}

function getContributionsSheet_(config) {
  var ss = SpreadsheetApp.getActive() ||
           SpreadsheetApp.openById(config.spreadsheetId);
  if (!config.sheetName) return ss.getSheets()[0];
  var sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) {
    var available = ss.getSheets().map(function(s) { return s.getName(); }).join(", ");
    throw new Error("Tab \"" + config.sheetName + "\" not found. Available tabs: " + available);
  }
  return sheet;
}

function getContributionHeaders_(sheet, config) {
  return sheet.getRange(config.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function getContributionData_(sheet, config) {
  var startRow = config.headerRow + 1;
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];
  return sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn()).getValues();
}

function dataIndexToRow_(dataIndex, config) {
  return config.headerRow + 1 + dataIndex;
}

function resolveColumns_(headers, config) {
  if (config && config.useFixedColumns) {
    return config.columns;
  }
  var cols = { name: 0, annualTotal: -1, receiptNo: -1, dateIssued: -1, email: -1, emailSentAt: -1 };
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || "").toLowerCase().trim();
    if (h.indexOf("annual total") !== -1)            cols.annualTotal = i;
    else if (h === "email" || h === "email address") cols.email = i;
    else if (h.indexOf("receipt no") !== -1)         cols.receiptNo = i;
    else if (h.indexOf("date issued") !== -1)        cols.dateIssued = i;
    else if (h.indexOf("email sent") !== -1)         cols.emailSentAt = i;
  }
  if (cols.annualTotal === -1) throw new Error("Cannot find ANNUAL TOTAL column");
  if (cols.email === -1) throw new Error("Cannot find Email column. Run setup first.");
  if (cols.emailSentAt === -1) throw new Error("Cannot find Email Sent At column. Run setup first.");
  if (cols.receiptNo === -1) cols.receiptNo = cols.annualTotal + 1;
  if (cols.dateIssued === -1) cols.dateIssued = cols.annualTotal + 2;
  return cols;
}

function findHeaderIndex_(headers, target) {
  var needle = target.toLowerCase().trim();
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").toLowerCase().trim() === needle) return i;
  }
  return -1;
}

function nextAnnualReceiptNo_(data, cols, config) {
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
  while (padded.length < config.receiptPad) { padded = "0" + padded; }
  return config.receiptPrefix + padded;
}

function createAnnualPdf_(mergeData, donorName, config) {
  var templateFile = DriveApp.getFileById(config.templateId);
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
  if (config.archiveFolderId) {
    try {
      DriveApp.getFolderById(config.archiveFolderId).createFile(pdfBlob);
      Logger.log("Saved PDF to Drive: " + outputName + ".pdf");
    } catch (e) {
      Logger.log("WARNING: Could not save PDF to Drive folder: " + e.message);
    }
  }
  copyFile.setTrashed(true);
  return pdfBlob;
}

function testDriveSave() {
  var folder = DriveApp.getFolderById("1cJK5PWnIMy_gN_cO-5g-muSEJbdZXE3p");
  var testBlob = Utilities.newBlob("test", "text/plain", "test-drive-save.txt");
  var file = folder.createFile(testBlob);
  Logger.log("SUCCESS: Created file " + file.getName() + " in folder " + folder.getName());
}

function emailAnnualReceipt_(toEmail, donorName, receiptNo, annualTotal, pdfBlob, config) {
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
    "<p>With gratitude,<br>Seeds of Love Online Community, Inc.</p>";
  var options = {
    name: config.email.fromName,
    htmlBody: htmlBody,
    attachments: [pdfBlob]
  };
  if (config.email.fromAlias) {
    options.from = config.email.fromAlias;
    options.replyTo = config.email.fromAlias;
  }
  if (config.email.bccArchive) {
    options.bcc = config.email.bccArchive;
  }
  GmailApp.sendEmail(toEmail, subject, plainBody, options);
}

function formatUsd_(amount) {
  var n = Number(amount || 0);
  return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function escapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
