var CONFIG = {
  sheets: {
    responses: "Form Responses 1",
    commitments: "Recurring Commitments",
    payments: "Payments Received",
    clean: "Clean Donations",
    fxHelper: "_FX_HELPER"
  },
  sequence: {
    receiptPrefix: "2026-",
    commitmentPrefix: "RC-2026-",
    pad: 4
  },
  email: {
    fromAlias: "finance@soloc.net",
    fromName: "Seeds of Love Online Community, Inc."
  },
  sponsorChild: {
    amountPerChildUsd: 7.5
  },
  templates: {
    generalDonationReceiptId: "1QdekD4BEBS484ToyZhkHdthQGHLu30i7UKZjpKDKoz0",
    sponsorDonationReceiptId: "1jJCWZtunKYYw6DFLOEtmM2wc7jKHlQuu9Ed6QyW_988",
    generalPledgeConfirmationId: "13ALtg2Io9fuoTyQFRdhzLSuqEze8d4dO5tvPFiwIR0k",
    sponsorPledgeConfirmationId: "1wjCo3GtqkgX37rjBZ2U_-pi2XFk05NyYzAF6cmIfs2A"
  },
  responseComputedColumns: {
    committedUsd: "Committed Amount USD",
    paidUsd: "Paid Amount USD",
    fxRate: "FX Rate to USD",
    fxRateDate: "FX Rate Date",
    fxStatus: "FX Status",
    receiptSentAt: "Receipt Sent At",
    lastProcessedAt: "Last Processed At",
    processingStatus: "Processing Status",
    processingNotes: "Processing Notes"
  }
};

var HEADER_ALIASES = {
  donorName: [
    "Full Name",
    "Full name",
    "Donor Name",
    "Name"
  ],
  donorEmail: [
    "Donor Email Address",
    "Donor Email address",
    "Donor email address",
    "Email address of Donor",
    "Email Address",
    "Email address",
    "Email"
  ],
  donorPhone: [
    "Contact Number of Donor",
    "Contact Number",
    "Phone",
    "Phone Number"
  ],
  donorAddress: [
    "Address of Donor",
    "Donor Address",
    "Address"
  ],
  givingType: [
    "How would you like to give?",
    "Type of Giving Commitment",
    "Giving Type",
    "Type of Donation",
    "Donation Type",
    "Pledge Type"
  ],
  frequency: [
    "Donation Frequency",
    "Term of donation",
    "Donation Term",
    "Frequency"
  ],
  committedAmount: [
    "Pledged amount",
    "Pledge Amount",
    "Committed Amount",
    "Committed Donation Amount",
    "Amount pledged",
    "Amount",
    "Amount of pledge",
    "Pledge amount in your currency"
  ],
  paidAmount: [
    "Amount actually sent now",
    "Paid Amount",
    "Amount Sent",
    "Amount sent now",
    "Donation Amount",
    "Amount already sent",
    "Initial amount sent",
    "First payment amount"
  ],
  currency: [
    "Currency",
    "Donation Currency",
    "Pledge Currency",
    "Currency of donation",
    "What currency did you send?",
    "What currency is this amount in?"
  ],
  paidStatus: [
    "Has any payment already been sent?",
    "Has this payment already been sent?",
    "Paid Status",
    "Payment Sent?",
    "Did you already send this donation?",
    "Was payment already sent?"
  ],
  paymentDate: [
    "Payment Date",
    "Donation Date"
  ],
  paymentMethod: [
    "Payment Method",
    "Method of Donation",
    "Donation type"
  ],
  purpose: [
    "Donation for:",
    "Donation for",
    "Campaign / Purpose",
    "Purpose",
    "Donation Purpose",
    "Sponsor a child"
  ],
  startDate: [
    "Start month of pledge",
    "Start Date",
    "Start Month",
    "Start month"
  ],
  childCount: [
    "Number of children to sponsor",
    "Number of child to sponsor",
    "Number of Children",
    "Children to Sponsor",
    "Child Count",
    "No. of children"
  ],
  preferredDay: [
    "Preferred day of the month",
    "Preferred Day",
    "Preferred payment day"
  ],
  totalMonths: [
    "Total Months",
    "Number of months",
    "Months"
  ]
};

function setupSolocDonationAutomation() {
  ensureSupportSheets_();
  ensureResponseComputedColumns_();
  resetFormSubmitTrigger_("onFormSubmitMaster");
}

function onFormSubmitMaster(e) {
  if (!e) {
    throw new Error("Do not run onFormSubmitMaster manually. Use reprocessLastResponse() instead.");
  }

  var sheet = e.range ? e.range.getSheet() : SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.responses);
  var row = e.range ? e.range.getRow() : sheet.getLastRow();
  if (!sheet || sheet.getName() !== CONFIG.sheets.responses || row < 2) {
    return;
  }

  processResponseRow_(sheet, row);
}

function reprocessLastResponse() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.responses);
  var row = sheet.getLastRow();
  if (row < 2) {
    throw new Error("No response rows found.");
  }
  processResponseRow_(sheet, row);
}

function reprocessActiveResponseRow() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.responses);
  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    throw new Error("Select a response row first.");
  }
  processResponseRow_(sheet, row);
}

function processExternalPayment(payment) {
  ensureSupportSheets_();

  var record = {
    row: 0,
    donorName: stringValue_(payment.donorName),
    donorEmail: stringValue_(payment.donorEmail),
    donorPhone: stringValue_(payment.donorPhone),
    donorAddress: stringValue_(payment.donorAddress),
    givingType: stringValue_(payment.givingType || "External Payment"),
    frequency: stringValue_(payment.frequency),
    committedAmount: parseAmountNumber_(payment.committedAmount),
    paidAmount: parseAmountNumber_(payment.paidAmount),
    currency: normalizeCurrencyCode_(payment.currency || "USD"),
    paidStatus: "yes",
    paymentDate: normalizeDateValue_(payment.paymentDate, new Date()),
    paymentMethod: stringValue_(payment.paymentMethod || "Bank Transfer"),
    purpose: stringValue_(payment.purpose),
    startDate: normalizeDateValue_(payment.startDate || payment.paymentDate, new Date()),
    childCount: parseAmountNumber_(payment.childCount) || 0,
    preferredDay: stringValue_(payment.preferredDay),
    totalMonths: stringValue_(payment.totalMonths)
  };

  if (!record.donorName || !record.donorEmail || !record.paidAmount) {
    throw new Error("External payment requires donorName, donorEmail, and paidAmount");
  }

  if (/sponsor a child/i.test(record.purpose) && !record.committedAmount && record.childCount) {
    record.committedAmount = round2_(record.childCount * CONFIG.sponsorChild.amountPerChildUsd);
    record.currency = "USD";
  }

  record.committedAmountUsd = getAmountUsd_(record.committedAmount, record.currency, record.startDate);
  record.paidAmountUsd = getAmountUsd_(record.paidAmount, record.currency, record.paymentDate);
  record.fx = getFxToUsdSafe_(record.currency, record.paymentDate);

  var classification = classifyResponse_(record);
  var commitmentId = payment.commitmentId || findCommitmentId_(record.donorEmail, record.purpose, classification.normalizedFrequency);
  var receiptNumber = logPayment_(record, classification, commitmentId);
  sendDonationReceipt_(record, classification, receiptNumber);

  if (commitmentId) {
    updateCommitmentSummary_(commitmentId);
  }

  return {
    receiptNumber: receiptNumber,
    commitmentId: commitmentId || "",
    donorEmail: record.donorEmail,
    paidAmount: record.paidAmount,
    currency: record.currency
  };
}

function processResponseRow_(sheet, row) {
  var headers = getHeaders_(sheet);
  var record = readResponseRecord_(sheet, row, headers);
  var classification = classifyResponse_(record);

  if (!record.donorEmail) {
    updateResponseProcessingState_(sheet, row, "FAILED", "Missing donor email");
    throw new Error("Missing donor email for row " + row);
  }

  if (!record.donorName) {
    updateResponseProcessingState_(sheet, row, "FAILED", "Missing donor name");
    throw new Error("Missing donor name for row " + row);
  }

  if (!classification.isPaid && !classification.isPledgeOrRecurring) {
    updateResponseProcessingState_(sheet, row, "SKIPPED", "No paid donation or pledge/recurring commitment detected");
    return;
  }

  var commitmentId = "";
  if (classification.isPledgeOrRecurring) {
    commitmentId = upsertCommitment_(record, classification);
  }

  var receiptNumber = "";
  if (classification.isPaid) {
    receiptNumber = logPayment_(record, classification, commitmentId);
    sendDonationReceipt_(record, classification, receiptNumber);
    setResponseReceiptSentAt_(sheet, row, new Date());
  } else {
    sendPledgeConfirmation_(record, classification, commitmentId);
  }

  buildCleanDonationsTab_();
  updateResponseProcessingState_(
    sheet,
    row,
    "PROCESSED",
    classification.isPaid ? ("Receipt " + receiptNumber + " sent") : ("Pledge confirmation sent" + (commitmentId ? " for " + commitmentId : ""))
  );
}

function readResponseRecord_(sheet, row, headers) {
  var record = {
    row: row,
    donorName: stringValue_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.donorName)),
    donorEmail: stringValue_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.donorEmail)),
    donorPhone: stringValue_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.donorPhone)),
    donorAddress: stringValue_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.donorAddress)),
    givingType: stringValue_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.givingType)),
    frequency: stringValue_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.frequency)),
    committedAmount: parseAmountNumber_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.committedAmount)),
    paidAmount: parseAmountNumber_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.paidAmount)),
    currency: normalizeCurrencyCode_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.currency) || "USD"),
    paidStatus: stringValue_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.paidStatus)).toLowerCase(),
    paymentDate: normalizeDateValue_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.paymentDate), new Date()),
    paymentMethod: stringValue_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.paymentMethod)),
    purpose: stringValue_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.purpose)),
    startDate: normalizeDateValue_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.startDate), new Date()),
    childCount: parseAmountNumber_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.childCount)) || 0,
    preferredDay: stringValue_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.preferredDay)),
    totalMonths: stringValue_(getValueByAliases_(sheet, row, headers, HEADER_ALIASES.totalMonths))
  };

  if (!record.purpose && /sponsor a child/i.test(record.givingType)) {
    record.purpose = "Sponsor a Child";
  }
  if (!record.purpose && /soloc/i.test(record.givingType)) {
    record.purpose = "General Contribution to SOLOC";
  }
  if (/sponsor a child/i.test(record.purpose) && !record.committedAmount && record.childCount) {
    record.committedAmount = round2_(record.childCount * CONFIG.sponsorChild.amountPerChildUsd);
    record.currency = "USD";
  }
  if (!record.paidAmount && isYes_(record.paidStatus) && record.committedAmount) {
    record.paidAmount = record.committedAmount;
  }

  record.committedAmountUsd = getAmountUsd_(record.committedAmount, record.currency, record.startDate);
  record.paidAmountUsd = getAmountUsd_(record.paidAmount, record.currency, record.paymentDate);
  record.fx = getFxToUsdSafe_(record.currency, record.paymentDate);

  writeComputedValuesToResponse_(sheet, row, record);
  return record;
}

function classifyResponse_(record) {
  var normalizedPurpose = stringValue_(record.purpose).toLowerCase();
  var normalizedGivingType = stringValue_(record.givingType).toLowerCase();
  var normalizedFrequency = stringValue_(record.frequency || inferFrequencyFromGivingType_(record.givingType)).toLowerCase();
  var isSponsorChild = normalizedPurpose.indexOf("sponsor a child") !== -1 || normalizedGivingType.indexOf("sponsor a child") !== -1;
  var isPaid = isYes_(record.paidStatus) || Number(record.paidAmount || 0) > 0;
  var isRecurring = /monthly|quarterly|annual|annually|yearly|weekly/.test(normalizedGivingType) ||
    /monthly|quarterly|annual|annually|yearly|weekly/.test(normalizedFrequency);
  var isPledge = /pledge/.test(normalizedGivingType) ||
    /pledge/.test(normalizedFrequency) ||
    (Number(record.committedAmount || 0) > 0 && !isPaid);

  return {
    program: isSponsorChild ? "SPONSOR_CHILD" : "GENERAL_SOLOC",
    isSponsorChild: isSponsorChild,
    isPaid: isPaid,
    isPledge: isPledge,
    isRecurring: isRecurring,
    isPledgeOrRecurring: isPledge || isRecurring,
    normalizedFrequency: toTitleCase_(normalizedFrequency || "One-time")
  };
}

function upsertCommitment_(record, classification) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.commitments);
  var data = sheet.getDataRange().getValues();
  var donorEmail = stringValue_(record.donorEmail).toLowerCase();
  var purpose = stringValue_(record.purpose).toLowerCase();
  var frequency = stringValue_(classification.normalizedFrequency).toLowerCase();

  for (var i = 1; i < data.length; i++) {
    if (
      stringValue_(data[i][1]).toLowerCase() === "active" &&
      stringValue_(data[i][3]).toLowerCase() === donorEmail &&
      stringValue_(data[i][7]).toLowerCase() === frequency &&
      stringValue_(data[i][14]).toLowerCase() === purpose
    ) {
      return data[i][0];
    }
  }

  var commitmentId = nextSequence_("COMMITMENT_SEQ", CONFIG.sequence.commitmentPrefix);
  sheet.appendRow([
    commitmentId,
    "Active",
    record.donorName,
    record.donorEmail,
    record.donorPhone,
    record.donorAddress,
    classification.isPledge ? "Pledge" : "Recurring",
    classification.normalizedFrequency,
    record.committedAmount || "",
    record.currency || "USD",
    record.committedAmountUsd || "",
    record.startDate || "",
    record.preferredDay || "",
    record.paymentMethod || "",
    record.purpose || "",
    classification.isPaid ? "Yes" : "No",
    record.paidAmount || "",
    classification.isPaid ? record.paymentDate : "",
    "",
    "",
    "",
    "",
    "",
    new Date(),
    new Date()
  ]);
  return commitmentId;
}

function logPayment_(record, classification, commitmentId) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.payments);
  var receiptNumber = nextSequence_("RECEIPT_SEQ", CONFIG.sequence.receiptPrefix);
  sheet.appendRow([
    receiptNumber,
    commitmentId || "",
    record.paymentDate || new Date(),
    record.donorName || "",
    record.donorEmail || "",
    record.paidAmount || "",
    record.currency || "USD",
    record.fx.hasRate ? record.fx.rate : "",
    record.paidAmountUsd || "",
    record.paymentMethod || "",
    "",
    record.purpose || "",
    "",
    "PENDING",
    "",
    new Date()
  ]);

  if (commitmentId) {
    updateCommitmentSummary_(commitmentId);
  }
  return receiptNumber;
}

function sendDonationReceipt_(record, classification, receiptNumber) {
  var templateId = classification.isSponsorChild ?
    CONFIG.templates.sponsorDonationReceiptId :
    CONFIG.templates.generalDonationReceiptId;

  var pdfBlob = createPdfFromTemplate_(templateId, {
    RECEIPT_NUMBER: receiptNumber,
    DONATION_DATE: formatDate_(record.paymentDate),
    DONOR_NAME: record.donorName,
    ORIGINAL_AMOUNT: formatAmountWithCode_(record.paidAmount, record.currency),
    USD_EQUIVALENT: record.paidAmountUsd ? formatAmountWithCode_(record.paidAmountUsd, "USD") : "",
    FX_RATE: record.fx.hasRate ? String(record.fx.rate) : "",
    RATE_DATE: formatDate_(record.paymentDate),
    PURPOSE: record.purpose,
    DONATION_FREQUENCY: classification.normalizedFrequency,
    CHILD_COUNT: record.childCount ? String(record.childCount) : "",
    DONATION_TERM: buildDonationTerm_(record, classification)
  }, "Donation Receipt - " + receiptNumber + " - " + record.donorName);

  var subject = classification.isSponsorChild ?
    ("Sponsor a Child Donation Receipt " + receiptNumber) :
    ("SOLOC Donation Receipt " + receiptNumber);

  var plainBody =
    "Dear " + record.donorName + ",\n\n" +
    "Thank you for your donation to " + (classification.isSponsorChild ? "Sponsor a Child" : "Seeds of Love Online Community") + ".\n\n" +
    "Receipt Number: " + receiptNumber + "\n" +
    "Donation Date: " + formatDate_(record.paymentDate) + "\n" +
    "Amount Received: " + formatAmountWithCode_(record.paidAmount, record.currency) + "\n" +
    (record.paidAmountUsd ? "USD Equivalent: " + formatAmountWithCode_(record.paidAmountUsd, "USD") + "\n" : "") +
    (record.purpose ? "Purpose: " + record.purpose + "\n" : "") +
    "\nAttached is your donation receipt PDF.\n";

  var htmlBody =
    "<p>Dear " + escapeHtml_(record.donorName) + ",</p>" +
    "<p>Thank you for your donation to " + escapeHtml_(classification.isSponsorChild ? "Sponsor a Child" : "Seeds of Love Online Community") + ".</p>" +
    "<p><strong>Receipt Number:</strong> " + escapeHtml_(receiptNumber) + "<br>" +
    "<strong>Donation Date:</strong> " + escapeHtml_(formatDate_(record.paymentDate)) + "<br>" +
    "<strong>Amount Received:</strong> " + escapeHtml_(formatAmountWithCode_(record.paidAmount, record.currency)) +
    (record.paidAmountUsd ? "<br><strong>USD Equivalent:</strong> " + escapeHtml_(formatAmountWithCode_(record.paidAmountUsd, "USD")) : "") +
    (record.purpose ? "<br><strong>Purpose:</strong> " + escapeHtml_(record.purpose) : "") +
    "</p><p>Attached is your donation receipt PDF.</p>";

  sendFromFinance_(record.donorEmail, subject, plainBody, htmlBody, pdfBlob ? [pdfBlob] : []);
  updatePaymentReceiptStatus_(receiptNumber, "SENT", "");
}

function sendPledgeConfirmation_(record, classification, commitmentId) {
  var templateId = classification.isSponsorChild ?
    CONFIG.templates.sponsorPledgeConfirmationId :
    CONFIG.templates.generalPledgeConfirmationId;

  var pdfBlob = createPdfFromTemplate_(templateId, {
    RECEIPT_NUMBER: commitmentId || "",
    DONATION_DATE: formatDate_(record.startDate),
    DONOR_NAME: record.donorName,
    ORIGINAL_AMOUNT: formatAmountWithCode_(record.committedAmount, record.currency),
    USD_EQUIVALENT: record.committedAmountUsd ? formatAmountWithCode_(record.committedAmountUsd, "USD") : "",
    FX_RATE: record.fx.hasRate ? String(record.fx.rate) : "",
    RATE_DATE: formatDate_(record.startDate),
    PURPOSE: record.purpose,
    DONATION_FREQUENCY: classification.normalizedFrequency,
    CHILD_COUNT: record.childCount ? String(record.childCount) : "",
    DONATION_TERM: buildDonationTerm_(record, classification)
  }, "Pledge Confirmation - " + (commitmentId || "pending") + " - " + record.donorName);

  var subject = classification.isSponsorChild ?
    "Sponsor a Child Pledge Confirmation" :
    "SOLOC Pledge Confirmation";

  var plainBody =
    "Dear " + record.donorName + ",\n\n" +
    "Thank you for your pledge to " + (classification.isSponsorChild ? "Sponsor a Child" : "Seeds of Love Online Community") + ".\n\n" +
    (commitmentId ? "Commitment ID: " + commitmentId + "\n" : "") +
    "Frequency: " + classification.normalizedFrequency + "\n" +
    "Amount Pledged: " + formatAmountWithCode_(record.committedAmount, record.currency) + "\n" +
    (record.purpose ? "Purpose: " + record.purpose + "\n" : "") +
    "\nAttached is your pledge confirmation PDF.\n";

  var htmlBody =
    "<p>Dear " + escapeHtml_(record.donorName) + ",</p>" +
    "<p>Thank you for your pledge to " + escapeHtml_(classification.isSponsorChild ? "Sponsor a Child" : "Seeds of Love Online Community") + ".</p>" +
    "<p>" +
    (commitmentId ? "<strong>Commitment ID:</strong> " + escapeHtml_(commitmentId) + "<br>" : "") +
    "<strong>Frequency:</strong> " + escapeHtml_(classification.normalizedFrequency) + "<br>" +
    "<strong>Amount Pledged:</strong> " + escapeHtml_(formatAmountWithCode_(record.committedAmount, record.currency)) +
    (record.purpose ? "<br><strong>Purpose:</strong> " + escapeHtml_(record.purpose) : "") +
    "</p><p>Attached is your pledge confirmation PDF.</p>";

  sendFromFinance_(record.donorEmail, subject, plainBody, htmlBody, pdfBlob ? [pdfBlob] : []);
}

function createPdfFromTemplate_(templateId, mergeData, outputName) {
  if (!templateId || templateId.indexOf("REPLACE_") === 0) {
    throw new Error("Template ID not configured for " + outputName);
  }

  var templateFile = DriveApp.getFileById(templateId);
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

function ensureSupportSheets_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(CONFIG.sheets.commitments)) {
    ss.insertSheet(CONFIG.sheets.commitments);
  }
  if (!ss.getSheetByName(CONFIG.sheets.payments)) {
    ss.insertSheet(CONFIG.sheets.payments);
  }
  if (!ss.getSheetByName(CONFIG.sheets.clean)) {
    ss.insertSheet(CONFIG.sheets.clean);
  }
  ensureCommitmentsHeader_();
  ensurePaymentsHeader_();
}

function ensureCommitmentsHeader_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.commitments);
  var headers = [
    "Commitment ID", "Status", "Donor Name", "Donor Email", "Phone", "Address", "Giving Type", "Frequency",
    "Committed Amount", "Currency", "Committed Amount USD", "Start Date", "Preferred Day", "Payment Method",
    "Campaign / Purpose", "First Payment Already Sent", "First Payment Amount", "First Payment Date", "Last Payment Date",
    "Next Expected Payment", "Total Received", "Payments Count", "Notes", "Created At", "Updated At"
  ];
  writeHeaderIfEmpty_(sheet, headers);
}

function ensurePaymentsHeader_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.payments);
  var headers = [
    "Receipt Number", "Commitment ID", "Payment Date", "Donor Name", "Donor Email", "Amount Received",
    "Currency", "USD Rate", "USD Equivalent", "Payment Method", "Payment Reference", "Campaign / Purpose",
    "Receipt Sent At", "Receipt Status", "Notes", "Created At"
  ];
  writeHeaderIfEmpty_(sheet, headers);
}

function ensureResponseComputedColumns_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.responses);
  var headers = getHeaders_(sheet);
  var desired = [
    CONFIG.responseComputedColumns.committedUsd,
    CONFIG.responseComputedColumns.paidUsd,
    CONFIG.responseComputedColumns.fxRate,
    CONFIG.responseComputedColumns.fxRateDate,
    CONFIG.responseComputedColumns.fxStatus,
    CONFIG.responseComputedColumns.receiptSentAt,
    CONFIG.responseComputedColumns.lastProcessedAt,
    CONFIG.responseComputedColumns.processingStatus,
    CONFIG.responseComputedColumns.processingNotes
  ];

  for (var i = 0; i < desired.length; i++) {
    if (headers.indexOf(desired[i]) === -1) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue(desired[i]);
      headers = getHeaders_(sheet);
    }
  }
}

function writeComputedValuesToResponse_(sheet, row, record) {
  writeResponseCellByHeader_(sheet, row, CONFIG.responseComputedColumns.committedUsd, record.committedAmountUsd || "");
  writeResponseCellByHeader_(sheet, row, CONFIG.responseComputedColumns.paidUsd, record.paidAmountUsd || "");
  writeResponseCellByHeader_(sheet, row, CONFIG.responseComputedColumns.fxRate, record.fx.hasRate ? record.fx.rate : "");
  writeResponseCellByHeader_(sheet, row, CONFIG.responseComputedColumns.fxRateDate, record.fx.hasRate ? record.paymentDate : "");
  writeResponseCellByHeader_(sheet, row, CONFIG.responseComputedColumns.fxStatus, record.fx.hasRate ? "OK" : "FX LOOKUP FAILED");
}

function setResponseReceiptSentAt_(sheet, row, value) {
  writeResponseCellByHeader_(sheet, row, CONFIG.responseComputedColumns.receiptSentAt, value);
}

function updateResponseProcessingState_(sheet, row, status, notes) {
  writeResponseCellByHeader_(sheet, row, CONFIG.responseComputedColumns.lastProcessedAt, new Date());
  writeResponseCellByHeader_(sheet, row, CONFIG.responseComputedColumns.processingStatus, status);
  writeResponseCellByHeader_(sheet, row, CONFIG.responseComputedColumns.processingNotes, notes || "");
}

function writeResponseCellByHeader_(sheet, row, headerName, value) {
  var headers = getHeaders_(sheet);
  var index = headers.indexOf(headerName);
  if (index === -1) {
    return;
  }
  sheet.getRange(row, index + 1).setValue(value);
}

function updatePaymentReceiptStatus_(receiptNumber, status, notes) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.payments);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(receiptNumber).trim()) {
      if (status === "SENT") {
        sheet.getRange(i + 1, 13).setValue(new Date());
      }
      sheet.getRange(i + 1, 14).setValue(status || "");
      if (notes) {
        sheet.getRange(i + 1, 15).setValue(String(notes).substring(0, 500));
      }
      return;
    }
  }
}

function updateCommitmentSummary_(commitmentId) {
  var commitmentSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.commitments);
  var paymentSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.payments);
  var commitments = commitmentSheet.getDataRange().getValues();
  var payments = paymentSheet.getDataRange().getValues();
  var rowIndex = -1;

  for (var i = 1; i < commitments.length; i++) {
    if (String(commitments[i][0]).trim() === String(commitmentId).trim()) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) {
    return;
  }

  var total = 0;
  var count = 0;
  var lastDate = "";
  for (var j = 1; j < payments.length; j++) {
    if (String(payments[j][1]).trim() === String(commitmentId).trim()) {
      total += Number(payments[j][5] || 0);
      count += 1;
      if (!lastDate || new Date(payments[j][2]) > new Date(lastDate)) {
        lastDate = payments[j][2];
      }
    }
  }

  commitmentSheet.getRange(rowIndex, 19).setValue(lastDate || "");
  commitmentSheet.getRange(rowIndex, 21).setValue(total || "");
  commitmentSheet.getRange(rowIndex, 22).setValue(count || "");
  commitmentSheet.getRange(rowIndex, 25).setValue(new Date());
}

function findCommitmentId_(donorEmail, purpose, frequency) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.commitments);
  var data = sheet.getDataRange().getValues();
  var emailNeedle = stringValue_(donorEmail).toLowerCase();
  var purposeNeedle = stringValue_(purpose).toLowerCase();
  var frequencyNeedle = stringValue_(frequency).toLowerCase();

  for (var i = 1; i < data.length; i++) {
    if (
      stringValue_(data[i][1]).toLowerCase() === "active" &&
      stringValue_(data[i][3]).toLowerCase() === emailNeedle &&
      (!purposeNeedle || stringValue_(data[i][14]).toLowerCase() === purposeNeedle) &&
      (!frequencyNeedle || stringValue_(data[i][7]).toLowerCase() === frequencyNeedle)
    ) {
      return data[i][0];
    }
  }
  return "";
}

function buildCleanDonationsTab_() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName(CONFIG.sheets.responses);
  var out = ss.getSheetByName(CONFIG.sheets.clean);
  var headers = getHeaders_(src);
  var data = src.getDataRange().getValues();
  var output = [[
    "Timestamp",
    "Donor Name",
    "Donor Email",
    "Donor Phone",
    "Donation For",
    "Giving Type",
    "Frequency",
    "Committed Amount",
    "Committed Amount USD",
    "Paid Amount",
    "Paid Amount USD",
    "Currency",
    "Payment Date",
    "Payment Method",
    "Child Count",
    "Receipt Sent At",
    "Processing Status",
    "Processing Notes"
  ]];

  for (var row = 2; row <= src.getLastRow(); row++) {
    output.push([
      readCellByAliases_(src, row, headers, ["Timestamp"]),
      readCellByAliases_(src, row, headers, HEADER_ALIASES.donorName),
      readCellByAliases_(src, row, headers, HEADER_ALIASES.donorEmail),
      readCellByAliases_(src, row, headers, HEADER_ALIASES.donorPhone),
      readCellByAliases_(src, row, headers, HEADER_ALIASES.purpose),
      readCellByAliases_(src, row, headers, HEADER_ALIASES.givingType),
      readCellByAliases_(src, row, headers, HEADER_ALIASES.frequency),
      readCellByAliases_(src, row, headers, HEADER_ALIASES.committedAmount),
      readCellByAliases_(src, row, headers, [CONFIG.responseComputedColumns.committedUsd]),
      readCellByAliases_(src, row, headers, HEADER_ALIASES.paidAmount),
      readCellByAliases_(src, row, headers, [CONFIG.responseComputedColumns.paidUsd]),
      readCellByAliases_(src, row, headers, HEADER_ALIASES.currency),
      readCellByAliases_(src, row, headers, HEADER_ALIASES.paymentDate),
      readCellByAliases_(src, row, headers, HEADER_ALIASES.paymentMethod),
      readCellByAliases_(src, row, headers, HEADER_ALIASES.childCount),
      readCellByAliases_(src, row, headers, [CONFIG.responseComputedColumns.receiptSentAt]),
      readCellByAliases_(src, row, headers, [CONFIG.responseComputedColumns.processingStatus]),
      readCellByAliases_(src, row, headers, [CONFIG.responseComputedColumns.processingNotes])
    ]);
  }

  out.clearContents();
  out.getRange(1, 1, output.length, output[0].length).setValues(output);
  out.setFrozenRows(1);
  out.autoResizeColumns(1, output[0].length);
}

function writeHeaderIfEmpty_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
}

function resetFormSubmitTrigger_(functionName) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger(functionName)
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();
}

function getHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function readCellByAliases_(sheet, row, headers, aliases) {
  return getValueByAliases_(sheet, row, headers, aliases);
}

function getValueByAliases_(sheet, row, headers, aliases) {
  var normalizedAliases = aliases.map(normalizeHeader_);
  var values = sheet.getRange(row, 1, 1, headers.length).getDisplayValues()[0];
  var firstMatchValue = "";

  for (var i = 0; i < headers.length; i++) {
    var normalizedHeader = normalizeHeader_(headers[i]);
    if (normalizedAliases.indexOf(normalizedHeader) !== -1) {
      var cellValue = values[i];
      if (firstMatchValue === "") {
        firstMatchValue = cellValue;
      }
      if (String(cellValue || "").trim() !== "") {
        return cellValue;
      }
    }
  }
  return firstMatchValue;
}

function normalizeHeader_(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/\s*:\s*/g, ":")
    .trim();
}

function normalizeCurrencyCode_(value) {
  var raw = String(value == null ? "" : value).trim();
  if (!raw) {
    return "USD";
  }

  var upper = raw.toUpperCase();
  var direct = upper.match(/\b(USD|AUD|CAD|EUR|GBP|PHP|SGD|NZD|JPY|INR|HKD|CNY|RMB|AED|SAR|QAR)\b/);
  if (direct) {
    return direct[1] === "RMB" ? "CNY" : direct[1];
  }
  if (/PHILIPPINE|PESO/.test(upper) || raw.indexOf("₱") !== -1) return "PHP";
  if (/SINGAPORE/.test(upper) || raw.indexOf("S$") !== -1 || raw.indexOf("SG$") !== -1) return "SGD";
  if (/AUSTRALIAN/.test(upper) || raw.indexOf("A$") !== -1) return "AUD";
  if (/CANADIAN/.test(upper) || raw.indexOf("C$") !== -1) return "CAD";
  if (/NEW ZEALAND/.test(upper) || raw.indexOf("NZ$") !== -1) return "NZD";
  if (/EURO/.test(upper) || raw.indexOf("€") !== -1) return "EUR";
  if (/POUND/.test(upper) || raw.indexOf("£") !== -1) return "GBP";
  if (/YEN/.test(upper) || raw.indexOf("¥") !== -1) return "JPY";
  if (/RUPEE/.test(upper)) return "INR";
  if (raw.indexOf("$") !== -1) return "USD";

  var fallback = upper.replace(/[^A-Z]/g, "").substring(0, 3);
  return fallback || "USD";
}

function normalizeDateValue_(value, fallback) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return value;
  }
  if (typeof value === "number" && isFinite(value)) {
    var ms = Math.round((value - 25569) * 86400 * 1000);
    var serialDate = new Date(ms);
    if (!isNaN(serialDate.getTime())) {
      return serialDate;
    }
  }

  var text = String(value || "").trim();
  if (!text) {
    return fallback ? normalizeDateValue_(fallback, new Date()) : new Date();
  }

  var parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }

  var monthMap = {
    january: 0, jan: 0, february: 1, feb: 1, march: 2, mar: 2, april: 3, apr: 3,
    may: 4, june: 5, jun: 5, july: 6, jul: 6, august: 7, aug: 7,
    september: 8, sep: 8, sept: 8, october: 9, oct: 9, november: 10, nov: 10,
    december: 11, dec: 11
  };
  var lower = text.toLowerCase();
  if (monthMap.hasOwnProperty(lower)) {
    var base = fallback ? normalizeDateValue_(fallback, new Date()) : new Date();
    var year = base.getFullYear();
    var month = monthMap[lower];
    if (month < base.getMonth()) {
      year += 1;
    }
    return new Date(year, month, 1);
  }

  return fallback ? normalizeDateValue_(fallback, new Date()) : new Date();
}

function getAmountUsd_(amount, currency, dateValue) {
  var numeric = Number(amount || 0);
  if (!numeric) {
    return "";
  }
  var fx = getFxToUsdSafe_(currency, dateValue);
  if (!fx.hasRate) {
    return "";
  }
  return round2_(numeric * Number(fx.rate));
}

function getFxToUsdSafe_(currency, dateValue) {
  var code = normalizeCurrencyCode_(currency || "USD");
  if (code === "USD") {
    return { hasRate: true, rate: 1, source: "identity" };
  }

  var props = PropertiesService.getScriptProperties();
  var dateKey = Utilities.formatDate(normalizeDateValue_(dateValue, new Date()), Session.getScriptTimeZone(), "yyyyMMdd");
  var cacheKey = "FX_" + code + "_" + dateKey;
  var cached = props.getProperty(cacheKey);
  if (cached && isFinite(Number(cached))) {
    return { hasRate: true, rate: Number(cached), source: "cache" };
  }

  var rate = getGoogleFinanceFxRate_(code, dateValue);
  if (rate && isFinite(rate)) {
    props.setProperty(cacheKey, String(rate));
    return { hasRate: true, rate: Number(rate), source: "googlefinance" };
  }

  return { hasRate: false, rate: "", source: "unavailable" };
}

function getGoogleFinanceFxRate_(fromCurrency, dateObj) {
  var code = normalizeCurrencyCode_(fromCurrency);
  if (code === "USD") {
    return 1;
  }

  var helper = ensureFxHelperSheet_();
  var cell = helper.getRange("A1");
  var date = normalizeDateValue_(dateObj, new Date());
  var formulas = [
    '=IFERROR(INDEX(GOOGLEFINANCE("CURRENCY:' + code + 'USD","price",DATE(' +
      date.getFullYear() + ',' + (date.getMonth() + 1) + ',' + date.getDate() +
      ')),2,2),"")',
    '=IFERROR(INDEX(GOOGLEFINANCE("CURRENCY:' + code + 'USD"),2,2),"")',
    '=IFERROR(GOOGLEFINANCE("CURRENCY:' + code + 'USD"),"")'
  ];

  for (var i = 0; i < formulas.length; i++) {
    cell.clearContent();
    cell.setFormula(formulas[i]);
    SpreadsheetApp.flush();
    for (var attempts = 0; attempts < 16; attempts++) {
      Utilities.sleep(500);
      SpreadsheetApp.flush();
      var value = cell.getValue();
      if (typeof value === "number" && isFinite(value) && value > 0) {
        return Number(value);
      }
      var display = Number(String(cell.getDisplayValue() || "").replace(/,/g, ""));
      if (isFinite(display) && display > 0) {
        return display;
      }
    }
  }
  return null;
}

function ensureFxHelperSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.sheets.fxHelper);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.sheets.fxHelper);
    sheet.hideSheet();
  }
  return sheet;
}

function nextSequence_(key, prefix) {
  var props = PropertiesService.getScriptProperties();
  var next = Number(props.getProperty(key) || "0") + 1;
  props.setProperty(key, String(next));
  return prefix + padNumber_(next, CONFIG.sequence.pad);
}

function sendFromFinance_(to, subject, plainBody, htmlBody, attachments) {
  to = stringValue_(to);
  if (!to) {
    throw new Error("Recipient email is blank");
  }

  var options = { name: CONFIG.email.fromName };
  if (htmlBody) {
    options.htmlBody = htmlBody;
  }
  if (CONFIG.email.fromAlias) {
    options.from = CONFIG.email.fromAlias;
    options.replyTo = CONFIG.email.fromAlias;
  }

  var normalizedAttachments = normalizeAttachments_(attachments);
  if (normalizedAttachments.length > 0) {
    options.attachments = normalizedAttachments;
  }

  GmailApp.sendEmail(to, subject, plainBody || "", options);
}

function normalizeAttachments_(attachments) {
  if (!attachments) {
    return [];
  }
  if (!Array.isArray(attachments)) {
    attachments = [attachments];
  }

  var result = [];
  for (var i = 0; i < attachments.length; i++) {
    var item = attachments[i];
    if (!item) {
      continue;
    }
    if (typeof item.getBlob === "function") {
      result.push(item.getBlob());
    } else if (typeof item.copyBlob === "function") {
      result.push(item);
    } else {
      throw new Error("Invalid attachment type at index " + i);
    }
  }
  return result;
}

function parseAmountNumber_(value) {
  if (value === "" || value === null || typeof value === "undefined") {
    return null;
  }
  if (typeof value === "number") {
    return isNaN(value) ? null : value;
  }

  var cleaned = String(value).replace(/[^0-9.\-]/g, "");
  if (!cleaned) {
    return null;
  }

  var numeric = Number(cleaned);
  return isNaN(numeric) ? null : numeric;
}

function buildDonationTerm_(record, classification) {
  if (record.totalMonths) {
    return String(record.totalMonths) + " month(s)";
  }
  if (classification.isSponsorChild && record.childCount) {
    return record.childCount + " child(ren), " + classification.normalizedFrequency;
  }
  return classification.normalizedFrequency;
}

function inferFrequencyFromGivingType_(givingType) {
  var text = String(givingType || "").toLowerCase();
  if (/weekly/.test(text)) return "Weekly";
  if (/monthly/.test(text)) return "Monthly";
  if (/quarterly/.test(text)) return "Quarterly";
  if (/annual|annually|yearly/.test(text)) return "Annual";
  if (/one[- ]?time|single|once/.test(text)) return "One-time";
  return "";
}

function isYes_(value) {
  return /^(yes|y|true|paid|already sent)$/i.test(String(value || "").trim());
}

function stringValue_(value) {
  return String(value || "").trim();
}

function formatDate_(value) {
  return Utilities.formatDate(normalizeDateValue_(value, new Date()), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatAmountWithCode_(amount, currency) {
  var code = normalizeCurrencyCode_(currency || "USD");
  var numeric = Number(amount || 0);
  return code + " " + numeric.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function padNumber_(num, size) {
  var text = String(num);
  while (text.length < size) {
    text = "0" + text;
  }
  return text;
}

function round2_(n) {
  var numeric = Number(n);
  if (!isFinite(numeric)) {
    return 0;
  }
  return Math.round(numeric * 100) / 100;
}

function escapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function toTitleCase_(s) {
  return String(s || "").replace(/\w\S*/g, function(part) {
    return part.charAt(0).toUpperCase() + part.substr(1).toLowerCase();
  });
}
