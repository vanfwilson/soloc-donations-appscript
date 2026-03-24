# SOLOC Donations Apps Script

This package is a clean replacement for the existing donation and pledge Apps Script logic.

## Business Goal

This project is meant to give Seeds of Love Online Community a lightweight donor administration system built on Google Forms, Google Sheets, Google Docs, PDF generation, and email.

The goal is to make sure each donor receives the correct acknowledgement automatically, whether they already paid or only made a pledge, without requiring a full accounting platform.

In practical terms, the system should:

- capture donor information from a Google Form
- determine whether the donor has already paid or is only pledging
- determine whether the donation is for general SOLOC support or Sponsor a Child
- generate the correct PDF acknowledgement from a Google Docs template
- send the correct email acknowledgement from `finance@soloc.net`
- maintain internal tracking for commitments, payments, receipt status, and reporting

## Core Business Rules

The workflow has four acknowledgement outcomes:

- paid general SOLOC donation -> donation receipt email + general receipt PDF
- paid Sponsor a Child donation -> donation receipt email + sponsor receipt PDF
- unpaid general SOLOC pledge -> pledge confirmation email + general pledge PDF
- unpaid Sponsor a Child pledge -> pledge confirmation email + sponsor pledge PDF

The response handling rules are:

- if money has already been paid, send a donation receipt
- if the donor has pledged but not yet paid, send a pledge confirmation
- use different wording and PDF templates for general SOLOC giving versus Sponsor a Child
- keep a separate internal record of commitments and received payments
- support foreign-currency donations by calculating USD equivalents when possible
- keep a simplified reporting tab for review and manual follow-up

## Future Direction

This project also leaves room for non-form payment events.

The same receipt engine can be reused for:

- recurring Zelle payments
- bank-triggered payment notifications
- manual back-office payment entry
- future webhook or email-triggered receipt automation

That is why the script includes `processExternalPayment(payment)` in addition to the Google Form submit flow.

## Files

- `Code.gs`: main Apps Script logic
- `appsscript.json`: optional manifest template with scopes

## What It Does

When a donor submits the Google Form, the script:

1. Reads the response row using header aliases so it can tolerate duplicate or inconsistent column names.
2. Determines whether the transaction is:
   - general SOLOC or Sponsor a Child
   - paid now or pledged for later
   - one-time or recurring
3. Writes commitment records into `Recurring Commitments` when needed.
4. Writes payment records into `Payments Received` when money has already been sent.
5. Calculates USD values when the donation is in another currency.
6. Generates a PDF from the correct Google Docs template.
7. Sends the correct email from `finance@soloc.net`.
8. Builds a simplified `Clean Donations` reporting tab.

For non-form payment events, the script can also:

1. Accept a bank-originated payment payload through `processExternalPayment(payment)`.
2. Match the payment to an existing commitment when possible.
3. Log the payment and send the same donation receipt email + PDF flow.

## Before You Deploy

Update the template IDs at the top of `Code.gs`:

- `generalDonationReceiptId`
- `sponsorDonationReceiptId`
- `generalPledgeConfirmationId`
- `sponsorPledgeConfirmationId`

These should be Google Docs template file IDs.

Current mapping already filled in:

- `generalDonationReceiptId`: `1QdekD4BEBS484ToyZhkHdthQGHLu30i7UKZjpKDKoz0`
- `sponsorDonationReceiptId`: `1jJCWZtunKYYw6DFLOEtmM2wc7jKHlQuu9Ed6QyW_988`
- `generalPledgeConfirmationId`: `13ALtg2Io9fuoTyQFRdhzLSuqEze8d4dO5tvPFiwIR0k`
- `sponsorPledgeConfirmationId`: `1wjCo3GtqkgX37rjBZ2U_-pi2XFk05NyYzAF6cmIfs2A`

Expected placeholders in the templates:

- `{{RECEIPT_NUMBER}}`
- `{{DONATION_DATE}}`
- `{{PAYMENT_DATE}}`
- `{{PLEDGE_DATE}}`
- `{{RATE_DATE}}`
- `{{DONOR_NAME}}`
- `{{ORIGINAL_AMOUNT}}`
- `{{USD_EQUIVALENT}}`
- `{{FX_RATE}}`
- `{{PURPOSE}}`
- `{{DONATION_FREQUENCY}}`
- `{{CHILD_COUNT}}`
- `{{DONATION_TERM}}`

## Sheet Setup

The script expects the form responses to land in:

- `Form Responses 1`

It will create or maintain these tabs:

- `Recurring Commitments`
- `Payments Received`
- `Clean Donations`
- `_FX_HELPER`

It will also add these computed columns to the response sheet if they are missing:

- `Committed Amount USD`
- `Paid Amount USD`
- `FX Rate to USD`
- `FX Rate Date`
- `FX Status`
- `Receipt Sent At`
- `Last Processed At`
- `Processing Status`
- `Processing Notes`

## Deployment Steps

1. Open the spreadsheet's Apps Script project.
2. Replace the current `Code.gs` contents with `Code.gs` from this folder.
3. Optionally replace the manifest with `appsscript.json`.
4. Update the template IDs in `CONFIG.templates`.
5. Run `setupSolocDonationAutomation()` once.
6. Grant the requested permissions.
7. Submit a test form for each of the four acknowledgement outcomes.

## Important Differences From The Old Script

- one trigger path instead of multiple competing submit handlers
- one canonical alias lookup function
- no test email fallback
- no duplicate function definitions
- no stubbed-out paid sponsor-child route
- clean commitment and payment logging

## Recommended Tests

Run one submission for each:

- paid general SOLOC donation
- paid Sponsor a Child donation
- unpaid general SOLOC pledge
- unpaid Sponsor a Child pledge

Also test one external payment payload for a recurring donor so you have a clean path for future Zelle or bank-triggered receipts.

Also test one foreign-currency donation and one recurring commitment.
