# EAF² Automation

Automates the Expense Approval Form (EAF) workflow for Shaastra Finance — generates the EAF document,
merges in up to 5 supporting bills (PDFs or images), and drafts the approval email automatically when
a row in the tracking sheet is marked "Ready".

## What it does
1. Watches a Google Sheet for rows marked "Ready" (via an installable onEdit trigger)
2. Fills in an EAF template (amount, payment category, bank details, amount-in-words) from the row data
3. Merges in up to 5 attached bills — handles PDFs, images, and even encrypted/locked PDFs (via a
   Google Slides round-trip that strips encryption on import)
4. Saves the merged EAF and the bills-only PDF to Drive, writes the links back to the sheet
5. Creates a Gmail draft addressed to the approver with the merged PDF attached

## Tech
Google Apps Script, Drive API, Gmail API, pdf-lib (loaded at runtime)

## Setup
1. Bind this script to your tracking Spreadsheet (Extensions → Apps Script)
2. Enable the Drive API: Services (+) → add "Drive API" (v2)
3. Set up the trigger: Triggers → Add Trigger → `onInstallableEdit` → On edit, from spreadsheet
4. Add the required Script Properties — see SETUP.md
