# Required Script Properties

Add these under Project Settings → Script Properties before running:

- TEMPLATE_DOC_ID — Google Doc ID of your EAF template
- ALL_EAFS_FOLDER_ID — Drive folder ID where merged EAFs are saved
- ALL_BILLS_FOLDER_ID — Drive folder ID where bills-only PDFs are saved
- DRAFT_TO — approver's email address
- DRAFT_CC — CC email address
- EMAIL_SIGNATURE — signature block text for the draft email
- DEFAULT_BANKS_JSON — JSON object mapping a short name to {acc, holder, bank, ifsc, branch}
