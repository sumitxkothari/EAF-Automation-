// ============================================================
//  EAF AUTOMATION v4 — Google Apps Script
//
//  REQUIRED ONE-TIME SETUP IN APPS SCRIPT:
//  1. Extensions → Apps Script → Services (+) → Add "Drive API" (v2)
//  2. Set up installable trigger:
//     Triggers → Add Trigger → onInstallableEdit
//     Event source: From spreadsheet | Event type: On edit
// ============================================================

const props = PropertiesService.getScriptProperties();

const CONFIG = {
  SHEET_NAME: "sheet 2",
  TEMPLATE_DOC_ID: props.getProperty('TEMPLATE_DOC_ID'),
  ALL_EAFS_FOLDER_ID: props.getProperty('ALL_EAFS_FOLDER_ID'),
  ALL_BILLS_FOLDER_ID: props.getProperty('ALL_BILLS_FOLDER_ID'),
  PDF_LIB_URL: "https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js",
  DRAFT_TO: props.getProperty('DRAFT_TO'),
  DRAFT_CC: props.getProperty('DRAFT_CC'),
  EMAIL_SIGNATURE: props.getProperty('EMAIL_SIGNATURE'),
};

const DEFAULT_BANKS = JSON.parse(props.getProperty('DEFAULT_BANKS_JSON'));


// ─────────────────────────────────────────────────────────────
//  TRIGGER ENTRY POINT
// ─────────────────────────────────────────────────────────────

function onInstallableEdit(e) {
  if (!e) { console.log("EXIT: no event object"); return; }

  const sheet = e.range.getSheet();
  const actualName = sheet.getName();
  console.log("Sheet: '" + actualName + "' | Expected: '" + CONFIG.SHEET_NAME + "'");
  if (actualName !== CONFIG.SHEET_NAME) { console.log("EXIT: sheet mismatch"); return; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf("Status of EAF") + 1;
  console.log("statusCol=" + statusCol + " | editedCol=" + e.range.columnStart + " | value='" + e.value + "'");

  if (statusCol === 0) { console.log("EXIT: Status of EAF col not found"); return; }
  if (e.range.columnStart !== statusCol) { console.log("EXIT: wrong column"); return; }
  if (e.value !== "Ready") { console.log("EXIT: value not Ready"); return; }

  console.log("✅ Launching processEAF for row " + e.range.rowStart);
  processEAF(e.range.rowStart, sheet, headers);
}


// ─────────────────────────────────────────────────────────────
//  MAIN PIPELINE
// ─────────────────────────────────────────────────────────────

async function processEAF(row, sheet, headers) {
  console.log("processEAF row=" + row);

  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const data = {};
  headers.forEach((h, i) => { data[h] = values[i]; });

  setColValue(sheet, row, headers, "Status of EAF", "Processing");
  SpreadsheetApp.flush();
  console.log("→ Processing");

  try {
    // Load pdf-lib (setTimeout polyfill needed by pdf-lib internally)
    const setTimeout = (f, t) => { Utilities.sleep(t); return f(); };
    console.log("Loading pdf-lib...");
    eval(UrlFetchApp.fetch(CONFIG.PDF_LIB_URL).getContentText());
    console.log("pdf-lib loaded ✅");

    // Step 1 — EAF doc → PDF bytes
    const eafBytes = generateEAFBytes(data);
    console.log("EAF PDF bytes: " + eafBytes.length);

    // Step 2 — Bills PDF bytes
    const billsBytes = await buildBillsPdf(data, PDFLib);
    console.log("Bills PDF bytes: " + billsBytes.length);

    // Step 3 — Merge
    const mergedBytes = await mergePdfs(eafBytes, billsBytes, PDFLib);
    console.log("Merged PDF bytes: " + mergedBytes.length);

    const eafName = (data["EAF name"] || "EAF_Row_" + row).toString().trim();
    const billsName = eafName + " Bills";

    // Step 4 — Save Bills PDF → ALL Bills folder
    const billsBlob = Utilities.newBlob(Array.from(billsBytes), MimeType.PDF, billsName + ".pdf");
    const billsFile = DriveApp.getFolderById(CONFIG.ALL_BILLS_FOLDER_ID).createFile(billsBlob);
    console.log("Bills saved: " + billsFile.getUrl());

    // Step 5 — Save Merged PDF → ALL EAFs folder
    const mergedBlob = Utilities.newBlob(Array.from(mergedBytes), MimeType.PDF, eafName + ".pdf");
    const eafFile = DriveApp.getFolderById(CONFIG.ALL_EAFS_FOLDER_ID).createFile(mergedBlob);
    console.log("Merged EAF saved: " + eafFile.getUrl());

    // Step 6 — Write links to sheet
    setColValue(sheet, row, headers, "EAF pdf link", eafFile.getUrl());
    setColValue(sheet, row, headers, "Only Bills pdf link", billsFile.getUrl());

    // Step 7 — Gmail draft
    setColValue(sheet, row, headers, "Status of Mail", "Drafting");
    SpreadsheetApp.flush();

    const draftBlob = Utilities.newBlob(Array.from(mergedBytes), MimeType.PDF, eafName + ".pdf");

GmailApp.createDraft(
  CONFIG.DRAFT_TO,
  "Request for approval of EAF",
  "Respected Sir,\n\n" +
  "This is regarding the EAF \"" + eafName + "\".\n\n" +
  "Please review and approve.\n\n" +
  "--\n" +
  "Regards\n" +
  CONFIG.EMAIL_SIGNATURE,
  {
    cc: CONFIG.DRAFT_CC,
    attachments: [draftBlob]
  }
);


    setColValue(sheet, row, headers, "Status of Mail", "Drafted");
    setColValue(sheet, row, headers, "Status of EAF", "Done");
    console.log("✅ Done for row " + row);

  } catch (err) {
    console.log("❌ ERROR: " + err.message);
    console.log("Stack: " + err.stack);
    setColValue(sheet, row, headers, "Status of EAF", "Error");
    setColValue(sheet, row, headers, "Error Message", err.message);
  }
}


// ─────────────────────────────────────────────────────────────
//  EAF DOC → PDF BYTES
//  Template is a .docx — Drive.Files.copy with mimeType
//  conversion turns it into a native Google Doc on the fly.
//  Requires: Services → Drive API (v2) enabled.
// ─────────────────────────────────────────────────────────────

function generateEAFBytes(data) {
  console.log("generateEAFBytes start");

  let copyId;
  try {
    const copied = Drive.Files.copy(
      { title: "_temp_eaf_" + Date.now(), mimeType: "application/vnd.google-apps.document" },
      CONFIG.TEMPLATE_DOC_ID
    );
    copyId = copied.id;
    console.log("Converted copy created: " + copyId);
  } catch (e) {
    throw new Error(
      "Drive.Files.copy failed — ensure 'Drive API' is added under Services (+). Error: " + e.message
    );
  }

  Utilities.sleep(2000);

  try {
    const doc = DocumentApp.openById(copyId);
    const body = doc.getBody();
    console.log("Doc opened, replacing placeholders...");

    body.replaceText("\\{1\\}", Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy"));
    body.replaceText("\\{2\\}", String(data["Event Name & Details"] || ""));
    body.replaceText("\\{3\\}", String(data["Remarks"] || ""));

    const amountNum = data["Amount (₹)"] || data["Amount"] || "";
    body.replaceText("\\{5\\}", String(amountNum));

    const amountWords = (data["Amount (Words)"] && String(data["Amount (Words)"]).trim() !== "")
      ? String(data["Amount (Words)"]).trim()
      : numberToWords(Number(amountNum) || 0);
    body.replaceText("\\{6\\}", amountWords);

    handlePaymentTicks(body, data["Payment for"]);
    handleBank(body, data);

    doc.saveAndClose();
    console.log("Doc saved, exporting as PDF...");

    const pdfBytes = DriveApp.getFileById(copyId).getAs(MimeType.PDF).getBytes();
    console.log("PDF export done, bytes: " + pdfBytes.length);
    return pdfBytes;

  } finally {
    try { DriveApp.getFileById(copyId).setTrashed(true); } catch (_) { }
    console.log("Temp doc trashed");
  }
}


function handlePaymentTicks(body, selectedType) {
  const MAP = {
    "Vendor Payment": "\\{4A\\}",
    "Reimbursement": "\\{4B\\}",
    "Prizemoney": "\\{4C\\}",
    "Advance": "\\{4D\\}"
  };
  const sel = (selectedType || "").toString().trim().toLowerCase();
  console.log("Payment for: '" + selectedType + "'");
  Object.entries(MAP).forEach(([label, placeholder]) => {
    body.replaceText(placeholder, label.toLowerCase() === sel ? "✓" : " ");
  });
}


function handleBank(body, data) {
  const type = (data["Bank Details"] || "").toString().trim();
  console.log("Bank Details: '" + type + "'");
  let d;
  if (DEFAULT_BANKS[type]) {
    d = DEFAULT_BANKS[type];
  }
  else if (type === "Other") {
    d = {
      acc: String(data["Account Number"] || ""),
      holder: String(data["Account Holder Name"] || ""),
      bank: String(data["Bank Name"] || ""),
      ifsc: String(data["IFSC Code"] || ""),
      branch: String(data["Branch"] || "")
    };
  } else {
    d = { acc: "", holder: "", bank: "", ifsc: "", branch: "" };
  }
  body.replaceText("\\{7A\\}", d.acc);
  body.replaceText("\\{7B\\}", d.holder);
  body.replaceText("\\{7C\\}", d.bank);
  body.replaceText("\\{7D\\}", d.ifsc);
  body.replaceText("\\{7E\\}", d.branch);
}


// ─────────────────────────────────────────────────────────────
//  DECRYPT PDF VIA GOOGLE SLIDES ROUND-TRIP
//
//  Some PDFs (bank statements, digitally signed certificates,
//  invoices) have encryption/DRM flags that make pdf-lib copy
//  blank pages even with ignoreEncryption:true — the content
//  streams themselves are locked.
//
//  Fix: import into Google Slides (Drive decrypts on import,
//  rendering each page as a slide), then export back as a clean
//  unencrypted PDF that pdf-lib can fully read.
// ─────────────────────────────────────────────────────────────

function decryptPdfViaSlides(fileId) {
  let slidesId = null;
  try {
    console.log("  Slides round-trip for: " + fileId);
    const imported = Drive.Files.copy(
      { title: "_temp_slides_" + Date.now(), mimeType: "application/vnd.google-apps.presentation" },
      fileId
    );
    slidesId = imported.id;
    Utilities.sleep(4000); // wait for Drive to finish rendering all pages as slides

    const cleanBytes = DriveApp.getFileById(slidesId).getAs(MimeType.PDF).getBytes();
    console.log("  Decrypted via Slides, bytes: " + cleanBytes.length);
    return cleanBytes;

  } finally {
    if (slidesId) {
      try { DriveApp.getFileById(slidesId).setTrashed(true); } catch (_) { }
      console.log("  Temp Slides file trashed");
    }
  }
}


// ─────────────────────────────────────────────────────────────
//  BILLS PDF — all pages, handles encrypted PDFs + images
// ─────────────────────────────────────────────────────────────

async function buildBillsPdf(data, PDFLib) {
  const newPdf = await PDFLib.PDFDocument.create();
  let pageCount = 0;
  const skipped = [];

  for (let i = 1; i <= 5; i++) {
    const cell = data["Bill " + i];
    if (!cell || String(cell).trim() === "") continue;

    const ids = extractFileIds(String(cell));
    console.log("Bill " + i + " IDs: " + JSON.stringify(ids));

    for (const id of ids) {
      try {
        const file = DriveApp.getFileById(id);
        const mime = file.getMimeType();
        console.log("Bill " + i + " id=" + id + " mime=" + mime);

        if (mime === MimeType.PDF) {
          // Step 1: Try direct load first (works for most normal PDFs)
          let pdfBytes = new Uint8Array(file.getBlob().getBytes());
          let src;
          let usedSlides = false;

          try {
            src = await PDFLib.PDFDocument.load(pdfBytes);
            // Verify pages actually have content by checking first page content streams
            // A blank/encrypted page will have 0 content streams
            const pages = src.getPages();
            const firstPage = pages[0];
            const hasContent = firstPage && firstPage.node.get(PDFLib.PDFName.of("Contents")) !== undefined;
            if (!hasContent && pages.length > 0) {
              throw new Error("Pages appear empty, trying Slides decryption");
            }
          } catch (loadErr) {
            // Direct load failed or content is empty → use Slides round-trip
            console.log("  Direct load issue: " + loadErr.message + " → falling back to Slides");
            const cleanBytes = decryptPdfViaSlides(id);
            pdfBytes = new Uint8Array(cleanBytes);
            src = await PDFLib.PDFDocument.load(pdfBytes);
            usedSlides = true;
          }

          const copied = await newPdf.copyPages(src, src.getPageIndices());
          copied.forEach(p => newPdf.addPage(p));
          pageCount += copied.length;
          console.log("  Added " + copied.length + " page(s)" + (usedSlides ? " (via Slides)" : ""));

        } else if (["image/jpeg", "image/jpg", "image/png"].includes(mime)) {
          const bytes = new Uint8Array(file.getBlob().getBytes());
          const page = newPdf.addPage([595, 842]); // A4
          const { width, height } = page.getSize();
          const img = mime === "image/png"
            ? await newPdf.embedPng(bytes)
            : await newPdf.embedJpg(bytes);
          const dims = img.scaleToFit(width - 40, height - 40);
          page.drawImage(img, {
            x: (width - dims.width) / 2,
            y: (height - dims.height) / 2,
            width: dims.width, height: dims.height
          });
          pageCount++;
          console.log("  Embedded image");

        } else {
          skipped.push("Bill " + i + " id=" + id + ": unsupported type " + mime);
        }

      } catch (err) {
        skipped.push("Bill " + i + " id=" + id + ": " + err.message);
        console.log("Bill " + i + " error: " + err.message);
      }
    }
  }

  if (pageCount === 0) throw new Error(
    "No valid bill files processed." + (skipped.length ? " Skipped: " + skipped.join(" | ") : "")
  );

  return await newPdf.save();
}


// ─────────────────────────────────────────────────────────────
//  MERGE EAF + BILLS
// ─────────────────────────────────────────────────────────────

async function mergePdfs(eafBytes, billsBytes, PDFLib) {
  const merged = await PDFLib.PDFDocument.create();
  const eafDoc = await PDFLib.PDFDocument.load(new Uint8Array(eafBytes));
  const billsDoc = await PDFLib.PDFDocument.load(billsBytes);

  const ep = await merged.copyPages(eafDoc, eafDoc.getPageIndices());
  ep.forEach(p => merged.addPage(p));

  const bp = await merged.copyPages(billsDoc, billsDoc.getPageIndices());
  bp.forEach(p => merged.addPage(p));

  return await merged.save();
}


// ─────────────────────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────────────────────

function extractFileIds(text) {
  const ids = new Set();
  let m;
  const re1 = /\/d\/([a-zA-Z0-9_-]{25,})/g;
  while ((m = re1.exec(text)) !== null) ids.add(m[1]);
  const re2 = /[?&]id=([a-zA-Z0-9_-]{25,})/g;
  while ((m = re2.exec(text)) !== null) ids.add(m[1]);
  if (ids.size === 0) {
    text.split(",").forEach(part => {
      const t = part.trim();
      if (/^[a-zA-Z0-9_-]{25,}$/.test(t)) ids.add(t);
    });
  }
  return Array.from(ids);
}

function setColValue(sheet, row, headers, colName, value) {
  const idx = headers.indexOf(colName) + 1;
  if (idx > 0) {
    sheet.getRange(row, idx).setValue(value);
  } else {
    console.log("⚠️  Column not found: '" + colName + "'");
  }
}

function numberToWords(num) {
  if (!num || isNaN(num)) return "";
  num = Math.round(Number(num));
  if (num === 0) return "Zero Only";

  const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
    "Seventeen", "Eighteen", "Nineteen"];
  const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];

  function spell(n) {
    if (n === 0) return "";
    if (n < 20) return ones[n];
    if (n < 100) return tens[Math.floor(n / 10)] + (n % 10 ? " " + ones[n % 10] : "");
    return ones[Math.floor(n / 100)] + " Hundred" + (n % 100 ? " " + spell(n % 100) : "");
  }

  let result = "", n = num;
  if (n >= 10000000) { result += spell(Math.floor(n / 10000000)) + " Crore "; n %= 10000000; }
  if (n >= 100000) { result += spell(Math.floor(n / 100000)) + " Lakh "; n %= 100000; }
  if (n >= 1000) { result += spell(Math.floor(n / 1000)) + " Thousand "; n %= 1000; }
  if (n > 0) { result += spell(n); }
  return result.trim() + " Only";
}