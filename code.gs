/**
 * ============================================================================
 * POLYFILLS FOR EXTERNAL LIBRARIES (Like pdf-lib)
 * ============================================================================
 * Google Apps Script lacks browser timer APIs. We mock them here so libraries 
 * relying on them don't crash.
 */
if (typeof setTimeout === 'undefined') {
  var setTimeout = function(cb, ms) {
    Utilities.sleep(ms || 0);
    cb();
  };
}
if (typeof clearTimeout === 'undefined') {
  var clearTimeout = function() {};
}

/**
 * ============================================================================
 * CONFIGURATION
 * ============================================================================
 * For Target and Data Source, provide EITHER a NAMED_RANGE (string) 
 * OR BOTH a SHEET_NAME and A1_NOTATION. Set NAMED_RANGE to null if not using it.
 */
const CONFIG = {
  // 1. Where does the dropdown value go?
  TARGET_CELL: {
    NAMED_RANGE: 'selectedDept',      // e.g., 'selected'
    SHEET_NAME: 'deptLoading',        // Fallback if NAMED_RANGE is null
    A1_NOTATION: 'B3'                 // Fallback if NAMED_RANGE is null
  },
  
  // 2. Where is the list of values to iterate through?
  DATA_SOURCE: {
    NAMED_RANGE: 'deptList',          // e.g., 'list'
    SHEET_NAME: 'deptLookup',         // Fallback if NAMED_RANGE is null
    A1_NOTATION: 'A2:A17'             // Fallback if NAMED_RANGE is null
  },

  // 3. The "Wait for Data" Logic
  READY_FLAG: {
    USE_FLAG: false,                  // Set to false to disable waiting entirely
    SHEET_NAME: 'deptLoading',
    A1_NOTATION: 'AL1',                // Cell with formula that outputs TRUE when data is ready
    MAX_WAIT_MS: 10000,               // Max time to wait for TRUE (10 seconds)
    RENDER_DELAY_MS: 500              // Hard pause AFTER data is ready to let Conditional Formatting paint
  },

  // 4. Export & File Settings
  EXPORT: {
    FOLDER_ID: '1_BgcjHTi0YMYn14NbAdpoFnidoBM2hlr',
    PDF_NAME_PREFIX: 'deptLoading',
    TIMEZONE: 'Europe/London',
    DATE_FORMAT: 'yyyy-MM-dd_HH-mm',
    COMBINE_PDFS: true,               // Set to true to generate the master PDF
    COMBINED_FILE_NAME: 'All_Departments_Summary.pdf'
  },
  
  // 5. PDF Print URL Parameters
  PRINT_SETTINGS: {
    format: 'pdf',
    size: '7',           // 7 = A4, 6 = A3
    fzc: 'false',        // Frozen columns
    fzr: 'false',        // Frozen rows
    portrait: 'false',   // Landscape
    fitw: 'true',        // Fit to width
    gridlines: 'false',
    printtitle: 'true',
    sheetnames: 'true',
    printdate: 'true',
    printtime: 'true',
    top_margin: '0.8',
    bottom_margin: '0.2',
    left_margin: '0.2',
    right_margin: '0.2'
  }
};

/**
 * ============================================================================
 * MAIN FUNCTION
 * ============================================================================
 */
async function exportDropdownPDFs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  
  try {
    console.log("Starting PDF Export Process...");
    
    // 1. Validate and Resolve Ranges
    const targetCell = resolveRange(ss, CONFIG.TARGET_CELL, "TARGET_CELL");
    const dataSourceRange = resolveRange(ss, CONFIG.DATA_SOURCE, "DATA_SOURCE");
    
    // We also need the specific ID of the sheet we are printing
    const printSheetId = targetCell.getSheet().getSheetId(); 

    // 2. Get Data & Format Date
    const dropdownValues = dataSourceRange.getValues();
    const validValues = dropdownValues.flat().filter(val => val !== "");
    
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, CONFIG.EXPORT.TIMEZONE, CONFIG.EXPORT.DATE_FORMAT);
    
    // Create Dynamic Export Folder
    const baseFolder = DriveApp.getFolderById(CONFIG.EXPORT.FOLDER_ID);
    const exportFolder = baseFolder.createFolder(`PDF_Exports_${formattedDate}`);
    console.log(`Created new export folder: PDF_Exports_${formattedDate}`);

    // Legacy timestamp math from original script (from A1)
    const rawDateValue = targetCell.getSheet().getRange('A1').getValue();
    const dateValueAdjusted = rawDateValue / 100000000000000;

    let generatedBlobs = [];
    let successCount = 0;

    // 3. Iterate through valid values
    for (let i = 0; i < validValues.length; i++) {
      let value = validValues[i];
      
      // Update Dropdown
      targetCell.setValue(value);
      
      // Wait for Calculation & Rendering
      waitForReady(ss);

      // Prepare Filename
      const safeValue = value.toString().replace(/\s+/g, '-');
      const pdfName = `${CONFIG.EXPORT.PDF_NAME_PREFIX}_${formattedDate}_Đ${(i + 1).toString().padStart(2, '0')}_${safeValue}.pdf`;
      
      console.log(`Generating ${i + 1} of ${validValues.length}: ${pdfName}...`);
      
      // Fetch PDF Blob (Now includes exponential backoff)
      const blob = fetchSinglePdfBlob(ssId, printSheetId, dateValueAdjusted, pdfName);
      
      if (blob) {
        // Save Individual File to the dynamic folder
        exportFolder.createFile(blob);
        generatedBlobs.push(blob);
        successCount++;
        console.log(`✅ Success: ${pdfName}`);
      } else {
        console.error(`❌ Failed permanently: ${pdfName}`);
      }
    }

    // 4. Combine PDFs if requested
    if (CONFIG.EXPORT.COMBINE_PDFS && generatedBlobs.length > 0) {
      console.log("Stitching multi-page PDF together...");
      await mergePDFs(generatedBlobs, CONFIG.EXPORT.COMBINED_FILE_NAME, exportFolder);
    }

    // 5. Completion Summary
    console.log(`🎉 Process Complete! Generated ${successCount} individual files inside 'PDF_Exports_${formattedDate}'.`);
    
  } catch (error) {
    console.error("🚨 Script Error: " + error.message);
    console.error(error.stack);
  }
}

/**
 * ============================================================================
 * HELPER FUNCTIONS
 * ============================================================================
 */

/**
 * Validates the CONFIG input and returns the correct SpreadsheetApp Range.
 */
function resolveRange(ss, configBlock, contextName) {
  if (configBlock.NAMED_RANGE) {
    const range = ss.getRangeByName(configBlock.NAMED_RANGE);
    if (!range) {
      const msg = `🚨 ERROR: Named Range '${configBlock.NAMED_RANGE}' defined in ${contextName} was not found in this spreadsheet.`;
      throw new Error(msg);
    }
    return range;
  } else if (configBlock.SHEET_NAME && configBlock.A1_NOTATION) {
    const sheet = ss.getSheetByName(configBlock.SHEET_NAME);
    if (!sheet) {
      const msg = `🚨 ERROR: Sheet '${configBlock.SHEET_NAME}' defined in ${contextName} was not found.`;
      throw new Error(msg);
    }
    return sheet.getRange(configBlock.A1_NOTATION);
  } else {
    const msg = `🚨 ERROR: Invalid Config for ${contextName}. You must provide EITHER a NAMED_RANGE OR both a SHEET_NAME and A1_NOTATION.`;
    throw new Error(msg);
  }
}

/**
 * Polls the spreadsheet until the Ready cell becomes TRUE, then applies Render Delay.
 */
function waitForReady(ss) {
  SpreadsheetApp.flush(); // Force initial calculation push
  
  if (CONFIG.READY_FLAG.USE_FLAG) {
    const sheet = ss.getSheetByName(CONFIG.READY_FLAG.SHEET_NAME);
    if (sheet) {
      const cell = sheet.getRange(CONFIG.READY_FLAG.A1_NOTATION);
      const startTime = Date.now();
      
      // Poll until TRUE or Timeout
      while (Date.now() - startTime < CONFIG.READY_FLAG.MAX_WAIT_MS) {
        if (cell.getValue() === true) {
          break; // Data is ready!
        }
        Utilities.sleep(500); // Check every half second
        SpreadsheetApp.flush();
      }
    }
  }
  
  // Final pause for visual rendering (Conditional Formatting)
  Utilities.sleep(CONFIG.READY_FLAG.RENDER_DELAY_MS);
}

/**
 * Reaches out to Google's internal export URL to generate a PDF Blob.
 * Utilizes Exponential Backoff to prevent rate-limit flooding.
 */
function fetchSinglePdfBlob(ssId, sheetId, dateValue, pdfName) {
  const queryParams = Object.keys(CONFIG.PRINT_SETTINGS)
    .map(key => `${key}=${CONFIG.PRINT_SETTINGS[key]}`)
    .join('&');
    
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?${queryParams}&timestamp=${dateValue}&gid=${sheetId}`;
  
  const params = { 
    method: "GET", 
    headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  
  const maxRetries = 5;
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    const response = UrlFetchApp.fetch(url, params);
    
    if (response.getResponseCode() === 200) {
      return response.getBlob().setName(pdfName);
    }
    
    // If we fail, wait and retry. Wait time increases exponentially: 1.5s, 3s, 6s, 12s...
    const waitTime = 1500 * Math.pow(2, attempt);
    console.warn(`Fetch attempt ${attempt + 1} failed for ${pdfName} (Status: ${response.getResponseCode()}). Retrying in ${waitTime}ms...`);
    Utilities.sleep(waitTime);
  }
  
  return null; // Failed after all retries
}

/**
 * Uses the local pdf-lib library to merge GAS Blobs in server memory.
 */
async function mergePDFs(blobsArray, combinedFileName, destinationFolder) {
  try {
    // 1. Create a new, empty PDF document using pdf-lib
    const mergedPdf = await PDFLib.PDFDocument.create();

    // 2. Loop through our generated blobs
    for (let blob of blobsArray) {
      // GAS Blobs give Signed Ints (-128 to 127). pdf-lib needs Unsigned Ints (0 to 255).
      const uint8Array = new Uint8Array(blob.getBytes()); 
      
      // Load the individual PDF
      const donorPdf = await PDFLib.PDFDocument.load(uint8Array);
      
      // Copy all pages from the donor to the master document
      const copiedPages = await mergedPdf.copyPages(donorPdf, donorPdf.getPageIndices());
      copiedPages.forEach((page) => mergedPdf.addPage(page));
    }

    // 3. Save the master document to bytes
    const mergedBytes = await mergedPdf.save();
    
    // 4. Convert back to GAS-friendly Signed Int array
    const gasBytes = Array.from(new Int8Array(mergedBytes));
    
    // 5. Create final Blob and save to the dynamic Drive folder
    const finalBlob = Utilities.newBlob(gasBytes, 'application/pdf', combinedFileName);
    destinationFolder.createFile(finalBlob);
    
    console.log(`✅ Successfully merged and saved: ${combinedFileName}`);
    
  } catch (e) {
    console.error("Failed to merge PDFs. Did you install pdf-lib.gs? Error: " + e.message);
    throw new Error("Multi-page PDF generation failed. See execution logs.");
  }
}