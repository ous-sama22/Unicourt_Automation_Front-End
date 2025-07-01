// Sheetmanager.gs

/**
 * Updates the "Unicourt Processor Case Details" sheet with data from the backend.
 * Existing rows are updated, and new rows are appended.
 * Cases not in the current refresh data (caseDataMap/errorsMap) are preserved.
 */
function updateCaseDetailsSheet(sheet, caseDataMap, errorsMap) {
  if (!sheet) {
    Logger.log("Unicourt Processor Case Details sheet not provided for update.");
    showUserToast("Unicourt Processor Case Details sheet not found, update skipped.", "Error", 5);
    return;
  }
  showUserToast("Updating 'Unicourt Processor Case Details' sheet with latest case data...", "Processing", 5);

  const headerRow = 1;
  const firstDataRow = headerRow + 1;
  const caseNumColIdx = UNICOURT_CONFIG.CASE_DETAILS_COLS.CASE_NUMBER_FOR_DB_ID - 1; // 0-based
  const maxCols = UNICOURT_CONFIG.CASE_DETAILS_COLS.FINAL_JUDGMENT_AWARDED_TO_CREDITOR_CONTEXT; // 1-based

  // 1. Read existing data from "Unicourt Processor Case Details" sheet
  let existingSheetData = [];
  const existingCaseDataMap = new Map(); // Map<case_number_for_db_id, Array<{rowIndexInArray: number, rowData: Array<any>}>>

  if (sheet.getLastRow() >= firstDataRow) {
    existingSheetData = sheet.getRange(firstDataRow, 1, sheet.getLastRow() - headerRow, maxCols).getValues();
    existingSheetData.forEach((row, index) => {
      const caseNumber = String(row[caseNumColIdx] || "").trim();
      if (caseNumber) {
        // Handle duplicate case numbers by storing arrays of row info
        if (!existingCaseDataMap.has(caseNumber)) {
          existingCaseDataMap.set(caseNumber, []);
        }
        existingCaseDataMap.get(caseNumber).push({ 
          rowIndexInArray: index, // 0-based index within existingSheetData
          rowData: row 
        });
      }
    });
    const totalMappedRows = Array.from(existingCaseDataMap.values()).reduce((sum, rowArray) => sum + rowArray.length, 0);
    Logger.log(`updateCaseDetailsSheet: Read ${existingSheetData.length} existing data rows. Mapped ${existingCaseDataMap.size} unique cases with ${totalMappedRows} total rows (including duplicates).`);
  } else {
    Logger.log("updateCaseDetailsSheet: 'Unicourt Processor Case Details' sheet has no existing data rows (or only header).");
  }

  // 2. Prepare a list of cases to process from the current refresh cycle
  const refreshedCaseNumbers = new Set([...Object.keys(caseDataMap), ...Object.keys(errorsMap)]);
  let updatedCount = 0;
  let newCount = 0;

  // 3. Update existing rows or collect new rows
  const newRowsToAppend = [];

  refreshedCaseNumbers.forEach(caseNum => {
    const backendData = caseDataMap[caseNum];
    const backendErrorMsg = errorsMap[caseNum];
    let rowValues;

    if (backendData) { // Data successfully fetched from backend      
      // Helper function to check if array/JSON is empty
      const isEmptyArrayOrJSON = (value) => {
        if (!value) return true;
        if (Array.isArray(value) && value.length === 0) return true;
        if (typeof value === 'string' && (value.trim() === '[]' || value.trim() === '{}')) return true;
        return false;
      };

      rowValues = [
        backendData.id || "",
        backendData.case_number_for_db_id || caseNum,
        backendData.unicourt_actual_case_number_on_page || "",
        backendData.case_name_for_search || "",
        backendData.input_creditor_name || "",
        backendData.is_business ? "TRUE" : "FALSE",
        backendData.creditor_type || "",
        backendData.unicourt_case_name_on_page || "",
        backendData.case_url_on_unicourt || "",
        backendData.status || (backendErrorMsg ? "Data with Error" : "N/A"),
        backendData.last_submitted_at ? Utilities.formatDate(new Date(backendData.last_submitted_at), "America/New_York", "yyyy-MM-dd HH:mm:ss z") : "",
        backendData.original_creditor_name_from_doc || "",
        backendData.original_creditor_name_source_doc_title || "",
        backendData.creditor_address_from_doc || "",
        backendData.creditor_address_source_doc_title || "",
        backendData.associated_parties ? backendData.associated_parties.join("; ") : "",
        // Use N/A for empty associated parties data
        isEmptyArrayOrJSON(backendData.associated_parties_data) ? "N/A" : JSON.stringify(backendData.associated_parties_data.map(party => ({ name: party.name, address: party.address })), null, 2),
        backendData.creditor_registration_state_from_doc || "",
        backendData.creditor_registration_state_source_doc_title || "",
        // Use N/A for empty processed documents summary
        isEmptyArrayOrJSON(backendData.processed_documents_summary) ? "N/A" : JSON.stringify(backendData.processed_documents_summary, null, 2),
        backendData.final_judgment_awarded_to_creditor || "",
        backendData.final_judgment_awarded_source_doc_title || "",
        backendData.final_judgment_awarded_to_creditor_context || ""
      ];
    } else if (backendErrorMsg) { // Error fetching details from backend
      rowValues = [            
        "", // Backend ID
        caseNum, // Case Number
        "", // Unicourt Actual Case Number
        "(Error updating)", // Case Name for Search
        "", // Input Creditor Name
        "", // Is Business
        "", // Creditor Type
        "", // Unicourt Case Name
        "", // Case URL
        "Error Fetching Detail", // Overall Status
        Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd HH:mm:ss z"), // Last Submitted (effectively last attempt time)
        "", "", "", "", "", "", "", "", backendErrorMsg // Put error in a relevant place, or extend columns
      ];
      // Ensure rowValues has the correct number of columns
      while (rowValues.length < maxCols) {
        rowValues.push("");
      }
      rowValues = rowValues.slice(0, maxCols);
    } else {
      // Should not happen if caseNum is in refreshedCaseNumbers
      Logger.log(`updateCaseDetailsSheet: Case ${caseNum} was in refreshed list but no data or error found.`);
      return; // Skip this case
    }

    if (existingCaseDataMap.has(caseNum)) {
      // Update all existing rows for this case number (handles duplicates)
      const existingRowsArray = existingCaseDataMap.get(caseNum);
      existingRowsArray.forEach(existingInfo => {
        existingSheetData[existingInfo.rowIndexInArray] = rowValues;
      });
      updatedCount += existingRowsArray.length; // Count all updated rows including duplicates
    } else {
      // This is a new case, add to newRowsToAppend
      newRowsToAppend.push(rowValues);
      newCount++;
    }
  });

  // 4. Write back data
  // Batch update approach - collect all changes then write at once
  if (existingSheetData.length > 0 && updatedCount > 0) {
    // Write all existing data back to sheet in one operation
    sheet.getRange(firstDataRow, 1, existingSheetData.length, maxCols).setValues(existingSheetData);
    
    Logger.log(`updateCaseDetailsSheet: Updated ${updatedCount} existing rows using batch operations.`);
  }

  if (newRowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRowsToAppend.length, maxCols)
      .setValues(newRowsToAppend);
    Logger.log(`updateCaseDetailsSheet: Appended ${newCount} new rows.`);
  }

  if (updatedCount > 0 || newCount > 0) {
    showUserToast(`'Unicourt Processor Case Details' updated: ${updatedCount} rows modified, ${newCount} rows added.`, "Complete", 4);
  } else {
    showUserToast("'Unicourt Processor Case Details': No new data or updates for currently listed cases.", "Info", 3);
  }
}

/**
 * Updates specific columns in the "Research" sheet with data from the backend.
 */
function updateSampleResearchedCaseSheet(sheet, caseDataMap) {
  if (!sheet || Object.keys(caseDataMap).length === 0) {
    Logger.log("Research sheet not provided or no data to update.");
    return;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(`Updating '${sheet.getName()}'...`, "Processing", 5);
  const requiredHeaders = [
    UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_NUMBER,
    UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_BIZ_STATE,
    UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.ORIGINAL_CREDITOR_NAME,
    UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.ORIGINAL_CREDITOR_ADDRESS,
    UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.ADDITIONAL_CREDITOR_NAME,
    UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.ADDITIONAL_CREDITOR_ADDRESS,
    UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.VOLUNTARY_DISMISSAL,
    UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.FINAL_JUDGMENT_AWARDED_TO_CREDITOR
  ];
  const headerIndices = getHeadersAndIndices(sheet, requiredHeaders);
  if (!headerIndices) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Header setup error in '${sheet.getName()}'. Update skipped.`, "Error", 5);
    return;
  }

  const caseNumColIdx = headerIndices[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_NUMBER]; // 0-based
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`'${sheet.getName()}' is empty (or only headers). Nothing to update.`, "Info", 3);
    return;
  }

  // Get all data from sheet for processing in memory
  const sheetRange = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
  const sheetValues = sheetRange.getValues();
  
  let updatedCount = 0;

  // Helper function to check if array/JSON is empty
  const isEmptyArrayOrJSON = (value) => {
    if (!value) return true;
    if (Array.isArray(value) && value.length === 0) return true;
    if (typeof value === 'string' && (value.trim() === '[]' || value.trim() === '{}')) return true;
    return false;
  };

  // Iterate through sheet data (starting from row 1, which is index 0 in sheetValues after header)
  for (let i = 1; i < sheetValues.length; i++) { // i = 0 is header row
    const sheetCaseNum = String(sheetValues[i][caseNumColIdx] || "").trim();
    if (sheetCaseNum && caseDataMap.hasOwnProperty(sheetCaseNum)) {
      const backendData = caseDataMap[sheetCaseNum];
      if (backendData) {
        // Update values directly in the sheetValues array
        sheetValues[i][headerIndices[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_BIZ_STATE]] = backendData.creditor_registration_state_from_doc || "";
        sheetValues[i][headerIndices[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.ORIGINAL_CREDITOR_NAME]] = backendData.original_creditor_name_from_doc || "";
        sheetValues[i][headerIndices[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.ORIGINAL_CREDITOR_ADDRESS]] = backendData.creditor_address_from_doc || "";
        sheetValues[i][headerIndices[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.ADDITIONAL_CREDITOR_NAME]] = isEmptyArrayOrJSON(backendData.associated_parties) ? "N/A" : backendData.associated_parties.join("; ");
        sheetValues[i][headerIndices[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.ADDITIONAL_CREDITOR_ADDRESS]] = isEmptyArrayOrJSON(backendData.associated_parties_data) ? "N/A" : JSON.stringify(backendData.associated_parties_data.map(party => ({ name: party.name, address: party.address })), null, 2);
        sheetValues[i][headerIndices[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.VOLUNTARY_DISMISSAL]] = backendData.status === 'Voluntary_Dismissal_Found_Skipped' ? 'Y' : 'N';
        sheetValues[i][headerIndices[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.FINAL_JUDGMENT_AWARDED_TO_CREDITOR]] = backendData.final_judgment_awarded_to_creditor || "";
        updatedCount++;
      }
    }
  }

  if (updatedCount > 0) {
    // Write the entire modified data back to the sheet
    sheetRange.setValues(sheetValues);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Updated ${updatedCount} rows in '${sheet.getName()}'.`, "Complete", 3);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(`No matching rows found/updated in '${sheet.getName()}'.`, "Info", 3);
  }
}

/**
 * Updates the "Unicourt Processor Case Submissions" sheet with jump links to "Unicourt Processor Case Details".
 * Called after "Unicourt Processor Case Details" sheet is successfully updated.
 */
function updateSubmissionsSheetWithJumpLinks(submissionsSheet, caseDetailsSheet) {
  if (!submissionsSheet || !caseDetailsSheet) {
     Logger.log("Submissions or Unicourt Processor Case Details sheet not provided for jump link update.");
     return;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast("Updating jump links in 'Unicourt Processor Case Submissions'...", "Processing", 3);
  const lastSubRow = submissionsSheet.getLastRow();
  if (lastSubRow < 2) return; // No data rows

  const caseDetailsSheetGid = caseDetailsSheet.getSheetId();
  const caseDetailNumbersMap = {}; // Map case number to its actual row in Unicourt Processor Case Details
  
  if (caseDetailsSheet.getLastRow() > 1) {
    const caseDetailsDataRange = caseDetailsSheet.getRange(2, UNICOURT_CONFIG.CASE_DETAILS_COLS.CASE_NUMBER_FOR_DB_ID, caseDetailsSheet.getLastRow() - 1, 1);
    const caseDetailNumbers = caseDetailsDataRange.getValues();
    caseDetailNumbers.forEach((row, index) => {
        const cn = String(row[0]).trim();
        if (cn) caseDetailNumbersMap[cn] = index + 2; // Actual sheet row (1-based array index + 2)
    });
  }


  const submissionCaseNumbersRange = submissionsSheet.getRange(2, UNICOURT_CONFIG.SUBMISSIONS_COLS.CASE_NUMBER_FOR_DB_ID, lastSubRow - 1, 1);
  const submissionCaseNumbers = submissionCaseNumbersRange.getValues();
  
  const jumpLinkFormulas = [];

  for (let i = 0; i < submissionCaseNumbers.length; i++) {
    const currentSubCaseNum = String(submissionCaseNumbers[i][0]).trim();
    let linkFormula = "";
    if (currentSubCaseNum && caseDetailNumbersMap[currentSubCaseNum]) {
        const targetRowInDetails = caseDetailNumbersMap[currentSubCaseNum];
        linkFormula = `=HYPERLINK("#gid=${caseDetailsSheetGid}&range=A${targetRowInDetails}", "View Detail")`;
    } else if (currentSubCaseNum) {
        linkFormula = "(Detail N/A)";
    }
    jumpLinkFormulas.push([linkFormula]);
  }
  
  if (jumpLinkFormulas.length > 0) {
    submissionsSheet.getRange(2, UNICOURT_CONFIG.SUBMISSIONS_COLS.JUMP_TO_CASE_DETAIL_ROW, jumpLinkFormulas.length, 1).setFormulas(jumpLinkFormulas);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast("Jump links updated.", "Complete", 2);
}


/**
 * Ensures a sheet with the given name exists and optionally sets its headers.
 * @param {string} sheetName - The name of the sheet to ensure exists.
 * @param {string[]} [headers] - Optional array of header strings. If null, the sheet will be created without headers. If provided, row 1 will be cleared and the new headers will be written.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet object.
 */
function ensureSheetExists(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
  } else { 
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, sheet.getMaxColumns()).clearContent(); 
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
      sheet.setFrozenRows(1); 
    }
  }
  return sheet;
}

/**
 * Helper function to identify cases that are eligible for refresh based on their status
 * in both submissions and details sheets.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} submissionsSheet - The submissions sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} detailsSheet - The case details sheet
 * @returns {Object} An object with casesToRefresh array and detailsRowMap for matching rows
 */
function getEligibleCasesForRefresh(submissionsSheet, detailsSheet) {
  const casesToRefresh = new Set();
  const detailsRowMap = new Map(); // Maps case number to row in details sheet for updating
  // First, build a map of all cases in the details sheet and their statuses
  const detailsCaseMap = new Map(); // Map<case_number, Array<{status, rowIndex}>>
  if (detailsSheet && detailsSheet.getLastRow() > 1) {    const detailsData = detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 
      UNICOURT_CONFIG.CASE_DETAILS_COLS.FINAL_JUDGMENT_AWARDED_TO_CREDITOR_CONTEXT).getValues();
    detailsData.forEach((row, index) => {
      const caseNumber = String(row[UNICOURT_CONFIG.CASE_DETAILS_COLS.CASE_NUMBER_FOR_DB_ID - 1] || "").trim();
      const status = String(row[UNICOURT_CONFIG.CASE_DETAILS_COLS.OVERALL_STATUS - 1] || "").trim().toLowerCase();
      if (caseNumber) {
        // Handle duplicate case numbers by storing arrays of status info
        if (!detailsCaseMap.has(caseNumber)) {
          detailsCaseMap.set(caseNumber, []);
        }
        detailsCaseMap.get(caseNumber).push({
          status: status,
          rowIndex: index + 2 // Convert to 1-based sheet row
        });
      }
    });
  }

  // Then check cases in submissions sheet
  if (submissionsSheet && submissionsSheet.getLastRow() > 1) {
    const submissionsData = submissionsSheet.getRange(2, 1, submissionsSheet.getLastRow() - 1, 
      UNICOURT_CONFIG.SUBMISSIONS_COLS.JUMP_TO_CASE_DETAIL_ROW).getValues();
    
    submissionsData.forEach((row) => {
      const caseNumber = String(row[UNICOURT_CONFIG.SUBMISSIONS_COLS.CASE_NUMBER_FOR_DB_ID - 1] || "").trim();
      const submissionStatus = String(row[UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS - 1] || "").trim().toLowerCase();        
      // Only process cases that are "Submitted to Backend" in submissions sheet
      if (caseNumber && submissionStatus === "submitted to backend") {
        const detailsInfoArray = detailsCaseMap.get(caseNumber);
        
        if (detailsInfoArray && detailsInfoArray.length > 0) {
          // Check if ANY of the duplicate rows has processing/queued status
          const hasEligibleStatus = detailsInfoArray.some(detailsInfo => {
            const status = detailsInfo.status.toLowerCase();
            return status === "processing" || status === "queued";
          });
          
          if (hasEligibleStatus) {
            casesToRefresh.add(caseNumber);
            // Use the first occurrence for the row map (could be any, since we update all duplicates anyway)
            detailsRowMap.set(caseNumber, detailsInfoArray[0].rowIndex);
          }
        } else {
          // Case doesn't exist in details sheet yet - needs to be appended
          casesToRefresh.add(caseNumber);
        }
      }
    });
  }

  return {
    casesToRefresh: Array.from(casesToRefresh),
    detailsRowMap: detailsRowMap
  };
}