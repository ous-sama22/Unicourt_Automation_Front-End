// Global constants for sheet names and important column indices/headers
const UNICOURT_CONFIG = {
  SHEET_NAMES: {
    SUBMISSIONS: "Unicourt Processor Case Submissions",
    CASE_DETAILS: "Unicourt Processor Case Details",
    ALL_CASES: "Unicourt Processor All Cases",
    SAMPLE_RESEARCHED_CASE: "Research"
  },
  SUBMISSIONS_COLS: { // 1-based
    CASE_NUMBER_FOR_DB_ID: 1,       // Column A
    CASE_NAME_FOR_SEARCH: 2,        // Column B
    INPUT_CREDITOR_NAME: 3,         // Column C
    IS_BUSINESS: 4,                 // Column D (TRUE/FALSE)
    CREDITOR_TYPE: 5,               // Column E 
    SUBMISSION_STATUS: 6,           // Column F (Status like 'Submitted to Backend', 'Error')
    LAST_SUBMIT_ATTEMPT: 7,         // Column G
    JUMP_TO_CASE_DETAIL_ROW: 8      // Column H
  },  CASE_DETAILS_COLS: { // 1-based
    BACKEND_ID: 1,
    CASE_NUMBER_FOR_DB_ID: 2,
    UNICOURT_ACTUAL_CASE_NUMBER: 3,
    CASE_NAME_FOR_SEARCH: 4,        
    INPUT_CREDITOR_NAME: 5,         
    IS_BUSINESS: 6,                 
    CREDITOR_TYPE: 7,               
    UNICOURT_CASE_NAME_ON_PAGE: 8,
    CASE_URL: 9,
    OVERALL_STATUS: 10,
    LAST_SUBMITTED_SHEET: 11, 
    ORIGINAL_CREDITOR_NAME_FROM_DOC: 12,
    ORIGINAL_CREDITOR_NAME_DOC_TITLE: 13,
    CREDITOR_ADDRESS_FROM_DOC: 14,
    CREDITOR_ADDRESS_DOC_TITLE: 15,
    ASSOCIATED_PARTIES_NAMES: 16, 
    ASSOCIATED_PARTIES_DATA_JSON: 17, 
    CREDITOR_REGISTRATION_STATE_FROM_DOC: 18,
    CREDITOR_REGISTRATION_STATE_DOC_TITLE: 19,    
    PROCESSED_DOCUMENTS_SUMMARY_JSON: 20,
    FINAL_JUDGMENT_AWARDED_TO_CREDITOR: 21,
    FINAL_JUDGMENT_AWARDED_SOURCE_DOC_TITLE: 22,
    FINAL_JUDGMENT_AWARDED_TO_CREDITOR_CONTEXT: 23
  },
  // For "Research" - these are header names, not column indices
  SAMPLE_RESEARCHED_CASE_HEADERS: {
    COURT_CASE_NUMBER: "Court Case Number", // For case_number_for_db_id
    COURT_CASE_TITLE: "Court Case Title",   // For case_name_for_search
    CREDITOR_PARSED_FIRST_NAME: "(Optional) Creditor Parsed First Name",
    CREDITOR_PARSED_LAST_NAME: "(Optional) Creditor Parsed Last Name",
    CREDITOR_BUSINESS_NAME: "Creditor Business Name",
    CREDITOR_TYPE: "Type",
    // Output columns for refresh
    CREDITOR_BIZ_STATE: "Creditor Business State of Incorporation",
    ORIGINAL_CREDITOR_NAME: "Original Creditor Name",
    ORIGINAL_CREDITOR_ADDRESS: "Original Creditor Address on Judgment",
    ADDITIONAL_CREDITOR_NAME: "Additional Creditor Name",
    ADDITIONAL_CREDITOR_ADDRESS: "Additional Creditor Address on Judgment",
    VOLUNTARY_DISMISSAL: "Voluntary Dismissal/Vacate Judgment",
    FINAL_JUDGMENT_AWARDED_TO_CREDITOR: "Final Judgment Awarded To Listed Creditor?"
  },
  SETTINGS_KEYS: {
    BACKEND_URL: "BACKEND_URL",
    BACKEND_API_KEY: "BACKEND_API_KEY",
    UNICOURT_EMAIL: "UNICOURT_EMAIL",
    UNICOURT_PASSWORD: "UNICOURT_PASSWORD",
    OPENROUTER_KEY: "OPENROUTER_KEY",
    OPENROUTER_MODEL: "OPENROUTER_MODEL",
    EXTRACT_ASSOCIATED_PARTY_ADDRESSES: "EXTRACT_ASSOCIATED_PARTY_ADDRESSES"
  },  ERROR_LOG_KEY: "LAST_API_ERRORS_LIST",
  MAX_ERROR_LOG_ENTRIES: 100,
  SUBMISSION_LOG_KEY: "LAST_SUBMISSION_ATTEMPTS_LIST", 
  MAX_SUBMISSION_LOG_ENTRIES: 100,                      
  DIALOG_DATA_KEY: "CURRENT_DIALOG_DISPLAY_DATA",
  BATCH_SIZE: 25  // Maximum cases per submission batch
};

/**
 * Helper function to chunk an array into batches of specified size
 * @param {Array} array - The array to chunk
 * @param {number} size - The size of each chunk (default: 5)
 * @returns {Array} Array of chunks
 */
function chunkArray(array, size = UNICOURT_CONFIG.BATCH_SIZE) {
  const chunks = [];
  for (let i = 0; i < array.length; i += size) {
    chunks.push(array.slice(i, i + size));
  }
  return chunks;
}


/**
 * Validates case data before submission to ensure all required fields are present and valid.
 * @param {Object} caseData - The case data object to validate
 * @returns {Object} Validation result with isValid boolean and error message if invalid
 */
function validateCaseData(caseData) {
  if (!caseData) return { isValid: false, error: "Case data is missing" };

  // Required field validations with detailed messages
  if (!caseData.case_number_for_db_id?.trim()) {
    return { isValid: false, error: "Case number is required" };
  }
  if (!caseData.case_name_for_search?.trim()) {
    return { isValid: false, error: "Case name is required" };
  }
  if (!caseData.input_creditor_name?.trim()) {
    return { isValid: false, error: "Creditor name is required" };
  }
  if (!caseData.creditor_type?.trim()) {
    return { isValid: false, error: "Creditor type is required" };
  }

  // Type validation for isBusiness
  if (typeof caseData.is_business !== "boolean") {
    // Also handle string representations from sheet data
    if (typeof caseData.is_business === "string") {
      const isBusinessStr = caseData.is_business.trim().toUpperCase();
      if (isBusinessStr !== "TRUE" && isBusinessStr !== "FALSE") {
        return { isValid: false, error: "Is business must be TRUE or FALSE" };
      }
    } else {
      return { isValid: false, error: "Is business must be a boolean value" };
    }
  }

  return { isValid: true, error: null };
}

/**
 * Helper function to submit cases in batches with proper error handling and user feedback
 * @param {Array} casesToSubmit - Array of case objects to submit
 * @param {Array} submissionSourceInfo - Array of source info for sheet updates
 * @param {string} context - Context string for logging (e.g., "Manual", "Auto", "Selected")
 * @returns {Object} Combined response with success status and data
 */
function submitCasesInBatches(casesToSubmit, submissionSourceInfo, context = "Batch") {
  if (!casesToSubmit || casesToSubmit.length === 0) {
    return { success: false, error: "No cases to submit" };
  }

  // Validate all cases before proceeding
  const validationErrors = [];
  const validCases = [];
  casesToSubmit.forEach((caseData, index) => {
    const validation = validateCaseData(caseData);
    if (!validation.isValid) {
      validationErrors.push(`Case ${caseData.case_number_for_db_id || `at index ${index}`}: ${validation.error}`);
    } else {
      validCases.push(caseData);
    }
  });

  if (validCases.length === 0) {
    return {
      success: false,
      error: "No valid cases to submit",
      detail: validationErrors.join('; ')
    };
  }

  const batches = chunkArray(validCases);
  const batchCount = batches.length;

  Logger.log(`${context} Submission: Found ${validCases.length} valid cases out of ${casesToSubmit.length} total. ` +
            `${validationErrors.length > 0 ? 'Submission Errors: ' + validationErrors.join('; ') : ''}`);
  
  Logger.log(`${context} Submission: Submitting ${casesToSubmit.length} cases in ${batchCount} batch(es) of up to ${UNICOURT_CONFIG.BATCH_SIZE} cases each.`);
  
  let totalSubmitted = 0;
  let totalReplaced = 0;
  let totalSkipped = 0;
  let currentQueueSize = 0;
  let batchErrors = [];
  let successfulBatches = 0;
  
  for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
    const batch = batches[batchIndex];
    const batchNumber = batchIndex + 1;
    
    showUserToast(`Submitting batch ${batchNumber} of ${batchCount} (${batch.length} cases)...`, "Processing", 5);
    Logger.log(`${context} Submission: Processing batch ${batchNumber}/${batchCount} with ${batch.length} cases.`);
    
    const payload = { cases: batch };
    const response = callBackendSubmitCases(payload);
    
    // Log submission attempts for each case in the batch
    batch.forEach(caseDto => {
      if (caseDto && caseDto.case_number_for_db_id) {
        logSubmissionAttempt(caseDto.case_number_for_db_id, context);
      }
    });
    
    if (response.success && response.data) {
      successfulBatches++;
      totalSubmitted += response.data.submitted_cases || 0;
      totalReplaced += response.data.deleted_and_resubmitted_cases || 0;
      totalSkipped += response.data.already_queued_or_processing || 0;
      currentQueueSize = response.data.current_queue_size || 0;
      
      Logger.log(`${context} Submission: Batch ${batchNumber} successful - ${response.data.submitted_cases} submitted, ${response.data.deleted_and_resubmitted_cases} replaced, ${response.data.already_queued_or_processing} skipped.`);
    } else {
      const errorMsg = `Batch ${batchNumber} failed: ${response.error || 'Unknown error'}${response.detail ? ' (' + response.detail + ')' : ''}`;
      batchErrors.push(errorMsg);
      Logger.log(`${context} Submission: ${errorMsg}`);
      
      // For failed batches, we'll continue with remaining batches but track the errors
      showUserToast(`Batch ${batchNumber} failed, continuing with remaining batches...`, "Warning", 3);
    }
    
    // Small delay between batches to avoid overwhelming the backend
    if (batchIndex < batches.length - 1) {
      Utilities.sleep(1000);
    }
  }
  
  // Prepare combined response
  const allBatchesSuccessful = batchErrors.length === 0;
  const combinedResponse = {
    success: allBatchesSuccessful,
    data: allBatchesSuccessful ? {
      submitted_cases: totalSubmitted,
      deleted_and_resubmitted_cases: totalReplaced,
      already_queued_or_processing: totalSkipped,
      current_queue_size: currentQueueSize,
      batches_processed: successfulBatches,
      total_batches: batchCount
    } : null,
    error: batchErrors.length > 0 ? `${batchErrors.length} out of ${batchCount} batches failed` : null,
    detail: batchErrors.length > 0 ? batchErrors.join('; ') : null,
    partial_success: successfulBatches > 0 && batchErrors.length > 0,
    successful_batches: successfulBatches,
    failed_batches: batchErrors.length
  };
  
  if (batchErrors.length > 0) {
    Logger.log(`${context} Submission: Completed with errors. ${successfulBatches}/${batchCount} batches successful. Errors: ${batchErrors.join('; ')}`);
  } else {
    Logger.log(`${context} Submission: All ${batchCount} batches completed successfully.`);
  }
  
  return combinedResponse;
}

/**
 * Prompts user and then ensures/resets headers on all primary sheets.
 */
function ensureAllSheetHeadersWithPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Confirm Header Reset & Sheet Structure',
    'This will ensure/reset headers on "Unicourt Processor Case Submissions", "Unicourt Processor Case Details", and "Unicourt Processor All Cases" sheets, and ensure "Research" sheet exists. ' +
    '"Unicourt Processor Case Submissions", "Unicourt Processor Case Details", and "Unicourt Processor All Cases" sheets will be hidden. Existing data below headers will NOT be affected. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response === ui.Button.YES) {
    ensureAllSheetHeaders();
    ui.alert('Sheet headers and structure have been ensured/reset. Specified sheets are now hidden.');
  }
}

/**
 * Ensures/Resets headers on all primary sheets and hides specific ones.
 */
function ensureAllSheetHeaders() {
  const submissionsHeaders = [
    "Case Number (DB Key)", "Case Name for Search", "Input Creditor Name", 
    "Is Business (TRUE/FALSE)", "Creditor Type", "Submission Status", 
    "Last Submit Attempt", "Jump to Case Detail"
  ];  const caseDetailsHeaders = [
    "Backend ID", "Case Number (DB Key)", "Unicourt Actual Case Number", "Case Name (for Search)", 
    "Input Creditor Name", "Is Business", "Creditor Type", "Case Name (from Unicourt)", 
    "Unicourt Case URL", "Overall Status", "Last Submitted At (Backend Time)", 
    "Original Creditor Name (Doc)", "Original Creditor Name Source Doc Title", "Creditor Address (Doc)", 
    "Creditor Address Source Doc Title", "Associated Parties (Names)", "Associated Parties (Data JSON)", 
    "Creditor Reg State (Doc)", "Creditor Reg State Source Doc Title", "Processed Documents Summary (JSON)",
    "Final Judgment Awarded to Creditor?", "Final Judgment Awarded Source Doc Title", "Final Judgment Context"
  ]; const allCasesHeaders = [
    "Backend ID", "Case Number (DB Key)", "Unicourt Case Number", "Case Name (for Search)", 
    "Input Creditor Name", "Is Business", "Creditor Type", "Case Name (from Unicourt)", 
    "Unicourt Case URL", "Overall Status", "Last Submitted At", "Original Creditor Name (Doc)", 
    "Creditor Name Source Doc Title", "Creditor Address (Doc)", "Creditor Address Source Doc Title", 
    "Associated Parties (Names)", "Associated Parties (Data JSON)", "Creditor Reg State (Doc)", 
    "Creditor Reg State Source Doc Title", "Processed Documents Summary (JSON)",
    "Final Judgment Awarded to Creditor?", "Final Judgment Awarded Source Doc Title", "Final Judgment Context"          
  ];

  ensureSheetExists(UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS, submissionsHeaders);
  ensureSheetExists(UNICOURT_CONFIG.SHEET_NAMES.CASE_DETAILS, caseDetailsHeaders);
  ensureSheetExists(UNICOURT_CONFIG.SHEET_NAMES.ALL_CASES, allCasesHeaders);
  ensureSheetExists(UNICOURT_CONFIG.SHEET_NAMES.SAMPLE_RESEARCHED_CASE, null); // Ensure it exists, headers are user-managed

  // Hide specified sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToHide = [
    UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS,
    UNICOURT_CONFIG.SHEET_NAMES.CASE_DETAILS,
    UNICOURT_CONFIG.SHEET_NAMES.ALL_CASES
  ];
  sheetsToHide.forEach(name => {
    const sheetToHide = ss.getSheetByName(name);
    if (sheetToHide) {
      try {
        sheetToHide.hideSheet();
      } catch (e) {
        Logger.log(`Could not hide sheet "${name}": ${e.toString()}`);
        // Optionally alert user or handle if hiding is critical
      }
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast("Sheet headers ensured/reset and relevant sheets hidden.", "Complete", 5);
}

// Code.gs

/**
 * Refreshes case data from the backend.
 * Prioritizes manually submitted cases for refresh.
 * Also refreshes cases from Unicourt Processor Case Details if status is Processing, Queued, or empty,
 * and cases from Submissions sheet not yet in Unicourt Processor Case Details.
 * @param {string[]} [manualSubmitCaseNumbers] - Optional array of case numbers manually submitted, to force refresh.
 */
function refreshAllCaseData(manualSubmitCaseNumbers) {
  showUserToast("Preparing to refresh specific case data...", "Processing...", 5);

  const caseDetailsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNICOURT_CONFIG.SHEET_NAMES.CASE_DETAILS);
  const submissionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS);
  const sampleResearchedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNICOURT_CONFIG.SHEET_NAMES.SAMPLE_RESEARCHED_CASE);

  if (!caseDetailsSheet || !submissionsSheet) {
    showUserMessage(
      "Required sheets ('Unicourt Processor Case Submissions', 'Unicourt Processor Case Details') not found. Please run 'Ensure/Reset Sheet Headers & Structure'.",
      "Sheet Not Found",
      true
    );
    return;
  }

  let caseNumbersToRefresh = new Set();

  // 1. Add manually submitted case numbers if provided
  if (manualSubmitCaseNumbers && Array.isArray(manualSubmitCaseNumbers) && manualSubmitCaseNumbers.length > 0) {
    manualSubmitCaseNumbers.forEach(cn => {
      if (cn && String(cn).trim() !== "") {
        caseNumbersToRefresh.add(String(cn).trim());
      }
    });
    Logger.log(`Refresh: Added ${caseNumbersToRefresh.size} manually submitted cases to refresh queue.`);
  }
  // 2. Get eligible cases based on status in both submission and details sheets
  Logger.log("Refresh: Identifying eligible cases from submissions and details sheets...");
  const eligibleCases = getEligibleCasesForRefresh(submissionsSheet, caseDetailsSheet);
  
  // Add eligible cases to refresh set if not already added manually
  eligibleCases.casesToRefresh.forEach(caseNumber => {
    if (!caseNumbersToRefresh.has(caseNumber)) {
      caseNumbersToRefresh.add(caseNumber);
    }
  });
  
  Logger.log(`Refresh: Found ${eligibleCases.casesToRefresh.length} cases eligible for refresh based on status (excluding manual submissions).`);

  const uniqueCaseNumbersArray = Array.from(caseNumbersToRefresh);

  if (uniqueCaseNumbersArray.length === 0) {
    showUserToast("No cases meet criteria for active refresh or were manually specified. Checking for existing delayed triggers.", "Status", 5);
    Logger.log("Refresh: No cases to actively refresh.");
    manageDelayedRefreshTrigger(false);
    if (submissionsSheet && caseDetailsSheet) {
      updateSubmissionsSheetWithJumpLinks(submissionsSheet, caseDetailsSheet);
    }
    return;
  }

  showUserToast(`Found ${uniqueCaseNumbersArray.length} total case(s) to refresh. Fetching data...`, "Processing...", 10);
  Logger.log(`Refresh: Total unique cases to refresh: ${uniqueCaseNumbersArray.length}`);

  const CHUNK_SIZE = 50;
  let allCaseData = {};
  let allErrors = {};
  let anErrorOccurred = false;

  for (let i = 0; i < uniqueCaseNumbersArray.length; i += CHUNK_SIZE) {
    const chunk = uniqueCaseNumbersArray.slice(i, i + CHUNK_SIZE);
    showUserToast(`Fetching details for chunk ${Math.floor(i / CHUNK_SIZE) + 1} of ${Math.ceil(uniqueCaseNumbersArray.length / CHUNK_SIZE)} (${chunk.length} cases)...`, "Processing...", 5);
    Logger.log(`Refresh: Fetching chunk ${Math.floor(i / CHUNK_SIZE) + 1} with ${chunk.length} cases.`);
    const response = callBackendBatchDetails({ case_numbers_for_db_id: chunk });

    if (response.success) {
      if (response.data.results) Object.assign(allCaseData, response.data.results);
      if (response.data.errors) Object.assign(allErrors, response.data.errors);
    } else {
      anErrorOccurred = true;
      const errorMsg = `Failed to fetch details for a chunk starting with ${chunk[0]}: ${response.error} (Detail: ${response.detail || ''}). Subsequent operations on sheets will be skipped for this refresh cycle.`;
      Logger.log(`Refresh: Error fetching batch details for chunk starting with ${chunk[0]}: ${response.error} (Detail: ${response.detail || ''})`);
      showUserMessage(errorMsg, "Error Fetching Batch Details", true);
      break;
    }
  }

  if (anErrorOccurred) {
    showUserToast("Refresh incomplete due to an API error. Sheets NOT fully updated.", "Error", 5);
    Logger.log("Refresh: Incomplete due to API error.");
    manageDelayedRefreshTrigger(false);
    return;
  }

  Logger.log(`Refresh: Successfully fetched data for ${Object.keys(allCaseData).length} cases. Errors for ${Object.keys(allErrors).length} cases.`);

  updateCaseDetailsSheet(caseDetailsSheet, allCaseData, allErrors);
  updateSubmissionsSheetWithJumpLinks(submissionsSheet, caseDetailsSheet);

  if (sampleResearchedSheet && Object.keys(allCaseData).length > 0) {
    updateSampleResearchedCaseSheet(sampleResearchedSheet, allCaseData);
  } else if (!sampleResearchedSheet) {
    Logger.log("Refresh: 'Research' sheet not found, skipping its update.");
  }

  let hasProcessingOrQueuedCases = false;
  if (Object.keys(allCaseData).length > 0) {
    for (const caseNum in allCaseData) {
      if (allCaseData.hasOwnProperty(caseNum) && allCaseData[caseNum]) {
        // Only consider for auto-follow-up if it wasn't a manually submitted case *unless* it's still processing
        // This logic can get complex. For now, any processing/queued case will trigger follow-up.
        const caseStatus = String(allCaseData[caseNum].status || "").trim().toLowerCase();
        if (caseStatus === 'processing' || caseStatus === 'queued') {
          hasProcessingOrQueuedCases = true;
          Logger.log(`Refresh: Case ${caseNum} is still '${caseStatus}'. Follow-up needed.`);
          break;
        }
      }
    }
  }

  if (manualSubmitCaseNumbers && manualSubmitCaseNumbers.length > 0 && hasProcessingOrQueuedCases) {
      Logger.log("Refresh: Follow-up refresh will be scheduled because some cases (possibly including manually submitted ones) are still processing/queued.");
  }


  manageDelayedRefreshTrigger(hasProcessingOrQueuedCases);

  if (hasProcessingOrQueuedCases) {
    // Toast handled by manageDelayedRefreshTrigger
  } else {
    showUserToast("Case data refresh complete. All processed cases in this batch are finalized.", "Complete", 5);
  }
  Logger.log("Refresh: All case data refresh process finished.");
}

function manageDelayedRefreshTrigger(shouldSchedule) {
  const handlerFunctionName = 'refreshAllCaseData';
  let existingTriggerDeleted = false;

  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === handlerFunctionName && 
        trigger.getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(trigger);
      existingTriggerDeleted = true;
      Logger.log(`manageDelayedRefreshTrigger: Deleted an existing CLOCK trigger for ${handlerFunctionName}.`);
    }
  }

  if (shouldSchedule) {
    ScriptApp.newTrigger(handlerFunctionName)
      .timeBased()
      .after(60 * 1000) // 60 seconds
      .create();
    Logger.log(`manageDelayedRefreshTrigger: Scheduled a delayed refresh for ${handlerFunctionName} in approx. 1 minute.`);
    showUserToast("Follow-up refresh scheduled in ~1 minute.", "Auto-Refresh", 4); // Uses helper
  } else {
    if (existingTriggerDeleted) {
      Logger.log(`manageDelayedRefreshTrigger: Ensured no delayed refresh is scheduled for ${handlerFunctionName} (existing one was cleared).`);
    } else {
      Logger.log(`manageDelayedRefreshTrigger: No delayed refresh needed for ${handlerFunctionName}, and no existing one to clear.`);
    }
  }
}

// --- HTML Dialog Management ---
/**
 * Sets data to be fetched by the HTML dialog.
 */
function setDialogDisplayData(title, htmlContent) {
  PropertiesService.getScriptProperties().setProperty(UNICOURT_CONFIG.DIALOG_DATA_KEY, JSON.stringify({ title: title, content: htmlContent }));
}

/**
 * Called by the HTML dialog to get its data.
 */
function getDialogDisplayData() {
  const dataStr = PropertiesService.getScriptProperties().getProperty(UNICOURT_CONFIG.DIALOG_DATA_KEY);
  PropertiesService.getScriptProperties().deleteProperty(UNICOURT_CONFIG.DIALOG_DATA_KEY); // Clean up
  return dataStr ? JSON.parse(dataStr) : null;
}

function showHtmlDialog(title, htmlFileName, width, height) {
  const htmlOutput = HtmlService.createHtmlOutputFromFile(htmlFileName)
    .setWidth(width)
    .setHeight(height);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}

// --- Menu Item Functions for Backend Info ---

function displayBackendServiceStatus_HTML() {
  SpreadsheetApp.getActiveSpreadsheet().toast("Fetching backend service status...", "Processing", 5);
  const response = callBackendApi('/service/status', 'get', null);
  if (response.success) {
    let htmlContent = "<ul>";
    for (const key in response.data) {
      htmlContent += `<li><strong>${String(key).replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase())}:</strong> ${response.data[key]}</li>`;
    }
    htmlContent += "</ul>";
    setDialogDisplayData("Backend Service Status", htmlContent);
    showHtmlDialog("Backend Service Status", "Unicourt Processor DialogDisplay", 450, 450);
  } else {
    SpreadsheetApp.getUi().alert('Error', `Could not retrieve backend service status: ${response.error} ${response.detail ? '('+response.detail+')' : ''}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function displayBackendHealth_HTML() {
  SpreadsheetApp.getActiveSpreadsheet().toast("Checking backend health...", "Processing", 5);
  const response = callBackendApi('/healthz', 'get', null);
  if (response.success) {
    let htmlContent = `
      <p><strong>Status:</strong> ${response.data.status || 'N/A'}</p>
      <p><strong>Message:</strong> ${response.data.message || 'N/A'}</p>
    `;
    setDialogDisplayData("Backend Health Check", htmlContent);
    showHtmlDialog("Backend Health Check", "Unicourt Processor DialogDisplay", 400, 250);
  } else {
    SpreadsheetApp.getUi().alert('Error Checking Health', `Could not retrieve backend health: ${response.error} ${response.detail ? '('+response.detail+')' : ''}`);
  }
}

function promptAndDisplayBatchCaseStatus_HTML() {
  const ui = SpreadsheetApp.getUi();
  const activeRange = SpreadsheetApp.getActiveRange();
  if (!activeRange) {
    ui.alert("Please select cells containing case data in a relevant sheet.");
    return;
  }

  const activeSheet = activeRange.getSheet();
  const activeSheetName = activeSheet.getName();
  let caseNumberColumnIndex; // 0-based index

  if (activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS) {
    caseNumberColumnIndex = UNICOURT_CONFIG.SUBMISSIONS_COLS.CASE_NUMBER_FOR_DB_ID - 1;
  } else if (activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.CASE_DETAILS) {
    caseNumberColumnIndex = UNICOURT_CONFIG.CASE_DETAILS_COLS.CASE_NUMBER_FOR_DB_ID - 1;
  } else if (activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.SAMPLE_RESEARCHED_CASE) {
    const headerMap = getHeadersAndIndices(activeSheet, [UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_NUMBER]);
    if (!headerMap) return; // Error already shown by getHeadersAndIndices
    caseNumberColumnIndex = headerMap[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_NUMBER];
  } else {
    ui.alert(`Please select cells in a supported sheet ("${UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS}", "${UNICOURT_CONFIG.SHEET_NAMES.CASE_DETAILS}", or "${UNICOURT_CONFIG.SHEET_NAMES.SAMPLE_RESEARCHED_CASE}").`);
    return;
  }
  
  const data = activeRange.getValues();
  const caseNumbersArray = [];

  for (let i = 0; i < data.length; i++) {
    const caseNumber = String(data[i][caseNumberColumnIndex] || "").trim();
    if (caseNumber) {
      caseNumbersArray.push(caseNumber);
    }
  }

  const uniqueCaseNumbers = [...new Set(caseNumbersArray)];
  if (uniqueCaseNumbers.length === 0) {
    ui.alert("No case numbers found in the selected cells. Please ensure you've selected cells that contain case numbers in the correct column.");
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Fetching batch case status...", "Processing", 10);
  const response = callBackendApi('/cases/batch-status', 'post', { case_numbers_for_db_id: uniqueCaseNumbers });

  if (response.success) {
    let htmlContent = "<h4>Results:</h4>";
    if (response.data.results && Object.keys(response.data.results).length > 0) {
      htmlContent += "<ul>";
      for (const caseNum in response.data.results) {
        const item = response.data.results[caseNum];
        if (item) {
          htmlContent += `<li><strong>${caseNum}:</strong> Status: ${item.status || 'N/A'}. Message: ${item.message || '(No message)'}`;
          if (item.data) { 
            htmlContent += `<br/>  <em>Unicourt Name:</em> ${item.data.unicourt_case_name_on_page || 'N/A'}`;
            htmlContent += `<br/>  <em>Creditor (Doc):</em> ${item.data.original_creditor_name_from_doc || 'N/A'}`;
            htmlContent += `<br/>  <em>Creditor Addr (Doc):</em> ${item.data.creditor_address_from_doc || 'N/A'}`;
          }
          htmlContent += `</li>`;
        } else {
          htmlContent += `<li><strong>${caseNum}:</strong> No data returned.</li>`;
        }
      }
      htmlContent += "</ul>";
    } else {
      htmlContent += "<p>No results found for the provided case numbers.</p>";
    }

    if (response.data.errors && Object.keys(response.data.errors).length > 0) {
      htmlContent += "<h4>Errors:</h4><ul>";
      for (const caseNum in response.data.errors) {
        htmlContent += `<li><strong>${caseNum}:</strong> ${response.data.errors[caseNum]}</li>`;
      }
      htmlContent += "</ul>";
    }
    setDialogDisplayData("Batch Case Status", htmlContent);
    showHtmlDialog("Batch Case Status", "Unicourt Processor DialogDisplay", 600, 500);
  } else {
    ui.alert('Error Fetching Batch Status', `Could not retrieve batch status: ${response.error} ${response.detail ? '('+response.detail+')' : ''}`);
  }
}

function viewLastApiErrorLog_HTML() {
  const errorsJson = PropertiesService.getScriptProperties().getProperty(UNICOURT_CONFIG.ERROR_LOG_KEY);
  let errors = [];
  if (errorsJson) {
    try {
      errors = JSON.parse(errorsJson);
    } catch (e) {
      setDialogDisplayData("API Error Log", "<p>Error parsing stored error log.</p>");
      showHtmlDialog("API Error Log", "Unicourt Processor DialogDisplay", 600, 400);
      return;
    }
  }

  let htmlContent;
  if (errors.length === 0) {
    htmlContent = "<p>No API errors logged yet.</p>";
  } else {
    htmlContent = "<ul>";
    errors.forEach(err => {
      htmlContent += `<li>
        <strong>Timestamp:</strong> ${new Date(err.timestamp).toLocaleString()}<br/>
        <strong>Method:</strong> ${err.method || 'N/A'}<br/>
        <strong>Endpoint:</strong> ${err.endpoint || 'N/A'}<br/>
        <strong>Code:</strong> ${err.responseCode || 'N/A'}<br/>
        <strong>Response:</strong> <pre>${JSON.stringify(err.responseBody, null, 2) || 'N/A'}</pre>
      </li><hr/>`;
    });
    htmlContent += "</ul>";
  }
  setDialogDisplayData("API Error Log (Last " + errors.length + ")", htmlContent);
  showHtmlDialog("API Error Log", "Unicourt Processor DialogDisplay", 700, 500);
}

function viewBackendDocs_HTML() {
  const properties = PropertiesService.getScriptProperties();
  let backendUrl = properties.getProperty(UNICOURT_CONFIG.SETTINGS_KEYS.BACKEND_URL);
  
  if (!backendUrl) {
    SpreadsheetApp.getUi().alert('Error', 'Backend URL not configured in Script Properties.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  backendUrl = backendUrl.replace(/\/api\/v1$/, '');
  const docsUrl = `${backendUrl}/docs`;
  const html = HtmlService.createHtmlOutput(
    `<script>window.open('${docsUrl}', '_blank'); google.script.host.close();</script>`
  ).setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Documentation...');
}

// --- Configuration Sidebar Functions ---

function showConfigurationSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Unicourt Processor UISidebar')
    .setTitle('Backend Configuration')
    .setWidth(1000);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getAppSettingsFromScriptProperties() {
  const properties = PropertiesService.getScriptProperties();
  return {
    backendUrl: properties.getProperty(UNICOURT_CONFIG.SETTINGS_KEYS.BACKEND_URL) || "http://localhost:8000/api/v1",
    backendApiKey: properties.getProperty(UNICOURT_CONFIG.SETTINGS_KEYS.BACKEND_API_KEY) || "",
    unicourtEmail: properties.getProperty(UNICOURT_CONFIG.SETTINGS_KEYS.UNICOURT_EMAIL) || "",
    unicourtPassword: "", // Never return saved password to UI
    openrouterKey: "",    // Never return saved key to UI
    openrouterModel: properties.getProperty(UNICOURT_CONFIG.SETTINGS_KEYS.OPENROUTER_MODEL) || "google/gemini-2.0-flash-001",
    extractAssociatedPartyAddresses: properties.getProperty(UNICOURT_CONFIG.SETTINGS_KEYS.EXTRACT_ASSOCIATED_PARTY_ADDRESSES) === 'true' 
  };
}

function saveBackendConnectionSettings_scriptOnly(settings) {
  try {
    const properties = PropertiesService.getScriptProperties();
    if (settings.backendUrl !== undefined) properties.setProperty(UNICOURT_CONFIG.SETTINGS_KEYS.BACKEND_URL, settings.backendUrl);
    // Only save API key if a new one is provided
    if (settings.backendApiKey !== undefined && settings.backendApiKey !== "") properties.setProperty(UNICOURT_CONFIG.SETTINGS_KEYS.BACKEND_API_KEY, settings.backendApiKey);
    return { success: true, message: "Backend Connection settings saved to Apps Script." };
  } catch (e) {
    Logger.log(`Error in saveBackendConnectionSettings_scriptOnly: ${e.toString()}`);
    return { success: false, error: `Failed to save connection settings: ${e.toString()}` };
  }
}

function saveClientCredentials_scriptAndBackend(settings) {
  let scriptSaveSuccess = false;
  let backendUpdateResponse = { success: false, error: "Not attempted", data: {} };

  try {
    const properties = PropertiesService.getScriptProperties();
    if (settings.unicourtEmail !== undefined) properties.setProperty(UNICOURT_CONFIG.SETTINGS_KEYS.UNICOURT_EMAIL, settings.unicourtEmail);
    if (settings.unicourtPassword !== undefined && settings.unicourtPassword !== "") properties.setProperty(UNICOURT_CONFIG.SETTINGS_KEYS.UNICOURT_PASSWORD, settings.unicourtPassword);
    if (settings.openrouterKey !== undefined && settings.openrouterKey !== "") properties.setProperty(UNICOURT_CONFIG.SETTINGS_KEYS.OPENROUTER_KEY, settings.openrouterKey);
    if (settings.openrouterModel !== undefined) properties.setProperty(UNICOURT_CONFIG.SETTINGS_KEYS.OPENROUTER_MODEL, settings.openrouterModel);
    if (settings.extractAssociatedPartyAddresses !== undefined) properties.setProperty(UNICOURT_CONFIG.SETTINGS_KEYS.EXTRACT_ASSOCIATED_PARTY_ADDRESSES, String(settings.extractAssociatedPartyAddresses));
    scriptSaveSuccess = true;

    const backendPayload = {};
    if (settings.unicourtEmail) backendPayload.UNICOURT_EMAIL = settings.unicourtEmail;
    if (settings.unicourtPassword) backendPayload.UNICOURT_PASSWORD = settings.unicourtPassword; // Only send if provided
    if (settings.openrouterKey) backendPayload.OPENROUTER_API_KEY = settings.openrouterKey; // Only send if provided
    if (settings.openrouterModel) backendPayload.OPENROUTER_LLM_MODEL = settings.openrouterModel;
    if (settings.extractAssociatedPartyAddresses !== undefined) backendPayload.EXTRACT_ASSOCIATED_PARTY_ADDRESSES = settings.extractAssociatedPartyAddresses;
    
    if (Object.keys(backendPayload).length > 0) {
      backendUpdateResponse = callBackendConfigUpdate(backendPayload);
    } else {
      backendUpdateResponse = { success: true, data: { message: "No client credentials or settings provided to update backend.", updated_fields: {}, restart_required: false }};
    }

    return { 
      scriptSaveSuccess: scriptSaveSuccess, 
      backendUpdateResult: backendUpdateResponse 
    };

  } catch (e) {
    Logger.log(`Error in saveClientCredentials_scriptAndBackend: ${e.toString()}`);
    return { 
      scriptSaveSuccess: scriptSaveSuccess, 
      backendUpdateResult: { success: false, error: `Script error: ${e.toString()}` }
    };
  }
}
    
function saveAllSettings_scriptAndBackend(settings) {
  const connectionResult = saveBackendConnectionSettings_scriptOnly(settings);
  const credentialsResult = saveClientCredentials_scriptAndBackend(settings); 

  return {
    connectionSaveResult: connectionResult,
    credentialsSaveResult: credentialsResult
  };
}

function requestBackendRestartFromSidebar() {
  return callBackendRestart(); 
}

function displayCurrentBackendConfigFromSidebar() {
  const response = fetchCurrentBackendConfig();
  
  if (response.success && response.data) {
    const prettyConfig = JSON.stringify(response.data, null, 2);
    let htmlContent = '<pre style="white-space: pre-wrap; word-wrap: break-word;">';
    htmlContent += prettyConfig.replace(/</g, '&lt;').replace(/>/g, '&gt;');
    htmlContent += '</pre>';
    
    setDialogDisplayData("Backend Configuration", htmlContent);
    showHtmlDialog("Backend Configuration", "Unicourt Processor DialogDisplay", 700, 500);
  }
  return response;
}


/**
 * Displays a message to the user if in a UI context, otherwise logs it.
 * @param {string} message The main message.
 * @param {string} title The title for the dialog (if UI is available).
 * @param {boolean} isError If true, logs as an error and uses an error title prefix.
 */
function showUserMessage(message, title, isError = false) {
  try {
    // Attempt to get UI. If it fails, we are not in a UI context.
    const ui = SpreadsheetApp.getUi(); 
    const fullTitle = isError ? `Error: ${title}` : title;
    ui.alert(fullTitle, message, ui.ButtonSet.OK);
  } catch (e) {
    // Not in UI context, log instead.
    const logMessage = `${title}: ${message}`;
    if (isError) {
      Logger.log(`ERROR: ${logMessage}`);
    } else {
      Logger.log(`INFO: ${logMessage}`);
    }
  }
}

/**
 * Shows a toast message if in a UI context, otherwise logs it.
 * @param {string} message The message for the toast/log.
 * @param {string} title The title for the toast (optional).
 * @param {number} timeoutSeconds The duration for the toast (optional).
 */
function showUserToast(message, title, timeoutSeconds) {
  try {
    // Attempt to get UI.
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, timeoutSeconds);
  } catch (e) {
    // Not in UI context, log instead.
    const logMessage = title ? `${title}: ${message}` : message;
    Logger.log(`TOAST_EQUIVALENT: ${logMessage}`);
  }
}

/**
 * Gets 0-based column indices for given header names from a sheet.
 * Logs an error and returns null if any required header is not found.
 * This version is safe for non-UI contexts like triggers.
 */
function getHeadersAndIndices(sheet, requiredHeaderNames) {
  if (!sheet) {
    Logger.log("getHeadersAndIndices: Sheet object is null or undefined.");
    return null;
  }
  // Check if sheet has any columns before trying to get a range
  if (sheet.getLastColumn() === 0 || sheet.getLastRow() === 0) { 
    Logger.log(`getHeadersAndIndices: Sheet "${sheet.getName()}" appears to be empty or headers cannot be read (lastColumn: ${sheet.getLastColumn()}, lastRow: ${sheet.getLastRow()}).`);
    return null; 
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  const missingHeaders = [];

  requiredHeaderNames.forEach(reqHeader => {
    const index = headers.findIndex(h => String(h || "").trim().toLowerCase() === reqHeader.trim().toLowerCase());
    if (index !== -1) {
      headerMap[reqHeader] = index; // 0-based index
    } else {
      missingHeaders.push(reqHeader);
    }
  });

  if (missingHeaders.length > 0) {
    const errorMessage = `Header Error in Sheet: "${sheet.getName()}". The following required column headers were not found: ${missingHeaders.join(", ")}. Please ensure the headers are correct.`;
    Logger.log(errorMessage); // Log the error
    return null; // Return null when headers are missing
  }
  return headerMap;
}

function submitSelectedCasesToBackend() {
  const ui = SpreadsheetApp.getUi(); // Still useful for initial checks if run from menu
  showUserToast("Preparing to submit selected cases...", "Processing...", 5);
  
  let validationErrors = [];
  let invalidCases = [];
  
  const activeRange = SpreadsheetApp.getActiveRange();
  if (!activeRange) {
    showUserMessage("Please select cells/rows containing case data in a relevant sheet.", "Selection Required", true);
    return;
  }

  const activeSheet = activeRange.getSheet();
  const activeSheetName = activeSheet.getName();
  const casesToSubmit = [];
  const submissionSourceInfo = [];

  let columnMapping = {};
  let headerMap_SampleResearched;

  if (activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS || activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.CASE_DETAILS) {
    columnMapping = (activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS) ? {
      caseNumber: UNICOURT_CONFIG.SUBMISSIONS_COLS.CASE_NUMBER_FOR_DB_ID - 1,
      caseName: UNICOURT_CONFIG.SUBMISSIONS_COLS.CASE_NAME_FOR_SEARCH - 1,
      creditorName: UNICOURT_CONFIG.SUBMISSIONS_COLS.INPUT_CREDITOR_NAME - 1,
      isBusiness: UNICOURT_CONFIG.SUBMISSIONS_COLS.IS_BUSINESS - 1,
      creditorType: UNICOURT_CONFIG.SUBMISSIONS_COLS.CREDITOR_TYPE - 1,
      statusColumn: UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS,
      lastAttemptColumn: UNICOURT_CONFIG.SUBMISSIONS_COLS.LAST_SUBMIT_ATTEMPT
    } : {
      caseNumber: UNICOURT_CONFIG.CASE_DETAILS_COLS.CASE_NUMBER_FOR_DB_ID - 1,
      caseName: UNICOURT_CONFIG.CASE_DETAILS_COLS.CASE_NAME_FOR_SEARCH - 1,
      creditorName: UNICOURT_CONFIG.CASE_DETAILS_COLS.INPUT_CREDITOR_NAME - 1,
      isBusiness: UNICOURT_CONFIG.CASE_DETAILS_COLS.IS_BUSINESS - 1,
      creditorType: UNICOURT_CONFIG.CASE_DETAILS_COLS.CREDITOR_TYPE - 1,
      statusColumn: UNICOURT_CONFIG.CASE_DETAILS_COLS.OVERALL_STATUS,
      lastAttemptColumn: UNICOURT_CONFIG.CASE_DETAILS_COLS.LAST_SUBMITTED_SHEET
    };
  } else if (activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.SAMPLE_RESEARCHED_CASE) {
    const requiredHeaders = [
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_NUMBER,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_TITLE,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_PARSED_FIRST_NAME,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_PARSED_LAST_NAME,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_BUSINESS_NAME,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_TYPE
    ];
    headerMap_SampleResearched = getHeadersAndIndices(activeSheet, requiredHeaders);
    if (!headerMap_SampleResearched) {
      // getHeadersAndIndices already logged the specific missing headers.
      // Use ui.alert here since this function is menu-driven
      try {
        ui.alert(`Cannot proceed: Missing required headers in "${activeSheetName}". Check logs for details. Please ensure headers are correct and try again.`);
      } catch (e) { // If UI context is somehow lost, log it.
        Logger.log(`UI ALERT FAILED in submitSelectedCasesToBackend (Header Error): Cannot proceed: Missing required headers in "${activeSheetName}".`);
      }
      return;
    }
  } else {
    showUserMessage(
        `Please select cells in a supported sheet ("${UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS}", "${UNICOURT_CONFIG.SHEET_NAMES.CASE_DETAILS}", or "${UNICOURT_CONFIG.SHEET_NAMES.SAMPLE_RESEARCHED_CASE}").`,
        "Unsupported Sheet",
        true
    );
    return;
  }

  const data = activeRange.getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    let caseData = {};

    if (activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.SAMPLE_RESEARCHED_CASE) {
      // headerMap_SampleResearched is guaranteed to be non-null here if we reached this point
      const caseNumberForDbId = String(row[headerMap_SampleResearched[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_NUMBER]] || "").trim();
      const caseNameForSearch = String(row[headerMap_SampleResearched[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_TITLE]] || "").trim();
      const creditorFirstName = String(row[headerMap_SampleResearched[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_PARSED_FIRST_NAME]] || "").trim();
      const creditorLastName = String(row[headerMap_SampleResearched[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_PARSED_LAST_NAME]] || "").trim();
      const creditorBusinessName = String(row[headerMap_SampleResearched[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_BUSINESS_NAME]] || "").trim();
      const creditorType = String(row[headerMap_SampleResearched[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_TYPE]] || "").trim();

      let inputCreditorName;
      let isBusiness = false;
      if (creditorBusinessName) {
        inputCreditorName = creditorBusinessName;
        isBusiness = true;
      } else {
        inputCreditorName = `${creditorFirstName} ${creditorLastName}`.trim();
      }

      if (caseNumberForDbId && caseNameForSearch && inputCreditorName && creditorType) {
        caseData = {
          case_number_for_db_id: caseNumberForDbId,
          case_name_for_search: caseNameForSearch,
          input_creditor_name: inputCreditorName,
          is_business: isBusiness,
          creditor_type: creditorType
        };
      }
    } else { // Submissions or Unicourt Processor Case Details sheet
      const caseNumberForDbId = String(row[columnMapping.caseNumber] || "").trim();
      const caseNameForSearch = String(row[columnMapping.caseName] || "").trim();
      const inputCreditorName = String(row[columnMapping.creditorName] || "").trim();
      const isBusinessStr = String(row[columnMapping.isBusiness] || "FALSE").trim().toUpperCase();
      const isBusiness = isBusinessStr === 'TRUE';
      const creditorType = String(row[columnMapping.creditorType] || "").trim();

      if (caseNumberForDbId && caseNameForSearch && inputCreditorName && creditorType) {
        caseData = {
          case_number_for_db_id: caseNumberForDbId,
          case_name_for_search: caseNameForSearch,
          input_creditor_name: inputCreditorName,
          is_business: isBusiness,
          creditor_type: creditorType
        };
      }
    }    // Validate case data before adding it to submission batch
    if (caseData.case_number_for_db_id) {
      const validation = validateCaseData(caseData);
      if (validation.isValid) {
        casesToSubmit.push(caseData);
        submissionSourceInfo.push({
          sheetRowIndex: activeRange.getRow() + i, // 1-based sheet row
          originalCaseData: caseData,
          isValid: true
        });
      } else {
        Logger.log(`Validation failed for case ${caseData.case_number_for_db_id}: ${validation.error}`);
        validationErrors.push(`Case ${caseData.case_number_for_db_id}: ${validation.error}`);
        invalidCases.push(caseData.case_number_for_db_id);
        // Still track invalid cases for proper status update
        submissionSourceInfo.push({
          sheetRowIndex: activeRange.getRow() + i, // 1-based sheet row
          originalCaseData: caseData,
          isValid: false
        });
      }
    }
  }  if (casesToSubmit.length === 0) {
    let message = "No valid cases selected or extracted from the selection. ";
    if (validationErrors.length > 0) {
      message += "\n\nSubmission Errors:\n" + validationErrors.join("\n");
    }
    message += "\n\nEnsure all required fields are filled.";
    showUserMessage(message, "No Cases to Submit", true);
    return;
  }
  
  if (validationErrors.length > 0) {
    showUserMessage(`Found ${validationErrors.length} invalid cases that will be skipped:\n\n${validationErrors.join("\n")}\n\nProceeding with ${casesToSubmit.length} valid cases.`, "Validation Warnings", false);
  }

  showUserToast(`Found ${casesToSubmit.length} case(s) to submit. Processing in batches of ${UNICOURT_CONFIG.BATCH_SIZE}...`, "Processing...", 10);
  
  // Use the new batching function
  const response = submitCasesInBatches(casesToSubmit, submissionSourceInfo, "Selected");  
  const currentTime = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd HH:mm:ss z");
  const submittedCaseNumbersForRefresh = casesToSubmit.map(c => c.case_number_for_db_id).filter(cn => cn && String(cn).trim() !== "");

  if (response.success && response.data) {
    let successMsg = `Submission results: ${response.data.submitted_cases} submitted. ` +
                     `${response.data.deleted_and_resubmitted_cases} replaced. ` +
                     `${response.data.already_queued_or_processing} skipped (active/queued). ` +
                     `Queue size: ${response.data.current_queue_size}.`;
    
    if (response.data.batches_processed) {
      successMsg += ` Processed in ${response.data.batches_processed} batch(es).`;
    }
    
    showUserMessage(successMsg, 'Submission Processed');    // Update status in active sheet (Submissions or Unicourt Processor Case Details)
    if (activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS || activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.CASE_DETAILS) {
      submissionSourceInfo.forEach(info => {
        if (info.isValid) {
          // Only update valid cases that were actually submitted
          activeSheet.getRange(info.sheetRowIndex, columnMapping.statusColumn).setValue('Submitted to Backend');
          activeSheet.getRange(info.sheetRowIndex, columnMapping.lastAttemptColumn).setValue(currentTime);
        } else {
          // Mark invalid cases with Submission Error
          activeSheet.getRange(info.sheetRowIndex, columnMapping.statusColumn).setValue('Submission Error');
          activeSheet.getRange(info.sheetRowIndex, columnMapping.lastAttemptColumn).setValue(currentTime);
        }
      });
      if (activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.CASE_DETAILS) {
        const submissionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS);
        if (submissionsSheet) {
          const caseNumbersFromDetails = submissionSourceInfo.map(info => info.originalCaseData.case_number_for_db_id);
          const matchingSubmissionsRows = findMatchingSubmissionsRows(submissionsSheet, caseNumbersFromDetails);
          matchingSubmissionsRows.forEach(subInfo => {
            submissionsSheet.getRange(subInfo.sheetRowIndex, UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS).setValue('Submitted to Backend');
            submissionsSheet.getRange(subInfo.sheetRowIndex, UNICOURT_CONFIG.SUBMISSIONS_COLS.LAST_SUBMIT_ATTEMPT).setValue(currentTime);
          });
        }
      }
    } else if (activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.SAMPLE_RESEARCHED_CASE) {
      const submissionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS);
      if (submissionsSheet) {
        const submissionsData = submissionsSheet.getDataRange().getValues();
        const caseNumColSubmissions = UNICOURT_CONFIG.SUBMISSIONS_COLS.CASE_NUMBER_FOR_DB_ID - 1;        submissionSourceInfo.forEach(info => {
          if (info.isValid) {
            const caseData = info.originalCaseData;
            let foundRowInSubmissions = -1;
            for (let j = 1; j < submissionsData.length; j++) {
              if (String(submissionsData[j][caseNumColSubmissions]).trim() === caseData.case_number_for_db_id) {
                foundRowInSubmissions = j + 1;
                break;
              }
            }
            if (foundRowInSubmissions !== -1) {
              submissionsSheet.getRange(foundRowInSubmissions, UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS).setValue('Submitted to Backend');
              submissionsSheet.getRange(foundRowInSubmissions, UNICOURT_CONFIG.SUBMISSIONS_COLS.LAST_SUBMIT_ATTEMPT).setValue(currentTime);          
            } else {
              // Only add to submission sheet if case data is valid
              const validation = validateCaseData(caseData);
              if (validation.isValid) {
                const newRowData = [
                  caseData.case_number_for_db_id, caseData.case_name_for_search, caseData.input_creditor_name,
                  caseData.is_business, caseData.creditor_type, 'Submitted to Backend', currentTime, ""
                ];
                submissionsSheet.appendRow(newRowData);
              } else {
                Logger.log(`Skipping invalid case ${caseData.case_number_for_db_id}: ${validation.error}`);
              }
            }
          } else {
            // Handle invalid cases - mark them as Submission Error in submissions sheet
            const caseData = info.originalCaseData;
            let foundRowInSubmissions = -1;
            for (let j = 1; j < submissionsData.length; j++) {
              if (String(submissionsData[j][caseNumColSubmissions]).trim() === caseData.case_number_for_db_id) {
                foundRowInSubmissions = j + 1;
                break;
              }
            }
            if (foundRowInSubmissions !== -1) {
              submissionsSheet.getRange(foundRowInSubmissions, UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS).setValue('Submission Error');
              submissionsSheet.getRange(foundRowInSubmissions, UNICOURT_CONFIG.SUBMISSIONS_COLS.LAST_SUBMIT_ATTEMPT).setValue(currentTime);          
            } else {
              const newRowData = [
                caseData.case_number_for_db_id, caseData.case_name_for_search || "(Error)", caseData.input_creditor_name || "(Error)",
                caseData.is_business || false, caseData.creditor_type || "(Error)", 'Submission Error', currentTime, ""
              ];
              submissionsSheet.appendRow(newRowData);
            }
          }
        });
      }
    }    showUserToast("Cases submitted. Refreshing data after a few moments...", "Status", 5);
    Utilities.sleep(3000);
    refreshAllCaseData(submittedCaseNumbersForRefresh);

  } else if (response.partial_success) {    // Handle partial success - some batches succeeded, some failed
    let partialMsg = `Partial success: ${response.successful_batches}/${response.successful_batches + response.failed_batches} batches completed successfully. `;
    if (response.data) {
      partialMsg += `${response.data.submitted_cases} submitted, ${response.data.deleted_and_resubmitted_cases} replaced, ${response.data.already_queued_or_processing} skipped. `;
    }
    partialMsg += `Errors: ${response.detail}`;
    
    showUserMessage(partialMsg, 'Partial Success', true);
    
    // Only update submission sheet for cases that passed validation
    const validCases = casesToSubmit.filter(caseData => validateCaseData(caseData).isValid);
    
    // Update successful submissions and refresh
    if (submittedCaseNumbersForRefresh.length > 0) {
      showUserToast("Updating sheet statuses and refreshing data...", "Processing", 5);
      Utilities.sleep(2000);
      refreshAllCaseData(submittedCaseNumbersForRefresh);
    }

  } else {
    const errorMsg = `Error: ${response.error || 'Unknown error'} (Detail: ${response.detail || ''}).`;
    showUserMessage(`${errorMsg} Statuses updated to reflect error.`, 'Submission Failed', true);    if (activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS || activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.CASE_DETAILS) {
      submissionSourceInfo.forEach(info => {
        if (info.isValid) {
          // Only mark valid cases that were attempted to be submitted
          activeSheet.getRange(info.sheetRowIndex, columnMapping.statusColumn).setValue('Submission Error');
          activeSheet.getRange(info.sheetRowIndex, columnMapping.lastAttemptColumn).setValue(currentTime);
        } else {
          // Invalid cases should remain as Submission Error
          activeSheet.getRange(info.sheetRowIndex, columnMapping.statusColumn).setValue('Submission Error');
          activeSheet.getRange(info.sheetRowIndex, columnMapping.lastAttemptColumn).setValue(currentTime);
        }
      });
    } else if (activeSheetName === UNICOURT_CONFIG.SHEET_NAMES.SAMPLE_RESEARCHED_CASE) {
      const submissionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS);
      if (submissionsSheet) {
        const submissionsData = submissionsSheet.getDataRange().getValues();
        const caseNumColSubmissions = UNICOURT_CONFIG.SUBMISSIONS_COLS.CASE_NUMBER_FOR_DB_ID - 1;        submissionSourceInfo.forEach(info => {
          const caseData = info.originalCaseData;
          let foundRowInSubmissions = -1;
          for (let j = 1; j < submissionsData.length; j++) {
            if (String(submissionsData[j][caseNumColSubmissions]).trim() === caseData.case_number_for_db_id) {
              foundRowInSubmissions = j + 1;
              break;
            }
          }
          if (foundRowInSubmissions !== -1) {
            if (info.isValid) {
              submissionsSheet.getRange(foundRowInSubmissions, UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS).setValue('Submission Error');
            } else {
              submissionsSheet.getRange(foundRowInSubmissions, UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS).setValue('Submission Error');
            }
            submissionsSheet.getRange(foundRowInSubmissions, UNICOURT_CONFIG.SUBMISSIONS_COLS.LAST_SUBMIT_ATTEMPT).setValue(currentTime);
          } else {
            const statusValue = info.isValid ? 'Submission Error' : 'Submission Error';
            const newRowData = [
              caseData.case_number_for_db_id, 
              caseData.case_name_for_search || "(Error)", 
              caseData.input_creditor_name || "(Error)",
              caseData.is_business || false, 
              caseData.creditor_type || "(Error)", 
              statusValue, 
              currentTime, 
              ""
            ];
            submissionsSheet.appendRow(newRowData);
          }
        });
      }
    }
    // Refresh even on error to update local views if possible
    if (submittedCaseNumbersForRefresh.length > 0) {
        showUserToast("Attempting to refresh submitted cases despite backend error...", "Status", 5);
        Utilities.sleep(1000);
        refreshAllCaseData(submittedCaseNumbersForRefresh);
    }
  }
}

/**
 * Find rows in Submissions sheet that match given case numbers
 * Returns array of { sheetRowIndex: number (1-based), caseNum: string }
 */
function findMatchingSubmissionsRows(submissionsSheet, caseNumbers) {
  if (!submissionsSheet || caseNumbers.length === 0) return [];
  const submissionsData = submissionsSheet.getDataRange().getValues();
  const matchingRows = [];
  const caseNumCol = UNICOURT_CONFIG.SUBMISSIONS_COLS.CASE_NUMBER_FOR_DB_ID - 1;
  
  for (let i = 1; i < submissionsData.length; i++) { // Skip header
    const currentCaseNumber = String(submissionsData[i][caseNumCol] || "").trim();
    if (currentCaseNumber && caseNumbers.includes(currentCaseNumber)) {
      matchingRows.push({
        sheetRowIndex: i + 1, 
        caseNum: currentCaseNumber
      });
    }  }
  return matchingRows;
}

/**
 * Manual version of autoCheckAndSubmitCases - searches for new cases in research sheet and submits them.
 * This is the menu-triggered version that provides user feedback via toasts and alerts.
 */
function manualSearchAndSubmitNewCases() {
  showUserToast("Searching for new cases to submit...", "Processing", 5);
  Logger.log("Manual Search & Submit: Starting check...");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const submissionsSheet = ss.getSheetByName(UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS);
  const sampleResearchedSheet = ss.getSheetByName(UNICOURT_CONFIG.SHEET_NAMES.SAMPLE_RESEARCHED_CASE);

  if (!submissionsSheet) {
    showUserMessage(
      "Required sheet 'Unicourt Processor Case Submissions' not found. Please run 'Ensure/Reset Sheet Headers & Structure' first.",
      "Sheet Not Found",
      true
    );
    return;
  }

  const casesToSubmit = [];
  const submissionSourceInfo = []; // To track origin for sheet updates
  const submittedCaseNumbersSet = new Set();

  // 1. Process "Unicourt Processor Case Submissions" sheet for new/errored cases
  Logger.log("Manual Search & Submit: Processing 'Unicourt Processor Case Submissions' sheet...");
  if (submissionsSheet.getLastRow() > 1) {
    const submissionsDataRange = submissionsSheet.getRange(2, 1, submissionsSheet.getLastRow() - 1, submissionsSheet.getLastColumn());
    const submissionsData = submissionsDataRange.getValues();

    for (let i = 0; i < submissionsData.length; i++) {
      const row = submissionsData[i];
      const caseNumberForDbId = String(row[UNICOURT_CONFIG.SUBMISSIONS_COLS.CASE_NUMBER_FOR_DB_ID - 1] || "").trim();

      if (caseNumberForDbId) {
        submittedCaseNumbersSet.add(caseNumberForDbId);
      }

      const submissionStatus = String(row[UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS - 1] || "").trim().toLowerCase();
      if (!submissionStatus || submissionStatus === 'submission error') {
        const caseNameForSearch = String(row[UNICOURT_CONFIG.SUBMISSIONS_COLS.CASE_NAME_FOR_SEARCH - 1] || "").trim();
        const inputCreditorName = String(row[UNICOURT_CONFIG.SUBMISSIONS_COLS.INPUT_CREDITOR_NAME - 1] || "").trim();        
        const isBusinessStr = String(row[UNICOURT_CONFIG.SUBMISSIONS_COLS.IS_BUSINESS - 1] || "FALSE").trim().toUpperCase();
        const isBusiness = isBusinessStr === 'TRUE';
        const creditorType = String(row[UNICOURT_CONFIG.SUBMISSIONS_COLS.CREDITOR_TYPE - 1] || "").trim();

        if (caseNumberForDbId && caseNameForSearch && inputCreditorName && creditorType) {
          const caseData = {
            case_number_for_db_id: caseNumberForDbId,
            case_name_for_search: caseNameForSearch,
            input_creditor_name: inputCreditorName,
            is_business: isBusiness,
            creditor_type: creditorType
          };
          
          // Validate case data before adding to submission batch
          const validation = validateCaseData(caseData);
          if (validation.isValid) {
            casesToSubmit.push(caseData);
            submissionSourceInfo.push({
              source: 'Submissions',
              sheetRowIndex: i + 2, // 1-based sheet row
              originalCaseData: caseData,
              isValid: true
            });
          } else {
            Logger.log(`Manual Search & Submit: Validation failed for case ${caseData.case_number_for_db_id}: ${validation.error}`);
            submissionSourceInfo.push({
              source: 'Submissions',
              sheetRowIndex: i + 2, // 1-based sheet row
              originalCaseData: caseData,
              isValid: false,
              validationError: validation.error
            });
          }
        }
      }
    }
    Logger.log(`Manual Search & Submit: Found ${submissionSourceInfo.filter(s => s.source === 'Submissions').length} cases from 'Unicourt Processor Case Submissions' to process.`);
  } else {
    Logger.log("Manual Search & Submit: 'Unicourt Processor Case Submissions' sheet has no data rows (or only header).");
  }

  // 2. Process "Research" sheet for cases NOT in "Unicourt Processor Case Submissions"
  if (sampleResearchedSheet) {
    Logger.log("Manual Search & Submit: Processing 'Research' sheet...");
    showUserToast("Searching research sheet for new cases...", "Processing", 3);
    
    const requiredHeadersSample = [
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_NUMBER,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_TITLE,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_PARSED_FIRST_NAME,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_PARSED_LAST_NAME,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_BUSINESS_NAME,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_TYPE
    ];
    const headerMap_Sample = getHeadersAndIndices(sampleResearchedSheet, requiredHeadersSample);

    if (headerMap_Sample) {
      if (sampleResearchedSheet.getLastRow() > 1) {
        const sampleDataRange = sampleResearchedSheet.getRange(2, 1, sampleResearchedSheet.getLastRow() - 1, sampleResearchedSheet.getLastColumn());
        const sampleData = sampleDataRange.getValues();
        let foundFromSample = 0;

        sampleData.forEach((row, rowIndex) => {
          const caseNumberForDbId = String(row[headerMap_Sample[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_NUMBER]] || "").trim();

          if (caseNumberForDbId && !submittedCaseNumbersSet.has(caseNumberForDbId)) {
            const caseNameForSearch = String(row[headerMap_Sample[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_TITLE]] || "").trim();
            const creditorFirstName = String(row[headerMap_Sample[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_PARSED_FIRST_NAME]] || "").trim();
            const creditorLastName = String(row[headerMap_Sample[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_PARSED_LAST_NAME]] || "").trim();
            const creditorBusinessName = String(row[headerMap_Sample[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_BUSINESS_NAME]] || "").trim();
            const creditorType = String(row[headerMap_Sample[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_TYPE]] || "").trim();            
            let inputCreditorName;
            let isBusiness = false;
            if (creditorBusinessName) {
              inputCreditorName = creditorBusinessName;
              isBusiness = true;
            } else {
              inputCreditorName = `${creditorFirstName} ${creditorLastName}`.trim();
            }

            if (caseNameForSearch && inputCreditorName && creditorType) {
              const caseData = {
                case_number_for_db_id: caseNumberForDbId,
                case_name_for_search: caseNameForSearch,
                input_creditor_name: inputCreditorName,
                is_business: isBusiness,
                creditor_type: creditorType
              };
              
              // Validate case data before adding to submission batch
              const validation = validateCaseData(caseData);
              if (validation.isValid) {
                casesToSubmit.push(caseData);
                submissionSourceInfo.push({
                  source: 'SampleResearched',
                  sheetRowIndex: rowIndex + 2,
                  originalCaseData: caseData,
                  isValid: true
                });
                foundFromSample++;
              } else {
                Logger.log(`Manual Search & Submit: Validation failed for Research case ${caseData.case_number_for_db_id}: ${validation.error}`);
                submissionSourceInfo.push({
                  source: 'SampleResearched',
                  sheetRowIndex: rowIndex + 2,
                  originalCaseData: caseData,
                  isValid: false,
                  validationError: validation.error
                });
              }
            }
          }
        });
        Logger.log(`Manual Search & Submit: Found ${foundFromSample} new cases from 'Research' to submit.`);
      } else {
        Logger.log("Manual Search & Submit: 'Research' sheet has no data rows (or only header).");
      }
    } else {
      showUserMessage(
        "Could not process 'Research' sheet due to missing required headers. Please ensure the research sheet has the correct headers.",
        "Header Error",
        true
      );
      Logger.log("Manual Search & Submit: Could not process 'Research' due to missing headers. Skipping this sheet for this run.");
    }
  } else {
    Logger.log("Manual Search & Submit: 'Research' sheet not found. Skipping.");
  }

  // 3. Submit collected cases
  if (casesToSubmit.length === 0) {
    showUserMessage(
      "No new cases found to submit from any source. All cases in the research sheet may already be submitted or missing required data.",
      "No New Cases",
      false
    );
    Logger.log("Manual Search & Submit: No new or errored cases to submit from any source.");
    return;
  }  showUserToast(`Found ${casesToSubmit.length} new case(s) to submit. Processing in batches of ${UNICOURT_CONFIG.BATCH_SIZE}...`, "Processing", 10);
  Logger.log(`Manual Search & Submit: Attempting to submit a total of ${casesToSubmit.length} case(s) in batches of ${UNICOURT_CONFIG.BATCH_SIZE}.`);
    // Use the batching function instead of direct submission
  const response = submitCasesInBatches(casesToSubmit, submissionSourceInfo, "Manual");
  const currentTime = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd HH:mm:ss z");
  const submittedCaseNumbers = casesToSubmit.map(c => c.case_number_for_db_id).filter(cn => cn && String(cn).trim() !== "");

  // 4. Update sheets based on submission response
  if (response.success && response.data) {    Logger.log(`Manual Search & Submit: Backend Success. Submitted: ${response.data.submitted_cases}, Replaced: ${response.data.deleted_and_resubmitted_cases}, Skipped: ${response.data.already_queued_or_processing}`);

    let successMsg = `Submission results: ${response.data.submitted_cases} submitted, ` +
                      `${response.data.deleted_and_resubmitted_cases} replaced, ` +
                      `${response.data.already_queued_or_processing} skipped (already active/queued). ` +
                      `Queue size: ${response.data.current_queue_size}.`;
      if (response.data.batches_processed) {
      successMsg += ` Processed in ${response.data.batches_processed} batch(es).`;
    }

    submissionSourceInfo.forEach(info => {
      if (info.source === 'Submissions') {
        // Check validation status - only mark valid cases as submitted
        const statusValue = info.isValid ? 'Submitted to Backend' : 'Submission Error';
        submissionsSheet.getRange(info.sheetRowIndex, UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS).setValue(statusValue);
        submissionsSheet.getRange(info.sheetRowIndex, UNICOURT_CONFIG.SUBMISSIONS_COLS.LAST_SUBMIT_ATTEMPT).setValue(currentTime);
      } else if (info.source === 'SampleResearched') {
        const caseData = info.originalCaseData;
        // Check validation status - only mark valid cases as submitted
        const statusValue = info.isValid ? 'Submitted to Backend' : 'Submission Error';
        const newRowData = [
          caseData.case_number_for_db_id, caseData.case_name_for_search, caseData.input_creditor_name,
          caseData.is_business, caseData.creditor_type, statusValue, currentTime, ""
        ];
        submissionsSheet.appendRow(newRowData);
      }
    });    showUserMessage(successMsg, 'Search & Submit Complete');
    
    showUserToast("Cases submitted successfully. Refreshing data...", "Processing", 5);
    Logger.log("Manual Search & Submit: Triggering data refresh after successful submission(s).");
    Utilities.sleep(3000);
    refreshAllCaseData(submittedCaseNumbers);
    
  } else if (response.partial_success) {
    // Handle partial success - some batches succeeded, some failed
    let partialMsg = `Partial success: ${response.successful_batches}/${response.successful_batches + response.failed_batches} batches completed successfully. `;
    if (response.data) {
      partialMsg += `${response.data.submitted_cases} submitted, ${response.data.deleted_and_resubmitted_cases} replaced, ${response.data.already_queued_or_processing} skipped. `;
    }
    partialMsg += `Errors: ${response.detail}`;
    
    showUserMessage(partialMsg, 'Partial Success', true);
      // Update successful submissions only
   
    submissionSourceInfo.forEach(info => {
      if (info.source === 'Submissions') {
        // For partial success: valid cases that were attempted get 'Submitted to Backend', invalid get 'Submission Error'  
        const statusValue = info.isValid ? 'Submitted to Backend' : 'Submission Error';
        submissionsSheet.getRange(info.sheetRowIndex, UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS).setValue(statusValue);
        submissionsSheet.getRange(info.sheetRowIndex, UNICOURT_CONFIG.SUBMISSIONS_COLS.LAST_SUBMIT_ATTEMPT).setValue(currentTime);
      } else if (info.source === 'SampleResearched') {
        const caseData = info.originalCaseData;
        // For partial success: valid cases that were attempted get 'Submitted to Backend', invalid get 'Submission Error'  
        const statusValue = info.isValid ? 'Submitted to Backend' : 'Submission Error';
        const newRowData = [
          caseData.case_number_for_db_id, caseData.case_name_for_search, caseData.input_creditor_name,
          caseData.is_business, caseData.creditor_type, statusValue, currentTime, ""
        ];
        submissionsSheet.appendRow(newRowData);
      }
    });
    
    // Refresh data for successful submissions
    if (submittedCaseNumbers.length > 0) {
      showUserToast("Updating sheet statuses and refreshing data...", "Processing", 5);
      Utilities.sleep(2000);
      refreshAllCaseData(submittedCaseNumbers);
    }
    
  } else {
    const errorMsg = `Submission failed: ${response.error || 'Unknown error'}. Detail: ${response.detail || ''}`;
    Logger.log(`Manual Search & Submit: Backend Submission failed. Error: ${response.error || 'Unknown error'}. Detail: ${response.detail || ''}`);
      submissionSourceInfo.forEach(info => {
      if (info.source === 'Submissions') {
        // Check validation status - only mark valid cases as submission error, invalid as Submission Error
        const statusValue = info.isValid ? 'Submission Error' : 'Submission Error';
        submissionsSheet.getRange(info.sheetRowIndex, UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS).setValue(statusValue);
        submissionsSheet.getRange(info.sheetRowIndex, UNICOURT_CONFIG.SUBMISSIONS_COLS.LAST_SUBMIT_ATTEMPT).setValue(currentTime);
      } else if (info.source === 'SampleResearched') {
        const caseData = info.originalCaseData;
        let foundRowInSubmissions = -1;
        const currentSubmissionsData = submissionsSheet.getDataRange().getValues();
        for (let j = 1; j < currentSubmissionsData.length; j++) {
          if (String(currentSubmissionsData[j][UNICOURT_CONFIG.SUBMISSIONS_COLS.CASE_NUMBER_FOR_DB_ID - 1]).trim() === caseData.case_number_for_db_id) {
            foundRowInSubmissions = j + 1;
            break;
          }
        }
        if (foundRowInSubmissions !== -1) {
          // Check validation status - only mark valid cases as submission error, invalid as Submission Error
          const statusValue = info.isValid ? 'Submission Error' : 'Submission Error';
          submissionsSheet.getRange(foundRowInSubmissions, UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS).setValue(statusValue);
          submissionsSheet.getRange(foundRowInSubmissions, UNICOURT_CONFIG.SUBMISSIONS_COLS.LAST_SUBMIT_ATTEMPT).setValue(currentTime);
        } else {
          // Check validation status - only mark valid cases as submission error, invalid as Submission Error
          const statusValue = info.isValid ? 'Submission Error' : 'Submission Error';
          const newRowData = [
            caseData.case_number_for_db_id, caseData.case_name_for_search, caseData.input_creditor_name,
            caseData.is_business, caseData.creditor_type, statusValue, currentTime, ""
          ];
          submissionsSheet.appendRow(newRowData);
        }
      }
    });

    showUserMessage(`${errorMsg} Status updated to reflect error.`, 'Submission Failed', true);
    
    // Refresh even on error to update local views if possible
    if (submittedCaseNumbers.length > 0) {
        showUserToast("Attempting to refresh case data despite submission error...", "Processing", 5);
        Utilities.sleep(1000);
        refreshAllCaseData(submittedCaseNumbers);
    }
  }
  Logger.log("Manual Search & Submit: Process finished.");
}


function autoCheckAndSubmitCases() {
  Logger.log("Auto-Submit: Starting check...");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const submissionsSheet = ss.getSheetByName(UNICOURT_CONFIG.SHEET_NAMES.SUBMISSIONS);
  const sampleResearchedSheet = ss.getSheetByName(UNICOURT_CONFIG.SHEET_NAMES.SAMPLE_RESEARCHED_CASE);

  if (!submissionsSheet) {
    Logger.log("Auto-Submit: 'Unicourt Processor Case Submissions' sheet not found. Aborting.");
    return;
  }

  const casesToSubmit = [];
  const submissionSourceInfo = []; // To track origin for sheet updates
  const submittedCaseNumbersSet = new Set();

  // 1. Process "Unicourt Processor Case Submissions" sheet for new/errored cases
  Logger.log("Auto-Submit: Processing 'Unicourt Processor Case Submissions' sheet...");
  if (submissionsSheet.getLastRow() > 1) {
    const submissionsDataRange = submissionsSheet.getRange(2, 1, submissionsSheet.getLastRow() - 1, submissionsSheet.getLastColumn());
    const submissionsData = submissionsDataRange.getValues();

    for (let i = 0; i < submissionsData.length; i++) {
      const row = submissionsData[i];
      const caseNumberForDbId = String(row[UNICOURT_CONFIG.SUBMISSIONS_COLS.CASE_NUMBER_FOR_DB_ID - 1] || "").trim();

      if (caseNumberForDbId) {
        submittedCaseNumbersSet.add(caseNumberForDbId);
      }

      const submissionStatus = String(row[UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS - 1] || "").trim().toLowerCase();
      if (!submissionStatus || submissionStatus === 'submission error') {
        const caseNameForSearch = String(row[UNICOURT_CONFIG.SUBMISSIONS_COLS.CASE_NAME_FOR_SEARCH - 1] || "").trim();
        const inputCreditorName = String(row[UNICOURT_CONFIG.SUBMISSIONS_COLS.INPUT_CREDITOR_NAME - 1] || "").trim();
        const isBusinessStr = String(row[UNICOURT_CONFIG.SUBMISSIONS_COLS.IS_BUSINESS - 1] || "FALSE").trim().toUpperCase();
        const isBusiness = isBusinessStr === 'TRUE';
        const creditorType = String(row[UNICOURT_CONFIG.SUBMISSIONS_COLS.CREDITOR_TYPE - 1] || "").trim();        if (caseNumberForDbId && caseNameForSearch && inputCreditorName && creditorType) {
          const caseData = {
            case_number_for_db_id: caseNumberForDbId,
            case_name_for_search: caseNameForSearch,
            input_creditor_name: inputCreditorName,
            is_business: isBusiness,
            creditor_type: creditorType
          };
          
          // Validate case data before adding to submission batch
          const validation = validateCaseData(caseData);
          if (validation.isValid) {
            casesToSubmit.push(caseData);
            submissionSourceInfo.push({
              source: 'Submissions',
              sheetRowIndex: i + 2, // 1-based sheet row
              originalCaseData: caseData,
              isValid: true
            });
          } else {
            Logger.log(`Auto-Submit: Validation failed for case ${caseData.case_number_for_db_id}: ${validation.error}`);
            submissionSourceInfo.push({
              source: 'Submissions',
              sheetRowIndex: i + 2, // 1-based sheet row
              originalCaseData: caseData,
              isValid: false,
              validationError: validation.error
            });
          }
        }
      }
    }
    Logger.log(`Auto-Submit: Found ${submissionSourceInfo.filter(s => s.source === 'Submissions').length} cases from 'Unicourt Processor Case Submissions' to process.`);
  } else {
    Logger.log("Auto-Submit: 'Unicourt Processor Case Submissions' sheet has no data rows (or only header).");
  }

  // 2. Process "Research" sheet for cases NOT in "Unicourt Processor Case Submissions"
  if (sampleResearchedSheet) {
    Logger.log("Auto-Submit: Processing 'Research' sheet...");
    const requiredHeadersSample = [
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_NUMBER,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_TITLE,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_PARSED_FIRST_NAME,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_PARSED_LAST_NAME,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_BUSINESS_NAME,
      UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_TYPE
    ];
    const headerMap_Sample = getHeadersAndIndices(sampleResearchedSheet, requiredHeadersSample);

    if (headerMap_Sample) {
      if (sampleResearchedSheet.getLastRow() > 1) {
        const sampleDataRange = sampleResearchedSheet.getRange(2, 1, sampleResearchedSheet.getLastRow() - 1, sampleResearchedSheet.getLastColumn());
        const sampleData = sampleDataRange.getValues();
        let foundFromSample = 0;

        sampleData.forEach((row, rowIndex) => {
          const caseNumberForDbId = String(row[headerMap_Sample[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_NUMBER]] || "").trim();

          if (caseNumberForDbId && !submittedCaseNumbersSet.has(caseNumberForDbId)) {
            const caseNameForSearch = String(row[headerMap_Sample[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.COURT_CASE_TITLE]] || "").trim();
            const creditorFirstName = String(row[headerMap_Sample[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_PARSED_FIRST_NAME]] || "").trim();
            const creditorLastName = String(row[headerMap_Sample[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_PARSED_LAST_NAME]] || "").trim();
            const creditorBusinessName = String(row[headerMap_Sample[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_BUSINESS_NAME]] || "").trim();
            const creditorType = String(row[headerMap_Sample[UNICOURT_CONFIG.SAMPLE_RESEARCHED_CASE_HEADERS.CREDITOR_TYPE]] || "").trim();

            let inputCreditorName;
            let isBusiness = false;
            if (creditorBusinessName) {
              inputCreditorName = creditorBusinessName;
              isBusiness = true;
            } else {
              inputCreditorName = `${creditorFirstName} ${creditorLastName}`.trim();
            }            if (caseNameForSearch && inputCreditorName && creditorType) {
              const caseData = {
                case_number_for_db_id: caseNumberForDbId,
                case_name_for_search: caseNameForSearch,
                input_creditor_name: inputCreditorName,
                is_business: isBusiness,
                creditor_type: creditorType
              };
              
              // Validate case data before adding to submission batch
              const validation = validateCaseData(caseData);
              if (validation.isValid) {
                casesToSubmit.push(caseData);
                submissionSourceInfo.push({
                  source: 'SampleResearched',
                  sheetRowIndex: rowIndex + 2,
                  originalCaseData: caseData,
                  isValid: true
                });
                foundFromSample++;
              } else {
                Logger.log(`Auto-Submit: Validation failed for Research case ${caseData.case_number_for_db_id}: ${validation.error}`);
                submissionSourceInfo.push({
                  source: 'SampleResearched',
                  sheetRowIndex: rowIndex + 2,
                  originalCaseData: caseData,
                  isValid: false,
                  validationError: validation.error
                });
              }
            }
          }
        });
        Logger.log(`Auto-Submit: Found ${foundFromSample} new cases from 'Research' to submit.`);
      } else {
        Logger.log("Auto-Submit: 'Research' sheet has no data rows (or only header).");
      }
    } else {
      Logger.log("Auto-Submit: Could not process 'Research' due to missing headers. Skipping this sheet for this run.");
    }
  } else {
    Logger.log("Auto-Submit: 'Research' sheet not found. Skipping.");
  }
  // 3. Submit collected cases
  if (casesToSubmit.length === 0) {
    Logger.log("Auto-Submit: No new or errored cases to submit from any source.");
    return;
  }

  Logger.log(`Auto-Submit: Attempting to submit a total of ${casesToSubmit.length} case(s) in batches of ${UNICOURT_CONFIG.BATCH_SIZE}.`);
  
  // Use the new batching function  
  const response = submitCasesInBatches(casesToSubmit, submissionSourceInfo, "Auto");
  const currentTime = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd HH:mm:ss z");

  const autoSubmittedCaseNumbers = casesToSubmit.map(c => c.case_number_for_db_id).filter(cn => cn && String(cn).trim() !== "");

  // 4. Update sheets based on submission response
  if (response.success && response.data) {
    Logger.log(`Auto-Submit: Backend Success. Submitted: ${response.data.submitted_cases}, Replaced: ${response.data.deleted_and_resubmitted_cases}, Skipped: ${response.data.already_queued_or_processing}`);    submissionSourceInfo.forEach(info => {
      if (info.source === 'Submissions') {
        // Check validation status - only mark valid cases as submitted
        const statusValue = info.isValid ? 'Submitted to Backend' : 'Submission Error';
        submissionsSheet.getRange(info.sheetRowIndex, UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS).setValue(statusValue);
        submissionsSheet.getRange(info.sheetRowIndex, UNICOURT_CONFIG.SUBMISSIONS_COLS.LAST_SUBMIT_ATTEMPT).setValue(currentTime);
      } else if (info.source === 'SampleResearched') {
        const caseData = info.originalCaseData;
        // Check validation status - only mark valid cases as submitted
        const statusValue = info.isValid ? 'Submitted to Backend' : 'Submission Error';
        const newRowData = [
          caseData.case_number_for_db_id, caseData.case_name_for_search, caseData.input_creditor_name,
          caseData.is_business, caseData.creditor_type, statusValue, currentTime, ""
        ];
        submissionsSheet.appendRow(newRowData);
        submittedCaseNumbersSet.add(caseData.case_number_for_db_id);
      }
    });

    Logger.log("Auto-Submit: Triggering data refresh after successful submission(s).");
    Utilities.sleep(5000);
    refreshAllCaseData(autoSubmittedCaseNumbers);
  } else {
    Logger.log(`Auto-Submit: Backend Submission failed. Error: ${response.error || 'Unknown error'}. Detail: ${response.detail || ''}`);    submissionSourceInfo.forEach(info => {
      if (info.source === 'Submissions') {
        // Check validation status - only mark valid cases as submission error, invalid as Submission Error
        const statusValue = info.isValid ? 'Submission Error' : 'Submission Error';
        submissionsSheet.getRange(info.sheetRowIndex, UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS).setValue(statusValue);
        submissionsSheet.getRange(info.sheetRowIndex, UNICOURT_CONFIG.SUBMISSIONS_COLS.LAST_SUBMIT_ATTEMPT).setValue(currentTime);
      } else if (info.source === 'SampleResearched') {
        const caseData = info.originalCaseData;
        let foundRowInSubmissions = -1;
        const currentSubmissionsData = submissionsSheet.getDataRange().getValues();
        for (let j = 1; j < currentSubmissionsData.length; j++) {
          if (String(currentSubmissionsData[j][UNICOURT_CONFIG.SUBMISSIONS_COLS.CASE_NUMBER_FOR_DB_ID - 1]).trim() === caseData.case_number_for_db_id) {
            foundRowInSubmissions = j + 1;
            break;
          }
        }
        if (foundRowInSubmissions !== -1) {
          // Check validation status - only mark valid cases as submission error, invalid as Submission Error
          const statusValue = info.isValid ? 'Submission Error' : 'Submission Error';
          submissionsSheet.getRange(foundRowInSubmissions, UNICOURT_CONFIG.SUBMISSIONS_COLS.SUBMISSION_STATUS).setValue(statusValue);
          submissionsSheet.getRange(foundRowInSubmissions, UNICOURT_CONFIG.SUBMISSIONS_COLS.LAST_SUBMIT_ATTEMPT).setValue(currentTime);
        } else {
          // Check validation status - only mark valid cases as submission error, invalid as Submission Error
          const statusValue = info.isValid ? 'Submission Error' : 'Submission Error';
          const newRowData = [
            caseData.case_number_for_db_id, caseData.case_name_for_search, caseData.input_creditor_name,
            caseData.is_business, caseData.creditor_type, statusValue, currentTime, ""
          ];
          submissionsSheet.appendRow(newRowData);
          submittedCaseNumbersSet.add(caseData.case_number_for_db_id);
        }
      }
    });
    // Optionally refresh these auto-submitted cases even on error
    if (autoSubmittedCaseNumbers.length > 0) {
        Logger.log("Auto-Submit: Attempting to refresh auto-submitted cases despite backend error...");
        Utilities.sleep(1000);
        refreshAllCaseData(autoSubmittedCaseNumbers);
    }
  }
  Logger.log("Auto-Submit: Check finished.");
}

function createAutoSubmitTrigger(minutes) {
  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoCheckAndSubmitCases') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  PropertiesService.getScriptProperties().setProperty('AUTO_SUBMIT_INTERVAL', String(minutes));
  
  let triggerBuilder = ScriptApp.newTrigger('autoCheckAndSubmitCases').timeBased();
  let intervalDescription = '';
  
  // Handle different time intervals based on the value
  if (minutes <= 30 && [1, 5, 10, 15, 30].includes(minutes)) {
    // Use everyMinutes for valid minute intervals
    triggerBuilder = triggerBuilder.everyMinutes(minutes);
    intervalDescription = `${minutes} minute${minutes > 1 ? 's' : ''}`;
  } else if (minutes === 60) {
    // 1 hour
    triggerBuilder = triggerBuilder.everyHours(1);
    intervalDescription = '1 hour';
  } else if (minutes === 120) {
    // 2 hours
    triggerBuilder = triggerBuilder.everyHours(2);
    intervalDescription = '2 hours';
  } else if (minutes === 240) {
    // 4 hours
    triggerBuilder = triggerBuilder.everyHours(4);
    intervalDescription = '4 hours';
  } else if (minutes === 360) {
    // 6 hours
    triggerBuilder = triggerBuilder.everyHours(6);
    intervalDescription = '6 hours';
  } else if (minutes === 480) {
    // 8 hours
    triggerBuilder = triggerBuilder.everyHours(8);
    intervalDescription = '8 hours';
  } else if (minutes === 720) {
    // 12 hours
    triggerBuilder = triggerBuilder.everyHours(12);
    intervalDescription = '12 hours';
  } else if (minutes === 1440) {
    // 1 day
    triggerBuilder = triggerBuilder.everyDays(1);
    intervalDescription = '1 day';
  } else if (minutes === 2880) {
    // 2 days
    triggerBuilder = triggerBuilder.everyDays(2);
    intervalDescription = '2 days';
  } else if (minutes === 4320) {
    // 3 days
    triggerBuilder = triggerBuilder.everyDays(3);
    intervalDescription = '3 days';
  } else if (minutes === 10080) {
    // 1 week
    triggerBuilder = triggerBuilder.everyWeeks(1);
    intervalDescription = '1 week';
  } else {
    // Fallback for unsupported intervals - use closest supported interval
    throw new Error(`Unsupported interval: ${minutes} minutes. Please select a supported interval from the dropdown.`);
  }
  
  triggerBuilder.create();
  SpreadsheetApp.getUi().alert('Auto-Submit Enabled', `Automatic case submission enabled to run every ${intervalDescription}.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

function enableAutoSubmit() {
  const ui = SpreadsheetApp.getUi();
  const existingTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of existingTriggers) {
    if (trigger.getHandlerFunction() === 'autoCheckAndSubmitCases') {
      const interval = PropertiesService.getScriptProperties().getProperty('AUTO_SUBMIT_INTERVAL') || 'unknown interval';
      ui.alert('Auto-Submit Already Enabled', 
        `The auto-submit feature is already enabled to run at an interval of approximately ${interval} minutes. Please disable it first if you want to change the interval.`,
        ui.ButtonSet.OK);
      return;
    }
  }

  const html = HtmlService.createHtmlOutputFromFile('Unicourt Processor AutoSubmitIntervalDialog') // Assumes this HTML file exists or will be created
    .setWidth(600)
    .setHeight(400); // Adjusted height
  ui.showModalDialog(html, 'Configure Auto-Submit Interval');
}

function disableAutoSubmit() {
  const ui = SpreadsheetApp.getUi();
  let triggerFound = false;
  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoCheckAndSubmitCases') {
      ScriptApp.deleteTrigger(trigger);
      triggerFound = true;
    }
  });

  if (triggerFound) {
    PropertiesService.getScriptProperties().deleteProperty('AUTO_SUBMIT_INTERVAL');
    ui.alert('Auto-Submit Disabled', 'The automatic case submission has been disabled.', ui.ButtonSet.OK);
  } else {
    ui.alert('Auto-Submit Not Running', 'The auto-submit feature was not enabled.', ui.ButtonSet.OK);
  }
}

// --- Case Processing Functions ---

function getAllCasesFromBackend() {
  const ui = SpreadsheetApp.getUi();
  SpreadsheetApp.getActiveSpreadsheet().toast("Fetching all cases from backend...", "Processing", 5);

  const response = callBackendGetAllCases();

  if (response.success && response.data) {    const headers = [ // From UNICOURT_CONFIG.allCasesHeaders definition
      "Backend ID", "Case Number (DB Key)", "Unicourt Case Number", "Case Name (for Search)", 
      "Input Creditor Name", "Is Business", "Creditor Type", "Case Name (from Unicourt)", 
      "Unicourt Case URL", "Overall Status", "Last Submitted At", "Original Creditor Name (Doc)", 
      "Creditor Name Source Doc Title", "Creditor Address (Doc)", "Creditor Address Source Doc Title", 
      "Associated Parties (Names)", "Associated Parties (Data JSON)", "Creditor Reg State (Doc)", 
      "Creditor Reg State Source Doc Title", "Processed Documents Summary (JSON)",
      "Final Judgment Awarded to Creditor?", "Final Judgment Awarded Source Doc Title", "Final Judgment Context"
    ];
    const sheet = ensureSheetExists(UNICOURT_CONFIG.SHEET_NAMES.ALL_CASES, headers);

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }

    const cases = response.data || [];
    if (cases.length > 0) {      const rowsToAdd = cases.map(caseData => [
        caseData.id,
        caseData.case_number_for_db_id,
        caseData.unicourt_actual_case_number_on_page || "",
        caseData.case_name_for_search || "",
        caseData.input_creditor_name || "",
        caseData.is_business ? "TRUE" : "FALSE",
        caseData.creditor_type || "",
        caseData.unicourt_case_name_on_page || "",
        caseData.case_url_on_unicourt ? `=HYPERLINK("${caseData.case_url_on_unicourt}","View on Unicourt")` : "",        
        caseData.status || "N/A",
        caseData.last_submitted_at ? Utilities.formatDate(new Date(caseData.last_submitted_at), "America/New_York", "yyyy-MM-dd HH:mm:ss z") : "",
        caseData.original_creditor_name_from_doc || "",
        caseData.original_creditor_name_source_doc_title || "",
        caseData.creditor_address_from_doc || "",
        caseData.creditor_address_source_doc_title || "",
        caseData.associated_parties ? caseData.associated_parties.join("; ") : "",
        caseData.associated_parties_data ? JSON.stringify(caseData.associated_parties_data, null, 2) : "[]",
        caseData.creditor_registration_state_from_doc || "",
        caseData.creditor_registration_state_source_doc_title || "",
        caseData.processed_documents_summary ? JSON.stringify(caseData.processed_documents_summary, null, 2) : "[]",
        caseData.final_judgment_awarded_to_creditor || "",
        caseData.final_judgment_awarded_source_doc_title || "",
        caseData.final_judgment_awarded_to_creditor_context || ""
      ]);

      sheet.getRange(2, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
      SpreadsheetApp.getActiveSpreadsheet().toast(`Updated ${cases.length} cases in '${UNICOURT_CONFIG.SHEET_NAMES.ALL_CASES}' sheet.`, "Complete", 3);
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast("No cases found in the database.", "Complete", 3);
    }
    // Ensure sheet is hidden after populating
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetToHide = ss.getSheetByName(UNICOURT_CONFIG.SHEET_NAMES.ALL_CASES);
    if(sheetToHide && !sheetToHide.isSheetHidden()) sheetToHide.hideSheet();

  } else {
    ui.alert('Error', `Failed to fetch cases: ${response.error || 'Unknown error'} ${response.detail ? '('+response.detail+')' : ''}`, ui.ButtonSet.OK);
  }
}


function viewSubmissionLog_HTML() {
  const submissionsJson = PropertiesService.getScriptProperties().getProperty(UNICOURT_CONFIG.SUBMISSION_LOG_KEY);
  let submissionLogs = [];
  if (submissionsJson) {
    try {
      submissionLogs = JSON.parse(submissionsJson);
    } catch (e) {
      setDialogDisplayData("Submission Log", "<p>Error parsing stored submission log.</p>");
      showHtmlDialog("Submission Log", "Unicourt Processor DialogDisplay", 600, 400);
      return;
    }
  }

  let htmlContent;
  if (submissionLogs.length === 0) {
    htmlContent = "<p>No submission attempts logged yet.</p>";
  } else {
    htmlContent = "<ul>";
    submissionLogs.forEach(log => {
      htmlContent += `<li>
        <strong>Timestamp:</strong> ${new Date(log.timestamp).toLocaleString()}<br/>
        <strong>Case Number:</strong> ${log.caseNumber || 'N/A'}<br/>
        <strong>Type:</strong> ${log.type || 'N/A'}
      </li><hr/>`;
    });
    htmlContent += "</ul>";
  }
  setDialogDisplayData("Submission Attempt Log (Last " + submissionLogs.length + ")", htmlContent);
  showHtmlDialog("Submission Attempt Log", "Unicourt Processor DialogDisplay", 700, 500);
}