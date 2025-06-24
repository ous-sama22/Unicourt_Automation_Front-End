// BackendService.gs

/**
 * Makes a generic call to the backend API.
 * Logs errors for later viewing.
 */
function callBackendApi(endpoint, method, payload, isRetry = false) { 
  const properties = PropertiesService.getScriptProperties();
  const backendUrl = properties.getProperty(UNICOURT_CONFIG.SETTINGS_KEYS.BACKEND_URL);
  const apiKey = properties.getProperty(UNICOURT_CONFIG.SETTINGS_KEYS.BACKEND_API_KEY);

  if (!backendUrl || !apiKey) {
    const errorMsg = "Backend URL or API Key not configured in Script Properties.";
    Logger.log(errorMsg);
    logApiError(backendUrl ? `${backendUrl}${endpoint}` : "Unknown Endpoint", method, "N/A", errorMsg);
    return { success: false, error: errorMsg };
  }

  const options = {
    method: method.toLowerCase(), 
    contentType: 'application/json',
    headers: {
      'X-API-KEY': apiKey
    },
    muteHttpExceptions: true
  };

  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  let response;
  let responseCode;
  let responseBody;
  const fullUrl = `${backendUrl}${endpoint}`;

  try {
    Logger.log(`Calling backend: ${method.toUpperCase()} ${fullUrl} with payload: ${payload ? JSON.stringify(payload).substring(0, 200) + (JSON.stringify(payload).length > 200 ? '...' : '') : 'None'}`);
    response = UrlFetchApp.fetch(fullUrl, options);
    responseCode = response.getResponseCode();
    responseBody = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      return { success: true, data: JSON.parse(responseBody), code: responseCode };
    } else {
      let errorDetail = responseBody;
      try {
        const parsedError = JSON.parse(responseBody);
        errorDetail = parsedError.detail || responseBody;
      } catch (e) { /* Keep responseBody as errorDetail */ }
      
      Logger.log(`Backend Error: ${responseCode} from ${method.toUpperCase()} ${fullUrl}. Detail: ${errorDetail}`);
      logApiError(fullUrl, method, responseCode, errorDetail);
      return { success: false, error: `HTTP Error ${responseCode}`, detail: errorDetail, code: responseCode };
    }
  } catch (e) {
    Logger.log(`Network or Apps Script error calling backend: ${e.toString()}\n${e.stack}`);
    logApiError(fullUrl, method, "Network/Script Error", e.toString());
    return { success: false, error: `Network/Script Error: ${e.toString()}` };
  }
}

/**
 * Logs an API error to a rolling list in ScriptProperties.
 */
function logApiError(endpoint, method, responseCode, responseBody) {
  const scriptProps = PropertiesService.getScriptProperties();
  const errorsJson = scriptProps.getProperty(UNICOURT_CONFIG.ERROR_LOG_KEY);
  let errors = [];
  if (errorsJson) {
    try {
      errors = JSON.parse(errorsJson);
    } catch (e) {
      Logger.log("Error parsing existing error log: " + e.toString());
      errors = []; 
    }
  }

  const newError = {
    timestamp: new Date().toISOString(),
    method: method.toUpperCase(),
    endpoint: endpoint,
    responseCode: String(responseCode),
    responseBody: responseBody // Storing the raw string or parsed detail
  };

  errors.unshift(newError); 

  if (errors.length > UNICOURT_CONFIG.MAX_ERROR_LOG_ENTRIES) {
    errors = errors.slice(0, UNICOURT_CONFIG.MAX_ERROR_LOG_ENTRIES); 
  }

  try {
    scriptProps.setProperty(UNICOURT_CONFIG.ERROR_LOG_KEY, JSON.stringify(errors));
  } catch (e) {
     Logger.log("Error saving updated error log: " + e.toString());
  }
}

// --- Specific Backend Call Wrappers ---

function callBackendSubmitCases(payload) {
  return callBackendApi('/cases/submit', 'post', payload);
}

function callBackendBatchDetails(payload) {
  return callBackendApi('/cases/batch-details', 'post', payload);
}


function callBackendConfigUpdate(payload) {
  return callBackendApi('/service/config', 'put', payload);
}

function callBackendRestart() {
  return callBackendApi('/service/request-restart', 'post', null);
}

function fetchCurrentBackendConfig() {
  return callBackendApi('/service/config', 'get', null);
}

function callBackendGetAllCases() {
  return callBackendApi('/cases', 'get', null);
}

/**
 * Logs a case submission attempt to a rolling list in ScriptProperties.
 */
function logSubmissionAttempt(caseNumberForDbId, submissionType) {
  if (!caseNumberForDbId || !submissionType) {
    Logger.log("logSubmissionAttempt: Missing caseNumberForDbId or submissionType. Skipping log.");
    return;
  }
  const scriptProps = PropertiesService.getScriptProperties();
  const submissionsLogJson = scriptProps.getProperty(UNICOURT_CONFIG.SUBMISSION_LOG_KEY);
  let submissionLogs = [];
  if (submissionsLogJson) {
    try {
      submissionLogs = JSON.parse(submissionsLogJson);
    } catch (e) {
      Logger.log("Error parsing existing submission log: " + e.toString());
      submissionLogs = [];
    }
  }

  const newLogEntry = {
    timestamp: new Date().toISOString(),
    caseNumber: caseNumberForDbId,
    type: submissionType // 'Manual' or 'Auto'
  };

  submissionLogs.unshift(newLogEntry);

  if (submissionLogs.length > UNICOURT_CONFIG.MAX_SUBMISSION_LOG_ENTRIES) {
    submissionLogs = submissionLogs.slice(0, UNICOURT_CONFIG.MAX_SUBMISSION_LOG_ENTRIES);
  }

  try {
    scriptProps.setProperty(UNICOURT_CONFIG.SUBMISSION_LOG_KEY, JSON.stringify(submissionLogs));
    Logger.log(`Logged submission attempt for Case: ${caseNumberForDbId}, Type: ${submissionType}`);
  } catch (e) {
    Logger.log("Error saving updated submission log: " + e.toString());
  }
}