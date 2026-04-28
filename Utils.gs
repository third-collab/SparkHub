/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Utils.gs
 * VERSION: 1.2 (Generalized Base64 Upload Engine)
 * SYNC STATUS: Fully Synchronized with SystemsGovernance.md
 */

/**
 * Utility Module - Backend
 * Standardized under SparkHub Architecture Blueprint.
 * Handles: File uploads, subfolder management, and global template helpers.
 */

/**
 * Server-side helper to include separate HTML files into the master template.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Standardized helper to find or create a folder.
 * Centralized here to support Installation.gs, Settings.gs
 */
function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    var newFolder = parentFolder.createFolder(folderName);
    // Ensure the folder is viewable by the system for UI rendering (e.g. logo/photos)
    newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return newFolder;
  }
}

/**
 * Dynamically retrieves a subfolder within the configured Root Folder.
 * Utilizes getOrCreateFolder to ensure structural integrity.
 */
function getSystemSubfolder(subfolderName) {
  var settings = getSystemSettings();
  var rootId = settings.rootFolderId;
  
  if (!rootId) {
    throw new Error("System Configuration Error: Root Folder is not set in Settings.");
  }
  
  var rootFolder = DriveApp.getFolderById(rootId);
  return getOrCreateFolder(rootFolder, subfolderName);
}

/**
 * GENERIC UPLOAD ENGINE: Decodes a base64 string and saves it to a specified folder.
 * This is the standardized function for all system file uploads (Logos, Photos, Docs).
 * * @param {string} base64 - The data URI or raw base64 string.
 * @param {string} filename - The name to save the file as.
 * @param {GoogleAppsScript.Drive.Folder} folderObj - The target Drive Folder object.
 * @return {Object} Contains the permanent Drive URL and File ID.
 */
function uploadBase64File(base64, filename, folderObj) {
  try {
    var base64Data = base64.split(',')[1] || base64;
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, 'image/png', filename);
    
    var file = folderObj.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      url: file.getUrl(),
      id: file.getId()
    };
  } catch (e) {
    console.error("Generic upload error: " + e.message);
    throw new Error("File Upload Failed: " + e.message);
  }
}

/**
 * Standardized System Logging Engine
 * Records auditable events across the SparkHub architecture.
 * @param {string} module - Originating module (e.g., 'Users', 'Templates', 'Settings').
 * @param {string} actionType - CRUD category (CREATE, UPDATE, DELETE, SYSTEM, ALERT).
 * @param {string} actionName - Short name of the action (e.g., 'Create User').
 * @param {string} severity - INFO, WARN, ERROR, CRITICAL.
 * @param {string} targetEntity - The specific entity affected (e.g., 'johndoe', 'Welcome Email').
 * @param {string} logText - Detailed description of the event.
 */
function logSystemAction(module, actionType, actionName, severity, targetEntity, logText) {
  try {
    var ss = getLogsDb(); 
    var sheet = ss.getSheetByName("System Logs");
    if (!sheet) return; // Failsafe if not yet initialized
    
    // Try to get the active user, default to "System" if automated
    var actor = "System";
    try { actor = getLoggedInUsername(); } catch(e) {}

    // Grab the current environment
    var env = "Unknown";
    try { env = PropertiesService.getScriptProperties().getProperty('ENVIRONMENT') || "Sandbox"; } catch(e) {}

    // Build the payload
    var logData = [[
      new Date(),
      module,
      actionType,
      actionName,
      severity,
      actor,
      targetEntity,
      logText,
      env
    ]];

    // INSERT AT THE TOP: Create a new row right under the header (Row 1) and insert data
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, 9).setValues(logData);
    
  } catch (e) {
    console.error("Logging Engine Failure: " + e.message);
  }
}


/**
 * Helper to retrieve the system logo as a blob for email attachments.
 * Priority: 1. Drive Logo, 2. Fallback URL (Centralized logic).
 * @return {Blob} The logo image blob.
 */
function getLogoBlob() {
  var settings = getSystemSettings();
  
  // 1. Try Custom Uploaded Logo from Drive
  if (settings.systemLogoId) {
    try {
      return DriveApp.getFileById(settings.systemLogoId).getBlob().setName("logo");
    } catch(e) {
      console.warn("Drive logo fetch failed, proceeding to fallback: " + e.message);
    }
  }
  
  // 2. Use Centralized Fallback URL 
  // (The Imgur URL failsafe is managed once in Settings.gs:getSystemSettings)
  try {
    return UrlFetchApp.fetch(settings.fallbackLogoUrl).getBlob().setName("logo");
  } catch(e) {
    console.error("Critical: All logo blob fetches failed: " + e.message);
    // Return an empty transparent pixel or empty blob to prevent MailApp crash
    return Utilities.newBlob("", "image/png", "logo");
  }
}

/**
 * Sends an automated email based on a Trigger Event mapped in the Templates database.
 * Merges user data, applies Sandbox overrides, and attaches branding.
 * @param {string} triggerName - The name of the trigger event.
 * @param {string} toEmail - The intended recipient's email address.
 * @param {object} dataMap - Key-value pairs for {{placeholders}}.
 */
function sendTriggerEmail(triggerName, toEmail, dataMap) {
  var sheet = getMainDb().getSheetByName("Templates");
  var data = sheet.getDataRange().getValues();
  
  // Fetch system-wide branding and environment settings
  var settings = getSystemSettings();
  var sysName = settings.systemName;
  
  var subject = "";
  var templateFound = false;
  var templateIdx = -1;

  for (var i = 1; i < data.length; i++) {
    var rowTrigger = data[i][5];
    var rowStatus = data[i][8];
    
    if (rowTrigger === triggerName && rowStatus === "Active") {
      subject = data[i][6];
      templateIdx = i;
      templateFound = true;
      break;
    }
  }

  if (!templateFound) {
    console.log("No active template found for trigger: " + triggerName);
    return;
  }

  var logoBlob = getLogoBlob();
  var rawHtml = data[templateIdx][7];
  var wrapperType = data[templateIdx][9]; 

  // 1. Prepare Layout
  var wrapperHtml = getWrapperContent(wrapperType);
  var fullLayoutHtml = wrapperHtml.replace("{{USER_MESSAGE_CONTENT}}", rawHtml);

  // 2. Perform Placeholder Swap
  var finalSubject = subject;
  var finalHtml = fullLayoutHtml;
  
  for (var key in dataMap) {
    var regex = new RegExp("\\{\\{" + key + "\\}\\}", "gi");
    var replacement = dataMap[key] || "";
    finalSubject = finalSubject.replace(regex, replacement);
    finalHtml = finalHtml.replace(regex, replacement);
  }

  // 3. Sandbox Environment Interceptor
  var finalToEmail = toEmail;

  if (settings.environment === 'Sandbox' && settings.adminEmail !== '') {
    finalToEmail = settings.adminEmail;
    finalSubject = "[Sandbox Mail] " + finalSubject;
    
    // Aggressive "program code" block for the override notification
    var sandboxWarning = "<br><br><div style='padding: 20px; background-color: #000; color: #0f0; font-family: \"Courier New\", Courier, monospace; font-size: 14px; border: 2px solid #333; margin-top: 50px;'>";
    sandboxWarning += "=========================================<br>";
    sandboxWarning += " SYSTEM OVERRIDE: SANDBOX ENVIRONMENT    <br>";
    sandboxWarning += "=========================================<br>";
    sandboxWarning += "&gt; STATUS: INTERCEPTED<br>";
    sandboxWarning += "&gt; INTENDED RECIPIENT(S): " + toEmail + "<br>";
    sandboxWarning += "&gt; REROUTED TO ADMIN: " + settings.adminEmail + "<br>";
    sandboxWarning += "=========================================";
    sandboxWarning += "</div>";
    
    finalHtml += sandboxWarning;
  }

  // 4. Dispatch Email
  MailApp.sendEmail({
    to: finalToEmail,
    subject: finalSubject,
    htmlBody: finalHtml,
    noReply: true,
    name: sysName,
    inlineImages: {
      logo: logoBlob 
    }
  });
}

/**
 * Sends a test email with dummy data for template verification.
 * @param {number} rowIndex - Template row index.
 * @param {string} testEmail - Recipient for the test.
 * @return {string} Success or failure message.
 */
function sendTestEmailAction(rowIndex, testEmail) {
  var sheet = getMainDb().getSheetByName("Templates");
  var data = sheet.getDataRange().getValues();
  
  if (rowIndex < 1 || rowIndex >= data.length) return "Error: Template not found.";
  
  var settings = getSystemSettings();
  var rowData = data[rowIndex];
  var subject = rowData[6] || "No Subject";
  var rawHtml = rowData[7] || "";
  var wrapperType = rowData[9] || "Internal";
  
  var wrapperHtml = getWrapperContent(wrapperType);
  var fullLayoutHtml = wrapperHtml.replace("{{USER_MESSAGE_CONTENT}}", rawHtml);

  // Comprehensive Dummy Data
  var dummyData = {
    "companyName": "Acme Corp (Test)",
    "brandName": "Acme Brand",
    "address": "123 Test Ave, Suite 100",
    "website": "www.megarhino.com",
    "priFirstName": "John",
    "priLastName": "Doe",
    "priEmail": "john@example.com",
    "monthlyContractValue": "$2,500",
    "contractStartDate": "2026-05-01",
    "services": "SEO & Content Marketing",
    "notes": "Sample test note.",
    "username": "jdoe",
    "firstName": "John",
    "lastName": "Doe",
    "role": "Account Manager"
  };

  var finalSubject = "[TEST] " + subject;
  var finalHtml = fullLayoutHtml;

  for (var key in dummyData) {
    var regex = new RegExp("\\{\\{" + key + "\\}\\}", "gi");
    finalSubject = finalSubject.replace(regex, dummyData[key]);
    finalHtml = finalHtml.replace(regex, dummyData[key]);
  }

  try {
    MailApp.sendEmail({
      to: testEmail,
      subject: finalSubject,
      htmlBody: finalHtml,
      noReply: true,
      name: settings.systemName,
      inlineImages: {
        logo: getLogoBlob()
      }
    });
    return "Test email successfully sent to " + testEmail;
  } catch (e) {
    return "Failed to send test email: " + e.message;
  }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */