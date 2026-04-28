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
 * Sends an automated email based on a Trigger Event Handle.
 * Updated: Restored Sandbox Interceptor and 11-column index logic.
 */
function sendTriggerEmail(triggerHandle, toEmail, dataMap) {
  var sheet = getMainDb().getSheetByName("Templates");
  var data = sheet.getDataRange().getValues();
  var settings = getSystemSettings();
  
  var templateIdx = -1;
  for (var i = 1; i < data.length; i++) {
    // Column G (Index 6) is the Handle; Column J (Index 9) is Status
    if (data[i][6] === triggerHandle && data[i][9] === "Active") {
      templateIdx = i;
      break;
    }
  }

  if (templateIdx === -1) return;

  var finalToEmail = toEmail;
  var finalSubject = data[templateIdx][7]; // Column H (Subject)
  var finalHtmlBody = data[templateIdx][8]; // Column I (Body)
  var wrapperName = data[templateIdx][10]; // Column K (Wrapper)
  
  // 1. Prepare Layout from dynamic Database Wrappers
  var wrapperHtml = getWrapperContent(wrapperName);
  var fullHtml = wrapperHtml.replace("{{USER_MESSAGE_CONTENT}}", finalHtmlBody);

  // 2. Perform Placeholder Swap
  for (var key in dataMap) {
    var regex = new RegExp("\\{\\{" + key + "\\}\\}", "gi");
    var replacement = dataMap[key] || "";
    finalSubject = finalSubject.replace(regex, replacement);
    fullHtml = fullHtml.replace(regex, replacement);
  }

  // 3. RESTORED: Sandbox Environment Interceptor
  if (settings.environment === 'Sandbox' && settings.adminEmail !== '') {
    finalToEmail = settings.adminEmail;
    finalSubject = "[Sandbox Mail] " + finalSubject;
    
    var sandboxWarning = "<br><br><div style='padding: 20px; background-color: #000; color: #0f0; font-family: \"Courier New\", Courier, monospace; font-size: 14px; border: 2px solid #333; margin-top: 50px;'>";
    sandboxWarning += "=========================================<br>";
    sandboxWarning += " SYSTEM OVERRIDE: SANDBOX ENVIRONMENT    <br>";
    sandboxWarning += "=========================================<br>";
    sandboxWarning += "&gt; STATUS: INTERCEPTED<br>";
    sandboxWarning += "&gt; INTENDED RECIPIENT(S): " + toEmail + "<br>";
    sandboxWarning += "&gt; REROUTED TO ADMIN: " + settings.adminEmail + "<br>";
    sandboxWarning += "=========================================";
    sandboxWarning += "</div>";
    
    fullHtml += sandboxWarning;
  }

  // 4. Dispatch Branded Email
  MailApp.sendEmail({
    to: finalToEmail,
    subject: finalSubject,
    htmlBody: fullHtml,
    noReply: true,
    name: settings.systemName,
    inlineImages: {
      logo: getLogoBlob() 
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