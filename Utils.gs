/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Utils.gs
 * VERSION: 1.4 (Consolidated Drive + Email Engine + 11-Col Indices)
 * SYNC STATUS: Fully Synchronized with Installation.gs & Templates.gs
 */

/**
 * Server-side helper to include separate HTML files into the master template.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ========================================================================
   1. CORE INFRASTRUCTURE UTILITIES (Drive & Sheets)
   ======================================================================== */

/**
 * Standardized helper to find or create a folder.
 * Ignores trashed folders to prevent "Service error: Drive".
 */
function getOrCreateFolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  while (folders.hasNext()) {
    var f = folders.next();
    if (!f.isTrashed()) return f; 
  }
  var newFolder = parent.createFolder(name);
  try {
    newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) { console.warn("Sharing restricted by domain policy: " + e.message); }
  return newFolder;
}

function getSystemSubfolder(subfolderName) {
  var settings = getSystemSettings();
  var rootId = settings.rootFolderId;
  if (!rootId) throw new Error("System Configuration Error: Root Folder is not set.");
  
  var rootFolder = DriveApp.getFolderById(rootId);
  return getOrCreateFolder(rootFolder, subfolderName);
}

/**
 * Safely moves a file to a target folder using a retry loop.
 * Bypasses Google Drive API indexing latency.
 */
function moveFileWithRetry(fileId, targetFolder) {
  var maxRetries = 5;
  for (var i = 0; i < maxRetries; i++) {
    try {
      var file = DriveApp.getFileById(fileId);
      file.moveTo(targetFolder);
      return; 
    } catch (e) {
      if (i === maxRetries - 1) throw new Error("Drive Indexing Error: " + e.message);
      Utilities.sleep(3000); 
    }
  }
}

/**
 * Helper to initialize a sheet with bold headers and frozen top row.
 */
function initializeSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setFontWeight("bold")
         .setBackground("#F3F3F3");
    sheet.setFrozenRows(1);
  }
}

/* ========================================================================
   2. ASSET & UPLOAD UTILITIES
   ======================================================================== */

function uploadBase64File(base64, filename, folderObj) {
  try {
    var base64Data = base64.split(',')[1] || base64;
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, 'image/png', filename);
    var file = folderObj.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { url: file.getUrl(), id: file.getId() };
  } catch (e) { throw new Error("File Upload Failed: " + e.message); }
}

/**
 * Helper to retrieve the system logo as a blob for email attachments.
 */
function getLogoBlob() {
  var settings = getSystemSettings();
  if (settings.systemLogoId) {
    try { return DriveApp.getFileById(settings.systemLogoId).getBlob().setName("logo"); } 
    catch(e) { console.warn("Drive logo fetch failed: " + e.message); }
  }
  try { return UrlFetchApp.fetch(settings.appFallbackLogo).getBlob().setName("logo"); } 
  catch(e) { return Utilities.newBlob("", "image/png", "logo"); }
}

/* ========================================================================
   3. EMAIL ENGINE (11-Column Indices + Sandbox Interceptor)
   ======================================================================== */

/**
 * Sends an automated email based on a Trigger Event Handle.
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
  var finalSubject = data[templateIdx][7];  // Column H
  var finalHtmlBody = data[templateIdx][8]; // Column I
  var wrapperName = data[templateIdx][10]; // Column K
  
  var wrapperHtml = getWrapperContent(wrapperName);
  var fullHtml = wrapperHtml.replace("{{USER_MESSAGE_CONTENT}}", finalHtmlBody);

  for (var key in dataMap) {
    var regex = new RegExp("\\{\\{" + key + "\\}\\}", "gi");
    finalSubject = finalSubject.replace(regex, dataMap[key] || "");
    fullHtml = fullHtml.replace(regex, dataMap[key] || "");
  }

  // RESTORED: Sandbox Environment Interceptor
  if (settings.environment === 'Sandbox' && settings.adminEmail !== '') {
    finalToEmail = settings.adminEmail;
    finalSubject = "[Sandbox Mail] " + finalSubject;
    var sandboxWarning = "<br><br><div style='padding: 20px; background-color: #000; color: #0f0; font-family: monospace; font-size: 14px; border: 2px solid #333; margin-top: 50px;'>";
    sandboxWarning += "SYSTEM OVERRIDE: SANDBOX ENVIRONMENT INTERCEPTED<br>";
    sandboxWarning += "&gt; INTENDED RECIPIENT: " + toEmail + "<br></div>";
    fullHtml += sandboxWarning;
  }

  MailApp.sendEmail({
    to: finalToEmail, subject: finalSubject, htmlBody: fullHtml,
    noReply: true, name: settings.systemName,
    inlineImages: { logo: getLogoBlob() }
  });
}

/**
 * Sends a test email with dummy data for template verification.
 * Updated for 11-column indices.
 */
function sendTestEmailAction(rowIndex, testEmail) {
  try {
    var sheet = getMainDb().getSheetByName("Templates");
    var data = sheet.getDataRange().getValues();
    var rowData = data[rowIndex];
    
    var subject = rowData[7] || "No Subject"; // Col H
    var rawHtml = rowData[8] || "";           // Col I
    var wrapperType = rowData[10] || "Internal Hub"; // Col K
    
    var fullLayoutHtml = getWrapperContent(wrapperType).replace("{{USER_MESSAGE_CONTENT}}", rawHtml);
    var dummyData = { "companyName": "Acme Corp (Test)", "username": "jdoe", "systemName": getSystemSettings().systemName };
    
    var finalSubject = "[TEST] " + subject;
    var finalHtml = fullLayoutHtml;
    for (var key in dummyData) {
      var regex = new RegExp("\\{\\{" + key + "\\}\\}", "gi");
      finalSubject = finalSubject.replace(regex, dummyData[key]);
      finalHtml = finalHtml.replace(regex, dummyData[key]);
    }

    MailApp.sendEmail({ to: testEmail, subject: finalSubject, htmlBody: finalHtml, noReply: true, inlineImages: { logo: getLogoBlob() } });
    return "Test email successfully sent to " + testEmail;
  } catch (e) { return "Failed to send test: " + e.message; }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */