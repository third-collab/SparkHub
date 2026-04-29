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
 * Sends automated emails based on a Trigger Event Handle.
 * Upgraded to support multiple templates sharing the same trigger.
 * [IMMUTABLE ANCHOR: Sandbox Environment Interceptor preserved]
 */
function sendTriggerEmail(triggerHandle, toEmail, dataMap) {
  var data = getMainDb().getSheetByName("Templates").getDataRange().getValues();
  var settings = getSystemSettings();
  
  // 1. Gather ALL matching active templates instead of stopping at the first one
  var matchedTemplates = [];
  for (var i = 1; i < data.length; i++) {
    // 10-COLUMN INDICES: Trigger is 5 (Col F), Status is 8 (Col I)
    if (data[i][5] === triggerHandle && data[i][8] === "Active") {
      matchedTemplates.push(data[i]); 
    }
  }

  if (matchedTemplates.length === 0) return;

  // 2. Loop through every matched template and dispatch its email
  for (var t = 0; t < matchedTemplates.length; t++) {
    var templateRow = matchedTemplates[t];
    
    var finalToEmail = toEmail;
    var finalSubject = templateRow[6];  // Col G (Subject)
    var finalHtmlBody = templateRow[7]; // Col H (Body)
    var wrapperName = templateRow[9];   // Col J (Wrapper)
    
    var wrapperHtml = getWrapperContent(wrapperName);
    var fullHtml = wrapperHtml.replace("{{USER_MESSAGE_CONTENT}}", finalHtmlBody);

    for (var key in dataMap) {
      var regex = new RegExp("\\{\\{" + key + "\\}\\}", "gi");
      finalSubject = finalSubject.replace(regex, dataMap[key] || "");
      fullHtml = fullHtml.replace(regex, dataMap[key] || "");
    }

    // ========================================================================
    // [IMMUTABLE ANCHOR: SANDBOX ENVIRONMENT INTERCEPTOR]
    // Applied individually to each template in the batch
    // ========================================================================
    if (settings.environment === 'Sandbox' && settings.adminEmail !== '') {
      finalToEmail = settings.adminEmail;
      finalSubject = "[Sandbox Mail] " + finalSubject;
      var sandboxWarning = "<br><br><div style='padding: 20px; background-color: #000; color: #0f0; font-family: monospace; font-size: 14px; border: 2px solid #333; margin-top: 50px;'>";
      sandboxWarning += "SYSTEM OVERRIDE: SANDBOX ENVIRONMENT INTERCEPTED<br>";
      sandboxWarning += "&gt; INTENDED RECIPIENT: " + toEmail + "<br></div>";
      fullHtml += sandboxWarning;
    }
    // ========================================================================

    MailApp.sendEmail({
      to: finalToEmail, 
      subject: finalSubject, 
      htmlBody: fullHtml, 
      noReply: true,
      name: settings.systemName, 
      inlineImages: { logo: getLogoBlob() }
    });
  }
}

/**
 * Sends a test email with dummy data for template verification.
 * Updated for 11-column indices.
 */
function sendTestEmailAction(rowIndex, testEmail) {
  try {
    // Fetch 10 columns
    var rowData = getMainDb().getSheetByName("Templates").getRange(rowIndex, 1, 1, 10).getValues()[0];
    
    // NEW INDICES: Wrapper is 9, Body is 7, Subject is 6
    var fullHtml = getWrapperContent(rowData[9]).replace("{{USER_MESSAGE_CONTENT}}", rowData[7]);
    var finalHtml = fullHtml.replace("{{username}}", "jdoe").replace("{{systemName}}", getSystemSettings().systemName);
    
    MailApp.sendEmail({ to: testEmail, subject: "[TEST] " + rowData[6], htmlBody: finalHtml, inlineImages: { logo: getLogoBlob() } });
    return "Test email sent to " + testEmail;
  } catch (e) { return "Error: " + e.message; }
}

/**
 * Securely hashes passwords using SHA-256 for database storage.
 */
function hashPassword(password) {
  if (!password) return "";
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  var txtHash = '';
  for (var i = 0; i < rawHash.length; i++) {
    var hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length == 1) txtHash += '0';
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

/**
 * Helper to ensure the cell actually contains a valid Date object before checking it.
 * @param {any} d - The value to check.
 * @return {boolean} - Returns true if the value is a valid Date object.
 */
function isValidDate(d) {
  return d instanceof Date && !isNaN(d);
}