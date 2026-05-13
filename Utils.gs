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
  var mainDb = getMainDb();
  var tData = mainDb.getSheetByName("Templates").getDataRange().getValues();
  var wData = mainDb.getSheetByName("Wrappers").getDataRange().getValues();
  var settings = getSystemSettings();
  var matchedTemplates = [];
  
  // Helper to check wrapper status directly from memory
  function isWrapperActive(wName) {
    for (var w=1; w<wData.length; w++) {
      if (wData[w][2] === wName && wData[w][5] === "Active") return true;
    }
    return false;
  }

  for (var i = 1; i < tData.length; i++) {
    // Trigger is Col G (Index 6), Status is Col K (Index 10), Wrapper is Col J (Index 9)
    // CRITICAL: Block email dispatch if the assigned wrapper is inactive
    if (tData[i][6] === triggerHandle && tData[i][10] === "Active" && isWrapperActive(tData[i][9])) {
      matchedTemplates.push(tData[i]);
    }
  }

  if (matchedTemplates.length === 0) return;

  for (var t = 0; t < matchedTemplates.length; t++) {
    var templateRow = matchedTemplates[t];
    var finalToEmail = toEmail;
    var finalSubject = templateRow[7];
    var finalHtmlBody = templateRow[8];
    var wrapperName = templateRow[9];
    
    // Fetch wrapper HTML
    var wrapperHtml = "{{USER_MESSAGE_CONTENT}}";
    for (var w=1; w<wData.length; w++) {
      if (wData[w][2] === wrapperName && wData[w][5] === "Active") wrapperHtml = wData[w][4];
    }

    finalHtmlBody = applyGlobalSignature(finalHtmlBody);
    
    var fullHtml = wrapperHtml.replace("{{USER_MESSAGE_CONTENT}}", finalHtmlBody);
    for (var key in dataMap) {
      var regex = new RegExp("\\{\\{" + key + "\\}\\}", "gi");
      finalSubject = finalSubject.replace(regex, dataMap[key] || "");
      fullHtml = fullHtml.replace(regex, dataMap[key] || "");
    }

    if (settings.environment === 'Sandbox' && settings.adminEmail !== '') {
      finalToEmail = settings.adminEmail;
      finalSubject = "[Sandbox Mail] " + finalSubject;
      var sandboxWarning = "<br><br><div style='padding: 20px; background-color: #000; color: #0f0; font-family: monospace; font-size: 14px; border: 2px solid #333; margin-top: 50px;'>";
      sandboxWarning += "SYSTEM OVERRIDE: SANDBOX ENVIRONMENT INTERCEPTED<br>";
      sandboxWarning += "&gt; INTENDED RECIPIENT: " + toEmail + "<br></div>";
      fullHtml += sandboxWarning;
    }

    // 2. Wrap all links for tracking right before sending
    fullHtml = applyLinkTracking(fullHtml, finalToEmail);

    MailApp.sendEmail({
      to: finalToEmail, subject: finalSubject, htmlBody: fullHtml, 
      noReply: true, name: settings.systemName, inlineImages: { logo: getLogoBlob() }
    });
  }
}

function sendTestEmail(rowIndex, testEmail) {
  try {
    var rowData = getMainDb().getSheetByName("Templates").getRange(rowIndex, 1, 1, 11).getValues()[0];
    var fullHtml = getWrapperContent(rowData[9]).replace("{{USER_MESSAGE_CONTENT}}", rowData[8]);
    var finalHtml = fullHtml.replace("{{username}}", "jdoe").replace("{{systemName}}", getSystemSettings().systemName);
    
    MailApp.sendEmail({ to: testEmail, subject: "[TEST] " + rowData[7], htmlBody: finalHtml, inlineImages: { logo: getLogoBlob() } });
    return "Test email sent to " + testEmail;
  } catch (e) { return "Error: " + e.message; }
}

/* ========================================================================
   4. SECURITY & DATA FORMATTING
   ======================================================================== */

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

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */