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
    
    // Extracts advanced custom recipient strings and distribution rules from columns 12, 13, 14, and 15
    var toRule = templateRow[11] ? String(templateRow[11]).trim() : "";
    var finalCcEmail = templateRow[12] ? String(templateRow[12]).trim() : "";
    var finalBccEmail = templateRow[13] ? String(templateRow[13]).trim() : "";
    var dispatchMode = templateRow[14] ? String(templateRow[14]).trim() : "Individual";
    
    var baseSubject = templateRow[7];
    var baseHtmlBody = templateRow[8];
    var wrapperName = templateRow[9];

    var wrapperHtml = "{{USER_MESSAGE_CONTENT}}";
    for (var w=1; w<wData.length; w++) {
      if (wData[w][2] === wrapperName && wData[w][5] === "Active") wrapperHtml = wData[w][4];
    }

    baseHtmlBody = applyGlobalSignature(baseHtmlBody);
    var baseFullHtml = wrapperHtml.replace("{{USER_MESSAGE_CONTENT}}", baseHtmlBody);

    // 1. Evaluate routing conditions and resolve user accounts directly to active System Emails
    var resolvedEmails = [];
    var resolvedProfiles = [];

    if (!toRule || toRule === "" || toRule === "TRIGGER_DEFAULT") {
      resolvedEmails.push(toEmail);
      resolvedProfiles.push({ systemEmail: toEmail, firstName: "User", lastName: "", username: toEmail.split('@')[0] });
    } else {
      try {
        if (typeof getUsersList === 'function') {
          var directory = getUsersList();
          if (toRule === "ALL_ACTIVE_USERS") {
            directory.forEach(function(u) {
              if (u.status === 'Active' && u.systemEmail) {
                resolvedEmails.push(u.systemEmail); resolvedProfiles.push(u);
              }
            });
          } else if (toRule.indexOf("CUSTOM_LOOKUP:") === 0) {
            var targetIds = toRule.replace("CUSTOM_LOOKUP:", "").split(",").map(id => id.trim());
            directory.forEach(function(u) {
              if (targetIds.indexOf(u.userId) > -1 && u.status === 'Active' && u.systemEmail) {
                resolvedEmails.push(u.systemEmail); resolvedProfiles.push(u);
              }
            });
          }
        }
      } catch(resolveErr) { console.warn("Recipient extractor defaulted to baseline: " + resolveErr.message); }
    }

    if (resolvedEmails.length === 0) { resolvedEmails.push(toEmail); }

    // 2. Branch processing execution logic to enforce the selected Dispatch Mode type
    if (dispatchMode === "Individual") {
      // Loop execution spins up a separate high-fidelity personalized email package transaction per user
      resolvedProfiles.forEach(function(profile) {
        var currentTo = profile.systemEmail;
        var currentSubject = baseSubject;
        var currentHtml = baseFullHtml;
        
        var localContextMap = {};
        for (var key in dataMap) { localContextMap[key] = dataMap[key]; }
        localContextMap.username = profile.username || "";
        localContextMap.firstName = profile.firstName || "User";
        localContextMap.lastName = profile.lastName || "";
        localContextMap.userFirst = profile.firstName || "User";
        localContextMap.userEmail = profile.systemEmail || "";

        for (var token in localContextMap) {
          var regex = new RegExp("\\{\\{" + token + "\\}\\}", "gi");
          currentSubject = currentSubject.replace(regex, localContextMap[token] || "");
          currentHtml = currentHtml.replace(regex, localContextMap[token] || "");
        }

        var currentCc = finalCcEmail; var currentBcc = finalBccEmail;
        for (var token in localContextMap) {
          var regex = new RegExp("\\{\\{" + token + "\\}\\}", "gi");
          currentCc = currentCc.replace(regex, localContextMap[token] || "");
          currentBcc = currentBcc.replace(regex, localContextMap[token] || "");
        }

        if (settings.environment === 'Sandbox' && settings.adminEmail !== '') {
          currentTo = settings.adminEmail; currentCc = ""; currentBcc = "";
          currentSubject = "[Sandbox Mail] " + currentSubject;
          currentHtml += "<br><br><div style='padding: 20px; background-color: #000; color: #0f0; font-family: monospace; font-size: 14px; border: 2px solid #333;'>SYSTEM OVERRIDE: SANDBOX INTERCEPTED<br>&gt; INTENDED RECIPIENT: " + profile.systemEmail + "<br></div>";
        }

        currentHtml = applyLinkTracking(currentHtml, currentTo);
        var mailOptions = { to: currentTo, subject: currentSubject, htmlBody: currentHtml, noReply: true, name: settings.systemName, inlineImages: { logo: getLogoBlob() } };
        if (currentCc) mailOptions.cc = currentCc;
        if (currentBcc) mailOptions.bcc = currentBcc;
        MailApp.sendEmail(mailOptions);
      });
    } else {
      // Collective group dispatches execute a single shared email payload asset delivery
      var currentTo = resolvedEmails.join(', ');
      var currentSubject = baseSubject;
      var currentHtml = baseFullHtml;
      var currentCc = finalCcEmail;
      var currentBcc = finalBccEmail;

      for (var token in dataMap) {
        var regex = new RegExp("\\{\\{" + token + "\\}\\}", "gi");
        currentTo = currentTo.replace(regex, dataMap[token] || "");
        currentCc = currentCc.replace(regex, dataMap[token] || "");
        currentBcc = currentBcc.replace(regex, dataMap[token] || "");
        currentSubject = currentSubject.replace(regex, dataMap[token] || "");
        currentHtml = currentHtml.replace(regex, dataMap[token] || "");
      }

      if (settings.environment === 'Sandbox' && settings.adminEmail !== '') {
        currentTo = settings.adminEmail; currentCc = ""; currentBcc = "";
        currentSubject = "[Sandbox Mail] " + currentSubject;
        currentHtml += "<br><br><div style='padding: 20px; background-color: #000; color: #0f0; font-family: monospace; font-size: 14px; border: 2px solid #333;'>SYSTEM OVERRIDE: SANDBOX INTERCEPTED<br>&gt; INTENDED RECIPIENTS: " + resolvedEmails.join(', ') + "<br></div>";
      }

      currentHtml = applyLinkTracking(currentHtml, currentTo.split(',')[0].trim());
      var mailOptions = { to: currentTo, subject: currentSubject, htmlBody: currentHtml, noReply: true, name: settings.systemName, inlineImages: { logo: getLogoBlob() } };
      if (currentCc) mailOptions.cc = currentCc;
      if (finalBccEmail || currentBcc) mailOptions.bcc = finalBccEmail || currentBcc;
      MailApp.sendEmail(mailOptions);
    }
  }
}

function sendTestEmail(rowIndex, testEmail) {
  try {
    // Widened manual testing scan width bounds to 15 columns matching live engine parameters
    var rowData = getMainDb().getSheetByName("Templates").getRange(parseInt(rowIndex), 1, 1, 15).getValues()[0];
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