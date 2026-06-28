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
    if (tData[i][6] === triggerHandle && tData[i][10] === "Active" && isWrapperActive(tData[i][9])) {
      
      // Dynamic Trigger Inclusion/Exclusion Evaluation Engine (Columns 16 & 17)
      var incRule = tData[i][15] ? String(tData[i][15]).trim().toLowerCase() : "";
      var excRule = tData[i][16] ? String(tData[i][16]).trim().toLowerCase() : "";
      var contextValue = "";
      
      if (triggerHandle.indexOf("Users:") === 0) {
        contextValue = dataMap.target_role || dataMap.roleName || "";
      } else if (triggerHandle.indexOf("Clients:") === 0) {
        contextValue = dataMap.status || "";
      }
      contextValue = String(contextValue).trim().toLowerCase();
      
      if (incRule !== "") {
        var allowedArr = incRule.split(',').map(function(s){ return s.trim(); });
        if (allowedArr.indexOf(contextValue) === -1) continue; // Skip: Fails inclusion check
      }
      if (excRule !== "") {
        var blockedArr = excRule.split(',').map(function(s){ return s.trim(); });
        if (blockedArr.indexOf(contextValue) > -1) continue; // Skip: Trapped by exclusion criteria
      }

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

    // Dynamic Centralized Ecosystem Distribution Rule Processor Engine 
    function resolveEcosystemRecipientRule(ruleStr, defaultEmail) {
      var emails = [];
      var profiles = [];
      if (!ruleStr || ruleStr.trim() === "" || ruleStr === "NONE") return { emails: emails, profiles: profiles };
      
      var conditionPart = ruleStr;
      var manualPart = "";
      var pipeIdx = ruleStr.indexOf('|');
      if (pipeIdx > -1) {
        conditionPart = ruleStr.substring(0, pipeIdx).trim();
        manualPart = ruleStr.substring(pipeIdx + 1).trim();
      }
      
      if (conditionPart === "TRIGGER_DEFAULT") {
        if (defaultEmail) {
          emails.push(defaultEmail);
          var foundUser = null;
          try {
            if (typeof getUsersList === 'function') {
              var uList = getUsersList();
              for (var u = 0; u < uList.length; u++) {
                if (String(uList[u].email).toLowerCase() === defaultEmail.toLowerCase() || String(uList[u].systemEmail).toLowerCase() === defaultEmail.toLowerCase()) {
                  foundUser = uList[u];
                  break;
                }
              }
            }
          } catch(e) {}
          if (foundUser) {
            profiles.push(foundUser);
          } else {
            profiles.push({ systemEmail: defaultEmail, firstName: "Recipient", lastName: "", username: defaultEmail.split('@')[0], role: "User" });
          }
        }
      } else if (conditionPart === "ALL_ACTIVE_USERS") {
        try {
          if (typeof getUsersList === 'function') {
            getUsersList().forEach(function(u) {
              if (u.status === 'Active' && u.systemEmail) {
                emails.push(u.systemEmail); profiles.push(u);
              }
            });
          }
        } catch(e) { console.warn(e.message); }
      } else if (conditionPart.indexOf("CUSTOM_LOOKUP:") === 0) {
        try {
          if (typeof getUsersList === 'function') {
            var targetIds = conditionPart.replace("CUSTOM_LOOKUP:", "").split(",").map(function(id) { return id.trim(); });
            getUsersList().forEach(function(u) {
              if (targetIds.indexOf(u.userId) > -1 && u.status === 'Active' && u.systemEmail) {
                emails.push(u.systemEmail); profiles.push(u);
              }
            });
          }
        } catch(e) { console.warn(e.message); }
      } else if (conditionPart !== "NONE" && conditionPart !== "MANUAL_ONLY") {
        manualPart = manualPart ? (conditionPart + "," + manualPart) : conditionPart;
      }
      
      if (manualPart && manualPart.trim() !== "") {
        manualPart.split(',').forEach(function(em) {
          var cleanEm = em.trim();
          if (cleanEm !== "" && emails.indexOf(cleanEm) === -1) {
            emails.push(cleanEm);
            profiles.push({ systemEmail: cleanEm, firstName: "Recipient", lastName: "", username: cleanEm.split('@')[0] });
          }
        });
      }
      return { emails: emails, profiles: profiles };
    }

    // 1. Evaluate routing conditions and resolve user accounts directly to active System Emails
    var toResult = resolveEcosystemRecipientRule(toRule, toEmail);
    var resolvedEmails = toResult.emails;
    var resolvedProfiles = toResult.profiles;

    if (resolvedEmails.length === 0 && toEmail) {
      resolvedEmails.push(toEmail);
      resolvedProfiles.push({ systemEmail: toEmail, firstName: "Recipient", lastName: "", username: toEmail.split('@')[0] });
    }

    // Dynamic Recipient Inclusion/Exclusion Evaluation Engine (Columns 18 & 19)
    var recIncRule = templateRow[17] ? String(templateRow[17]).trim().toLowerCase() : "";
    var recExcRule = templateRow[18] ? String(templateRow[18]).trim().toLowerCase() : "";
    
    if (recIncRule !== "" || recExcRule !== "") {
      var filteredEmails = []; var filteredProfiles = [];
      for (var pIdx = 0; pIdx < resolvedProfiles.length; pIdx++) {
        var prof = resolvedProfiles[pIdx];
        var trackingRole = String(prof.role || "User").trim().toLowerCase();
        var targetEmailMatch = String(dataMap.priEmail || dataMap.userEmail || "").trim().toLowerCase();
        var currentProfileEmail = String(prof.systemEmail || "").trim().toLowerCase();
        
        var isTargetUser = (targetEmailMatch !== "" && currentProfileEmail === targetEmailMatch);
        var evalMatchToken = isTargetUser ? "target_user" : trackingRole;
        
        if (recIncRule !== "") {
          var allowedRec = recIncRule.split(',').map(function(s){ return s.trim(); });
          if (allowedRec.indexOf(evalMatchToken) === -1) continue;
        }
        if (recExcRule !== "") {
          var blockedRec = recExcRule.split(',').map(function(s){ return s.trim(); });
          if (blockedRec.indexOf(evalMatchToken) > -1) continue;
        }
        filteredEmails.push(resolvedEmails[pIdx]);
        filteredProfiles.push(prof);
      }
      resolvedEmails = filteredEmails;
      resolvedProfiles = filteredProfiles;
    }

    // Extract CC and BCC distributions utilizing the matching dynamic framework tokens rule
    var ccResult = resolveEcosystemRecipientRule(finalCcEmail, "");
    var bccResult = resolveEcosystemRecipientRule(finalBccEmail, "");
    
    // Apply recipient exclusion drop-rules to CC and BCC rows if target_user is explicitly filtered
    if (recExcRule !== "" && recExcRule.split(',').map(function(s){ return s.trim(); }).indexOf("target_user") > -1) {
      var targetEmailMatch = String(dataMap.priEmail || dataMap.userEmail || "").trim().toLowerCase();
      if (targetEmailMatch !== "") {
        ccResult.emails = ccResult.emails.filter(function(e) { return e.trim().toLowerCase() !== targetEmailMatch; });
        bccResult.emails = bccResult.emails.filter(function(e) { return e.trim().toLowerCase() !== targetEmailMatch; });
      }
    }
    
    var baseCcEmails = ccResult.emails.join(', ');
    var baseBccEmails = bccResult.emails.join(', ');

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
        localContextMap.userRole = profile.role || "User";

        for (var token in localContextMap) {
          var regex = new RegExp("\\{\\{" + token + "\\}\\}", "gi");
          currentSubject = currentSubject.replace(regex, localContextMap[token] || "");
          currentHtml = currentHtml.replace(regex, localContextMap[token] || "");
        }

        var currentCc = baseCcEmails;
        var currentBcc = baseBccEmails;
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

        enqueueEmailRow(templateId, currentTo, baseCcEmails, baseBccEmails, localContextMap, "Individual");
      });
    } else {
      var currentTo = resolvedEmails.join(', ');
      enqueueEmailRow(templateId, currentTo, baseCcEmails, baseBccEmails, dataMap, "Collective");
    }
  }
}

function enqueueEmailRow(templateId, toEmail, ccEmail, bccEmail, dataMap, dispatchMode) {
  try {
    var sheet = getQueueDb().getSheetByName("Email Queue");
    if (!sheet) return;
    var queueId = "Q-" + Math.floor(100000 + Math.random() * 900000);
    sheet.appendRow([new Date(), queueId, templateId, toEmail, ccEmail, bccEmail, JSON.stringify(dataMap), dispatchMode, "Pending", "", ""]);
    SpreadsheetApp.flush();
  } catch(e) {
    console.error("Failed to enqueue email row: " + e.message);
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