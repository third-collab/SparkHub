/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Templates.gs
 * VERSION: 1.2
 * SYNC STATUS: Fully Synchronized with TemplatesData.html & EventBroker.gs
 */

// ========================================================================
// 1. REGISTRY EXPORTS
// ========================================================================
function Templates_getTriggers() {
  return [
    "Templates:CREATE", "Templates:UPDATE", 
    "Templates:Wrappers:CREATE", "Templates:Wrappers:UPDATE"
  ];
}

function Templates_getPlaceholders() {
  return ["templateName", "wrapperName", "triggerEvent"];
}

function Templates_getPermissions() {
  return ["View Templates", "Create Templates", "Manage Templates", "View Wrappers", "Create Wrappers", "Manage Wrappers", "Manage Settings"];
}

function Templates_getLookups() {
  var vectors = { templates: [], wrappers: [] };
  try {
    var tplSheet = getMainDb().getSheetByName("Templates");
    var tData = tplSheet.getDataRange().getValues();
    for (var j = 1; j < tData.length; j++) {
      if (tData[j][1] && tData[j][10] !== 'Inactive') {
        vectors.templates.push({
          id: String(tData[j][1]), 
          name: String(tData[j][2]) 
        });
      }
    }
  } catch(e) { console.warn("Templates lookup broadcast failed: " + e.message); }
  
  try {
    var wrapSheet = getMainDb().getSheetByName("Wrappers");
    if (wrapSheet) {
      var wData = wrapSheet.getDataRange().getValues();
      for (var w = 1; w < wData.length; w++) {
        if (wData[w][1] && wData[w][5] === 'Active') {
          vectors.wrappers.push({
            id: String(wData[w][2]), 
            name: String(wData[w][2]) 
          });
        }
      }
    }
  } catch(e) { console.warn("Wrappers lookup broadcast failed: " + e.message); }
  return vectors;
}

function getDynamicTriggerRegistry() {
  var triggers = [];
  var globalScope = (1, eval)("this");
  var activeModules = getActiveModules();

  // 3. Only execute triggers for active modules
  activeModules.forEach(function(modName) {
    var funcName = modName + "_getTriggers";
    if (typeof globalScope[funcName] === 'function') {
      triggers = triggers.concat(globalScope[funcName]());
    } else if (modName !== "System" && modName !== "Settings" && modName !== "Logs") {
      // Fallback defaults if the custom registry function is missing
      triggers.push(modName + ":CREATE");
      triggers.push(modName + ":UPDATE");
    }
  });
  // 5. IMMUTABLE ANCHOR: Hide hardcoded core templates from the UI dropdown
  var hiddenTriggers = ["Logs:EXPORT", "Logs:PURGE", "Settings:UPDATE", "System:INSTALL", "System:MODULE_INSTALLED"];
  triggers = triggers.filter(function(t) { return hiddenTriggers.indexOf(t) === -1; });
  
  return [...new Set(triggers)].sort();
}

function getPlaceholderMetadataRegistry() {
  var registry = {};
  var globalScope = (1, eval)("this");
  var activeModules = getActiveModules();
  
  activeModules.forEach(function(modName) {
    var funcName = modName + "_getPlaceholderMetadata";
    if (typeof globalScope[funcName] === 'function') {
      var meta = globalScope[funcName]();
      for (var token in meta) {
        registry[token] = meta[token];
      }
    }
  });
  return registry;
}

function getPlaceholderSuggestions() {
  var placeholders = ["details", "systemName"];
  var globalScope = (1, eval)("this");
  // Pulls the unified collection natively from our central config module service
  var activeModules = getActiveModules();
  
  // 3. Only fetch placeholders for active modules
  activeModules.forEach(function(modName) {
    var funcName = modName + "_getPlaceholders";
    if (typeof globalScope[funcName] === 'function') {
      placeholders = placeholders.concat(globalScope[funcName]());
    }
  });
  
  return [...new Set(placeholders)].sort();
}

// ========================================================================
// 2. CORE PROCESSORS
// ========================================================================
// ========================================================================
// 2. CORE PROCESSORS
// ========================================================================
var Templates = {
  handleEventEmail: function(payload) {
    if (!payload.recipientEmail) return;
    var dataMap = payload.extraData || {};
    dataMap.username = payload.entity;
    dataMap.details = payload.details;
    dataMap.systemName = getSystemSettings().systemName || "SparkHub";
    
    var hardcodedTriggers = ["Logs:EXPORT", "Logs:PURGE", "Settings:UPDATE", "System:INSTALL", "System:MODULE_INSTALLED"];
    
    if (hardcodedTriggers.indexOf(payload.handle) > -1) {
      sendHardcodedEmail(payload.handle, payload.recipientEmail, dataMap);
    } else {
      sendTriggerEmail(payload.handle, payload.recipientEmail, dataMap);
    }
  }
};

/**
 * Generates and dispatches immutable system emails bypassing the database registry.
 * Forces the use of the "Internal Communication" wrapper.
 */
function getQueueList() {
  try {
    var data = getQueueDb().getSheetByName("Email Queue").getDataRange().getDisplayValues();
    data.shift();
    return data.map(function(row, index) {
      return {
        rowIndex: index + 2,
        timestamp: row[0],
        id: row[1],
        templateId: row[2],
        recipient: row[3],
        status: row[8],
        details: row[9] || (row[8] === 'Sent' ? 'Dispatched at ' + row[10] : 'Awaiting trigger execution')
      };
    }).reverse();
  } catch (e) { return []; }
}

function sendHardcodedEmail(triggerHandle, toEmail, dataMap) {
  var immediateTriggers = ["System:INSTALL", "Users:RESET_REQUEST", "Users:PASSWORD_UPDATED"];
  var delivery = (immediateTriggers.indexOf(triggerHandle) > -1) ? "Immediate" : "Queue";
  enqueueEmailRow(triggerHandle, toEmail, "", "", dataMap, "Collective", delivery);
}

function legacySendHardcodedEmailBypass(triggerHandle, toEmail, dataMap) {
  var subject = "";
  var htmlBody = "";
  var wrapperName = "Internal Communication";

  switch (triggerHandle) {
    case "Logs:EXPORT":
      subject = "System Logs Exported";
      htmlBody = "<div style='font-family: sans-serif; padding: 20px;'><h2>Logs Exported</h2><p>The system logs have been successfully exported.</p><p>Details: {{details}}</p></div>";
      break;
    case "Logs:PURGE":
      subject = "System Logs Purged";
      htmlBody = "<div style='font-family: sans-serif; padding: 20px;'><h2>Logs Purged</h2><p>The system logs have been purged by an administrator.</p><p>Details: {{details}}</p></div>";
      break;
    case "Settings:UPDATE":
      subject = "System Settings Updated";
      htmlBody = "<div style='font-family: sans-serif; padding: 20px;'><h2>Settings Updated</h2><p>The global system configuration has been modified.</p><p>Details: {{details}}</p></div>";
      break;
    case "System:INSTALL":
      subject = "Installation Successful";
      htmlBody = "<div style='font-family: sans-serif; padding: 20px;'><h2>Installation Successful</h2><p>The system has been fully deployed and is ready for use as <strong>{{sysName}}</strong>.</p><hr><p><strong>Your Initial Credentials:</strong><br>Username: {{adminUsername}}<br>Role: Administrator</p></div>";
      break;
    case "System:MODULE_INSTALLED":
      subject = "New System Feature Available: {{moduleName}}";
      htmlBody = "<div style='font-family: sans-serif; padding: 20px;'><h2>System Upgrade Complete</h2><p>The <strong>{{moduleName}}</strong> module has been successfully installed into the core ecosystem.</p><p>Please refresh your dashboard to access the new features and logic.</p></div>";
      break;
    default:
      return;
  }

  htmlBody = applyGlobalSignature(htmlBody);

  // Inject the raw HTML into the wrapper
  var fullHtml = getWrapperContent(wrapperName).replace("{{USER_MESSAGE_CONTENT}}", htmlBody);
  
  // Replace variables
  for (var key in dataMap) {
      var regex = new RegExp("\\{\\{" + key + "\\}\\}", "gi");
      subject = subject.replace(regex, dataMap[key] || "");
      fullHtml = fullHtml.replace(regex, dataMap[key] || "");
  }

  var settings = getSystemSettings();
  var finalToEmail = toEmail;
  
  // Sandbox environment interceptor
  if (settings.environment === 'Sandbox' && settings.adminEmail !== '') {
      finalToEmail = settings.adminEmail;
      subject = "[Sandbox Mail] " + subject;
      var sandboxWarning = "<br><br><div style='padding: 20px; background-color: #000; color: #0f0; font-family: monospace; font-size: 14px; border: 2px solid #333; margin-top: 50px;'>";
      sandboxWarning += "SYSTEM OVERRIDE: SANDBOX ENVIRONMENT INTERCEPTED<br>";
      sandboxWarning += "&gt; INTENDED RECIPIENT: " + toEmail + "<br></div>";
      fullHtml += sandboxWarning;
  }

  // 2. Wrap all links for tracking right before sending
  fullHtml = applyLinkTracking(fullHtml, finalToEmail);

  // Dispatch via core mail app, using the dynamic system name and attached logo
  try {
      MailApp.sendEmail({
          to: finalToEmail, 
          subject: subject, 
          htmlBody: fullHtml, 
          noReply: true, 
          name: settings.systemName, 
          inlineImages: { logo: getLogoBlob() }
      });
  } catch(e) { console.warn("Failed to send hardcoded email: " + e.message); }
}

function getRenderedTemplatePreview(rowIndex) {
  // Widened scanning range block bounds to 15 columns to prevent parameter truncations
  var rowData = getMainDb().getSheetByName("Templates").getRange(parseInt(rowIndex), 1, 1, 15).getValues()[0];
  var rawHtml = rowData[8] || "";
  var wrapperType = rowData[9] || "Internal"; 
  
  var fullHtml = getWrapperContent(wrapperType).replace("{{USER_MESSAGE_CONTENT}}", rawHtml);
  var settings = getSystemSettings();
  var logo = settings.systemLogoId ? settings.systemLogoUrl : settings.appFallbackLogo;
  return fullHtml.replace(/src="cid:logo"/g, 'src="' + logo + '"');
}

/**
 * Saves the Templates module specific configuration.
 */
function saveTemplatesModuleConfig(payload) {
  try {
    if (!payload) throw new Error("No payload provided.");
    const props = PropertiesService.getScriptProperties();
    
    // Verify target schema context natively inside the communication layer before persistence commit
    if (payload.queueDbId) {
      validateSpreadsheetSchema(payload.queueDbId, ["Email Queue"], "Email Queue Database");
    }
    
    props.setProperty('TPL_DEFAULT_WRAPPER', payload.defaultWrapper);
    props.setProperty('TPL_BCC_ARCHIVE', payload.bccArchive);
    props.setProperty('TPL_LINK_TRACKING', payload.linkTracking ? 'true' : 'false');
    props.setProperty('TPL_WHITELIST', payload.whitelist);
    props.setProperty('TPL_GLOBAL_SIGNATURE', payload.signature);
    if (payload.queueDbId) props.setProperty('QUEUE_DATABASE_ID', payload.queueDbId);

    SystemEvent.emit("System", "UPDATE", "Config Updated", "INFO", "-", "Template module settings updated locally.", "", { targetName: "Templates" });
    return { success: true };
  } catch (e) {
    return { error: "Templates.gs: " + e.message };
  }
}

function getTemplatesModuleConfig() {
  try {
    var props = PropertiesService.getScriptProperties();
    return {
      tplDefaultWrapper: props.getProperty('TPL_DEFAULT_WRAPPER') || 'Internal Communication',
      tplBccArchive: props.getProperty('TPL_BCC_ARCHIVE') || '',
      tplLinkTracking: props.getProperty('TPL_LINK_TRACKING') || 'false',
      tplWhitelist: props.getProperty('TPL_WHITELIST') || '',
      tplGlobalSignature: props.getProperty('TPL_GLOBAL_SIGNATURE') || '',
      queueDbId: props.getProperty('QUEUE_DATABASE_ID') || ''
    };
  } catch(e) {
    return { tplDefaultWrapper: 'Internal Communication', tplBccArchive: '', tplLinkTracking: 'false', tplWhitelist: '', tplGlobalSignature: '', queueDbId: '' };
  }
}

// ========================================================================
// 3. READ / GET FUNCTIONS
// ========================================================================
function getTemplatesList() {
  try {
    var data = getMainDb().getSheetByName("Templates").getDataRange().getDisplayValues();
    data.shift();
    return data.map(function(row, index) {
      return {
        rowIndex: index + 2, timestamp: row[0], id: row[1], name: row[2], 
        description: row[3], category: row[4], module: row[5], trigger: row[6], 
        subject: row[7], wrapper: row[9], status: row[10],
        to: row[11] || "", cc: row[12] || "", bcc: row[13] || "", dispatchMode: row[14] || "Individual",
        triggerInclusions: row[15] || "", triggerExclusions: row[16] || "",
        recipientInclusions: row[17] || "", recipientExclusions: row[18] || "",
        deliveryMethod: row[19] || "Queue"
      };
    });
  } catch (e) { return []; }
}

function getTemplateById(rowIndex) {
  try {
    // Extends spreadsheet lookup range to 20 columns to safely process delivery method values
    var row = getMainDb().getSheetByName("Templates").getRange(parseInt(rowIndex), 1, 1, 20).getDisplayValues()[0];
    var templateName = row[2];
    var lastUpdated = getEventTimestampFromLogs("Templates", "UPDATE", templateName);
    return {
      rowIndex: rowIndex, timestamp: row[0], id: row[1], name: templateName, description: row[3], category: row[4],
      module: row[5], trigger: row[6], subject: row[7], body: row[8], wrapper: row[9], status: row[10],
      to: row[11] || "", cc: row[12] || "", bcc: row[13] || "", dispatchMode: row[14] || "Individual",
      triggerInclusions: row[15] || "", triggerExclusions: row[16] || "",
      recipientInclusions: row[17] || "", recipientExclusions: row[18] || "",
      deliveryMethod: row[19] || "Queue",
      lastUpdated: lastUpdated
    };
  } catch (e) { return { error: e.message }; }
}

function getWrappersList() {
  try {
    var sheet = ensureWrappersSheet();
    var data = sheet.getDataRange().getDisplayValues();
    data.shift();
    return data.map(function(row, i) {
      return { 
        rowIndex: i + 2, timestamp: row[0], id: row[1], name: row[2], 
        description: row[3], html: row[4], status: row[5] 
      };
    });
  } catch (e) { return []; }
}

function getWrapperById(rowIndex) {
  try {
    var row = ensureWrappersSheet().getRange(parseInt(rowIndex), 1, 1, 6).getDisplayValues()[0];
    var wrapperName = row[2];
    var lastUpdated = getEventTimestampFromLogs("Templates:Wrappers", "UPDATE", wrapperName);
    return {
      rowIndex: rowIndex, timestamp: row[0], id: row[1], name: wrapperName, description: row[3],
      html: row[4], status: row[5], lastUpdated: lastUpdated
    };
  } catch (e) { return { error: e.message }; }
}

function getWrapperContent(wrapperName) {
  try {
    var data = ensureWrappersSheet().getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === wrapperName && data[i][5] === "Active") return data[i][4];
    }
    return "{{USER_MESSAGE_CONTENT}}";
  } catch (e) { return "{{USER_MESSAGE_CONTENT}}"; }
}

// ========================================================================
// 4. WRITE / SAVE FUNCTIONS
// ========================================================================
function saveTemplateRecord(data) {
  try {
    var sheet = getMainDb().getSheetByName("Templates");
    var autoModule = data.trigger ? data.trigger.split(':')[0] : "System";
    if (autoModule === "Roles") autoModule = "Users";
    var values = [
      data.timestamp || new Date(), 
      data.id || "TPL-" + Utilities.getUuid().substring(0,8),
      data.name, data.description, data.category, autoModule, 
      data.trigger, data.subject, data.body, data.wrapper, data.status,
      data.to || "", data.cc || "", data.bcc || "", data.dispatchMode || "Individual",
      data.triggerInclusions || "", data.triggerExclusions || "",
      data.recipientInclusions || "", data.recipientExclusions || "",
      data.deliveryMethod || "Queue"
    ];
    var targetRow;
    var tplId = data.id || values[1];
    if (data.rowIndex) {
      targetRow = parseInt(data.rowIndex);
      var oldStatus = sheet.getRange(targetRow, 11).getValue();
      // Expanded structural write array parameters to column index 20
      sheet.getRange(targetRow, 1, 1, 20).setValues([values]);
      SystemEvent.emit("Templates", "UPDATE", "Edit Template", "INFO", tplId, "Template content or logic updated.", "", { targetName: data.name });
      // Granular Activation Logging
      if (oldStatus !== data.status) {
        var actionVerb = data.status === "Active" ?
        "activated" : "deactivated";
        SystemEvent.emit("Templates", "UPDATE", "Template Status Changed", "WARN", tplId, "Template was manually " + actionVerb + ".", "", { targetName: data.name });
      }
    } else {
      sheet.appendRow(values);
      targetRow = sheet.getLastRow();
      SystemEvent.emit("Templates", "CREATE", "Create Template", "INFO", tplId, "New template created.", "", { targetName: data.name });
    }
    return { success: true, rowIndex: targetRow, message: "Success! Template synced." };
  } catch (e) { return { error: "Error: " + e.message }; }
}

function saveWrapperRecord(data) {
  try {
    var sheet = ensureWrappersSheet();
    var values = [
      data.timestamp || new Date(), 
      data.id || "W-" + Utilities.getUuid().substring(0,8).toUpperCase(), 
      data.name, data.description || "", data.html, data.status
    ];
    var targetRow;
    var wrapId = data.id || values[1];
    if (data.rowIndex) {
      targetRow = parseInt(data.rowIndex);
      var oldStatus = sheet.getRange(targetRow, 6).getValue();
      sheet.getRange(targetRow, 1, 1, 6).setValues([values]);
      SystemEvent.emit("Templates:Wrappers", "UPDATE", "Edit Wrapper", "INFO", wrapId, "Wrapper layout HTML or settings updated.", "", { targetName: data.name });
      // Granular Activation Logging
      if (oldStatus !== data.status) {
        var actionVerb = data.status === "Active" ?
        "activated" : "deactivated";
        SystemEvent.emit("Templates:Wrappers", "UPDATE", "Wrapper Status Changed", "WARN", wrapId, "Wrapper layout was manually " + actionVerb + ".", "", { targetName: data.name });
      }
    } else {
      sheet.appendRow(values);
      targetRow = sheet.getLastRow();
      SystemEvent.emit("Templates:Wrappers", "CREATE", "Create Wrapper", "INFO", wrapId, "New wrapper layout created.", "", { targetName: data.name });
    }
    return { success: true, rowIndex: targetRow, message: "Success! Wrapper updated." };
  } catch (e) { return { error: "Error: " + e.message }; }
}

// ========================================================================
// 5. INTERNAL HELPERS
// ========================================================================
function ensureWrappersSheet() {
  var ss = getMainDb();
  var sheet = ss.getSheetByName("Wrappers") || ss.insertSheet("Wrappers");
  if (sheet.getLastRow() < 2) {
    sheet.getRange(1, 1, 1, 6).setValues([["Timestamp", "Wrapper ID", "Name", "Description", "HTML Content", "Status"]]).setFontWeight("bold");
    sheet.setFrozenRows(1);

    var internalHtml = `<div style="background-color: #f4f6f9; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;"><div style="background-color: #323232; padding: 25px; text-align: center; border-bottom: 4px solid #666DF2;"><img src="cid:logo" alt="Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"></div><div style="padding: 30px; color: #444; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; border-top: 1px solid #eee; background-color: #fcfcfc; text-align: center; font-size: 11px; color: #888;">This is an automated system notification.</div></div></div>`;
    var externalHtml = `<div style="background-color: #ffffff; padding: 40px 20px; font-family: Arial, sans-serif; border: 1px solid #eee;"><div style="max-width: 600px; margin: 0 auto;"><div style="padding-bottom: 20px; border-bottom: 1px solid #ddd; margin-bottom: 20px; text-align: center;"><img src="cid:logo" alt="Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"></div><div style="color: #555; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="margin-top: 40px; font-size: 12px; color: #999; border-top: 1px solid #eee; padding-top: 15px;">Sent from the Hub Team.</div></div></div>`;
    var userHtml = `<div style="background-color: #f8fafc; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 500px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px rgba(0,0,0,0.05);"><div style="padding: 30px; text-align: center; border-bottom: 1px solid #f1f5f9;"><img src="cid:logo" alt="Logo" style="max-width: 120px; height: auto;"></div><div style="padding: 30px; color: #334155; line-height: 1.6; font-size: 15px;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; background-color: #f8fafc; text-align: center; font-size: 12px; color: #94a3b8; border-top: 1px solid #e2e8f0;">Security & Access Notification</div></div></div>`;
    
    sheet.appendRow([new Date(), "W-INT", "Internal Communication", "Standard internal messaging", internalHtml, "Active"]);
    sheet.appendRow([new Date(), "W-EXT", "External Communication", "Client-facing messaging", externalHtml, "Active"]);
    sheet.appendRow([new Date(), "W-USER", "User Communications", "Dedicated layout for user access and security emails", userHtml, "Active"]);
  }
  return sheet;
}

/**
 * Scans an HTML body and wraps valid URLs in the SparkHub tracking endpoint.
 */
function applyLinkTracking(htmlBody, recipientEmail) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('TPL_LINK_TRACKING') !== 'true') return htmlBody;
  
  var scriptUrl = ScriptApp.getService().getUrl();
  if (!scriptUrl) return htmlBody; // Failsafe if script URL isn't available
  
  // Find all href="http..." attributes and rewrite them
  return htmlBody.replace(/href=["'](https?:\/\/[^"']+)["']/gi, function(match, originalUrl) {
    // Prevent double-wrapping if it already points to a Google Script
    if (originalUrl.indexOf('script.google.com') > -1) return match;
    
    var trackingUrl = scriptUrl + "?action=trackEvent&dest=" + encodeURIComponent(originalUrl);
    if (recipientEmail) {
      trackingUrl += "&u=" + encodeURIComponent(recipientEmail);
    }
    
    return 'href="' + trackingUrl + '"';
  });
}

/**
 * Appends the Global Signature to the email content.
 */
function applyGlobalSignature(htmlBody) {
  var signature = PropertiesService.getScriptProperties().getProperty('TPL_GLOBAL_SIGNATURE');
  if (!signature || signature.trim() === '') return htmlBody;
  
  // Convert standard newlines to HTML breaks if it's plain text
  var formattedSig = signature;
  if (formattedSig.indexOf('<') === -1) {
    formattedSig = formattedSig.replace(/\n/g, '<br>');
  }
  
  var signatureBlock = '<div style="margin-top: 30px; padding-top: 15px; border-top: 1px solid #e2e8f0; font-size: 12px; color: #64748b;">' + formattedSig + '</div>';
  return htmlBody + signatureBlock;
}

function Templates_getPlaceholderMetadata() {
  return {
    "templateName": { desc: "The customizable descriptive name assigned to an email template inside the library configuration panel.", tag: "Templates & Wrappers" },
    "wrapperName": { desc: "The designated layout blueprint wrapping component attached to the template broadcast context.", tag: "Templates & Wrappers" },
    "triggerEvent": { desc: "The specialized execution hook descriptor string mapping an action to a trigger (e.g., Users:CREATE).", tag: "Templates & Wrappers" }
  };
}

function Templates_getModuleLabel() { return "Templates & Wrappers"; }

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */