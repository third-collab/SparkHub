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

function getDynamicTriggerRegistry() {
  var triggers = [];
  var globalScope = typeof globalThis !== 'undefined' ? globalThis : this;
  
  for (var key in globalScope) {
    if (typeof key === 'string' && key.endsWith('_getTriggers') && typeof globalScope[key] === 'function') {
      triggers = triggers.concat(globalScope[key]());
    }
  }
  
  var installed = PropertiesService.getScriptProperties().getProperty('INSTALLED_MODULES');
  if (installed) {
    installed.split(',').forEach(function(modName) {
      var mod = modName.trim();
      var funcName = mod + "_getTriggers";
      if (typeof globalScope[funcName] !== 'function') {
        triggers.push(mod + ":CREATE");
        triggers.push(mod + ":UPDATE");
      }
    });
  }
  
  return [...new Set(triggers)].sort();
}

function getPlaceholderSuggestions() {
  var placeholders = ["details"];
  var globalScope = typeof globalThis !== 'undefined' ? globalThis : this;
  
  for (var key in globalScope) {
    if (typeof key === 'string' && key.endsWith('_getPlaceholders') && typeof globalScope[key] === 'function') {
      placeholders = placeholders.concat(globalScope[key]());
    }
  }
  
  return [...new Set(placeholders)].sort();
}

// ========================================================================
// 2. CORE PROCESSORS
// ========================================================================
var Templates = {
  handleEventEmail: function(payload) {
    if (!payload.recipientEmail) return;
    var dataMap = payload.extraData || {};
    dataMap.username = payload.entity;
    dataMap.details = payload.details;
    dataMap.systemName = getSystemSettings().systemName;
    
    sendTriggerEmail(payload.handle, payload.recipientEmail, dataMap);
  }
};

function getRenderedTemplatePreview(rowIndex) {
  var rowData = getMainDb().getSheetByName("Templates").getRange(parseInt(rowIndex), 1, 1, 11).getValues()[0];
  var rawHtml = rowData[8] || "";
  var wrapperType = rowData[9] || "Internal"; 
  
  var fullHtml = getWrapperContent(wrapperType).replace("{{USER_MESSAGE_CONTENT}}", rawHtml);
  var settings = getSystemSettings();
  var logo = settings.systemLogoId ? settings.systemLogoUrl : settings.appFallbackLogo;
  return fullHtml.replace(/src="cid:logo"/g, 'src="' + logo + '"');
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
        subject: row[7], status: row[10], wrapper: row[9]
      };
    });
  } catch (e) { return []; }
}

function getTemplateById(rowIndex) {
  try {
    var row = getMainDb().getSheetByName("Templates").getRange(parseInt(rowIndex), 1, 1, 11).getDisplayValues()[0];
    var templateName = row[2];
    var lastUpdated = getEventTimestampFromLogs("Templates", "UPDATE", templateName);
    return {
      rowIndex: rowIndex, timestamp: row[0], id: row[1], name: templateName, description: row[3], category: row[4],
      module: row[5], trigger: row[6], subject: row[7], body: row[8], wrapper: row[9], status: row[10],
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
      data.trigger, data.subject, data.body, data.wrapper, data.status
    ];
    var targetRow;
    if (data.rowIndex) {
      targetRow = parseInt(data.rowIndex);
      var oldStatus = sheet.getRange(targetRow, 11).getValue();
      sheet.getRange(targetRow, 1, 1, 11).setValues([values]);
      SystemEvent.emit("Templates", "UPDATE", "Edit Template", "INFO", data.name, "Template content or logic updated.");
      if (oldStatus !== data.status) {
        var actionVerb = data.status === "Active" ? "activated" : "deactivated";
        SystemEvent.emit("Templates", "UPDATE", "Template Status Changed", "WARN", data.name, "Template was manually " + actionVerb + ".");
      }
    } else {
      sheet.appendRow(values);
      targetRow = sheet.getLastRow();
      SystemEvent.emit("Templates", "CREATE", "Create Template", "INFO", data.name, "New template created.");
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
    if (data.rowIndex) {
      targetRow = parseInt(data.rowIndex);
      var oldStatus = sheet.getRange(targetRow, 6).getValue();
      sheet.getRange(targetRow, 1, 1, 6).setValues([values]);
      SystemEvent.emit("Templates:Wrappers", "UPDATE", "Edit Wrapper", "INFO", data.name, "Wrapper layout HTML or settings updated.");
      if (oldStatus !== data.status) {
        var actionVerb = data.status === "Active" ? "activated" : "deactivated";
        SystemEvent.emit("Templates:Wrappers", "UPDATE", "Wrapper Status Changed", "WARN", data.name, "Wrapper layout was manually " + actionVerb + ".");
      }
    } else {
      sheet.appendRow(values);
      targetRow = sheet.getLastRow();
      SystemEvent.emit("Templates:Wrappers", "CREATE", "Create Wrapper", "INFO", data.name, "New wrapper layout created.");
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

function getWrapperContent(wrapperName) {
  try {
    var data = ensureWrappersSheet().getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === wrapperName && data[i][5] === "Active") return data[i][4];
    }
    return "{{USER_MESSAGE_CONTENT}}";
  } catch (e) { return "{{USER_MESSAGE_CONTENT}}"; }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */