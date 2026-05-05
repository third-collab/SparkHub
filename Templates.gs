// --- TEMPLATES MODULE REGISTRY EXPORTS ---
function Templates_getTriggers() {
  return [
    "Templates:CREATE", "Templates:UPDATE", 
    "Templates:Wrappers:CREATE", "Templates:Wrappers:UPDATE"
  ];
}

function Templates_getPlaceholders() {
  return ["templateName", "wrapperName", "triggerEvent"];
}

/**
 * Templates Module - Backend
 * Standardized under SparkHub Architecture Blueprint.
 */
var Templates = {
  handleEventEmail: function(payload) {
    if (!payload.recipientEmail) return;
    
    // Merge the standard variables with any extra custom variables passed from the event
    var dataMap = payload.extraData || {};
    dataMap.username = payload.entity;
    dataMap.details = payload.details;
    dataMap.systemName = getSystemSettings().systemName;
    
    sendTriggerEmail(payload.handle, payload.recipientEmail, dataMap);
  }
};

function getTemplatesList() {
  try {
    var data = getMainDb().getSheetByName("Templates").getDataRange().getDisplayValues();
    data.shift();
    return data.map(function(row, index) {
      // NEW ORDER: TS(0), ID(1), Name(2), Desc(3), Cat(4), Module(5), Trig(6), Sub(7), Body(8), Wrap(9), Status(10)
      return {
        rowIndex: index + 2,
        timestamp: row[0], id: row[1], name: row[2], description: row[3], category: row[4],
        module: row[5], trigger: row[6], subject: row[7], status: row[10], wrapper: row[9]
      };
    });
  } catch (e) { return []; }
}

function getTemplateById(rowIndex) {
  try {
    // CRITICAL FIX: Changed .getValues() to .getDisplayValues()
    // This safely serializes the Timestamp so the network request doesn't crash!
    var row = getMainDb().getSheetByName("Templates").getRange(parseInt(rowIndex), 1, 1, 11).getDisplayValues()[0];
    
    return {
      rowIndex: rowIndex, timestamp: row[0], id: row[1], name: row[2], description: row[3], category: row[4],
      module: row[5], trigger: row[6], subject: row[7], body: row[8], wrapper: row[9], status: row[10]
    };
  } catch (e) { return { error: e.message }; }
}

function updateTemplateRecord(data) {
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
      var oldStatus = sheet.getRange(targetRow, 11).getValue(); // Col K (11) is Status
      sheet.getRange(targetRow, 1, 1, 11).setValues([values]);
      
      SystemEvent.emit("Templates", "UPDATE", "Edit Template", "INFO", data.name, "Template content or logic updated.");
      
      // NEW: Granular Activation Logging
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

function getRenderedTemplatePreview(rowIndex) {
  var rowData = getMainDb().getSheetByName("Templates").getRange(parseInt(rowIndex), 1, 1, 11).getValues()[0];
  var rawHtml = rowData[8] || ""; 
  var wrapperType = rowData[9] || "Internal"; 
  
  var fullHtml = getWrapperContent(wrapperType).replace("{{USER_MESSAGE_CONTENT}}", rawHtml);
  var settings = getSystemSettings();
  
  // FIX: Using the exact configured fallback logo instead of the old MR SVG
  var logo = settings.systemLogoId ? settings.systemLogoUrl : settings.appFallbackLogo;
  return fullHtml.replace(/src="cid:logo"/g, 'src="' + logo + '"');
}

/**
 * Returns a dynamic registry of all available template placeholders.
 * Uses V8 Auto-Discovery to scan global memory for placeholder declarations.
 */
function getPlaceholderSuggestions() {
  var placeholders = ["details"]; // Universal fallback variable
  var globalScope = typeof globalThis !== 'undefined' ? globalThis : this;
  
  // V8 Auto-Discovery: Scan the entire system's memory for placeholder declarations
  for (var key in globalScope) {
    if (typeof key === 'string' && key.endsWith('_getPlaceholders') && typeof globalScope[key] === 'function') {
      placeholders = placeholders.concat(globalScope[key]());
    }
  }
  
  return [...new Set(placeholders)].sort();
}

/** * WRAPPER ENGINE: Database-driven HTML frames
 */
function ensureWrappersSheet() {
  var ss = getMainDb();
  var sheet = ss.getSheetByName("Wrappers") || ss.insertSheet("Wrappers");
  if (sheet.getLastRow() < 2) {
    sheet.getRange(1, 1, 1, 6).setValues([["Timestamp", "Wrapper ID", "Name", "Description", "HTML Content", "Status"]]).setFontWeight("bold");
    sheet.setFrozenRows(1);

    // FIX: Removed hardcoded "MegaRhino" and branding names
    var internalHtml = `<div style="background-color: #f4f6f9; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;"><div style="background-color: #323232; padding: 25px; text-align: center; border-bottom: 4px solid #666DF2;"><img src="cid:logo" alt="Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"></div><div style="padding: 30px; color: #444; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; border-top: 1px solid #eee; background-color: #fcfcfc; text-align: center; font-size: 11px; color: #888;">This is an automated system notification.</div></div></div>`;
    var externalHtml = `<div style="background-color: #ffffff; padding: 40px 20px; font-family: Arial, sans-serif; border: 1px solid #eee;"><div style="max-width: 600px; margin: 0 auto;"><div style="padding-bottom: 20px; border-bottom: 1px solid #ddd; margin-bottom: 20px; text-align: center;"><img src="cid:logo" alt="Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"></div><div style="color: #555; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="margin-top: 40px; font-size: 12px; color: #999; border-top: 1px solid #eee; padding-top: 15px;">Sent from the Hub Team.</div></div></div>`;
    var userHtml = `<div style="background-color: #f8fafc; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 500px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px rgba(0,0,0,0.05);"><div style="padding: 30px; text-align: center; border-bottom: 1px solid #f1f5f9;"><img src="cid:logo" alt="Logo" style="max-width: 120px; height: auto;"></div><div style="padding: 30px; color: #334155; line-height: 1.6; font-size: 15px;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; background-color: #f8fafc; text-align: center; font-size: 12px; color: #94a3b8; border-top: 1px solid #e2e8f0;">Security & Access Notification</div></div></div>`;
    
    sheet.appendRow([new Date(), "W-INT", "Internal Communication", "Standard internal messaging", internalHtml, "Active"]);
    sheet.appendRow([new Date(), "W-EXT", "External Communication", "Client-facing messaging", externalHtml, "Active"]);
    sheet.appendRow([new Date(), "W-USER", "User Communications", "Dedicated layout for user access and security emails", userHtml, "Active"]);
  }
  return sheet;
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

function updateWrapperRecord(data) {
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
      var oldStatus = sheet.getRange(targetRow, 6).getValue(); // Col F (6) is Status
      sheet.getRange(targetRow, 1, 1, 6).setValues([values]);
      
      SystemEvent.emit("Templates:Wrappers", "UPDATE", "Edit Wrapper", "INFO", data.name, "Wrapper layout HTML or settings updated.");

      // NEW: Granular Activation Logging
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

function getWrapperContent(wrapperName) {
  try {
    var data = ensureWrappersSheet().getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === wrapperName && data[i][5] === "Active") return data[i][4]; // HTML is Index 4
    }
    return "{{USER_MESSAGE_CONTENT}}";
  } catch (e) { return "{{USER_MESSAGE_CONTENT}}"; }
}

/**
 * Returns a dynamic registry of all available system trigger events.
 * Uses V8 Auto-Discovery to scan global memory for module declarations.
 */
function getDynamicTriggerRegistry() {
  var triggers = [];
  var globalScope = typeof globalThis !== 'undefined' ? globalThis : this;
  
  // 1. V8 Auto-Discovery: Scan the entire system's memory for trigger declarations
  for (var key in globalScope) {
    if (typeof key === 'string' && key.endsWith('_getTriggers') && typeof globalScope[key] === 'function') {
      triggers = triggers.concat(globalScope[key]());
    }
  }
  
  // 2. External Fallback: Inject standard CRUD triggers for installed modules that lack custom triggers
  var installed = PropertiesService.getScriptProperties().getProperty('INSTALLED_MODULES');
  if (installed) {
    installed.split(',').forEach(function(modName) {
      var mod = modName.trim();
      var funcName = mod + "_getTriggers";
      
      // If the module didn't announce itself above, give it the default triggers
      if (typeof globalScope[funcName] !== 'function') {
        triggers.push(mod + ":CREATE");
        triggers.push(mod + ":UPDATE");
      }
    });
  }
  
  return [...new Set(triggers)].sort();
}