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
    var row = getMainDb().getSheetByName("Templates").getRange(parseInt(rowIndex), 1, 1, 11).getValues()[0];
    return {
      rowIndex: rowIndex, timestamp: row[0], id: row[1], name: row[2], description: row[3], category: row[4],
      module: row[5], trigger: row[6], subject: row[7], body: row[8], wrapper: row[9], status: row[10]
    };
  } catch (e) { return { error: e.message }; }
}

function updateTemplateRecord(data) {
  try {
    var sheet = getMainDb().getSheetByName("Templates");
    
    // Auto-populate the Module based on the text before the colon in the trigger
    var autoModule = data.trigger ? data.trigger.split(':')[0] : "System";
    
    var values = [
      data.timestamp || new Date(), 
      data.id || "TPL-" + Utilities.getUuid().substring(0,8),
      data.name, data.description, data.category, autoModule, 
      data.trigger, data.subject, data.body, data.wrapper, data.status
    ];
    
    if (data.rowIndex) {
      sheet.getRange(parseInt(data.rowIndex), 1, 1, 11).setValues([values]);
      SystemEvent.emit("Templates", "UPDATE", "Edit Template", "INFO", data.name, "Template content or logic updated.");
    } else {
      sheet.appendRow(values);
      SystemEvent.emit("Templates", "CREATE", "Create Template", "INFO", data.name, "New template created.");
    }
    return "Success! Template synced.";
  } catch (e) { return "Error: " + e.message; }
}

function getRenderedTemplatePreview(rowIndex) {
  var rowData = getMainDb().getSheetByName("Templates").getRange(parseInt(rowIndex), 1, 1, 11).getValues()[0];
  var rawHtml = rowData[8] || ""; // Body is now Col 9 (Index 8)
  var wrapperType = rowData[9] || "Internal"; // Wrapper is now Col 10 (Index 9)
  
  var fullHtml = getWrapperContent(wrapperType).replace("{{USER_MESSAGE_CONTENT}}", rawHtml);
  var settings = getSystemSettings();
  var logo = settings.systemLogoId ? settings.systemLogoUrl : "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'%3E%3Crect width='100' height='100' rx='20' fill='%23C40004'/%3E%3Ctext x='50' y='65' font-family='Arial' font-size='40' font-weight='bold' fill='white' text-anchor='middle'%3EMR%3C/text%3E%3C/svg%3E";
  return fullHtml.replace(/src="cid:logo"/g, 'src="' + logo + '"');
}

function getPlaceholderSuggestions() {
  return ["companyName", "brandName", "address", "website", "priFirstName", "priLastName", "priEmail", "monthlyContractValue", "contractStartDate", "services", "notes", "username", "firstName", "lastName", "role"];
}

/** * WRAPPER ENGINE: Database-driven HTML frames
 */
function ensureWrappersSheet() {
  var ss = getMainDb();
  var sheet = ss.getSheetByName("Wrappers") || ss.insertSheet("Wrappers");
  if (sheet.getLastRow() < 2) {
    // NEW ORDER: TS, ID, Name, Description, HTML, Status
    sheet.getRange(1, 1, 1, 6).setValues([["Timestamp", "Wrapper ID", "Name", "Description", "HTML Content", "Status"]]).setFontWeight("bold");
    sheet.setFrozenRows(1);

    var internalHtml = `<div style="background-color: #f4f6f9; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;"><div style="background-color: #323232; padding: 25px; text-align: center; border-bottom: 4px solid #F1C404;"><img src="cid:logo" alt="Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"></div><div style="padding: 30px; color: #444; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; border-top: 1px solid #eee; background-color: #fcfcfc; text-align: center; font-size: 11px; color: #888;">This is an automated system notification.</div></div></div>`;
    var externalHtml = `<div style="background-color: #ffffff; padding: 40px 20px; font-family: Arial, sans-serif; border: 1px solid #eee;"><div style="max-width: 600px; margin: 0 auto;"><div style="padding-bottom: 20px; border-bottom: 1px solid #ddd; margin-bottom: 20px; text-align: center;"><img src="cid:logo" alt="Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"></div><div style="color: #555; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="margin-top: 40px; font-size: 12px; color: #999; border-top: 1px solid #eee; padding-top: 15px;">Sent from the Hub Team.</div></div></div>`;
    var userHtml = `<div style="background-color: #f8fafc; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 500px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px rgba(0,0,0,0.05);"><div style="padding: 30px; text-align: center; border-bottom: 1px solid #f1f5f9;"><img src="cid:logo" alt="Logo" style="max-width: 120px; height: auto;"></div><div style="padding: 30px; color: #334155; line-height: 1.6; font-size: 15px;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; background-color: #f8fafc; text-align: center; font-size: 12px; color: #94a3b8; border-top: 1px solid #e2e8f0;">Security & Access Notification</div></div></div>`;
    
    // UPDATED: Replaced "Internal" and "External" with your custom names
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
    if (data.rowIndex) {
      sheet.getRange(parseInt(data.rowIndex), 1, 1, 6).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
    return "Success! Wrapper updated.";
  } catch (e) { return "Error: " + e.message; }
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