/**
 * Templates Module - Backend
 * Standardized under SparkHub Architecture Blueprint.
 */

var Templates = {
  handleEventEmail: function(payload) {
    sendTriggerEmail(payload.name, payload.adminEmail, {
      "username": payload.user,
      "details": payload.details
    });
  }
};

function getTemplatesList() {
  try {
    var sheet = getMainDb().getSheetByName("Templates");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    data.shift();
    return data.map(function(row, index) {
      return {
        rowIndex: index + 2,
        timestamp: row[0], // NEW
        id: row[1],
        name: row[2],
        category: row[3],
        trigger: row[6], 
        subject: row[7],
        status: row[9] || "Draft",
        wrapper: row[10] || "Internal"
      };
    });
  } catch (e) { return []; }
}

function getTemplateById(rowIndex) {
  try {
    // Fetch 11 columns
    var row = getMainDb().getSheetByName("Templates").getRange(parseInt(rowIndex, 10), 1, 1, 11).getValues()[0];
    return {
      rowIndex: rowIndex,
      timestamp: row[0],
      id: row[1],
      name: row[2],
      category: row[3],
      description: row[4],
      lastUpdatedBy: row[5],
      trigger: row[6],
      subject: row[7],
      body: row[8],
      status: row[9],
      wrapper: row[10]
    };
  } catch (e) { return { error: e.message }; }
}

function updateTemplateRecord(data) {
  try {
    var sheet = getMainDb().getSheetByName("Templates");
    var values = [
      data.timestamp || new Date(), // Col 1: Timestamp
      data.id || ("TPL-" + Utilities.formatDate(new Date(), "GMT", "yyyyMMdd-HHmm")),
      data.name, data.category, data.description, getLoggedInUsername(), 
      data.trigger, data.subject, data.body, data.status, data.wrapper
    ];
    if (data.rowIndex) {
      sheet.getRange(parseInt(data.rowIndex), 1, 1, 11).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
    return "Success! Template synced.";
  } catch (e) { return "Error: " + e.message; }
}

function getPlaceholderSuggestions() {
  return ["companyName", "brandName", "address", "website", "priFirstName", "priLastName", "priEmail", "monthlyContractValue", "contractStartDate", "services", "notes", "username", "firstName", "lastName", "role"];
}

/** * WRAPPER ENGINE: Database-driven HTML frames
 */
function ensureWrappersSheet() {
  var ss = getMainDb();
  var sheet = ss.getSheetByName("Wrappers") || ss.insertSheet("Wrappers");
  // SEED FIX: Only seed if row count is less than 2 (Headers only or empty)
  if (sheet.getLastRow() < 2) {
    sheet.getRange(1, 1, 1, 4).setValues([["Wrapper ID", "Name", "HTML Content", "Status"]]).setFontWeight("bold");
    sheet.setFrozenRows(1);

    var internalHtml = `<div style="background-color: #f4f6f9; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;"><div style="background-color: #323232; padding: 25px; text-align: center; border-bottom: 4px solid #F1C404;"><img src="cid:logo" alt="Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"></div><div style="padding: 30px; color: #444; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; border-top: 1px solid #eee; background-color: #fcfcfc; text-align: center; font-size: 11px; color: #888;">This is an automated system notification.</div></div></div>`;
    var externalHtml = `<div style="background-color: #ffffff; padding: 40px 20px; font-family: Arial, sans-serif; border: 1px solid #eee;"><div style="max-width: 600px; margin: 0 auto;"><div style="padding-bottom: 20px; border-bottom: 1px solid #ddd; margin-bottom: 20px; text-align: center;"><img src="cid:logo" alt="Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"></div><div style="color: #555; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="margin-top: 40px; font-size: 12px; color: #999; border-top: 1px solid #eee; padding-top: 15px;">Sent from the Hub Team.</div></div></div>`;

    sheet.appendRow(["W-INT", "Internal", internalHtml, "Active"]);
    sheet.appendRow(["W-EXT", "External", externalHtml, "Active"]);
  }
  return sheet;
}

function getWrappersList() {
  try {
    var sheet = ensureWrappersSheet();
    var data = sheet.getDataRange().getValues();
    data.shift();
    return data.map(function(row, i) {
      // row[0] is Timestamp, row[1] is ID, row[2] is Name...
      return { rowIndex: i + 2, timestamp: row[0], id: row[1], name: row[2], html: row[3], status: row[4] };
    });
  } catch (e) { return []; }
}

function updateWrapperRecord(data) {
  try {
    var sheet = ensureWrappersSheet();
    var values = [
      data.timestamp || new Date(), // NEW
      data.id || "W-" + Utilities.getUuid().substring(0,8).toUpperCase(), 
      data.name, data.html, data.status
    ];
    if (data.rowIndex) {
      sheet.getRange(parseInt(data.rowIndex), 1, 1, 5).setValues([values]);
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
      if (data[i][1] === wrapperName && data[i][3] === "Active") return data[i][2];
    }
    return "{{USER_MESSAGE_CONTENT}}";
  } catch (e) { return "{{USER_MESSAGE_CONTENT}}"; }
}

function getRenderedTemplatePreview(rowIndex) {
  var rowData = getMainDb().getSheetByName("Templates").getRange(parseInt(rowIndex), 1, 1, 10).getValues()[0];
  var rawHtml = rowData[7] || "";
  var wrapperType = rowData[9] || "Internal";
  var fullHtml = getWrapperContent(wrapperType).replace("{{USER_MESSAGE_CONTENT}}", rawHtml);
  var settings = getSystemSettings();
  var logo = settings.systemLogoId ? settings.systemLogoUrl : "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'%3E%3Crect width='100' height='100' rx='20' fill='%23C40004'/%3E%3Ctext x='50' y='65' font-family='Arial' font-size='40' font-weight='bold' fill='white' text-anchor='middle'%3EMR%3C/text%3E%3C/svg%3E";
  return fullHtml.replace(/src="cid:logo"/g, 'src="' + logo + '"');
}