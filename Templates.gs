/**
 * Templates Module - Backend
 * Standardized under SparkHub Architecture Blueprint.
 * * CORE RESPONSIBILITIES:
 * - CRUD operations for system-wide email/doc templates.
 * - Logic synchronization with EmailEngine.gs (10-column schema).
 * - Placeholder harvesting for UI suggestions.
 */

// Inside a module file like Templates.gs
var Templates = {
  handleEventEmail: function(payload) {
    // Use the payload to determine if an email should be sent
    // and call your existing sendTriggerEmail logic.
    sendTriggerEmail(payload.name, payload.adminEmail, {
      "username": payload.user,
      "details": payload.details
    });
  }
}

/**
 * Fetches a summarized list of all templates for the management table.
 * Standardized handle: getMainDb()
 */
function getTemplatesList() {
  try {
    var sheet = getMainDb().getSheetByName("Templates");
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    data.shift(); // Remove headers

    return data.map(function(row, index) {
      return {
        rowIndex: index + 2,
        id: row[0],
        name: row[1],
        category: row[2],
        // Aligning indices with EmailEngine.gs logic
        trigger: row[5], 
        subject: row[6],
        status: row[8] || "Draft",
        wrapper: row[9] || "Internal"
      };
    });
  } catch (e) {
    console.error("getTemplatesList error: " + e.message);
    return [];
  }
}

/**
 * Retrieves the full content and metadata for a specific template.
 * Restored 10-column parity with EmailEngine.gs
 */
function getTemplateById(rowIndex) {
  try {
    var ss = getMainDb();
    var sheet = ss.getSheetByName("Templates");
    var idx = parseInt(rowIndex, 10);
    var row = sheet.getRange(idx, 1, 1, 10).getValues()[0];

    if (!row[0] && !row[1]) throw new Error("Template logic error: record not found at row " + rowIndex);

    return {
      rowIndex: idx,
      id: row[0],
      name: row[1],
      category: row[2],
      description: row[3],
      lastUpdatedBy: row[4],
      trigger: row[5],
      subject: row[6],
      body: row[7],
      status: row[8],
      wrapper: row[9]
    };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Updates an existing template or creates a new one.
 * Enforces the 10-column schema to prevent EmailEngine.gs index shifts.
 */
function updateTemplateRecord(data) {
  try {
    var sheet = getMainDb().getSheetByName("Templates");
    var values = [
      data.id || ("TPL-" + Utilities.formatDate(new Date(), "GMT", "yyyyMMdd-HHmm")), // Col 1: ID
      data.name,        // Col 2: Name
      data.category,    // Col 3: Category
      data.description, // Col 4: Internal Desc
      CURRENT_USERNAME, // Col 5: Last Editor
      data.trigger,     // Col 6: Trigger Event (EmailEngine index 5)
      data.subject,     // Col 7: Subject (EmailEngine index 6)
      data.body,        // Col 8: Body (EmailEngine index 7)
      data.status,      // Col 9: Status (EmailEngine index 8)
      data.wrapper      // Col 10: Wrapper (EmailEngine index 9)
    ];

    if (data.rowIndex) {
      var idx = parseInt(data.rowIndex, 10);
      sheet.getRange(idx, 1, 1, 10).setValues([values]);
      // NEW LOG:
      SystemEvent.emit("Templates", "UPDATE", "Edit Template", "INFO", data.name, "Template logic or design was modified.");
    } else {
      sheet.appendRow(values);
      // NEW LOG:
      SystemEvent.emit("Templates", "CREATE", "Add Template", "INFO", data.name, "New system communication template registered.");
    }

    return "Success! Template synced to master registry.";
  } catch (e) {
    return "Error: " + e.message;
  }
}

/**
 * Returns a list of supported placeholders for the UI editor.
 */
function getPlaceholderSuggestions() {
  return [
    "companyName", "brandName", "address", "website", "priFirstName", 
    "priLastName", "priEmail", "monthlyContractValue", "contractStartDate", 
    "services", "notes", "username", "firstName", "lastName", "role"
  ];
}


/**
 * Helper to fetch the raw HTML content of a specific wrapper file.
 * @param {string} type - The wrapper filename prefix (e.g., 'Internal', 'External').
 * @return {string} The raw HTML content of the file.
 */
function getWrapperContent(type) {
  var fileName = type + "Wrapper"; 
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

/**
 * Merges the designated wrapper and body content for the UI template preview.
 * Swaps CID references for live URLs (SVG or Drive Thumbnail) so images render in the browser.
 * @param {number} rowIndex - The row index of the template in the spreadsheet.
 * @return {string} The fully rendered HTML string.
 */
function getRenderedTemplatePreview(rowIndex) {
  var sheet = getMainDb().getSheetByName("Templates");
  var data = sheet.getDataRange().getValues();
  var rowData = data[rowIndex];
  
  var rawHtml = rowData[7] || "";
  var wrapperType = rowData[9] || "Internal";
  
  var wrapperHtml = getWrapperContent(wrapperType);
  var fullHtml = wrapperHtml.replace("{{USER_MESSAGE_CONTENT}}", rawHtml);
  
  // Use centralized settings to find the correct logo for browser display
  var settings = getSystemSettings();
  var defaultSvgLogo = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'%3E%3Crect width='100' height='100' rx='20' fill='%23C40004'/%3E%3Ctext x='50' y='65' font-family='Arial' font-size='40' font-weight='bold' fill='white' text-anchor='middle'%3EMR%3C/text%3E%3C/svg%3E";
  
  // Browser preview can render SVG; if no Drive logo exists, use the SVG
  var displayLogoUrl = settings.systemLogoId ? settings.systemLogoUrl : defaultSvgLogo;
  
  // Replace CID with the dynamic URL for browser rendering
  fullHtml = fullHtml.replace(/src="cid:logo"/g, 'src="' + displayLogoUrl + '"');
  return fullHtml;
}

/**
 * Failsafe: Builds the Wrappers sheet if it doesn't exist.
 */
/**
 * Failsafe: Builds the Wrappers sheet and seeds it with your provided files.
 */
function ensureWrappersSheet() {
  var ss = getMainDb();
  var sheet = ss.getSheetByName("Wrappers") || ss.insertSheet("Wrappers");
  
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([["Wrapper ID", "Name", "HTML Content", "Status"]]).setFontWeight("bold");
    sheet.setFrozenRows(1);

    // SEEDING FROM YOUR PROVIDED FILES
    var internalHtml = `<div style="background-color: #f4f6f9; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;"><div style="background-color: #323232; padding: 25px; text-align: center; border-bottom: 4px solid #F1C40F;"><img src="cid:logo" alt="Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"><h1 style="color: #ffffff; margin: 0; font-size: 20px;">MegaRhino</h1></div><div style="padding: 30px; color: #444; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; border-top: 1px solid #eee; background-color: #fcfcfc; text-align: center; font-size: 11px; color: #888;">This is an automated system notification.<br>Please do not reply to this email.</div></div></div>`;
    
    var externalHtml = `<div style="background-color: #ffffff; padding: 40px 20px; font-family: Arial, sans-serif; border: 1px solid #eee;"><div style="max-width: 600px; margin: 0 auto;"><div style="padding-bottom: 20px; border-bottom: 1px solid #ddd; margin-bottom: 20px; text-align: center;"><img src="cid:logo" alt="Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"><h2 style="color: #333; margin: 0;">MegaRhino</h2></div><div style="color: #555; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="margin-top: 40px; font-size: 12px; color: #999; border-top: 1px solid #eee; padding-top: 15px;">Sent from the MegaRhino Team.<br><span style="font-size: 11px;">Please do not reply to this email.</span></div></div></div>`;

    sheet.appendRow(["W-INT-HUB", "Internal Hub", internalHtml, "Active"]);
    sheet.appendRow(["W-EXT-CLIENT", "External Client", externalHtml, "Active"]);
  }
  return sheet;
}

function getWrappersList() {
  try {
    var sheet = ensureWrappersSheet();
    var data = sheet.getDataRange().getValues();
    data.shift();
    return data.map(function(row, i) {
      return { rowIndex: i + 2, id: row[0], name: row[1], html: row[2], status: row[3] };
    });
  } catch (e) { return []; }
}

/**
 * REPLACED: Helper to fetch wrapper content from Database instead of File.
 */
function getWrapperContent(wrapperName) {
  try {
    var sheet = ensureWrappersSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === wrapperName && data[i][3] === "Active") {
        return data[i][2]; // Return the HTML content from Column 3
      }
    }
    // Fallback if not found
    return "{{USER_MESSAGE_CONTENT}}";
  } catch (e) { return "{{USER_MESSAGE_CONTENT}}"; }
}

/**
 * Saves or updates a wrapper record.
 */
function updateWrapperRecord(data) {
  try {
    var sheet = ensureWrappersSheet();
    var values = [data.id || "W-" + Utilities.getUuid().substring(0,8).toUpperCase(), data.name, data.html, data.status];
    if (data.rowIndex) {
      sheet.getRange(parseInt(data.rowIndex), 1, 1, 4).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
    return "Success! Wrapper updated.";
  } catch (e) { return "Error: " + e.message; }
}