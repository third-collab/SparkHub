/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Templates.gs
 * VERSION: 1.2 (100% Logic Parity + EmailEngine Synchronization)
 * SYNC STATUS: Fully Synchronized with EmailEngine.gs & TemplatesData.html
 */

/**
 * Templates Module - Backend
 * Standardized under SparkHub Architecture Blueprint.
 * * CORE RESPONSIBILITIES:
 * - CRUD operations for system-wide email/doc templates.
 * - Logic synchronization with EmailEngine.gs (10-column schema).
 * - Placeholder harvesting for UI suggestions.
 */

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
      logNotification("Templates", "Template Updated", "Modified template logic for: " + data.name, "All", "System Update");
    } else {
      sheet.appendRow(values);
      logNotification("Templates", "New Template", "Added new system communication: " + data.name, "All", "System Update");
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
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */