/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Patch.gs
 * VERSION: 1.2
 * SYNC STATUS: Standalone Utility
 */

/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Patch.gs
 * VERSION: 2.1 (Users Side-by-Side Email Schema Remap)
 * SYNC STATUS: Standalone Architecture Utility
 */

/**
 * Migrates and re-arranges the legacy 10-column Users table to an optimized 12-column 
 * schema, placing Google Email and System Email side-by-side, and moving Status to the end.
 * @return {string} Success or failure summary message.
 */
function patch_v2_1_UsersDatabase() {
  try {
    var db = getMainDb(); // Safely retrieves the main registry database spreadsheet
    var sheet = db.getSheetByName("Users");
    if (!sheet) return "Migration Aborted: Sheet 'Users' not found.";

    var rawData = sheet.getDataRange().getValues();
    if (rawData.length === 0) return "Migration Aborted: Sheet is completely empty.";

    // Advanced 12-column header array configuration matching blueprint guidelines
    var newHeaders = [
      "Timestamp", "User ID", "Username", "Google Email", "System Email", 
      "Role", "Password", "First Name", "Last Name", "Last Login", 
      "Dashboard Config", "Status"
    ];

    var migratedRows = [newHeaders];
    var uniqueUsernamesRegistry = [];
    var counterId = 1001;

    // Process rows skipping the old header index row [0]
    for (var i = 1; i < rawData.length; i++) {
      var oldRow = rawData[i];
      if (!oldRow[1]) continue; // Skip malformed rows lacking handles

      var remappedRow = new Array(12).fill("");
      
      // A (1): Map original creation timestamp
      remappedRow[0] = oldRow[0]; 
      
      // B (2): Generate and apply the immutable unique User ID anchor
      remappedRow[1] = "U-" + counterId;
      counterId++;

      // C (3): Uniqueness handle sanitization and iterative safety routing
      var baseUsername = String(oldRow[1]).trim().toLowerCase().replace(/[^a-z0-9_]/g, "");
      var finalizedUsername = baseUsername;
      var suffixCounter = 1;

      while (uniqueUsernamesRegistry.indexOf(finalizedUsername) > -1) {
        finalizedUsername = baseUsername + suffixCounter;
        suffixCounter++;
      }
      uniqueUsernamesRegistry.push(finalizedUsername);
      remappedRow[2] = finalizedUsername; 

      // D (4) & E (5): Place the Authentication Email and Communication Email side-by-side
      var rawEmailValue = String(oldRow[3]).trim();
      remappedRow[3] = rawEmailValue; // Google Email (SSO Principal Check)
      remappedRow[4] = rawEmailValue; // System Email (Custom Alerts Target)

      // F (6) through K (11): Map standard profile attributes and configurations
      remappedRow[5] = oldRow[2];  // Role
      remappedRow[6] = oldRow[4];  // Password
      remappedRow[7] = oldRow[5];  // First Name
      remappedRow[8] = oldRow[6];  // Last Name
      remappedRow[9] = oldRow[8];  // Last Login
      remappedRow[10] = oldRow[9]; // Dashboard Config

      // L (12): Force Status to the final column index to comply with Section 3.F
      remappedRow[11] = oldRow[7] || "Active"; 

      migratedRows.push(remappedRow);
    }

    // Overwrite the previous layout structure with the re-arranged array block safely
    sheet.clearContents();
    sheet.getRange(1, 1, migratedRows.length, newHeaders.length).setValues(migratedRows);
    sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight("bold").setBackground("#f1f5f9");
    
    // Mandate immediate physical commit flush sequence to circumvent data latency profile race conditions
    SpreadsheetApp.flush(); 

    return "Success: Re-arranged database schema to 12-column layout. Synchronized " + (migratedRows.length - 1) + " profiles.";

  } catch(e) {
    console.error("User re-arrangement patch failed: " + e.message);
    return "Migration Failed: " + e.message;
  }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */

function patchDashboardConfigColumns() {
  var db = getMainDb();
  var usersSheet = db.getSheetByName("Users");
  
  if (usersSheet.getRange("J1").getValue() !== "Dashboard Config") {
    usersSheet.getRange("J1").setValue("Dashboard Config").setFontWeight("bold");
  }
  
  var rolesSheet = db.getSheetByName("Roles");
  if (rolesSheet.getRange("G1").getValue() !== "Dashboard Config") {
    rolesSheet.getRange("G1").setValue("Dashboard Config").setFontWeight("bold");
  }
  console.log("Dashboard Config columns successfully patched into Users and Roles tables!");
}

/**
 * Utility function to forcefully inject the standard suite of custom fields 
 * into the Clients Module Configuration.
 */
function seedClientsCustomFields() {
  try {
    var props = PropertiesService.getScriptProperties();
    var configStr = props.getProperty('CLIENTS_MODULE_CONFIG');
    // Setup base config object in case it doesn't exist or is in the old array format
    var configObj = { clientRole: "Client", customFields: [] };
    if (configStr) {
      try {
        var parsed = JSON.parse(configStr);
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
          configObj = parsed;
          // Preserve existing settings like the clientRole
        }
      } catch(e) {}
    }

    // The fields we migrated out of the core schema
    var customFields = [
      { id: "brand_registry", label: "Brand Registry", type: "select", options: "Registered, Pending, Not Registered", required: false },
      { id: "commission_rate", label: "Commission Rate (%)", type: "number", options: "", required: false },
      { id: "commission_basis", label: "Commission Basis", type: "text", options: "", required: false },
      { id: "ppc_date", label: "PPC Launch Date", type: "date", options: "", required: false },
      { id: "dsp_date", label: "DSP Launch Date", type: "date", options: "", required: false },
      { id: "brand_code", label: "Brand Code", type: "text", options: "", required: true },
      { id: "brand_folder", label: "Brand Folder URL", type: "url", options: "", required: false }
    ];
    // Inject the fields
    configObj.customFields = customFields;
    
    // Save to global properties
    props.setProperty('CLIENTS_MODULE_CONFIG', JSON.stringify(configObj));
    Logger.log("Success! Custom fields injected. Refresh your UI.");
    return "Success: Custom fields injected into the Clients module configuration.";
  } catch(e) {
    Logger.log("Error: " + e.message);
    return "Error injecting custom fields: " + e.message;
  }
}

/**
 * Patches the v4.1 Clients Database to the v4.2 Schema.
 * Removes the hardcoded "Additional Contacts" column and updates the headers and JSON config.
 */
function patch_v4_2_clientsDatabase() {
  try {
    var id = PropertiesService.getScriptProperties().getProperty('CLIENTS_DB_ID');
    if (!id) return "Clients database ID not found in properties.";
    
    var ss = SpreadsheetApp.openById(id);
    var sheet = ss.getSheetByName("Clients");
    if (!sheet) return "Clients sheet not found.";

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // Check if Additional Contacts is still sitting at column L (Index 11)
    if (headers[11] === "Additional Contacts") {
       sheet.deleteColumn(12);
    }
    
    // Apply the new 26-column header structure
    var newHeaders = [
      "Timestamp", "Client ID", "Company Name", "Address", "Company Phone", "Website",
      "Primary First Name", "Primary Last Name", "Primary Email", "Primary Phone", "Primary Position",
      "Services", "Rate", "Original Start Date", "Current Start Date",
      "Term Count", "Term Unit", "Original Exp Date", "Current Exp Date", "Original End Date",
      "Latest End Date", "Operational Notes", "Remarks", "Additional Fields", "History", "Status" 
    ];
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]).setFontWeight("bold").setBackground("#f1f5f9");
    
    // Update config to remove the Additional Contacts accordion group if it exists
    var confStr = PropertiesService.getScriptProperties().getProperty('CLIENTS_MODULE_CONFIG');
    if (confStr) {
        var conf = JSON.parse(confStr);
        if (conf && conf.accordionGroups) {
            conf.accordionGroups = conf.accordionGroups.filter(function(g) { return g.id !== "grp_add_contact"; });
            PropertiesService.getScriptProperties().setProperty('CLIENTS_MODULE_CONFIG', JSON.stringify(conf));
        }
    }
    
    return "Success: Clients database successfully patched to v4.2 schema.";
  } catch(e) {
    return "Migration Error: " + e.message;
  }
}

/**
 * Patches the Clients Database to the new 26-Column Schema (v4.3).
 * Extracts existing data, dynamically remaps it to the new column indices, and applies new headers.
 */
function patch_v4_3_clientsDatabase() {
  try {
    var id = PropertiesService.getScriptProperties().getProperty('CLIENTS_DB_ID');
    if (!id) return "Clients database ID not found.";
    
    var sheet = SpreadsheetApp.openById(id).getSheetByName("Clients");
    if (!sheet) return "Clients sheet not found.";

    var data = sheet.getDataRange().getValues();
    if (data.length === 0) return "Sheet is completely empty.";
    
    var newHeaders = [
      "Timestamp", "Client ID", "Company Name", "Address", "Company Email", "Company Phone", "Website", 
      "Primary First Name", "Primary Last Name", "Primary Email", "Primary Phone", "Services", "Rate", 
      "Term Unit", "Term Count", "Original Start Date", "Current Start Date", "Original Exp Date", 
      "Current Exp Date", "Original End Date", "Latest End Date", "Operational Notes", "Remarks", 
      "Additional Fields", "History", "Status"
    ];
    
    var newData = [newHeaders];
    for (var i = 1; i < data.length; i++) {
        var oldRow = data[i];
        var newRow = new Array(26).fill("");
        
        // Dynamically remap data from the previous state into the new strict 26-column index locations
        newRow[0] = oldRow[0]; // Timestamp
        newRow[1] = oldRow[1]; // Client ID
        newRow[2] = oldRow[2]; // Company Name
        newRow[3] = oldRow[3]; // Address
        newRow[4] = "";        // Company Email (NEW - Empty by Default)
        newRow[5] = oldRow[4] !== undefined ? oldRow[4] : ""; // Company Phone
        newRow[6] = oldRow[5] !== undefined ? oldRow[5] : ""; // Website
        newRow[7] = oldRow[6] !== undefined ? oldRow[6] : ""; // Primary First Name
        newRow[8] = oldRow[7] !== undefined ? oldRow[7] : ""; // Primary Last Name
        newRow[9] = oldRow[8] !== undefined ? oldRow[8] : ""; // Primary Email
        newRow[10] = oldRow[9] !== undefined ? oldRow[9] : ""; // Primary Phone
        // Index 10 in oldRow was "Primary Position" which is completely dropped
        newRow[11] = oldRow[11] !== undefined ? oldRow[11] : ""; // Services
        newRow[12] = oldRow[12] !== undefined ? oldRow[12] : ""; // Rate
        newRow[13] = oldRow[16] !== undefined ? oldRow[16] : ""; // Term Unit (Reordered)
        newRow[14] = oldRow[15] !== undefined ? oldRow[15] : ""; // Term Count (Reordered)
        newRow[15] = oldRow[13] !== undefined ? oldRow[13] : ""; // Original Start Date
        newRow[16] = oldRow[14] !== undefined ? oldRow[14] : ""; // Current Start Date
        newRow[17] = oldRow[17] !== undefined ? oldRow[17] : ""; // Original Exp Date
        newRow[18] = oldRow[18] !== undefined ? oldRow[18] : ""; // Current Exp Date
        newRow[19] = oldRow[19] !== undefined ? oldRow[19] : ""; // Original End Date
        newRow[20] = oldRow[20] !== undefined ? oldRow[20] : ""; // Latest End Date
        newRow[21] = oldRow[21] !== undefined ? oldRow[21] : "[]"; // Operational Notes
        newRow[22] = oldRow[22] !== undefined ? oldRow[22] : ""; // Remarks
        newRow[23] = oldRow[23] !== undefined ? oldRow[23] : "{}"; // Additional Fields
        newRow[24] = oldRow[24] !== undefined ? oldRow[24] : "[]"; // History
        newRow[25] = oldRow[25] !== undefined ? oldRow[25] : "Active"; // Status
        
        newData.push(newRow);
    }
    
    // Non-destructive overwrite of the remapped payload
    sheet.clearContents();
    sheet.getRange(1, 1, newData.length, newHeaders.length).setValues(newData);
    sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight("bold").setBackground("#f1f5f9");
    
    return "Database Patched to v4.3 Schema successfully.";
  } catch(e) {
    return "Error: " + e.message;
  }
}

/**
 * Architectural Upgrade Patch v5.1
 * Append To, CC, and BCC recipient routing cells onto your active Templates spreadsheet table.
 * Run this function manually from your Apps Script editor workspace panel once to upgrade sheets safely.
 */
function patch_v5_1_TemplatesDatabaseSchema() {
  try {
    var db = getMainDb();
    var sheet = db.getSheetByName("Templates");
    if (!sheet) return "Migration Cancelled: Templates sheet container element is missing.";
    
    var lastCol = sheet.getLastColumn();
    // Verify if columns have already been extended
    if (lastCol < 14) {
      sheet.getRange(1, 12, 1, 3).setValues([["To Recipients", "CC Recipients", "BCC Recipients"]]).setFontWeight("bold").setBackground("#f1f5f9");
      SpreadsheetApp.flush();
    }
    
    // Upgrades legacy 14-column systems safely up to a standard 15-column matrix ruleset bounds
    if (sheet.getLastColumn() < 15) {
      sheet.getRange(1, 15).setValue("Dispatch Mode").setFontWeight("bold").setBackground("#f1f5f9");
      SpreadsheetApp.flush();
      return "Success! Database successfully migrated to v5.2 15-Column template configuration engine schema.";
    }
    return "Notice: Sheet layout was already upgraded to v5.2 template schema bounds.";
  } catch(e) {
    return "Migration Failed: " + e.message;
  }
}

/**
 * Upgrade structural table schema to support dynamic 20-column delivery method parameters.
 * Automatically tags all core installation and password links as "Immediate".
 * @return {string} Migration outcome metrics log summary.
 */
function patch_v5_3_TemplatesDeliveryMethodSchema() {
  try {
    var db = getMainDb();
    var sheet = db.getSheetByName("Templates");
    if (!sheet) return "Migration Aborted: Sheet 'Templates' not found.";
    
    var lastCol = sheet.getLastColumn();
    // 1. Force structural sheet boundaries out to a standard 20 column layout ceiling
    if (lastCol < 20) {
      sheet.getRange(1, 20).setValue("Delivery Method").setFontWeight("bold").setBackground("#f1f5f9");
    }
    
    var data = sheet.getDataRange().getValues();
    var immediateTriggers = ["TPL-USER-NEW", "TPL-ROLE-NEW", "TPL-ADMIN-NEW", "TPL-PWD-RESET", "TPL-PWD-UPDATE"];
    var updatedCount = 0;
    
    // 2. Loop through row matrices to map parameters to matching rows safely
    for (var i = 1; i < data.length; i++) {
      var templateId = data[i][1];
      var cell = sheet.getRange(i + 1, 20);
      
      if (immediateTriggers.indexOf(templateId) > -1) {
        cell.setValue("Immediate");
      } else if (cell.getValue() === "") {
        cell.setValue("Queue");
      }
      updatedCount++;
    }
    
    SpreadsheetApp.flush();
    return "Success: Upgraded database to v5.3 20-column schema. Configured " + updatedCount + " record entries.";
  } catch(e) {
    return "Migration Failed: " + e.message;
  }
}

/**
 * Retroactively configures the pre-seeded administrative templates in the live database.
 * Restricts delivery to active administrators, ensures immediate dispatch, and applies correct wrappers.
 * @return {string} Migration status narrative log summary.
 */
function patch_v5_4_configureAdministrativeTemplates() {
  try {
    var db = getMainDb();
    var sheet = db.getSheetByName("Templates");
    if (!sheet) return "Migration Aborted: Sheet 'Templates' not found.";
    
    // Enforce 20-column schema expansion parameters if not already performed
    var lastCol = sheet.getLastColumn();
    if (lastCol < 20) {
      sheet.getRange(1, 20).setValue("Delivery Method").setFontWeight("bold").setBackground("#f1f5f9");
    }
    
    var data = sheet.getDataRange().getValues();
    var updatedCount = 0;
    
    for (var i = 1; i < data.length; i++) {
      var templateId = data[i][1]; // Column B (ID)
      
      if (templateId === "TPL-ROLE-NEW") {
        // Col 12 (L): To Recipients | Col 18 (R): Recipient Inclusions | Col 20 (T): Delivery Method
        sheet.getRange(i + 1, 12).setValue("ALL_ACTIVE_USERS");
        sheet.getRange(i + 1, 18).setValue("Administrator");
        sheet.getRange(i + 1, 20).setValue("Immediate");
        updatedCount++;
      } else if (templateId === "TPL-ADMIN-NEW") {
        // Col 10 (J): Wrapper | Col 12 (L): To Recipients | Col 18 (R): Recipient Inclusions | Col 20 (T): Delivery Method
        sheet.getRange(i + 1, 10).setValue("Internal Communication");
        sheet.getRange(i + 1, 12).setValue("ALL_ACTIVE_USERS");
        sheet.getRange(i + 1, 18).setValue("Administrator");
        sheet.getRange(i + 1, 20).setValue("Immediate");
        updatedCount++;
      }
    }
    
    SpreadsheetApp.flush();
    return "Success: Retroactively reconfigured " + updatedCount + " administrative security templates.";
  } catch(e) {
    return "Migration Error: " + e.message;
  }
}

/**
 * Architectural Upgrade Patch v5.5
 * Retroactively updates the existing Client Welcome Email template body inside your live database.
 * @return {string} Migration outcome status metrics log narrative.
 */
function patch_v5_5_updateClientWelcomeTemplate() {
  try {
    var db = getMainDb();
    var sheet = db.getSheetByName("Templates");
    if (!sheet) return "Migration Aborted: Sheet 'Templates' not found.";

    var data = sheet.getDataRange().getValues();
    var updated = false;

    for (var i = 1; i < data.length; i++) {
      var templateName = data[i][2]; // Column C (Name)
      
      if (templateName === "Client Welcome Email") {
        var updatedBodyHtml = "<div style='font-family: sans-serif; padding: 20px;'><h2>Welcome to {{systemName}}</h2><p>Hi {{priFirstName}},</p><p>We are thrilled to officially partner with <strong>{{brandName}}</strong>.</p><p>Your dedicated workspace for <strong style='color:#666DF2;'>{{services}}</strong> is fully prepared and ready for use.</p></div>";
        
        // Enforce the update cleanly on Column 9 (Column I: Body) without breaking recipient routing layout parameters
        sheet.getRange(i + 1, 9).setValue(updatedBodyHtml);
        updated = true;
        break;
      }
    }

    SpreadsheetApp.flush(); // Mandated race condition safeguard
    
    if (updated) {
      return "Success: Client Welcome Email template text successfully migrated to reader-friendly corporate language.";
    } else {
      return "Notice: Client Welcome Email template record not found in database. No updates applied.";
    }
  } catch(e) {
    return "Migration Failed: " + e.message;
  }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */