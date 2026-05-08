/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Patch.gs
 * VERSION: 1.2
 * SYNC STATUS: Standalone Utility
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
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */