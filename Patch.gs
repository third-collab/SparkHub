/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Patch.gs
 * VERSION: 1.1
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
          configObj = parsed; // Preserve existing settings like the clientRole
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
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */