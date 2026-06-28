/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Clients.gs
 * VERSION: 5.0 (Audit Engine & Flush Restoration)
 * SYNC STATUS: Fully Synchronized with ClientsData.html
 */

/**
 * ========================================================================
 * REGISTRY EXPORTS (V8 Auto-Discovery)
 * ========================================================================
 */

function Clients_getTriggers() { 
  return [
    "Clients:CREATE", "Clients:UPDATE", "Clients:STATUS_CHANGE",
    "Clients:Services:CREATE", "Clients:Services:UPDATE"
  ];
}

function Clients_getPlaceholders() { 
  var placeholders = ["companyName", "brandName", "priFirstName", "priContactFull", "priEmail", "services", "rate", "contractStartDate", "notes", "systemName", "core_assetFolderLink"];
  try {
    if (typeof getClientsModuleConfig === 'function') {
      var confRes = getClientsModuleConfig();
      if (confRes && confRes.success && confRes.data && confRes.data.customFields) {
        confRes.data.customFields.forEach(function(field) {
          if (field.id && placeholders.indexOf(field.id) === -1) {
            placeholders.push(field.id);
          }
        });
      }
    }
  } catch(e) {
    console.warn("Failed to harvest custom fields for template placeholder registry: " + e.message);
  }
  return placeholders;
}

function Clients_getPermissions() { 
  return ["View Clients", "Onboard Clients", "Manage Clients", "View Services", "Define Services", "Manage Services", "Manage Settings"];
}

/**
 * ========================================================================
 * CORE PROCESSORS (Database & Config)
 * ========================================================================
 */

function setupClientsDatabase() {
  try {
    var settings = getSystemSettings();
    if (!settings.rootFolderId) return { error: "Root folder not found." };
    
    var folder = DriveApp.getFolderById(settings.rootFolderId);
    var sysName = PropertiesService.getScriptProperties().getProperty('SYSTEM_NAME') || "SparkHub";
    var dbName = sysName + " Clients Database";
    
    var files = folder.getFilesByName(dbName);
    var dbId = ""; 
    var ss;
    
    if (files.hasNext()) { 
      ss = SpreadsheetApp.openById(files.next().getId()); 
      dbId = ss.getId();
    } else {
      ss = SpreadsheetApp.create(dbName);
      var file = DriveApp.getFileById(ss.getId()); 
      file.moveTo(folder); 
      dbId = file.getId();
      
      var sheet = ss.getSheets()[0]; 
      sheet.setName("Clients");
      var headers = [
        "Timestamp", "Client ID", "Company Name", "Address", "Company Email", "Company Phone", "Website", 
        "Primary First Name", "Primary Last Name", "Primary Email", "Primary Phone", "Services", "Rate", 
        "Term Unit", "Term Count", "Original Start Date", "Current Start Date", "Original Exp Date", 
        "Current Exp Date", "Original End Date", "Latest End Date", "Operational Notes", "Remarks", 
        "Additional Fields", "History", "Status"
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f1f5f9");
      sheet.setFrozenRows(1);
    }

    var srvSheet = ss.getSheetByName("Services");
    if (!srvSheet) {
       srvSheet = ss.insertSheet("Services");
       srvSheet.getRange(1, 1, 1, 5).setValues([["Timestamp", "Service ID", "Service Name", "Description", "Status"]]).setFontWeight("bold").setBackground("#f1f5f9");
       srvSheet.appendRow([new Date(), "SRV-1001", "Primary services", "Core operational capabilities.", "Active"]);
       srvSheet.appendRow([new Date(), "SRV-1002", "Secondary services", "Ad-hoc, out-of-scope retainers.", "Active"]);
       srvSheet.setFrozenRows(1);
    }
    
    PropertiesService.getScriptProperties().setProperty('CLIENTS_DB_ID', dbId);
    
    var defaultRoleName = "Client";
    try { 
      if (typeof saveRoleRecord === 'function') {
        saveRoleRecord({ name: defaultRoleName, description: "Default restricted access role for clients.", permissions: "[]", dashConfig: "[]", status: "Active" });
      }
    } catch(e) {}

    var confObj = {
       clientRole: defaultRoleName,
       accordionGroups: [
           { id: "grp_client", label: "Client Details", isDefault: true },
           { id: "grp_contact", label: "Primary Contact", isDefault: true },
           { id: "grp_contract", label: "Contract Information", isDefault: true }
       ],
       customFields: [] 
    };
    PropertiesService.getScriptProperties().setProperty('CLIENTS_MODULE_CONFIG', JSON.stringify(confObj));

    return { success: true, message: "Clients database initialized." };
  } catch (e) { 
    return { error: "Database setup error: " + e.message };
  }
}

function seedClientsTemplates(extWrapperName, intWrapperName) {
  try {
    var tplList = typeof getTemplatesList === 'function' ? getTemplatesList() : [];
    
    // 1. Seed External Welcome Template for the incoming client target
    if (!tplList.some(function(t) { return t.name === "Client Welcome Email"; }) && typeof saveTemplateRecord === 'function') {
      saveTemplateRecord({ name: "Client Welcome Email", description: "External welcome to the implementation.", category: "Operations", trigger: "Clients:CREATE", subject: "Welcome to {{systemName}} - {{brandName}}", body: "<div style='font-family: sans-serif; padding: 20px;'><h2>Welcome to {{systemName}}</h2><p>Hi {{priFirstName}},</p><p>We are thrilled to officially partner with <strong>{{brandName}}</strong>.</p><p>Your workspace logic for <strong style='color:#666DF2;'>{{services}}</strong> is provisioned.</p></div>", wrapper: extWrapperName, status: "Active", to: "TRIGGER_DEFAULT", dispatchMode: "Individual" });
    }
    
    // 2. Seed Internal Notification Template targeted at staff distribution cohorts using the 15-column schema matrix
    if (!tplList.some(function(t) { return t.name === "Internal Client Announcement"; }) && typeof saveTemplateRecord === 'function') {
      saveTemplateRecord({ name: "Internal Client Announcement", description: "Alerts staff when a client joins.", category: "Operations", trigger: "Clients:CREATE", subject: "New Client Onboarded: {{brandName}}", body: "<div style='font-family: sans-serif; padding: 20px;'><h2>New Client Onboarded</h2><p>Team,</p><p><strong>{{brandName}}</strong> has joined the network.</p><ul><li><strong>Services:</strong> {{services}}</li><li><strong>Start:</strong> {{contractStartDate}}</li></ul></div>", wrapper: intWrapperName, status: "Active", to: "ALL_ACTIVE_USERS", dispatchMode: "Collective" });
    }
  } catch(e) { console.error("Error seeding unified client templates: " + e.message); }
}

/**
 * ========================================================================
 * READ / GET FUNCTIONS
 * ========================================================================
 */

function getClientsList() {
  try {
    var sheet = getClientsSheet();
    var data = sheet.getDataRange().getValues(); 
    var clients = [];
    
    for (var i = 1; i < data.length; i++) {
      if (!data[i][1]) continue;
      
      var addl = {}; 
      try { 
        addl = JSON.parse(data[i][23] || "{}");
      } catch(e){}
      
      var brand = addl.brand_name || data[i][2] || "Unknown Brand";
      var am = addl.account_manager || "Unassigned";
      
      clients.push({ 
        rowIndex: i, 
        clientId: data[i][1], 
        companyName: data[i][2], 
        brandName: brand, 
        pFirstName: data[i][7], 
        pLastName: data[i][8], 
        pEmail: data[i][9], 
        services: data[i][11], 
        acctMgr: am, 
        status: data[i][25] 
      });
    }
    return { success: true, data: clients };
  } catch(e) { 
    return { error: "List error: " + e.message };
  }
}

function getClientById(rowIndex) {
  try {
    var sheet = getClientsSheet(); 
    var tz = getClientsDb().getSpreadsheetTimeZone();
    var row = sheet.getDataRange().getValues()[parseInt(rowIndex, 10)];
    
    if (!row) return { error: "Record empty." };
    
    function safeVal(val) { 
      return (val instanceof Date) ? Utilities.formatDate(val, tz, "yyyy-MM-dd") : (val === undefined ? "" : val);
    }
    
    return { success: true, data: {
      rowIndex: rowIndex, 
      clientId: safeVal(row[1]), 
      companyName: safeVal(row[2]), 
      address: safeVal(row[3]), 
      companyEmail: safeVal(row[4]), 
      companyPhone: safeVal(row[5]), 
      website: safeVal(row[6]), 
      pFirstName: safeVal(row[7]), 
      pLastName: safeVal(row[8]), 
      pEmail: safeVal(row[9]), 
      pPhone: safeVal(row[10]), 
      services: safeVal(row[11]), 
      rate: safeVal(row[12]), 
      termUnit: safeVal(row[13]), 
      termCount: safeVal(row[14]), 
      origStartDate: safeVal(row[15]), 
      currentStartDate: safeVal(row[16]), 
      origExpDate: safeVal(row[17]), 
      currentExpDate: safeVal(row[18]), 
      origEndDate: safeVal(row[19]), 
      latestEndDate: safeVal(row[20]),
      operationalNotes: safeVal(row[21]), 
      remarks: safeVal(row[22]), 
      addlFields: safeVal(row[23]), 
      history: safeVal(row[24]), 
      status: safeVal(row[25]) || "Active"
    }};
  } catch(e) { 
    return { error: "Backend crash: " + e.message };
  }
}

function getServicesList() {
  try {
    var sheet = getClientsDb().getSheetByName("Services");
    if (!sheet) return { success: true, data: [] };
    
    var data = sheet.getDataRange().getValues(); 
    var srvs = [];
    
    for (var i = 1; i < data.length; i++) { 
      if (!data[i][1]) continue;
      srvs.push({ 
        rowIndex: i, 
        serviceId: data[i][1], 
        name: data[i][2], 
        description: data[i][3], 
        status: data[i][4] 
      });
    }
    return { success: true, data: srvs };
  } catch(e) { 
    return { error: e.message };
  }
}

function getClientsModuleConfig() {
  try {
    var conf = PropertiesService.getScriptProperties().getProperty('CLIENTS_MODULE_CONFIG');
    if (!conf) return { success: true, data: { clientRole: "Client", accordionGroups: [], customFields: [] } };
    return { success: true, data: JSON.parse(conf) };
  } catch(e) { 
    return { error: e.message };
  }
}

/**
 * ========================================================================
 * WRITE / SAVE FUNCTIONS
 * ========================================================================
 */

function saveClientsModuleConfig(p) {
  try { 
    PropertiesService.getScriptProperties().setProperty('CLIENTS_MODULE_CONFIG', JSON.stringify(p));
    // Emit active trace records straight to the central 'System' workspace block registry
    try {
      SystemEvent.emit("System", "UPDATE", "Config Updated", "INFO", "Clients", "Client module settings updated locally.");
    } catch(logErr) {}
    return { success: true };
  } catch(e) {
    return { error: e.message };
  }
}

function saveServiceRecord(p) {
  try {
    var sheet = getClientsDb().getSheetByName("Services");
    if (p.rowIndex) { 
      var oldData = sheet.getRange(parseInt(p.rowIndex,10)+1, 1, 1, 5).getValues()[0];
      var srvId = oldData[1];
      sheet.getRange(parseInt(p.rowIndex,10)+1, 3, 1, 3).setValues([[p.name, p.description, p.status]]); 
      SpreadsheetApp.flush();
      
      try {
        SystemEvent.emit("Clients:Services", "UPDATE", "Service Updated", "INFO", srvId, "Service definition details modified.", "system", { serviceName: p.name, description: p.description, status: p.status, targetName: p.name });
      } catch(logErr) { console.warn("Service log failed: " + logErr.message); }
      
      return { success: true };
    } else { 
      var srvId = "SRV-"+Math.floor(1000+Math.random()*9000);
      sheet.appendRow([new Date(), srvId, p.name, p.description, p.status || "Active"]);
      SpreadsheetApp.flush();
      
      try {
        SystemEvent.emit("Clients:Services", "CREATE", "Service Created", "INFO", srvId, "New service capability defined.", "system", { serviceName: p.name, description: p.description, status: p.status, targetName: p.name });
      } catch(logErr) { console.warn("Service log failed: " + logErr.message); }
      
      return { success: true };
    }
  } catch(e) {
    return { error: e.message };
  }
}

function createClientRecord(p) {
  try {
    var sheet = getClientsSheet();
    var addl = {};
    try { addl = JSON.parse(p.addlFields || "{}"); } catch(e){}
    var brand = addl.brand_name || p.companyName;
    
    var generatedId = "C-" + Math.floor(1000+Math.random()*9000);
    
    // Evaluate Cloud Storage Automation settings flags before provisioning assets
    try {
      var configRes = typeof getClientsModuleConfig === 'function' ? getClientsModuleConfig() : null;
      var conf = (configRes && configRes.success) ? configRes.data : {};
      
      // DEFENSIVE SAFEGUARD: Only run automation if enabled AND user left the field completely blank during onboarding
      if (conf && conf.autoCreateFolder && (!addl.core_assetFolderLink || addl.core_assetFolderLink.trim() === "")) {
        var parentFolder;
        // Determine destination tree path mapping bounds
        if (conf.parentFolderId && conf.parentFolderId.trim() !== "") {
          parentFolder = DriveApp.getFolderById(conf.parentFolderId.trim());
        } else {
          parentFolder = getSystemSubfolder("Client Assets");
        }
        
        var namingBlueprint = conf.folderTemplate && conf.folderTemplate.trim() !== "" ? conf.folderTemplate.trim() : "[{{clientId}}] {{companyName}}";
        var resolvedFolderName = namingBlueprint.replace(/\{\{clientId\}\}/g, generatedId).replace(/\{\{companyName\}\}/g, p.companyName);
        
        var clientFolder = parentFolder.createFolder(resolvedFolderName);
        clientFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        // Maps programmatically provisioned cloud URL into metadata fields sink slots
        addl.core_assetFolderLink = clientFolder.getUrl();
        p.addlFields = JSON.stringify(addl);
      }
    } catch(folderProvisioningErr) {
      console.warn("Automated asset folder generation bypassed: " + folderProvisioningErr.message);
    }

    var newRow = new Array(26).fill("");
    newRow[0] = new Date(); 
    newRow[1] = generatedId; 
    newRow[2] = p.companyName;
    newRow[3] = p.address;
    newRow[4] = p.companyEmail; 
    newRow[5] = p.companyPhone; 
    newRow[6] = p.website;
    newRow[7] = p.pFirstName; 
    newRow[8] = p.pLastName; 
    newRow[9] = p.pEmail;
    newRow[10] = p.pPhone;
    newRow[11] = p.services; 
    newRow[12] = p.rate; 
    newRow[13] = p.termUnit; 
    newRow[14] = p.termCount;
    newRow[15] = p.startDate;
    newRow[16] = p.startDate; 
    newRow[17] = p.expDate; 
    newRow[18] = p.expDate; 
    newRow[21] = "[]"; 
    newRow[23] = p.addlFields || "{}"; 
    newRow[24] = "[]";
    // Empty History
    newRow[25] = "Onboarding";

    sheet.appendRow(newRow); 
    SpreadsheetApp.flush();
    
    var targetRow = sheet.getLastRow() - 1;

    try {
      var conf = getClientsModuleConfig().data;
      var targetRole = conf.clientRole || "Client";
      // Enforces connection to the authentic createUserRecord token provisioning sequence defined in Users.gs
      if (typeof createUserRecord === 'function' && p.pEmail) {
        createUserRecord({ username: p.pEmail, email: p.pEmail, firstName: p.pFirstName, lastName: p.pLastName, role: targetRole, status: "Active" });
      }

      var sysName = getSystemSettings().systemName || "SparkHub";
      var dataMap = { "companyName": p.companyName, "brandName": brand, "priFirstName": p.pFirstName, "priContactFull": p.pFirstName+" "+p.pLastName, "priEmail": p.pEmail, "services": p.services, "rate": p.rate, "contractStartDate": p.startDate, "systemName": sysName };
      
      // Unpack all custom fields and assets folder links stored in the addl metadata payload straight into dataMap tokens
      if (addl && typeof addl === 'object') {
        for (var key in addl) {
          if (addl.hasOwnProperty(key)) {
            dataMap[key] = addl[key];
          }
        }
      }
      
      // Issues a single atomic transaction signal. Core lookup matrices handle internal/external routing automatically
      SystemEvent.emit("Clients", "CREATE", "New Client", "INFO", brand, "Client onboarded.", p.pEmail || "no-reply@local", dataMap);
    } catch(postError) {
      console.warn("Client saved, but post-save actions failed: " + postError.message);
    }
    
    return { success: true, rowIndex: targetRow };
  } catch (e) { 
    return { error: "Backend error: " + e.message };
  }
}

function updateClientRecord(p) {
  try {
    const sheet = getClientsSheet();
    const data = sheet.getDataRange().getValues();
    const rowIndex = parseInt(p.rowIndex, 10);
    const oldRow = data[rowIndex];
    
    if (!oldRow) throw new Error("Could not find record at row " + rowIndex);

    const oldStatus = oldRow[25];
    const newStatus = p.status;
    const hasStatusChanged = (oldStatus !== newStatus);

    // [REPLACE EXPLICITLY: Lines 75-81 in Clients.gs]
    const newRow = [...oldRow];
    while(newRow.length < 26) newRow.push("");

    // Defensive mapping layer safeguards array elements against accidental 'undefined' properties
    newRow[2]  = p.companyName !== undefined ? p.companyName : (oldRow[2] || "");
    newRow[3]  = p.address !== undefined ? p.address : (oldRow[3] || "");
    newRow[4]  = p.companyEmail !== undefined ? p.companyEmail : (oldRow[4] || "");
    newRow[5]  = p.companyPhone !== undefined ? p.companyPhone : (oldRow[5] || "");
    newRow[6]  = p.website !== undefined ? p.website : (oldRow[6] || "");
    newRow[7]  = p.pFirstName !== undefined ? p.pFirstName : (oldRow[7] || "");
    newRow[8]  = p.pLastName !== undefined ? p.pLastName : (oldRow[8] || "");
    newRow[9]  = p.pEmail !== undefined ? p.pEmail : (oldRow[9] || "");
    newRow[10] = p.pPhone !== undefined ? p.pPhone : (oldRow[10] || "");
    newRow[11] = p.services !== undefined ? p.services : (oldRow[11] || "");
    newRow[12] = p.rate !== undefined ? p.rate : (oldRow[12] || "");
    newRow[13] = p.termUnit !== undefined ? p.termUnit : (oldRow[13] || "");
    newRow[14] = p.termCount !== undefined ? p.termCount : (oldRow[14] || "");
    
    newRow[16] = p.currentStartDate || p.startDate || (oldRow[16] ? Utilities.formatDate(new Date(oldRow[16]), getClientsDb().getSpreadsheetTimeZone(), "yyyy-MM-dd") : "");
    newRow[18] = p.currentExpDate || p.expDate || (oldRow[18] ? Utilities.formatDate(new Date(oldRow[18]), getClientsDb().getSpreadsheetTimeZone(), "yyyy-MM-dd") : "");
    
    var inputEndDate = p.endDate !== undefined ? p.endDate : (p.latestEndDate !== undefined ? p.latestEndDate : (oldRow[20] ? Utilities.formatDate(new Date(oldRow[20]), getClientsDb().getSpreadsheetTimeZone(), "yyyy-MM-dd") : ""));
    if (!oldRow[19] && inputEndDate) {
      newRow[19] = inputEndDate;
      newRow[20] = inputEndDate;
    } else {
      if (p.origEndDate !== undefined) newRow[19] = p.origEndDate;
      newRow[20] = inputEndDate || "";
    }
    
    newRow[21] = p.operationalNotes || oldRow[21] || "[]";
    newRow[22] = p.remarks !== undefined ? p.remarks : (oldRow[22] || "");
    newRow[23] = p.addlFields || oldRow[23] || "{}";
    newRow[25] = p.status || oldRow[25] || "Active";

    // ========================================================================
    // BACKEND AUDIT NARRATIVE NATIVE LEXICON DISPATCH ENGINE
    // ========================================================================
    var changes = [];
    var fieldMap = {
      2: "Company Name", 3: "Address", 4: "Company Email", 5: "Company Phone", 6: "Website",
      7: "First Name", 8: "Last Name", 9: "Email", 10: "Phone",
      11: "Services Included", 12: "Rate", 13: "Term Unit", 14: "Term Count",
      16: "Start Date", 18: "Expiration Date", 20: "End Date", 25: "Status"
    };
    var editorUsername = "U-SYSTEM";
    try { 
      if (typeof getLoggedInUserId === 'function') {
        editorUsername = getLoggedInUserId(); // History records use the unchangeable tracking ID anchor
      }
    } catch(e){}

    for (var colIdx in fieldMap) {
      var oldV = String(oldRow[colIdx] || "").trim();
      var newV = String(newRow[colIdx] || "").trim();
      if (oldV !== newV) {
        changes.push({
          field: fieldMap[colIdx],
          type: "standard",
          action: (oldV === "") ? "Added" : "Edited",
          oldVal: oldV,
          newVal: newV
        });
      }
    }

    // High-Fidelity Differential Logic for Operational Note Collections
    var oldNotes = [], newNotes = [];
    try { oldNotes = JSON.parse(oldRow[21] || "[]"); } catch(e){}
    try { newNotes = JSON.parse(newRow[21] || "[]"); } catch(e){}
    
    if (JSON.stringify(oldNotes) !== JSON.stringify(newNotes)) {
      if (newNotes.length > oldNotes.length) {
        var addedNote = newNotes[newNotes.length - 1];
        changes.push({
          field: "Operational Notes",
          type: "note",
          action: "Added",
          oldVal: "",
          newVal: addedNote.content
        });
      } else if (newNotes.length < oldNotes.length) {
        var deletedText = "";
        for (var oldIdx = 0; oldIdx < oldNotes.length; oldIdx++) {
          if (!newNotes.some(n => n.timestamp === oldNotes[oldIdx].timestamp)) {
            deletedText = oldNotes[oldIdx].content;
            break;
          }
        }
        changes.push({
          field: "Operational Notes",
          type: "note",
          action: "Deleted",
          oldVal: deletedText,
          newVal: ""
        });
      } else {
        for (var noteIdx = 0; noteIdx < newNotes.length; noteIdx++) {
          if (oldNotes[noteIdx] && oldNotes[noteIdx].content !== newNotes[noteIdx].content) {
            changes.push({
              field: "Operational Notes",
              type: "note",
              action: "Edited",
              oldVal: oldNotes[noteIdx].content,
              newVal: newNotes[noteIdx].content
            });
            break;
          }
        }
      }
    }

    // Interrogate Custom Meta Metadata Fields and Array Sub-Fields
    var oldAddl = {}, newAddl = {};
    try { oldAddl = JSON.parse(oldRow[23] || "{}"); } catch(e){}
    try { newAddl = JSON.parse(p.addlFields || "{}"); } catch(e){}
    
    var customFieldsObj = [];
    try {
      var conf = PropertiesService.getScriptProperties().getProperty('CLIENTS_MODULE_CONFIG');
      if (conf) customFieldsObj = JSON.parse(conf).customFields || [];
    } catch(e){}

    var allKeys = new Set([...Object.keys(oldAddl), ...Object.keys(newAddl)]);
    allKeys.forEach(function(k) {
      var oldVal = oldAddl[k] !== undefined ? String(oldAddl[k]).trim() : "";
      var newVal = newAddl[k] !== undefined ? String(newAddl[k]).trim() : "";
      if (oldVal !== newVal) {
        var cfDef = customFieldsObj.find(f => f.id === k);
        var label = cfDef ? cfDef.label : k;
        var isRichText = cfDef && cfDef.type === 'richtext';
        var isJson = cfDef && cfDef.type === 'json';
        
        if (isRichText) {
          var cleanOld = oldVal.replace(/<[^>]*>/g, "").trim();
          var cleanNew = newVal.replace(/<[^>]*>/g, "").trim();
          
          if (cleanOld.length > 200) cleanOld = cleanOld.substring(0, 100) + "..." + cleanOld.substring(cleanOld.length - 100);
          if (cleanNew.length > 200) cleanNew = cleanNew.substring(0, 100) + "..." + cleanNew.substring(cleanNew.length - 100);
          
          changes.push({
            field: label,
            type: "richtext",
            action: (cleanOld === "") ? "Added" : "Edited",
            oldVal: cleanOld,
            newVal: cleanNew
          });
        } else if (isJson) {
          var pOld = [], pNew = [];
          try { pOld = JSON.parse(oldVal || "[]"); } catch(e){}
          try { pNew = JSON.parse(newVal || "[]"); } catch(e){}
          var subChanges = [];
          
          pNew.forEach(function(item, idx) {
            for (var subK in item) {
              var oS = (pOld[idx] && pOld[idx][subK] !== undefined) ? String(pOld[idx][subK]).trim() : "";
              var nS = String(item[subK] !== undefined ? item[subK] : "").trim();
              if (oS !== nS) {
                subChanges.push({ subField: subK, old: oS, new: nS });
              }
            }
          });
          
          if (subChanges.length > 0) {
            changes.push({
              field: label,
              type: "json",
              subChanges: subChanges
            });
          }
        } else {
          changes.push({
            field: label,
            type: "standard",
            action: (oldVal === "") ? "Added" : "Edited",
            oldVal: oldVal,
            newVal: newVal
          });
        }
      }
    });

    if (changes.length > 0) {
      var historyArr = [];
      try { historyArr = JSON.parse(oldRow[24] || "[]"); } catch(e){}
      historyArr.unshift({ timestamp: new Date().toISOString(), editor: editorUsername, changes: changes });
      newRow[24] = JSON.stringify(historyArr);
    } else {
      newRow[24] = oldRow[24] || "[]";
    }

    sheet.getRange(rowIndex + 1, 1, 1, newRow.length).setValues([newRow]);
    SpreadsheetApp.flush();

    // MANDATE: Generate global system-wide tracking logs for every profile or note modification
    try {
      var clientId = oldRow[1];
      var brandName = p.companyName || "Unknown Brand";
      var logAction = "Client Updated";
      var logDetails = "Client profile details updated.";
      // Interrogate calculated changes array to produce high-fidelity system logs
      var noteChange = changes.find(function(c) { return c.type === "note"; });
      if (noteChange) {
        logAction = "Client Note Modified";
        // References newVal property correctly to avoid throwing fatal unhandled undefined property TypeError breaks
        logDetails = noteChange.newVal.replace(/<[^>]*>/g, "");
      } else if (changes.length > 0) {
        logDetails = "Updated fields: " + changes.map(function(c) { return c.field; }).join(", ");
      }
      
      SystemEvent.emit("Clients", "UPDATE", logAction, "INFO", clientId, logDetails, p.pEmail || "system", { targetName: brandName });
    } catch(logErr) {
      console.warn("Global system log emission failed for client update: " + logErr.message);
    }

    if (hasStatusChanged) {
      var clientId = oldRow[1];
      const brand = p.companyName || "Unknown Brand";
      SystemEvent.emit("Clients", "STATUS_CHANGE", "Client Status Updated", "WARN", clientId, `Status changed from ${oldStatus} to ${newStatus}.`, p.pEmail || "system", { oldStatus: oldStatus, newStatus: newStatus, targetName: brand });
    }

    return { success: true, message: "Updated successfully.", data: { operationalNotes: newRow[21], history: newRow[24] } };
  } catch (e) {
    console.error("updateClientRecord error: " + e.message);
    return { error: "Update failed: " + e.message };
  }
}

/**
 * ========================================================================
 * INTERNAL HELPERS
 * ========================================================================
 */

function getClientsDb() { 
  var id = PropertiesService.getScriptProperties().getProperty('CLIENTS_DB_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.openById(setupClientsDatabase().dbId); 
}

function getClientsSheet() { 
  var ss = getClientsDb(); 
  var s = ss.getSheetByName("Clients");
  return s ? s : ss.insertSheet("Clients"); 
}

function Clients_getLookups() {
  var vectors = { clients: [], services: [] };
  try {
    var sheet = getClientsSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][25] !== 'Inactive') {
        var addl = {};
        try { addl = JSON.parse(data[i][23] || "{}"); } catch(e){}
        var brand = addl.brand_name || data[i][2] || "Unknown Brand";
        vectors.clients.push({
          id: String(data[i][1]), 
          name: String(brand)
        });
      }
    }
  } catch(e) { console.warn("Clients lookup broadcast failed: " + e.message); }
  
  try {
    var srvSheet = getClientsDb().getSheetByName("Services");
    if (srvSheet) {
      var sData = srvSheet.getDataRange().getValues();
      for (var s = 1; s < sData.length; s++) {
        if (sData[s][1] && sData[s][4] === 'Active') {
          vectors.services.push({
            id: String(sData[s][2]), 
            name: String(sData[s][2])
          });
        }
      }
    }
  } catch(e) { console.warn("Services lookup broadcast failed: " + e.message); }
  return vectors;
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */