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
  return ["Clients:CREATE", "Clients:UPDATE", "Clients:STATUS_CHANGE", "Clients:ANNOUNCE_NEW"];
}

function Clients_getPlaceholders() { 
  return ["companyName", "brandName", "priFirstName", "priContactFull", "priEmail", "services", "rate", "contractStartDate", "notes", "systemName"];
}

function Clients_getPermissions() { 
  return ["View Clients", "Manage Clients", "Onboard Clients", "Manage Services"];
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
    if (!tplList.some(function(t) { return t.trigger === "Clients:CREATE"; }) && typeof saveTemplateRecord === 'function') {
      saveTemplateRecord({ name: "Client Welcome Email", description: "External welcome to the implementation.", category: "Operations", trigger: "Clients:CREATE", subject: "Welcome to {{systemName}} - {{brandName}}", body: "<div style='font-family: sans-serif; padding: 20px;'><h2>Welcome to {{systemName}}</h2><p>Hi {{priFirstName}},</p><p>We are thrilled to officially partner with <strong>{{brandName}}</strong>.</p><p>Your workspace logic for <strong style='color:#666DF2;'>{{services}}</strong> is provisioned.</p></div>", wrapper: extWrapperName, status: "Active" });
    }
    if (!tplList.some(function(t) { return t.trigger === "Clients:ANNOUNCE_NEW"; }) && typeof saveTemplateRecord === 'function') {
      saveTemplateRecord({ name: "Internal Client Announcement", description: "Alerts staff when a client joins.", category: "Operations", trigger: "Clients:ANNOUNCE_NEW", subject: "New Client: {{brandName}}", body: "<div style='font-family: sans-serif; padding: 20px;'><h2>New Client Onboarded</h2><p>Team,</p><p><strong>{{brandName}}</strong> has joined the network.</p><ul><li><strong>Services:</strong> {{services}}</li><li><strong>Start:</strong> {{contractStartDate}}</li></ul></div>", wrapper: intWrapperName, status: "Active" });
    }
  } catch(e) {}
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
    return { success: true };
  } catch(e) { 
    return { error: e.message };
  }
}

function saveServiceRecord(p) {
  try {
    var sheet = getClientsDb().getSheetByName("Services");
    if (p.rowIndex) { 
      sheet.getRange(parseInt(p.rowIndex,10)+1, 3, 1, 3).setValues([[p.name, p.description, p.status]]); 
      SpreadsheetApp.flush(); // ENFORCED: Prevent Stale Data
      return { success: true };
    } else { 
      sheet.appendRow([new Date(), "SRV-"+Math.floor(1000+Math.random()*9000), p.name, p.description, p.status || "Active"]);
      SpreadsheetApp.flush(); // ENFORCED: Prevent Stale Data
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
    
    var newRow = new Array(26).fill("");
    newRow[0] = new Date(); 
    newRow[1] = "C-" + Math.floor(1000+Math.random()*9000); 
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
    newRow[24] = "[]"; // Empty History
    newRow[25] = "Onboarding";

    sheet.appendRow(newRow); 
    SpreadsheetApp.flush(); // ENFORCED: Prevent Stale Data on Dashboard Read
    
    var targetRow = sheet.getLastRow() - 1;

    try {
      var conf = getClientsModuleConfig().data;
      var targetRole = conf.clientRole || "Client";
      if (typeof saveUserRecord === 'function' && p.pEmail) {
        saveUserRecord({ username: p.pEmail, email: p.pEmail, firstName: p.pFirstName, lastName: p.pLastName, role: targetRole, status: "Active" });
      }

      var sysName = getSystemSettings().systemName || "SparkHub";
      var dataMap = { "companyName": p.companyName, "brandName": brand, "priFirstName": p.pFirstName, "priContactFull": p.pFirstName+" "+p.pLastName, "priEmail": p.pEmail, "services": p.services, "rate": p.rate, "contractStartDate": p.startDate, "systemName": sysName };
      SystemEvent.emit("Clients", "CREATE", "New Client", "INFO", brand, "Client onboarded.", p.pEmail || "no-reply@local", dataMap);
      
      if (typeof getUsersList === 'function') {
        var staff = getUsersList().data.filter(function(u){ return u.status === 'Active' && u.email; });
        if (staff.length > 0) {
          var bccEmails = staff.map(function(u){ return u.email; }).join(',');
          dataMap.bcc = bccEmails;
          SystemEvent.emit("Clients", "ANNOUNCE_NEW", "Announcement", "INFO", brand, brand+" joined.", staff[0].email, dataMap);
        }
      }
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

    const newRow = [...oldRow];
    while(newRow.length < 26) newRow.push("");

    newRow[2]  = p.companyName; 
    newRow[3]  = p.address;
    newRow[4]  = p.companyEmail; 
    newRow[5]  = p.companyPhone;
    newRow[6]  = p.website; 
    newRow[7]  = p.pFirstName; 
    newRow[8]  = p.pLastName; 
    newRow[9]  = p.pEmail;
    newRow[10] = p.pPhone;
    newRow[11] = p.services; 
    newRow[12] = p.rate; 
    newRow[13] = p.termUnit; 
    newRow[14] = p.termCount; 
    newRow[16] = p.currentStartDate || p.startDate;
    newRow[18] = p.currentExpDate || p.expDate;
    
    if (!oldRow[19]) { 
      newRow[19] = p.endDate;
      newRow[20] = p.endDate; 
    } else { 
      newRow[20] = p.endDate;
    }
    
    newRow[21] = p.operationalNotes || oldRow[21]; // Handle incoming op notes
    newRow[22] = p.remarks; 
    newRow[23] = p.addlFields; 
    newRow[25] = p.status;

    // --- ENFORCED: AUDIT DIFFERENTIAL ENGINE ---
    var changes = [];
    var fieldMap = {
      2: "Company Name", 3: "Address", 4: "Company Email", 5: "Company Phone", 6: "Website",
      7: "Primary First Name", 8: "Primary Last Name", 9: "Primary Email", 10: "Primary Phone",
      11: "Services", 12: "Rate", 13: "Term Unit", 14: "Term Count", 25: "Status"
    };

    for (var colIdx in fieldMap) {
      var oldV = String(oldRow[colIdx] || "").trim();
      var newV = String(newRow[colIdx] || "").trim();
      if (oldV !== newV) {
        changes.push({ field: fieldMap[colIdx], old: oldV, new: newV });
      }
    }

    if (changes.length > 0) {
      var historyArr = [];
      try { historyArr = JSON.parse(oldRow[24] || "[]"); } catch(e){}
      
      var editorEmail = "System User";
      try { editorEmail = Session.getActiveUser().getEmail() || "System User"; } catch(e){}

      historyArr.unshift({
        timestamp: new Date().toISOString(),
        editor: editorEmail,
        changes: changes
      });
      newRow[24] = JSON.stringify(historyArr);
    } else {
      newRow[24] = oldRow[24] || "[]"; // Retain existing history if no tracked fields changed
    }

    // Write array to database
    sheet.getRange(rowIndex + 1, 1, 1, newRow.length).setValues([newRow]);
    SpreadsheetApp.flush(); // ENFORCED: Prevent Stale Data on read-only reload

    if (hasStatusChanged) {
      const brand = p.companyName || "Unknown Brand";
      SystemEvent.emit(
        "Clients", 
        "STATUS_CHANGE", 
        "Client Status Updated", 
        "WARN", 
        brand, 
        `Status changed from ${oldStatus} to ${newStatus}.`,
        p.pEmail || "system",
        { oldStatus: oldStatus, newStatus: newStatus }
      );
    }

    return { success: true, message: "Client record updated successfully." };
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

function getSystemDynamicLookups() {
  var lookups = { users: [], templates: [] };
  try {
    var userDb = ensureSheet('W-USERS');
    var uData = userDb.getDataRange().getValues();
    for (var i = 1; i < uData.length; i++) {
      if (uData[i][2] && uData[i][8] !== 'Inactive') {
        lookups.users.push({
           name: uData[i][1] + " " + uData[i][2],
           username: uData[i][3],
           role: uData[i][4]
        });
      }
    }
  } catch(e) {}
  
  try {
    var tplDb = ensureSheet('W-TEMPLATES');
    var tData = tplDb.getDataRange().getValues();
    for (var j = 1; j < tData.length; j++) {
      if (tData[j][1] && tData[j][6] !== 'Inactive') {
         lookups.templates.push({
            name: tData[j][1]
         });
      }
    }
  } catch(e) {}
  
  return lookups;
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */