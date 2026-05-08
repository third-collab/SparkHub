/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Clients.gs
 * VERSION: 1.7 (100% Functional Parity + Standardized Database Handles)
 * SYNC STATUS: Fully Synchronized with Clients.gs.txt Baseline
 */

/**
 * Helper to deep-diff two JSON arrays and return human-readable field-level changes.
 * This correctly tracks exact edits to Phone, Email, Name, Address, etc., 
 * and handles drag/drop reorders silently.
 */
function diffJson(oldJson, newJson, headerName, type) {
  var oldArr = [];
  var newArr = [];
  try { oldArr = JSON.parse(oldJson || "[]"); } catch(e){}
  try { newArr = JSON.parse(newJson || "[]"); } catch(e){}
  
  var changes = [];

  var fieldLabels = {
    firstName: "First Name", lastName: "Last Name", email: "Email", phone: "Phone", 
    position: "Position", birthday: "Birthday", address: "Address",
    contactName: "Contact Name", street: "Street", city: "City", state: "State", 
    countryCode: "Country Code", remarks: "Remarks", content: "Content"
  };

  // Helper to establish the identifying "name" for the audit log subject
  function getItemName(item) {
    if (type === 'contacts') return [item.firstName, item.lastName].filter(Boolean).join(" ") || "Contact";
    if (type === 'addresses') return item.contactName || "Address";
    if (type === 'opnotes') return item.content ? (item.content.substring(0, 30) + "...") : "Note";
    return "Item";
  }

  // Deep comparison ignoring metadata keys (editor, timestamp)
  function isExactMatch(o, n) {
    var allKeys = [];
    Object.keys(o).concat(Object.keys(n)).forEach(function(k) {
      if (k !== 'editor' && k !== 'timestamp' && allKeys.indexOf(k) === -1) allKeys.push(k);
    });
    for (var i = 0; i < allKeys.length; i++) {
      var oVal = String(o[allKeys[i]] || "").trim();
      var nVal = String(n[allKeys[i]] || "").trim();
      if (oVal !== nVal) return false;
    }
    return true;
  }

  var unmatchedOld = [];
  var unmatchedNew = [...newArr];

  // 1. Find exact matches first (This allows drag/drop reorders to be silently ignored)
  for (var i = 0; i < oldArr.length; i++) {
    var o = oldArr[i];
    var matchIdx = -1;
    for (var j = 0; j < unmatchedNew.length; j++) {
      if (isExactMatch(o, unmatchedNew[j])) {
        matchIdx = j;
        break;
      }
    }
    if (matchIdx > -1) {
      unmatchedNew.splice(matchIdx, 1); 
    } else {
      unmatchedOld.push(o);
    }
  }

  // 2. Process remaining unmatched items
  var maxLen = Math.max(unmatchedOld.length, unmatchedNew.length);
  for (var i = 0; i < maxLen; i++) {
    var o = unmatchedOld[i];
    var n = unmatchedNew[i];
    
    // Operational Notes Logic
    if (type === 'opnotes') {
      if (o && !n) changes.push({ field: headerName, old: o.content, new: "Removed" });
      else if (!o && n) changes.push({ field: headerName, old: "", new: "Added: " + n.content });
      else if (o && n && o.content !== n.content) {
         changes.push({ field: headerName + " (Content)", old: o.content, new: n.content });
      }
      continue;
    }

    // Dynamic Context Name (e.g., "Shipping Addresses - Paul Luz")
    var name = n ? getItemName(n) : getItemName(o);
    var itemContext = headerName + " - " + name;

    if (o && !n) {
      changes.push({ field: itemContext, old: "Present", new: "Removed" });
    } else if (!o && n) {
      changes.push({ field: itemContext, old: "None", new: "Added" });
    } else if (o && n) {
      var uProps = [];
      Object.keys(o).concat(Object.keys(n)).forEach(function(p) {
        if(uProps.indexOf(p) === -1) uProps.push(p);
      });
      
      uProps.forEach(function(prop) {
        if (prop === 'editor' || prop === 'timestamp') return; 
        var oVal = String(o[prop] || "").trim();
        var nVal = String(n[prop] || "").trim();
        
        if (oVal !== nVal) {
          var label = fieldLabels[prop] || prop;
          changes.push({ field: itemContext + " (" + label + ")", old: oVal, new: nVal });
        }
      });
    }
  }
  return changes;
}

/**
 * Fetches the summarized list of clients for the main table view.
 * Standardized handle: getMainDb()
 */
function getClientsList() {
  try {
    var sheet = getMainDb().getSheetByName("Clients");
    var data = sheet.getDataRange().getValues();
    var clients = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[1]) continue; 
      
      clients.push({
        rowIndex: i, 
        companyName: row[1],
        brandName: row[2],
        website: row[4],
        pFirstName: row[6],
        pLastName: row[7],
        pEmail: row[8],
        pPhone: row[9],
        services: row[17], 
        status: row[34], 
        acctMgr: row[35] 
      });
    }
    return clients;
  } catch(e) { 
    return [];
  }
}

/**
 * Fetches all details for a specific client by spreadsheet row index.
 */
function getClientById(rowIndex) {
  try {
    var ss = getMainDb();
    var sheet = ss.getSheetByName("Clients");
    if (!sheet) return { error: "Sheet 'Clients' not found." };
    
    var data = sheet.getDataRange().getValues();
    var idx = parseInt(rowIndex, 10);
    if (isNaN(idx)) return { error: "Row index is invalid: " + rowIndex };
    
    var row = data[idx];
    if (!row) return { error: "Row " + idx + " is empty." };

    var tz = ss.getSpreadsheetTimeZone();
    function safeVal(val) {
      if (val instanceof Date) return Utilities.formatDate(val, tz, "yyyy-MM-dd");
      return val === undefined ? "" : val;
    }
    
    return {
      rowIndex: idx,
      companyName: safeVal(row[1]), 
      brandName: safeVal(row[2]), 
      address: safeVal(row[3]), 
      website: safeVal(row[4]), 
      anniversary: safeVal(row[5]),
      pFirstName: safeVal(row[6]), 
      pLastName: safeVal(row[7]), 
      pEmail: safeVal(row[8]), 
      pPhone: safeVal(row[9]), 
      pAddress: safeVal(row[10]), 
      pBirthday: safeVal(row[11]), 
      pPosition: safeVal(row[12]), 
      addContacts: safeVal(row[13]), 
      shipAddress: safeVal(row[14]), 
      retAddress: safeVal(row[15]), 
      brandReg: safeVal(row[16]), 
      services: safeVal(row[17]), 
      monthlyVal: safeVal(row[18]), 
      origStartDate: safeVal(row[19]), 
      currentStartDate: safeVal(row[20]),
      termCount: safeVal(row[21]), 
      termUnit: safeVal(row[22]), 
      origExpDate: safeVal(row[23]),
      currentExpDate: safeVal(row[24]), 
      origEndDate: safeVal(row[25]), 
      latestEndDate: safeVal(row[26]),
      commRate: safeVal(row[27]), 
      commBasis: safeVal(row[28]), 
      ppcDate: safeVal(row[29]),
      dspDate: safeVal(row[30]), 
      handoverNotes: safeVal(row[31]), 
      operationalNotes: safeVal(row[32]), 
      brandCode: safeVal(row[33]), 
      status: safeVal(row[34]) || "Active", 
      acctMgr: safeVal(row[35]),
      brandFolder: safeVal(row[36]), 
      remarks: safeVal(row[37]), 
      history: safeVal(row[38])           
    };
  } catch(e) {
    return { error: "Backend crash: " + e.message };
  }
}

/**
 * Onboards a new client with Drive folder creation and multi-column mapping.
 */
function processNewClient(clientData) {
  try {
    var sheet = getMainDb().getSheetByName("Clients");
    var brand = clientData.brandName || clientData.companyName;

    // DRIVE ENGINE: Dynamic storage allocation
    var assetsFolder = getSystemSubfolder("System Assets");
    var clientRoot = getOrCreateFolder(assetsFolder, "Clients");
    var brandFolder = clientRoot.createFolder(brand);
    brandFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // RESTORED: Explicit Onboarding Field Mapping (39-Columns)
    var newRow = new Array(39).fill("");
    newRow[0] = new Date(); 
    newRow[1] = clientData.companyName;
    newRow[2] = clientData.brandName;
    newRow[3] = clientData.address;
    newRow[4] = clientData.website;
    newRow[5] = ""; // Anniversary
    newRow[6] = clientData.pFirstName;
    newRow[7] = clientData.pLastName;
    newRow[8] = clientData.pEmail;
    newRow[9] = clientData.pPhone;
    newRow[10] = clientData.pAddress;
    newRow[11] = clientData.pBirthday;
    newRow[12] = clientData.pPosition;
    newRow[13] = clientData.addContacts || "[]"; 
    newRow[14] = "[]"; // Initial Shipping
    newRow[15] = "[]"; // Initial Return
    newRow[16] = clientData.brandReg; 
    newRow[17] = clientData.services;
    newRow[18] = clientData.monthlyVal;
    
    newRow[19] = clientData.startDate; // Orig Start
    newRow[20] = clientData.startDate; // Curr Start
    newRow[21] = clientData.termCount; 
    newRow[22] = clientData.termUnit;  
    newRow[23] = clientData.expDate;   // Orig Exp
    newRow[24] = clientData.expDate;   // Curr Exp
    
    newRow[25] = ""; // Orig End
    newRow[26] = ""; // Latest End
    newRow[27] = ""; // Comm Rate
    newRow[28] = ""; // Comm Basis
    newRow[29] = ""; // PPC Date
    newRow[30] = ""; // DSP Date
    newRow[31] = clientData.handoverNotes; 
    newRow[32] = "[]"; // Op Notes
    newRow[33] = ""; // Brand Code
    newRow[34] = "Onboarding"; 
    newRow[35] = "Unassigned"; 
    newRow[36] = brandFolder.getUrl(); // Sync Drive link
    newRow[37] = ""; // Remarks
    newRow[38] = "[]"; // History

    sheet.appendRow(newRow);

    logNotification("Clients", "New Client", "A new client (" + brand + ") was added.", "All", ""); 

    var priContactFull = clientData.pFirstName + " " + clientData.pLastName;
    var clientDataMap = {
      "companyName": clientData.companyName || "",
      "brandName": clientData.brandName || "",
      "priFirstName": clientData.pFirstName || "",
      "priContactFull": priContactFull || "",
      "priEmail": clientData.pEmail || "",
      "services": clientData.services || "",
      "monthlyContractValue": clientData.monthlyVal || "",
      "contractStartDate": clientData.startDate || "",
      "notes": clientData.handoverNotes || ""
    };

    if (clientData.pEmail) sendTriggerEmail("Client Welcome Email", clientData.pEmail, clientDataMap);
    sendTriggerEmail("New Client Announcement", "operations@yourbusiness.com", clientDataMap);

    return "Success! Client onboarded and Drive folder initialized.";
  } catch (error) { 
    return "Error: " + error.toString();
  }
}

/**
 * Updates client record and performs field-level diff for history log.
 */
function updateClientRecord(clientData) {
  try {
    var ss = getMainDb();
    var sheet = ss.getSheetByName("Clients");
    var tz = ss.getSpreadsheetTimeZone();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rowIndex = parseInt(clientData.rowIndex, 10); 
    var oldRow = data[rowIndex]; 
    
    var newRow = [...oldRow]; 
    while(newRow.length < 39) { newRow.push(""); } 

    // RESTORED: Full Logic Parity for Update Payload
    newRow[1] = clientData.companyName;
    newRow[2] = clientData.brandName;
    newRow[3] = clientData.address;
    newRow[4] = clientData.website;
    newRow[5] = clientData.anniversary;
    newRow[6] = clientData.pFirstName;
    newRow[7] = clientData.pLastName;
    newRow[8] = clientData.pEmail;
    newRow[9] = clientData.pPhone;
    newRow[10] = clientData.pAddress;
    newRow[11] = clientData.pBirthday;
    newRow[12] = clientData.pPosition;
    newRow[13] = clientData.addContacts;
    newRow[14] = clientData.shipAddress;
    newRow[15] = clientData.retAddress;
    newRow[16] = clientData.brandReg;
    newRow[17] = clientData.services; 
    newRow[18] = clientData.monthlyVal;
    
    newRow[19] = oldRow[19]; // Explicitly preserve original
    newRow[20] = clientData.currentStartDate; 
    newRow[21] = clientData.termCount;
    newRow[22] = clientData.termUnit;
    newRow[23] = oldRow[23]; // Explicitly preserve original
    newRow[24] = clientData.currentExpDate;   

    // RESTORED: Anniversary and End-Date preservation logic
    if (!oldRow[25] || String(oldRow[25]).trim() === "") {
      newRow[25] = clientData.endDate;
      newRow[26] = clientData.endDate;
    } else {
      newRow[25] = oldRow[25]; 
      newRow[26] = clientData.endDate;
    }

    newRow[27] = clientData.commRate;
    newRow[28] = clientData.commBasis;
    newRow[29] = clientData.ppcDate;
    newRow[30] = clientData.dspDate;
    newRow[31] = clientData.handoverNotes;     
    newRow[32] = clientData.operationalNotes;  
    newRow[33] = clientData.brandCode;         
    newRow[34] = clientData.status;            
    newRow[35] = clientData.acctMgr;           
    newRow[36] = clientData.brandFolder;       
    newRow[37] = clientData.remarks;           

    // RESTORED: Comprehensive Notification Flags
    var notifyContactUpdated = false, notifyContractUpdated = false, notifyShipAddress = false, 
        notifyRetAddress = false, notifyOpNoteAdded = false, notifyOpNoteUpdated = false;
    var contractIndexes = [16, 18, 19, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30]; 
    var jsonColumnsMap = { 13: 'contacts', 14: 'addresses', 15: 'addresses', 32: 'opnotes' };

    var changes = [];
    for (var i = 1; i <= 37; i++) { 
      var oldVal = (oldRow[i] instanceof Date) ? Utilities.formatDate(oldRow[i], tz, "yyyy-MM-dd") : String(oldRow[i] || "").trim();
      var newVal = (newRow[i] instanceof Date) ? Utilities.formatDate(newRow[i], tz, "yyyy-MM-dd") : String(newRow[i] || "").trim();
      
      if (oldVal !== newVal) {
        var headerName = headers[i] || "Field " + i;
        if (jsonColumnsMap[i]) {
          var jChanges = diffJson(oldVal, newVal, headerName, jsonColumnsMap[i]);
          changes = changes.concat(jChanges);
          if (i === 13 && jChanges.length > 0) notifyContactUpdated = true;
          if (i === 14 && jChanges.length > 0) notifyShipAddress = true;
          if (i === 15 && jChanges.length > 0) notifyRetAddress = true;
          if (i === 32 && jChanges.length > 0) {
             jChanges.forEach(function(c) {
                if (c.new && String(c.new).startsWith("Added:")) notifyOpNoteAdded = true;
                else if (c.new !== "Removed") notifyOpNoteUpdated = true;
             });
          }
        } else {
          changes.push({ field: headerName, old: oldVal, new: newVal });
          if (i >= 6 && i <= 12) notifyContactUpdated = true;
          if (contractIndexes.includes(i)) notifyContractUpdated = true;
        }
      }
    }

    if (changes.length > 0) {
      var historyArray = [];
      try { if (oldRow[38]) historyArray = JSON.parse(oldRow[38]); } catch (e) {} 
      historyArray.unshift({ timestamp: new Date().toISOString(), editor: clientData.editor || "Unknown User", changes: changes });
      newRow[38] = JSON.stringify(historyArray);
      newRow[0] = new Date(); 

      // RESTORED: Exhaustive Notification logic from Baseline
      try {
        var brand = clientData.brandName || clientData.companyName;
        var oldStatus = String(oldRow[34]).trim() || "Active";
        var newStatus = String(newRow[34]).trim();
        
        // 1. Status Changes
        if (oldStatus !== newStatus) {
           if (newStatus === "Paused") logNotification("Clients", "Client Paused", brand + " has paused their services.", "All", "");
           else if (newStatus === "Inactive") logNotification("Clients", "Client Offboarded", brand + " is now inactive.", "All", "");
           else if (oldStatus === "Paused" && newStatus === "Active") logNotification("Clients", "Client Resumed", brand + " has resumed services.", "All", "");
           else if (oldStatus === "Inactive" && newStatus === "Active") logNotification("Clients", "Client Rejoined", brand + " has rejoined the active roster.", "All", "");
        }

        // 2. AM Assignment
        if (String(oldRow[35]).trim() !== String(newRow[35]).trim()) {
           logNotification("Clients", "Account Manager Assigned", brand + " is now managed by " + newRow[35], "All", "");
        }

        // 3. Folder Changes
        if (String(oldRow[36]).trim() !== String(newRow[36]).trim()) {
           logNotification("Clients", "Brand Folder Changed", "The brand folder was changed for " + brand + ".", "All", "");
        }

        // 4. Aggregated Baseline Updates
        if (notifyOpNoteAdded) logNotification("Clients", "New Operational Note", "A new operational note was added for " + brand + ".", "All", "");
        if (notifyOpNoteUpdated) logNotification("Clients", "Operational Note Updated", "An operational note was updated for " + brand + ".", "All", "");
        if (notifyShipAddress) logNotification("Clients", "Shipping Address Updated", "A shipping address was added or updated for " + brand + ".", "All", "");
        if (notifyRetAddress) logNotification("Clients", "Return Address Updated", "A return address was added or updated for " + brand + ".", "All", "");
        if (notifyContactUpdated) logNotification("Clients", "Contact Info Updated", "A client's contact information was updated (" + brand + ").", "All", "");
        if (notifyContractUpdated) logNotification("Clients", "Contract Info Updated", "Contract information was updated for " + brand + ".", "All", "");

      } catch(e) { console.error("Notification sync error: " + e.message); }
    }

    sheet.getRange(rowIndex + 1, 1, 1, newRow.length).setValues([newRow]);
    return "Success! Master record synchronized.";
  } catch(e) {
    return "Error updating database: " + e.message;
  }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */