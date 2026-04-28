/**
 * Users Module - Backend
 * Standardized under SparkHub Architecture Blueprint.
 * Handles: Staff profiles, permissions, RBAC, and directory management.
 */

function processNewUser(obj) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    sheet.appendRow([
      new Date(), obj.username, obj.role, obj.email, obj.password, 
      obj.firstName, obj.lastName, obj.status, "" 
    ]);
    logSystemAction("Users", "CREATE", "Add User", "INFO", obj.username, "New user access profile created.");
    return "Success! User created.";
  } catch (e) { return "Error: " + e.message; }
}

function updateUserRecord(obj) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var row = parseInt(obj.rowIndex);
    // Write 7 columns starting at Column 2 (Username)
    sheet.getRange(row, 2, 1, 7).setValues([[
      obj.username, obj.role, obj.email, obj.password, obj.firstName, 
      obj.lastName, obj.status
    ]]);
    logSystemAction("Users", "UPDATE", "Edit User", "INFO", obj.username, "User access profile updated.");
    return "Success! User updated.";
  } catch (e) { return "Error: " + e.message; }
}

function getUserById(rowIndex) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    // Fetch 9 columns
    var rowData = sheet.getRange(rowIndex, 1, 1, 9).getDisplayValues()[0];
    return {
      rowIndex: rowIndex, username: rowData[1], role: rowData[2], email: rowData[3], password: rowData[4],
      firstName: rowData[5], lastName: rowData[6], status: rowData[7], lastLogin: rowData[8]
    };
  } catch (e) { return { error: e.message }; }
}

function getUsersList() {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    data.shift();
    return data.map(function(row, i) {
      return { 
        rowIndex: i + 2, 
        username: row[1], role: row[2], email: row[3], 
        firstName: row[5], lastName: row[6], status: row[7] 
      };
    });
  } catch (e) { return []; }
}

/* ========================================================================
   ROLE-BASED ACCESS CONTROL (RBAC) INTENTS
   ======================================================================== */

function getUserRole() {
  try {
    var email = Session.getActiveUser().getEmail();
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === email) { // Email is index 3
        return data[i][7] === 'Inactive' ? 'Inactive' : data[i][2]; // Status is index 7
      }
    }
    
    if (email === PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')) {
      return "Administrator";
    }
    return "Inactive";
  } catch (e) { return "Inactive"; }
}

function getLoggedInUsername() {
  try {
    var email = Session.getActiveUser().getEmail();
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === email) return data[i][1];
    }
    return email.split('@')[0];
  } catch (e) { return "User"; }
}

function updateLastLogin() {
  try {
    var email = Session.getActiveUser().getEmail();
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === email) { 
        sheet.getRange(i + 1, 9).setValue(new Date()); // Column 9 is Last Login
        break;
      }
    }
  } catch (e) { console.error("Failed to update last login: " + e.message); }
}