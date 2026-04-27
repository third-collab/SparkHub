/**
 * Users Module - Backend
 * Standardized under SparkHub Architecture Blueprint.
 * Handles: Staff profiles, permissions, RBAC, and directory management.
 */

/**
 * Processes and saves a new user record in the 'Users' sheet.
 * @param {Object} obj - The user data object from the Add User form.
 * @return {string} Success confirmation message.
 */
function processNewUser(obj) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    sheet.appendRow([
      new Date(), 
      obj.username, 
      obj.role, 
      obj.workEmail, 
      obj.firstName, 
      obj.lastName, 
      obj.birthday, 
      obj.personalEmail, 
      obj.phone, 
      obj.address, 
      obj.facebookUrl, 
      obj.profilePhotoUrl, 
      obj.position, 
      obj.employmentType, 
      obj.dateHired, 
      obj.status 
    ]);
    logNotification("Users", "New User", "Team Member Joined: " + obj.username, "All", "");
    return "Success! User created.";
  } catch (e) { return "Error: " + e.message; }
}

/**
 * Updates an existing user record in the spreadsheet.
 * @param {Object} obj - Data object including the spreadsheet rowIndex.
 * @return {string} Success confirmation message.
 */
function updateUserRecord(obj) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var row = parseInt(obj.rowIndex);
    sheet.getRange(row, 2, 1, 15).setValues([[
      obj.username, 
      obj.role, 
      obj.workEmail, 
      obj.firstName, 
      obj.lastName, 
      obj.birthday, 
      obj.personalEmail, 
      obj.phone, 
      obj.address, 
      obj.facebookUrl, 
      obj.profilePhotoUrl, 
      obj.position, 
      obj.employmentType, 
      obj.dateHired, 
      obj.status
    ]]);
    return "Success! User updated.";
  } catch (e) { return "Error: " + e.message; }
}

/**
 * Retrieves full details for a specific user.
 * @param {number} rowIndex - The spreadsheet row index.
 * @return {Object} Detailed user data or error object.
 */
function getUserById(rowIndex) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var rowData = sheet.getRange(rowIndex, 1, 1, 16).getDisplayValues()[0];
    return {
      rowIndex: rowIndex, username: rowData[1], role: rowData[2], workEmail: rowData[3],
      firstName: rowData[4], lastName: rowData[5], birthday: rowData[6], personalEmail: rowData[7],
      phone: rowData[8], address: rowData[9], facebookUrl: rowData[10], profilePhotoUrl: rowData[11],
      position: rowData[12], employmentType: rowData[13], dateHired: rowData[14], status: rowData[15]
    };
  } catch (e) { return { error: e.message }; }
}

/** * Fetches a summarized list of all users for the management table.
 */
function getUsersList() {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    data.shift(); 
    return data.map(function(row, i) {
      return { 
        rowIndex: i + 2, 
        username: row[1], role: row[2], workEmail: row[3], 
        firstName: row[4], lastName: row[5], position: row[12], status: row[15] 
      };
    });
  } catch (e) { return []; }
}

/* ========================================================================
   ROLE-BASED ACCESS CONTROL (RBAC) INTENTS
   ======================================================================== */

/**
 * Retrieves the role of the currently logged-in user from the Users sheet.
 * Defaults to 'Inactive' if the user is not found.
 */
function getUserRole() {
  try {
    var email = Session.getActiveUser().getEmail();
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === email) { // Work Email column (D)
        return data[i][15] === 'Inactive' ? 'Inactive' : data[i][2]; // Status (P) and Role (C)
      }
    }
    
    // Admin Failsafe: Check project settings for primary admin
    if (email === PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')) {
      return "Administrator";
    }
    
    return "Inactive";
  } catch (e) { return "Inactive"; }
}

/**
 * Fetches the specific Username for the UI greeting.
 */
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

/**
 * Compiles a list of usernames for all active Account Managers.
 */
function getAccountManagers() {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    var managers = [];
    for (var i = 1; i < data.length; i++) {
      var role = data[i][2];
      var status = data[i][15];
      if (status === 'Active' && (role === 'Account Manager' || role === 'Administrator' || role === 'Admin')) {
        managers.push(data[i][1]);
      }
    }
    return managers;
  } catch (e) { return []; }
}