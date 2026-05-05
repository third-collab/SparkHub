// --- USERS MODULE REGISTRY EXPORTS ---
function Users_getTriggers() {
  return [
    "Users:CREATE", "Users:UPDATE", "Users:RESET_REQUEST", "Users:PASSWORD_UPDATED",
    "Users:Roles:CREATE", "Users:Roles:UPDATE"
  ];
}

function Users_getPlaceholders() {
  return [
    "username", "firstName", "lastName", "role", "userFirst", "resetLink",
    "userEmail", "userStatus", "userTimestamp",
    "roleName", "roleDescription", "roleTimestamp"
  ];
}

/**
 * Users Module - Backend
 * Standardized under SparkHub Architecture Blueprint.
 * Handles: Staff profiles, permissions, RBAC, and directory management.
 */
function processNewUser(obj) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var hashedPw = hashPassword(obj.password);
    sheet.appendRow([
      new Date(), obj.username, obj.role, obj.email, hashedPw, 
      obj.firstName, obj.lastName, obj.status, "" 
    ]);
    var targetRow = sheet.getLastRow();
    SpreadsheetApp.flush();
    
    SystemEvent.emit("Users", "CREATE", "Add User", "INFO", obj.username, "New user access profile established via UI.", obj.email);
    return { success: true, rowIndex: targetRow, message: "Success! User created." };
  } catch (e) { return { error: "Error: " + e.message }; }
}

function updateUserRecord(obj) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    var row = parseInt(obj.rowIndex);

    var existingUsername = data[row-1][1];
    obj.username = existingUsername;
    var existingRole = data[row-1][2];
    var existingStatus = data[row-1][7];
    
    if ((existingRole === 'Administrator' || existingRole === 'Admin') && existingStatus === 'Active') {
      if (obj.role !== existingRole || obj.status !== 'Active') {
        var activeAdminCount = 0;
        for (var i = 1; i < data.length; i++) {
          if ((data[i][2] === 'Administrator' || data[i][2] === 'Admin') && data[i][7] === 'Active') activeAdminCount++;
        }
        if (activeAdminCount <= 1) return { error: "Error: Cannot modify the role or status of the last active Administrator." };
      }
    }

    var oldPassword = data[row-1][4];
    var newPassword = obj.password ? hashPassword(obj.password) : oldPassword;

    sheet.getRange(row, 2, 1, 7).setValues([[ obj.username, obj.role, obj.email, newPassword, obj.firstName, obj.lastName, obj.status ]]);
    SpreadsheetApp.flush();
    
    SystemEvent.emit("Users", "UPDATE", "Edit User", "INFO", obj.username, "User access profile updated.");
    if (existingStatus !== obj.status) {
      var actionVerb = obj.status === 'Active' ? 'activated' : 'deactivated';
      SystemEvent.emit("Users", "UPDATE", "User Status Changed", "WARN", obj.username, "User account was manually " + actionVerb + ".");
    }
    return { success: true, rowIndex: row, message: "Success! User updated." };
  } catch (e) { return { error: "Error: " + e.message }; }
}

function updateMyProfileRecord(obj) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    var rowToUpdate = -1;
    
    var loggedInUser = getLoggedInUsername();
    if (obj.username !== loggedInUser) return "Error: Unauthorized profile modification.";

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === obj.username) {
        rowToUpdate = i + 1;
        break;
      }
    }

    if (rowToUpdate === -1) return "Error: User profile not found.";

    var oldPassword = data[rowToUpdate-1][4];
    var newPassword = obj.password ? hashPassword(obj.password) : oldPassword;

    sheet.getRange(rowToUpdate, 4, 1, 4).setValues([[
      obj.email, newPassword, obj.firstName, obj.lastName
    ]]);
    
    SpreadsheetApp.flush();
    SystemEvent.emit("Users", "UPDATE", "Update Profile", "INFO", obj.username, "User updated their personal profile details.");
    
    // If the password field was filled out, emit the Password Updated trigger
    if (obj.password) {
      SystemEvent.emit(
        "Users", 
        "PASSWORD_UPDATED", 
        "Password Updated", 
        "WARN", 
        obj.username, 
        "User updated their password via their profile.", 
        obj.email,
        { userFirst: obj.firstName }
      );
    }

    return { success: true, firstName: obj.firstName, message: "Success! Profile updated." };
  } catch (e) { return "Error: " + e.message; }
}

/**
 * Creates a secure table for temporary password reset tokens
 */
function ensureTokensSheet() {
  var ss = getMainDb();
  var sheet = ss.getSheetByName("Password Tokens");
  if (!sheet) {
    sheet = ss.insertSheet("Password Tokens");
    sheet.getRange(1, 1, 1, 3).setValues([["Token", "Email", "Expiration Date"]]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Generates a 15-minute secure token and emails it to the user
 */
function sendPasswordResetEmail(email) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    var userExists = false;
    var userFirst = "User";
    var username = "User";

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]).toLowerCase() === String(email).toLowerCase()) {
        userExists = true;
        username = data[i][1];
        userFirst = data[i][5];
        break;
      }
    }

    if (!userExists) return "If that email is in our system, a reset link has been sent.";

    var token = Utilities.getUuid();
    var expiry = new Date(new Date().getTime() + 15 * 60000); // 15 mins
    var tokenSheet = ensureTokensSheet();
    tokenSheet.appendRow([token, email, expiry]);
    var resetLink = ScriptApp.getService().getUrl() + "?token=" + token;

    // Trigger the template engine to dispatch the email
    SystemEvent.emit(
      "Users", 
      "RESET_REQUEST", 
      "Password Reset Request", 
      "INFO", 
      username, 
      "User requested a password reset link.", 
      email, 
      { userFirst: userFirst, resetLink: resetLink }
    );

    return "If that email is in our system, a reset link has been sent.";
  } catch(e) { return "Error: " + e.message; }
}

/**
 * Verifies the token and overwrites the old password hash
 */
function processPasswordReset(token, newPassword) {
  try {
    var tokenSheet = ensureTokensSheet();
    var data = tokenSheet.getDataRange().getValues();
    var emailToReset = null;
    var tokenRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === token) {
        var expiry = new Date(data[i][2]);
        if (new Date() > expiry) return { success: false, message: "This reset link has expired." };
        emailToReset = data[i][1];
        tokenRow = i + 1;
        break;
      }
    }

    if (!emailToReset) return { success: false, message: "Invalid or expired reset token." };

    var userSheet = getMainDb().getSheetByName("Users");
    var userData = userSheet.getDataRange().getValues();
    var userRow = -1;

    for (var u = 1; u < userData.length; u++) {
      if (String(userData[u][3]).toLowerCase() === String(emailToReset).toLowerCase()) {
        userRow = u + 1;
        break;
      }
    }

    if (userRow === -1) return { success: false, message: "User account no longer exists." };

    var hashedPw = hashPassword(newPassword);
    userSheet.getRange(userRow, 5).setValue(hashedPw);
    tokenSheet.deleteRow(tokenRow);

    var username = userData[userRow-1][1];
    var userFirst = userData[userRow-1][5];

    // Emit the Password Updated trigger
    SystemEvent.emit(
      "Users", 
      "PASSWORD_UPDATED", 
      "Password Updated", 
      "WARN", 
      username, 
      "User reset their password via email link.", 
      emailToReset,
      { userFirst: userFirst }
    );

    return { success: true, message: "Password updated successfully!" };
  } catch(e) { return { success: false, message: "Error: " + e.message }; }
}

function getUserById(rowIndex) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var rowData = sheet.getRange(rowIndex, 1, 1, 9).getDisplayValues()[0];
    var username = rowData[1];
    var lastUpdated = getEventTimestampFromLogs("Users", "UPDATE", username);

    return {
      rowIndex: rowIndex, timestamp: rowData[0], lastUpdated: lastUpdated, username: username, role: rowData[2], email: rowData[3], 
      password: rowData[4], firstName: rowData[5], lastName: rowData[6], status: rowData[7], lastLogin: rowData[8]
    };
  } catch (e) { return { error: e.message }; }
}

function getRoleById(rowIndex) {
  try {
    var sheet = ensureRolesSheet();
    var rowData = sheet.getRange(rowIndex, 1, 1, 5).getDisplayValues()[0];
    var roleName = rowData[1];
    var createdOn = getEventTimestampFromLogs("Users:Roles", "CREATE", roleName);
    if (createdOn === "Not recorded" && (roleName === "Administrator" || roleName === "Admin")) createdOn = "System Default";
    var lastUpdated = getEventTimestampFromLogs("Users:Roles", "UPDATE", roleName);

    return {
      rowIndex: rowIndex, id: rowData[0], name: roleName, description: rowData[2],
      permissions: rowData[3], status: rowData[4], createdOn: createdOn, lastUpdated: lastUpdated
    };
  } catch (e) { return { error: e.message }; }
}

function getUserProfileByUsername(username) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getDisplayValues(); 
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(username)) { 
        return {
          rowIndex: i + 1, username: data[i][1], role: data[i][2], email: data[i][3], 
          password: data[i][4], firstName: data[i][5], lastName: data[i][6], 
          status: data[i][7], lastLogin: data[i][8] 
        };
      }
    }
    return { error: "User profile not found in directory." };
  } catch (e) { return { error: e.message }; }
}

function getUsersList() {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    data.shift(); 
    var validUsers = [];
    data.forEach(function(row, i) {
      if (row[1]) { 
        validUsers.push({ 
          rowIndex: i + 2, 
          username: String(row[1]), role: String(row[2]), email: String(row[3]), 
          firstName: String(row[5]), lastName: String(row[6]), status: String(row[7]) 
        });
      }
    });
    return validUsers;
  } catch (e) { return []; }
}

function getUserRole() {
  try {
    var email = Session.getActiveUser().getEmail();
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === email) { 
        return data[i][7] === 'Inactive' ? 'Inactive' : data[i][2]; 
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

function getLoggedInUserFirstName() {
  try {
    var email = Session.getActiveUser().getEmail();
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === email) return data[i][5];
    }
    return "User";
  } catch (e) { return "User"; }
}

function updateLastLogin() {
  try {
    var email = Session.getActiveUser().getEmail();
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === email) { 
        sheet.getRange(i + 1, 9).setValue(new Date());
        break;
      }
    }
  } catch (e) { console.error("Failed to update last login: " + e.message); }
}

function ensureRolesSheet() {
  var ss = getMainDb();
  var sheet = ss.getSheetByName("Roles");
  if (!sheet) {
    initializeSheet(ss, "Roles", ["Role ID", "Role Name", "Description", "Permissions JSON", "Status"]);
    sheet = ss.getSheetByName("Roles");
    var adminPerms = JSON.stringify({ "Core System": ["Manage Settings", "Manage Roles"], "Access & Users": ["View Users", "Manage Users"], "Templates": ["View Templates", "Manage Templates"] });
    sheet.appendRow(["R-ADMIN", "Administrator", "Unrestricted system access.", adminPerms, "Active"]);
  }
  return sheet;
}

function getRolesList() {
  try {
    var sheet = ensureRolesSheet();
    var data = sheet.getDataRange().getValues();
    data.shift();
    return data.map(function(row, i) {
      return { 
        rowIndex: i + 2, id: row[0], name: row[1], 
        description: row[2], permissions: row[3], status: row[4] 
      };
    });
  } catch(e) { return []; }
}

function saveRoleRecord(obj) {
  try {
    var sheet = ensureRolesSheet();
    var activeUserEmail = Session.getActiveUser().getEmail(); 
    var targetRow;
    
    if (obj.name === 'Administrator' || obj.name === 'Admin') obj.permissions = '{"ALL":["ALL"]}';
    
    if (obj.rowIndex) {
      targetRow = parseInt(obj.rowIndex);
      var oldData = sheet.getRange(targetRow, 1, 1, 5).getValues()[0];
      var oldStatus = oldData[4];
      
      sheet.getRange(targetRow, 2, 1, 4).setValues([[ obj.name, obj.description, obj.permissions, obj.status ]]);
      
      // NEW: Uses Sub-Module SystemEvent identifier
      SystemEvent.emit("Users:Roles", "UPDATE", "Edit Role", "INFO", obj.name, "Role permissions matrix updated.", activeUserEmail);
      if (oldStatus !== obj.status) {
        var actionVerb = obj.status === 'Active' ? 'activated' : 'deactivated';
        SystemEvent.emit("Users:Roles", "UPDATE", "Role Status Changed", "WARN", obj.name, "Role was manually " + actionVerb + ".", activeUserEmail);
      }
    } else {
      var roleId = "R-" + Utilities.getUuid().substring(0, 6).toUpperCase();
      sheet.appendRow([ roleId, obj.name, obj.description, obj.permissions, obj.status ]);
      targetRow = sheet.getLastRow();
      
      // NEW: Uses Sub-Module SystemEvent identifier
      SystemEvent.emit("Users:Roles", "CREATE", "Add Role", "INFO", obj.name, "New system role established.", activeUserEmail);
    }
    return { success: true, rowIndex: targetRow, message: "Success! Role saved." };
  } catch (e) { return { error: "Error: " + e.message }; }
}

function getUserPermissions(roleName) {
  if (roleName === 'Administrator' || roleName === 'Admin') {
    return '{"ALL":["ALL"]}';
  }
  try {
    var sheet = ensureRolesSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === roleName && data[i][4] === 'Active') {
        return data[i][3] ? data[i][3] : "{}"; 
      }
    }
  } catch(e) {}
  return "{}";
}

function verifyUserCredentials(loginId, password) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    var loginLower = loginId.toLowerCase();

    for (var i = 1; i < data.length; i++) {
      var rowUsername = String(data[i][1]).toLowerCase();
      var rowEmail = String(data[i][3]).toLowerCase();
      var rowPassword = String(data[i][4]);
      
      // Hash the incoming plaintext password and compare it against the database hash
      var hashedInput = hashPassword(password);
      
      if ((rowUsername === loginLower || rowEmail === loginLower) && rowPassword === hashedInput) {
        if (String(data[i][7]) === 'Inactive') {
          return { success: false, message: "Account is inactive." };
        }

        var role = String(data[i][2]);
        return {
          success: true,
          username: String(data[i][1]),
          firstName: String(data[i][5]), 
          role: role,
          permissions: getUserPermissions(role) 
        };
      }
    }
    return { success: false, message: "Invalid credentials." };
  } catch(e) {
    return { success: false, message: "System error during authentication." };
  }
}

function getDynamicPermissionMatrix() {
  var matrix = {
    "Core System": ["Manage Settings"],
    "Access & Users": ["View Users", "Manage Users", "Manage Roles"],
    "Templates": ["View Templates", "Manage Templates"],
    "System Logs": ["View Logs"]
  };
  
  var installed = PropertiesService.getScriptProperties().getProperty('INSTALLED_MODULES');
  if (installed) {
    installed.split(',').forEach(function(modName) {
      var mod = modName.trim();
      var funcName = mod + "_getPermissions";
      if (typeof this[funcName] === 'function') {
        matrix[mod + " Module"] = this[funcName]();
      } else {
        matrix[mod + " Module"] = ["View " + mod, "Manage " + mod];
      }
    });
  }
  return matrix;
}