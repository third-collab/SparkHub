/**
 * Users Module - Backend
 * Standardized under SparkHub Architecture Blueprint.
 * Handles: Staff profiles, permissions, RBAC, and directory management.
 */

/**
 * Securely hashes passwords using SHA-256 for database storage.
 */
function processNewUser(obj) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var hashedPw = hashPassword(obj.password);
    
    sheet.appendRow([
      new Date(), obj.username, obj.role, obj.email, hashedPw, 
      obj.firstName, obj.lastName, obj.status, "" 
    ]);
    
    SpreadsheetApp.flush();
    SystemEvent.emit("Users", "CREATE", "Add User", "INFO", obj.username, "New user access profile established via UI.");
    return "Success! User created.";
  } catch (e) { return "Error: " + e.message; }
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
          if ((data[i][2] === 'Administrator' || data[i][2] === 'Admin') && data[i][7] === 'Active') {
            activeAdminCount++;
          }
        }
        if (activeAdminCount <= 1) {
          return "Error: Cannot modify the role or status of the last active Administrator. Ensure another active admin exists first.";
        }
      }
    }

    // Preserve the old hash if the field was left blank in the UI
    var oldPassword = data[row-1][4];
    var newPassword = obj.password ? hashPassword(obj.password) : oldPassword;

    sheet.getRange(row, 2, 1, 7).setValues([[
      obj.username, obj.role, obj.email, newPassword, obj.firstName, 
      obj.lastName, obj.status
    ]]);
    
    SpreadsheetApp.flush();
    SystemEvent.emit("Users", "UPDATE", "Edit User", "INFO", obj.username, "User access profile updated.");
    return "Success! User updated.";
  } catch (e) { return "Error: " + e.message; }
}

function updateMyProfileRecord(obj) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    var rowToUpdate = -1;
    
    var loggedInUser = getLoggedInUsername();
    if (obj.username !== loggedInUser) {
      return "Error: Unauthorized profile modification.";
    }

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
    
    return {
      success: true,
      firstName: obj.firstName,
      message: "Success! Profile updated."
    };
  } catch (e) { return "Error: " + e.message; }
}

function getUserById(rowIndex) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var rowData = sheet.getRange(rowIndex, 1, 1, 9).getDisplayValues()[0];
    return {
      rowIndex: rowIndex, username: rowData[1], role: rowData[2], email: rowData[3], password: rowData[4],
      firstName: rowData[5], lastName: rowData[6], status: rowData[7], lastLogin: rowData[8]
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
    if (obj.name === 'Administrator' || obj.name === 'Admin') {
      obj.permissions = '{"ALL":["ALL"]}';
    }
    
    if (obj.rowIndex) {
      sheet.getRange(obj.rowIndex, 2, 1, 4).setValues([[ obj.name, obj.description, obj.permissions, obj.status ]]);
      SystemEvent.emit("Users", "UPDATE", "Edit Role", "INFO", obj.name, "Role permissions matrix updated.");
    } else {
      var roleId = "R-" + Utilities.getUuid().substring(0, 6).toUpperCase();
      sheet.appendRow([ roleId, obj.name, obj.description, obj.permissions, obj.status ]);
      SystemEvent.emit("Users", "CREATE", "Add Role", "INFO", obj.name, "New system role established.");
    }
    return "Success! Role saved.";
  } catch (e) { return "Error: " + e.message; }
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
    "Core System": ["Manage Settings", "Manage Roles"],
    "Access & Users": ["View Users", "Manage Users"],
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