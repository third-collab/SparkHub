/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Users.gs
 * VERSION: 1.6
 * SYNC STATUS: Fully Synchronized with UsersData.html
 */

// ========================================================================
// 1. REGISTRY EXPORTS
// ========================================================================
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

function Users_getPermissions() { 
  return ["View Users", "Add Users", "Manage Users", "View Roles", "Create Roles", "Manage Roles", "Manage Settings"];
}

function Users_getLookups() {
  var vectors = { users: [], roles: [] };
  try {
    var userSheet = getMainDb().getSheetByName("Users");
    var uData = userSheet.getDataRange().getValues();
    for (var i = 1; i < uData.length; i++) {
      if (uData[i][1] && uData[i][11] !== 'Inactive') {
        var first = String(uData[i][7] || "").trim();
        var last = String(uData[i][8] || "").trim();
        var shortName = first + (last ? " " + last.charAt(0).toUpperCase() + "." : "");
        
        vectors.users.push({
          id: String(uData[i][1]), 
          name: shortName, // Default display standard: First Name + Last Initial
          fullName: first + " " + last
        });
      }
    }
  } catch(e) { console.warn("Users lookup broadcast failed: " + e.message); }
  
  try {
    var rolesSheet = getMainDb().getSheetByName("Roles");
    if (rolesSheet) {
      var rData = rolesSheet.getDataRange().getDisplayValues();
      for (var r = 1; r < rData.length; r++) {
        if (rData[r][2] && rData[r][5] === 'Active') {
          vectors.roles.push({
            id: String(rData[r][2]), 
            name: String(rData[r][2])
          });
        }
      }
    }
  } catch(e) { console.warn("Roles lookup broadcast failed: " + e.message); }
  return vectors;
}

// ========================================================================
// 2. CORE PROCESSORS
// ========================================================================
function verifyUserCredentials(loginId, password) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    var loginLower = loginId.toLowerCase();

    for (var i = 1; i < data.length; i++) {
      var rowUserId = String(data[i][1]);
      var rowUsername = String(data[i][2]).toLowerCase();
      var rowGoogleEmail = String(data[i][3]).toLowerCase();
      var rowPassword = String(data[i][6]);
      
      var hashedInput = hashPassword(password);
      if ((rowUsername === loginLower || rowGoogleEmail === loginLower) && rowPassword === hashedInput) {
        if (String(data[i][11]) === 'Inactive') return { success: false, message: "Account is inactive." };
        var role = String(data[i][5]);
        return {
          success: true,
          userId: rowUserId,
          username: String(data[i][2]),
          firstName: String(data[i][7]), 
          role: role,
          permissions: getUserPermissions(role),
          dashboardLayout: getResolvedDashboardLayout(rowGoogleEmail, role)
        };
      }
    }
    return { success: false, message: "Invalid credentials." };
  } catch(e) { return { success: false, message: "System error during authentication." }; }
}

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
        username = data[i][2];  // Index 2 maps to customizable Username handle string
        userFirst = data[i][7]; // Index 7 maps to First Name string
        break;
      }
    }

    if (!userExists) return "If that email is in our system, a reset link has been sent.";
    var token = Utilities.getUuid();
    var expiry = new Date(new Date().getTime() + 15 * 60000); 
    var tokenSheet = ensureTokensSheet();
    tokenSheet.appendRow([token, email, expiry]);
    var resetLink = ScriptApp.getService().getUrl() + "?token=" + token;

    SystemEvent.emit("Users", "RESET_REQUEST", "Password Reset Request", "INFO", username, "User requested a password reset link.", email, { userFirst: userFirst, resetLink: resetLink });
    return "If that email is in our system, a reset link has been sent.";
  } catch(e) { return "Error: " + e.message; }
}

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
    // Enforces write on 1-based Column 7 (Password) instead of corrupting Column 5 (System Email)
    userSheet.getRange(userRow, 7).setValue(hashedPw);
    tokenSheet.deleteRow(tokenRow);

    var username = userData[userRow-1][2]; // Index 2 maps to customizable Username handle string
    var userFirst = userData[userRow-1][7]; // Index 7 maps to First Name string
    SystemEvent.emit("Users", "PASSWORD_UPDATED", "Password Updated", "WARN", username, "User reset their password via email link.", emailToReset, { userFirst: userFirst });
    return { success: true, message: "Password updated successfully!" };
  } catch(e) { return { success: false, message: "Error: " + e.message };
  }
}

function updateLastLogin() {
  try {
    var email = Session.getActiveUser().getEmail();
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === email) { 
        // Enforces write on 1-based Column 10 (Last Login) instead of overriding Column 9 (Last Name)
        sheet.getRange(i + 1, 10).setValue(new Date());
        break;
      }
    }
  } catch (e) { console.error("Failed to update last login: " + e.message);
  }
}

function saveUsersModuleConfig(payload) {
  try {
    if (!payload) throw new Error("No payload provided.");
    const props = PropertiesService.getScriptProperties();
    
    // Check for values before saving to prevent undefined crashes
    if (payload.defaultRole) props.setProperty('CONF_USERS_DEFAULT_ROLE', payload.defaultRole);
    if (payload.adminEmail) props.setProperty('ADMIN_EMAIL', payload.adminEmail);
    if (payload.authMode) props.setProperty('AUTH_MODE', payload.authMode);
    
    // Group updates inside the primary 'System' routing layer per core directive standard
    SystemEvent.emit("System", "UPDATE", "Config Updated", "INFO", "Users", "User module settings updated locally.");
    return { success: true, message: "User settings saved." };
  } catch (e) {
    return { error: "Users.gs: " + e.message };
  }
}

// ========================================================================
// 3. READ / GET FUNCTIONS
// ========================================================================
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
          userId: String(row[1]),
          username: String(row[2]), 
          email: String(row[3]), 
          systemEmail: String(row[4]),
          role: String(row[5]), 
          firstName: String(row[7]), 
          lastName: String(row[8]), 
          status: String(row[11]) 
        });
      }
    });
    return validUsers;
  } catch (e) { return []; }
}

function getUserById(rowIndex) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var rowData = sheet.getRange(rowIndex, 1, 1, 12).getDisplayValues()[0];
    var userId = rowData[1];
    var lastUpdated = getEventTimestampFromLogs("Users", "UPDATE", userId);
    return {
      rowIndex: rowIndex, timestamp: rowData[0], lastUpdated: lastUpdated, userId: userId, username: rowData[2], 
      email: rowData[3], systemEmail: rowData[4], role: rowData[5], password: rowData[6], firstName: rowData[7], 
      lastName: rowData[8], lastLogin: rowData[9], status: rowData[11]
    };
  } catch (e) { return { error: e.message }; }
}

function getUserProfileByUsername(username) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    // Swap array ingestion line to utilize raw values comparisons to protect unassigned password data-type checks
    var data = sheet.getDataRange().getValues(); 
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === String(username).toLowerCase()) { 
        return {
          rowIndex: i + 1,
          userId: String(data[i][1]),
          username: String(data[i][2]),
          email: String(data[i][3]),
          systemEmail: String(data[i][4]),
          role: String(data[i][5]),
          password: String(data[i][6] || ""),
          firstName: String(data[i][7]),
          lastName: String(data[i][8]),
          lastLogin: data[i][9] instanceof Date ? Utilities.formatDate(data[i][9], getMainDb().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss") : String(data[i][9] || ""),
          status: String(data[i][11])
        };
      }
    }
    return { error: "User profile not found in directory." };
  } catch (e) { return { error: e.message }; }
}

function getRolesList() {
  try {
    var sheet = ensureRolesSheet();
    var data = sheet.getDataRange().getDisplayValues(); 
    data.shift();
    var validRoles = [];
    data.forEach(function(row, i) {
      if (row[2]) {
        validRoles.push({ 
          rowIndex: i + 2, timestamp: row[0], id: row[1], name: row[2], 
          description: row[3], permissions: row[4], status: row[5] 
        });
      }
    });
    return validRoles;
  } catch(e) { return []; }
}

function getRoleById(rowIndex) {
  try {
    var sheet = ensureRolesSheet();
    var rowData = sheet.getRange(rowIndex, 1, 1, 6).getDisplayValues()[0];
    var roleName = rowData[2];
    var createdOn = rowData[0] || "System Default";
    var lastUpdated = getEventTimestampFromLogs("Users:Roles", "UPDATE", roleName);

    return {
      rowIndex: rowIndex, timestamp: rowData[0], id: rowData[1], name: roleName, description: rowData[3],
      permissions: rowData[4], status: rowData[5], createdOn: createdOn, lastUpdated: lastUpdated
    };
  } catch (e) { return { error: e.message }; }
}

function getUserPermissions(roleName) {
  // Standardized onto the explicit string token string to prevent loose syntax evaluation risks
  if (roleName === 'Administrator') return '{"ALL":["ALL"]}';
  try {
    var sheet = ensureRolesSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === roleName && data[i][5] === 'Active') return data[i][4] ? data[i][4] : "{}"; 
    }
  } catch(e) {}
  return "{}";
}

function getResolvedDashboardLayout(email, roleName) {
  try {
    var db = getMainDb();
    var usersData = db.getSheetByName("Users").getDataRange().getValues();
    for (var i = 1; i < usersData.length; i++) {
      if (usersData[i][3] === email) {
        // Enforces lookup on index 10 (Column K) to isolate the genuine widget layout matrix string
        var userConfig = usersData[i][10];
        if (userConfig && userConfig.trim() !== "" && userConfig !== "[]") return userConfig;
        break;
      }
    }
    var rolesData = ensureRolesSheet().getDataRange().getValues();
    for (var r = 1; r < rolesData.length; r++) {
      if (rolesData[r][2] === roleName) {
        var roleConfig = rolesData[r][6];
        if (roleConfig && roleConfig.trim() !== "" && roleConfig !== "[]") return roleConfig;
        break;
      }
    }
  } catch(e) {}
  return "DEFAULT";
}

function getUserRole() {
  try {
    var email = Session.getActiveUser().getEmail();
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === email) {
        // Index 11 maps to Status, Index 5 maps to Role under the 12-column schema
        return data[i][11] === 'Inactive' ? 'Inactive' : data[i][5];
      }
    }
    if (email === PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')) return "Administrator";
    return "Inactive";
  } catch (e) { return "Inactive"; }
}

function getLoggedInUsername() {
  try {
    var email = Session.getActiveUser().getEmail();
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === email) return data[i][2]; // Index 2 extracts the customizable Username handle string
    }
    return email.split('@')[0];
  } catch (e) { return "User";
  }
}

function getLoggedInUserFirstName() {
  try {
    var email = Session.getActiveUser().getEmail();
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === email) return data[i][7];
    }
    return "User";
  } catch (e) { return "User"; }
}

function getLoggedInUserId() {
  try {
    var email = Session.getActiveUser().getEmail();
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === email) return data[i][1] || "U-SYSTEM";
    }
    return "U-SYSTEM";
  } catch (e) { return "U-SYSTEM"; }
}

// ========================================================================
// 4. WRITE / SAVE FUNCTIONS
// ========================================================================
function createUserRecord(obj, isAutomatedOnboarding) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    
    var baseUsername = String(obj.username || "").trim().toLowerCase().replace(/[^a-z0-9_]/g, "");
    if (!baseUsername) throw new Error("A valid alphanumeric username handle is required.");
    
    var finalizedUsername = baseUsername;
    var suffixCounter = 1;
    var exists = false;
    
    // Scan existing rows to check for username availability
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === baseUsername) {
        exists = true;
        break;
      }
    }
    
    if (exists) {
      if (isAutomatedOnboarding) {
        // Automatic onboarding engine loops until hitting a vacant suffix row location
        var activeUsernames = data.map(function(r) { return String(r[2]).toLowerCase(); });
        while (activeUsernames.indexOf(finalizedUsername) > -1) {
          finalizedUsername = baseUsername + suffixCounter;
          suffixCounter++;
        }
      } else {
        return { success: false, error: "Error: The username handle '" + baseUsername + "' is already allocated." };
      }
    }
    
    var generatedId = "U-" + Math.floor(1000 + Math.random() * 9000);
    var hashedPw = hashPassword(obj.password || Utilities.getUuid().substring(0, 10));
    
    // Map data fields strictly matching the optimized 12-column layout mapping rules
    var newRow = [
      new Date(),
      generatedId,
      finalizedUsername,
      String(obj.email || "").trim(),       // Google Email (SSO Principal Check)
      String(obj.systemEmail || obj.email || "").trim(), // System Email (Outbound Routing Target)
      obj.role || "Client",
      hashedPw,
      obj.firstName || "",
      obj.lastName || "",
      "",  // Last Login
      "",  // Dashboard Config
      obj.status || "Active" // Status locks column 12 per Section 3.F mandate
    ];
    
    sheet.appendRow(newRow);
    var targetRow = sheet.getLastRow();
    SpreadsheetApp.flush(); // Mandated race condition safeguard
    
    var dataContextMap = {
      username: finalizedUsername, 
      firstName: obj.firstName || "User", 
      lastName: obj.lastName || "", 
      roleName: obj.role || "Client",
      target_username: finalizedUsername,
      target_firstName: obj.firstName || "User",
      target_lastName: obj.lastName || "",
      target_role: obj.role || "Client"
    };

    // Packages detailed contextual data parameters into extraData to shield against placeholder collisions
    SystemEvent.emit("Users", "CREATE", "Add User", "INFO", generatedId, "New user profile established.", obj.email, dataContextMap);
    
    // Fire separate security privilege assignment trigger if user is created as an Administrator
    if (obj.role === 'Administrator') {
      SystemEvent.emit("Users", "ADMIN_ASSIGNED", "Admin Privileges Assigned", "WARN", generatedId, "User created with Administrator role privileges.", obj.email, dataContextMap);
    }

    return { success: true, rowIndex: targetRow, username: finalizedUsername, userId: generatedId };
  } catch (e) { 
    return { error: "Error: " + e.message }; 
  }
}

function updateUserRecord(obj) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    var row = parseInt(obj.rowIndex, 10);

    var existingUserId = data[row-1][1];
    var oldUsername = data[row-1][2];
    var existingRole = data[row-1][5];
    var existingStatus = data[row-1][11];
    
    var newUsername = String(obj.username || "").trim().toLowerCase().replace(/[^a-z0-9_]/g, "");
    if (!newUsername) throw new Error("A valid alphanumeric username handle is required.");
    
    // Validate uniqueness if the username handle is modified
    if (newUsername !== oldUsername.toLowerCase()) {
      for (var i = 1; i < data.length; i++) {
        if (i !== (row - 1) && String(data[i][2]).toLowerCase() === newUsername) {
          return { success: false, error: "Error: The username handle '" + newUsername + "' is already allocated." };
        }
      }
    }

    // Standardized onto the explicit string token string to prevent loose syntax evaluation risks
    if (existingRole === 'Administrator' && existingStatus === 'Active') {
      if (obj.role !== existingRole || obj.status !== 'Active') {
        var activeAdminCount = 0;
        for (var i = 1; i < data.length; i++) {
          if (data[i][5] === 'Administrator' && data[i][11] === 'Active') activeAdminCount++;
        }
        if (activeAdminCount <= 1) return { error: "Error: Cannot modify the role or status of the last active Administrator." };
      }
    }

    var oldPassword = data[row-1][6];
    var newPassword = obj.password ? hashPassword(obj.password) : oldPassword;

    sheet.getRange(row, 2, 1, 11).setValues([[
      existingUserId,
      newUsername,
      obj.email,
      obj.systemEmail || obj.email,
      obj.role,
      newPassword,
      obj.firstName,
      obj.lastName,
      data[row-1][9],  // Preserve Last Login column placement
      data[row-1][10], // Preserve Dashboard Config layout
      obj.status       // Status locks column 12 per Section 
    ]]);
    SpreadsheetApp.flush(); // Mandated race condition safeguard
    
    var dataContextMap = {
      username: newUsername,
      firstName: obj.firstName,
      lastName: obj.lastName,
      roleName: obj.role,
      target_username: newUsername,
      target_firstName: obj.firstName,
      target_lastName: obj.lastName,
      target_role: obj.role
    };

    SystemEvent.emit("Users", "UPDATE", "Edit User", "INFO", existingUserId, "User access profile updated.", obj.email, dataContextMap);
    
    // Fire separate security privilege assignment trigger if user is elevated to an Administrator role
    if (obj.role === 'Administrator' && existingRole !== 'Administrator') {
      SystemEvent.emit("Users", "ADMIN_ASSIGNED", "Admin Privileges Assigned", "WARN", existingUserId, "User role elevated to Administrator.", obj.email, dataContextMap);
    }
    if (existingStatus !== obj.status) {
      var actionVerb = obj.status === 'Active' ? 'activated' : 'deactivated';
      SystemEvent.emit("Users", "UPDATE", "User Status Changed", "WARN", existingUserId, "User account was manually " + actionVerb + ".");
    }
    return { success: true, rowIndex: row, message: "Success! User updated." };
  } catch (e) { 
    return { error: "Error: " + e.message };
  }
}

function updateMyProfileRecord(obj) {
  try {
    var sheet = getMainDb().getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    var rowToUpdate = -1;
    
    var loggedInUser = getLoggedInUsername();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === loggedInUser.toLowerCase()) {
        rowToUpdate = i + 1;
        break;
      }
    }

    if (rowToUpdate === -1) return "Error: User profile not found.";
    var existingUserId = data[rowToUpdate-1][1];
    var oldUsername = data[rowToUpdate-1][2];
    
    var newUsername = String(obj.username || "").trim().toLowerCase().replace(/[^a-z0-9_]/g, "");
    if (!newUsername) return "Error: A valid alphanumeric username handle is required.";
    
    // Validate uniqueness if the username handle is modified during self-profile editing
    if (newUsername !== oldUsername.toLowerCase()) {
      for (var i = 1; i < data.length; i++) {
        if (i !== (rowToUpdate - 1) && String(data[i][2]).toLowerCase() === newUsername) {
          return "Error: The username handle '" + newUsername + "' is already allocated.";
        }
      }
    }

    sheet.getRange(rowToUpdate, 3).setValue(newUsername);

    var oldPassword = data[rowToUpdate-1][6];
    var newPassword = obj.password ? hashPassword(obj.password) : oldPassword;
    // GOVERNANCE MANTRA: Force use of data[rowToUpdate-1][3] to keep the Google SSO Email authentication token locked
    sheet.getRange(rowToUpdate, 4, 1, 2).setValues([[ data[rowToUpdate-1][3], obj.systemEmail || data[rowToUpdate-1][3] ]]);
    sheet.getRange(rowToUpdate, 7, 1, 3).setValues([[ newPassword, obj.firstName, obj.lastName ]]);
    
    SpreadsheetApp.flush(); // Mandated race condition safeguard
    SystemEvent.emit("Users", "UPDATE", "Update Profile", "INFO", existingUserId, "User updated their personal profile details.");
    if (obj.password) {
      SystemEvent.emit("Users", "PASSWORD_UPDATED", "Password Updated", "WARN", existingUserId, "User updated their password via their profile.", data[rowToUpdate-1][3], { userFirst: obj.firstName });
    }

    return { success: true, username: newUsername, firstName: obj.firstName, message: "Success! Profile updated." };
  } catch (e) { 
    return "Error: " + e.message;
  }
}

function saveRoleRecord(obj) {
  try {
    var sheet = ensureRolesSheet();
    var activeUserEmail = Session.getActiveUser().getEmail(); 
    var targetRow;
    // Standardized onto the explicit string token string to prevent loose syntax evaluation risks
    if (obj.name === 'Administrator') obj.permissions = '{"ALL":["ALL"]}';
    if (obj.rowIndex) {
      targetRow = parseInt(obj.rowIndex);
      var oldData = sheet.getRange(targetRow, 1, 1, 6).getValues()[0];
      var oldStatus = oldData[5];
      
      sheet.getRange(targetRow, 3, 1, 4).setValues([[ obj.name, obj.description, obj.permissions, obj.status ]]);
      
      SystemEvent.emit("Users:Roles", "UPDATE", "Edit Role", "INFO", obj.name, "Role permissions matrix updated.", activeUserEmail, { roleTimestamp: oldData[0], roleName: obj.name, roleDescription: obj.description });
      
      if (oldStatus !== obj.status) {
        var actionVerb = obj.status === 'Active' ? 'activated' : 'deactivated';
        SystemEvent.emit("Users:Roles", "UPDATE", "Role Status Changed", "WARN", obj.name, "Role was manually " + actionVerb + ".", activeUserEmail, { roleTimestamp: oldData[0], roleName: obj.name, roleDescription: obj.description });
      }
    } else {
      var roleId = "R-" + Utilities.getUuid().substring(0, 6).toUpperCase();
      var now = new Date();
      sheet.appendRow([ now, roleId, obj.name, obj.description, obj.permissions, obj.status, "" ]);
      targetRow = sheet.getLastRow();
      
      SystemEvent.emit("Users:Roles", "CREATE", "Add Role", "INFO", obj.name, "New system role established.", activeUserEmail, { roleTimestamp: now, roleName: obj.name, roleDescription: obj.description });
    }
    return { success: true, rowIndex: targetRow, message: "Success! Role saved." };
  } catch (e) { return { error: "Error: " + e.message }; }
}

// ========================================================================
// 5. INTERNAL HELPERS
// ========================================================================
function ensureRolesSheet() {
  var ss = getMainDb();
  var sheet = ss.getSheetByName("Roles");
  if (!sheet) {
    initializeSheet(ss, "Roles", ["Timestamp", "Role ID", "Role Name", "Description", "Permissions JSON", "Status", "Dashboard Config"]);
    sheet = ss.getSheetByName("Roles");
    var adminPerms = JSON.stringify({ "Core System": ["Manage Settings", "Manage Roles"], "Access & Users": ["View Users", "Manage Users"], "Templates": ["View Templates", "Manage Templates"], "System Logs": ["View Logs"] });
    sheet.appendRow([new Date(), "R-ADMIN", "Administrator", "Unrestricted system access.", adminPerms, "Active", ""]);
  }
  return sheet;
}

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
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */