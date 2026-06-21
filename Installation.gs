/**
 * [SPARKHUB INTEGRITY MANIFEST]
 * VERSION: 2.2
 * MANDATORY LOGIC: performUiInstallation, runInstallation, setupCoreDatabase, 
 * seedCoreAssets, setupLogsDatabase, Master Webhook Call, Deep DB Integrity Loop,
 * System:INSTALL Trigger, Users:CREATE Trigger
 */

var SPARKHUB_VERSION = "1.0.0";
var MASTER_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbyderUFTDgJjjSb4ML2xpXzRnfKp_yBLkYlpaKdZcWZLowtmiutt-QZsg7OMq0enBJljw/exec";

function performUiInstallation(data) {
  try {
    if (data.sysName.trim().toLowerCase() === 'sparkhub') throw new Error("Restricted name.");
    var props = PropertiesService.getScriptProperties();
    var installerEmail = Session.getActiveUser().getEmail();
    var adminUsername = installerEmail.split('@')[0];

    props.setProperty('ROOT_FOLDER_ID', data.rootId);
    props.setProperty('SYSTEM_NAME', data.sysName);
    props.setProperty('ADMIN_EMAIL', installerEmail);
    // Cache personal names temporarily for the multi-step structural spreadsheet generation process
    props.setProperty('TMP_FIRST_NAME', data.firstName || "System");
    props.setProperty('TMP_LAST_NAME', data.lastName || "Admin");
    // Set default communication and maintenance standards
    props.setProperty('LOGS_RETENTION_DAYS', '90'); 
    props.setProperty('TPL_WHITELIST', installerEmail);
    props.setProperty('TPL_DEFAULT_WRAPPER', 'Internal Communication');
    props.setProperty('TPL_LINK_TRACKING', 'false');
    props.setProperty('CLIENT_ID', "CID-" + Utilities.getUuid().substring(0, 8).toUpperCase());
    props.setProperty('INSTANCE_SECRET', Utilities.getUuid());
    
    // NEW: Save the Webhook Secret so daily cron jobs can authenticate!
    var masterSecret = "SparkHub-Sec-92vM4xL7qP8nR3wK1bC6";
    props.setProperty('WEBHOOK_SECRET', masterSecret);
    
    // Register the integrated UserBar plugin
    props.setProperty('INSTALLED_PLUGINS', 'UserBar');

    runInstallation();
    setupSystemTriggers();
    // IMMUTABLE ANCHOR: Prevent concurrent write-locks from overwriting the User creation log
    Utilities.sleep(1500);
    // Route the installation success via the Event Broker, passing the dynamic data
    SystemEvent.emit("System", "INSTALL", "System Installation", "WARN", "Core Architecture", "SparkHub core deployed.", installerEmail, { sysName: data.sysName, adminUsername: adminUsername });
    props.setProperty('ENVIRONMENT', 'Sandbox');
    
    // IMMUTABLE ANCHOR: Master Webhook Reporting
    if (MASTER_WEBHOOK_URL) {
      var payload = { 
        action: "install", 
        secretKey: masterSecret,
        clientId: props.getProperty('CLIENT_ID'), 
        clientName: data.sysName, 
        clientEmail: installerEmail, 
        version: SPARKHUB_VERSION,
        appUrl: ScriptApp.getService().getUrl(),
        timezone: Session.getScriptTimeZone(),
        rootId: data.rootId,
        databaseId: props.getProperty('DATABASE_ID'),
        status: "Active",
        lastPing: new Date().toISOString(),
        instanceSecret: props.getProperty('INSTANCE_SECRET'),
        date: new Date().toISOString()
      };
      UrlFetchApp.fetch(MASTER_WEBHOOK_URL, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
    }
    
    return "Success|" + ScriptApp.getService().getUrl();
  } catch (e) { return "Error: " + e.message; }
}

function runInstallation() {
  var rootFolder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID'));
  var assetsFolder = getOrCreateFolder(rootFolder, "System Assets");
  
  getOrCreateFolder(getOrCreateFolder(assetsFolder, "Settings"), "Images");
  getOrCreateFolder(getOrCreateFolder(assetsFolder, "Users"), "Photos");
  getOrCreateFolder(getOrCreateFolder(assetsFolder, "Templates"), "Emails");

  // Create Logs first so that SystemEvent.emit can record the core installation
  setupLogsDatabase(rootFolder);
  setupCoreDatabase(rootFolder);
}

function setupLogsDatabase(rootFolder) {
  // Dynamically fetch the system name defined in the wizard
  var sysName = PropertiesService.getScriptProperties().getProperty('SYSTEM_NAME') || "SparkHub";
  var ss = SpreadsheetApp.create(sysName + " Logs Database");
  
  var dbId = ss.getId();
  // 1. Immediate Save
  PropertiesService.getScriptProperties().setProperty('LOGS_DATABASE_ID', dbId);

  // 2. Database Integrity Loop
  var verifiedDb = null;
  for (var i = 0; i < 5; i++) {
    try {
      verifiedDb = SpreadsheetApp.openById(dbId);
      DriveApp.getFileById(dbId).moveTo(rootFolder);
      break;
    } catch(e) {
      if (i === 4) throw new Error("Drive Indexing Timeout for Logs Database.");
      Utilities.sleep(3000);
    }
  }

  initializeSheet(verifiedDb, "System Logs", ["Timestamp", "Module", "Action Type", "Action Name", "Severity", "Actor", "Target Entity", "Log Details", "Environment"]);
  if (verifiedDb.getSheetByName("Sheet1")) verifiedDb.deleteSheet(verifiedDb.getSheetByName("Sheet1"));
}

function setupCoreDatabase(rootFolder) {
  // Dynamically fetch the system name defined in the wizard
  var sysName = PropertiesService.getScriptProperties().getProperty('SYSTEM_NAME') || "SparkHub";
  var ss = SpreadsheetApp.create(sysName + " Database");
  
  var dbId = ss.getId();
  // 1. Immediate Save
  PropertiesService.getScriptProperties().setProperty('DATABASE_ID', dbId);

  // 2. Database Integrity Loop
  var verifiedDb = null;
  for (var i = 0; i < 5; i++) {
    try {
      verifiedDb = SpreadsheetApp.openById(dbId);
      DriveApp.getFileById(dbId).moveTo(rootFolder);
      break;
    } catch(e) {
      if (i === 4) throw new Error("Drive Indexing Timeout for Core Database.");
      Utilities.sleep(3000);
    }
  }

  // 3. Schema Initialization (Remapped to support side-by-side emails and locked Status column layout)
  initializeSheet(verifiedDb, "Users", ["Timestamp", "User ID", "Username", "Google Email", "System Email", "Role", "Password", "First Name", "Last Name", "Last Login", "Dashboard Config", "Status"]);
  initializeSheet(verifiedDb, "Roles", ["Timestamp", "Role ID", "Role Name", "Description", "Permissions JSON", "Status", "Dashboard Config"]);
  // Initializes fresh deployments with a standardized 15-column templates table configuration matrix
  initializeSheet(verifiedDb, "Templates", ["Timestamp", "ID", "Name", "Description", "Category", "Module", "Trigger", "Subject", "Body", "Wrapper", "Status", "To Recipients", "CC Recipients", "BCC Recipients", "Dispatch Mode"]);
  initializeSheet(verifiedDb, "Wrappers", ["Timestamp", "Wrapper ID", "Name", "HTML Content", "Status"]);

  // 4. Admin Creation
  var adminEmail = Session.getActiveUser().getEmail();
  var adminUsername = adminEmail.split('@')[0];
  var props = PropertiesService.getScriptProperties();
  var finalFirst = props.getProperty('TMP_FIRST_NAME') || "System";
  var finalLast = props.getProperty('TMP_LAST_NAME') || "Admin";
  
  verifiedDb.getSheetByName("Users").appendRow([new Date(), "U-1001", adminUsername, adminEmail, adminEmail, "Administrator", "", finalFirst, finalLast, new Date(), "", "Active"]);
  
  // Seeds the exact granular actions, tabs, and categories mapped inside our Ecosystem Permission Matrix Blueprint
  var adminPerms = JSON.stringify({ 
    "Core System": ["View Settings", "Manage Settings"], 
    "Access & Users": ["View Users", "Add Users", "Manage Users", "View Roles", "Create Roles", "Manage Roles", "Manage Settings"], 
    "Templates & Wrappers": ["View Templates", "Create Templates", "Manage Templates", "View Wrappers", "Create Wrappers", "Manage Wrappers", "Manage Settings"], 
    "Clients & Services": ["View Clients", "Onboard Clients", "Manage Clients", "View Services", "Define Services", "Manage Services", "Manage Settings"], 
    "System Logs": 
    ["View Logs", "Manage Settings"] 
  });
  verifiedDb.getSheetByName("Roles").appendRow([new Date(), "R-ADMIN", "Administrator", "Unrestricted system access.", adminPerms, "Active", ""]);
  // 5. Asset Seeding
  seedCoreAssets(verifiedDb);
  
  SpreadsheetApp.flush();
  
  // Clean up temporary bootstrap keys from properties registry
  props.deleteProperty('TMP_FIRST_NAME');
  props.deleteProperty('TMP_LAST_NAME');
  
  // Emit Role Creation with the necessary email payload matching exact variable tokens
  SystemEvent.emit("Users:Roles", "CREATE", "Add Role", "INFO", "R-ADMIN", "Default Administrator role generated during installation.", adminEmail, { roleName: "Administrator", details: "Unrestricted system access." });
  // IMMUTABLE ANCHOR: Prevent Google Sheets concurrent write-locks from overwriting logs
  Utilities.sleep(1500);
  
  var initialDataContext = { 
    username: adminUsername, 
    firstName: finalFirst, 
    lastName: finalLast, 
    roleName: "Administrator",
    target_username: adminUsername,
    target_firstName: finalFirst,
    target_lastName: finalLast,
    target_role: "Administrator"
  };

  // Emit User Creation utilizing the concrete record anchor ID (U-1001) instead of raw username handles
  SystemEvent.emit("Users", "CREATE", "Add User", "INFO", "U-1001", "Master admin created.", adminEmail, initialPayloadContextMap || initialDataContext);
  Utilities.sleep(1500);
  // Emit dynamic Administrator Assignment event for the Master Admin on deployment
  SystemEvent.emit("Users", "ADMIN_ASSIGNED", "Admin Privileges Assigned", "WARN", "U-1001", "Master administrator account provisioned during setup.", adminEmail, initialPayloadContextMap || initialDataContext);
  
  if (verifiedDb.getSheetByName("Sheet1")) verifiedDb.deleteSheet(verifiedDb.getSheetByName("Sheet1"));
}

/**
 * Seeds the core email wrappers and baseline system templates into the database 
 * during the initial installation of the system.
 * 
 * @param {Spreadsheet} ss - The active Google Spreadsheet database object.
 */
function seedCoreAssets(ss) {
  var now = new Date();
  var wrapSheet = ss.getSheetByName("Wrappers");
  
  // Safely check if the sheet is empty or only contains headers (<= 1)
  if (wrapSheet && wrapSheet.getLastRow() <= 1) {
    // 100% White-labeled, generic system wrappers
    var intHtml = `<div style="background-color: #f8fafc; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;"><div style="background-color: #0f172a; padding: 25px; text-align: center; border-bottom: 4px solid #666DF2;"><img src="cid:logo" alt="System Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"></div><div style="padding: 30px; color: #334155; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; border-top: 1px solid #e2e8f0; background-color: #f8fafc; text-align: center; font-size: 11px; color: #64748b;">This is an automated system notification.</div></div></div>`;
    
    var extHtml = `<div style="background-color: #ffffff; padding: 40px 20px; font-family: Arial, sans-serif; border: 1px solid #e2e8f0;"><div style="max-width: 600px; margin: 0 auto;"><div style="padding-bottom: 20px; border-bottom: 1px solid #e2e8f0; margin-bottom: 20px; text-align: center;"><img src="cid:logo" alt="System Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"></div><div style="color: #334155; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="margin-top: 40px; font-size: 12px; color: #94a3b8; border-top: 1px solid #e2e8f0; padding-top: 15px;">Sent securely from our Workspace.</div></div></div>`;
    
    var userHtml = `<div style="background-color: #f8fafc; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 500px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px rgba(0,0,0,0.05);"><div style="padding: 30px; text-align: center; border-bottom: 1px solid #f1f5f9;"><img src="cid:logo" alt="System Logo" style="max-width: 120px; height: auto;"></div><div style="padding: 30px; color: #334155; line-height: 1.6; font-size: 15px;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; background-color: #f8fafc; text-align: center; font-size: 12px; color: #94a3b8; border-top: 1px solid #e2e8f0;">Security & Access Notification</div></div></div>`;
    
    wrapSheet.appendRow([now, "W-INTERNAL", "Internal Communication", "Standard internal messaging", intHtml, "Active"]);
    wrapSheet.appendRow([now, "W-EXTERNAL", "External Communication", "Client-facing messaging", extHtml, "Active"]);
    wrapSheet.appendRow([now, "W-USER", "User Communications", "Dedicated layout for user access and security emails", userHtml, "Active"]);
  }
  
  var tplSheet = ss.getSheetByName("Templates");
  
  // Safely check if the templates sheet is empty or only contains headers
  if (tplSheet && tplSheet.getLastRow() <= 1) {
    var welcomeHtml = `<div style='font-family: sans-serif;
padding: 20px;'><h2>Welcome to your Workspace</h2><p>Hello {{firstName}},</p><p>Your account is ready. You can now access your workspace using your system credentials.</p></div>`;
    
    var roleHtml = `<div style='font-family: sans-serif; padding: 20px;'><h2>New System Role</h2><p>The system role <strong>{{username}}</strong> has been successfully established.</p><p>Details: {{details}}</p></div>`;
    
    var resetHtml = `<div style='font-family: sans-serif; padding: 20px;'><h2>Password Reset Request</h2><p>Hi {{userFirst}},</p><p>We received a request to reset your local password. Click the link below to set a new password. This link will expire in 15 minutes.</p><a href='{{resetLink}}' style='display:inline-block; padding: 10px 20px; background: #666DF2; color: white; text-decoration: none; border-radius: 5px; margin-top: 15px;'>Reset Password</a></div>`;
    
    var updatedHtml = `<div style='font-family: sans-serif;
    padding: 20px;'><h2>Password Updated</h2><p>Hi {{userFirst}},</p><p>This is a confirmation that your system password has been successfully updated.
    If you did not make this change, please contact your administrator immediately.</p></div>`;
    
    var adminElevatedHtml = `<div style='font-family: sans-serif; padding: 20px;'><h2>Security Privilege Escalation</h2><p>Hello {{firstName}},</p><p>An account has been assigned Administrator privileges.</p><p><strong>Target User:</strong> {{target_firstName}} {{target_lastName}} (@{{target_username}})</p><p>If you did not authorized this elevation, audit system security settings instantly.</p></div>`;

    // Seed core templates mapping the Welcome template precisely onto User Communications wrapper
    tplSheet.appendRow([now, "TPL-USER-NEW", "User Welcome", "Access email", "Security", "Users", "Users:CREATE", "Welcome to the Workspace", welcomeHtml, "User Communications", "Active"]);
    tplSheet.appendRow([now, "TPL-ROLE-NEW", "Role Created", "Role creation alert", "Security", "Users:Roles", "Users:Roles:CREATE", "New System Role: {{username}}", roleHtml, "Internal Communication", "Active"]);
    tplSheet.appendRow([now, "TPL-ADMIN-NEW", "Administrator Elevation Alert", "Alerts when an administrator is set", "Security", "Users", "Users:ADMIN_ASSIGNED", "Security Alert: Administrator Privileges Assigned", adminElevatedHtml, "User Communications", "Active"]);
    
    // Updated wrappers for Password Templates -> "User Communications"
    tplSheet.appendRow([now, "TPL-PWD-RESET", "Password Reset Link", "Forgot password link", "Security", "Users", "Users:RESET_REQUEST", "Password Reset Request", resetHtml, "User Communications", "Active"]);
    
    tplSheet.appendRow([now, "TPL-PWD-UPDATE", "Password Updated", "Password change confirmation", "Security", "Users", "Users:PASSWORD_UPDATED", "Security Alert: Password Updated", updatedHtml, "User Communications", "Active"]);
  }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */