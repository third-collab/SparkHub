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

    props.setProperty('ROOT_FOLDER_ID', data.rootId);
    props.setProperty('SYSTEM_NAME', data.sysName);
    props.setProperty('ADMIN_EMAIL', installerEmail);
    props.setProperty('CLIENT_ID', "CID-" + Utilities.getUuid().substring(0, 8).toUpperCase());
    props.setProperty('INSTANCE_SECRET', Utilities.getUuid());
    
    // NEW LOGIC: Automatically register the integrated UserBar plugin
    props.setProperty('INSTALLED_PLUGINS', 'UserBar');

    runInstallation();
    setupSystemTriggers();

    // ORCHESTRATION: Trigger 1
    SystemEvent.emit("System", "INSTALL", "System Installation", "WARN", "Core Architecture", "SparkHub core deployed.", installerEmail);
    props.setProperty('ENVIRONMENT', 'Sandbox');

    // IMMUTABLE ANCHOR: Master Webhook Reporting
    // IMMUTABLE ANCHOR: Master Webhook Reporting
    if (MASTER_WEBHOOK_URL) {
      var payload = { 
        action: "install", clientId: props.getProperty('CLIENT_ID'), 
        clientName: data.sysName, clientEmail: installerEmail, databaseId: props.getProperty('DATABASE_ID') 
      };
      UrlFetchApp.fetch(MASTER_WEBHOOK_URL, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
    }
    
    // NEW LOGIC: Hardcoded Installation Email Pipeline
    try {
      var sysLogo = "https://i.imgur.com/0iPmgVk.png"; // SparkHub Logo fallback
      var installHtml = `<div style="background-color: #f4f6f9; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;"><div style="background-color: #323232; padding: 25px; text-align: center; border-bottom: 4px solid #666DF2;"><img src="${sysLogo}" alt="SparkHub Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"></div><div style="padding: 30px; color: #444; line-height: 1.6;"><div style='font-family: sans-serif; padding: 20px;'><h2>Installation Successful</h2><p>SparkHub has been fully deployed and is ready for use as <strong>${data.sysName}</strong>.</p></div></div><div style="padding: 20px; border-top: 1px solid #eee; background-color: #fcfcfc; text-align: center; font-size: 11px; color: #888;">This is an automated system notification from SparkHub.</div></div></div>`;
      MailApp.sendEmail({
        to: installerEmail,
        subject: "Installation Successful",
        htmlBody: installHtml,
        noReply: true,
        name: "SparkHub Setup"
      });
    } catch(mailErr) {
      console.warn("Failed to send installation email: " + mailErr.message);
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
  
  // SEQUENCING FIX: Create Logs first so that SystemEvent.emit can record the core installation
  setupLogsDatabase(rootFolder);
  setupCoreDatabase(rootFolder);
}

function setupLogsDatabase(rootFolder) {
  var ss = SpreadsheetApp.create("SparkHub Logs Database");
  var dbId = ss.getId();
  
  // 1. IMMEDIATE SAVE: Fixes the missing ID error
  PropertiesService.getScriptProperties().setProperty('LOGS_DATABASE_ID', dbId);
  
  // 2. DATABASE INTEGRITY LOOP
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
  var ss = SpreadsheetApp.create("SparkHub Database");
  var dbId = ss.getId();
  
  // 1. IMMEDIATE SAVE: Fixes the missing ID error before any triggers fire
  PropertiesService.getScriptProperties().setProperty('DATABASE_ID', dbId);
  
  // 2. DATABASE INTEGRITY LOOP
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

  // 3. Schema Initialization
  initializeSheet(verifiedDb, "Users", ["Timestamp", "Username", "Role", "Email", "Password", "First Name", "Last Name", "Status", "Last Login"]);
  initializeSheet(verifiedDb, "Roles", ["Role ID", "Role Name", "Description", "Permissions JSON", "Status"]);
  initializeSheet(verifiedDb, "Templates", ["Timestamp", "ID", "Name", "Category", "Description", "Trigger", "Subject", "Body", "Status", "Wrapper"]);
  initializeSheet(verifiedDb, "Wrappers", ["Timestamp", "Wrapper ID", "Name", "HTML Content", "Status"]);

  // 4. Admin Creation
  var adminEmail = Session.getActiveUser().getEmail();
  var adminUsername = adminEmail.split('@')[0];
  verifiedDb.getSheetByName("Users").appendRow([new Date(), adminUsername, "Administrator", adminEmail, "", "System", "Admin", "Active", new Date()]);
  
  var adminPerms = JSON.stringify({ "Core System": ["Manage Settings", "Manage Roles"], "Access & Users": ["View Users", "Manage Users"], "Templates": ["View Templates", "Manage Templates"], "System Logs": ["View Logs"] });
  verifiedDb.getSheetByName("Roles").appendRow(["R-ADMIN", "Administrator", "Unrestricted system access.", adminPerms, "Active"]);
  
  // 5. Asset Seeding (Happens BEFORE the trigger)
  seedCoreAssets(verifiedDb);
  
  SpreadsheetApp.flush();

  // 6. ORCHESTRATION: Trigger 2
  // Because the ID is saved, the DB is verified, and the Template is seeded, this will succeed.
  SystemEvent.emit("Roles", "CREATE", "Add Role", "INFO", "Administrator", "Default Administrator role generated during installation.", adminEmail);
  SystemEvent.emit("Users", "CREATE", "Add User", "INFO", adminUsername, "Master admin created.", adminEmail);
  
  if (verifiedDb.getSheetByName("Sheet1")) verifiedDb.deleteSheet(verifiedDb.getSheetByName("Sheet1"));
}

function seedCoreAssets(ss) {
  var now = new Date();
  var wrapSheet = ss.getSheetByName("Wrappers");
  
  if (wrapSheet.getLastRow() === 1) {
    var intHtml = `<div style="background-color: #f4f6f9; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;"><div style="background-color: #323232; padding: 25px; text-align: center; border-bottom: 4px solid #F1C40F;"><img src="cid:logo" alt="MegaRhino Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"><h1 style="color: #ffffff; margin: 0; font-size: 20px;">MegaRhino</h1></div><div style="padding: 30px; color: #444; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; border-top: 1px solid #eee; background-color: #fcfcfc; text-align: center; font-size: 11px; color: #888;">This is an automated system notification.</div></div></div>`;
    var extHtml = `<div style="background-color: #ffffff; padding: 40px 20px; font-family: Arial, sans-serif; border: 1px solid #eee;"><div style="max-width: 600px; margin: 0 auto;"><div style="padding-bottom: 20px; border-bottom: 1px solid #ddd; margin-bottom: 20px; text-align: center;"><img src="cid:logo" alt="MegaRhino Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"><h2 style="color: #333; margin: 0;">MegaRhino</h2></div><div style="color: #555; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="margin-top: 40px; font-size: 12px; color: #999; border-top: 1px solid #eee; padding-top: 15px;">Sent from the MegaRhino Team.</div></div></div>`;
    var userHtml = `<div style="background-color: #f8fafc; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 500px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px rgba(0,0,0,0.05);"><div style="padding: 30px; text-align: center; border-bottom: 1px solid #f1f5f9;"><img src="cid:logo" alt="Logo" style="max-width: 120px; height: auto;"></div><div style="padding: 30px; color: #334155; line-height: 1.6; font-size: 15px;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; background-color: #f8fafc; text-align: center; font-size: 12px; color: #94a3b8; border-top: 1px solid #e2e8f0;">Security & Access Notification</div></div></div>`;

    // Strict array mapping: TS, ID, Name, Description, HTML, Status
    wrapSheet.appendRow([now, "W-INTERNAL", "Internal Communication", "Standard internal messaging", intHtml, "Active"]);
    wrapSheet.appendRow([now, "W-EXTERNAL", "External Communication", "Client-facing messaging", extHtml, "Active"]);
    wrapSheet.appendRow([now, "W-USER", "User Communications", "Dedicated layout for user access and security emails", userHtml, "Active"]);
    
    SystemEvent.emit("Templates", "CREATE", "Seed Wrapper", "INFO", "Internal Communication", "Default internal wrapper seeded.");
    SystemEvent.emit("Templates", "CREATE", "Seed Wrapper", "INFO", "External Communication", "Default external client wrapper seeded.");
    SystemEvent.emit("Templates", "CREATE", "Seed Wrapper", "INFO", "User Communications", "Dedicated layout for user access and security emails seeded.");
  }
  
  var tplSheet = ss.getSheetByName("Templates");
  if (tplSheet.getLastRow() === 1) {
    // Define the styled HTML for the core templates
    var installHtml = `<div style='font-family: sans-serif; padding: 20px;'><h2>Installation Successful</h2><p>SparkHub has been fully deployed and is ready for use.</p></div>`;
    var welcomeHtml = `<div style='font-family: sans-serif; padding: 20px;'><h2>Welcome to SparkHub</h2><p>Hello {{username}},</p><p>Your account is ready. You can now access your workspace using your system credentials.</p></div>`;

    tplSheet.appendRow([now, "TPL-INSTALL", "System Installed", "System", "Admin alert", "System:INSTALL", "Installation Successful", installHtml, "Active", "Internal Hub"]);
    tplSheet.appendRow([now, "TPL-USER-NEW", "User Welcome", "Security", "Access email", "Users:CREATE", "Welcome to SparkHub", welcomeHtml, "Active", "Internal Hub"]);
    
    // NEW TEMPLATE: Role Creation Alert (Mapped to 10-column schema)
    var roleHtml = `<div style='font-family: sans-serif; padding: 20px;'><h2>New System Role</h2><p>The system role <strong>{{username}}</strong> has been successfully established.</p><p>Details: {{details}}</p></div>`;
    tplSheet.appendRow([now, "TPL-ROLE-NEW", "Role Created", "Security", "Role creation alert", "Roles:CREATE", "New System Role: {{username}}", roleHtml, "Active", "Internal Hub"]);
    
    SystemEvent.emit("Templates", "CREATE", "Seed Template", "INFO", "User Welcome", "User welcome template seeded.");
    SystemEvent.emit("Templates", "CREATE", "Seed Template", "INFO", "Role Created", "Role creation notification template seeded.");

    // NEW TEMPLATES: Password Control Flow
    var resetHtml = `<div style='font-family: sans-serif; padding: 20px;'><h2>Password Reset Request</h2><p>Hi {{userFirst}},</p><p>We received a request to reset your local password. Click the link below to set a new password. This link will expire in 15 minutes.</p><a href='{{resetLink}}' style='display:inline-block; padding: 10px 20px; background: #c40004; color: white; text-decoration: none; border-radius: 5px; margin-top: 15px;'>Reset Password</a></div>`;
    var updatedHtml = `<div style='font-family: sans-serif; padding: 20px;'><h2>Password Updated</h2><p>Hi {{userFirst}},</p><p>This is a confirmation that your system password has been successfully updated. If you did not make this change, please contact your administrator immediately.</p></div>`;

    tplSheet.appendRow([now, "TPL-PWD-RESET", "Password Reset Link", "Security", "Forgot password link", "Users:RESET_REQUEST", "Password Reset Request", resetHtml, "Active", "External Client"]);
    tplSheet.appendRow([now, "TPL-PWD-UPDATE", "Password Updated", "Security", "Password change confirmation", "Users:PASSWORD_UPDATED", "Security Alert: Password Updated", updatedHtml, "Active", "External Client"]);
    
    SystemEvent.emit("Templates", "CREATE", "Seed Template", "INFO", "Password Reset Link", "Password reset request template seeded.");
    SystemEvent.emit("Templates", "CREATE", "Seed Template", "INFO", "Password Updated", "Password updated confirmation template seeded.");
  }
}