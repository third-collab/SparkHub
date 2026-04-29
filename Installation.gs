/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Installation.gs
 * VERSION: 1.5 (Finalized Event Orchestration + Webhook + Governance)
 * SYNC STATUS: Fully Synchronized with EventBroker.gs & Utils.gs
 */

var SPARKHUB_VERSION = "1.0.0";
var MASTER_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbyderUFTDgJjjSb4ML2xpXzRnfKp_yBLkYlpaKdZcWZLowtmiutt-QZsg7OMq0enBJljw/exec";

/**
 * UI INSTALLATION HANDLER
 */
function performUiInstallation(data) {
  try {
    if (data.sysName.trim().toLowerCase() === 'sparkhub') {
      throw new Error("System name 'SparkHub' is restricted.");
    }

    var props = PropertiesService.getScriptProperties();
    var installerEmail = Session.getActiveUser().getEmail();
    var clientId = "CID-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    var instanceSecret = Utilities.getUuid();

    props.setProperty('ROOT_FOLDER_ID', data.rootId);
    props.setProperty('SYSTEM_NAME', data.sysName);
    props.setProperty('ADMIN_EMAIL', installerEmail);
    props.setProperty('CLIENT_ID', clientId);
    props.setProperty('INSTANCE_SECRET', instanceSecret);
    props.setProperty('SYSTEM_VERSION', SPARKHUB_VERSION);
    
    // 1. Run Physical & Schema setup
    runInstallation();
    setupSystemTriggers();

    // 2. ORCHESTRATION: System Level Trigger
    SystemEvent.emit("System", "INSTALL", "System Installation", "Core Architecture", "SparkHub system installed and databases initialized.", installerEmail);
    
    props.setProperty('ENVIRONMENT', 'Sandbox');

    // 3. MASTER WEBHOOK REPORTING
    if (MASTER_WEBHOOK_URL && MASTER_WEBHOOK_URL !== "") {
      var payload = {
        action: "install",
        secretKey: props.getProperty('WEBHOOK_SECRET') || "MISSING_KEY",
        instanceSecret: instanceSecret,
        date: new Date().toISOString(),
        clientId: clientId,
        clientName: data.sysName,
        clientEmail: installerEmail,
        version: SPARKHUB_VERSION,
        appUrl: ScriptApp.getService().getUrl(),
        databaseId: props.getProperty('DATABASE_ID')
      };
      try {
        UrlFetchApp.fetch(MASTER_WEBHOOK_URL, {
          method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true 
        });
      } catch (webhookError) { console.error("Webhook reporting failed: " + webhookError.message); }
    }

    return "Success|" + ScriptApp.getService().getUrl();
  } catch (e) {
    return "Error: " + e.message;
  }
}

function runInstallation() {
  var settings = getSystemSettings();
  var rootFolder = DriveApp.getFolderById(settings.rootFolderId);
  var assetsFolder = getOrCreateFolder(rootFolder, "System Assets");

  // Storage Partitions
  getOrCreateFolder(getOrCreateFolder(assetsFolder, "Settings"), "Images");
  getOrCreateFolder(getOrCreateFolder(assetsFolder, "Users"), "Photos");
  getOrCreateFolder(getOrCreateFolder(assetsFolder, "Templates"), "Emails");
  
  setupCoreDatabase(rootFolder);
  setupLogsDatabase(rootFolder);
}

function setupCoreDatabase(rootFolder) {
  var ss = SpreadsheetApp.create("SparkHub Database");
  moveFileWithRetry(ss.getId(), rootFolder);
  
  // 1. Governance Schemas
  initializeSheet(ss, "Users", ["Timestamp", "Username", "Role", "Email", "Password", "First Name", "Last Name", "Status", "Last Login"]);
  initializeSheet(ss, "Roles", ["Role ID", "Role Name", "Description", "Permissions JSON", "Status"]);
  initializeSheet(ss, "Templates", ["Timestamp", "ID", "Name", "Category", "Description", "Last Editor", "Trigger", "Subject", "Body", "Status", "Wrapper"]);
  initializeSheet(ss, "Wrappers", ["Timestamp", "Wrapper ID", "Name", "HTML Content", "Status"]);

  // 2. Admin Creation & DB Commit
  var adminEmail = Session.getActiveUser().getEmail();
  var adminUsername = adminEmail.split('@')[0];
  ss.getSheetByName("Users").appendRow([new Date(), adminUsername, "Administrator", adminEmail, "", "System", "Admin", "Active", new Date()]);
  
  // Default Admin Role
  var adminPerms = JSON.stringify({ "Core System": ["Manage Settings", "Manage Roles"], "Access & Users": ["View Users", "Manage Users"], "Templates": ["View Templates", "Manage Templates"], "System Logs": ["View Logs"] });
  ss.getSheetByName("Roles").appendRow(["R-ADMIN", "Administrator", "Unrestricted system access.", adminPerms, "Active"]);
  
  SpreadsheetApp.flush();

  // 3. ORCHESTRATION: User Level Trigger
  SystemEvent.emit("Users", "CREATE", "Add User", adminUsername, "Master Administrator profile auto-generated during installation.", adminEmail);

  // 4. Asset Seeding
  seedCoreAssets(ss);

  if (ss.getSheetByName("Sheet1")) ss.deleteSheet(ss.getSheetByName("Sheet1"));
  PropertiesService.getScriptProperties().setProperty('DATABASE_ID', ss.getId());
}

function seedCoreAssets(ss) {
  var now = new Date();
  var wrapSheet = ss.getSheetByName("Wrappers");
  if (wrapSheet.getLastRow() === 1) {
    // Professional Wrappers
    var intHtml = `<div style="background-color: #f4f6f9; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;"><div style="background-color: #323232; padding: 25px; text-align: center; border-bottom: 4px solid #F1C40F;"><img src="cid:logo" alt="MegaRhino Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"><h1 style="color: #ffffff; margin: 0; font-size: 20px;">MegaRhino</h1></div><div style="padding: 30px; color: #444; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; border-top: 1px solid #eee; background-color: #fcfcfc; text-align: center; font-size: 11px; color: #888;">This is an automated system notification.</div></div></div>`;
    var extHtml = `<div style="background-color: #ffffff; padding: 40px 20px; font-family: Arial, sans-serif; border: 1px solid #eee;"><div style="max-width: 600px; margin: 0 auto;"><div style="padding-bottom: 20px; border-bottom: 1px solid #ddd; margin-bottom: 20px; text-align: center;"><img src="cid:logo" alt="MegaRhino Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"><h2 style="color: #333; margin: 0;">MegaRhino</h2></div><div style="color: #555; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="margin-top: 40px; font-size: 12px; color: #999; border-top: 1px solid #eee; padding-top: 15px;">Sent from the MegaRhino Team.</div></div></div>`;
    wrapSheet.appendRow([now, "W-INTERNAL", "Internal Hub", intHtml, "Active"]);
    wrapSheet.appendRow([now, "W-EXTERNAL", "External Client", extHtml, "Active"]);
  }

  var tplSheet = ss.getSheetByName("Templates");
  if (tplSheet.getLastRow() === 1) {
    // Initial Communication Logic with Event Handles
    tplSheet.appendRow([now, "TPL-INSTALL", "System Installed", "System", "Admin alert", "Installer", "System:INSTALL", "Installation Successful", "<p>SparkHub has been deployed.</p>", "Active", "Internal Hub"]);
    tplSheet.appendRow([now, "TPL-WELCOME", "User Welcome", "Security", "Account access email", "Installer", "Users:CREATE", "Welcome to SparkHub", "<p>Hello {{username}}, your account is ready.</p>", "Active", "Internal Hub"]);
  }
}

function setupLogsDatabase(rootFolder) {
  var ss = SpreadsheetApp.create("SparkHub Logs Database");
  moveFileWithRetry(ss.getId(), rootFolder);
  initializeSheet(ss, "System Logs", ["Timestamp", "Module", "Action Type", "Action Name", "Severity", "Actor", "Target Entity", "Log Details", "Environment"]);
  if (ss.getSheetByName("Sheet1")) ss.deleteSheet(ss.getSheetByName("Sheet1"));
  PropertiesService.getScriptProperties().setProperty('LOGS_DATABASE_ID', ss.getId());
}