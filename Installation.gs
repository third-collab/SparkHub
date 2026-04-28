/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Installation.gs
 * VERSION: 1.2 (Registry-Based Modular Installation)
 * SYNC STATUS: Fully Synchronized with Settings.gs & Master_Succession.md
 */

/**
 * Installation Module - Backend
 * Standardized under SparkHub Architecture Blueprint.
 * * CORE RESPONSIBILITIES:
 * - Orchestrates the "Core" system installation (Folders, Users DB, Templates DB).
 * - Provides dynamic entry points for decoupled Add-on installations.
 * - Manages physical storage hierarchy in Google Drive.
 */
// Establish the system version identifier
var SPARKHUB_VERSION = "1.0.0";
var MASTER_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbyderUFTDgJjjSb4ML2xpXzRnfKp_yBLkYlpaKdZcWZLowtmiutt-QZsg7OMq0enBJljw/exec"; 

/**
 * UI INSTALLATION HANDLER
 * Triggered by the Installation Wizard in Index.html.
 */
function performUiInstallation(data) {
  try {
    // Backend Validation: Prevent "SparkHub"
    if (data.sysName.trim().toLowerCase() === 'sparkhub') {
      throw new Error("System name 'SparkHub' is restricted. Please choose another name.");
    }

    var props = PropertiesService.getScriptProperties();
    var installerEmail = Session.getActiveUser().getEmail();
    
    // Auto-generate unique Identifiers
    var clientId = "CID-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    var instanceSecret = Utilities.getUuid(); // The secure instance key

    props.setProperty('ROOT_FOLDER_ID', data.rootId);
    props.setProperty('SYSTEM_NAME', data.sysName);
    props.setProperty('ADMIN_EMAIL', installerEmail);
    props.setProperty('CLIENT_ID', clientId);
    props.setProperty('INSTANCE_SECRET', instanceSecret); // Save locally
    props.setProperty('SYSTEM_VERSION', SPARKHUB_VERSION);
    
    // Executes Core and Logs infrastructure
    runInstallation();
    
    // Set up the automated daily triggers immediately after install
    setupSystemTriggers();
    
    // ==========================================
    // INITIAL SYSTEM LOGS
    // Now that both DBs are built and IDs are saved, it is safe to log!
    // ==========================================
    var adminUsername = installerEmail.split('@')[0];
    logSystemAction("System", "SYSTEM", "System Installation", "INFO", "Core Architecture", "SparkHub system installed and databases initialized.");
    logSystemAction("Users", "CREATE", "Add User", "INFO", adminUsername, "Master Administrator profile auto-generated during installation.");
    // ==========================================
    
    props.setProperty('ENVIRONMENT', 'Sandbox');

    // FIRE POST REQUEST TO MASTER REGISTRY
    if (MASTER_WEBHOOK_URL !== "https://script.google.com/macros/s/AKfycbyderUFTDgJjjSb4ML2xpXzRnfKp_yBLkYlpaKdZcWZLowtmiutt-QZsg7OMq0enBJljw/exec") {
      var payload = {
        action: "install", // Tell the webhook this is a new setup
        secretKey: props.getProperty('WEBHOOK_SECRET') || "MISSING_KEY",
        instanceSecret: instanceSecret, // Send the key for the master sheet to store
        date: new Date().toISOString(),
        clientId: clientId,
        clientName: data.sysName,
        clientEmail: installerEmail,
        version: SPARKHUB_VERSION,
        appUrl: ScriptApp.getService().getUrl(),
        timezone: Session.getScriptTimeZone(),
        rootId: data.rootId,
        databaseId: props.getProperty('DATABASE_ID'),
        status: "Active",
        lastPing: new Date().toISOString()
      };
      
      try {
        UrlFetchApp.fetch(MASTER_WEBHOOK_URL, {
          method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true 
        });
      } catch (webhookError) { console.error("Webhook reporting failed: " + webhookError.message); }
    }
    // ==========================================

    return "Success|" + ScriptApp.getService().getUrl();
  } catch (e) {
    return "Error: " + e.message;
  }
}

/**
 * CORE INSTALLATION
 * Sets up the minimum viable infrastructure for the system to boot.
 * Creates only the Users and Templates sheets.
 */
function runInstallation() {
  var settings = getSystemSettings();
  var rootId = settings.rootFolderId;
  if (!rootId) throw new Error("INSTALLATION HALTED: Root Folder ID not found.");
  
  try {
    var rootFolder = DriveApp.getFolderById(rootId);
    var assetsFolder = getOrCreateFolder(rootFolder, "System Assets");

    // Core Folder Partitions
    getOrCreateFolder(getOrCreateFolder(assetsFolder, "Settings"), "Images");
    getOrCreateFolder(getOrCreateFolder(assetsFolder, "Users"), "Photos");
    getOrCreateFolder(getOrCreateFolder(assetsFolder, "Templates"), "Emails");
    
    // Core Database Initialization (Users & Templates Only)
    setupCoreDatabase(rootFolder);
    
    // Dedicated Logs Database Initialization
    setupLogsDatabase(rootFolder);
    
    console.log("SUCCESS: Core infrastructure ready.");
  } catch (e) {
    throw new Error("Core installation failed: " + e.message);
  }
}

/**
 * Initializes the Main Database with Core-only sheets (Users, Templates, Wrappers, Roles).
 * Patterned after: https://docs.google.com/spreadsheets/d/1UwEG4lY7Gs8BK_5NkdEqAicyZ_pyBpfwg2CEyLAd2qI/
 */
function setupCoreDatabase(rootFolder) {
  var dbName = "SparkHub Database";
  var files = rootFolder.getFilesByName(dbName);
  var ss;
  var isNew = false;
  
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(dbName);
    isNew = true;
  }
  
  if (isNew) {
    moveFileWithRetry(ss.getId(), rootFolder);
  }

  // 1. Schema: Users
  initializeSheet(ss, "Users", ["Timestamp", "Username", "Role", "Email", "Password", "First Name", "Last Name", "Status", "Last Login"]);

  // Auto-create Initial Admin
  var usersSheet = ss.getSheetByName("Users");
  if (usersSheet.getLastRow() === 1) {
    var adminEmail = Session.getActiveUser().getEmail();
    usersSheet.appendRow([new Date(), adminEmail.split('@')[0], "Administrator", adminEmail, "", "System", "Admin", "Active", new Date()]);
  }

  // 2. Schema: Roles
  initializeSheet(ss, "Roles", ["Role ID", "Role Name", "Description", "Permissions JSON", "Status"]);
  var rolesSheet = ss.getSheetByName("Roles");
  if (rolesSheet.getLastRow() === 1) {
    var adminPerms = JSON.stringify({ "Core System": ["Manage Settings", "Manage Roles"], "Access & Users": ["View Users", "Manage Users"], "Templates": ["View Templates", "Manage Templates"], "System Logs": ["View Logs"] });
    rolesSheet.appendRow(["R-ADMIN", "Administrator", "Unrestricted system access.", adminPerms, "Active"]);
  }

  // 3. Schema: Templates (11-Column Schema with Timestamp)
  var templateHeaders = ["Timestamp", "ID", "Name", "Category", "Description", "Last Editor", "Trigger", "Subject", "Body", "Status", "Wrapper"];
  initializeSheet(ss, "Templates", templateHeaders);

  // 4. Schema: Wrappers (5-Column Schema with Timestamp)
  var wrapperHeaders = ["Timestamp", "Wrapper ID", "Name", "HTML Content", "Status"];
  initializeSheet(ss, "Wrappers", wrapperHeaders);

  // Seed provided assets [cite: 421, 424]
  seedCoreAssets(ss);

  if (ss.getSheetByName("Sheet1")) ss.deleteSheet(ss.getSheetByName("Sheet1"));
  PropertiesService.getScriptProperties().setProperty('DATABASE_ID', ss.getId());
}

/**
 * Seeds the provided Wrapper HTML and initial Templates into the Registry.
 */
function seedCoreAssets(ss) {
  var wrapSheet = ss.getSheetByName("Wrappers");
  if (wrapSheet.getLastRow() === 1) {
    var internalHtml = `<div style="background-color: #f4f6f9; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;"><div style="background-color: #323232; padding: 25px; text-align: center; border-bottom: 4px solid #F1C40F;"><img src="cid:logo" alt="MegaRhino Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"><h1 style="color: #ffffff; margin: 0; font-size: 20px;">MegaRhino</h1></div><div style="padding: 30px; color: #444; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; border-top: 1px solid #eee; background-color: #fcfcfc; text-align: center; font-size: 11px; color: #888;">This is an automated system notification.<br>Please do not reply to this email.</div></div></div>`; // [cite: 421, 422, 423]
    
    var externalHtml = `<div style="background-color: #ffffff; padding: 40px 20px; font-family: Arial, sans-serif; border: 1px solid #eee;"><div style="max-width: 600px; margin: 0 auto;"><div style="padding-bottom: 20px; border-bottom: 1px solid #ddd; margin-bottom: 20px; text-align: center;"><img src="cid:logo" alt="MegaRhino Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"><h2 style="color: #333; margin: 0;">MegaRhino</h2></div><div style="color: #555; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="margin-top: 40px; font-size: 12px; color: #999; border-top: 1px solid #eee; padding-top: 15px;">Sent from the MegaRhino Team.<br><span style="font-size: 11px;">Please do not reply to this email.</span></div></div></div>`; // [cite: 424, 425]

    wrapSheet.appendRow([new Date(), "W-INTERNAL", "Internal Hub", internalHtml, "Active"]);
    wrapSheet.appendRow([new Date(), "W-EXTERNAL", "External Client", externalHtml, "Active"]);
  }

  var tplSheet = ss.getSheetByName("Templates");
  if (tplSheet.getLastRow() === 1) {
    // Initial System Template
    tplSheet.appendRow([
      new Date(), "TPL-WELCOME", "System Welcome", "Security", "Initial access email", 
      "Installer", "User:CREATE", "Welcome to {{systemName}}", 
      "<p>Hello {{username}},</p><p>Your account has been established.</p>", "Active", "Internal Hub"
    ]);
  }
}

/**
 * Initializes a strictly dedicated Database for System Logs.
 */
function setupLogsDatabase(rootFolder) {
  var dbName = "SparkHub Logs Database";
  var files = rootFolder.getFilesByName(dbName);
  var ss;
  var isNew = false;
  
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(dbName);
    isNew = true;
  }
  
  // ROBUST FIX: Use the retry loop to move the file
  if (isNew) {
    moveFileWithRetry(ss.getId(), rootFolder);
  }

  var logHeaders = [
    "Timestamp", "Module", "Action Type", "Action Name", "Severity", "Actor", "Target Entity", "Log Details", "Environment"
  ];
  initializeSheet(ss, "System Logs", logHeaders);
  
  // Cleanup default Sheet1 if it exists
  var defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet) {
    ss.deleteSheet(defaultSheet);
  }
  
  PropertiesService.getScriptProperties().setProperty('LOGS_DATABASE_ID', ss.getId());
}

/* ========================================================================
   GLOBAL HELPERS
   ======================================================================== */

/**
 * Safely moves a file to a target folder using a retry loop.
 * Bypasses Google Drive API indexing latency.
 */
function moveFileWithRetry(fileId, targetFolder) {
  var maxRetries = 4;
  for (var i = 0; i < maxRetries; i++) {
    try {
      var file = DriveApp.getFileById(fileId);
      file.moveTo(targetFolder);
      return; // Success, exit the loop
    } catch (e) {
      if (i === maxRetries - 1) {
        throw new Error("Drive indexing timeout. File created but could not be moved.");
      }
      Utilities.sleep(2000); // Wait 2 seconds and try again
    }
  }
}

/**
 * Helper to get or create a folder within a parent.
 */
function getOrCreateFolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  var newFolder = parent.createFolder(name);
  newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return newFolder;
}

/**
 * Helper to initialize a sheet with bold headers and frozen top row.
 */
function initializeSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setFontWeight("bold")
         .setBackground("#F3F3F3");
    sheet.setFrozenRows(1);
  }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */