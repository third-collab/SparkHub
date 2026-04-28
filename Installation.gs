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
 * Initializes the Main Database with Core-only sheets (Users, Templates).
 */
/**
 * Initializes the Main Database with Core-only sheets (Users, Templates).
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
  
  // ROBUST FIX: Use the retry loop to move the file
  if (isNew) {
    moveFileWithRetry(ss.getId(), rootFolder);
  }

  // Core Schema: Users (Access Management Focus)
  var userHeaders = [
    "Timestamp", "Username", "Role", "Email", "Password", "First Name", "Last Name", "Status", "Last Login"
  ];
  initializeSheet(ss, "Users", userHeaders);

  // ==========================================
  // AUTO-CREATE INITIAL ADMINISTRATOR
  // ==========================================
  var usersSheet = ss.getSheetByName("Users");
  // Only append if the sheet is empty (meaning only the Header row exists)
  if (usersSheet && usersSheet.getLastRow() === 1) {
    var adminEmail = Session.getActiveUser().getEmail() || "admin@example.com";
    var adminUsername = adminEmail.split('@')[0]; // Auto-generate username from email
    
    usersSheet.appendRow([
      new Date(),           // Timestamp
      adminUsername,        // Username
      "Administrator",      // Role
      adminEmail,           // Email
      "",                   // Password (Leave blank to enforce Google SSO)
      "System",             // First Name placeholder
      "Admin",              // Last Name placeholder
      "Active",             // Status
      new Date()            // Last Login (Timestamp of installation)
    ]);
  }
  // ==========================================

  // Core Schema: Roles & Permissions Matrix
  var roleHeaders = [
    "Role ID", "Role Name", "Description", "Permissions JSON", "Status"
  ];
  initializeSheet(ss, "Roles", roleHeaders);
  
  var rolesSheet = ss.getSheetByName("Roles");
  if (rolesSheet && rolesSheet.getLastRow() === 1) {
    // Inject the Master Administrator role with a wildcard permission payload
    var adminPerms = JSON.stringify({ 
      "Core System": ["Manage Settings", "Manage Roles"], 
      "Access & Users": ["View Users", "Manage Users"], 
      "Templates": ["View Templates", "Manage Templates"] 
    });
    rolesSheet.appendRow(["R-ADMIN", "Administrator", "Unrestricted system access.", adminPerms, "Active"]);
  }

  // Core Schema: Templates
  var templateHeaders = [
    "ID", "Name", "Category", "Trigger", "Subject", "Body", "Status", "Wrapper"
  ];
  initializeSheet(ss, "Templates", templateHeaders);

  // Cleanup default Sheet1 if it exists
  var defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet) {
    ss.deleteSheet(defaultSheet);
  }
  
  PropertiesService.getScriptProperties().setProperty('DATABASE_ID', ss.getId());
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