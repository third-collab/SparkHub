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
 * - Provides modular entry points for "Add-on" installations (Clients, Notifications, Calendar).
 * - Manages physical storage hierarchy in Google Drive.
 */

/**
 * UI INSTALLATION HANDLER (CORE ONLY)
 * Triggered by the Installation Wizard in Index.html.
 * Sets the system to "Sandbox" by default upon first install.
 */
function performUiInstallation(data) {
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('ROOT_FOLDER_ID', data.rootId);
    props.setProperty('SYSTEM_NAME', data.sysName || 'SparkHub');
    props.setProperty('ADMIN_EMAIL', Session.getActiveUser().getEmail());
    
    // Executes only Core infrastructure
    runInstallation();
    
    props.setProperty('ENVIRONMENT', 'Sandbox');
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
    
    console.log("SUCCESS: Core infrastructure ready.");
  } catch (e) {
    throw new Error("Core installation failed: " + e.message);
  }
}

/**
 * Initializes the Main Database with Core-only sheets (Users, Templates).
 */
function setupCoreDatabase(rootFolder) {
  var dbName = "SparkHub Database";
  var files = rootFolder.getFilesByName(dbName);
  var ss = files.hasNext() ? SpreadsheetApp.open(files.next()) : SpreadsheetApp.create(dbName);
  
  if (!files.hasNext()) {
    DriveApp.getFileById(ss.getId()).moveTo(rootFolder);
  }

  // Core Schema: Users
  var userHeaders = [
    "Timestamp", "Username", "Role", "Work Email", "First Name", "Last Name", 
    "Birthday", "Personal Email", "Phone", "Address", "Facebook URL", 
    "Profile Photo URL", "Position", "Employment Type", "Date Hired", "Status"
  ];
  initializeSheet(ss, "Users", userHeaders);

  // Core Schema: Templates
  var templateHeaders = [
    "ID", "Name", "Category", "Trigger", "Subject", "Body", "Status", "Wrapper"
  ];
  initializeSheet(ss, "Templates", templateHeaders);
  
  PropertiesService.getScriptProperties().setProperty('DATABASE_ID', ss.getId());
}

/* ========================================================================
   ADD-ON MODULE INSTALLERS (Triggered via Settings Registry)
   ======================================================================== */

/**
 * Installs the Clients Module.
 * Creates the 'Clients' sheet in the Main DB and the 'Clients' Asset folder.
 */
function installClients() {
  var settings = getSystemSettings();
  var ss = SpreadsheetApp.openById(settings.mainDbId);
  var rootFolder = DriveApp.getFolderById(settings.rootFolderId);
  var assetsFolder = getOrCreateFolder(rootFolder, "System Assets");

  // 1. Create Sheet with the full 39-column Architecture
  var headers = [
    "Timestamp", "Company Name", "Brand Name", "Address", "Website", "Anniversary", 
    "P-FirstName", "P-LastName", "P-Email", "P-Phone", "P-Address", "P-Birthday", 
    "P-Position", "Add-Contacts", "Ship-Address", "Ret-Address", "Brand-Reg", 
    "Services", "Monthly-Val", "Orig-Start", "Curr-Start", "Term-Count", 
    "Term-Unit", "Orig-Exp", "Curr-Exp", "Orig-End", "Latest-End", "Comm-Rate", 
    "Comm-Basis", "PPC-Date", "DSP-Date", "Handover-Notes", "Op-Notes", 
    "Brand-Code", "Status", "Acct-Mgr", "Folder-URL", "Remarks", "History"
  ];
  initializeSheet(ss, "Clients", headers);

  // 2. Create Storage
  getOrCreateFolder(assetsFolder, "Clients");
  return true;
}

/**
 * Installs the Notifications Module.
 * Creates a dedicated Notifications Database file.
 */
function installNotifications() {
  var settings = getSystemSettings();
  var rootFolder = DriveApp.getFolderById(settings.rootFolderId);
  var dbName = "SparkHub Notifications";
  
  var ss = SpreadsheetApp.create(dbName);
  DriveApp.getFileById(ss.getId()).moveTo(rootFolder);
  
  var headers = [
    "Timestamp", "Module", "Event", "Message", "Recipient", "IsRead"
  ];
  initializeSheet(ss, "Notifications", headers);
  
  PropertiesService.getScriptProperties().setProperty('NOTIF_DATABASE_ID', ss.getId());
  return true;
}

/**
 * Installs the Calendar Module.
 */
function installCalendar() {
  var settings = getSystemSettings();
  var ss = SpreadsheetApp.openById(settings.mainDbId);
  var rootFolder = DriveApp.getFolderById(settings.rootFolderId);
  var assetsFolder = getOrCreateFolder(rootFolder, "System Assets");

  // 1. Create Sheet
  var headers = [
    "Timestamp", "Event Name", "Date", "Start Time", "End Time", "Description", "Assigned To"
  ];
  initializeSheet(ss, "Calendar", headers);

  // 2. Create Storage
  getOrCreateFolder(assetsFolder, "Calendar");
  return true;
}

/* ========================================================================
   GLOBAL HELPERS
   ======================================================================== */

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