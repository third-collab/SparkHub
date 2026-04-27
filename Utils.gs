/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Utils.gs
 * VERSION: 1.2 (Generalized Base64 Upload Engine)
 * SYNC STATUS: Fully Synchronized with SystemsGovernance.md
 */

/**
 * Utility Module - Backend
 * Standardized under SparkHub Architecture Blueprint.
 * Handles: File uploads, subfolder management, and global template helpers.
 */

/**
 * Server-side helper to include separate HTML files into the master template.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Standardized helper to find or create a folder.
 * Centralized here to support Installation.gs, Settings.gs, and Clients.gs.
 */
function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    var newFolder = parentFolder.createFolder(folderName);
    // Ensure the folder is viewable by the system for UI rendering (e.g. logo/photos)
    newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return newFolder;
  }
}

/**
 * Dynamically retrieves a subfolder within the configured Root Folder.
 * Utilizes getOrCreateFolder to ensure structural integrity.
 */
function getSystemSubfolder(subfolderName) {
  var settings = getSystemSettings();
  var rootId = settings.rootFolderId;
  
  if (!rootId) {
    throw new Error("System Configuration Error: Root Folder is not set in Settings.");
  }
  
  var rootFolder = DriveApp.getFolderById(rootId);
  return getOrCreateFolder(rootFolder, subfolderName);
}

/**
 * GENERIC UPLOAD ENGINE: Decodes a base64 string and saves it to a specified folder.
 * This is the standardized function for all system file uploads (Logos, Photos, Docs).
 * * @param {string} base64 - The data URI or raw base64 string.
 * @param {string} filename - The name to save the file as.
 * @param {GoogleAppsScript.Drive.Folder} folderObj - The target Drive Folder object.
 * @return {Object} Contains the permanent Drive URL and File ID.
 */
function uploadBase64File(base64, filename, folderObj) {
  try {
    var base64Data = base64.split(',')[1] || base64;
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, 'image/png', filename);
    
    var file = folderObj.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      url: file.getUrl(),
      id: file.getId()
    };
  } catch (e) {
    console.error("Generic upload error: " + e.message);
    throw new Error("File Upload Failed: " + e.message);
  }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */