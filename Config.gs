/**
 * Configuration Module - Backend
 * Standardized under MR Hub Architecture Blueprint.
 * Dynamically retrieves Database IDs from Script Properties.
 */

/**
 * Helper function to retrieve the main spreadsheet database object.
 * Fetches the ID from ScriptProperties.
 * Throws a specific error if the system has not been installed/configured.
 * @return {SpreadsheetApp.Spreadsheet} The main database spreadsheet object.
 */
function getMainDb() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('DATABASE_ID');
  
  if (!id) {
    throw new Error("Configuration Error: Main Database ID is missing. Please run the System Installation or configure the ID in Settings.");
  }
  
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error("System Error: Unable to open Main Database. Verify that the ID in Settings is correct and that the system has access to the file.");
  }
}

/**
 * Helper function to retrieve the notification spreadsheet database object.
 * Fetches the ID from ScriptProperties.
 * @return {SpreadsheetApp.Spreadsheet} The notification database spreadsheet object.
 */
function getNotifDb() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('NOTIF_DATABASE_ID');
  
  if (!id) {
    throw new Error("Configuration Error: Notification Database ID is missing. Please run the System Installation.");
  }
  
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error("System Error: Unable to open Notification Database. Verify that the file exists and is accessible.");
  }
}