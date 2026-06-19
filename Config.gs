/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Config.gs
 * VERSION: 1.1
 * SYNC STATUS: Fully Synchronized with Settings.gs
 */

/**
 * Configuration Module - Backend
 * Standardized under SparkHub Architecture Blueprint.
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
 * Helper function to retrieve the dedicated Logs database object.
 * @return {SpreadsheetApp.Spreadsheet} The logs database spreadsheet object.
 */
function getLogsDb() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('LOGS_DATABASE_ID');
  if (!id) {
    throw new Error("Configuration Error: Logs Database ID is missing.");
  }
  
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error("System Error: Unable to open Logs Database.");
  }
}

/**
 * Centralized Ecosystem Lookup Aggregator
 * Automatically discovers and harvests all lookable entity reference models 
 * across core infrastructure and dynamically installed external extensions.
 * Follows the V8 Auto-Discovery pattern.
 */
function getSystemDynamicLookups() {
  var masterLookups = { users: [], roles: [], templates: [], wrappers: [], clients: [], services: [] };
  var activeModules = ["System", "Users", "Templates", "Settings", "Logs"];
  var props = PropertiesService.getScriptProperties();
  var installed = props.getProperty('INSTALLED_MODULES');
  
  if (installed) {
    installed.split(',').forEach(function(m) {
      var mod = m.trim();
      if (mod && activeModules.indexOf(mod) === -1) activeModules.push(mod);
    });
  }
  
  var globalScope = typeof globalThis !== 'undefined' ? globalThis : this;
  
  activeModules.forEach(function(modName) {
    var funcName = modName + "_getLookups";
    if (typeof globalScope[funcName] === 'function') {
      try {
        var modulePayload = globalScope[funcName]();
        if (modulePayload && typeof modulePayload === 'object') {
          for (var key in modulePayload) {
            masterLookups[key] = modulePayload[key];
          }
        }
      } catch(e) { console.warn("Lookup discovery bypassed for " + modName + ": " + e.message); }
    }
  });
  return masterLookups;
}

/**
 * Global Metadata Audit Helper.
 * Queries the centralized system log sheet to identify the last operation timestamp.
 */
function getEventTimestampFromLogs(module, type, entity) {
  try {
    var sheet = getLogsDb().getSheetByName("System Logs");
    if (!sheet) return "Not recorded";
    var data = sheet.getDataRange().getDisplayValues();
    for (var i = data.length - 1; i > 0; i--) {
      if (data[i][1] === module && data[i][2] === type && data[i][6] === entity) {
        return data[i][0];
      }
    }
    return "Not recorded";
  } catch(e) { return "Unknown"; }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */