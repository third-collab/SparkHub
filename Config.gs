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

function getQueueDb() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('QUEUE_DATABASE_ID');
  if (!id) {
    throw new Error("Configuration Error: Queue Database ID is missing. Please initialize or map it via Templates configuration.");
  }
  
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error("System Error: Unable to open Queue Database.");
  }
}

/**
 * Centralized Ecosystem Active Modules Registry.
 * Combines immutable core architectural layers with dynamically registered extensions.
 * Reusable globally across all cross-module discovery and routing workflows.
 */
function getActiveModules() {
  var modules = ["System", "Users", "Templates", "Settings", "Logs"];
  var installed = PropertiesService.getScriptProperties().getProperty('INSTALLED_MODULES');
  if (installed) {
    installed.split(',').forEach(function(m) {
      var mod = m.trim();
      if (mod && modules.indexOf(mod) === -1) modules.push(mod);
    });
  }
  return modules;
}

/**
 * Centralized Ecosystem Lookup Aggregator
 * Automatically discovers and harvests all lookable entity reference models 
 * across core infrastructure and dynamically installed external extensions.
 * Follows the V8 Auto-Discovery pattern.
 */
function getSystemDynamicLookups() {
  // Initialized as a pure open object to ensure zero hardcoded trace of any core or external module keys
  var masterLookups = {};
  var activeModules = getActiveModules();
  
  var globalScope = (1, eval)("this");
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
 * Universal Ecosystem Permission Discovery Engine.
 * Promoted to the centralized system config layer to enforce complete module independence.
 * Loops over active extensions and auto-builds user-friendly permission capability maps.
 */
function getDynamicPermissionMatrix() {
  var matrix = {};
  // Explicitly seeds the fundamental baseline Core System settings group
  matrix["Core System"] = ["View Settings", "Manage Settings"];
  var activeModules = getActiveModules();
  var globalScope = (1, eval)("this");
  
  activeModules.forEach(function(modName) {
    var funcName = modName + "_getPermissions";
    if (typeof globalScope[funcName] === 'function') {
      try {
        var perms = globalScope[funcName]();
        if (Array.isArray(perms)) {
          // Maps technical module handles dynamically into standard operational UI category labels
          var groupLabel = modName + " Module";
          var globalScope = (1, eval)("this");
          
          if (typeof globalScope[modName + "_getModuleLabel"] === 'function') {
            groupLabel = globalScope[modName + "_getModuleLabel"]();
          } else if (modName === "System" || modName === "Settings") {
            return; // Covered under Core System settings
          }
        
          matrix[groupLabel] = perms;
        }
      } catch(e) { console.warn("Permission matrix extraction bypassed for " + modName + ": " + e.message); }
    }
  });
  
  return matrix;
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