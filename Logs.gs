/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Logs.gs
 * VERSION: 1.1
 * SYNC STATUS: Fully Synchronized with LogsData.html
 */

// ========================================================================
// 1. REGISTRY EXPORTS
// ========================================================================
function Logs_getTriggers() {
  return ["Logs:EXPORT", "Logs:PURGE"];
}

function Logs_getPlaceholders() {
  return ["logTimestamp", "logModule", "logAction", "logActor", "logEntity", "logDetails"];
}

// ========================================================================
// 2. CORE PROCESSORS
// ========================================================================
var Logs = {
  handleSystemEvent: function(payload) {
    try {
      var ss = getLogsDb();
      var sheet = ss.getSheetByName("System Logs");
      if (!sheet) return;

      var env = PropertiesService.getScriptProperties().getProperty('ENVIRONMENT') || "Sandbox";
      sheet.appendRow([
        payload.timestamp,
        payload.module,
        payload.type,
        payload.name,
        payload.severity, 
        payload.user,
        payload.entity,
        payload.details,
        env
      ]);
      SpreadsheetApp.flush();
    } catch (e) { console.error("Logs Handler Error: " + e.message); }
  }
};

// ========================================================================
// 3. READ / GET FUNCTIONS
// ========================================================================
function getLogsList() {
  try {
    var sheet = getLogsDb().getSheetByName("System Logs");
    var data = sheet.getDataRange().getDisplayValues();
    data.shift(); 
    
    // Fetch user mapping cache to convert User ID codes into friendly handles dynamically
    var userLookupCache = {};
    try {
      if (typeof getUsersList === 'function') {
        getUsersList().forEach(function(u) {
          userLookupCache[u.userId] = u.firstName + " " + u.lastName + " (" + u.username + ")";
        });
      }
    } catch(uErr) { console.warn("User lookup cache mapping failed for logs view: " + uErr.message); }
    
    var formattedLogs = data.map(function(row, i) {
      var actorTrackingKey = row[5];
      var friendlyDisplayActor = userLookupCache[actorTrackingKey] || actorTrackingKey || "System";
      
      return {
        rowIndex: i + 2,
        timestamp: row[0],
        module: row[1],
        type: row[2],
        name: row[3],
        severity: row[4],
        actor: friendlyDisplayActor,
        entity: row[6],
        details: 
        row[7],
        env: row[8]
      };
    });
    return formattedLogs.reverse();
  } catch (e) { return []; }
}

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

// ========================================================================
// 4. WRITE / SAVE FUNCTIONS
// ========================================================================
/**
 * Saves the Logs module specific configuration.
 */
function saveLogsModuleConfig(payload) {
  try {
    var props = PropertiesService.getScriptProperties();
    // Move the ID here and save the new retention policy
    props.setProperty('LOGS_DATABASE_ID', payload.logsDbId);
    props.setProperty('LOGS_RETENTION_DAYS', payload.retentionDays);
    
    // Ensure the nightly trigger is active
    setupLogJanitorTrigger();
    
    SystemEvent.emit("Logs", "UPDATE", "Config Updated", "INFO", "Logs", "Logs registry and retention settings updated.");
    return { success: true };
  } catch (e) { 
    return { error: "Logs.gs: " + e.message }; 
  }
}

/**
 * Nightly Janitor: Purges logs older than the retention period.
 * Triggered automatically by setupLogJanitorTrigger()
 */
function runLogJanitor() {
  var props = PropertiesService.getScriptProperties();
  var retentionDays = parseInt(props.getProperty('LOGS_RETENTION_DAYS') || "90");
  
  // If set to 0 (Indefinite), stop here.
  if (retentionDays === 0) return; 

  var sheet = getLogsDb().getSheetByName("System Logs");
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return; // Only header exists

  var now = new Date();
  // Calculate the cutoff date based on settings
  var cutoff = new Date(now.getTime() - (retentionDays * 24 * 60 * 60 * 1000));
  
  var rowsToKeep = [data[0]]; // Always keep the header
  var deletedCount = 0;

  for (var i = 1; i < data.length; i++) {
    var logDate = new Date(data[i][0]); // Assumes Timestamp is in first column
    if (logDate >= cutoff) {
      rowsToKeep.push(data[i]);
    } else {
      deletedCount++;
    }
  }
  
  // If we found old logs, overwrite the sheet with only the fresh logs
  if (deletedCount > 0) {
    sheet.clearContents();
    sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
    console.log("Log Janitor: Purged " + deletedCount + " old logs.");
  }
}

// ========================================================================
// 5. INTERNAL HELPERS
// ========================================================================
// No internal helpers currently needed for Logs

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */