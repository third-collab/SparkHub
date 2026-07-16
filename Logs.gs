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

function Logs_getPermissions() {
  return ["View Logs", "Manage Settings"];
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
        payload.entity,       // Target Entity ID (Index 6)
        payload.entityName,   // Target Entity Name (Index 7)
        payload.details,      // Log Details (Index 8)
        env                   // Environment (Index 9)
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
    
    var userMap = {};
    try {
      if (typeof getUsersList === 'function') {
        getUsersList().forEach(function(u) {
          userMap[u.userId] = u.firstName + " " + (u.lastName ? u.lastName.charAt(0).toUpperCase() + "." : "");
        });
      }
    } catch(e){}
    
    var formattedLogs = data.map(function(row, i) {
      var actorTrackingKey = row[5];
      var friendlyDisplayActor = userMap[actorTrackingKey] || actorTrackingKey || "System";
      
      return {
        rowIndex: i + 2,
        timestamp: row[0],
        module: row[1],
        type: row[2],
        name: row[3],
        severity: row[4],
        actor: friendlyDisplayActor,
        entity: row[7] || "-",       // Natively pulls pre-saved friendly Target Entity Name directly
        entityId: row[6] || "",      // Target Entity ID Column
        details: row[8],             // Log Details
        env: row[9]                  // Environment Profile
      };
    });
    return formattedLogs.reverse();
  } catch (e) { return []; }
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
    // Verify target schema context natively inside the logging layer before persistence commit
    if (payload.logsDbId) {
      validateSpreadsheetSchema(payload.logsDbId, ["System Logs"], "Logs Database");
    }
    props.setProperty('LOGS_DATABASE_ID', payload.logsDbId);
    props.setProperty('LOGS_RETENTION_DAYS', payload.retentionDays);
    // Ensure the nightly trigger is active
    setupLogJanitorTrigger();
    // Group updates inside the primary 'System' routing layer per core directive standard
    SystemEvent.emit("System", "UPDATE", "Config Updated", "INFO", "-", "Logs registry and retention settings updated.", "", { targetName: "Logs" });
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
function getLogsModuleConfig() {
  try {
    var props = PropertiesService.getScriptProperties();
    return {
      logsDbId: props.getProperty('LOGS_DATABASE_ID') || '',
      logsRetentionDays: props.getProperty('LOGS_RETENTION_DAYS') || '90'
    };
  } catch(e) {
    return { logsDbId: '', logsRetentionDays: '90' };
  }
}

function Logs_getPlaceholderMetadata() {
  return {
    "logTimestamp": { desc: "The chronological date and time stamp logging precisely when an action transaction occurred.", tag: "System Logs" },
    "logModule": { desc: "The specific architectural application layer or extension routing code that emitted the event log.", tag: "System Logs" },
    "logAction": { desc: "The descriptive high-level operational verb or process handle recorded (e.g., User Status Changed).", tag: "System Logs" },
    "logActor": { desc: "The identifier handle or user profile name of the account who physically authorized the system action.", tag: "System Logs" },
    "logEntity": { desc: "The target object reference unique key identifier or name subjected to the action modification layer.", tag: "System Logs" },
    "logDetails": { desc: "A narrative contextual log tracking exactly what attributes or fields were updated during the process cycle.", tag: "System Logs" }
  };
}

function Logs_getModuleLabel() { return "System Logs"; }

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */