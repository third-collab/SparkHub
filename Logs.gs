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
    
    var formattedLogs = data.map(function(row, i) {
      return {
        rowIndex: i + 2,
        timestamp: row[0],
        module: row[1],
        type: row[2],
        name: row[3],
        severity: row[4],
        actor: row[5],
        entity: row[6],
        details: row[7],
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
// Currently handled by CORE PROCESSORS (handleSystemEvent)

// ========================================================================
// 5. INTERNAL HELPERS
// ========================================================================
// No internal helpers currently needed for Logs

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */