// --- LOGS MODULE REGISTRY EXPORTS ---
function Logs_getTriggers() {
  return ["Logs:EXPORT", "Logs:PURGE"];
}

function Logs_getPlaceholders() {
  return ["logTimestamp", "logModule", "logAction", "logActor", "logEntity", "logDetails"];
}

/**
 * Logs Module - Backend
 * Core module for system audit trails and event monitoring.
 */
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
        payload.severity, // FIX: Now dynamically pulls INFO, WARN, or ERROR
        payload.user,
        payload.entity,
        payload.details,
        env
      ]);
      
      SpreadsheetApp.flush();
    } catch (e) { console.error("Logs Handler Error: " + e.message); }
  }
};

/**
 * UI Data Provider
 */
function getLogsList() {
  try {
    var sheet = getLogsDb().getSheetByName("System Logs");
    var data = sheet.getDataRange().getDisplayValues();
    data.shift(); // Remove headers
    
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
    
    // Reverse the array so the frontend dashboard still displays the newest logs at the top
    return formattedLogs.reverse();
  } catch (e) { return []; }
}

function getEventTimestampFromLogs(module, type, entity) {
  try {
    var sheet = getLogsDb().getSheetByName("System Logs");
    if (!sheet) return "Not recorded";
    var data = sheet.getDataRange().getDisplayValues();
    // Loop backwards to find the most recent matching record
    for (var i = data.length - 1; i > 0; i--) {
      // Module is Col B (1), Type is Col C (2), Entity is Col G (6)
      if (data[i][1] === module && data[i][2] === type && data[i][6] === entity) {
        return data[i][0]; // Return the Timestamp
      }
    }
    return "Not recorded";
  } catch(e) { return "Unknown"; }
}