/**
 * Logs Module - Backend
 * Core module for system audit trails and event monitoring.
 */

var Logs = {
  /**
   * Subscriber function called by the Event Broker.
   */
  handleSystemEvent: function(payload) {
    try {
      var ss = getLogsDb();
      var sheet = ss.getSheetByName("System Logs");
      if (!sheet) return;

      var env = PropertiesService.getScriptProperties().getProperty('ENVIRONMENT') || "Sandbox";

      var rowData = [[
        payload.timestamp,
        payload.module,
        payload.type,
        payload.name,
        "INFO", // Default severity
        payload.user,
        payload.entity,
        payload.details,
        env
      ]];

      sheet.insertRowAfter(1);
      sheet.getRange(2, 1, 1, 9).setValues(rowData);
      SpreadsheetApp.flush();
    } catch (e) {
      console.error("Logs Handler Error: " + e.message);
    }
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
    
    return data.map(function(row, i) {
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
  } catch (e) { return []; }
}