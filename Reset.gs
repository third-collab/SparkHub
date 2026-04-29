/**
 * Utility function to clear all system properties during testing,
 * while securely preserving the master webhook secret.
 * Run this manually from the Apps Script editor between test installs.
 */
function resetTestingEnvironment() {
  var scriptProps = PropertiesService.getScriptProperties();
  var allProps = scriptProps.getProperties();
  var deletedCount = 0;

  for (var key in allProps) {
    // Check against the exact key name used in your setup
    if (key !== 'WEBHOOK_SECRET') {
      scriptProps.deleteProperty(key);
      deletedCount++;
    }
  }

  console.log("Environment reset complete. Wiped " + deletedCount + " properties. WEBHOOK_SECRET was safely preserved.");
}