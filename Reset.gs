/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Reset.gs
 * VERSION: 1.1
 * SYNC STATUS: Standalone Utility
 */

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

/**
 * Uninstalls all external modules by clearing them from the system registry,
 * while strictly preserving the Core Engine modules.
 */
function removeExternalModules() {
  try {
    var props = PropertiesService.getScriptProperties();
    var installed = props.getProperty('INSTALLED_MODULES') || "";
    
    if (!installed) {
      return { success: true, message: "No external modules are currently installed." };
    }
    
    // Define the untouchable Core Modules
    var coreModules = ["Settings", "Users", "Templates", "Logs"];
    var currentModules = installed.split(',').map(function(m){ return m.trim(); }).filter(function(m){ return m !== ""; });
    
    var keptModules = [];
    var removedModules = [];
    
    // Separate core from external
    for (var i = 0; i < currentModules.length; i++) {
      if (coreModules.indexOf(currentModules[i]) > -1) {
        keptModules.push(currentModules[i]);
      } else {
        removedModules.push(currentModules[i]);
      }
    }
    
    // Update the system property to only include kept modules (if any)
    props.setProperty('INSTALLED_MODULES', keptModules.join(','));
    
    if (removedModules.length > 0) {
      var removedString = removedModules.join(', ');
      SystemEvent.emit("Settings", "UPDATE", "Modules Uninstalled", "WARN", "System", "The following external modules were uninstalled: " + removedString, "System", {});
      return { success: true, message: "Successfully uninstalled external modules: " + removedString };
    } else {
      return { success: true, message: "Only core modules were found in the registry. Nothing removed." };
    }
    
  } catch (e) {
    return { error: "Error uninstalling modules: " + e.message };
  }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */