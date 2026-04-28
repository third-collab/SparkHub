/**
 * Event Broker - Core Orchestrator
 * Automatically routes system events to interested modules.
 */
var SystemEvent = (function() {
  
  function emit(module, type, name, entity, details) {
    var payload = {
      module: module, type: type, name: name,
      entity: entity, details: details,
      timestamp: new Date(), user: getLoggedInUsername()
    };

    // 1. MANDATORY: System Logging
    if (typeof Logs !== 'undefined' && Logs.handleSystemEvent) {
      Logs.handleSystemEvent(payload);
    }

    // 2. DYNAMIC DISCOVERY: Notify all extension modules
    var installed = PropertiesService.getScriptProperties().getProperty('INSTALLED_MODULES');
    if (installed) {
      installed.split(',').forEach(function(modName) {
        var mod = modName.trim();
        // Look for: [Module]_on[Action] (e.g., Clients_onCREATE)
        var handlerName = mod + "_on" + type;
        if (typeof this[handlerName] === 'function') {
          try {
            this[handlerName](payload);
          } catch(e) {
            console.error("Extension handler error [" + handlerName + "]: " + e.message);
          }
        }
      });
    }
    
    // 3. SPECIAL CASE: Templates (Emails)
    if (typeof Templates !== 'undefined' && Templates.handleEventEmail) {
      Templates.handleEventEmail(payload);
    }
  }

  return { emit: emit };
})();