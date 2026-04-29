/**
 * Event Broker - Core Orchestrator
 * Automatically routes system events to interested modules.
 */
/**
 * Event Broker - Core Orchestrator
 * Automatically routes system events to interested modules.
 */
var SystemEvent = (function() {
  
  // UPGRADED: Added 'extraData' to pass dynamic template placeholders
  function emit(module, type, name, severity, entity, details, recipientEmail, extraData) {
    var payload = {
      module: module, 
      type: type, 
      handle: module + ":" + type, 
      name: name,
      severity: severity || "INFO",
      entity: entity, 
      details: details,
      recipientEmail: recipientEmail || "",
      extraData: extraData || {}, // Carries variables like {{resetLink}}
      timestamp: new Date(), 
      user: getLoggedInUsername()
    };

    // 1. Core Logging
    if (typeof Logs !== 'undefined' && Logs.handleSystemEvent) {
      Logs.handleSystemEvent(payload);
    }

    // 2. Dynamic Extensions Notification
    var props = PropertiesService.getScriptProperties();
    var installed = props.getProperty('INSTALLED_MODULES');
    if (installed) {
      installed.split(',').forEach(function(modName) {
        var mod = modName.trim();
        var handlerName = mod + "_on" + type;
        if (typeof globalThis[handlerName] === 'function') {
          try { globalThis[handlerName](payload); } catch(e) { console.error(e.message); }
        }
      });
    }
    
    // 3. Email Dispatcher
    if (typeof Templates !== 'undefined' && Templates.handleEventEmail) {
      Templates.handleEventEmail(payload);
    }
  }

  return { emit: emit };
})();