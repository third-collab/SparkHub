/**
 * Event Broker - Core Orchestrator
 * Updated to support targeted email recipients.
 */
var SystemEvent = (function() {
  
  function emit(module, type, name, entity, details, recipientEmail) {
    var payload = {
      module: module, 
      type: type, 
      handle: module + ":" + type, // The unique Event Handle (e.g., Users:CREATE)
      name: name,
      entity: entity, 
      details: details,
      recipientEmail: recipientEmail || "",
      timestamp: new Date(), 
      user: getLoggedInUsername()
    };

    // 1. Core Logging
    if (typeof Logs !== 'undefined' && Logs.handleSystemEvent) {
      Logs.handleSystemEvent(payload);
    }

    // 2. Extensions Notification
    var installed = PropertiesService.getScriptProperties().getProperty('INSTALLED_MODULES');
    if (installed) {
      installed.split(',').forEach(function(modName) {
        var mod = modName.trim();
        var handlerName = mod + "_on" + type;
        if (typeof this[handlerName] === 'function') {
          try { this[handlerName](payload); } catch(e) { console.error("Handler error: " + e.message); }
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