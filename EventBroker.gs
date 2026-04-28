/**
 * Event Broker - Core Orchestrator
 * Manages system-wide hooks and module subscriptions.
 */
var SystemEvent = (function() {
  
  // The Master Hook Registry
  // Format: "Module:Action": [functionReference, ...]
  var registry = {
    // Every action is automatically hooked to the Logs handler
    "ANY:ANY": ["Logs.handleSystemEvent"],
    
    // Example of specific module subscriptions
    "Users:CREATE": ["Templates.handleEventEmail"],
    "Users:UPDATE": ["Templates.handleEventEmail"]
  };

  /**
   * Broadcasts a system event to all subscribers.
   * @param {string} module - Originating module.
   * @param {string} type - Action type (CREATE, UPDATE, DELETE, etc).
   * @param {string} name - Action name (e.g., 'Add User').
   * @param {string} entity - Target entity name/ID.
   * @param {string} details - Log text or data payload.
   */
  function emit(module, type, name, entity, details) {
    var eventKey = module + ":" + type;
    var subscribers = [];

    // 1. Get Global Subscribers
    if (registry["ANY:ANY"]) subscribers = subscribers.concat(registry["ANY:ANY"]);

    // 2. Get Module/Action Specific Subscribers
    if (registry[eventKey]) subscribers = subscribers.concat(registry[eventKey]);

    // 3. Execute Handlers
    subscribers.forEach(function(handlerPath) {
      try {
        var parts = handlerPath.split('.');
        var func = parts.length > 1 ? this[parts[0]][parts[1]] : this[parts[0]];
        
        if (typeof func === 'function') {
          func({
            module: module,
            type: type,
            name: name,
            entity: entity,
            details: details,
            timestamp: new Date(),
            user: getLoggedInUsername()
          });
        }
      } catch (e) {
        console.error("Broker Execution Failure [" + handlerPath + "]: " + e.message);
      }
    });
  }

  return { emit: emit };
})();