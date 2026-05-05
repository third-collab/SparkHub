/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Triggers.gs
 * VERSION: 1.1
 * SYNC STATUS: Fully Synchronized with Installation.gs
 */

/**
 * Triggers Module
 * Orchestrates automated system cron jobs.
 */

/**
 * Initializes all system-wide triggers.
 * This is called automatically by Installation.gs during setup.
 */
function setupSystemTriggers() {
  // Clear existing triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  // Create the new daily trigger to fire between 1 AM and 2 AM
  ScriptApp.newTrigger("executeDailyCronJobs")
    .timeBased()
    .everyDays(1)
    .atHour(1) 
    .create();
  console.log("Daily triggers initialized!");
}

/**
 * The main engine that runs every night.
 * Fires the system ping and executes module-specific daily checks.
 */
function executeDailyCronJobs() {
  // 1. Report health to Master Webhook
  sendDailyPing();
  
  // 2. Future Module Executions (e.g., Client alerts, Calendar reminders)
  // checkClientEvents();
  // processOverdueInvoices();
}

/**
 * Sends a lightweight health check to the Master Installations webhook.
 * Authenticates using the specific Instance Secret.
 */
function sendDailyPing() {
  var props = PropertiesService.getScriptProperties();
  var webhookUrl = "https://script.google.com/macros/s/AKfycbyderUFTDgJjjSb4ML2xpXzRnfKp_yBLkYlpaKdZcWZLowtmiutt-QZsg7OMq0enBJljw/exec";
  
  var masterSecret = props.getProperty('WEBHOOK_SECRET');
  var clientId = props.getProperty('CLIENT_ID');
  var instanceSecret = props.getProperty('INSTANCE_SECRET');
  
  // If the system hasn't been fully configured, abort the ping.
  if (!webhookUrl || !masterSecret || !clientId || !instanceSecret) return;
  
  var payload = {
    action: "ping",
    secretKey: masterSecret,
    clientId: clientId,
    instanceSecret: instanceSecret,
    timestamp: new Date().toISOString()
  };
  
  try {
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    // If ping fails (e.g. no internet, webhook down), fail silently to not interrupt other cron jobs
    console.warn("Daily ping failed to send: " + e.message);
  }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */