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
  ScriptApp.newTrigger("executeDailyCronJobs")
    .timeBased()
    .everyDays(1)
    .atHour(1) 
    .create();
    
  setupEmailQueueTrigger();
}

function setupEmailQueueTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerExists = triggers.some(t => t.getHandlerFunction() === 'processEmailQueue');
  if (!triggerExists) {
    ScriptApp.newTrigger('processEmailQueue')
      .timeBased()
      .everyMinutes(5)
      .create();
  }
}

function processEmailQueue() {
  try {
    var queueDb = getQueueDb();
    var sheet = queueDb.getSheetByName("Email Queue");
    if (!sheet) return;
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    
    var settings = getSystemSettings();
    var mainDb = getMainDb();
    var tData = mainDb.getSheetByName("Templates").getDataRange().getValues();
    var wData = mainDb.getSheetByName("Wrappers").getDataRange().getValues();
    
    function isWrapperActive(wName) {
      for (var w=1; w<wData.length; w++) {
        if (wData[w][2] === wName && wData[w][5] === "Active") return true;
      }
      return false;
    }
    
    var hardcodedTriggers = ["Logs:EXPORT", "Logs:PURGE", "Settings:UPDATE", "System:INSTALL", "System:MODULE_INSTALLED"];

    for (var i = 1; i < data.length; i++) {
      var status = data[i][8];
      if (status === "Pending") {
        var rowNum = i + 1;
        var templateId = data[i][2];
        var recipient = data[i][3];
        var cc = data[i][4];
        var bcc = data[i][5];
        var contextMap = {};
        try { contextMap = JSON.parse(data[i][6] || "{}"); } catch(e){}
        var dispatchMode = data[i][7];
        
        try {
          if (hardcodedTriggers.indexOf(templateId) > -1) {
            executeSendHardcodedEmail(templateId, recipient, contextMap, settings, wData);
          } else {
            executeSendTemplateEmail(templateId, recipient, cc, bcc, contextMap, dispatchMode, tData, wData, settings, isWrapperActive);
          }
          sheet.getRange(rowNum, 9, 1, 3).setValues([["Sent", "", new Date()]]);
        } catch(err) {
          sheet.getRange(rowNum, 9, 1, 2).setValues([["Failed", err.message]]);
        }
      }
    }
    SpreadsheetApp.flush();
  } catch(e) {
    console.error("Queue execution error: " + e.message);
  }
}

function executeSendHardcodedEmail(triggerHandle, toEmail, dataMap, settings, wData) {
  var subject = "";
  var htmlBody = "";
  var wrapperName = "Internal Communication"; 

  switch (triggerHandle) {
    case "Logs:EXPORT":
      subject = "System Logs Exported";
      htmlBody = "<div style='font-family: sans-serif; padding: 20px;'><h2>Logs Exported</h2><p>The system logs have been successfully exported.</p><p>Details: {{details}}</p></div>";
      break;
    case "Logs:PURGE":
      subject = "System Logs Purged";
      htmlBody = "<div style='font-family: sans-serif; padding: 20px;'><h2>Logs Purged</h2><p>The system logs have been purged by an administrator.</p><p>Details: {{details}}</p></div>";
      break;
    case "Settings:UPDATE":
      subject = "System Settings Updated";
      htmlBody = "<div style='font-family: sans-serif; padding: 20px;'><h2>Settings Updated</h2><p>The global system configuration has been modified.</p><p>Details: {{details}}</p></div>";
      break;
    case "System:INSTALL":
      subject = "Installation Successful";
      htmlBody = "<div style='font-family: sans-serif; padding: 20px;'><h2>Installation Successful</h2><p>The system has been fully deployed and is ready for use as <strong>{{sysName}}</strong>.</p><hr><p><strong>Your Initial Credentials:</strong><br>Username: {{adminUsername}}<br>Role: Administrator</p></div>";
      break;
    case "System:MODULE_INSTALLED":
      subject = "New System Feature Available: {{moduleName}}";
      htmlBody = "<div style='font-family: sans-serif; padding: 20px;'><h2>System Upgrade Complete</h2><p>The <strong>{{moduleName}}</strong> module has been successfully installed into the core ecosystem.</p><p>Please refresh your dashboard to access the new features and logic.</p></div>";
      break;
    default:
      return;
  }

  htmlBody = applyGlobalSignature(htmlBody);
  
  var wrapperHtml = "{{USER_MESSAGE_CONTENT}}";
  for (var w = 1; w < wData.length; w++) {
    if (wData[w][2] === wrapperName && wData[w][5] === "Active") {
      wrapperHtml = wData[w][4];
      break;
    }
  }
  var fullHtml = wrapperHtml.replace("{{USER_MESSAGE_CONTENT}}", htmlBody);
  
  for (var key in dataMap) {
      var regex = new RegExp("\\{\\{" + key + "\\}\\}", "gi");
      subject = subject.replace(regex, dataMap[key] || "");
      fullHtml = fullHtml.replace(regex, dataMap[key] || "");
  }

  var finalToEmail = toEmail;
  if (settings.environment === 'Sandbox' && settings.adminEmail !== '') {
      finalToEmail = settings.adminEmail;
      subject = "[Sandbox Mail] " + subject;
      var sandboxWarning = "<br><br><div style='padding: 20px; background-color: #000; color: #0f0; font-family: monospace; font-size: 14px; border: 2px solid #333; margin-top: 50px;'>";
      sandboxWarning += "SYSTEM OVERRIDE: SANDBOX ENVIRONMENT INTERCEPTED<br>";
      sandboxWarning += "&gt; INTENDED RECIPIENT: " + toEmail + "<br></div>";
      fullHtml += sandboxWarning;
  }

  fullHtml = applyLinkTracking(fullHtml, finalToEmail);
  MailApp.sendEmail({
      to: finalToEmail, 
      subject: subject, 
      htmlBody: fullHtml, 
      noReply: true, 
      name: settings.systemName, 
      inlineImages: { logo: getLogoBlob() }
  });
}

function executeSendTemplateEmail(templateId, recipient, cc, bcc, contextMap, dispatchMode, tData, wData, settings, isWrapperActive) {
  var targetTpl = null;
  for (var t = 1; t < tData.length; t++) {
    if (tData[t][1] === templateId) {
      targetTpl = tData[t];
      break;
    }
  }
  if (!targetTpl) throw new Error("Template " + templateId + " not found.");
  
  var baseSubject = targetTpl[7];
  var baseHtmlBody = targetTpl[8];
  var wrapperName = targetTpl[9];
  
  var wrapperHtml = "{{USER_MESSAGE_CONTENT}}";
  for (var w = 1; w < wData.length; w++) {
    if (wData[w][2] === wrapperName && wData[w][5] === "Active") {
      wrapperHtml = wData[w][4];
      break;
    }
  }
  
  baseHtmlBody = applyGlobalSignature(baseHtmlBody);
  var fullHtml = wrapperHtml.replace("{{USER_MESSAGE_CONTENT}}", baseHtmlBody);
  
  for (var token in contextMap) {
    var regex = new RegExp("\\{\\{" + token + "\\}\\}", "gi");
    baseSubject = baseSubject.replace(regex, contextMap[token] || "");
    fullHtml = fullHtml.replace(regex, contextMap[token] || "");
  }
  
  var finalTo = recipient;
  var finalCc = cc;
  var finalBcc = bcc;
  
  if (settings.environment === 'Sandbox' && settings.adminEmail !== '') {
    finalTo = settings.adminEmail;
    finalCc = "";
    finalBcc = "";
    baseSubject = "[Sandbox Mail] " + baseSubject;
    var sandboxWarning = "<br><br><div style='padding: 20px; background-color: #000; color: #0f0; font-family: monospace; font-size: 14px; border: 2px solid #333; margin-top: 50px;'>";
    sandboxWarning += "SYSTEM OVERRIDE: SANDBOX ENVIRONMENT INTERCEPTED<br>";
    sandboxWarning += "&gt; INTENDED RECIPIENT: " + recipient;
    if (cc) sandboxWarning += "<br>&gt; INTENDED CC: " + cc;
    if (bcc) sandboxWarning += "<br>&gt; INTENDED BCC: " + bcc;
    sandboxWarning += "<br></div>";
    fullHtml += sandboxWarning;
  }
  
  fullHtml = applyLinkTracking(fullHtml, finalTo.split(',')[0].trim());
  
  var mailOptions = {
    to: finalTo,
    subject: baseSubject,
    htmlBody: fullHtml,
    noReply: true,
    name: settings.systemName,
    inlineImages: { logo: getLogoBlob() }
  };
  
  if (finalCc) mailOptions.cc = finalCc;
  var tplBccArchive = PropertiesService.getScriptProperties().getProperty('TPL_BCC_ARCHIVE') || '';
  if (finalBcc || (settings.environment !== 'Sandbox' && tplBccArchive)) {
    mailOptions.bcc = finalBcc || tplBccArchive;
  }
  
  MailApp.sendEmail(mailOptions);
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
 * Creates a nightly trigger for the Log Janitor if it doesn't already exist.
 */
function setupLogJanitorTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerExists = triggers.some(t => t.getHandlerFunction() === 'runLogJanitor');
  
  if (!triggerExists) {
    ScriptApp.newTrigger('runLogJanitor')
      .timeBased()
      .atHour(2) // Runs at 2:00 AM
      .everyDays(1)
      .create();
  }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */