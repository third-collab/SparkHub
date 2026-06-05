function testWebhookConnection() {
  var props = PropertiesService.getScriptProperties();
  var clientId = props.getProperty('CLIENT_ID');
  var instanceSecret = props.getProperty('INSTANCE_SECRET');
  
  Logger.log("Client ID: " + clientId);
  Logger.log("Instance Secret: " + instanceSecret);
  Logger.log("Webhook URL: " + MASTER_WEBHOOK_URL);
  
  if (!MASTER_WEBHOOK_URL || !clientId || !instanceSecret) {
    Logger.log("FAILED: Missing required credentials or Webhook URL. Cannot send ping.");
    return;
  }
  
  var payload = { 
    action: "module_install", 
    secretKey: "SparkHub-Sec-92vM4xL7qP8nR3wK1bC6",
    clientId: clientId, 
    instanceSecret: instanceSecret,
    moduleName: "DiagnosticTestModule"
  };
  
  var response = UrlFetchApp.fetch(MASTER_WEBHOOK_URL, { 
    method: 'post', 
    contentType: 'application/json', 
    payload: JSON.stringify(payload), 
    muteHttpExceptions: true 
  });
  
  Logger.log("REGISTRY RESPONSE: " + response.getContentText());
}

function testLinkTrackingAndSignature() {
  // Get the email of the person running the script (you)
  var myEmail = Session.getActiveUser().getEmail();
  
  var testSubject = "SparkHub System Test: Tracking & Signatures";
  
  // A simple HTML body with a standard link inside it
  var testBody = "<h3>Hello!</h3>" + 
                 "<p>This is an automated test from SparkHub.</p>" + 
                 "<p>Please click the link below to test the tracking redirect:</p>" + 
                 "<p><a href='https://www.google.com'>Test The Tracking Link</a></p>" +
                 "<br><p>If signatures are working, it should appear right below this message.</p>";
                 
  // Fetch your configured default wrapper
  var props = PropertiesService.getScriptProperties();
  var defaultWrapper = props.getProperty('TPL_DEFAULT_WRAPPER') || 'Internal Communication';
  
  // Dispatch the email using your standard engine
  try {
    sendHardcodedEmail(myEmail, testSubject, testBody, defaultWrapper, {});
    Logger.log("Test email successfully sent to: " + myEmail);
    Logger.log("Check your inbox!");
  } catch (e) {
    Logger.log("Error sending email: " + e.message);
  }
}


function testDirectTrackingAndSignature() {
  var myEmail = Session.getActiveUser().getEmail();
  if (!myEmail) {
    Logger.log("Error: Cannot detect your email. Are you logged into multiple Google accounts?");
    return;
  }
  
  var testSubject = "Direct Bypass Test: Tracking & Signatures";
  
  // A simple raw HTML body
  var rawBody = "<h3>Direct Test</h3>" + 
                "<p>This bypasses the SparkHub wrappers to test only the signature and tracking.</p>" + 
                "<p><a href='https://www.google.com'>Test The Tracking Link</a></p>";
                 
  try {
    // 1. Manually apply the signature
    var bodyWithSig = applyGlobalSignature(rawBody);
    
    // 2. Manually apply the tracking wraps
    var finalBody = applyLinkTracking(bodyWithSig, myEmail);
    
    // 3. Send directly via Google (bypassing Sandbox/Wrappers)
    MailApp.sendEmail({
      to: myEmail,
      subject: testSubject,
      htmlBody: finalBody
    });
    
    Logger.log("DIRECT SEND SUCCESS to: " + myEmail);
  } catch (e) {
    Logger.log("DIRECT SEND FAILED: " + e.message);
  }
}