/*



  <div id="ro-clients-core-config" style="background:#fff; padding:20px; border-radius:8px; border:1px solid #e2e8f0; margin-bottom: 20px;"></div>
  
  <h4 class="settings-section-header">Form Accordions</h4>
  <div id="ro-clients-accordions" style="background:var(--brand-bg); padding:20px; border-radius:8px; border:1px solid #e2e8f0; margin-bottom: 20px;"></div>
  
  <h4 class="settings-section-header">Custom Fields</h4>
  <div id="ro-clients-fields" style="background:var(--brand-bg); padding:20px; border-radius:8px; border:1px solid #e2e8f0;"></div>
*/


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