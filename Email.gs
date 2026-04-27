/**
 * Email Engine - Backend
 * Standardized under MR Hub Architecture Blueprint.
 * Handles: Template merging, placeholder swapping, and environment-based routing.
 */

/**
 * Helper to fetch the raw HTML content of a specific wrapper file.
 * @param {string} type - The wrapper filename prefix (e.g., 'Internal', 'External').
 * @return {string} The raw HTML content of the file.
 */
function getWrapperContent(type) {
  var fileName = type + "Wrapper"; 
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

/**
 * Merges the designated wrapper and body content for the UI template preview.
 * Swaps CID references for live URLs (SVG or Drive Thumbnail) so images render in the browser.
 * @param {number} rowIndex - The row index of the template in the spreadsheet.
 * @return {string} The fully rendered HTML string.
 */
function getRenderedTemplatePreview(rowIndex) {
  var sheet = getMainDb().getSheetByName("Templates");
  var data = sheet.getDataRange().getValues();
  var rowData = data[rowIndex];
  
  var rawHtml = rowData[7] || "";
  var wrapperType = rowData[9] || "Internal";
  
  var wrapperHtml = getWrapperContent(wrapperType);
  var fullHtml = wrapperHtml.replace("{{USER_MESSAGE_CONTENT}}", rawHtml);
  
  // Use centralized settings to find the correct logo for browser display
  var settings = getSystemSettings();
  var defaultSvgLogo = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'%3E%3Crect width='100' height='100' rx='20' fill='%23C40004'/%3E%3Ctext x='50' y='65' font-family='Arial' font-size='40' font-weight='bold' fill='white' text-anchor='middle'%3EMR%3C/text%3E%3C/svg%3E";
  
  // Browser preview can render SVG; if no Drive logo exists, use the SVG
  var displayLogoUrl = settings.systemLogoId ? settings.systemLogoUrl : defaultSvgLogo;
  
  // Replace CID with the dynamic URL for browser rendering
  fullHtml = fullHtml.replace(/src="cid:logo"/g, 'src="' + displayLogoUrl + '"');
  return fullHtml;
}

/**
 * Helper to retrieve the system logo as a blob for email attachments.
 * Priority: 1. Drive Logo, 2. Fallback URL (Centralized logic).
 * @return {Blob} The logo image blob.
 */
function getLogoBlob() {
  var settings = getSystemSettings();
  
  // 1. Try Custom Uploaded Logo from Drive
  if (settings.systemLogoId) {
    try {
      return DriveApp.getFileById(settings.systemLogoId).getBlob().setName("logo");
    } catch(e) {
      console.warn("Drive logo fetch failed, proceeding to fallback: " + e.message);
    }
  }
  
  // 2. Use Centralized Fallback URL 
  // (The Imgur URL failsafe is managed once in Settings.gs:getSystemSettings)
  try {
    return UrlFetchApp.fetch(settings.fallbackLogoUrl).getBlob().setName("logo");
  } catch(e) {
    console.error("Critical: All logo blob fetches failed: " + e.message);
    // Return an empty transparent pixel or empty blob to prevent MailApp crash
    return Utilities.newBlob("", "image/png", "logo");
  }
}

/**
 * Sends an automated email based on a Trigger Event mapped in the Templates database.
 * Merges user data, applies Sandbox overrides, and attaches branding.
 * @param {string} triggerName - The name of the trigger event.
 * @param {string} toEmail - The intended recipient's email address.
 * @param {object} dataMap - Key-value pairs for {{placeholders}}.
 */
function sendTriggerEmail(triggerName, toEmail, dataMap) {
  var sheet = getMainDb().getSheetByName("Templates");
  var data = sheet.getDataRange().getValues();
  
  // Fetch system-wide branding and environment settings
  var settings = getSystemSettings();
  var sysName = settings.systemName;
  
  var subject = "";
  var templateFound = false;
  var templateIdx = -1;

  for (var i = 1; i < data.length; i++) {
    var rowTrigger = data[i][5];
    var rowStatus = data[i][8];
    
    if (rowTrigger === triggerName && rowStatus === "Active") {
      subject = data[i][6];
      templateIdx = i;
      templateFound = true;
      break;
    }
  }

  if (!templateFound) {
    console.log("No active template found for trigger: " + triggerName);
    return;
  }

  var logoBlob = getLogoBlob();
  var rawHtml = data[templateIdx][7];
  var wrapperType = data[templateIdx][9]; 

  // 1. Prepare Layout
  var wrapperHtml = getWrapperContent(wrapperType);
  var fullLayoutHtml = wrapperHtml.replace("{{USER_MESSAGE_CONTENT}}", rawHtml);

  // 2. Perform Placeholder Swap
  var finalSubject = subject;
  var finalHtml = fullLayoutHtml;
  
  for (var key in dataMap) {
    var regex = new RegExp("\\{\\{" + key + "\\}\\}", "gi");
    var replacement = dataMap[key] || "";
    finalSubject = finalSubject.replace(regex, replacement);
    finalHtml = finalHtml.replace(regex, replacement);
  }

  // 3. Sandbox Environment Interceptor
  var finalToEmail = toEmail;

  if (settings.environment === 'Sandbox' && settings.adminEmail !== '') {
    finalToEmail = settings.adminEmail;
    finalSubject = "[Sandbox Mail] " + finalSubject;
    
    // Aggressive "program code" block for the override notification
    var sandboxWarning = "<br><br><div style='padding: 20px; background-color: #000; color: #0f0; font-family: \"Courier New\", Courier, monospace; font-size: 14px; border: 2px solid #333; margin-top: 50px;'>";
    sandboxWarning += "=========================================<br>";
    sandboxWarning += " SYSTEM OVERRIDE: SANDBOX ENVIRONMENT    <br>";
    sandboxWarning += "=========================================<br>";
    sandboxWarning += "&gt; STATUS: INTERCEPTED<br>";
    sandboxWarning += "&gt; INTENDED RECIPIENT(S): " + toEmail + "<br>";
    sandboxWarning += "&gt; REROUTED TO ADMIN: " + settings.adminEmail + "<br>";
    sandboxWarning += "=========================================";
    sandboxWarning += "</div>";
    
    finalHtml += sandboxWarning;
  }

  // 4. Dispatch Email
  MailApp.sendEmail({
    to: finalToEmail,
    subject: finalSubject,
    htmlBody: finalHtml,
    noReply: true,
    name: sysName,
    inlineImages: {
      logo: logoBlob 
    }
  });
}

/**
 * Sends a test email with dummy data for template verification.
 * @param {number} rowIndex - Template row index.
 * @param {string} testEmail - Recipient for the test.
 * @return {string} Success or failure message.
 */
function sendTestEmailAction(rowIndex, testEmail) {
  var sheet = getMainDb().getSheetByName("Templates");
  var data = sheet.getDataRange().getValues();
  
  if (rowIndex < 1 || rowIndex >= data.length) return "Error: Template not found.";
  
  var settings = getSystemSettings();
  var rowData = data[rowIndex];
  var subject = rowData[6] || "No Subject";
  var rawHtml = rowData[7] || "";
  var wrapperType = rowData[9] || "Internal";
  
  var wrapperHtml = getWrapperContent(wrapperType);
  var fullLayoutHtml = wrapperHtml.replace("{{USER_MESSAGE_CONTENT}}", rawHtml);

  // Comprehensive Dummy Data
  var dummyData = {
    "companyName": "Acme Corp (Test)",
    "brandName": "Acme Brand",
    "address": "123 Test Ave, Suite 100",
    "website": "www.megarhino.com",
    "priFirstName": "John",
    "priLastName": "Doe",
    "priEmail": "john@example.com",
    "monthlyContractValue": "$2,500",
    "contractStartDate": "2026-05-01",
    "services": "SEO & Content Marketing",
    "notes": "Sample test note.",
    "username": "jdoe",
    "firstName": "John",
    "lastName": "Doe",
    "role": "Account Manager"
  };

  var finalSubject = "[TEST] " + subject;
  var finalHtml = fullLayoutHtml;

  for (var key in dummyData) {
    var regex = new RegExp("\\{\\{" + key + "\\}\\}", "gi");
    finalSubject = finalSubject.replace(regex, dummyData[key]);
    finalHtml = finalHtml.replace(regex, dummyData[key]);
  }

  try {
    MailApp.sendEmail({
      to: testEmail,
      subject: finalSubject,
      htmlBody: finalHtml,
      noReply: true,
      name: settings.systemName,
      inlineImages: {
        logo: getLogoBlob()
      }
    });
    return "Test email successfully sent to " + testEmail;
  } catch (e) {
    return "Failed to send test email: " + e.message;
  }
}