/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Forms.gs
 * VERSION: 1.0
 * SYNC STATUS: Fully Synchronized with FormsData.html
 */

// ============================================================================
// 1. REGISTRY EXPORTS (V8 Auto-Discovery)
// ============================================================================

function Forms_getTriggers() {
  return [
    "Forms:TEMPLATE_CREATED",
    "Forms:INSTANCE_GENERATED",
    "Forms:RESPONSE_SUBMITTED"
  ];
}

function Forms_getPlaceholders() {
  return [
    "formTitle",
    "formLink",
    "formDueDate",
    "respondentName"
  ];
}

function Forms_getPermissions() {
  return [
    "View Forms",
    "Manage Form Templates",
    "Submit Forms",
    "View Form Responses"
  ];
}

// ============================================================================
// 2. CORE PROCESSORS & EVENT HANDLERS
// ============================================================================

/**
 * Resolves dynamic lookup sources for form inputs (e.g. SYS_LOOKUP:users)
 * and handles self-exclusion rules (e.g. excluding the logged-in respondent).
 */
function Forms_resolveDynamicOptions(optionSource, excludeCurrentUser, currentUserEmail) {
  try {
    var options = [];
    if (!optionSource || typeof optionSource !== "string") {
      return { success: true, options: [] };
    }

    if (optionSource === "SYS_LOOKUP:users") {
      var usersList = getUsersList(); // Core Users getter
      if (Array.isArray(usersList)) {
        usersList.forEach(function(user) {
          // Exclude current respondent if excludeCurrentUser flag is set
          if (excludeCurrentUser && currentUserEmail && user.email && user.email.toLowerCase() === currentUserEmail.toLowerCase()) {
            return;
          }
          if (user.status === "Active") {
            options.push({
              value: user.email || user.id,
              label: (user.firstName && user.lastName) ? (user.firstName + " " + user.lastName) : user.username
            });
          }
        });
      }
    } else if (optionSource === "SYS_LOOKUP:clients") {
      var clientsList = getClientsList(); // Core Clients getter if present
      if (Array.isArray(clientsList)) {
        clientsList.forEach(function(client) {
          if (client.status === "Active") {
            options.push({
              value: client.id,
              label: client.clientName || client.company
            });
          }
        });
      }
    }

    return { success: true, options: options };
  } catch (e) {
    return { error: "Failed to resolve dynamic options: " + e.message };
  }
}

/**
 * Scheduled processor for automated form generation routines (e.g. Monthly 25th Peer Recognition).
 * Triggered by daily system scheduler.
 */
function Forms_processScheduledGenerations() {
  try {
    var today = new Date();
    var dayOfMonth = today.getDate();
    
    var templatesResult = getFormTemplatesList();
    if (templatesResult.error) throw new Error(templatesResult.error);
    var templates = templatesResult.data || [];

    templates.forEach(function(template) {
      if (template.status !== "Active") return;

      if (template.scheduleRule === "MONTHLY_25" && dayOfMonth === 25) {
        // Fetch all active users to generate individual instances
        var users = getUsersList();
        if (Array.isArray(users)) {
          users.forEach(function(user) {
            if (user.status !== "Active") return;

            var token = "FI-" + Utilities.getUuid();
            var dueDate = new Date();
            dueDate.setDate(dueDate.getDate() + 7); // 7-day completion window

            saveFormInstanceRecord({
              templateId: template.id,
              targetRespondent: user.email,
              publicToken: token,
              dueDate: dueDate.toISOString().split("T")[0],
              metadata: { source: "SCHEDULED_MONTHLY_25" },
              status: "Pending"
            });
          });
        }
      }
    });

    return { success: true, message: "Scheduled form generations processed." };
  } catch (e) {
    return { error: "Scheduled generation failed: " + e.message };
  }
}

// ============================================================================
// 3. READ / GET FUNCTIONS
// ============================================================================

function getFormTemplatesList() {
  try {
    var sheets = ensureFormsSheets();
    var sheet = sheets.templatesSheet;
    var data = sheet.getDataRange().getValues();
    var templates = [];

    if (data.length > 1) {
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var questionsJson = {};
        try { questionsJson = row[23] ? JSON.parse(row[23]) : []; } catch(err) {}

        templates.push({
          rowIndex: i + 1,
          timestamp: row[0],
          id: row[1],
          title: row[2],
          category: row[3],
          scheduleRule: row[4],
          questions: questionsJson,
          status: row[25]
        });
      }
    }
    return { success: true, data: templates };
  } catch (e) {
    return { error: "Failed to fetch form templates: " + e.message };
  }
}

function getFormInstancesList(filterEmail) {
  try {
    var sheets = ensureFormsSheets();
    var sheet = sheets.instancesSheet;
    var data = sheet.getDataRange().getValues();
    var instances = [];

    if (data.length > 1) {
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var metadata = {};
        try { metadata = row[23] ? JSON.parse(row[23]) : {}; } catch(err) {}

        if (filterEmail && row[3] && row[3].toLowerCase() !== filterEmail.toLowerCase()) {
          continue;
        }

        instances.push({
          rowIndex: i + 1,
          timestamp: row[0],
          id: row[1],
          templateId: row[2],
          targetRespondent: row[3],
          publicToken: row[4],
          dueDate: row[5],
          metadata: metadata,
          status: row[25]
        });
      }
    }
    return { success: true, data: instances };
  } catch (e) {
    return { error: "Failed to fetch form instances: " + e.message };
  }
}

function getPublicFormInstanceByToken(token) {
  try {
    if (!token) return { error: "Access token is required." };

    var sheets = ensureFormsSheets();
    var instanceSheet = sheets.instancesSheet;
    var instData = instanceSheet.getDataRange().getValues();
    var matchedInstance = null;

    for (var i = 1; i < instData.length; i++) {
      if (instData[i][4] === token) {
        matchedInstance = {
          rowIndex: i + 1,
          id: instData[i][1],
          templateId: instData[i][2],
          targetRespondent: instData[i][3],
          publicToken: instData[i][4],
          dueDate: instData[i][5],
          status: instData[i][25]
        };
        break;
      }
    }

    if (!matchedInstance) return { error: "Invalid or expired form link." };
    if (matchedInstance.status === "Completed") return { error: "This form has already been submitted." };

    // Fetch parent template details
    var templatesResult = getFormTemplatesList();
    if (templatesResult.error) throw new Error(templatesResult.error);
    var parentTemplate = (templatesResult.data || []).find(function(t) { return t.id === matchedInstance.templateId; });

    if (!parentTemplate) return { error: "Associated form template not found." };

    return {
      success: true,
      instance: matchedInstance,
      template: parentTemplate
    };
  } catch (e) {
    return { error: "Failed to load public form: " + e.message };
  }
}

function getFormResponsesList(templateId) {
  try {
    var sheets = ensureFormsSheets();
    var sheet = sheets.responsesSheet;
    var data = sheet.getDataRange().getValues();
    var responses = [];

    if (data.length > 1) {
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (templateId && row[2] !== templateId) continue;

        var answersJson = {};
        try { answersJson = row[23] ? JSON.parse(row[23]) : {}; } catch(err) {}

        responses.push({
          rowIndex: i + 1,
          timestamp: row[0],
          id: row[1],
          instanceId: row[2],
          respondent: row[3],
          answers: answersJson,
          status: row[25]
        });
      }
    }
    return { success: true, data: responses };
  } catch (e) {
    return { error: "Failed to fetch form responses: " + e.message };
  }
}

// ============================================================================
// 4. WRITE / SAVE FUNCTIONS
// ============================================================================

function saveFormTemplateRecord(payload) {
  try {
    if (!payload.title) return { error: "Form Title is required." };

    var sheets = ensureFormsSheets();
    var sheet = sheets.templatesSheet;
    var templateId = payload.id || ("FT-" + Math.floor(1000 + Math.random() * 9000));
    var questionsJsonString = JSON.stringify(payload.questions || []);
    var status = payload.status || "Active";
    var targetRow = -1;

    if (payload.rowIndex && payload.rowIndex > 1) {
      targetRow = payload.rowIndex;
      sheet.getRange(targetRow, 3).setValue(payload.title);
      sheet.getRange(targetRow, 4).setValue(payload.category || "General");
      sheet.getRange(targetRow, 5).setValue(payload.scheduleRule || "NONE");
      sheet.getRange(targetRow, 24).setValue(questionsJsonString);
      sheet.getRange(targetRow, 26).setValue(status);
    } else {
      var newRow = new Array(26).fill("");
      newRow[0] = new Date();
      newRow[1] = templateId;
      newRow[2] = payload.title;
      newRow[3] = payload.category || "General";
      newRow[4] = payload.scheduleRule || "NONE";
      newRow[23] = questionsJsonString; // Col 24 (X)
      newRow[25] = status;              // Col 26 (Z)
      sheet.appendRow(newRow);
      targetRow = sheet.getLastRow();
    }

    SpreadsheetApp.flush();

    SystemEvent.emit("Forms", "MUTATION", "Save Template", "INFO", templateId, "Saved form template: " + payload.title, "", {});

    return { success: true, rowIndex: targetRow, id: templateId, message: "Form Template saved successfully." };
  } catch (e) {
    return { error: "Failed to save form template: " + e.message };
  }
}

function saveFormInstanceRecord(payload) {
  try {
    if (!payload.templateId) return { error: "Template ID is required." };

    var sheets = ensureFormsSheets();
    var sheet = sheets.instancesSheet;
    var instanceId = payload.id || ("FI-" + Math.floor(1000 + Math.random() * 9000));
    var token = payload.publicToken || ("FI-" + Utilities.getUuid());
    var metadataString = JSON.stringify(payload.metadata || {});
    var status = payload.status || "Pending";

    var newRow = new Array(26).fill("");
    newRow[0] = new Date();
    newRow[1] = instanceId;
    newRow[2] = payload.templateId;
    newRow[3] = payload.targetRespondent || "OPEN";
    newRow[4] = token;
    newRow[5] = payload.dueDate || "";
    newRow[23] = metadataString; // Col 24 (X)
    newRow[25] = status;         // Col 26 (Z)

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();

    SystemEvent.emit("Forms", "INSTANCE_GENERATED", "Generate Instance", "INFO", instanceId, "Generated instance for template: " + payload.templateId, payload.targetRespondent, {});

    return { success: true, instanceId: instanceId, token: token, message: "Form instance generated successfully." };
  } catch (e) {
    return { error: "Failed to generate form instance: " + e.message };
  }
}

function submitFormResponseRecord(payload) {
  try {
    if (!payload.instanceId) return { error: "Instance ID is required." };

    var sheets = ensureFormsSheets();
    var responsesSheet = sheets.responsesSheet;
    var instancesSheet = sheets.instancesSheet;

    var responseId = "FR-" + Math.floor(1000 + Math.random() * 9000);
    var answersString = JSON.stringify(payload.answers || {});

    var newRow = new Array(26).fill("");
    newRow[0] = new Date();
    newRow[1] = responseId;
    newRow[2] = payload.instanceId;
    newRow[3] = payload.respondent || "ANONYMOUS";
    newRow[23] = answersString; // Col 24 (X)
    newRow[25] = "Submitted";   // Col 26 (Z)

    responsesSheet.appendRow(newRow);

    // Update Instance Status to Completed if rowIndex passed
    if (payload.instanceRowIndex && payload.instanceRowIndex > 1) {
      instancesSheet.getRange(payload.instanceRowIndex, 26).setValue("Completed");
    }

    SpreadsheetApp.flush();

    SystemEvent.emit("Forms", "RESPONSE_SUBMITTED", "Submit Response", "INFO", responseId, "Form response submitted for instance: " + payload.instanceId, payload.respondent, {});

    return { success: true, responseId: responseId, message: "Response submitted successfully!" };
  } catch (e) {
    return { error: "Failed to submit form response: " + e.message };
  }
}

// ============================================================================
// 5. INTERNAL HELPERS
// ============================================================================

function ensureFormsSheets() {
  var settings = getSystemSettings();
  var ss = SpreadsheetApp.openById(settings.mainDatabaseId);

  var templatesSheet = ss.getSheetByName("FORMS_TEMPLATES");
  if (!templatesSheet) {
    templatesSheet = ss.insertSheet("FORMS_TEMPLATES");
    var headers = new Array(26).fill("");
    headers[0] = "Timestamp";
    headers[1] = "Template ID";
    headers[2] = "Title";
    headers[3] = "Category";
    headers[4] = "Schedule Rule";
    headers[23] = "Questions JSON";
    headers[25] = "Status";
    templatesSheet.appendRow(headers);
  }

  var instancesSheet = ss.getSheetByName("FORMS_INSTANCES");
  if (!instancesSheet) {
    instancesSheet = ss.insertSheet("FORMS_INSTANCES");
    var instHeaders = new Array(26).fill("");
    instHeaders[0] = "Timestamp";
    instHeaders[1] = "Instance ID";
    instHeaders[2] = "Template ID";
    instHeaders[3] = "Target Respondent";
    instHeaders[4] = "Public Token";
    instHeaders[5] = "Due Date";
    instHeaders[23] = "Metadata JSON";
    instHeaders[25] = "Status";
    instancesSheet.appendRow(instHeaders);
  }

  var responsesSheet = ss.getSheetByName("FORMS_RESPONSES");
  if (!responsesSheet) {
    responsesSheet = ss.insertSheet("FORMS_RESPONSES");
    var respHeaders = new Array(26).fill("");
    respHeaders[0] = "Timestamp";
    respHeaders[1] = "Response ID";
    respHeaders[2] = "Instance ID";
    respHeaders[3] = "Respondent";
    respHeaders[23] = "Answers JSON";
    respHeaders[25] = "Status";
    responsesSheet.appendRow(respHeaders);
  }

  return {
    templatesSheet: templatesSheet,
    instancesSheet: instancesSheet,
    responsesSheet: responsesSheet
  };
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */