function hardResetTemplatesAndWrappers() {
  var db = getMainDb();
  var now = new Date();

  // =========================================================
  // 1. REBUILD WRAPPERS REGISTRY
  // =========================================================
  var wSheet = db.getSheetByName("Wrappers");
  if (wSheet) {
    wSheet.clear(); // Wipe the corrupted data
    
    // Set exact 6-column schema
    wSheet.getRange(1, 1, 1, 6).setValues([["Timestamp", "Wrapper ID", "Name", "Description", "HTML Content", "Status"]]).setFontWeight("bold");
    wSheet.setFrozenRows(1);

    var internalHtml = `<div style="background-color: #f4f6f9; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;"><div style="background-color: #323232; padding: 25px; text-align: center; border-bottom: 4px solid #F1C404;"><img src="cid:logo" alt="Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"></div><div style="padding: 30px; color: #444; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; border-top: 1px solid #eee; background-color: #fcfcfc; text-align: center; font-size: 11px; color: #888;">This is an automated system notification.</div></div></div>`;
    var externalHtml = `<div style="background-color: #ffffff; padding: 40px 20px; font-family: Arial, sans-serif; border: 1px solid #eee;"><div style="max-width: 600px; margin: 0 auto;"><div style="padding-bottom: 20px; border-bottom: 1px solid #ddd; margin-bottom: 20px; text-align: center;"><img src="cid:logo" alt="Logo" style="max-width: 150px; height: auto; margin-bottom: 10px;"></div><div style="color: #555; line-height: 1.6;">{{USER_MESSAGE_CONTENT}}</div><div style="margin-top: 40px; font-size: 12px; color: #999; border-top: 1px solid #eee; padding-top: 15px;">Sent from the Hub Team.</div></div></div>`;
    var userHtml = `<div style="background-color: #f8fafc; padding: 40px 20px; font-family: sans-serif;"><div style="max-width: 500px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px rgba(0,0,0,0.05);"><div style="padding: 30px; text-align: center; border-bottom: 1px solid #f1f5f9;"><img src="cid:logo" alt="Logo" style="max-width: 120px; height: auto;"></div><div style="padding: 30px; color: #334155; line-height: 1.6; font-size: 15px;">{{USER_MESSAGE_CONTENT}}</div><div style="padding: 20px; background-color: #f8fafc; text-align: center; font-size: 12px; color: #94a3b8; border-top: 1px solid #e2e8f0;">Security & Access Notification</div></div></div>`;

    wSheet.appendRow([now, "W-INT", "Internal", "Standard internal messaging", internalHtml, "Active"]);
    wSheet.appendRow([now, "W-EXT", "External", "Client-facing messaging", externalHtml, "Active"]);
    wSheet.appendRow([now, "W-USER", "User Communications", "Dedicated layout for user access and security emails", userHtml, "Active"]);
  }

  // =========================================================
  // 2. REBUILD TEMPLATES REGISTRY
  // =========================================================
  var tSheet = db.getSheetByName("Templates");
  if (tSheet) {
    tSheet.clear(); // Wipe the corrupted data
    
    // Set exact 11-column schema
    tSheet.getRange(1, 1, 1, 11).setValues([["Timestamp", "ID", "Name", "Description", "Category", "Module", "Trigger", "Subject", "Body", "Wrapper", "Status"]]).setFontWeight("bold");
    tSheet.setFrozenRows(1);

    var installHtml = `<div style='font-family: sans-serif; padding: 20px;'><h2>Installation Successful</h2><p>SparkHub has been fully deployed and is ready for use.</p></div>`;
    var welcomeHtml = `<div style='font-family: sans-serif; padding: 20px;'><h2>Welcome to SparkHub</h2><p>Hello {{username}},</p><p>Your account is ready. You can now access your workspace using your system credentials.</p></div>`;
    var roleHtml = `<div style='font-family: sans-serif; padding: 20px;'><h2>New System Role</h2><p>The system role <strong>{{username}}</strong> has been successfully established.</p><p>Details: {{details}}</p></div>`;
    var resetHtml = `<div style='font-family: sans-serif; padding: 20px;'><h2>Password Reset Request</h2><p>Hi {{userFirst}},</p><p>We received a request to reset your local password. Click the link below to set a new password. This link will expire in 15 minutes.</p><a href='{{resetLink}}' style='display:inline-block; padding: 10px 20px; background: #c40004; color: white; text-decoration: none; border-radius: 5px; margin-top: 15px;'>Reset Password</a></div>`;
    var updatedHtml = `<div style='font-family: sans-serif; padding: 20px;'><h2>Password Updated</h2><p>Hi {{userFirst}},</p><p>This is a confirmation that your system password has been successfully updated. If you did not make this change, please contact your administrator immediately.</p></div>`;

    // Strict array mapping: TS(0), ID(1), Name(2), Desc(3), Cat(4), Module(5), Trigger(6), Subject(7), Body(8), Wrapper(9), Status(10)
    tSheet.appendRow([now, "TPL-INSTALL", "System Installed", "System installation confirmation", "System", "System", "System:INSTALL", "Installation Successful", installHtml, "Internal", "Active"]);
    tSheet.appendRow([now, "TPL-USER-NEW", "User Welcome", "Access email", "Security", "Users", "Users:CREATE", "Welcome to SparkHub", welcomeHtml, "User Communications", "Active"]);
    tSheet.appendRow([now, "TPL-ROLE-NEW", "Role Created", "Role creation alert", "Security", "Roles", "Roles:CREATE", "New System Role: {{username}}", roleHtml, "Internal", "Active"]);
    tSheet.appendRow([now, "TPL-PWD-RESET", "Password Reset Link", "Forgot password link", "Security", "Users", "Users:RESET_REQUEST", "Password Reset Request", resetHtml, "User Communications", "Active"]);
    tSheet.appendRow([now, "TPL-PWD-UPDATE", "Password Updated", "Password change confirmation", "Security", "Users", "Users:PASSWORD_UPDATED", "Security Alert: Password Updated", updatedHtml, "User Communications", "Active"]);
  }

  console.log("HARD RESET COMPLETE! Templates and Wrappers have been restored to perfect structural alignment.");
}