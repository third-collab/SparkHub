function patchRolesToUsersModule() {
  var db = getMainDb();
  var tSheet = db.getSheetByName("Templates");
  var tData = tSheet.getDataRange().getValues();
  
  for (var i = 1; i < tData.length; i++) {
    var trigger = tData[i][6]; // Trigger is Index 6
    var module = tData[i][5];  // Module is Index 5
    
    // Convert Roles module to Users module to reflect Access structure
    if (module === "Roles" || trigger.startsWith("Roles:")) {
      tSheet.getRange(i + 1, 6).setValue("Users");
    }
  }
  console.log("Templates successfully patched! Roles module converted to Users module.");
}