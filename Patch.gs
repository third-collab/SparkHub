/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Patch.gs
 * VERSION: 1.1
 * SYNC STATUS: Standalone Utility
 */

function patchDashboardConfigColumns() {
  var db = getMainDb();
  var usersSheet = db.getSheetByName("Users");
  
  if (usersSheet.getRange("J1").getValue() !== "Dashboard Config") {
    usersSheet.getRange("J1").setValue("Dashboard Config").setFontWeight("bold");
  }
  
  var rolesSheet = db.getSheetByName("Roles");
  if (rolesSheet.getRange("G1").getValue() !== "Dashboard Config") {
    rolesSheet.getRange("G1").setValue("Dashboard Config").setFontWeight("bold");
  }
  console.log("Dashboard Config columns successfully patched into Users and Roles tables!");
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */