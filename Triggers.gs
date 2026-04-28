/**
 * Run this function ONCE manually from the editor to set up the daily automation.
 * It tells Google to run the 'checkDailyEvents' function every day between 8 AM and 9 AM.
 */
function setupDailyTriggers() {
  // Clear existing triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "checkDailyEvents") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create the new daily trigger
  ScriptApp.newTrigger("checkDailyEvents")
    .timeBased()
    .everyDays(1)
    .atHour(8) // Runs between 8 AM and 9 AM
    .create();
  console.log("Daily trigger setup complete!");
}

/**
 * The main engine that scans for dates.
 * Orchestrates the scanning of both databases for relevant events.
 */
function checkDailyEvents() {
  var today = new Date();
  var currentMonth = today.getMonth(); // 0-indexed (Jan = 0)
  var currentDate = today.getDate();

  checkUserEvents(currentMonth, currentDate);
}

/**
 * Scans the Users sheet for Birthdays and Work Anniversaries.
 * Uses indices based on the Users sheet structure to log system notifications.
 * @param {number} currentMonth - The current month (0-11).
 * @param {number} currentDate - The current day of the month (1-31).
 */
function checkUserEvents(currentMonth, currentDate) {
  // Uses the Global Helper from Config.gs
  var sheet = getMainDb().getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  var today = new Date();
  
  for (var i = 1; i < data.length; i++) {
    var username = data[i][2];
    var userEmail = data[i][5];
    var birthday = new Date(data[i][7]); 
    var hireDate = new Date(data[i][8]);

    // 1. Check User Birthday
    if (isValidDate(birthday) && birthday.getMonth() === currentMonth && birthday.getDate() === currentDate) {
      // Notification for everyone else: Hidden from the birthday person
      logNotification("Users", "User Birthday", "It's " + username + "'s birthday today! 🎂", "All", username);
      // Notification JUST for the birthday person
      logNotification("Users", "User Birthday", "Happy Birthday! 🎉 Have a great day!", username, "All");
    }

    // 2. Check User Work Anniversary
    if (isValidDate(hireDate) && hireDate.getMonth() === currentMonth && hireDate.getDate() === currentDate) {
      var years = today.getFullYear() - hireDate.getFullYear();
      if (years > 0) {
        logNotification("Users", "User Anniversary", username + " is celebrating " + years + " year(s) with MegaRhino! 🎈", "All", "");
      }
    }
  }
}

/**
 * Helper to ensure the cell actually contains a valid Date object before checking it.
 * @param {any} d - The value to check.
 * @return {boolean} - Returns true if the value is a valid Date object.
 */
function isValidDate(d) {
  return d instanceof Date && !isNaN(d);
}