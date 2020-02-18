/***************************************************
 * Automated Lighthouse Report
 * template by james@upbuild.io
 * built for JML by tronan@jmclaughlin.com
 ***************************************************/

// ================== //
// -  Trigger Times - //
// ================== //
// runTool || runLog  //
// --------||-------- //
//  9-10   ||         //
//  10-11  || 11-12   //
//  12-1   ||         //
//  1-2    || 2-3     //
//  4-5    ||         //
//  5-6    || 6-7     //
//         ||         //
// ================== //


function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
      name: "Set Report & Log Schedule",
      functionName: "scheduleboth"
    },
    {
      name: "Manual Push Report",
      functionName: "runTool"
    },
    {
      name: "Manual Push Log",
      functionName: "runLog"
    },
    {
      name: "Reset Schedule",
      functionName: "resetSchedule"
    }
  ];
  sheet.addMenu("PageSpeed Menu", entries);

  // Add JML Menu to run custom schedule
  var dailySchedule = [{
      name: "Set Daily Schedule",
      functionName: "dailyTrigger"
    },
    {
      name: "Clear Schedule",
      functionName: "resetSchedule"
    }
  ];
  sheet.addMenu("JML Menu", dailySchedule)
}

function resetsuccess() {
  Browser.msgBox('Success! - Report and Log Times Reset');
}

function resetFailure() {
  Browser.msgBox("Your schedule is already empty.")
}

function scheduleConfirm() {
  Browser.msgBox("Reports Scheduled!")
}

// Deletes all triggers in the current project.
function resetSchedule() {
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.length > 0) {
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    resetsuccess();
  } else {
    resetFailure();
  }
}

// Set Default Triggers
function scheduleboth() {
  startScheduledReportOne();
  startScheduledReportTwo();
  startScheduledReportThree();
  startScheduledReportFour();
  startScheduledLog();
  scheduleConfirm();
}

// ***************
// Schedule Custom Triggers
// ***************

function dailyTrigger() {
  //  ********** Reports *********** //
  createDailyReportTriggerMorningOne();
  createDailyReportTriggerMorningTwo();
  createDailyReportTriggerAfternoonOne();
  createDailyReportTriggerAfternoonTwo();
  createDailyReportTriggerEveningOne();
  createDailyReportTriggerEveningTwo();
  //  *********** Logs ************ //
  createDailyLogTriggerMorning();
  createDailyLogTriggerAfternoon();
  createDailyLogTriggerEvening();
  scheduleConfirm();
}

// ***************
// Create Triggers for Reports (3x/day) ##################################################
// ***************

function createDailyReportTriggerMorningOne() {
  ScriptApp.newTrigger('runTool')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
}

function createDailyReportTriggerMorningTwo() {
  ScriptApp.newTrigger('runTool')
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .create();
}

//######################

function createDailyReportTriggerAfternoonOne() {
  ScriptApp.newTrigger('runTool')
    .timeBased()
    .atHour(12)
    .everyDays(1)
    .create();
}

function createDailyReportTriggerAfternoonTwo() {
  ScriptApp.newTrigger('runTool')
    .timeBased()
    .atHour(13)
    .everyDays(1)
    .create();
}

//######################

function createDailyReportTriggerEveningOne() {
  ScriptApp.newTrigger('runTool')
    .timeBased()
    .atHour(16)
    .everyDays(1)
    .create();
}

function createDailyReportTriggerEveningTwo() {
  ScriptApp.newTrigger('runTool')
    .timeBased()
    .atHour(17)
    .everyDays(1)
    .create();
}

// ***************
// Create Triggers for Logs (1 hour after each report trigger) ##################################################
// ***************

function createDailyLogTriggerMorning() {
  ScriptApp.newTrigger('runLog')
    .timeBased()
    .atHour(11)
    .everyDays(1)
    .create();
}

function createDailyLogTriggerAfternoon() {
  ScriptApp.newTrigger('runLog')
    .timeBased()
    .atHour(14)
    .everyDays(1)
    .create();
}

function createDailyLogTriggerEvening() {
  ScriptApp.newTrigger('runLog')
    .timeBased()
    .atHour(18)
    .everyDays(1)
    .create();
}

//==============

//Send email on log

function emailLink(email, subject, date, time) {
  var message = "<span style='margin-bottom:10px;'><p style='font-size:16px;'>Your latest Lighthouse report has finished updating!<br><br> Please visit one of the links below to either view the results in your browser, or download as a csv file:</p><br></span>";
  var liveLink = "<a style='width:125px;margin-right:20px;margin-bottom:5px;background-color:#ffffff;color:#51bceb; border:2px solid #51bceb;text-decoration:none;border-radius:6px;padding:14px 25px;text-transform:uppercase;text-align:center;font-weight:bolder;display:inline-block;'href=\'<<< ADD LINK TO BROWSER LINK HERE --- >>>'>View in Browser</a>"
  var downloadCsv = "<a style='width:125px;margin-bottom:5px;background-color:#ffffff;color:#51bceb;border:2px solid #51bceb;text-decoration:none;border-radius:6px;padding:14px 25px;text-transform:uppercase;text-align:center;font-weight:bolder;display:inline-block;'href=\'<<< ADD LINK TO CSV DOWNLOAD HERE --- >>>'>Download</a>"
  var jmlLogo = "<<< ADD LINK TO IMAGE HERE --- >>>"
  var embedLogo = "<a href='<<< ADD LINK TO STORE HERE >>>'><img height='20%' width='20%' src='cid:jmlLogo'></a>"
  var jmlLogoBlob = UrlFetchApp
                         .fetch(jmlLogo)
                         .getBlob()
                         .setName("jmlLogoBlob");
    MailApp.sendEmail({
    to: email,
    subject: subject + date + "-" + time,
      htmlBody: (message + liveLink + downloadCsv + "<br><br>" + "<hr style='margin-bottom:15px;'>" + embedLogo),
    inlineImages: {jmlLogo: jmlLogoBlob}
  })
    Browser.msgBox("Email sent to " + email);
}

function triggerOnEdit(e) {
  showMessageOnUpdate(e);
}

function validateEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)
}

function sendEmail(email, subject, date, time, content) {
  MailApp.sendEmail({
    to: email,
    subject: subject + date + "-" + time,
    htmlBody: content
  })
}

function showMessageOnUpdate(e) {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert("Report Processed!", "Would you like an emailed copy of results?", ui.ButtonSet.YES_NO);
  if (resp == ui.Button.YES) {

    var getUserEmail = ui.prompt("What address would you like them sent to?");
    var userEmail = getUserEmail.getResponseText();
    var d = new Date();
    var currentTime = d.toLocaleTimeString().replace(/:\d{2}\s/, ' ').slice(0, 8);
    var date = Utilities.formatDate(new Date(), "GMT-5", "MM/dd ");
    var subjectLine = "Lighthouse Report Results for "
    getUserEmail;

    if (validateEmail(userEmail) == true) {
      emailLink(userEmail, subjectLine, date, currentTime);

    } else {
      var failed = ui.alert("Sorry, we weren't able to verify your email address.");

      if (failed = ui.Button.OK) {
        var tryAgain = ui.prompt("Please try again");
        var newEmail = tryAgain.getResponseText();

        if (validateEmail(newEmail) == true) {
          emailLink(newEmail, subjectLine, date, currentTime);
        }
      }
    }
  }
}

// ***************
// Run Scheduled Reports via Settings Tab  ##################################################
// ***************


//Run the Report - Phase One
function startScheduledReportOne() {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var reportDay = settingsSheet.getRange("C10:C10").getValue();
  var reportTime = settingsSheet.getRange("E10:E10").getValue();

  if (reportTime == "1AM") {
    var newhour = "1"
  }
  if (reportTime == "2AM") {
    var newhour = "2"
  }
  if (reportTime == "3AM") {
    var newhour = "3"
  }
  if (reportTime == "4AM") {
    var newhour = "4"
  }
  if (reportTime == "5AM") {
    var newhour = "5"
  }
  if (reportTime == "6AM") {
    var newhour = "6"
  }
  if (reportTime == "7AM") {
    var newhour = "7"
  }
  if (reportTime == "8AM") {
    var newhour = "8"
  }
  if (reportTime == "9AM") {
    var newhour = "9"
  }
  if (reportTime == "10AM") {
    var newhour = "10"
  }
  if (reportTime == "11AM") {
    var newhour = "11"
  }
  if (reportTime == "12PM") {
    var newhour = "12"
  }
  if (reportTime == "1PM") {
    var newhour = "13"
  }
  if (reportTime == "2PM") {
    var newhour = "14"
  }
  if (reportTime == "3PM") {
    var newhour = "15"
  }
  if (reportTime == "4PM") {
    var newhour = "16"
  }
  if (reportTime == "5PM") {
    var newhour = "17"
  }
  if (reportTime == "6PM") {
    var newhour = "18"
  }
  if (reportTime == "7PM") {
    var newhour = "19"
  }
  if (reportTime == "8PM") {
    var newhour = "20"
  }
  if (reportTime == "9PM") {
    var newhour = "21"
  }
  if (reportTime == "10PM") {
    var newhour = "22"
  }
  if (reportTime == "11PM") {
    var newhour = "23"
  }
  if (reportTime == "12AM") {
    var newhour = "24"
  }

  if (reportDay == "MONDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .create();
  }

  if (reportDay == "TUESDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.TUESDAY)
      .create();

  }

  if (reportDay == "WEDNESDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
      .create();

  }

  if (reportDay == "THURSDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.THURSDAY)
      .create();

  }

  if (reportDay == "FRIDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.FRIDAY)
      .create();

  }

  if (reportDay == "SATURDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.SATURDAY)
      .create();

  }

  if (reportDay == "SUNDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .create();
  }
}

//Run the Report - Phase Two
function startScheduledReportTwo() {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var reportDay = settingsSheet.getRange("C11:C11").getValue();
  var reportTime = settingsSheet.getRange("E11:E11").getValue();

  if (reportTime == "1AM") {
    var newhour = "1"
  }
  if (reportTime == "2AM") {
    var newhour = "2"
  }
  if (reportTime == "3AM") {
    var newhour = "3"
  }
  if (reportTime == "4AM") {
    var newhour = "4"
  }
  if (reportTime == "5AM") {
    var newhour = "5"
  }
  if (reportTime == "6AM") {
    var newhour = "6"
  }
  if (reportTime == "7AM") {
    var newhour = "7"
  }
  if (reportTime == "8AM") {
    var newhour = "8"
  }
  if (reportTime == "9AM") {
    var newhour = "9"
  }
  if (reportTime == "10AM") {
    var newhour = "10"
  }
  if (reportTime == "11AM") {
    var newhour = "11"
  }
  if (reportTime == "12PM") {
    var newhour = "12"
  }
  if (reportTime == "1PM") {
    var newhour = "13"
  }
  if (reportTime == "2PM") {
    var newhour = "14"
  }
  if (reportTime == "3PM") {
    var newhour = "15"
  }
  if (reportTime == "4PM") {
    var newhour = "16"
  }
  if (reportTime == "5PM") {
    var newhour = "17"
  }
  if (reportTime == "6PM") {
    var newhour = "18"
  }
  if (reportTime == "7PM") {
    var newhour = "19"
  }
  if (reportTime == "8PM") {
    var newhour = "20"
  }
  if (reportTime == "9PM") {
    var newhour = "21"
  }
  if (reportTime == "10PM") {
    var newhour = "22"
  }
  if (reportTime == "11PM") {
    var newhour = "23"
  }
  if (reportTime == "12AM") {
    var newhour = "24"
  }

  if (reportDay == "MONDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .create();
  }

  if (reportDay == "TUESDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.TUESDAY)
      .create();

  }

  if (reportDay == "WEDNESDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
      .create();

  }

  if (reportDay == "THURSDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.THURSDAY)
      .create();

  }

  if (reportDay == "FRIDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.FRIDAY)
      .create();

  }

  if (reportDay == "SATURDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.SATURDAY)
      .create();

  }

  if (reportDay == "SUNDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .create();

  }
}

//Run the Report - Phase Three
function startScheduledReportThree() {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var reportDay = settingsSheet.getRange("C12:C12").getValue();
  var reportTime = settingsSheet.getRange("E12:E12").getValue();

  if (reportTime == "1AM") {
    var newhour = "1"
  }
  if (reportTime == "2AM") {
    var newhour = "2"
  }
  if (reportTime == "3AM") {
    var newhour = "3"
  }
  if (reportTime == "4AM") {
    var newhour = "4"
  }
  if (reportTime == "5AM") {
    var newhour = "5"
  }
  if (reportTime == "6AM") {
    var newhour = "6"
  }
  if (reportTime == "7AM") {
    var newhour = "7"
  }
  if (reportTime == "8AM") {
    var newhour = "8"
  }
  if (reportTime == "9AM") {
    var newhour = "9"
  }
  if (reportTime == "10AM") {
    var newhour = "10"
  }
  if (reportTime == "11AM") {
    var newhour = "11"
  }
  if (reportTime == "12PM") {
    var newhour = "12"
  }
  if (reportTime == "1PM") {
    var newhour = "13"
  }
  if (reportTime == "2PM") {
    var newhour = "14"
  }
  if (reportTime == "3PM") {
    var newhour = "15"
  }
  if (reportTime == "4PM") {
    var newhour = "16"
  }
  if (reportTime == "5PM") {
    var newhour = "17"
  }
  if (reportTime == "6PM") {
    var newhour = "18"
  }
  if (reportTime == "7PM") {
    var newhour = "19"
  }
  if (reportTime == "8PM") {
    var newhour = "20"
  }
  if (reportTime == "9PM") {
    var newhour = "21"
  }
  if (reportTime == "10PM") {
    var newhour = "22"
  }
  if (reportTime == "11PM") {
    var newhour = "23"
  }
  if (reportTime == "12AM") {
    var newhour = "24"
  }

  if (reportDay == "MONDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .create();
  }

  if (reportDay == "TUESDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.TUESDAY)
      .create();

  }

  if (reportDay == "WEDNESDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
      .create();

  }

  if (reportDay == "THURSDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.THURSDAY)
      .create();

  }

  if (reportDay == "FRIDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.FRIDAY)
      .create();

  }

  if (reportDay == "SATURDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.SATURDAY)
      .create();

  }

  if (reportDay == "SUNDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .create();

  }
}

//Run the Report - Phase Four
function startScheduledReportFour() {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var reportDay = settingsSheet.getRange("C13:C13").getValue();
  var reportTime = settingsSheet.getRange("E13:E13").getValue();

  if (reportTime == "1AM") {
    var newhour = "1"
  }
  if (reportTime == "2AM") {
    var newhour = "2"
  }
  if (reportTime == "3AM") {
    var newhour = "3"
  }
  if (reportTime == "4AM") {
    var newhour = "4"
  }
  if (reportTime == "5AM") {
    var newhour = "5"
  }
  if (reportTime == "6AM") {
    var newhour = "6"
  }
  if (reportTime == "7AM") {
    var newhour = "7"
  }
  if (reportTime == "8AM") {
    var newhour = "8"
  }
  if (reportTime == "9AM") {
    var newhour = "9"
  }
  if (reportTime == "10AM") {
    var newhour = "10"
  }
  if (reportTime == "11AM") {
    var newhour = "11"
  }
  if (reportTime == "12PM") {
    var newhour = "12"
  }
  if (reportTime == "1PM") {
    var newhour = "13"
  }
  if (reportTime == "2PM") {
    var newhour = "14"
  }
  if (reportTime == "3PM") {
    var newhour = "15"
  }
  if (reportTime == "4PM") {
    var newhour = "16"
  }
  if (reportTime == "5PM") {
    var newhour = "17"
  }
  if (reportTime == "6PM") {
    var newhour = "18"
  }
  if (reportTime == "7PM") {
    var newhour = "19"
  }
  if (reportTime == "8PM") {
    var newhour = "20"
  }
  if (reportTime == "9PM") {
    var newhour = "21"
  }
  if (reportTime == "10PM") {
    var newhour = "22"
  }
  if (reportTime == "11PM") {
    var newhour = "23"
  }
  if (reportTime == "12AM") {
    var newhour = "24"
  }

  if (reportDay == "MONDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .create();
  }

  if (reportDay == "TUESDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.TUESDAY)
      .create();

  }

  if (reportDay == "WEDNESDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
      .create();

  }

  if (reportDay == "THURSDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.THURSDAY)
      .create();

  }

  if (reportDay == "FRIDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.FRIDAY)
      .create();

  }

  if (reportDay == "SATURDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.SATURDAY)
      .create();

  }

  if (reportDay == "SUNDAY") {

    ScriptApp.newTrigger('runTool')
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .create();

  }
}

function startScheduledLog() {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var logDay = settingsSheet.getRange("C17:C17").getValue();
  var logTime = settingsSheet.getRange("E17:E17").getValue();

  if (logTime == "1AM") {
    var newloghour = "1"
  }
  if (logTime == "2AM") {
    var newloghour = "2"
  }
  if (logTime == "3AM") {
    var newloghour = "3"
  }
  if (logTime == "4AM") {
    var newloghour = "4"
  }
  if (logTime == "5AM") {
    var newloghour = "5"
  }
  if (logTime == "6AM") {
    var newloghour = "6"
  }
  if (logTime == "7AM") {
    var newloghour = "7"
  }
  if (logTime == "8AM") {
    var newloghour = "8"
  }
  if (logTime == "9AM") {
    var newloghour = "9"
  }
  if (logTime == "10AM") {
    var newloghour = "10"
  }
  if (logTime == "11AM") {
    var newloghour = "11"
  }
  if (logTime == "12PM") {
    var newloghour = "12"
  }
  if (logTime == "1PM") {
    var newloghour = "13"
  }
  if (logTime == "2PM") {
    var newloghour = "14"
  }
  if (logTime == "3PM") {
    var newloghour = "15"
  }
  if (logTime == "4PM") {
    var newloghour = "16"
  }
  if (logTime == "5PM") {
    var newloghour = "17"
  }
  if (logTime == "6PM") {
    var newloghour = "18"
  }
  if (logTime == "7PM") {
    var newloghour = "19"
  }
  if (logTime == "8PM") {
    var newloghour = "20"
  }
  if (logTime == "9PM") {
    var newloghour = "21"
  }
  if (logTime == "10PM") {
    var newloghour = "22"
  }
  if (logTime == "11PM") {
    var newloghour = "23"
  }
  if (logTime == "12AM") {
    var newloghour = "24"
  }

  if (logDay == "MONDAY") {

    ScriptApp.newTrigger('runLog')
      .timeBased()
      .atHour(newloghour)
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .create();

  }

  if (logDay == "TUESDAY") {

    ScriptApp.newTrigger('runLog')
      .timeBased()
      .atHour(newloghour)
      .onWeekDay(ScriptApp.WeekDay.TUESDAY)
      .create();

  }

  if (logDay == "WEDNESDAY") {

    ScriptApp.newTrigger('runLog')
      .timeBased()
      .atHour(newloghour)
      .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
      .create();

  }

  if (logDay == "THURSDAY") {

    ScriptApp.newTrigger('runLog')
      .timeBased()
      .atHour(newloghour)
      .onWeekDay(ScriptApp.WeekDay.THURSDAY)
      .create();

  }

  if (logDay == "FRIDAY") {

    ScriptApp.newTrigger('runLog')
      .timeBased()
      .atHour(newloghour)
      .onWeekDay(ScriptApp.WeekDay.FRIDAY)
      .create();

  }

  if (logDay == "SATURDAY") {

    ScriptApp.newTrigger('runLog')
      .timeBased()
      .atHour(newloghour)
      .onWeekDay(ScriptApp.WeekDay.SATURDAY)
      .create();

  }

  if (logDay == "SUNDAY") {

    ScriptApp.newTrigger('runLog')
      .timeBased()
      .atHour(newloghour)
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .create();

  }
}

// ***************
// Manually Run ##################################################
// ***************

//Run the formula to get the PageSpeed V5 data from each URL
function runTool() {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  var rows = activeSheet.getLastRow();

  for (var i = 6; i <= rows; i++) {
    var workingCell = activeSheet.getRange(i, 2).getValue();
    var stuff = "=runCheck"

    if (workingCell != "") {
      activeSheet.getRange(i, 3).setFormulaR1C1(stuff + "(R[0]C[-1])");
    }
  }
}


//Log the values and clear PageSpeed Results data:
function runLog() {
  var columnNumberToWatch = 10; // column A = 1, B = 2, etc.
  var valueToWatch = "complete";
  var sheetNameToMoveTheRowTo = "Log";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Results").activate();
  var cell = sheet.getRange("J6:J");
  var type = SpreadsheetApp.CopyPasteType.PASTE_VALUES;
  var lastRow = sheet.getLastRow();
  var Avals = ss.getRange("J6:J").getValues();
  var Alast = Avals.filter(String).length;

  if (sheet.getName() != sheetNameToMoveTheRowTo && cell.getColumn() == columnNumberToWatch && cell.getValue().toLowerCase() == valueToWatch) {
    var targetSheet = ss.getSheetByName(sheetNameToMoveTheRowTo);
    var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    sheet.getRange(cell.getRow(), 2, Alast, sheet.getLastColumn()).copyTo(targetRange, type, false);
    sheet.getRange('C6:J').clearContent();
    var d = new Date();
    var currentTime = d.toLocaleTimeString().replace(/:\d{2}\s/, ' ').slice(0, 8);
    var date = Utilities.formatDate(new Date(), "GMT-5", "MM/dd ");
  }
}

// ***************
// Parse and Convert ISO Time to EST, with AM/PM ##################################################
// ***************

function convertToAmPm(isoHour, min) {
  var estH = isoHour - 5;
  var suffix = (estH >= 12) ? "pm" : "am";
  var convertedH = ((estH > 12) ? estH - 12 : estH) || (estH == 0 ? 12 : estH);
  var newTime = convertedH + ":" + min + suffix;
  return newTime;
}

function runCheck(Url) {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var key = settingsSheet.getRange("C7:C7").getValue();
  var strategy = settingsSheet.getRange("C20").getValue();
  var serviceUrl = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=" + Url + "&key=" + key + "&strategy=" + strategy + "";
  var array = [];

  if (key == "YOUR_API_KEY") {
    return "Please enter your API key to the script";
  };
  var response = UrlFetchApp.fetch(serviceUrl);

  if (response.getResponseCode() == 200) {
    var content = JSON.parse(response.getContentText());
    if ((content != null) && (content["lighthouseResult"] != null)) {
      if (content["captchaResult"]) {
        var score = content["lighthouseResult"]["categories"]["performance"]["score"];
        var timetointeractive = content["lighthouseResult"]["audits"]["interactive"]["displayValue"].slice(0, -2);
        var firstcontentfulpaint = content["lighthouseResult"]["audits"]["first-contentful-paint"]["displayValue"].slice(0, -2);
        var firstmeaningfulpaint = content["lighthouseResult"]["audits"]["first-meaningful-paint"]["displayValue"].slice(0, -2);
        var timetofirstbyte = content["lighthouseResult"]["audits"]["time-to-first-byte"]["displayValue"].slice(19, -3);
        var speedindex = content["lighthouseResult"]["audits"]["speed-index"]["displayValue"].slice(0, -2);
        var newTime = convertToAmPm(parseInt(content["lighthouseResult"]["fetchTime"].slice(11, -11)), content["lighthouseResult"]["fetchTime"].slice(14, -8));
      } else {
        var score = "An error occured";
        var timetointeractive = "An error occured";
        var firstcontentfulpaint = "An error occured";
        var firstmeaningfulpaint = "An error occured";
        var timetofirstbyte = "An error occured";
        var speedindex = "An error occured";
        var newTime = "An error occured";
      }
    }

    var currentDate = new Date().toJSON().slice(0, 10).replace(/-/g, '/').slice(5);
    var datePlusTime = currentDate + " - " + newTime;
    Logger.log(datePlusTime)

    array.push([score, timetointeractive, firstcontentfulpaint, firstmeaningfulpaint, timetofirstbyte, speedindex, datePlusTime, "complete"]);
    Utilities.sleep(500);
    return array;
  }
}
