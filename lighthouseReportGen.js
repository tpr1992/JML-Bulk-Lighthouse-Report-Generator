/***************************************************
 * Automated Lighthouse Report
 * template by james@upbuild.io
 * for JML by tronan@jmclaughlin.com
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
  // Add JML Menu to run custom schedule
  var dailySchedule = [{
      name: "Run Report",
      functionName: "runTool"
    },
    {
      name: "Run Log",
      functionName: "runLog"
    },
    {
      name: "Set Daily Schedule",
      functionName: "checkIfTriggersAlreadyExistThenSet"
    },
    {
      name: "Set Weekly Email",
      functionName: "weeklyEmail"
    },
    {
      name: "Clear Schedule",
      functionName: "resetSchedule"
    },
    {
      name: "Share Logs",
      functionName: "triggerOnEdit"
    }
  ];
  sheet.addMenu("JML Menu", dailySchedule)
}

function resetConfirm() {
  Browser.msgBox("Success! All triggers removed.");
}

function scheduleConfirm() {
  Browser.msgBox("Reports Scheduled!")
}

function scheduleDuplicate() {
  Browser.msgBox("Your schedule is already set. If you would like to clear it and start over, please first select 'Clear Schedule' and try again.");
}

function resetFailure() {
  Browser.msgBox("Your schedule is already empty.")
}

function emailConfirm(email) {
  Browser.msgBox("Weekly email set! Your report will be sent to " + email + " on a weekly basis.");
}

// Deletes all triggers in the current project.
function resetSchedule() {
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.length > 0) {
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() != "checkEmail") {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    resetConfirm();
  } else {
    resetFailure();
  }
  checkTriggerStatusEmail();
  checkTriggerStatusReports();
}

function checkTriggerStatusEmail() {
  var triggers = ScriptApp.getProjectTriggers();
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var email = settingsSheet.getRange("C11:C11").getValue();
  var arr = [];
  for (var i = 0; i < triggers.length; i++) {
      arr.push(triggers[i].getHandlerFunction());
  }
  if (arr.indexOf("grabEmailFromSettingsAndSendWeeklyReport") > -1) {
    Logger.log(arr)
    settingsSheet.getRange("C18:C18").setValue("Yes");
  } else {
      settingsSheet.getRange("C18:C18").setValue("No");
  }
}

function checkTriggerStatusReports() {
  var triggers = ScriptApp.getProjectTriggers();
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var email = settingsSheet.getRange("C11:C11").getValue();
  var arr = [];
  for (var i = 0; i < triggers.length; i++) {
      arr.push(triggers[i].getHandlerFunction());
  }
  if (arr.indexOf("runTool") > -1 && arr.indexOf("runLog") > -1) {
    Logger.log(arr)
    settingsSheet.getRange("C17:C17").setValue("Yes");
  } else {
      settingsSheet.getRange("C17:C17").setValue("No");
  }
}

function checkIfTriggersAlreadyExistThenSet() {
  var triggers = ScriptApp.getProjectTriggers();
  var arr = [];
  for (var i = 0; i < triggers.length; i++) {
    arr.push(triggers[i].getHandlerFunction());
  }
  var filteredLogs = arr.filter(function(item) {
    if (item == "runLog") {
      return item;
    }
  });
  var filteredRun = arr.filter(function(item) {
    if (item == "runTool") {
      return item;
    }
  })
  if (filteredRun.length >= 6 && filteredLogs.length >= 3) {
   scheduleDuplicate();
  } else {
      dailyTrigger();
  }
}

// ***************
// Check Email Input - Update on Edit ##################################################
// ***************

function checkEmail() {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var email = settingsSheet.getRange("C11:C11").getValue();
  if (email.length < 1) {
    settingsSheet.getRange("C12:C12").setValue("Please enter your email above");
  } else if (validateEmail(email)) {
    settingsSheet.getRange("C12:C12").setValue("Email confirmed!");
  } else {
    settingsSheet.getRange("C12:C12").setValue("Please enter a valid email address");
  }
}

function weeklyEmail() {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var email = settingsSheet.getRange("C11:C11").getValue();
  var triggers = ScriptApp.getProjectTriggers();
  var arr = [];
  for (var i = 0; i < triggers.length; i++) {
    arr.push(triggers[i].getHandlerFunction());
  }
  var emailArr = arr.filter(function(item) {
    if (item == "grabEmailFromSettingsAndSendWeeklyReport") {
      return item;
    }
  })
  if (emailArr.length != 0) {
    Browser.msgBox("Your email is already set to go out, if you would like to cancel please select 'Clear Schedule' above and try again.")
  } else {
    setWeeklyEmailTrigger();
    emailConfirm(email);
  }
  checkTriggerStatusEmail();
}

// ***************
// Set Triggers
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
  checkTriggerStatusReports();
  checkTriggerStatusEmail();
}

// ***************
// Set Weekly Trigger to Email Logs ##################################################
// ***************

function setWeeklyEmailTrigger() {
  ScriptApp.newTrigger('grabEmailFromSettingsAndSendWeeklyReport')
  .timeBased()
  .everyWeeks(1)
  .onWeekDay(ScriptApp.WeekDay.MONDAY)
  .atHour(9)
  .create();
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

function grabEmailFromSettingsAndSendWeeklyReport() {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var email = settingsSheet.getRange("C11:C11").getValue();
  var message = "<span style='margin-bottom:10px;'><p style='font-size:16px;'>Your latest Lighthouse report has finished updating!<br><br> Please visit one of the links below to either view the results in your browser, or download as a csv file:</p><br></span>";
  var liveLink = "<a style='width:125px;margin-right:20px;margin-bottom:5px;background-color:#ffffff;color:#51bceb; border:2px solid #51bceb;text-decoration:none;border-radius:6px;padding:14px 25px;text-transform:uppercase;text-align:center;font-weight:bolder;display:inline-block;'href=\'<<< LINK TO SHEET >>>'>View in Browser</a>"
  var downloadCsv = "<a style='width:125px;margin-bottom:5px;background-color:#ffffff;color:#51bceb;border:2px solid #51bceb;text-decoration:none;border-radius:6px;padding:14px 25px;text-transform:uppercase;text-align:center;font-weight:bolder;display:inline-block;'href=<<< LINK TO CSV >>>'>Download</a>"
  var jmlLogo = "<<< LOGO LINK >>>"
  var embedLogo = "<a href='<<< HREF FOR PHOTO >>>'><img height='20%' width='20%' src='cid:jmlLogo'></a>"
  var d = new Date();
  var currentTime = d.toLocaleTimeString().replace(/:\d{2}\s/, ' ').slice(0, 8);
  var date = Utilities.formatDate(new Date(), "GMT-5", "MM/dd ");
  var subjectLine = "Your Lighthouse Report Results for "
  var jmlLogoBlob = UrlFetchApp
  .fetch(jmlLogo)
  .getBlob()
  .setName("jmlLogoBlob");
  MailApp.sendEmail({
    to: email,
    subject: subjectLine + date + "-" + currentTime,
    htmlBody: (message + liveLink + downloadCsv + "<br><br>" + "<hr style='margin-bottom:15px;'>" + embedLogo),
    inlineImages: {jmlLogo: jmlLogoBlob}
  })
}

//==============

//Send email on log

function emailLink(email, subject, date, time) {
  var message = "<span style='margin-bottom:10px;'><p style='font-size:16px;'>Your latest Lighthouse report has finished updating!<br><br> Please visit one of the links below to either view the results in your browser, or download as a csv file:</p><br></span>";
  var liveLink = "<a style='width:125px;margin-right:20px;margin-bottom:5px;background-color:#ffffff;color:#51bceb; border:2px solid #51bceb;text-decoration:none;border-radius:6px;padding:14px 25px;text-transform:uppercase;text-align:center;font-weight:bolder;display:inline-block;'href=\'<<< LINK TO SHEET >>>'>View in Browser</a>"
  var downloadCsv = "<a style='width:125px;margin-bottom:5px;background-color:#ffffff;color:#51bceb;border:2px solid #51bceb;text-decoration:none;border-radius:6px;padding:14px 25px;text-transform:uppercase;text-align:center;font-weight:bolder;display:inline-block;'href=\'<<< LINK TO CSV >>>'>Download</a>"
  var jmlLogo = "<<< LOGO LINK >>>"
  var embedLogo = "<a href='<<< HREF FOR PHOTO >>>'><img height='20%' width='20%' src='cid:jmlLogo'></a>"
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
  var columnNumberToWatch = 10;
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
  return newTime.length < 7 ? 0 + newTime : newTime;
}

function runCheck(Url) {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var key = settingsSheet.getRange("C7:C7").getValue();
  var strategy = settingsSheet.getRange("C15").getValue();
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
        var timetointeractive = parseFloat(content["lighthouseResult"]["audits"]["interactive"]["displayValue"].slice(0, -2));
        var firstcontentfulpaint = parseFloat(content["lighthouseResult"]["audits"]["first-contentful-paint"]["displayValue"].slice(0, -2));
        var firstmeaningfulpaint = parseFloat(content["lighthouseResult"]["audits"]["first-meaningful-paint"]["displayValue"].slice(0, -2));
        var timetofirstbyte = parseFloat(content["lighthouseResult"]["audits"]["time-to-first-byte"]["displayValue"].slice(19, -3));
        var speedindex = parseFloat(content["lighthouseResult"]["audits"]["speed-index"]["displayValue"].slice(0, -2));
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
