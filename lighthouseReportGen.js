/***************************************************
* Bulk URL PageSpeed Tool (PageSpeed Insights v5)
* by james@upbuild.io
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
},
{
  name: "Share Logs",
  functionName: "triggerOnEdit"
},
{
  name: "Schedule Weekly Email",
  functionName: "weeklyEmail"
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
  //  ********** Logs *********** //
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

// ***************
// Set Weekly Trigger to Email Results ##################################################
// ***************


function weeklyEmail() {
  setWeeklyEmailTrigger();
}

function setWeeklyEmailTrigger() {
  ScriptApp.newTrigger('grabEmailFromSettingsAndSendWeeklyReport')
  .timeBased()
  .everyWeeks(1)
  .onWeekDay(ScriptApp.WeekDay.MONDAY)
  .atHour(9)
  .create();
}

function grabEmailFromSettingsAndSendWeeklyReport() {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var email = settingsSheet.getRange("C11:C11").getValue();
  var message = "<span style='margin-bottom:10px;'><p style='font-size:16px;'>Your latest Lighthouse report has finished updating!<br><br> Please visit one of the links below to either view the results in your browser, or download as a csv file:</p><br></span>";
  var liveLink = "<a style='width:125px;margin-right:20px;margin-bottom:5px;background-color:#ffffff;color:#51bceb; border:2px solid #51bceb;text-decoration:none;border-radius:6px;padding:14px 25px;text-transform:uppercase;text-align:center;font-weight:bolder;display:inline-block;'href=\'<<< BROWSER LINK >>>'>View in Browser</a>"
  var downloadCsv = "<a style='width:125px;margin-bottom:5px;background-color:#ffffff;color:#51bceb;border:2px solid #51bceb;text-decoration:none;border-radius:6px;padding:14px 25px;text-transform:uppercase;text-align:center;font-weight:bolder;display:inline-block;'href=\'<<< CSV LINK >>>'>Download</a>"
  var jmlLogo = "<<< LOGO >>>"
  var embedLogo = "<a href=<<< LOGO CID LINK >>>'></a>"
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

// ***************
// Send Email with Data (WIP) ##################################################
// ***************

function getValues1(email, subjectLine, date, currentTime) {
  var columnNumberToWatch = 10; // column A = 1, B = 2, etc.
  var valueToWatch = "complete";
  var sheetNameToMoveTheRowTo = "Log";
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = sheet.getSheetByName("Results").activate();
  var cell = sheet.getRange("J6:J");
  var type = SpreadsheetApp.CopyPasteType.PASTE_VALUES;
  var lastRow = sheet.getLastRow();
  var Avals = sheet.getRange("J6:J").getValues();
  var Alast = Avals.filter(String).length;
  var test = activeSheet.getRange(cell.getRow(), 2, Alast, activeSheet.getLastColumn()).getValues();
  var ui = SpreadsheetApp.getUi()
  var data = sheet.getRange("B5:J").getValues()
  var columnHeader = "<p><b>Header<b><p>" + data[0] + "<br />" + "<br />";
  var messageContent = cleanData(JSON.stringify(data[1]));

  var csvStr = "";
  for (var i = 0; i < data.length; i++) {
    var row = ""
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j]) {
        row = row + data[i][j];
        row = row + " ///";
        row = row.substring(0, (row.length - 1));
        csvStr += row + "\n";
      }
    }
  }
  //Creates a blob of the csv file
  var csvBlob = Utilities.newBlob(csvStr, 'text/csv', 'pageSpeedData.csv');

  ui.alert(data[1] + "Email sent to " + email);
  //  sendEmail(email, "Lighthouse ", date, currentTime, csvBlob.getDataAsString());
  Logger.log(csvStr)
  DriveApp.createFile('Log', csvBlob);
}

function cleanData(message) {
  message.split(',').join(' ')
}

//==============

//Send email on log

function emailLink(email, subject, date, time) {
  var message = "<span style='margin-bottom:10px;'><p style='font-size:16px;'>Your latest Lighthouse report has finished updating!<br><br> Please visit one of the links below to either view the results in your browser, or download as a csv file:</p><br></span>";
  var liveLink = "<a style='width:125px;margin-right:20px;margin-bottom:5px;background-color:#ffffff;color:#51bceb; border:2px solid #51bceb;text-decoration:none;border-radius:6px;padding:14px 25px;text-transform:uppercase;text-align:center;font-weight:bolder;display:inline-block;'href=\'<<< BROWSER LINK >>>'>View in Browser</a>"
  var downloadCsv = "<a style='width:125px;margin-bottom:5px;background-color:#ffffff;color:#51bceb;border:2px solid #51bceb;text-decoration:none;border-radius:6px;padding:14px 25px;text-transform:uppercase;text-align:center;font-weight:bolder;display:inline-block;'href=\'<<< CSV LINK >>>'>Download</a>"
  var jmlLogo = "<<< IMAGE LINK >>>"
  var embedLogo = "<a href='<<< IMAGE CID LINK >>>'></a>"
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

function sendCSV() {
  getValues();
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
    var subjectLine = "Your Lighthouse Report Results for "
    getUserEmail;

    if (validateEmail(userEmail) == true) {
      emailLink(userEmail,subjectLine, date, currentTime);
      Browser.msgBox("Email sent to " + userEmail);
    } else {
      var failed = ui.alert("Sorry, we weren't able to verify your email address.");
      if (failed = ui.Button.OK) {
        var tryAgain = ui.prompt("Please try again");
        var newEmail = tryAgain.getResponseText();
        if (validateEmail(newEmail) == true) {
          emailLink(newEmail, subjectLine, date, currentTime);
          Browser.msgBox("Email sent to " + newEmail);
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
  var columnNumberToWatch = 10; // column A = 1, B = 2, etc.
  var valueToWatch = "complete";
  var sheetNameToMoveTheRowTo = "Log";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Results").activate();
  var cell = sheet.getRange("J6:J");
  var type = SpreadsheetApp.CopyPasteType.PASTE_VALUES;
  var lastRow = sheet.getLastRow();
  //30

  var Avals = ss.getRange("J6:J").getValues();
  //complete, complete, complete, etc...

  var Alast = Avals.filter(String).length;
  //25

  if (sheet.getName() != sheetNameToMoveTheRowTo && cell.getColumn() == columnNumberToWatch && cell.getValue().toLowerCase() == valueToWatch) {
    var targetSheet = ss.getSheetByName(sheetNameToMoveTheRowTo);
    var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    sheet.getRange(cell.getRow(), 2, Alast, sheet.getLastColumn()).copyTo(targetRange, type, false);
    //push to weekly log
    runWeeklyLog();
    // THEN clear content
    sheet.getRange('C6:J').clearContent();
    var d = new Date();
    var currentTime = d.toLocaleTimeString().replace(/:\d{2}\s/, ' ').slice(0, 8);
    var date = Utilities.formatDate(new Date(), "GMT-5", "MM/dd ");
    //    emailLink("terence.pataneronan@gmail.com", "Your Lighthouse Report Log Results Are In for ", date, currentTime);
  }
}

//############################### Testing Weekly Log ############################### //

//Log the values and clear PageSpeed Results data:
function runWeeklyLog() {
  var columnNumberToWatch = 10; // column A = 1, B = 2, etc.
  var valueToWatch = "complete";
  var sheetNameToMoveTheRowTo = "Weekly Log";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Results").activate();
  var cell = sheet.getRange("J6:J");
  var type = SpreadsheetApp.CopyPasteType.PASTE_VALUES;
  var lastRow = sheet.getLastRow();
  //30

  var Avals = ss.getRange("J6:J").getValues();
  //complete, complete, complete, etc...

  var Alast = Avals.filter(String).length;
  //25

  if (sheet.getName() != sheetNameToMoveTheRowTo && cell.getColumn() == columnNumberToWatch && cell.getValue().toLowerCase() == valueToWatch) {
    var targetSheet = ss.getSheetByName(sheetNameToMoveTheRowTo);
    var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    sheet.getRange(cell.getRow(), 2, Alast, sheet.getLastColumn()).copyTo(targetRange, type, false);
    //push to weekly log
    var d = new Date();
    var currentTime = d.toLocaleTimeString().replace(/:\d{2}\s/, ' ').slice(0, 8);
    var date = Utilities.formatDate(new Date(), "GMT-5", "MM/dd ");
    emailLink("terence.pataneronan@gmail.com", "Your Lighthouse Report Log Results Are In for ", date, currentTime);
  }
}

//// WIP - gather all matching urls and generate averages per page
//function matchUrls() {
//  var columnNumberToWatch = 10; // column A = 1, B = 2, etc.
//  var valueToWatch = "complete";
//  var sheetNameToMoveTheRowTo = "Averages/URL";
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sheet = ss.getSheetByName("Log").activate();
//  var cell = sheet.getRange("G3:G");
//  var type = SpreadsheetApp.CopyPasteType.PASTE_VALUES;
//  var lastRow = sheet.getLastRow();
//  var Avals = ss.getRange("G3:G").getValues();
//  var Alast = Avals.filter(String).length;
//
//  if (sheet.getName() != sheetNameToMoveTheRowTo && cell.getColumn() == columnNumberToWatch && cell.getValue().toLowerCase() == valueToWatch) {
//    var targetSheet = ss.getSheetByName(sheetNameToMoveTheRowTo);
//    var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
//    sheet.getRange(cell.getRow(), 2, Alast, sheet.getLastColumn()).copyTo(targetRange, type, false);
//    var d = new Date();
//    var currentTime = d.toLocaleTimeString().replace(/:\d{2}\s/, ' ').slice(0, 8);
//    var date = Utilities.formatDate(new Date(), "GMT-5", "MM/dd ");
//  }
//}

//Alt Version

function collectUrls() {
  var arr = [];
  var sheetNameToMoveTheRowTo = "Averages/URL";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Results").activate();
  var cell = sheet.getRange("B6:B");
  var type = SpreadsheetApp.CopyPasteType.PASTE_VALUES;
  var totalRows = sheet.getLastRow();
  var allValues = ss.getRange("B6:B").getValues();
  var newSheet = ss.getSheetByName(sheetNameToMoveTheRowTo).activate();
  var lastActiveRow = allValues.filter(String).length;
  var newSheetLastRow = (ss.getSheetByName(sheetNameToMoveTheRowTo).getLastRow() - 1);
  var allValuesNewSheet = newSheet.getRange("A2:A").getValues().filter(String);
  var allValuesNewSheetStr = JSON.stringify(newSheet.getRange("A2:A").getValues().filter(String));
  //  Logger.log(allValues.filter(String).length);
  //  Logger.log(ss.getSheetByName(sheetNameToMoveTheRowTo).getLastRow() - 1);
  //  Logger.log(lastActiveRow);
  //  Add links if they do not already exist in sheet
  if (sheet.getName() != sheetNameToMoveTheRowTo && cell.getValue() != "" && lastActiveRow != (ss.getSheetByName(sheetNameToMoveTheRowTo).getLastRow() - 1)) {
    var targetSheet = ss.getSheetByName(sheetNameToMoveTheRowTo);
    var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    sheet.getRange(cell.getRow(), 2, lastActiveRow, 3).copyTo(targetRange, type, false);
  }

  for (var i = 0; i < allValues.filter(String).length; i++) {
    arr.push(allValues.filter(String)[i]);
  }

  var obj = {};
  for (var i = 0; i < arr.length; i++) {
    obj["url: "] = arr[i];
  }
  return obj;
}

function matchEntries(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Log");
  var resultSheet = ss.getSheetByName("Results");
  var column = logSheet.getRange("A3:A");
  var logValues = column.getValues().filter(String);
  var resultColumn = resultSheet.getRange("B6:B");
  var resultValues = resultColumn.getValues().filter(String);
  Logger.log(logValues.length);
  //  var row = logValues.length - 1;
  var row = 0;
  Logger.log(resultValues[0][0] == logValues[2]);
  while (logValues[row + 1][0] != resultValues[row + 1]) {
    row++;
    Logger.log(row);
  }

  //  if (logValues[row][0] == resultValues[0]) {
  //    return row + 1;
  ////      Logger.log(row)
  //  } else {
  //    return - 1;
  //      Logger.log(row)
  //  }
}


//  if (cell.getValue() != "") {
//    for (var i = 0; i < Avals.length; i++) {
//      var link = [];
//      if (Avals[i].filter(String) != "") {
//        var values = Avals.filter(String); //object
//         link += JSON.stringify(values[i]);
//        Logger.log(link); //returns single result
//
//        for (var j = 0; j < link.length; j++) {
//          var item = "";
//          item += link[j];
//          return item;
//          Logger.log(item);
//        }
//      }
//    }
//  }
//  Logger.log(arr);

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

function testApiResponse() {
  var url = "ENTER PAGESPEED URL Example Here";
  var arr = [];
  var response = UrlFetchApp.fetch(url);
  var content = JSON.parse(response.getContentText());
  var newTime = convertToAmPm(parseInt(content["lighthouseResult"]["fetchTime"].slice(11, -11)), content["lighthouseResult"]["fetchTime"].slice(14, -8));
  Logger.log(newTime);
};

//weekly log page browser link
//https://docs.google.com/spreadsheets/d/e/2PACX-1vRzQ5m7sV55o5AN6jHM85qBso1JnjBcEWO7DImrvhxDT8vEzElKih2AhGLp80U7dF3-KQYq9EBh3hHF/pubhtml?gid=1520138531&single=true

//weekly log csv link
//https://docs.google.com/spreadsheets/d/e/2PACX-1vRzQ5m7sV55o5AN6jHM85qBso1JnjBcEWO7DImrvhxDT8vEzElKih2AhGLp80U7dF3-KQYq9EBh3hHF/pub?gid=1520138531&single=true&output=csv
