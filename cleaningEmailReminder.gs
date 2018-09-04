var SUFFIX = "@hkn.eecs.berkeley.edu";
var FIRST_SENT = "First";
var LAST_SENT = "Last";

var FIRST_MESSAGE = "Hello pPl,\n\n" + 
"One of the requirements of any officer or asst. officer is to participate in the weekly cleanup of the Cory/Soda offices, and you are on-duty this weekend to cleanup the offices. " + 
"If you don't know what to do, look at <a href=\"https://hkn.eecs.berkeley.edu/prot/Offices\">prot</a>.\n\n" + 
"Sincerely,\n" +
"Tutoring Committee and RSec";

var LAST_MESSAGE = "Hello pPl,\n\n" + 
"A friendly reminder that you are in charge of this week's cleanup of the Cory/Soda offices. " + 
"Don't forget to check <a href=\"https://hkn.eecs.berkeley.edu/prot/Offices\">prot</a>.\n\n" + 
"Sincerely,\n" +
"Tutoring Committee and RSec";

function sendEmails() {
  
  var emailsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EmailHandles")
  var emails = emailsheet.getRange(2, 1, emailsheet.getLastRow(), 2).getValues();
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intercommittee");
  var column_names = sheet.getRange(1, 1, 1, 3).getValues()[0];  // Get 1st 4 Column Names (Weekend of, Sent, # of People, People)
  var notify = column_names.indexOf("Sent"); // Modify the line above if more columns are added later
  if (notify == -1) {
      Logger.log("Incorrect Formatting: Missing 'Sent' Column");
      return;
  }
  
  var data = sheet.getRange(2, 1, sheet.getLastRow(), 4).getValues();
  for (var i = 0; i < data.length; i++) {
    var[time, sent, people] = data[i];
    
    var days_left = (time - Date.now()) / (1000 * 60 * 60 * 24);
    if (sent != LAST_SENT) {
      if (sent == FIRST_SENT) {
        if (days_left > 1.9) {  // Send 2 days early
          return;
        }
        var snt = LAST_SENT;
        var message = LAST_MESSAGE;
        var sbjct = "Office Cleanup Reminder";
      } else {
        if (days_left > 4.9) {  // Send 5 days early
          return; 
        }
        var snt = FIRST_SENT;
        var message = FIRST_MESSAGE;
        var sbjct = "Office Cleanup";
      }
      
      var to_cmt = "";  // Create List of Emails for Joint-Cleanups
      
      var split_people = people.split(',');
      for (var x = 0; x < split_people.length; x++) {
        for (var y = 0; y < emails.length; y++) {
          if (split_people[x] == emails[y][0]) {
            to_cmt = to_cmt + emails[y][1] + SUFFIX + ",";
          }
        }
      }
      
      MailApp.sendEmail({
        to: to_cmt,
        cc: "tutoring" + SUFFIX + "," + "rsec" + SUFFIX,
        subject: sbjct,
        htmlBody: message.replace(/pPl/g, people).replace(/\n/g, "<br>"), //.replace(/CMTE/g, committee) to replace
       });
      
      sheet.getRange(2 + i, notify + 1).setValue(snt);  // Why notify + 1 and 2 + i? Changes to 1-based index
      SpreadsheetApp.flush();
      return;
    }
  }
}
