// Load menu

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email menu')
      .addItem('Create drafts', 'draftEmails')
      .addItem('Clear sheet', 'clearSheet')
      .addToUi();
}

// Set status

var EMAIL_DRAFTED = "EMAIL DRAFTED";

// Create drafts

function draftEmails() {
  var sheet = SpreadsheetApp.getActiveSheet(); // Use data from the active sheet
  var startRow = 2;                            // First row of data to process
  var numRows = sheet.getLastRow() - 1;        // Number of rows to process
  var lastColumn = sheet.getLastColumn();      // Last column
  var dataRange = sheet.getRange(startRow, 1, numRows, lastColumn) // Fetch the data range of the active sheet
  var data = dataRange.getValues();            // Fetch values for each row in the range

  // Work through each row in the spreadsheet
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    // Assign each row a variable
    var contactName = row[0];               // Col A: Contact name
    var companyName = row[1];               // Col B: Company name
    var contactEmail = row[2];              // Col C: Contact email
    var emailStatus = row[lastColumn - 1];  // Col D: Email status
    var aeName = 'YOUR_NAME';
    var emailSubject = 'YOUR_SUBJECT';

    // Prevent from drafing duplicates and from drafting emails without a recipient
    if (emailStatus !== EMAIL_DRAFTED && contactEmail) {

      // Set the email styling
      var style = 'style="color:#000"';

      // Build the email message
      var emailBody =  '<p '+ style +'>Hi ' + contactName + ',</p>';
          emailBody += '<p '+ style +'>How are things going with ScheduleOnce at '+ companyName +'?</p>';
          emailBody += '<p '+ style +'>If you have any issues feel free to contact me at this email address</p>';
          emailBody += '<p '+ style +'>All the best,</p>';
          emailBody += '<p '+ style +'>'+ aeName +'</p>';

      // Create the email draft
      GmailApp.createDraft(
        contactEmail,                      // Recipient
        emailSubject,                      // Subject
        '',                                // Body (plain text)
        {
        htmlBody: emailBody                // Options: Body (HTML)
        }
      );

      sheet.getRange(startRow + i, lastColumn).setValue(EMAIL_DRAFTED); // Update the last column with "EMAIL_DRAFTED"
      SpreadsheetApp.flush(); // Make sure the last cell is updated right away
    }
  }
}

// Clear the sheet

function clearSheet(){
var sheet = SpreadsheetApp.getActiveSheet();
var ui = SpreadsheetApp.getUi();

var startRow = 2;                            // First row of data to process
var numRows = sheet.getLastRow() - 1;        // Number of rows to process
var lastColumn = sheet.getLastColumn();      // Last column
var dataRange = sheet.getRange(startRow, 1, numRows, lastColumn) // Fetch the data range of the active sheet
var data = dataRange.getValues();

dataRange.clear();

}
