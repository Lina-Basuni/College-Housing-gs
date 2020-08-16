
// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = 'EMAIL_SENT';

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendBoardingEmails() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Boarding");
  var startRow = 2; // First row of data to process
  var numRows = 2; // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 20);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[11]; // First column
    Logger.log(emailAddress);
    var name = row[8]; // Second column
    Logger.log(name);

    var message ="<p>Dear "+name+",</p>"+"<p>Thank you so much for your interest in renting your College residence with us! We are excited to have you join our community and can't wait for your move-in this month. </p> <p>For your reference, we've attached a blank copy of the lease agreement so you can review it before making any final decisions or deposit payments.</p>"+ "<p>If you have any questions regarding your application or the lease draft, please let us know at contact@collegehousing.us</p>"+"<p>Best,</p>";
    var file = DriveApp.getFileById('1_jkG9Px8VIsg1FSF8NTJJhh_gVULy2tO');
    var emailSent = row[0]; // Third column
    if (emailSent !== EMAIL_SENT) { // Prevents sending duplicates
      var subject = 'Sending emails from a Spreadsheet';
      MailApp.sendEmail(emailAddress, subject , '', {
        htmlBody: message,
        name: 'College Housing',
        attachments: [file.getAs(MimeType.PDF)]
      });
      sheet.getRange(startRow + i, 1).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}
