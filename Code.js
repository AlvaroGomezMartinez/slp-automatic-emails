/**
 * Automatic Emails
 * 
 * @function sendAutomaticEmails
 * Sends a personalized thank-you email to each respondent from the Google Form 
 * linked to this spreadsheet, if a confirmation timestamp has not already been recorded.
 *
 * The email is sent to the address in column C, using the first name from column B.
 * After sending, a timestamp is recorded in column L in the format M/d/yy hh:mm a.
 * 
 * This script is intended to run automatically via a time-driven trigger 
 * scheduled for weekdays at 5:00 PM.
 * 
 * You can customize the email content using the `subject` and `message` variables.
 * 
 * @author Alvaro Gomez
 * Academic Technology Coach, Dept. of Academic Technology  
 * Office: 210-397-9408  
 * Cell: 210-363-1577
 */

function sendAutomaticEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const data = sheet.getDataRange().getValues();
  const timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  
  for (let i = 1; i < data.length; i++) {
    const name = data[i][1];  // Column B
    const email = data[i][2]; // Column C
    const sent = data[i][11]; // Column L

    if (!sent) {
      const firstName = name.split(" ")[0];
      const subject = "Thank you for your response";
      const message = `Dear ${firstName},\n\nThank you for your response.`;

      MailApp.sendEmail(email, subject, message);

      const timestamp = Utilities.formatDate(new Date(), timeZone, "M/d/yy hh:mm a");
      sheet.getRange(i + 1, 12).setValue(timestamp); // Column L
    }
  }
}
