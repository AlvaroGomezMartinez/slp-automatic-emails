/**
 * Automatic Emails for SLP's Feedback Form
 * 
 * Sends a personalized thank-you email when a new Google Form response is submitted.
 * The email goes to the address in column C, using the first name from column B.
 * A timestamp is recorded in column L in the format M/d/yy hh:mm a.
 *
 * This version is optimized for use with an "on form submit" trigger.
 * 
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e - The event object from the form submission.
 */
function sendAutomaticEmails(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const values = e.values;

  const name = values[1];   // Column B
  const email = values[2];  // Column C
  const sent = values[11];  // Column L (might be empty on new submission)

  if (!sent) {
    const firstName = name.split(" ")[0];
    const subject = "Thank you for your response";
    const message = `Dear ${firstName},\n\nOn behalf of the NISD SLP Leadership Team, we appreciate you taking the time to share your thoughts and give us feedback. Your insights are valuable!\n\nThank you!`;

    MailApp.sendEmail(email, subject, message);

    const timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    const timestamp = Utilities.formatDate(new Date(), timeZone, "M/d/yy hh:mm a");
    
    sheet.getRange(row, 12).setValue(timestamp); // Column L
  }
}

/**
 * The function below is for testing.
 */

// function sendAutomaticEmailsTest() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
//   const data = sheet.getDataRange().getValues();
//   const timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  
//   for (let i = 1; i < data.length; i++) {
//     const name = data[i][1];  // Column B
//     const email = data[i][2]; // Column C
//     const sent = data[i][11]; // Column L

//     if (!sent) {
//       const firstName = name.split(" ")[0];
//       const subject = "Thank you for your response";
//       const message = `Dear ${firstName},\n\nOn behalf of the NISD SLP Leadership Team, we appreciate you taking the time to share your thoughts and give us feedback. Your insights are valuable!\n\nThank you!`;

//       MailApp.sendEmail(email, subject, message);

//       const timestamp = Utilities.formatDate(new Date(), timeZone, "M/d/yy hh:mm a");
//       sheet.getRange(i + 1, 12).setValue(timestamp); // Column L
//     }
//   }
// }