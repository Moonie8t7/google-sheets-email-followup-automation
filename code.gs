/**
 * @fileOverview This script retrieves email data from a Google Sheet and sends follow-up emails to recipients
 * based on the number of days elapsed since the last follow-up, scheduling them for weekdays. It also allows
 * manual scheduling of follow-ups via a custom menu in a Google Sheets document.
 *
 * @author u/IAmMoonie <https://www.reddit.com/user/IAmMoonie/>
 * @license MIT
 * @version 1.0
 */

/**
 * Object that stores configuration data for the script. It contains two properties: timeZone and spreadsheetId.
 * These properties are used throughout the script to specify the time zone and spreadsheet to be used.
 * @typedef {Object} Config
 * @property {string} timeZone - The time zone to be used in the script.
 * @property {string} spreadsheetId - The ID of the Google Sheets spreadsheet to be used in the script.
 */
const config = {
  timeZone: "America/New_York",
  // Replace with your timeZone
  spreadsheetId: "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" // Replace with your Spreadsheet ID
};
/**
 * This function retrieves email data from a Google Sheet.
 * @function
 * @returns {Array<Array<*>>} - An array of values from a Google Sheets spreadsheet. Specifically, it is returning
 * the values in columns A through G, starting from row 2 and ending at the last row with data in the sheet. These
 * values likely represent email addresses and associated data.
 */
function getEmailsFromSheet() {
  const sheet = SpreadsheetApp.openById(config.spreadsheetId).getActiveSheet();
  const lastRow = sheet.getLastRow();
  return sheet.getRange(2, 1, lastRow - 1, 7).getValues();
}
/**
 * This function updates a Google Sheets row with a follow-up number and today's date.
 * @function
 * @param {number} row - The row number of the cell to be updated in the spreadsheet.
 * @param {number} followupNum - A number that represents the column number in the spreadsheet where the
 * follow-up date should be updated.
 * @param {string} today - The current date in a specific format (not specified in the code snippet).
 */
function updateSheet(row, followupNum, today) {
  const sheet = SpreadsheetApp.openById(config.spreadsheetId).getActiveSheet();
  const timestamp = new Date().toLocaleString("en-US", {
    timeZone: config.timeZone
  });
  sheet.getRange(row + 2, followupNum + 3).setValue(timestamp);
  sheet.getRange(row + 2, 7).setValue(today);
}
/**
 * This function sends a follow-up email with a specific body based on the number of days elapsed since
 * the last follow-up, and schedules it to be sent on a weekday.
 * @function
 * @param {string} email - The email address to which the follow-up email will be sent.
 * @param {number} row - The row number in the sheet where the email data is stored.
 * @returns {void|undefined} If `followupNum` is 0, the function will return nothing (`undefined`).
 */
function sendFollowupEmail(email, row) {
  try {
    const { timeZone } = config;
    const emailData = getEmailsFromSheet()[row];
    const lastFollowup = emailData[5];
    const today = new Date().toLocaleDateString("en-US", {
      timeZone
    });
    const daysElapsed = daysBetween(lastFollowup, today);
    const followupNum =
      daysElapsed >= 2 && daysElapsed < 4
        ? 1
        : daysElapsed >= 4 && daysElapsed < 7
        ? 2
        : daysElapsed >= 7
        ? 3
        : 0;
    if (!followupNum) return;
    const thread = GmailApp.search(`to:${email}`)[0];
    const message = thread.getMessages()[0];
    const replies = message.getThread().getMessages();
    const userEmailAddress = Session.getActiveUser().getEmail();
    if (
      replies.some(
        (reply) =>
          reply.getFrom().toLowerCase() !== userEmailAddress.toLowerCase()
      )
    )
      return;
    const body = `Follow-up email #${followupNum} body goes here.`;
    const replyMessage = message.reply("", {
      htmlBody: body
    });
    let sendDate = getRandomTime(new Date(), timeZone);
    while (!isWeekday(sendDate)) {
      sendDate.setDate(sendDate.getDate() + 1);
      sendDate = getRandomTime(sendDate, timeZone);
    }
    replyMessage.scheduleSend(sendDate);
    updateSheet(row, followupNum, today);
  } catch (error) {
    console.error(`Error sending follow-up email to ${email}:`, error);
  }
}
/**
 * This function schedules follow-up emails for weekdays to the recipients listed in a Google Sheet.
 * @function
 * @returns {void|undefined} If the current date is not a weekday, nothing is returned (the function stops
 * executing). If the current date is a weekday, the function sends follow-up emails to the email addresses listed
 * in the sheet. No value is explicitly returned from the function.
 */
function scheduleFollowups() {
  const currentDate = new Date();
  if (!isWeekday(currentDate)) {
    return;
  }
  const emails = getEmailsFromSheet();
  emails.forEach((emailData, rowIndex) => {
    const email = emailData[0];
    sendFollowupEmail(email, rowIndex);
  });
}
/**
 * This function calculates the number of days between two given dates.
 * @function
 * @param {string|Date} date1 - The first date in the format of a string or a Date object.
 * @param {string|Date} date2 - The second date parameter that is being passed to the function. It is the date that
 * you want to calculate the number of days between, in relation to the first date parameter.
 * @returns {number} The number of days between two dates.
 */
function daysBetween(date1, date2) {
  const oneDay = 24 * 60 * 60 * 1000;
  const firstDate = new Date(date1);
  const secondDate = new Date(date2);
  return Math.round(Math.abs((firstDate - secondDate) / oneDay));
}
/**
 * The function checks if a given date is a weekday (Monday to Friday).
 * @function
 * @param {Date} date - The input parameter "date" is a JavaScript Date object representing a specific date
 * and time.
 * @returns {boolean} A boolean value indicating whether the given date is a weekday (Monday to Friday) or not.
 */
function isWeekday(date) {
  const day = date.getDay();
  return day !== 0 && day !== 6;
}
/**
 * The function generates a random time within a specified time zone.
 * @function
 * @param {Date} date - The date parameter is a JavaScript Date object that represents the date for which a
 * random time needs to be generated.
 * @param {string} timeZone - The timeZone parameter is a string representing the time zone to be used for the
 * date object. It should be in the format of the IANA Time Zone database, such as "America/New_York"
 * or "Europe/London".
 * @returns {Date} A Date object with a random time between 9:00 AM and 5:59 PM in the specified time zone.
 */
function getRandomTime(date, timeZone) {
  const localDate = new Date(
    date.toLocaleString("en-US", {
      timeZone
    })
  );
  const hours = Math.floor(Math.random() * (17 - 9) + 9);
  const minutes = Math.floor(Math.random() * 60);
  localDate.setHours(hours, minutes);
  return localDate;
}
/**
 * This function creates a custom menu in a Google Sheets document with an option to schedule
 * follow-ups - this is so it can be done manually if needed.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Follow-Up")
    .addItem("Schedule Follow-ups", "scheduleFollowups")
    .addToUi();
}
