# Email Follow-up Script
This is a Google Apps Script that helps automate follow-up emails for recipients listed in a Google Sheet. It can be used to send follow-up emails to contacts after a certain number of days since the last email, and schedule them to be sent on a weekday.

The script includes functions to retrieve email data from a Google Sheet, update a Google Sheet row with a follow-up number and today's date, and send a follow-up email with a specific body based on the number of days elapsed since the last follow-up.

To use the script, you will need to have a Google account and a Google Sheets document with the email data. The script is triggered to run on a schedule or can be run manually from a custom menu that is created in the Google Sheets document.

## Installation
To use this script, follow these steps:

1. Open your Google Sheets document.
2. Click on the "Extensions" menu and select "Apps Script".
3. Copy and paste the script code into the editor.
4. Update the config object with your own timeZone and spreadsheetId.
5. Save the script with a name of your choice.
6. Test the script by running the scheduleFollowups() function manually.
7. (Optional) Create a custom menu to run the scheduleFollowups() function more easily.

## Configuration
The config object in the script contains two properties: timeZone and spreadsheetId. timeZone should be set to your local time zone, and spreadsheetId should be set to the ID of your Google Sheets document.

## Usage
To use this script, you can either run the scheduleFollowups() function manually or set up a trigger to run the function on a schedule. The function will send follow-up emails to the email addresses listed in the Google Sheets document, based on the number of days elapsed since the last follow-up.

## License
This script is licensed under the MIT License.
