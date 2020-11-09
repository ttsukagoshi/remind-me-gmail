// MIT License
// 
// Copyright (c) 2020 Taro TSUKAGOSHI
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
// 
// For the latest information, see https://github.com/ttsukagoshi/remind-me-gmail

const REMINDER_SHEET_NAME = 'Reminder';
const PLACEHOLDER_MARKER = /\{\{[^\}]+\}\}/g;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Send Reminders')
    .addItem('月初', 'Reminder_MonthStart')
    .addItem('毎週月曜13時', 'Reminder_EveryMonday1300')
    .addToUi();
}

function Reminder_MonthStart() {
  const reminderName = '月初';
  var placeholderValues = {
    'date': Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyyMM"),
    'spreadsheetUrl': SpreadsheetApp.getActiveSpreadsheet().getUrl()
  };
  sendReminder(reminderName, placeholderValues);
}

function Reminder_EveryMonday1300() {
  const reminderName = '毎週月曜13:00';
  var placeholderValues = {
    'date': Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyyMM"),
    'spreadsheetUrl': SpreadsheetApp.getActiveSpreadsheet().getUrl()
  };
  sendReminder(reminderName, placeholderValues);
}

/**
 * Send the reminder, replacing placeholder values.
 * @param {string} reminderName Name of the reminder, which should correspond with the value in the spreadsheet.
 * @param {Object} placeholderValues [Optional] Placeholder values to replace in the email text. {'valueName': value}
 */
function sendReminder(reminderName, placeholderValues = {}) {
  const reminderContent = getReminderContent_().find(element => element[0] == reminderName);
  const myAddress = Session.getActiveUser().getEmail();
  // Replace the placeholder values and compose the email
  var [subjectReplaced, bodyReplaced] = [reminderContent[1], reminderContent[2]].map(content => {
    let textVars = content.match(PLACEHOLDER_MARKER);
    if (!textVars) {
      return content; // return text itself if no placeholder marker is found
    } else {
      // Get the text inside markers, e.g., {{field name}} => field name
      let markerTexts = textVars.map(value => value.substring(2, value.length - 2)); // assuming that the text length for opening and closing markers are 2 and 2, respectively
      // Replace variables in textVars with the actual values from the data object.
      // If no value is available, replace with 'NA'.
      textVars.forEach(
        (variable, i) => content = content.replace(variable, placeholderValues[markerTexts[i]] || 'NA')
      );
      return content;
    }
  });
  MailApp.sendEmail(myAddress, subjectReplaced, bodyReplaced);
}

/**
 * Gets the reminder email contents in a 2d JavaScript array and returns it.
 * @returns {array} Reminder content in 2d array with reminder name, email subject, and email body text as its respective values per row, 
 */
function getReminderContent_() {
  var reminders = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(REMINDER_SHEET_NAME)
    .getDataRange()
    .getValues();
  reminders.shift();
  return reminders;
}
