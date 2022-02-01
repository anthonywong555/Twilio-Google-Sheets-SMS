/**
 * Twilio Information
 */
const TWILIO_ACCOUNT_SID = 'TWILIO_ACCOUNT_SID';
const TWILIO_AUTH_TOKEN = 'TWILIO_AUTH_TOKEN';
const TWILIO_PHONE_NUMBER = 'TWILIO_PHONE_NUMBER';

/**
 * Contact Information
 * Note: Zero Based Indexing. A = 0, B = 1, ....
 * Will store as a A, B, C
 */
const CONTACT_NAME_COLUMN = 'CONTACT_NAME_COLUMN';
const CONTACT_PHONE_NUMBER_COLUMN = 'CONTACT_PHONE_NUMBER_COLUMN';
const MESSAGE_COLUMN = 'MESSAGE_COLUMN';
const RESULT_COLUMN = 'RESULT_COLUMN';

const ui = SpreadsheetApp.getUi();
const userProperties = PropertiesService.getUserProperties();

/**
 * Google Sheets Lifecycle Methods
 */

function onOpen() {
  ui.createMenu('Twilio - Credentials')
    .addItem('Set Twilio Account SID', 'setTwilioAccountSID')
    .addItem('Set Twilio Auth Token', 'setTwilioAuthToken')
    .addItem('Set Twilio phone number', 'setTwilioPhoneNumber')
    .addItem('Set Contact Phone Number Column', 'setContactPhoneNumberColumn')
    .addItem('Set Message Column', 'setMessageColumn')
    .addItem('Set Result Column', 'setResultColumn')
    .addItem('Delete Twilio Account SID', 'deleteTwilioAccountSID')
    .addItem('Delete Twilio Auth Token', 'deleteTwilioAuthToken')
    .addItem('Delete Twilio phone number', 'deleteTwilioPhoneNumber')
    .addToUi();
  ui.createMenu('Twilio - Actions')
    .addItem('Send SMS to All', 'sendSmsToAll')
    .addToUi();
};

/**
 * UI Getters and Setters
 */

function setTwilioAccountSID() {
  const scriptValue = ui.prompt('Enter your Twilio Account SID', ui.ButtonSet.OK);
  userProperties.setProperty(TWILIO_ACCOUNT_SID, scriptValue.getResponseText());
};

function setTwilioAuthToken() {
  const scriptValue = ui.prompt('Enter your Twilio Auth Token', ui.ButtonSet.OK);
  userProperties.setProperty(TWILIO_AUTH_TOKEN, scriptValue.getResponseText());
};

function setTwilioPhoneNumber() {
  const scriptValue = ui.prompt('Enter your Twilio phone number in this format: +12345678900', ui.ButtonSet.OK);
  userProperties.setProperty(TWILIO_PHONE_NUMBER, scriptValue.getResponseText());
};

function deleteTwilioAccountSID() {
  userProperties.deleteProperty(TWILIO_ACCOUNT_SID);
};

function deleteTwilioAuthToken() {
  userProperties.deleteProperty(TWILIO_AUTH_TOKEN);
};

function deleteTwilioPhoneNumber() {
  userProperties.deleteProperty(TWILIO_PHONE_NUMBER);
};

function setResultColumn() {
  const scriptValue = ui.prompt('Enter your Result Column', ui.ButtonSet.OK);
  userProperties.setProperty(RESULT_COLUMN, scriptValue.getResponseText());
}

function setContactPhoneNumberColumn () {
  const scriptValue = ui.prompt('Enter your Contact Phone Number Column', ui.ButtonSet.OK);
  userProperties.setProperty(CONTACT_PHONE_NUMBER_COLUMN, scriptValue.getResponseText());
}

function setMessageColumn () {
  const scriptValue = ui.prompt('Enter your Message Column', ui.ButtonSet.OK);
  userProperties.setProperty(MESSAGE_COLUMN, scriptValue.getResponseText());
}

/**
 * Twilio / Helper Methods
 */

function sendSms(payload) {
  const twilioAccountSID = userProperties.getProperty('TWILIO_ACCOUNT_SID');
  const twilioAuthToken = userProperties.getProperty('TWILIO_AUTH_TOKEN');
  const twilioUrl = 'https://api.twilio.com/2010-04-01/Accounts/' + twilioAccountSID + '/Messages.json';
  const authenticationString = twilioAccountSID + ':' + twilioAuthToken;

  try {
    const response = UrlFetchApp.fetch(twilioUrl, {
      method: 'post',
      headers: {
        Authorization: 'Basic ' + Utilities.base64Encode(authenticationString)
      },
      payload
    });

    console.log(response);
    return response.getContentText();
  } catch (err) {
    return 'error: ' + err;
  }
};

function sendSmsToAll() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  const resultColumnIndex = letterToColumn(userProperties.getProperty(RESULT_COLUMN));
  const messageColumnIndex = letterToColumn(userProperties.getProperty(MESSAGE_COLUMN));
  const contactPhoneNumberColumnIndex = letterToColumn(userProperties.getProperty(CONTACT_PHONE_NUMBER_COLUMN));
  
  const twilioPhoneNumber = userProperties.getProperty(TWILIO_PHONE_NUMBER);

  rows.forEach((aRow) => {
    const contactPhoneNumber = aRow[contactPhoneNumberColumnIndex];
    const message = aRow[messageColumnIndex];
    // Check to see is Contact Phone Number and Message
    if(contactPhoneNumber && message) {
      const payload = {
        To: contactPhoneNumber,
        From: twilioPhoneNumber,
        Body: message
      };

      const result = sendSms(payload);

      console.log(`result: ${result}`);

      // Check to see if the result row is declare.
      if(resultColumnIndex) {
        aRow[resultColumnIndex] = result;
      }
    }
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

/**
 * Source: https://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
 */
 function letterToColumn(letter) {
  let column = 0;
  if(letter) {
    let length = letter.length;
    for (let i = 0; i < length; i++)
    {
      column += (letter.toUpperCase().charCodeAt(i) - 64) * Math.pow(26, length - i - 1) - 1
    }
  } else {
    column = false;
  }

  return column;
}

/**
 * Source: https://blog.kevinchisholm.com/javascript/javascript-e164-phone-number-validation/
 */
function validatePhoneForE164 (phoneNumber) {
  const regEx = /^\+[1-9]\d{10,14}$/;
  return regEx.test(phoneNumber);
}