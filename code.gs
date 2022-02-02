/**
 * Twilio Information
 */
const TWILIO_ACCOUNT_SID = 'TWILIO_ACCOUNT_SID';
const TWILIO_AUTH_TOKEN = 'TWILIO_AUTH_TOKEN';
const TWILIO_PHONE_NUMBER = 'TWILIO_PHONE_NUMBER';

/**
 * Columns
 */
const CONTACT_NAME_COLUMN = 'CONTACT_NAME_COLUMN';
const CONTACT_PHONE_NUMBER_COLUMN = 'CONTACT_PHONE_NUMBER_COLUMN';
const MESSAGE_COLUMN = 'MESSAGE_COLUMN';
const SID_COLUMN = 'SID_COLUMN';
const STATUS_COLUMN = 'STATUS_COLUMN';

/**
 * Twilio Message Status
 */
const MESSAGE_STATUS_FINAL = ['delivered', 'undelivered', 'failed'];

/**
 * Google Sheets UI
 */
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
    .addItem('Set SID Column', 'setSIDColumn')
    .addItem('Set Status Column', 'setStatusColumn')
    .addItem('Delete Twilio Account SID', 'deleteTwilioAccountSID')
    .addItem('Delete Twilio Auth Token', 'deleteTwilioAuthToken')
    .addItem('Delete Twilio phone number', 'deleteTwilioPhoneNumber')
    .addToUi();
  ui.createMenu('Twilio - Actions')
    .addItem('Send SMS to All', 'sendSmsToAll')
    .addItem('Check Status', 'checkStatus')
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

function setContactPhoneNumberColumn () {
  const scriptValue = ui.prompt('Enter your Contact Phone Number Column', ui.ButtonSet.OK);
  userProperties.setProperty(CONTACT_PHONE_NUMBER_COLUMN, scriptValue.getResponseText());
}

function setMessageColumn () {
  const scriptValue = ui.prompt('Enter your Message Column', ui.ButtonSet.OK);
  userProperties.setProperty(MESSAGE_COLUMN, scriptValue.getResponseText());
}

function setSIDColumn () {
  const scriptValue = ui.prompt('Enter your SID Column', ui.ButtonSet.OK);
  userProperties.setProperty(SID_COLUMN, scriptValue.getResponseText());
}

function setStatusColumn () {
  const scriptValue = ui.prompt('Enter your Status Column', ui.ButtonSet.OK);
  userProperties.setProperty(STATUS_COLUMN, scriptValue.getResponseText());
}

/**
 * Twilio / Helper Methods
 */

function fetchMessage(messageSID) {
  const twilioAccountSID = userProperties.getProperty('TWILIO_ACCOUNT_SID');
  const twilioAuthToken = userProperties.getProperty('TWILIO_AUTH_TOKEN');
  const twilioUrl = `https://api.twilio.com/2010-04-01/Accounts/${twilioAccountSID}/Messages/${messageSID}.json`;
  const authenticationString = twilioAccountSID + ':' + twilioAuthToken;

  try {
    const response = UrlFetchApp.fetch(twilioUrl, {
      method: 'get',
      headers: {
        Authorization: 'Basic ' + Utilities.base64Encode(authenticationString)
      }
    });

    return response.getContentText();
  } catch (err) {
    return err;
  }
}

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
    return err;
  }
};

function checkStatus() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  const sidColumnIndex = letterToColumn(userProperties.getProperty(SID_COLUMN));
  const statusColumnIndex = letterToColumn(userProperties.getProperty(STATUS_COLUMN));

  for (const aRow of rows) {
    const sid = aRow[sidColumnIndex];
    const status = aRow[statusColumnIndex];

    if(sid && status && !MESSAGE_STATUS_FINAL.includes(status)) {
      // Fetch the Message
      const result = fetchMessage(sid);
      if(typeof result === 'string') {
        const {status} = JSON.parse(result);
        aRow[statusColumnIndex] = status;
      } else {
        aRow[statusColumnIndex] = result;
      }

      
    }
  }

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function sendSmsToAll() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  const sidColumnIndex = letterToColumn(userProperties.getProperty(SID_COLUMN));
  const statusColumnIndex = letterToColumn(userProperties.getProperty(STATUS_COLUMN));
  const messageColumnIndex = letterToColumn(userProperties.getProperty(MESSAGE_COLUMN));
  const contactPhoneNumberColumnIndex = letterToColumn(userProperties.getProperty(CONTACT_PHONE_NUMBER_COLUMN));
  
  const twilioPhoneNumber = userProperties.getProperty(TWILIO_PHONE_NUMBER);

  rows.forEach((aRow) => {
    const contactPhoneNumber = aRow[contactPhoneNumberColumnIndex];
    const message = aRow[messageColumnIndex];
    // Check to see is Contact Phone Number and Message
    if(validatePhoneForE164(contactPhoneNumber) && message) {
      const payload = {
        To: contactPhoneNumber,
        From: twilioPhoneNumber,
        Body: message
      };

      const result = sendSms(payload);

      // Check to see if the result row is declare.
      if(sidColumnIndex && statusColumnIndex && typeof result === 'string') {
        const {sid, status} = JSON.parse(result);
        aRow[sidColumnIndex] = sid;
        aRow[statusColumnIndex] = status;
      } else {
        aRow[sidColumnIndex] = result;
      }
    } else {
      let errorMessage = [];

      if(!validatePhoneForE164(contactPhoneNumber)) {
        errorMessage.push(`Please double check the phone number. It needs to be the following format: +155555555555`);
      }

      if(!message) {
        errorMessage.push(`Please include an message`);
      }

      // Check to see if the result row is declare.
      if(sidColumnIndex) {
        aRow[sidColumnIndex] = errorMessage.join('\n');
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