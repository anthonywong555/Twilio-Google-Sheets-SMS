var CUSTOMER_PHONE_NUMBER = 0;
var CUSTOMER_NAME = 1;
var AMOUNT_DUE = 2;
var PAYMENT_DUE_DATE = 3;
var PAYMENT_LINK = 4; // A URL where the payment can be made
var PAYMENT_INFO = 5; // Details of the service or product
var MESSAGE_STATUS = 6; // Whether the SMS was sent or not

var TWILIO_ACCOUNT_SID = 'placeholder';
var TWILIO_PHONE_NUMBER = 'placeholder';
var TWILIO_AUTH_TOKEN = 'placeholder';

var ui = SpreadsheetApp.getUi();
var userProperties = PropertiesService.getUserProperties();

function onOpen() {
  ui.createMenu('Credentials')
    .addItem('Set Twilio Account SID', 'setTwilioAccountSID')
    .addItem('Set Twilio Auth Token', 'setTwilioAuthToken')
    .addItem('Set Twilio phone number', 'setTwilioPhoneNumber')
    .addItem('Delete Twilio Account SID', 'deleteTwilioAccountSID')
    .addItem('Delete Twilio Auth Token', 'deleteTwilioAuthToken')
    .addItem('Delete Twilio phone number', 'deleteTwilioPhoneNumber')
    .addToUi();
  ui.createMenu('Send SMS')
    .addItem('Send to all', 'sendSmsToAll')
    .addItem('Send to customers with due date 1st-15th', 'sendSmsByDateFilter')
    .addToUi();
};

function setTwilioAccountSID() {
  var scriptValue = ui.prompt('Enter your Twilio Account SID', ui.ButtonSet.OK);
  userProperties.setProperty('TWILIO_ACCOUNT_SID', scriptValue.getResponseText());
};

function setTwilioAuthToken() {
  var scriptValue = ui.prompt('Enter your Twilio Auth Token', ui.ButtonSet.OK);
  userProperties.setProperty('TWILIO_AUTH_TOKEN', scriptValue.getResponseText());
};

function setTwilioPhoneNumber() {
  var scriptValue = ui.prompt('Enter your Twilio phone number in this format: +12345678900', ui.ButtonSet.OK);
  userProperties.setProperty('TWILIO_PHONE_NUMBER', scriptValue.getResponseText());
};

function deleteTwilioAccountSID() {
  userProperties.deleteProperty('TWILIO_ACCOUNT_SID');
};

function deleteTwilioAuthToken() {
  userProperties.deleteProperty('TWILIO_AUTH_TOKEN');
};

function deleteTwilioPhoneNumber() {
  userProperties.deleteProperty('TWILIO_PHONE_NUMBER');
};

function sendSms(customerPhoneNumber, amountDue, paymentLink, customerName, paymentInfo, paymentDueDate) {
  var twilioAccountSID = userProperties.getProperty('TWILIO_ACCOUNT_SID');
  var twilioAuthToken = userProperties.getProperty('TWILIO_AUTH_TOKEN');
  var twilioPhoneNumber = userProperties.getProperty('TWILIO_PHONE_NUMBER');
  var twilioUrl = 'https://api.twilio.com/2010-04-01/Accounts/' + twilioAccountSID + '/Messages.json';
  var authenticationString = twilioAccountSID + ':' + twilioAuthToken;
  try {
    const response = UrlFetchApp.fetch(twilioUrl, {
      method: 'post',
      headers: {
        Authorization: 'Basic ' + Utilities.base64Encode(authenticationString)
      },
      payload: {
        To: "+" + customerPhoneNumber.toString(),
        Body: "Hello, " + customerName + ", your payment of $" + amountDue + " is outstanding" + " for " + paymentInfo + ". It was due on " + paymentDueDate + "." + " Please visit " + paymentLink + " to pay your balance. If you have any questions, contact us at support@example.com. Thanks!",
        From: twilioPhoneNumber, // Your Twilio phone number
      },
    });

    console.log(response);
    return response.getContentText();
  } catch (err) {
    return 'error: ' + err;
  }
};

function sendSmsToAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  var headers = rows.shift();
  rows.forEach(function(row) {
    row[MESSAGE_STATUS] = sendSms(row[CUSTOMER_PHONE_NUMBER], row[AMOUNT_DUE], row[PAYMENT_LINK], row[CUSTOMER_NAME], row[PAYMENT_INFO], row[PAYMENT_DUE_DATE]);
  });
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
};