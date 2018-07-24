/*   FILE   : contact.gs (Apps Script)
 *   AUTHOR : Jiin Jeong
 *   DATE   : June 19, 2018 (Completed),
 *            July 23, 2018 (Cleaned)
 *   DESC   : Sends texts and e-mails from Google Sheets.
 *            (Sensitive info removed.)
 */

/******************************** SHEETS ********************************/
// Creates a custom menu for the Google Sheets.
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Contact')
      .addItem('E-mail', 'personalEmail')
      .addItem('Text', 'personalText')
      .addToUi();
}

// Gets contact data from Google Sheets.
function getData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var startCol = 1;
  var numRows = 500;  // Num of rows to process (MAX)
  var numCols = 6;
  
  // Fetch range of cells.
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols)
  var data = dataRange.getValues();
  Logger.log(data)  // Debugs.
  return data;
}

/******************************** CONTENT ********************************/
// Gets e-mail content from an external Google Document.
function Content() {
  var doc = DocumentApp.openById("DocumentID");  // Change.
  var body = doc.getBody();
  var text = body.getText();
  return text;
}

// Gets weather info from Yahoo Weather. (Method 1: JSON, 2: XML Response)
// Sometimes breaks - it seems like you have to run it after at least 1 min wait.
function weatherInfo(){
  // Pretty-prints JSON in Terminal shell to show the data more clearly.
  // https://stackoverflow.com/questions/12943819/how-to-prettyprint-a-json-file
  // cat some.json | python -m json.tool

  var url="https://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20weather.forecast%20where%20woeid%20=%202380633%20%20&format=json";
  var response = UrlFetchApp.fetch(url);  // Gets feed.
  var weatherData = JSON.parse(response.getContentText());
  
  // Gets 10-day forecast information from the JSON file.
  var forecast = weatherData.query.results.channel.item.forecast;
  var total_high = 0;
  
  // Gets high temperature of each day and adds up.
//  for (var i = 0; i < forecast.length; i ++) {
  for (var i = 0; i < 7; i ++) {
    total_high += Number(forecast[i].high); // Switches string to num.
  }

  // Finds average high temperature to the third decimal place and returns it.
  var avg_high = total_high/7;
  avg_high = avg_high.toFixed(3);
  return avg_high;
}

function weatherContent() {
  var content = "A friendly reminder to water your tree! This week's average high temperature will be %s. ";
  content = content.replace('%s', weatherInfo().toString());

  // Reminds water frequency based on the week's average temperature.
  if (weatherInfo() < 85) {
    content += "Please water your tree once this week.";
  }
  else if (weatherInfo() > 100) {
    content += "Please water your tree three times this week.";
  }
  else {
    content += "Please water your tree two times this week.";
  }
  
  content += "\n \nFrom your Sustainable Claremont Team.";
  return content;
}

/******************************** E-MAIL ********************************/
// Sends e-mail based on user preferrance.
function sendEmails(subject, content) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = getData();
  var today = Utilities.formatDate(new Date(), "America/Los_Angeles", "MM/dd")

  for (var i = 0; i < data.length; i ++) {
    var row = data[i];
    var name = row[0];  // Column 1: Name
    var rule = row[1];  // Column 2: Preferred

    var body = "Hello %s, \n \n"
    body = body.replace("%s", name);
    body += content;

    if (rule == "E-mail") {
        var emailAddress = row[3];  // Column 4: E-mail Address
        MailApp.sendEmail(emailAddress, subject, body);
        sheet.getRange(2 + i, 6).setValue(today);  // Start row, col
        // Make sure the cell is updated right away in case the script is interrupted
        SpreadsheetApp.flush();
    }
  }
}

// Sends user-generated e-mails based on content from Google Sheets.
function personalEmail() {
  var subject = "[Sustainable Claremont] From Your Sustainable Claremont Team";
  var content = Content();
  sendEmails(subject, content);
}

// Sends automatic weekly water-reminder e-mails.
function weeklyEmail() {
  var subject = "[Sustainable Claremont] Weekly Water Trees Reminder";
  var content = weatherContent();
  sendEmails(subject, content);
}

/******************************** TEXT ********************************/
// CITE : https://www.twilio.com/blog/2016/02/send-sms-from-a-google-spreadsheet.html
// DESC : Provided starting code for configuring environment to send text via Twilio.
function TwilioText(to, body)  // Change.
  // Twilio Messaing API.
  var messages_url = "https://api.twilio.com/2010-04-01/Accounts/SID/Messages.json";

  // Form data or "payload" information.
  var formData = {
    "To": to,
    "Body" : body,
    "From" : "phone#"
  };

  // Tell Apps Script that this is a POST request that uses the payload.
  var options = {
    "method" : "post",
    "payload" : formData
  };

  // Authorizes the request with Account SID & token.
  options.headers = { 
    "Authorization" : "Basic " + Utilities.base64Encode("SID:token")
  };
 
  // Executes HTTP request.
  UrlFetchApp.fetch(messages_url, options);
}

// Sends text based on user preferrance.
function sendTexts(content) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = getData();
  var today = Utilities.formatDate(new Date(), "America/Los_Angeles", "MM/dd")

  for (var i = 0; i < data.length; i ++) {
    var row = data[i];
    var name = row[0];  // Column 1: Name
    var rule = row[1];  // Column 2: Preferred.

    var body = "[Sustainable Claremont]\nHello %s, \n \n"
    body = body.replace("%s", name);
    body += content;

    if (rule == "Text") {
        var phoneNum = row[2];  // Column 3: Phone num.
        TwilioText(phoneNum, body);
        sheet.getRange(2 + i, 6).setValue(today);  // Start row, col
        // Updates cell right away in case the script is interrupted.
        SpreadsheetApp.flush();
    }
  }
}

// Sends user-generated texts based on content from Google Sheets.
function personalText() {
  var content = Content();
  sendTexts(content);
}

// Sends automatic weekly water-reminder texts.
function weeklyText() {
  var content = weatherContent();
  sendTexts(content);
}
