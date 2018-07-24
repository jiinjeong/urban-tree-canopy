/* File   : staticmap.gs (Apps Script)
 * Author : Jiin Jeong
 * Date   : June 29, 2018 (Completed),
 *          July 21, 2018 (Cleaned)
 * Desc   : Generates URL of a static map with the given tree locations.
 */

/******************************** SHEETS ********************************/
// Creates customized mens.
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu("Analyze")
    .addItem("Create Static Map", "createMap")
    .addToUi();
}

// Makes pop-ups.
function popUp(message){
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
      "Please complete before continuing",
      message,
      ui.ButtonSet.OK_CANCEL);

  // Processes the user's response.
  var button = result.getSelectedButton();
  var response = result.getResponseText();
  
  if (button == ui.Button.OK) {
    return response;
  }
  else if (button == ui.Button.CANCEL) {  // User clicks 'Cancel'.
    return false;
  }
  else if (button == ui.Button.CLOSE) {  // User clicks 'X'.
    ui.alert("You closed the dialog.");
    return false;
  }
}

/******************************** DATA ********************************/
// Gets data from Google Sheets.
function getData(sheetName, startRow, endRow) {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  // Fetch range of cells.
  var startCol = 6;  // Coordinates.
  var numRows = endRow - startRow + 1;
  var numCols = 2;
  
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols);
  return dataRange;
}

// Changes a color into a hex color code from a dictionary.
function toHex(color){
  color = color.toUpperCase();  // Accounts for upper/lowercase errors.
  var color_dict = {"RED": "0xFF0000", "ORANGE": "0xFF9900",
                    "YELLOW": "0xFFFF00", "GREEN": "0x00FF00",
                    "BLUE": "0x0000FF", "PURPLE": "0x9900FF",
                    "PINK": "0xFF00FF", "WHITE": "0xFFFFFF",
                    "BLACK": "0x000000"};

  return color_dict[color];  // Indexes.
}

/******************************** MAP ********************************/
// Makes static map (markers are default icons).
function staticMap(sheetName, startRow, endRow, hex) {
  var dataRange = getData(sheetName, startRow, endRow);
  var data = dataRange.getValues();
  var font = dataRange.getFontLines();

  var map = Maps.newStaticMap()
                .setSize(1500, 1200)
                .setMarkerStyle(Maps.StaticMap.MarkerSize.MID, hex, 'T');

  for (var i = 0; i < data.length; i ++) {
    var row = data[i];
    var row_font = font[i][0];  // Gets the line style of the row's first cell.

    // Skips strike-through rows and N/A location values.
    if (row_font != "line-through" && row[0] != "N/A") {
      var x = Number(row[0]);  // x = Longitude.
      var y = Number(row[1]);  // y = Latitude.
      map.addMarker(y, x);
    }
  }
  return map.getMapUrl();
}

// Appends map URL result to a different sheet called "Map."
function urlMap(sheetName, startRow, endRow, hex){
  var spreadsheet = SpreadsheetApp.getActive();
  var result = spreadsheet.getSheetByName("Map");
  result.clear();  // Clears spreadsheet.
  
  var url = staticMap(sheetName, startRow, endRow, hex);
  result.appendRow(["URL to map", url]);  // Appends new data.
}

/******************************** MAIN ********************************/
// Creates map with pop-up data that prompts value for range.
function createMap(){
  var sheetName = popUp("Enter name of the sheet you would like to analyze.\n"+
                        "REMINDER: Watch your spelling.");

  // Turns response (type str) into integer.
  var startRow = parseInt(popUp("Enter the starting row number.\n"+
                                "REMINDER: Row number, not tree number."));
  var endRow = parseInt(popUp("Enter the ending row number.\n"+
                              "REMINDER: Row number, not tree number."));

  var color = popUp("Choose a color for your marker.\n"+
                    "CHOICES: Red, Orange, Yellow, Green, Blue,\n"+
                    "         Purple, Pink, Black, White.");

  urlMap(sheetName, startRow, endRow, toHex(color));  // Result: Map URL
}
