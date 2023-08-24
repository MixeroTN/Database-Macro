// --------------------------------------------------------------------------------------------------------------------------

// LinkBack to this script:
// http://webapps.stackexchange.com/questions/7211/how-can-i-make-some-data-on-a-google-spreadsheet-auto-sorting/43036#43036


// What to sort:
var columnToSortBy = 1;
var tableRange = "A2:C1000";


/**
* Automatically sorts the 1st column (not the header row) Ascending.
*/

function sort() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet(); //event.source.getActiveSheet()
  var editedCell = sheet.getActiveCell();

  if (editedCell.getColumn() == columnToSortBy) {
    var range = sheet.getRange(tableRange);

    range.sort({ column: columnToSortBy, ascending: false });
  }
}

function onEdit() {
  sort();
}

function onOpen() {
  sort();
}

// --------------------------------------------------------------------------------------------------------------------------

//  1. Run > setup
//
//  2. Publish > Deploy as web app
//    - enter Project Version name and click 'Save New Version'
//    - set security level and enable service (execute as 'me' and access 'anyone, even anonymously)
//
//  3. Copy the 'Current web app URL' and post this in your form/script action
//
//  4. Insert column names on your destination sheet matching the parameter names of the data you are passing in (exactly matching case)

var SCRIPT_PROP = PropertiesService.getScriptProperties();

function doGet(e) {
  try {
    var key = e.parameter["key"];
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    var sheet = doc.getSheetByName(e.parameter["sheet"]);
    var data = sheet.getDataRange().getValues();
    var value;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == key) {
        value = data[i][1];
      }
    }

    if (value) {
      return ContentService
        .createTextOutput(JSON.stringify({ "result": "success", "value": value }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService
        .createTextOutput(JSON.stringify({ "result": "error", "error": "Key not found" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "error", "error": "Database does not exist" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  var lock = LockService.getPublicLock();

  lock.waitLock(30000);

  try {
    var sheetNom = e.parameter["sheet"];
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    var sheet = doc.getSheetByName(e.parameter["sheet"]);
    var data = sheet.getDataRange().getValues();
    var range = sheet.getRange(tableRange);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow() + 1;
    var row = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == e.parameter["key"]) {
        nextRow = i + 1;
      }
    }

    for (i in headers) {
      row.push(e.parameter[headers[i]]);
    }

    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    range.sort({ column: columnToSortBy, ascending: false });

    return ContentService
      .createTextOutput(JSON.stringify({ "result": "success", "row": nextRow }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "error", "error": e, "test": "Test", "sheet": sheetNom }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function setup() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();

  SCRIPT_PROP.setProperty("key", doc.getId());
}
