function formSubmitReply(e){  

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheets()[0];
  var lastRow = dataSheet.getLastRow();
  
  dataSheet.getRange(lastRow, getColIndexByName("Computed_Answer")).setFormula('=SUBSTITUTE('+columnToLetter(getColIndexByName("Answer"))+lastRow+',char(10),"<br>")');
     
  switch(dataSheet.getRange(lastRow, getColIndexByName("Success - Question Type 1")).getValue()) {
    case "Account Management": dataSheet.getRange(lastRow, getColIndexByName("Status")).setValue("Redirected");
                               dataSheet.getRange(lastRow, getColIndexByName("Answer")).setValue("No Answer Required! -> MC has already sent an Autoresponder");
        break;
    case "Technical Support": dataSheet.getRange(lastRow, getColIndexByName("Status")).setValue("Redirected");
                              dataSheet.getRange(lastRow, getColIndexByName("Answer")).setValue("No Answer Required! -> MC has already sent an Autoresponder");
        break;
    default: dataSheet.getRange(lastRow, getColIndexByName("Status")).setValue("New");
    }
}

function sendEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheets()[0];
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows() - 1, dataSheet.getMaxColumns());
  var emailErrorCount = [];
  var emailSendCount = [];
  var templateSheet = ss.getSheets()[1];

  // Create one JavaScript object per row of data.
  var objects = getRowsData(dataSheet, dataRange);

  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var i = 0; i < objects.length; ++i) {
          
    // Get a row object
      var rowData = objects[i];
    
    if(objects[i].status == "Send" && objects[i].answer != null)
    {
      var emailTemplate = templateSheet.getRange("A13").getValue();
      
      // Given a template string, replace markers (for instance ${"First Name"}) with
      // the corresponding value in a row object (for instance rowData.firstName).
      var emailText = fillInTemplateFromObject(emailTemplate, rowData);
      var alias = "mcsuccesshub@salesforce.com";
      //var fromemail = Session.getActiveUser().getEmail();
      var fromvalue = templateSheet.getRange("A4").getValue(); 
      var fromname = templateSheet.getRange("A7").getValue();
      var replyAdd = Session.getActiveUser().getEmail();
      var emailSubject = templateSheet.getRange("A10").getValue();    
      
      switch(fromvalue){
        case 'Personal' : GmailApp.sendEmail(rowData.email, emailSubject, emailText,{htmlBody: emailText,name:fromname, replyTo: replyAdd}); break;
        case 'MC Success Hub' : GmailApp.sendEmail(rowData.email, emailSubject, emailText,{htmlBody: emailText, from:alias, name:fromname, replyTo: replyAdd});
      }
      //update status column to completed
      var IndexData = rowData.id;
      dataSheet.getRange(getRowfromID(dataSheet,IndexData), getColIndexByName("Status")).setValue("Completed");
      dataSheet.getRange(getRowfromID(dataSheet,IndexData), getColIndexByName("User")).setValue(Session.getActiveUser().getEmail());
      emailSendCount.push(rowData.id);
    }
  
    else {
           if(objects[i].status == "Send")
           {emailErrorCount.push(rowData.id)};
          }
  }
  //After event alert
  var ui = SpreadsheetApp.getUi(); // Same variations.
  if(emailErrorCount.length > 0) //no emails sent
     {var ErrorResult = ui.alert('The Following Emails did not send: '+emailErrorCount);}
  else
     {var SendResult = ui.alert('The Following Emails were successfully sent: '+emailSendCount);}
}

function getRowfromID(SheetRange,data) {
// get the values in an array
var range = SheetRange.getRange(2, getColIndexByName("ID"), SheetRange.getMaxRows() - 1);
var values = range.getValues();
// examin the values in the array
var i = []; 
for (var y = 0; y < values.length; y++) {
   if(values[y] == data){
      i.push(y);
   }
}
// translate the indices into rows
return Number(i)+Number(range.getRow());    
}

function columnToLetter(column){
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter){
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance ${"Column name"}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker ${"Column name"}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function fillInTemplateFromObject(template, data) {
  var email = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);

  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string.
  for (var i = 0; i < templateVars.length; ++i) {
    // normalizeHeader ignores ${"} so we can call it directly here.
    var variableData = data[normalizeHeader(templateVars[i])];
    email = email.replace(templateVars[i], variableData || "");
  }

  return email;
}





//////////////////////////////////////////////////////////////////////////////////////////
//
// The code below is reused from the 'Reading Spreadsheet data using JavaScript Objects'
// tutorial.
//
//////////////////////////////////////////////////////////////////////////////////////////

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

function getColIndexByName(colName) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var numColumns = sheet.getLastColumn();
  var row = sheet.getRange(1, 1, 1, numColumns).getValues();
  for (i in row[0]) {
    var name = row[0][i];
    if (name == colName) {
      return parseInt(i) + 1;
    }
  }
  return -1;
}

function onOpen() {
  var subMenus = [{name:"Send Emails", functionName: "sendEmails"}];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Email", subMenus);
}
