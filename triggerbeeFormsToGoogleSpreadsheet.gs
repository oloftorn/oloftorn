//this is a function that fires when the webapp receives a GET request
function doGet(e) {
  return HtmlService.createHtmlOutput("request received");
}

//this is a function that fires when the webapp receives a POST request
function doPost(e) {
   var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = Math.max(sheet.getLastRow(),1);
  var jsonData = getJsonFromTriggerbeePayload(e.postData);

  AddResponses(sheet, lastRow, jsonData);

  SpreadsheetApp.flush();
  return HtmlService.createHtmlOutput("post request received");
}

function getJsonFromTriggerbeePayload(postData)
{
  var params = JSON.stringify(postData.contents);
  var jsonparams = JSON.parse(params);
  var decodedJson = decodeURIComponent(jsonparams);

  // remove "payload=" from json string
  decodedJson = decodedJson.substring(8);
  return JSON.parse(decodedJson);
}

function logmsg(msg){
    console.log(msg);
   // var debugCell = SpreadsheetApp.getActiveSheet().getRange(1,25,1,1).getCell(1,1);
   // debugCell.setValue(debugCell.getValue() + "\n" + msg);
}

function AddResponses(sheet, row, json)
{ 
  sheet.getRange(row + 1, 1).setValue(new Date().toISOString());
  var columnCounter = 2;

  if(json.visit.PersonalData)
  {
    if(json.visit.PersonalData.Email)
    {
      sheet.getRange(row + 1, columnCounter++).setValue(json.visit.PersonalData.Email);
    }
    if(json.visit.PersonalData.Firstname)
    {
      sheet.getRange(row + 1, columnCounter++).setValue(json.visit.PersonalData.Firstname);
    }
    if(json.visit.PersonalData.Lastname)
    {
      sheet.getRange(row + 1, columnCounter++).setValue(json.visit.PersonalData.Lastname);
    }
    if(json.visit.PersonalData.Name)
    {
      sheet.getRange(row + 1, columnCounter++).setValue(json.visit.PersonalData.Name);
    }
    if(json.visit.PersonalData.Telephone)
    {
      sheet.getRange(row + 1, columnCounter++).setValue(json.visit.PersonalData.Telephone);
    }
  }
  var fields = json.visit.Fields;
  for (let key in fields) {
      let val = fields[key];
      if(key.substring(0,3)!="pp_")
      {
        sheet.getRange(row + 1, columnCounter++).setValue(val);
      }
  }
  sheet.getRange(row + 1, columnCounter++).setValue("https://app.triggerbee.com/insights/visits?sessionid="+json.visit.SessionId);
  
  var lastHeader = getLastHeader(row);
  var headerColumnCounter = 1;
  if(lastHeader){
     logmsg("Header exists");
    for(let i = 1; i < lastHeader.getNumColumns(); i++)
    {
      var cell = lastHeader.getCell(1,i);
      if(cell.getValue()!="")
          headerColumnCounter++;
    }
  }else
    logmsg("Header does not exist");
  
  ApplyRegularStyle(sheet.getRange(row + 1, 1, 1, 10));

  logmsg("column: " + columnCounter + "; headerColumnCounter: " + headerColumnCounter + "; row: " + row);
    
  if(columnCounter != headerColumnCounter || row == 1)
  { 
     sheet.insertRowsAfter(row, 1);
     logmsg("Different column count in header (" + columnCounter + "; headerColumnCounter: " + headerColumnCounter + "). Building new header!");

     BuildHeaderRows(json, row+1);
  }

  SpreadsheetApp.flush();
}

// bold indicates a header row 
function getLastHeader(totRow)
{
    var sheet = SpreadsheetApp.getActiveSheet();
    for(let i = totRow; i > 0; i--)
  {
    var range = sheet.getRange(i, 1, i, 30);
     var cell = range.getCell(1,1);
     if(cell.getTextStyle().isBold())
      {
        return range; 
      }
  }
}

function ApplyHeaderStyle(range)
{
  var style = SpreadsheetApp.newTextStyle()
      .setFontSize(11)
    .setBold(true)
    .build();

    range.setBorder(null, null, true, null, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);

  range.setTextStyle(style);
}

function ApplyRegularStyle(range)
{
  var style = SpreadsheetApp.newTextStyle()
      .setFontSize(10)
    .setBold(false)
    .build();
  range.setTextStyle(style);
}

function BuildHeaderRows(json, row)
{
  logmsg("BuildHeaderRows");
  var sheet = SpreadsheetApp.getActiveSheet(); 
  ApplyHeaderStyle(sheet.getRange(row, 1, 1, 30))
  
  var columnCounter = 1;
  var cell = SpreadsheetApp.getActiveSheet().getRange(row, columnCounter);
  cell.setValue("Date");
  columnCounter++;
  if(json.visit.PersonalData)
  {
    if(json.visit.PersonalData.Email)
    {
      sheet.getRange(row, columnCounter++).setValue("Email");
    }
    if(json.visit.PersonalData.Firstname)
    {
      sheet.getRange(row, columnCounter++).setValue("Firstname");
    }
    if(json.visit.PersonalData.Lastname)
    {
      sheet.getRange(row, columnCounter++).setValue("Lastname");
    }
    if(json.visit.PersonalData.Name)
    {
      sheet.getRange(row, columnCounter++).setValue("Name");
    }
    if(json.visit.PersonalData.Telephone)
    {
      sheet.getRange(row, columnCounter++).setValue("Telephone");
    }
  }

  var fields = json.visit.Fields;
  for (let key in fields) {
        let val = fields[key];
        if(key.substring(0, 3)!="pp_"){
          SpreadsheetApp.getActiveSheet().getRange(row, columnCounter++).setValue(key);
        }
  }
  sheet.getRange(row, columnCounter++).setValue("Session");
}
