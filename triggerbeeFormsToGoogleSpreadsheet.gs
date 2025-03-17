// Handles GET requests
function doGet(e) {
  console.log("doGet fired at: " + new Date().toISOString());
  return HtmlService.createHtmlOutput("Request received");
}

// Handles POST requests
function doPost(e) {
  console.log("doPost fired at: " + new Date().toISOString());
  console.log("Incoming postData: " + e.postData.contents);
  
  try {
    // Open your spreadsheet by ID.
    var ss = SpreadsheetApp.SpreadsheetApp.getActiveSheet();
    
    var sheet = ss.getSheetByName("Sheet1");
    
    // Parse the incoming JSON payload.
    var jsonData = JSON.parse(e.postData.contents);
    console.log("Parsed JSON data: " + JSON.stringify(jsonData));
    
    // Ensure header row (row 1) is current with the incoming data.
    ensureHeader(sheet, jsonData);
    
    // Determine the row to insert data.
    var headerRow = 1;
    var dataRow = sheet.getLastRow() + 1; // Data rows follow header row.
    
    // Add the data row.
    AddResponses(sheet, dataRow, jsonData);
    
    SpreadsheetApp.flush();
    console.log("doPost completed at: " + new Date().toISOString());
    return HtmlService.createHtmlOutput("POST request received");
  } catch (error) {
    console.log("Error in doPost: " + error);
    return HtmlService.createHtmlOutput("Error: " + error);
  }
}

// Ensures that the header row (row 1) matches the expected structure based on the JSON.
function ensureHeader(sheet, json) {
  var expectedHeaders = getExpectedHeaders(json);
  var currentHeaders = [];
  
  // If the sheet is empty, currentHeaders stays empty.
  if (sheet.getLastRow() >= 1 && sheet.getLastColumn() > 0) {
    currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }
  
  // Compare the current header row with the expected header.
  if (currentHeaders.length === 0 || currentHeaders.join('|') !== expectedHeaders.join('|')) {
    // Write (or overwrite) the header row in row 1.
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    ApplyHeaderStyle(sheet.getRange(1, 1, 1, expectedHeaders.length));
    console.log("Header row updated to: " + expectedHeaders.join(", "));
  } else {
    console.log("Header row is up-to-date.");
  }
}

// Builds the expected header array based on known fields and dynamic fields.
function getExpectedHeaders(json) {
  // Fixed headers for known fields.
  var headers = [
    "Timestamp",   // Data insertion time
    "SessionId",
    "Uid",
    "Time",
    "Identifier",
    "GeoLocation",
    "Email",
    "Name",
    "FirstName",
    "LastName",
    "Organization",
    "Phone",
    "Username",
    "Title"
  ];
  
  // Dynamic headers from visit.Fields.
  if (json.visit && json.visit.Fields) {
    var fieldKeys = Object.keys(json.visit.Fields);
    fieldKeys.sort(); // Sort alphabetically for consistency.
    headers = headers.concat(fieldKeys);
  }
  
  // Append the root-level "text" field.
  headers.push("text");
  return headers;
}

// Appends a new data row based on the JSON payload.
function AddResponses(sheet, row, json) {
  var col = 1;
  // Column 1: Timestamp of insertion.
  sheet.getRange(row, col++).setValue(new Date().toISOString());
  
  var visit = json.visit;
  if (visit) {
    // Fixed properties.
    if (visit.SessionId !== undefined) {
      sheet.getRange(row, col++).setValue(visit.SessionId);
    } else { col++; }
    if (visit.Uid !== undefined) {
      sheet.getRange(row, col++).setValue(visit.Uid);
    } else { col++; }
    if (visit.Time) {
      sheet.getRange(row, col++).setValue(visit.Time);
    } else { col++; }
    if (visit.Identifier) {
      sheet.getRange(row, col++).setValue(visit.Identifier);
    } else { col++; }
    if (visit.GeoLocation) {
      sheet.getRange(row, col++).setValue(visit.GeoLocation);
    } else { col++; }
  } else {
    col += 5; // Skip 5 columns if no visit.
  }
  
  // PersonalData fields.
  var pd = (visit && visit.PersonalData) ? visit.PersonalData : {};
  // For each expected personal data field, set value or leave blank.
  var pdFields = ["Email", "Name", "FirstName", "LastName", "Organization", "Phone", "Username", "Title"];
  pdFields.forEach(function(field) {
    sheet.getRange(row, col++).setValue(pd[field] || "");
  });
  
  // Dynamic fields from visit.Fields.
  if (visit && visit.Fields) {
    // To match the header order, sort keys alphabetically.
    var fieldKeys = Object.keys(visit.Fields).sort();
    fieldKeys.forEach(function(key) {
      sheet.getRange(row, col++).setValue(visit.Fields[key]);
    });
  }
  
  // Root-level "text" field.
  sheet.getRange(row, col++).setValue(json.text || "");
  
  console.log("Data row " + row + " added, ending at column " + (col - 1));
  SpreadsheetApp.flush();
}

// Applies header style (bold, font size 11, border)
function ApplyHeaderStyle(range) {
  var style = SpreadsheetApp.newTextStyle()
      .setFontSize(11)
      .setBold(true)
      .build();
  
  range.setBorder(null, null, true, null, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  range.setTextStyle(style);
}

// Test function for manual execution in the Apps Script editor.
function testFunction() {
  var testData = {
    "visit": {
      "SessionId": 1026578280,
      "Uid": 1597130962,
      "Time": "2025-02-18T14:23:02",
      "Identifier": "testar@testar.com",
      "NetName": null,
      "GeoLocation": "Hägersten, Sweden",
      "Tags": null,
      "Goals": [
        "Submitted the form \"Test Webhook\""
      ],
      "Revenue": 0,
      "Consent": {
        "Title": null,
        "CurrentWebsiteUrl": null,
        "AcceptedTitle": null,
        "AcceptedUrl": null,
        "PolicyUrl": null,
        "PolicyText": null,
        "IsCheckboxPresent": null,
        "PolicyVersionDate": null,
        "PolicyRevisionNumber": null
      },
      "PersonalData": {
        "Email": "testar@testar.com",
        "Name": "Testar",
        "FirstName": null,
        "LastName": null,
        "Organization": null,
        "Phone": "11111111111111",
        "Username": null,
        "Title": null
      },
      "Fields": {
        "TextInput1": "Hello TestArFriurcuy",
        "SmileyRating1": "5",
        "responseListId": "7748ae88-e20e-4c82-be80-d032007ba35e",
        "submitPath": "/",
        "more test": "jajajajaj"
      },
      "Utm": null
    },
    "text": "Testar visited your website, viewed 1 pages and completed goals \"Submitted the form \"Test Webhook\"\".",
    "blocks": null
  };
  
  // Construct a fake event object as if it were a POST request.
  var e = {
    postData: {
      type: "application/json",
      contents: JSON.stringify(testData)
    }
  };
  
  // Call doPost with the fake event.
  var result = doPost(e);
  console.log("Test function result: " + result.getContent());
}
