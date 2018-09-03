var sheetID = "19LhOLxFuOVHx2bE5BrZQVdUDK5Re3A0T4U6Fe6c0KO8"; //needed as you cannot use getActiveSheet() while the sheet is not in use (as in a standalone application like this one)
var favicon_url = 'https://icons.iconarchive.com/icons/iconsmind/outline/32/Quill-3-icon.png';


// ******************************************************************************************************
// Function to create menus when you open the sheet
// ******************************************************************************************************
function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{name: "Configure Jira", functionName: "jiraConfigure"},
                     {name: "Refresh Data Now", functionName: "jiraPullManual"},
                     {name: "Schedule 4 Hourly Automatic Refresh", functionName: "scheduleRefresh"},
                     {name: "Stop Automatic Refresh", functionName: "removeTriggers"},
                     {name: "Log the card creation metadata", functionName: "sendMetaToLogger"}]; 
  ss.addMenu("Jira", menuEntries);
}


// ******************************************************************************************************
// Function to display the HTML as a webApp
// ******************************************************************************************************
function doGet(e) {
  
  //pass a parameter via the URL as ?action=XXX
  var userAction = e.parameter.action;
  
  switch(userAction) {
     case "create":  
       var template = HtmlService.createTemplateFromFile('submit'); 
       var pageTitle = "JIRA Create";
     break;      
     default:
       var template = HtmlService.createTemplateFromFile('display');
       var pageTitle = "CJ Search & Match Roadmap";
  }
  var htmlOutput = template.evaluate()
                   .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                   .setTitle(pageTitle)
                   .addMetaTag('viewport', 'width=device-width, initial-scale=1')
                   .setFaviconUrl(favicon_url);

  return htmlOutput;
};


// ******************************************************************************************************
// Function to print out the content of an HTML file into another (used to load the CSS and JS)
// ******************************************************************************************************
function getContent(filename) {
  var pageContent = HtmlService.createTemplateFromFile(filename).getRawContent();
  return pageContent;
}


function printVal(key) {
  var dummy = PropertiesService.getUserProperties().getProperty(key);
  return dummy;
}


// ******************************************************************************************************
// Function to create a new item in JIRA
// ******************************************************************************************************
function JIRAsubmit(data) {

  var url = "https://" + PropertiesService.getUserProperties().getProperty("host") + "/rest/api/2/issue/";
  var authVal = PropertiesService.getUserProperties().getProperty("digest");  
  var options = { 
              "Accept":"application/json", 
              "contentType":"application/json", 
              "method": "POST",
              "payload": data,
              "headers": {"Authorization": authVal}           
             };
  
  
  var resp = UrlFetchApp.fetch(url,options);
  
  //if it all goes well the response should be 201 - the JIRA documentation is incorrect. See https://jira.atlassian.com/browse/JRASERVER-39339
  if (resp.getResponseCode() != 201) {
    Logger.log("Error retrieving data for url " + url + ":" + resp.getContentText());
  }  
  else {
    var STRINGresp = resp.getContentText();
    Logger.log('Item has been created');
    Logger.log(STRINGresp);
    var respJSON = JSON.parse(STRINGresp);
    Browser.msgBox('Item ' + respJSON.self + ' has been created.'); 
  }  
  
}


// ******************************************************************************************************
// Get the content of a google sheet and convert it into a json - from https://gist.github.com/daichan4649/8877801
// ******************************************************************************************************
function convertSheet2JsonText(sheet) {
  // first line(title)
  var colStartIndex = 1;
  var rowNum = 1;
  var firstRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  var firstRowValues = firstRange.getValues();
  var titleColumns = firstRowValues[0];

  // after the second line(data)
  var lastRow = sheet.getLastRow();
  var rowValues = [];
  for(var rowIndex=2; rowIndex<=lastRow; rowIndex++) {
    var colStartIndex = 1;
    var rowNum = 1;
    var range = sheet.getRange(rowIndex, colStartIndex, rowNum, sheet.getLastColumn());
    var values = range.getValues();
    rowValues.push(values[0]);
  }

  // create json
  var jsonArray = [];
  for(var i=0; i<rowValues.length; i++) {
    var line = rowValues[i];
    var json = new Object();
    for(var j=0; j<titleColumns.length; j++) {
      json[titleColumns[j]] = line[j];
    }
    jsonArray.push(json);
  }
  return jsonArray;
}

// ******************************************************************************************************
// Wrapper for the convertSheet2JsonText - pass the sheetName to get the JSON array out
// ******************************************************************************************************
function sheet2Json(sheetName) {
  var ss = SpreadsheetApp.openById(sheetID);
  var sheets = ss.getSheets();
  var sh = ss.getSheetByName(sheetName);
  return convertSheet2JsonText(sh);
}


// ******************************************************************************************************
// Generate the HTML for the display page
// ******************************************************************************************************
function printColumnsWithItems(){
  
  var itemJSON = sheet2Json("SM");
  
  var htmlParked = '';
  var htmlIdeas = '';
  var htmlPlanned = '';
  var htmlOngoing = '';
  var htmlDone = '';
  
  var itemHasSprint = false;
  var itemIsDone = false;
  
  
  for (var i=0; i<itemJSON.length; i++) {
    var jItem = itemJSON[i];  
    var id = jItem["Key"]
    var type = jItem["Issuetype.name"];
    var displayText = jItem["Summary"];
    var epic = jItem["customfield_10008"];
    var sprint = jItem["Sprint.name"];
    var status = jItem["Status.name"];
    var itemLabel = jItem["Labels"];

    if (type != "Epic") {
      if (sprint){  
        if (status == "Done") {
            htmlDone += createItemHMTL(id, type, displayText, epic, "Done" )
          } else {
            htmlOngoing += createItemHMTL(id, type, displayText, epic, "Ongoing" )
          }
      } else {
        switch(itemLabel){
          case "Parked":
            htmlParked += createItemHMTL(id, type, displayText, epic, "Parked" )
            break;
          case "Planned":
            htmlPlanned += createItemHMTL(id, type, displayText, epic, "Planned" )
            break;
          default:
            htmlIdeas += createItemHMTL(id, type, displayText, epic, "Idea" )
        }  
      }
    }
  }
  
  var outHTML = '';
  outHTML += '<div class="column" id="Parked-wrapper">' +
		       '<h2>Parked</h2>' +
		       '<div class="item-container sortable-area" id="Parked">' +
		         htmlParked +
		       '</div>' +
             '</div>';
  outHTML += '<div class="column" id="Idea-wrapper">' +
		       '<h2>Ideas</h2>' +
		       '<div class="item-container sortable-area" id="Idea">' +
		         htmlIdeas +
		       '</div>' +
             '</div>';  
  outHTML += '<div class="column" id="Planned-wrapper">' +
		       '<h2>Planned</h2>' +
		       '<div class="item-container sortable-area" id="Planned">' +
		         htmlPlanned +
		       '</div>' +
             '</div>';
  outHTML += '<div class="column" id="Ongoing-wrapper">' +
		       '<h2>Ongoing</h2>' +
		       '<div class="item-container" id="Ongoing">' +
		         htmlOngoing +
		       '</div>' +
             '</div>';
  outHTML += '<div class="column" id="Done-wrapper">' +
		       '<h2>Done</h2>' +
		       '<div class="item-container" id="Done">' +
		         htmlDone +
		       '</div>' +
             '</div>';
  return outHTML;
}

// ******************************************************************************************************
// Generate the HTML for an item for the display page
// ******************************************************************************************************
function createItemHMTL(id, type, displayText, epic, column ){
   var itemHTML = '<li id="' + id + '" data-column="' + column + '">' +
                     '<span class="drag-handle">☰</span>' +
                     '<div class="type-' + type + '">' + displayText + '</div>' +
                     '<span class="epic" class="epic-'+ epic +'"></span>' +
                   '</li>'; 
   return itemHTML;
}