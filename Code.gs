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
// Function to return the HTML of the roadmap
// ******************************************************************************************************
function returnRoadmapHtml(e) {
  var html = HtmlService
        .createTemplateFromFile('display') // uses templated html
        .evaluate()
        .getContent();
      return html;
}

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
  if (userIsEditor()){
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
  } else {
    Browser.msgBox('You do not have the permission to create a new item via this interface'); 
  }
}




function userIsEditor(){
  var userEmail = Session.getActiveUser().getEmail();
  var editorEmails = PropertiesService.getScriptProperties().getProperty("editorEmails");
  if ( editorEmails.indexOf(userEmail) != -1 ) {
    return true;
  } else {  
    var file = DriveApp.getFileById(sheetID);
    var editors = file.getEditors();
    var isEditor = false;
    var eLength = editors.length
    for (var ed=0; ed<eLength; ed++) {
      if (editors[ed].getEmail() == userEmail ){
        isEditor = true
      }
    }
    if (file.getOwner().getEmail() == userEmail ){
      isEditor = true
    }
  }
}


// ******************************************************************************************************
// Generate the HTML for the display page
// ******************************************************************************************************
function printColumnsWithItems(){
  
  Logger.log('job started');
  
  var ss = SpreadsheetApp.openById(sheetID);
  var sheetName = PropertiesService.getUserProperties().getProperty("projectKey");
  var sheet = ss.getSheetByName(sheetName);
    
  //get all data
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn()
  var range = sheet.getRange(1, 1, lastRow, lastCol);
  var allValues = range.getValues();
  
  var id_col = -1;
  var type_col = -1;
  var displayText_col = -1;
  var itemLabel_col = -1;
  var sprint_col = -1;
  var sprintStartDate_col = -1;
  var status_col = -1;
  var epic_col = -1;
  var id_col_txt = '<li>Key</li>';
  var type_col_txt = '<li>Issuetype.name</li>';
  var displayText_col_txt = '<li>Summary</li>';
  var itemLabel_col_txt = '<li>Labels</li>';
  var sprint_col_txt = '<li>Sprint.name</li>';
  var sprintStartDate_col_txt = '<li>Sprint.startDate</li>';
  var status_col_txt = '<li>Status.name</li>';
  var epic_col_txt = '<li>customfield_10008</li>';
  
  for (var m=0; m<lastCol; m++){
    var dummy = allValues[0][m].toLowerCase();
    switch(dummy){
      case "key":   
        id_col = m;
        id_col_txt = '';
        break;
      case "issuetype.name":   
        type_col = m;
        type_col_txt = '';
        break;
      case "summary":   
        displayText_col = m;
        displayText_col_txt = '';
        break;  
      case "labels":   
        itemLabel_col = m;
        itemLabel_col_txt = '';
        break;  
      case "sprint.name":   
        sprint_col = m;
        sprint_col_txt = '';
        break;  
      case "sprint.startDate":   
        sprintStartDate_col = m;
        sprintStartDate_col_txt = '';
        break;  
      case "status.name":   
        status_col = m;
        status_col_txt = '';
        break;  
      case "customfield_10008":   
        epic_col = m;
        epic_col_txt = '';
        break;  
      default:   
        //do nothing				
    }
  }
  
  var missingColumnsHTML = id_col_txt + type_col_txt + displayText_col_txt + itemLabel_col_txt + sprint_col_txt + sprintStartDate_col_txt + status_col_txt + epic_col_txt;
  if ( missingColumnsHTML.length == 0 ) {
    var outHTML = '<div class="warning">' +
                    '<h2>The information in your sheet is incomplete. Please add these columns:</h2>' +
                    '<ul>' +
                      missingColumnsHTML +
                    '</ul>' +
                  '</div>';
  } else {
    var htmlParked = '';
    var htmlIdeas = '';
    var htmlPlanned = '';
    var htmlOngoing = '';
    var htmlDone = '';
  
    var itemHasSprint = false;
    var itemIsDone = false;
    
    for (var i=1; i<lastRow; i++) {  
      var id = allValues[i][0];
      var type = allValues[i][1];
      var displayText = cleanOutput(allValues[i][2]);
      var itemLabel = cleanOutput(allValues[i][3]);
      var sprint = cleanOutput(allValues[i][4]);
      var sprintStartDate = allValues[i][5];
      var status = allValues[i][6];
      var epic = allValues[i][7];
    
      if (sprint && sprintStartDate != "<null>"){  
        htmlOngoing += createItemHMTL(id, type, displayText, epic, "Ongoing" )
      } else if (itemLabel.indexOf("Parked") != -1) {
        htmlParked += createItemHMTL(id, type, displayText, epic, "Parked" )
      } else if (itemLabel.indexOf("Planned") != -1) {
        htmlPlanned += createItemHMTL(id, type, displayText, epic, "Planned" )
      } else {
        htmlIdeas += createItemHMTL(id, type, displayText, epic, "Idea" ) 
      }
    }
  
    Logger.log('html items completed');
  
    var outHTML = '<div class="column" id="Parked-wrapper" style="display: none;">' +
  	 	            '<h2>Parked</h2>' +
		            '<div class="item-container sortable-area" id="Parked">' +
		              htmlParked +
		            '</div>' +
                  '</div>' +
                  '<div class="column" id="Idea-wrapper">' +
		            '<h2>Ideas</h2>' +
		            '<div class="item-container sortable-area" id="Idea">' +
		              htmlIdeas +
		            '</div>' +
                  '</div>' +  
                  '<div class="column" id="Planned-wrapper">' +
		            '<h2>Planned</h2>' +
		            '<div class="item-container sortable-area" id="Planned">' +
		              htmlPlanned +
		            '</div>' +
                  '</div>' +
                  '<div class="column" id="Ongoing-wrapper">' +
		            '<h2>Ongoing</h2>' +
		            '<div class="item-container" id="Ongoing">' +
		              htmlOngoing +
		            '</div>' +
                  '</div>';
  }  
  Logger.log('outHTML completed');
  return outHTML;
}

// ******************************************************************************************************
// Generate the HTML for an item for the display page
// ******************************************************************************************************
function createItemHMTL(id, type, displayText, epic, column ){
   var itemHTML = '<li id="' + id + '" data-column="' + column + '">' +
                     '<span class="drag-handle">☰</span>' +
                     '<div class="type-' + type + '">' + displayText + '</div>' +
                     '<span class="epic epic-'+ epic +'"></span>' +
                     '<a class="link-item" href="https://jobladder.atlassian.net/browse/' + id + '" target="_blank"><i class="material-icons">launch</i></a>' +  
                   '</li>'; 
   return itemHTML;
}


// ******************************************************************************************************
// Clean the use input into something that we can embed in HTML without breaking it - from https://stackoverflow.com/questions/1787322/htmlspecialchars-equivalent-in-javascript/4835406#4835406
// ******************************************************************************************************
function cleanOutput(text) {
  var map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };

  return text.replace(/[&<>"']/g, function(m) { return map[m]; });
}