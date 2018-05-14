var sheetID = "19LhOLxFuOVHx2bE5BrZQVdUDK5Re3A0T4U6Fe6c0KO8"; //needed as you cannot use getActiveSheet() while the sheet is not in use (as in a standalone application like this one)
var favicon_url = 'http://icons.iconarchive.com/icons/iconsmind/outline/32/Quill-3-icon.png';


// ******************************************************************************************************
// Function to create menus when you open the sheet
// ******************************************************************************************************
function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{name: "Configure Jira", functionName: "jiraConfigure"},
                     {name: "Refresh Data Now", functionName: "jiraPullManual"},
                     {name: "Schedule 4 Hourly Automatic Refresh", functionName: "scheduleRefresh"},
                     {name: "Stop Automatic Refresh", functionName: "removeTriggers"},
                     {name: "Clean data", functionName: "cleanData"},
                     {name: "Log the card creation metadata", functionName: "sendMetaToLogger"}]; 
  ss.addMenu("Jira", menuEntries);
}


// ******************************************************************************************************
// Function to display the HTML as a webApp
// ******************************************************************************************************
function doGet(e) {
   var template = HtmlService.createTemplateFromFile('submit');
   var htmlOutput = template.evaluate()
                   .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                   .setTitle('JIRA Create')
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