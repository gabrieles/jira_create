// ---------------------------------------------------------------------------------------------------------------------------------------------
// The MIT License (MIT)
// 
// Copyright (c) 2019 https://github.com/gabrieles
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.





// NB: before you can use this you need to publish the project via "Publish > Deploy as web app"
// Once you do, the system will show you the URL of your web app. Visiting this url triggers the response defined in the doGet() function 



// ******************************************************************************************************
// Function to create menus when a user opens the sheet
// ******************************************************************************************************
function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
                     {name: "Configure Jira", functionName: "jiraConfigure"},
                     {name: "Configure Editors", functionName: "editorsConfigure"},
                     {name: "Refresh data on this sheet", functionName: "jiraPullOnSheet"}
                    ]; 
  ss.addMenu("Jira", menuEntries);
}



// ******************************************************************************************************
// Core function to trigger the response when a user visits the web app 
// ******************************************************************************************************
function doGet(e) {
  
  //you can pass a parameter via the URL as ?action=XXX to serve different pages
  var userAction = e.parameter.action;
  
  switch(userAction) {
     case "create":  
       var template = HtmlService.createTemplateFromFile('submit'); 
       var pageTitle = "JIRA Create";
     break;      
     default:
       var template = HtmlService.createTemplateFromFile('roadmap');
       var pageTitle = "CJ Roadmap";
  }
  var favicon_url = 'https://res.cloudinary.com/gabrieles/image/upload/v1536163902/quill.png';
  var htmlOutput = template.evaluate()
                   .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                   .setTitle(pageTitle)
                   .addMetaTag('viewport', 'width=device-width, initial-scale=1')
                   .setFaviconUrl(favicon_url);

  return htmlOutput;
};



// ******************************************************************************************************
// Print out the content of file in this project (used for HTML, CSS and JS)
// ******************************************************************************************************
function getContent(filename) {
  var pageContent = HtmlService.createTemplateFromFile(filename).getRawContent();
  return pageContent;
}



// ******************************************************************************************************
// Shortcut to retrieve the value of a property
// ******************************************************************************************************
function printVal(key) {
  if (key == 'sheetID' || key == 'host' || key == 'projectKey' || key == 'editorEmails') {
    var propVal = PropertiesService.getScriptProperties().getProperty(key);
  } else {
    var propVal = PropertiesService.getUserProperties().getProperty(key);
  }
  return propVal;
}



// ******************************************************************************************************
// Store the configuration used to connect to JIRA
// ******************************************************************************************************
function jiraConfigure() {
  //only the owner can access this
  var userEmail = Session.getActiveUser().getEmail();
  var sheetID = SpreadsheetApp.getActiveSpreadsheet().getId();
  var file = DriveApp.getFileById(sheetID);
  if (file.getOwner().getEmail() == userEmail ){
    
    //store the sheetID automatically
    PropertiesService.getScriptProperties().setProperty("sheetID", sheetID);
    
    var ui = SpreadsheetApp.getUi();
    
    var host = askToSetProperty(ui,"Enter the host name of your JIRA instance","Eg: jobladder.atlassian.net (Cancel to skip)", "host")
    
    var userID = askToSetProperty(ui,"Enter your Jira UserID/email","Eg: yourname@charityjob.co.uk (Cancel to skip)", "userID")
    
    var apiToken = askToSetProperty(ui,"Enter your Jira API Token","Get it at https://id.atlassian.com/manage/api-tokens (Cancel to skip)", "apiToken")
    
    if( userID && apiToken){
      var userAndToken = userID + ':' + apiToken;
      var x = Utilities.base64Encode(userAndToken);
      PropertiesService.getUserProperties().setProperty("digest", "Basic " + x);
    }
    
    var projectKey = askToSetProperty(ui,"Enter your Jira Board Key", "It will be something short like 'CJ' or 'Didi' (Cancel to skip)", "projectKey");
    
    var maxNumResults = askToSetProperty(ui,"Enter the max number of issues to retrieve", "1000 (Cancel to skip)", "maxNumResults");
    
    ui.alert("Jira configuration saved successfully.");
  } else {
    ui.alert("Only the file owner can complete this action");
  }
}  



// ******************************************************************************************************
// Shortcut to generate a modal to store a property value
// ******************************************************************************************************
function askToSetProperty(ui,title,prompt,propName){
  
  var result = ui.prompt(title, prompt, ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    if(propName == 'host' || propName == 'projectKey' || propName == 'editorEmails' ){
      PropertiesService.getScriptProperties().setProperty(propName, text);
    } else {
      PropertiesService.getUserProperties().setProperty(propName, text);
    }
    return text;
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    //do nothing
    return ;
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    //do nothing
    return ;
  }
}



// ******************************************************************************************************
// Set editors - only the file owner can run this
// ******************************************************************************************************
function editorsConfigure(){
  var userEmail = Session.getActiveUser().getEmail();
  var ui = SpreadsheetApp.getUi();
  if (file.getOwner().getEmail() == userEmail ){
    var editorEmails = printVal("editorEmails");
    var dummy = askToSetProperty(ui,"Add or remove emails to determine who can move items", "(the owner will always have full access)", "editorEmails");
    if (dummy != editorEmails){
      PropertiesService.getUserProperties().setProperty("editorEmails", dummy);
    }
  } else {
    ui.alert('You do not have permission to do this. Only the file owner can');
  }
}



// ******************************************************************************************************
// Check that the user has the editor role
// ******************************************************************************************************
function userIsEditor(){
  var userEmail = Session.getActiveUser().getEmail();
  var editorEmails = printVal("editorEmails");
  if(!editorEmails){
    return true;
  } else {  
    if ( editorEmails.indexOf(userEmail) != -1 ) {
      return true;
    } else {  
      var sheetID = printVal('sheetID')
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
}




// ******************************************************************************************************
// Return the URL of this web app. Optionally, add a string at the end of it
// ******************************************************************************************************
function getScriptURL(addpath) {
  var url = ScriptApp.getService().getUrl();
  if (addpath) { url += addpath; }
  return url;
}



// ******************************************************************************************************
// Call JiraPull to print data on the current sheet
// ******************************************************************************************************
function jiraPullOnSheet() {
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getSheetName();
  jiraPull(sheetName);
  var ui = SpreadsheetApp.getUi();
  ui.alert("Jira backlog successfully imported");
}  



// ******************************************************************************************************
// Print data on the current sheet
// ******************************************************************************************************
function printDataToSheet(data,sheetName){

    //if you are on the instructions page, ask for the prefix
    if (prefix == "Instructions") {
      var ui = SpreadsheetApp.getUi();
      ui.alert("You cannot do it on this sheet")
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);   
    var headings = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];
    var extractArray = new Array(headings.length);
    
    //check if there is a "." in a headings name
    var hLength = headings.length;
    for (var ii=0; ii<hLength; ii++) {
      if (headings[ii].indexOf(".")) {
        var dummy = headings[ii].split(".");
        headings[ii] = dummy[0];
        extractArray[ii] = dummy[1];
      } else {
        extractArray = "";
      }
    }
    
    var y = new Array();
    var diLength = data.issues.length;
    for (var i=0; i<diLength; i++) {
      var d=data.issues[i];
      y.push(getIssue(d,headings,allFields));
    }  
    
    var last = ss.getLastRow();
    if (last >= 2) {
      ss.getRange(2, 1, ss.getLastRow()-1,ss.getLastColumn()).clearContent();  
    }  
    
    if (y.length > 0) {
      ss.getRange(2, 1, data.issues.length,y[0].length).setValues(y);
    }
    
    cleanData(sheetName);
}



// ******************************************************************************************************
// Clean up data stored in a sheet
// ******************************************************************************************************
function cleanData(sheetName) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var last = ss.getLastRow();
  
  var headings = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];
    
  //clean up the output of columns with a "."
  var hLength = headings.length +1;
  for (var j=1; j<hLength; j++) {
    var headingVal = headings[j-1]; 
    if (headingVal.indexOf(".") != -1 ) {
      Logger.log(headingVal);
      var range = ss.getRange(2,j,last,j);
      var values = range.getValues();
      var dummyVal = headingVal.split(".")[1];
      for (var k=0; k<last; k++) {
        if(values[k][0].indexOf(dummyVal) != -1 ) {
          var startExt = values[k][0].indexOf(dummyVal);
          if (startExt > 0 ) {
            var endExt = values[k][0].indexOf(',',startExt);     
            values[k][0] = values[k][0].slice(startExt + dummyVal.length + 1,endExt);
          }
        }  
      }
      range.setValues(values);
    }
  }

  //check if a column has dates in the JIRA format, and if so trim the values  
  //var regexDate = /20\d{2}(-|\/)((0[1-9])|(1[0-2]))(-|\/)((0[1-9])|([1-2][0-9])|(3[0-1]))(T|\s)(([0-1][0-9])|(2[0-3])):([0-5][0-9]):([0-5][0-9]).([0-9][0-9][0-9]\+[0-9][0-9][0-9][0-9])/
  var regexDate = /20\d{2}(-|\/)((0[1-9])|(1[0-2]))(-|\/)((0[1-9])|([1-2][0-9])|(3[0-1]))/
  var testValues = ss.getRange(2, 1, 2, ss.getLastColumn()).getValues()[0];
  var tLength = testValues.length+1;
  for (var i=1; i<tLength; i++) {

    dummyVal = testValues[i-1].slice(0, 10);
    
    if (dummyVal.match(regexDate)) {
      var dateRange = ss.getRange(2,i,last,i);
      var dateValues = dateRange.getValues();
      for (var h=0; h<last; h++) {
        dateValues[h][0] = dateValues[h][0].substring(0,10);
      }
      dateRange.setValues(dateValues);       
    }
  }    
}



// ******************************************************************************************************
// Return an issue as a new array of data based on the headings
// ******************************************************************************************************
function getIssue(data,headings,fields) {
  
  var issue = [];
  var hLength = headings.length;
  for (var i=0; i<hLength; i++) {
    if (headings[i] !== "") {
      issue.push(getDataForHeading(data,headings[i].toLowerCase(),fields));
    }  
  }        
  
  return issue;
  
}  



// ******************************************************************************************************
// Return a value based on the heading/key 
// ******************************************************************************************************
function getDataForHeading(data,heading,fields) {
  
      if (data.hasOwnProperty(heading)) {
        return data[heading];
      }  
      else if (data.fields.hasOwnProperty(heading)) {
        return data.fields[heading];
      }  
  
      var fieldName = getFieldName(heading,fields);
  
      if (fieldName !== "") {
        if (data.hasOwnProperty(fieldName)) {
          return data[fieldName];
        }  
        else if (data.fields.hasOwnProperty(fieldName)) {
          return data.fields[fieldName];
        }  
      }
  
      var splitName = heading.split(" ");
  
      if (splitName.length == 2) {
        if (data.fields.hasOwnProperty(splitName[0]) ) {
          if (data.fields[splitName[0]] && data.fields[splitName[0]].hasOwnProperty(splitName[1])) {
            return data.fields[splitName[0]][splitName[1]];
          }
          return "";
        }  
      }  
  
      return "Could not find value for " + heading;
      
}  



// ******************************************************************************************************
// Get the name/key of a field
// ******************************************************************************************************
function getFieldName(heading,fields) {
  var index = fields.names.indexOf(heading);
  if ( index > -1) {
     return fields.ids[index]; 
  }
  return "";
}  
