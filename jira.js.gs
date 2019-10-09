// ---------------------------------------------------------------------------------------------------------------------------------------------
// The MIT License (MIT)
// 
// Copyright (c) 2014 Iain Brown - http://www.littlebluemonkey.com/blog/automatically-import-jira-backlog-into-google-spreadsheet
//
// Inspired by http://gmailblog.blogspot.co.nz/2011/07/gmail-snooze-with-apps-script.html
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

//code extensively modified by https://github.com/gabrieles


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



function jiraPullOnSheet() {
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getSheetName();
  jiraPull(sheetName);
  var ui = SpreadsheetApp.getUi();
  ui.alert("Jira backlog successfully imported");
}  





function getAllFields() {
  var theFields = JSON.parse(getDataForAPI("field")); 
  var allFields = new Object();
  allFields.ids = new Array();
  allFields.names = new Array();
  var fLength = theFields.length;
  for (var i=0; i<fLength; i++) {
      allFields.ids.push(theFields[i].id);
      allFields.names.push(theFields[i].name.toLowerCase());
  }  
  return allFields;
}  


function getStories() {
  var allData = {issues:[]};
  var data = {startAt:0,maxResults:0,total:1};
  var startAt = 0;
  
  var searchKeyword = printVal("searchKeyword");
  if (searchKeyword) {
    //make sure that this will not break the GET call
    var searchFilter = "%20and%20text%20~%20" + '%22' + encodeURI(searchKeyword) + '%22';;
  } else {
    var searchFilter = "";  
  }
  Logger.log('searchFilter: ' + searchFilter);
  
  while (data.startAt + data.maxResults < data.total) {
    var C_MAX_RESULTS = printVal("maxNumResults")
    Logger.log("Making request for %s entries", C_MAX_RESULTS);  
    
    //exclude epics, and items with status=done
    var searchQuery = 'search?jql=project%20%3D%20' + printVal("projectKey") + 
                                  searchFilter + '%20and%20' + 
                                  'issuetype%20in%20(story,task,bug,hotfix,support)%20and%20' + 
                                  'status!=Done%20' + 
                                  'order%20by%20priority%20desc%20&' +
                                  'maxResults=' + C_MAX_RESULTS + 
                                  '&expand=versionedRepresentations&startAt=' + startAt;  
    Logger.log(searchQuery);
    data =  JSON.parse(getDataForAPI(searchQuery));  
    
    allData.issues = allData.issues.concat(data.issues);
    startAt = data.startAt + data.maxResults;
  }  
  
  return allData;
}  


function getEpics(){
  var allData = {issues:[]};
  var data = {startAt:0,maxResults:0,total:1};
  var startAt = 0;
  
  while (data.startAt + data.maxResults < data.total) {
    var C_MAX_RESULTS = printVal("maxNumResults")
    Logger.log("Making request for %s epics", C_MAX_RESULTS);  
    
    data =  JSON.parse(getDataForAPI("search?jql=project%20%3D%20" + printVal("projectKey") + "%20and%20type%20in%20(epic)%20order%20by%20created%20&maxResults=" + C_MAX_RESULTS + "&startAt=" + startAt));  
    
    allData.issues = allData.issues.concat(data.issues);
    startAt = data.startAt + data.maxResults;
  } 
  return data;
}





function getDataForAPI(path) {
   var url = "https://" + printVal("host") + "/rest/api/2/" + path;
   var digestfull = printVal("digest");
  
  var parameters = { "Accept":"application/json", 
                  "Content-Type":"application/json", 
                  "method": "GET",
                  "headers": {"Authorization": digestfull},
                  "muteHttpExceptions": true              
                 };
  
  var resp = UrlFetchApp.fetch(url,parameters );
  if (resp.getResponseCode() != 200) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Error retrieving data for url " + url + ":" + resp.getContentText());
    return "";
  }  
  else {
    return resp.getContentText();
  }  
  
}  




function sendMetaToLogger(){
  
  var urlAdd = "/rest/api/2/issue/createmeta?projectKeys=" + printVal("projectKey") + "&expand=projects.issuetypes.fields";
  var url = "https://" + printVal("host") + urlAdd ;
  var digestfull = printVal("digest");
  
  var headers = { "Accept":"application/json", 
              "Content-Type":"application/json", 
              "method": "GET",
               "headers": {"Authorization": digestfull},
                 "muteHttpExceptions": true              
             };
  
  var resp = UrlFetchApp.fetch(url,headers );
  if (resp.getResponseCode() != 200) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Error retrieving data for url " + url + ":" + resp.getContentText());
    return "";
  }  
  else {
    Logger.log(resp.getContentText());
  }  
  
}  





function jiraPull(sheetName) {
  
  var allFields = getAllFields();
  if (allFields === "" || data === "") {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Error pulling data from Jira - aborting now.");
    return;
  }  
  
  var data = getStories();
    
  if(sheetName){
    printDataToSheet(data, sheetName);   
  } else {
    return data;    
  }
  
}


function printDataToSheet(data,sheetName){
   //get the prefix and search keyword from the sheet name  
    if (sheetName.indexOf("-") >0) {
      var dummy = sheetName.split("-");
      var prefix = dummy[0];
      var searchKeyword = dummy[1];
      PropertiesService.getUserProperties().setProperty("searchKeyword",searchKeyword);
    } else {
      var prefix = sheetName;
      PropertiesService.getUserProperties().setProperty("searchKeyword","");
    }
    
  
  
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
      y.push(getStory(d,headings,allFields));
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







function getStory(data,headings,fields) {
  
  var story = [];
  var hLength = headings.length;
  for (var i=0; i<hLength; i++) {
    if (headings[i] !== "") {
      story.push(getDataForHeading(data,headings[i].toLowerCase(),fields));
    }  
  }        
  
  return story;
  
}  



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



function getFieldName(heading,fields) {
  var index = fields.names.indexOf(heading);
  if ( index > -1) {
     return fields.ids[index]; 
  }
  return "";
}  


// see https://docs.atlassian.com/jira-software/REST/7.3.1/#agile/1.0/issue-rankIssues for details
function updateRank(IssueMovedID,BeforeIssueID) {
  if (userIsEditor()){
    var url = "https://" + printVal("host") + "/rest/agile/1.0/issue/rank";
    var digestfull = printVal("digest");
    var payload = '{"issues":["' + IssueMovedID + '"],"rankBeforeIssue":"' + BeforeIssueID + '"}';
    var headers = { "method":"PUT",
                    "contentType":"application/json",
                    "headers":{"Authorization": digestfull},
                    "payload": payload,
                    "muteHttpExceptions":true
                   };
  
    var resp = UrlFetchApp.fetch(url,headers);
    if (resp.getResponseCode() != 204) {
      Logger.log("Error " + resp.getResponseCode() + " updating rank for " + IssueMovedID + " to be placed before: " + BeforeIssueID + " Response Text: " + resp.getContentText());
      return "";
    } else {
      Logger.log("Success: moved item " + IssueMovedID + " above item " + BeforeIssueID);
    }    
  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert('You do not have the permission to update the ranking');
  }
}  


function editJiraLabel(IssueID,addLabelValue,removeLabelValue) {
  if (userIsEditor()){
    var url = "https://" + printVal("host") + "/rest/api/2/issue/" + IssueID;
    var digestfull = printVal("digest");
    if (removeLabelValue) {
      var payload = '{ "update": { "labels": [ {"add" : "' + addLabelValue + '"},{"remove" : "' + removeLabelValue + '"}  ] } }';
    } else {
      var payload = '{ "update": { "labels": [ {"add" : "' + addLabelValue + '"} ] } }';
    }
    var headers = { "method":"PUT",
                    "contentType":"application/json",
                    "headers":{"Authorization": digestfull},
                    "payload": payload,
                    "muteHttpExceptions":true
                   };
  
    var resp = UrlFetchApp.fetch(url,headers);
    if (resp.getResponseCode() != 204) {
      Logger.log("Error " + resp.getResponseCode() + " updating label for " + IssueID + " Response Text: " + resp.getContentText());
      return "";
    } else {
      Logger.log("Success: updated label for " + IssueID + ". Added " + addLabelValue + " and removed " + removeLabelValue);
    }    
  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert('You do not have the permission to move the items');
  }
} 