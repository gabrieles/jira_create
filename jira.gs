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



// ******************************************************************************************************
// GET a JIRA endpoint and return its response if there is no error, otherwise nothing
// ******************************************************************************************************
function getDataFromJIRA(path) {
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
    Logger.log("Error " + resp.getResponseCode() + " for url " + url + " : " + resp.getContentText());
    return "";
  }  
  else {
    return resp.getContentText();
  }  
  
}  



// ******************************************************************************************************
// POST payload to a JIRA endpoint. Return its response if there is no error, otherwise nothing
// ******************************************************************************************************
function postDatatoJIRA(path, payload){
   var url = "https://" + printVal("host") + path;
    var authVal = printVal("digest");  
    var options = { 
                "Accept":"application/json", 
                "contentType":"application/json", 
                "method": "POST",
                "payload": payload,
                "headers": {"Authorization": authVal}           
               };
  
    var resp = UrlFetchApp.fetch(url,options);
  
    //if it all goes well the response should be 201 - the JIRA documentation is incorrect. See https://jira.atlassian.com/browse/JRASERVER-39339
    if (resp.getResponseCode() != 201) {
       Logger.log("Error " + resp.getResponseCode() + " for url " + url + " : " + resp.getContentText());
       return ;
    }  
    else {
      return resp.getContentText();
    }
}
  


// ******************************************************************************************************
// PUT payload to a JIRA endpoint. Return error message, otherwise nothing
// ******************************************************************************************************
function putDataOnJIRA(path,payload){
 
  var url = "https://" + printVal("host") + path;
  var digestfull = printVal("digest");
  var headers = { "method":"PUT",
                  "contentType":"application/json",
                  "headers":{"Authorization": digestfull},
                  "payload": payload,
                  "muteHttpExceptions":true
                };
  
  var resp = UrlFetchApp.fetch(url,headers);
  if (resp.getResponseCode() != 204) {
    var outText = 'Error ' + resp.getResponseCode() + ' - ' + resp.getContentText();
    return outText;
  }  
  else {
    return;
  }  
  
}  



// ******************************************************************************************************
// Create a new item in JIRA
// ******************************************************************************************************
function JIRAsubmit(data) {
  if (userIsEditor()){
    
    var response = postDatatoJIRA("/rest/api/2/issue/", data)
      
  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert('You do not have the permission to create a new item via this interface'); 
  }
}



// ******************************************************************************************************
// Call JIRA and get a list of all fields
// ******************************************************************************************************
function getAllFields() {
  var theFields = JSON.parse(getDataFromJIRA("field")); 
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



// ******************************************************************************************************
// Call JIRA to get all issues that are not epics and do not have status "done", ordered by rank (=customfield_10400). Optionally, pass a search keyword, too
// ******************************************************************************************************
function getIssues(searchKeyword) {
  var allData = {issues:[]};
  var data = {startAt:0,maxResults:0,total:1};
  var startAt = 0;
  
  if (searchKeyword) {
    //encode the searchKeyword to make sure that it will not break the GET call
    var searchFilter = "%20and%20text%20~%20" + '%22' + encodeURI(searchKeyword) + '%22';;
  } else {
    var searchFilter = "";  
  }
  
  while (data.startAt + data.maxResults < data.total) {
    var C_MAX_RESULTS = printVal("maxNumResults")
    
    //exclude epics, and items with status=done
    var searchQuery = 'search?jql=project%20%3D%20' + printVal("projectKey") + 
                                  searchFilter + '%20and%20' + 
                                  'issuetype%20in%20(story,task,bug,hotfix,support)' + 
                                  '%20and%20' + 'status!=Done%20' + 
                                  'order%20by%20cf[10400]%20asc%20&' +
                                  'maxResults=' + C_MAX_RESULTS + 
                                  '&expand=versionedRepresentations&startAt=' + startAt;  
    Logger.log(searchQuery);
    data =  JSON.parse(getDataFromJIRA(searchQuery));  
    
    allData.issues = allData.issues.concat(data.issues);
    startAt = data.startAt + data.maxResults;
  }  
  
  return allData;
}  



// ******************************************************************************************************
// Call JIRA and get a list of epics where epicStatus (=customfield_10010) is not "Done"
// ******************************************************************************************************
function getEpics(){
  var allData = {issues:[]};
  var data = {startAt:0,maxResults:0,total:1};
  var startAt = 0;
  
  while (data.startAt + data.maxResults < data.total) {
    var C_MAX_RESULTS = printVal("maxNumResults")
    var searchQuery = 'search?jql=project%20%3D%20' + printVal('projectKey') + 
                      '%20and%20' +  'type%20in%20(epic)' + 
                      '%20and%20' +  'cf[10010]!=Done%20' + 
                      '%20order%20by%20summary%20ASC%20' +
                      '&maxResults=' + C_MAX_RESULTS + 
                      '&startAt=' + startAt;
    Logger.log(searchQuery)
    data =  JSON.parse(getDataFromJIRA(searchQuery));  
    
    allData.issues = allData.issues.concat(data.issues);
    startAt = data.startAt + data.maxResults;
  } 
  return data;
}



// ******************************************************************************************************
// Wrapper to get all issues out of JIRA
// ******************************************************************************************************
function jiraPull(sheetName) {
  
  var allFields = getAllFields();
  if (allFields === "" || data === "") {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Error pulling data from Jira - aborting now.");
    return;
  }  
   
  if(sheetName){
    //get the prefix and search keyword from the sheet name  
    if (sheetName.indexOf("-") >0) {
      var dummy = sheetName.split("-");
      var prefix = dummy[0];
      var searchKeyword = dummy[1];
      var data = getIssues(searchKeyword);
    } else {
      var prefix = sheetName;
      var data = getIssues();
    }   
    
    printDataToSheet(data, sheetName);   
    
  } else {
    var data = getIssues();
    return data;    
  }
  
}



// ******************************************************************************************************
// Change rank of an item in JIRA
// This is not an absolute value, but a reference to the next item in the list...
// see https://docs.atlassian.com/jira-software/REST/7.3.1/#agile/1.0/issue-rankIssues for details
// ******************************************************************************************************
function updateRank(IssueMovedID,BeforeIssueID) {
  if (userIsEditor()){
    var path = "/rest/agile/1.0/issue/rank";
    var payload = '{"issues":["' + IssueMovedID + '"],"rankBeforeIssue":"' + BeforeIssueID + '"}';
    var response = putDataOnJIRA(path,payload)
    
    if (response) {
      Logger.log(response);
    } else {
      Logger.log("Success: moved item " + IssueMovedID + " above item " + BeforeIssueID);
    }    
  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert('You do not have the permission to update the ranking');
  }
}  



// ******************************************************************************************************
// Add (and optionally Remove) a label for an issue in JIRA
// ******************************************************************************************************
function editJiraLabel(IssueID,addLabelValue,removeLabelValue) {
  if (userIsEditor()){
    var path = "/rest/api/2/issue/" + IssueID;
    if (removeLabelValue) {
      var payload = '{ "update": { "labels": [ {"add" : "' + addLabelValue + '"},{"remove" : "' + removeLabelValue + '"}  ] } }';
    } else {
      var payload = '{ "update": { "labels": [ {"add" : "' + addLabelValue + '"} ] } }';
    }
  
    var response = putDataOnJIRA(path,payload)
    if (response) {
      Logger.log(response + 'for PUT in editJiraLabel using IssueID: ' + IssueID + ', addLabelValue: ' + addLabelValue + ' and removeLabelValue:' + removeLabelValue);
      return "";
    } else {
      Logger.log("Success: updated label for " + IssueID + ". Added " + addLabelValue + " and removed " + removeLabelValue);
    }    
  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert('You do not have the permission to move the items');
  }
} 



// ******************************************************************************************************
// Call JIRA to discover the meta-data used for creating issues
// See https://developer.atlassian.com/server/jira/platform/jira-rest-api-example-discovering-meta-data-for-creating-issues-6291669/
// ******************************************************************************************************
function sendIssueMetaToLogger(){
  
  var path = "/rest/api/2/issue/createmeta?projectKeys=" + printVal("projectKey") + "&expand=projects.issuetypes.fields";
  var response = getDataFromJIRA(path); 

  if (response){
    Logger.log(response);
  }  
  
}  

