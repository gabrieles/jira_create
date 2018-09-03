// ---------------------------------------------------------------------------------------------------------------------------------------------------
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

//code extensively modified by Gabriele Sani - gabrielesani@gmail.com


var C_MAX_RESULTS = 1000;


function jiraConfigure() {
  
  //var host = Browser.inputBox("Enter the host name of your on demand instance e.g. toothCamp.atlassian.net", "Host", Browser.Buttons.OK);
  PropertiesService.getUserProperties().setProperty("host", host);
  
  var userID = Browser.inputBox("Enter your Jira UserID/email", "yourname@email.com", Browser.Buttons.OK_CANCEL);
  var userPassword = Browser.inputBox("Enter your Jira Password (Note: This will be base64 Encoded and saved as a property on the spreadsheet)", "Password", Browser.Buttons.OK_CANCEL);
  var userAndPassword = userID + ':' + userPassword;
  var x = Utilities.base64Encode(userAndPassword);
  PropertiesService.getUserProperties().setProperty("digest", "Basic " + x);
  
  var projectKey = Browser.inputBox("Enter your Jira Board Key (it will be an acronym, like CJ or SM)", "ProjectKey", Browser.Buttons.OK_CANCEL);
  PropertiesService.getUserProperties().setProperty("projectKey", projectKey);
  
  Browser.msgBox("Jira configuration saved successfully.");
}  



function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  Browser.msgBox("Spreadsheet will no longer refresh automatically.");
  
}  



function scheduleRefresh() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  ScriptApp.newTrigger("jiraPull").timeBased().everyHours(4).create();
  
  Browser.msgBox("Spreadsheet will refresh automatically every 4 hours.");

}  



function jiraPullManual() {
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getSheetName();
  jiraPull(sheetName);
  Browser.msgBox("Jira backlog successfully imported");
}  



function getFields() {
  return JSON.parse(getDataForAPI("field"));  
}  



function getStories() {
  var allData = {issues:[]};
  var data = {startAt:0,maxResults:0,total:1};
  var startAt = 0;
  
  var searchKeyword = PropertiesService.getUserProperties().getProperty("searchKeyword");
  if (searchKeyword) {
    //make sure that this will not break the GET call
    var searchFilter = "%20and%20text%20~%20" + '%22' + encodeURI(searchKeyword) + '%22';;
  } else {
    var searchFilter = "";  
  }
  Logger.log('searchFilter: ' + searchFilter);
  
  while (data.startAt + data.maxResults < data.total) {
    Logger.log("Making request for %s entries", C_MAX_RESULTS);  
    //data =  JSON.parse(getDataForAPI("search?jql=project%20%3D%20" + PropertiesService.getUserProperties().getProperty("prefix") + searchFilter + "%20and%20type%20in%20("+ PropertiesService.getUserProperties().getProperty("issueTypes") + ")%20order%20by%20created%20&maxResults=" + C_MAX_RESULTS + "&startAt=" + startAt));  
    //data =  JSON.parse(getDataForAPI("search?jql=project%20%3D%20" + PropertiesService.getUserProperties().getProperty("prefix") + searchFilter + "%20and%20type%20in%20(story,task,bug)%20order%20by%20created%20&maxResults=" + C_MAX_RESULTS + "&startAt=" + startAt));  
    
    //removed filter on issue type
    data =  JSON.parse(getDataForAPI("search?jql=project%20%3D%20" + PropertiesService.getUserProperties().getProperty("prefix") + searchFilter + "%20order%20by%20created%20&maxResults=" + C_MAX_RESULTS + "&startAt=" + startAt));  
    
    allData.issues = allData.issues.concat(data.issues);
    startAt = data.startAt + data.maxResults;
  }  
  
  return allData;
}  



function printOptionsForEpic() {
  var allData = {issues:[]};
  var data = {startAt:0,maxResults:0,total:1};
  var startAt = 0;
  
  while (data.startAt + data.maxResults < data.total) {
    Logger.log("Making request for %s epics", C_MAX_RESULTS);  
    //data =  JSON.parse(getDataForAPI("search?jql=project%20%3D%20" + PropertiesService.getUserProperties().getProperty("prefix") + searchFilter + "%20and%20type%20in%20("+ PropertiesService.getUserProperties().getProperty("issueTypes") + ")%20order%20by%20created%20&maxResults=" + C_MAX_RESULTS + "&startAt=" + startAt));  
    //data =  JSON.parse(getDataForAPI("search?jql=project%20%3D%20" + PropertiesService.getUserProperties().getProperty("prefix") + searchFilter + "%20and%20type%20in%20(story,task,bug)%20order%20by%20created%20&maxResults=" + C_MAX_RESULTS + "&startAt=" + startAt));  
    
    //removed filter on issue type
    data =  JSON.parse(getDataForAPI("search?jql=project%20%3D%20" + PropertiesService.getUserProperties().getProperty("projectKey") + "%20and%20type%20in%20(epic)%20order%20by%20created%20&maxResults=" + C_MAX_RESULTS + "&startAt=" + startAt));  
    
    allData.issues = allData.issues.concat(data.issues);
    startAt = data.startAt + data.maxResults;
  }  
  
  var html = '';
  for (i=0;i<data.issues.length;i++) {
    html = html + '<option value="'+ data.issues[i].key + '">' + data.issues[i].fields.customfield_10009 + '</options>';
  }  
  return html;
  
}  


function getDataForAPI(path) {
   var url = "https://" + PropertiesService.getUserProperties().getProperty("host") + "/rest/api/2/" + path;
   var digestfull = PropertiesService.getUserProperties().getProperty("digest");
  
  var headers = { "Accept":"application/json", 
                  "Content-Type":"application/json", 
                  "method": "GET",
                  "headers": {"Authorization": digestfull},
                  "muteHttpExceptions": true              
                 };
  
  var resp = UrlFetchApp.fetch(url,headers );
  if (resp.getResponseCode() != 200) {
    Browser.msgBox("Error retrieving data for url " + url + ":" + resp.getContentText());
    return "";
  }  
  else {
    return resp.getContentText();
  }  
  
}  


function sendMetaToLogger(){
  
  var urlAdd = "/rest/api/2/issue/createmeta?projectKeys=" + PropertiesService.getUserProperties().getProperty("projectKey") + "&expand=projects.issuetypes.fields";
  var url = "https://" + PropertiesService.getUserProperties().getProperty("host") + urlAdd ;
  var digestfull = PropertiesService.getUserProperties().getProperty("digest");
  
  var headers = { "Accept":"application/json", 
              "Content-Type":"application/json", 
              "method": "GET",
               "headers": {"Authorization": digestfull},
                 "muteHttpExceptions": true              
             };
  
  var resp = UrlFetchApp.fetch(url,headers );
  if (resp.getResponseCode() != 200) {
    Browser.msgBox("Error retrieving data for url " + url + ":" + resp.getContentText());
    return "";
  }  
  else {
    Logger.log(resp.getContentText());
  }  
  
}  



function jiraPull(sheetName) {
  
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
  
  
  
  //if you are on the instructins page, ask for the prefix
  if (prefix == "Instructions") {
    prefix = Browser.inputBox("Enter the 2 to 4 digit prefix for your Jira Project. e.g. CJ, SM or COUR", "Prefix", Browser.Buttons.OK);
    prefix = prefix.toUpperCase();
    var yourNewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(prefix);
    //if there is no sheet, create it
    if (yourNewSheet == null) {
      yourNewSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
      yourNewSheet.setName(prefix);
      var defaultHeadingsArray = ['Key','Status.name','Issue Type.name','Summary','Description','Story Points','Created','Updated'];
      yourNewSheet.getRange(1, 1, 1, defaultHeadingsArray.length).setValues(defaultHeadingsArray);
    }
  }
  PropertiesService.getUserProperties().setProperty("prefix", prefix.toUpperCase());
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  var allFields = getAllFields();
  
  var data = getStories();
  
  if (allFields === "" || data === "") {
    Browser.msgBox("Error pulling data from Jira - aborting now.");
    return;
  }  

  var headings = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];
  var extractArray = new Array(headings.length);
  
  //check if there is a "." in a headings name
  for (ii=0;ii<headings.length;ii++) {
    if (headings[ii].indexOf(".")) {
       var dummy = headings[ii].split(".");
       headings[ii] = dummy[0];
       extractArray[ii] = dummy[1];
    } else {
      extractArray = "";
    }
  }
  
  var y = new Array();
  for (i=0;i<data.issues.length;i++) {
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
  for (var j=1;j<headings.length+1;j++) {
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
  var regexDate = /20\d{2}(-|\/)((0[1-9])|(1[0-2]))(-|\/)((0[1-9])|([1-2][0-9])|(3[0-1]))(T|\s)(([0-1][0-9])|(2[0-3])):([0-5][0-9]):([0-5][0-9]).([0-9][0-9][0-9]\+[0-9][0-9][0-9][0-9])/
  var testValues = ss.getRange(2, 1, 2, ss.getLastColumn()).getValues()[0];
  
  for (var i=1;i<testValues.length+1;i++) {

    dummyVal = testValues[i-1];
    if (dummyVal.length == 28) {
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
  
}



function getAllFields() {
  
  var theFields = getFields();
  var allFields = new Object();
  allFields.ids = new Array();
  allFields.names = new Array();
  
  for (var i = 0; i < theFields.length; i++) {
      allFields.ids.push(theFields[i].id);
      allFields.names.push(theFields[i].name.toLowerCase());
  }  
  
  return allFields;
  
}  



function getStory(data,headings,fields) {
  
  var story = [];
  for (var i = 0;i < headings.length;i++) {
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
  var url = "https://" + PropertiesService.getUserProperties().getProperty("host") + "/rest/agile/1.0/issue/rank";
  var digestfull = PropertiesService.getUserProperties().getProperty("digest");
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
}  


function editJiraLabel(IssueID,addLabelValue,removeLabelValue) {
  var url = "https://" + PropertiesService.getUserProperties().getProperty("host") + "/rest/api/2/issue/" + IssueID;
  var digestfull = PropertiesService.getUserProperties().getProperty("digest");
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
} 