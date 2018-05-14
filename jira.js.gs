var host = "jobladder.atlassian.net"; //where your JIRA is

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
  
  //var issueTypes = Browser.inputBox("Enter a comma separated list of the types of issues you want to import  e.g. story or story,epic,bug", "Issue Types", Browser.Buttons.OK);
  //PropertiesService.getUserProperties().setProperty("issueTypes", issueTypes);

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
  jiraPull();
  cleanData();
  Browser.msgBox("Jira backlog successfully imported");
}  



function getFields() {
  return JSON.parse(getDataForAPI("field"));  
}  



function getStories() {
  var allData = {issues:[]};
  var data = {startAt:0,maxResults:0,total:1};
  var startAt = 0;
  var searchFilter = "";
  var searchKeyword = PropertiesService.getUserProperties().getProperty("searchKeyword");
  if (searchKeyword) {
    //make sure that this will not break the GET call
    searchKeyword = '%22' + encodeURI(searchKeyword) + '%22';
    searchFilter = "%20and%20text%20~%20" + searchKeyword;
  }
  
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
    Logger.log("Making request for %s entries", C_MAX_RESULTS);  
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



function jiraPull() {
  
  //get the prefix and search keyword from the sheet name
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getSheetName();
  var prefix = sheetName;
  
  if (prefix.indexOf("-") >0) {
    var dummy = prefix.split("-");
    prefix = dummy[0];
    var searchKeyword = dummy[1];
    PropertiesService.getUserProperties().setProperty("searchKeyword",searchKeyword);
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
   
}




function cleanData() {
  
  var prefix = SpreadsheetApp.getActiveSpreadsheet().getSheetName();
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(prefix);
  var last = ss.getLastRow();
  
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
  
  //import the moment.js library to convert dates - deprectaed for copying in the code into a local file
  //eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.20.1/moment.min.js').getContentText());
  
  
  //clean up the output
  var dummyVal;
  var regexDate = /20\d{2}(-|\/)((0[1-9])|(1[0-2]))(-|\/)((0[1-9])|([1-2][0-9])|(3[0-1]))(T|\s)(([0-1][0-9])|(2[0-3])):([0-5][0-9]):([0-5][0-9]).([0-9][0-9][0-9]\+[0-9][0-9][0-9][0-9])/
  for (jj=0;jj<extractArray.length+1;jj++) {
    
    if (extractArray[jj]) {
      for (kk=2;kk<last+1;kk++) {
        dummyVal = ss.getRange(kk,jj+1).getValue();
        var startExt = dummyVal.indexOf(extractArray[jj]);
        if (startExt > 0 ) {
          var endExt = dummyVal.indexOf(',',startExt);     
          ss.getRange(kk,jj+1).setValue(dummyVal.slice(startExt + extractArray[jj].length + 1,endExt));
        }
      }
    }
    
    //check if a column has dates, and trim it
    dummyVal = ss.getRange(2,jj+1).getValue();
    if (dummyVal.length == 28) {
      if (dummyVal.match(regexDate)) {
        for (kkk=2;kkk<last+1;kkk++) {
          //commented out since I do not see the need to use moment.js at this stage
          //dummyVal = ss.getRange(kkk,jj+1).getValue().substring(0,19);
          //ss.getRange(kk,jj+1).setValue( moment(dummyVal).format('YYYY/MM/DD') );
          dummyVal = ss.getRange(kkk,jj+1).getValue().substring(0,10);
          ss.getRange(kkk,jj+1).setValue( dummyVal );
        }
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
        