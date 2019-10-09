// ******************************************************************************************************
// Function to create menus when you open the sheet
// ******************************************************************************************************
function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{name: "Configure Jira", functionName: "jiraConfigure"},
                     {name: "Configure Editors", functionName: "editorsConfigure"},
                     {name: "Refresh data on this sheet", functionName: "jiraPullOnSheet"}
                    ]; 
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
       var template = HtmlService.createTemplateFromFile('roadmap');
       var pageTitle = "CJ Search & Match Roadmap";
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
  if (key == 'sheetID' || key == 'host' || key == 'projectKey' || key == 'editorEmails') {
    var propVal = PropertiesService.getScriptProperties().getProperty(key);
  } else {
    var propVal = PropertiesService.getUserProperties().getProperty(key);
  }
  return propVal;
}


// ******************************************************************************************************
// Function to create a new item in JIRA
// ******************************************************************************************************
function JIRAsubmit(data) {
  if (userIsEditor()){
    var url = "https://" + printVal("host") + "/rest/api/2/issue/";
    var authVal = printVal("digest");  
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
      var ui = SpreadsheetApp.getUi();
      ui.alert('Item ' + respJSON.self + ' has been created.'); 
    }  
  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert('You do not have the permission to create a new item via this interface'); 
  }
}






// ******************************************************************************************************
// Generate the HTML for the roadmap display page
// ******************************************************************************************************
function printColumnsWithItems(){
    
  var htmlOut = {};
  htmlOut.Parked = '';
  htmlOut.Idea = '';
  htmlOut.Planned = '';                     
  htmlOut.Ongoing = '';   
                     
  var data = jiraPull();
  Logger.log('data.issues.length: ' +data.issues.length)
  for (var i=0; i<data.issues.length; i++) {
    var issue=data.issues[i];
    Logger.log('Issue ' + issue.key);
    var sprintArray = issue.versionedRepresentations.customfield_10007[2]
    var sprintState = 'nosprint';
  
    if (sprintArray){
      var sprint = sprintArray[sprintArray.length-1]
      if (sprint){
        if(sprint.hasOwnProperty('state')){
          sprintState = sprint.state;
        }
      }  
    } 
  
    var issueLabels = issue.versionedRepresentations.labels[1]
   
    var outCol = "Idea" //default output column
    
    if (issueLabels){ //change output column if it has labels
      if (issueLabels.indexOf("Parked") != -1) {
        outCol = 'Parked';
      } else if (issueLabels.indexOf("Planned") != -1) {
        outCol = 'Planned';
      }
    }  
    if (sprintState == 'active'){
      outCol = 'Ongoing'     
    }
  
    var issueKey = issue.key;
    var issueType = issue.versionedRepresentations.issuetype[1].name;
    var issueTitle = issue.versionedRepresentations.summary[1];
    var epicId = issue.versionedRepresentations.customfield_10008[1];
    htmlOut[outCol] += createItemHMTL(issueKey, issueType, issueTitle, epicId, sprintState, outCol );
    
  }  
 
  
  var outHTML = '<div class="column" id="Parked-wrapper" style="display: none;">' +
  	 	          '<h2>Parked</h2>' +
		          '<div class="item-container sortable-area" id="Parked">' +
		            htmlOut.Parked +
		          '</div>' +
                '</div>' +
                '<div class="column" id="Idea-wrapper">' +
		          '<h2>Ideas</h2>' +
		          '<div class="item-container sortable-area" id="Idea">' +
		            htmlOut.Idea +
		          '</div>' +
                '</div>' +  
                '<div class="column" id="Planned-wrapper">' +
		          '<h2>Planned</h2>' +
		          '<div class="item-container sortable-area" id="Planned">' +
		            htmlOut.Planned +
		          '</div>' +
                '</div>' +
                '<div class="column" id="Ongoing-wrapper">' +
		          '<h2>Ongoing</h2>' +
		          '<div class="item-container" id="Ongoing">' +
		            htmlOut.Ongoing +
		          '</div>' +
                '</div>';
    
  Logger.log('outHTML completed');
  return outHTML;
}

// ******************************************************************************************************
// Generate the HTML for an item for the roadmap page
// ******************************************************************************************************
function createItemHMTL(key, type, displayText, epic, sprintState, column ){
   var itemHTML = '<li id="' + key + '" class="epic-state-' +sprintState + '" data-column="' + column + '">' +
                     '<span class="drag-handle">☰</span>' +
                     '<div class="type-' + type + '">' + displayText + '</div>' +
                     '<span class="epic epic-'+ epic +'"></span>' +
                     '<a class="link-item" href="https://jobladder.atlassian.net/browse/' + key + '" target="_blank"><i class="material-icons">launch</i></a>' +  
                   '</li>'; 
   return itemHTML;
}



// ******************************************************************************************************
// Print dropdown to choose the epic for an item
// ******************************************************************************************************
function printOptionsForEpic() {
 
  var html = '';
  var data = getEpics();
  var diLength = data.issues.length;
  for (var i=0; i<diLength; i++) {
    html = html + '<option value="'+ data.issues[i].key + '">' + data.issues[i].fields.customfield_10009 + '</options>';
  }  
  return html;
  
}  



// ******************************************************************************************************
// Clean the input so we can embed it in HTML without breaking it - from https://stackoverflow.com/questions/1787322/htmlspecialchars-equivalent-in-javascript/4835406#4835406
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



// ******************************************************************************************************
// Function to check that the user has the permission to create items
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
// Function to set who has permission to create items
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