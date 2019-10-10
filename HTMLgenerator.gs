// ******************************************************************************************************
// Generate the HTML used for the columns of items in the roadmap page
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
    htmlOut[outCol] += createIssueHMTL(issueKey, issueType, issueTitle, epicId, sprintState, outCol );
    
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
function createIssueHMTL(key, type, displayText, epic, sprintState, column ){
   var itemHTML = '<li id="' + key + '" class="epic-state-' +sprintState + '" data-column="' + column + '">' +
                     '<span class="drag-handle">☰</span>' +
                     '<div class="type-' + type + '">' + displayText + '</div>' +
                     '<span class="epic epic-'+ epic +'"></span>' +
                     '<a class="link-item" href="https://jobladder.atlassian.net/browse/' + key + '" target="_blank"><i class="material-icons">launch</i></a>' +  
                   '</li>'; 
   return itemHTML;
}



// ******************************************************************************************************
// Print dropdown to choose the epic when creating an item
// ******************************************************************************************************
function printOptionsForEpic() {
 
  var html = '';
  var data = getEpics();  
  var diLength = data.issues.length;
  Logger.log(diLength);
  for (var i=0; i<diLength; i++) {
    if ( data.issues[i].key == 'CJ-5309') {    Logger.log(data.issues[i]) }
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