//set the global vars
var sheetID = "1qPR7oEYMi-Rzfr-D8nTqUDH5kWdX43PbP9tmeIrDI3s"; //needed as you cannot use getActiveSheet() while the sheet is not in use (as in a standalone application like this one)
var scriptURL = "https://script.google.com/macros/s/AKfycbytiMYA6sUnb5dRE3yVzXUJApEDzk5f9Sm_Ihx1pXI0NW4pfLk/exec";  //the URL of this web app


/* 
To run the script that interact with Google drive, you need to enable the Google Drive API: 
You can do it by clikcing on "Tools > Script Editor" and then on "Resources > Advanced Google Services"

Important: the REST calls are restricted by the Google Apps Script quotas. The two you are most likely to hit into are: 
1) The maximum runtime for a script is 6 minutes. Individual scripts have a limit of 30 seconds. 
2) The URLfetch (the HTTP/HTTPS service used to make the API calls) has a 20MB maximum payload size per call.
   For full details, see https://developers.google.com/apps-script/guides/services/quotas   
   
*/


// ******************************************************************************************************
// Function to display the HTML as a webApp
// ******************************************************************************************************
function doGet(e) {
  
  //you can also pass a parameter via the URL as ?add=XXX 

  var template = HtmlService.createTemplateFromFile('Dashboard');  

  var htmlOutput = template.evaluate()
                   .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                   .setTitle('Dashboard WFDF Sport Dev')
                   .addMetaTag('viewport', 'width=device-width, initial-scale=1')
                   .setFaviconUrl('http://threeflamesproductions.com/wp-content/uploads/2017/01/Favicon_ThreeFlames_FireIcon_Color.png');

  return htmlOutput;
};



// ******************************************************************************************************
// Function to print out the content of an HTML file into another (used to load the CSS and JS)
// ******************************************************************************************************
function getContent(filename) {
  var pageContent = HtmlService.createTemplateFromFile(filename).getRawContent();
  return pageContent;
}


// ******************************************************************************************************
// Function to return the full HTML of a page by stitching together different page components
// ******************************************************************************************************
function createHTML(pagename, pageTitle, bodyClass) {
 
  var html = '<!DOCTYPE html>' +
             '<html>' +
               '<head>' +
                 '<base target="_top">' + 
                  getContent('head') +                   
                  '<title>' + pageTitle + '</title>' +
               '</head>' +
               '<body id="page-top" class="' + bodyClass + '">' +
               getContent('navigation') +  
               getContent(pagename) +
               getContent('footer') + 
               '</body>' +
             '</html>'  
  return html;               
}



// ******************************************************************************************************
// Function to shortcut writing a call for a user property
// ******************************************************************************************************
function printVal(key) {

  if (key == "git_token" || key == "git_user") {
     var dummy = PropertiesService.getUserProperties().getProperty(key);    
  } else {
     var dummy = PropertiesService.getScriptProperties().getProperty(key);    
  }
  return dummy;
  //TODO - add logic for when the property is not defined
}



// ******************************************************************************************************
// Function to create menus when you open the sheet
// ******************************************************************************************************
function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{name: "Configure the github repo", functionName: "githubRepoConfigure"},
                     {name: "Set GitHub authentication", functionName: "setGithubAuthToken"}
                    ]; 
  ss.addMenu("GitHub", menuEntries);
}




// ******************************************************************************************************
// Function to generate the HTML for a project page
// ******************************************************************************************************
function generateProjectHTML(id) {
  
  var html = '';
  var idString = id.toString();
  var data = sheet2Json('Projects');
  //filter the data to get the project with thr required id
  var projectDataArray = data.filter(function (entry) {
    return entry.id == idString;
  });
  //the above returns an array - you want just the first item
  var projectData = projectDataArray[0];
  
  var bodyClass = 'project-page';

  html += '<!DOCTYPE html>' +
            '<html>' +
              '<head>' +
                '<base target="_top">' + 
                getContent('head') +                   
                '<title>' + projectData.title + ' - ToGetThere</title>' +
              '</head>' +
              '<body id="page-top" class="' + bodyClass + '">' +
                getContent('navigation');  
               
	
  // Header
  var headerImage_url = projectData.headerImage_url;
  if(!headerImage_url) { headerImage_url = '/images/header.jpg';}    
    html += '<header class="masthead" style="background-image: url(' + headerImage_url + ')">' +
			  '<div class="container">' +
				'<div class="intro-text">' +
					'<div class="intro-lead-in">' + projectData['headline'] + '</div>' +
					'<div class="intro-heading text-uppercase">' + projectData['title'] + '</div>' +
					'<a class="btn btn-primary btn-xl text-uppercase js-scroll-trigger" href="#donate">Give</a>' +
				'</div>' +
			  '</div>' +
			'</header>';

  // Project description - you have 3 rows
  html += '<section id="description">' +
     	    '<div class="container">';
    
  if (projectData.top_row) { 
    html += '<div class="row">' +
              '<div class="col-lg-12">' +
                projectData.top_row +
              '</div>' +
            '</div>';
  }
    
  if (projectData.middle_row) { 
    html += '<div class="row">' +
              '<div class="col-lg-12">' +
                projectData.middle_row +
              '</div>' +
            '</div>';
  }
    
	
  if (projectData.bottom_row) { 
    html += '<div class="row">' +
              '<div class="col-lg-12">' +
                projectData.bottom_row +
              '</div>' +
            '</div>';
  }
	
  html += '</section>';

	
  
  // Donate section - up to 3 calls to action
  var colCount = 0;
  if(projectData.money_url) { colCount = colCount+1 }
  if(projectData.equipment_text) { colCount = colCount+1 }
  if(projectData.service_text) { colCount = colCount+1 }
	
  if (colCount> 0) {
    var colSize = 12/colCount;
    
    html += '<section id="donate">' +
              '<div class="container">' +
				'<div class="row text-center">';
				  if(projectData.money_url) {
					html += '<div class="col-md-' + colSize + '">' +
							  '<a class="cta cta-ask cta-button" data-toggle="modal" href="' + projectData.money_url +'">' +
								'<span class="fa-stack fa-4x">' +
								  '<i class="fa fa-circle fa-stack-2x text-primary"></i>' +
								  '<i class="fa fa-question fa-stack-1x fa-inverse"></i>' +
								'</span>' +
							  '</a>' +
							  '<h4 class="service-heading">' +
							    '<a class="cta cta-ask cta-text" data-toggle="modal" href="' + projectData.money_url + '">Donate</a>' +
							  '</h4>' +
							  '<p class="text-muted">' + projectData.money_text + '</p>' +
							'</div>';
				  }
						
				  if(projectData.equipment_text) {
					if (!projectData.equipment_url) { projectData.equipment_url = '#give-equipment'}
					  html += '<div class="col-md-' + colSize + '">' +
					 		    '<a class="cta cta-ask cta-button" data-toggle="modal" href="' + projectData.equipment_url + '">' +
								  '<span class="fa-stack fa-4x">' +
									'<i class="fa fa-circle fa-stack-2x text-primary"></i>' +
									'<i class="fa fa-question fa-stack-1x fa-inverse"></i>' +
								  '</span>' +
								'</a>' +
								'<h4 class="service-heading">' +
								  '<a class="cta cta-ask cta-text" data-toggle="modal" href="' + projectData.equipment_url + '">Give</a>' +
								'</h4>' +
								'<p class="text-muted">' + projectData.equipment_text + '</p>' +
							  '</div>';
				  }
						
				  if(projectData.service_text) {
					if (!projectData.service_url) { projectData.service_url = '#give-service'}
					  html += '<div class="col-md-' + colSize + '">' +
								'<a class="cta cta-ask cta-button" data-toggle="modal" href="' + projectData.service_url + '">' +
					   			  '<span class="fa-stack fa-4x">' +
									'<i class="fa fa-circle fa-stack-2x text-primary"></i>' +
									'<i class="fa fa-question fa-stack-1x fa-inverse"></i>' +
								  '</span>' +
								'</a>' +
								'<h4 class="service-heading">' +
								  '<a class="cta cta-ask cta-text" data-toggle="modal" href="' + projectData.service_url + '">Give</a>' +
								'</h4>' +
								'<p class="text-muted">' + projectData.service_text + '</p>' +
							  '</div>';
				  }

				  html += '</div>' +
				        '</div>' +
			          '</section>';
  }
	
  // Team - TODO
  if(projectData.has_team) {
    html += '<section class="bg-light" id="team">' +
			'</section>';
  }
  
  // Links - TODO
  var colCount = 0;
  if(projectData.facebook_url) { colCount = colCount+1 }
  if(projectData.twitter_url) { colCount = colCount+1 }
  if(projectData.site_url) { colCount = colCount+1 }
  
  if (colCount> 0) {
    var colSize = 12/colCount;
    
  }  
    
  html += getContent('footer') + 
        '</body>' +
      '</html>';  

  return html;

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
// Wrapper for the convertSheet2JsonText - pass the sheetName to get the JSOn array out
// ******************************************************************************************************
function sheet2Json(sheetName) {
  var ss = SpreadsheetApp.openById(sheetID);
  var sheets = ss.getSheets();
  var sh = ss.getSheetByName(sheetName);
  return convertSheet2JsonText(sh);
}

// ******************************************************************************************************
// Function to convert a string into a SEO-friendly URL
// from https://stackoverflow.com/questions/14107522/producing-seo-friendly-url-in-javascript
// ******************************************************************************************************
function toSeoUrl(textToConvert) {
    return textToConvert.toString()               // Convert to string
        .normalize('NFD')               // Change diacritics
        .replace(/[\u0300-\u036f]/g,'') // Remove illegal characters
        .replace(/\s+/g,'-')            // Change whitespace to dashes
        .toLowerCase()                  // Change to lowercase
        .replace(/&/g,'-and-')          // Replace ampersand
        .replace(/[^a-z0-9\-]/g,'')     // Remove anything that is not a letter, number or dash
        .replace(/-+/g,'-')             // Remove duplicate dashes
        .replace(/^-*/,'')              // Remove starting dashes
        .replace(/-*$/,'');             // Remove trailing dashes
}