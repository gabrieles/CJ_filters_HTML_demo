//set the global vars
var sheetID = "1FLrj1S5pjnSM5btjKHiuC7WzMRwR4t03ZhddbG9mHks"; //needed as you cannot use getActiveSheet() while the sheet is not in use (as in a standalone application like this one)
var settingsName = "Settings"; //the name of the sheet where all the settings are stored - handy to avoid editing code all the time
var scriptURL = "https://script.google.com/macros/s/AKfycbx7gdszBquToEL_Iw6RRQYaa2-X9Qrs1y8FdrnajoFu3HnfmUKq/exec";  //the URL of this web app


// ******************************************************************************************************
// Function to display the HTML as a webApp
// ******************************************************************************************************
function doGet(e) {
  
  //you can also pass a parameter via the URL as ?add=XXX 

  var template = HtmlService.createTemplateFromFile('filters');  

  var htmlOutput = template.evaluate()
                   .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                   .setTitle('CJ Filters')
                   .addMetaTag('viewport', 'width=device-width, initial-scale=1')
                   .setFaviconUrl('https://www.charityjob.co.uk/Assets/img/favicon.ico');;

  return htmlOutput;
};


// ******************************************************************************************************
// Function to automatically run functions when teh sheet is opened
// ******************************************************************************************************
function onOpen(e) {
  updateSheets();
}


// ******************************************************************************************************
// Function to update list of sheets every time the worksheets is opened
// ******************************************************************************************************
function updateSheets() {
  var ss = SpreadsheetApp.openById(sheetID);
  var sheets = ss.getSheets();
  var sSettings = ss.getSheetByName(settingsName);
  var range = sSettings.getRange(2,1,sheets.length,1)
  for(var i=0; i<sheets.length; i++){  
    range.getCell(i+1, 1).setValue(sheets[i].getName());
  }  
}


// ******************************************************************************************************
// Function to print out the content of an HTML file into another (used to load the CSS and JS)
// ******************************************************************************************************
function getContent(filename) {
  var pageContent = HtmlService.createTemplateFromFile(filename).getRawContent();
  return pageContent;
}


// ******************************************************************************************************
// Function to print out the content of the sheets with the menu items
// ******************************************************************************************************
function printFilters() {
  
  var sheetName = '';
  var codeM = ''; //code for the menu
  var sheetAction = '';
  var sheetOut = [];
  
  //get settings
  
  var ss = SpreadsheetApp.openById(sheetID);
  var sSettings = ss.getSheetByName(settingsName);
  var sheets = ss.getSheets();
  
  var showJobs = sSettings.getRange(3, 7, 1,1).getValue();
  
  //loop over all the sheets, and add their content to the respective menus as needed
  for(var h=0; h<sheets.length; h++){  
    
    var sheet = ss.getSheets()[h]; // look at each sheet in workbook 
    var sheetName = sheet.getName(); //get the sheet name
    
    //grab the action from the "settings" sheet, and add the code to the respective menu
    sheetAction = sSettings.getRange(h+2, 2, 1,1).getValue();
    
    Logger.log("Executing " + sheetAction + " for " + sheetName);

    switch (sheetAction) {
        
      case "Dropdown":
        codeM += createFiltersFromSheet(sheet,h);
      break;
    
      default:
        //do nothing
    }
    
  }
  
  if (showJobs == "Yes") {
    
    var codeJ = '<div id="charityJobsContainer"><div id="dummyjobs"></div></div>' + 
                '<div id="jobPanel"></div>' +
                '<script>' +
                   'jQuery(document).ready(function(){ ' + 
                     '$("body").addClass("show-jobs"); ' + 
                     'refreshPage();' +
                   '});'+ 
                '</script>';
  } else {
    
    var codeJ = '';
    
  }
  
  var codeB = '<a class="btn btn-primary btn-center" href="javascript:void(0)">Save search</a>'+
              '<a class="btn btn-primary btn-center" href="javascript:void(0)" onclick="resetFilters();">Reset</a>'; 
  
  var codeMenu = '<div class="filters-wrapper">' + 
                   '<ul class="main-ul">' + 
                      '<li class="filters-header"><h3><i class="fa fa-filter"></i>Filters</h3><span id="filtersCounter" class="filters-counter"></span></li>' + 
                        codeM + 
                   '</ul>' + 
                   codeB + 
                  '</div>' + 
                  codeJ;
  
  return codeMenu; 
  
}



// ******************************************************************************************************
// Utility function to return the content of a sheet as menu
// ******************************************************************************************************

function createFiltersFromSheet(sh, index) {
  var outC = '<li class="has-children">';
  var parentID = '';
  var uniqueID = '';
  var thisCellVal = '';
  var rightCellVal = '';
  var belowCellVal = '';
  var belowLeftCellVal = '';
  var refineBtn = '';
  var collapseClass = '';
  var defaultSelected = '';
  var cellFontWeight = '';
  var hasChildren = false;
  var shName = sh.getName();  
  var lastRow = sh.getLastRow(); 
  var lastCol = sh.getLastColumn(); 
  
  outC = '<li data-toggle="collapse" data-target="#items-in-' + index + '" class="has-children collapsed has-arrow level-0">' + shName + '</li>' +
           '<ul id="items-in-' + index + '" class="collapse">';
  
  var range = sh.getRange(1, 1, lastRow+1, lastCol+1)
  
  //loop over all rows and columns 
  for ( var i=1; i<=lastRow; i++ ){
    for ( var j=1; j<=lastCol; j++ ){
      
    //uniqueID = sheet    -   row   -  column
      uniqueID = index + '-' + i + '-' + j;
      
      
      cellFontWeight = range.getCell(i,j).getFontWeight();      
      if ( cellFontWeight == "bold" ) {
        defaultSelected = ' selected';
      } else {
        defaultSelected = '';
      }
      
      thisCellVal = range.getCell(i,j).getValue();
     
      //skip if empty
      if (thisCellVal) {
        
        rightCellVal = range.getCell(i,j+1).getValue();
        
        //if the cell that has a value is in the first column, or if it has a value on its right, it is a parent 
        if( rightCellVal ) {
          refineBtn = '<a class="refine btn" data-toggle="collapse" data-target="#children-of-' + uniqueID + '">Refine</a>';
          parentID = uniqueID;
          outC += '<li id="id-' + uniqueID + '" class="has-children level-' + j + '">' +
                    '<a class="selectable has-icon' + defaultSelected + '">' + 
                      '<i class="far fa-square"></i><i class="fa fa-check overlap"></i>' + thisCellVal + 
                    '</a>' + 
                    refineBtn + 
                    '<ul id="children-of-' + uniqueID + '" class="collapse">';
        } else {
          outC += '<li id="id-' + uniqueID + '" class="no-children level-' + j + '">' + 
                    '<a class="selectable has-icon' + defaultSelected + '">' + 
                      '<i class="far fa-square"></i><i class="fa fa-check overlap"></i>' + thisCellVal + 
                    '</a>' + 
                  '</li>';
        }
        
        
        //check if it is time to close a menu
        belowCellVal = range.getCell(i+1,j).getValue();
        if(j>1) { 
          belowLeftCellVal = range.getCell(i+1,j-1).getValue();
        } else {
          belowLeftCellVal = '';
        }
        
        
        if ( i==lastRow ) {
          //close the last row only if the cell is not in the first column
          if(j>1) { 
            for ( var k = j; k>1; k-- ) {
              outC += '</ul>';
            }
          } 
          
        } else {
        
          if ( (!belowCellVal) || belowLeftCellVal){
            if(j>1) { 
              for ( var k = j; k>1; k-- ) {
                var cellUnderLeftVal = range.getCell(i+1,k-1).getValue();
                if(cellUnderLeftVal) {
                  outC += '</ul>';
                } 
              }
            } 
          }
          
        }
        
      }
    }
  }
  
  outC += "</ul></li>";
  return outC;
}


// ******************************************************************************************************
// Function to print out the bottom text
// ******************************************************************************************************
function printAdditionalText() {
    
  //get settings
  
  var ss = SpreadsheetApp.openById(sheetID);
  var sSettings = ss.getSheetByName(settingsName);
  
  var textVal = sSettings.getRange(4, 7, 1,1).getValue();
  
  var codeBT = '<div class="additional-text">' + textVal + '</div>'
  return codeBT;
}



// ******************************************************************************************************
// Function to print out the content of the sheets with the menu items
// ******************************************************************************************************
function printSidebar() {
  
  var sheetName = '';
  var codeM = ''; //code for the menu
  var codeC = ''; //code for Chosen search/select 
  var sheetAction = '';
  var sheetOut = [];
  
  //get settings
  
  var ss = SpreadsheetApp.openById(sheetID);
  var sSettings = ss.getSheetByName(settingsName);
  var sheets = ss.getSheets();
  
  var showJobs = sSettings.getRange(3, 7, 1,1).getValue();

  
  //print menus

  //loop over all the sheets, and add their content to the respective menus as needed
  for(var h=0; h<sheets.length; h++){  
    
    var sheet = ss.getSheets()[h]; // look at each sheet in workbook 
    var sheetName = sheet.getName(); //get the sheet name
    
    //grab the action from the "settings" sheet, and add the code to the respective menu
    sheetAction = sSettings.getRange(h+2, 2, 1,1).getValue();
    
    switch (sheetAction) {
      case "Dropdown":
        sheetOut = createSidebarItems(sheet,h);
        codeM += '<li id="li-' + h + '" class="has-children dropdown"><input type="checkbox" name ="group-' + h + '" id="group-' + h + '"><label for="group-' + h + '">' + sheetName + '</label><ul>';
        codeM += sheetOut[0];
        codeM += '</ul></li>';  
        
        codeC += sheetOut[1];
        
        break;
      case "Multiselect":
        
        break;
      case "Subcategory":
        sheetOut = createSubcategoryItems(sheet,h);
        codeM += '<li id="li-' + h + '" class="has-children subcategory"><input type="checkbox" name ="group-' + h + '" id="group-' + h + '"><label for="group-' + h + '">' + sheetName + '</label><ul>';
        codeM += sheetOut[0];
        codeM += '</ul></li>'; 
        
        codeC += sheetOut[1];
               
        break;  
      default:
        //do nothing
    }  
  }
  
  codeC = '<div class="select-wrapper"><select data-placeholder="Choose a filter here or in the list below" multiple class="chosen-select">' + codeC + '</select></div>';
  //insert the code in the menu wrapper
  if(codeM) { codeM = '<ul class="cd-accordion-menu animated">' + codeM + '</ul>'; }
  
  
  //not currently in use
  var codeB = '<a class="btn" href="' + ss.getUrl() + '" target="_blank">Edit Options</a>'; 
  
  if (showJobs == "Yes") {
    
    var codeJ = '<div id="charityJobsContainer"><div id="dummyjobs"></div></div>' + 
                '<div id="jobPanel"></div>' + 
                '<script>' +
                   'jQuery(document).ready(function(){ ' + 
                     '$("body").addClass("show-jobs"); ' + 
                   '});'+ 
                '</script>';
  } else {
    
    var codeJ = '';
    
  }
     
  // var codeMenu = '<div class="filters-wrapper">' + codeC + codeM +'</div>'  + codeJ ; 
  var codeMenu = '<div class="filters-wrapper">' + codeM +'</div>'  + codeJ ; 
  return codeMenu; 
  
}


// ******************************************************************************************************
// Utility function to return the content of a sheet as menu
// ******************************************************************************************************

function createSidebarItems(sh, index) {
  var outC = '';
  var shName = sh.getName();
  var outS = '<optgroup label="' + shName + '">';
  var gIndex = '';
  var lastRow = sh.getLastRow(); 
  var lastCol = sh.getLastColumn(); 
  
  var isIt = false;
  
  var range = sh.getRange(1, 1, lastRow+1, lastCol+1)
  var addCode = '';
  
  //loop over the rows and columns - each one is a level
  for ( var i=1; i<=lastRow; i++ ){
    for ( var j=1; j<=lastCol; j++ ){
      gIndex = index + '-' + i + '-' + j;
      var thisCellVal = range.getCell(i,j).getValue();
      
      //skip if empty
      if (thisCellVal) {
        
        outS += '<option class="item-' + gIndex + ' col-class-' + j + ' " value="' + thisCellVal + '">' + thisCellVal + '</option>';

        //if it has children, set the class, print an input box and open an <ul>
        var cellRightVal = range.getCell(i,j+1).getValue();
        if (cellRightVal) {
          outC += '<li id="item-' + gIndex + '" class="has-children"><input type="checkbox" name ="sub-group-' + gIndex + '" id="group-' + gIndex + '"><label for="group-' + gIndex + '">' + thisCellVal + '</label><ul>';
        } else {       
          outC += '<li id="item-' + gIndex + '" class="no-children"><a class="selectable">' + thisCellVal + '</a></li>';
        }
                         
        //check if it is time to close a menu
        var cellUnderVal = range.getCell(i+1,j).getValue();
        var cellUnderLeftVal = '';
        if(j>1) { 
          var cellUnderLeftVal = range.getCell(i+1,j-1).getValue();
        } 
        
        
        if ( i==lastRow ) {
          //close the last row only if the cell is not in the first column
          if(j>1) { 
            for ( var k = j; k>1; k-- ) {
              outC += '</ul></li>';
            }
          } 
          
        } else {
        
          if ( (!cellUnderVal && !cellRightVal) || cellUnderLeftVal){
            if(j>1) { 
              for ( var k = j; k>1; k-- ) {
                var cellUnderLeftVal = range.getCell(i+1,k-1).getValue();
                if(cellUnderLeftVal) {
                  outC += '</ul></li>';
                } 
              }
            } 
          }
          
        }
      }
    }
  }
  
  outC += '</optgroup>';
  
  return [outC, outS];
}



// ******************************************************************************************************
// Utility function to return the content of a sheet as menu
// ******************************************************************************************************

function createSubcategoryItems(sh, index) {
  var outC = '';
  var shName = sh.getName();
  var outS = '<optgroup label="' + shName + '">';
  var outT = '<div class="subcategory sub-' + shName + '"><select>'; 
  var parentIndex = 0;
  var lastRow = sh.getLastRow(); 
  var gIndex = '';
  
  var range = sh.getRange(1, 1, lastRow+1, 2)
  
  //loop over all orws, but only on 2 columns 
  for ( var i=1; i<=lastRow; i++ ){
    
    parentIndex = 0;
    
    for ( var j=1; j<=2; j++ ){
      
      gIndex = index + '-' + i + '-' + j;
      var thisCellVal = range.getCell(i,j).getValue();
      
      //skip if empty
      if (thisCellVal) {
        
        //add value to chosen
        outS += '<option class="item-' + gIndex + ' col-class-' + j + ' " value="' + thisCellVal + '">' + thisCellVal + '</option>';
        
        if (j==1) { 
          parentIndex = shName + '-' + i;
          var cellRightVal = range.getCell(i,j+1).getValue();
          if (cellRightVal) {
            outC += '<li id="item-' + gIndex + '" class="has-children subcat parent-item"><a class="selectable">' + thisCellVal + '</a><a class="refine" data-toggle="collapse" data-target="#list-' + parentIndex + '">Refine</a><ul id="list-' + parentIndex + '">';
          } else {
            outC += '<li id="item-' + gIndex + '" class="no-children subcat child-item"><a class="selectable">' + thisCellVal + '</a></li>';
          }
        } else {       
          outC += '<li id="item-' + gIndex + '" class="no-children subcat child-item"><a class="selectable">' + thisCellVal + '</a></li>';
        }  
        
        //check if it is time to close a menu
        var cellUnderVal = range.getCell(i+1,j).getValue();
        var cellUnderLeftVal = '';
        if(j>1) { 
          var cellUnderLeftVal = range.getCell(i+1,j-1).getValue();
        } 
        
        
        if ( i==lastRow ) {
          //close the last row only if the cell is not in the first column
          if(j>1) { 
            for ( var k = j; k>1; k-- ) {
              outC += '</ul></li>';
            }
          } 
          
        } else {
        
          if ( (!cellUnderVal) || cellUnderLeftVal){
            if(j>1) { 
              for ( var k = j; k>1; k-- ) {
                var cellUnderLeftVal = range.getCell(i+1,k-1).getValue();
                if(cellUnderLeftVal) {
                  outC += '</ul></li>';
                } 
              }
            } 
          }
          
        }
        
      }
    }
  }
  

  
  outC += '</optgroup>';
  
  return [outC, outS];
}












// ******************************************************************************************************
// Function to print out the tasks
// ******************************************************************************************************
function printTasks() {
  
  var codeTask = ''; 
  var sheetName = '';
  var headingName = '';
  var cellVal = '' ;
  var j = 0; 
  var nextID = '';
  var solVal = '';
  var taskList = [];
  var colIndex = 0;
  var maxTasks = 0;
 
  var ss = SpreadsheetApp.openById(sheetID);
  var sheet = ss.getSheetByName("Tasks");
 
  
  var lastCol = sheet.getLastColumn();  
 
  var doWeLimit = sheet.getRange(1, 1, 1,1).getValue();
 
  
  //create array of all task column identifiers, including only tasks that have been marked as "Included"
  for (j=2; j<lastCol; j++ ) { 
    if ( sheet.getRange(3, j, 1,1).getValue() == 'Included' ) {
      taskList.push(j); 
      maxTasks++;
    }  
  } 

  //reset the number of max tasks to display if needed (and if it is lower than the number of included tasks)
  if( doWeLimit != 'Display all') { 
    var mTasks = Number(/\d+/.exec(doWeLimit));
    if ( mTasks < maxTasks ) { maxTasks = mTasks; }
  }
  
  //create list of elements to display
  taskList = shuffle(taskList);           //shuffle the list of tasks
  taskList = taskList.slice(0, maxTasks); //truncate the list of tasks to the desired number
  taskList.push('thankyou');              //the l;ast item to display is the "thank you" box
  
  
  
  //print welcome message
  codeTask += '<div id="welcome" class="boxMessage">'; 
    codeTask += '<p>' + ss.getSheetByName(settingsName).getRange(13, 7, 1,1).getValue().replace(/\n/gm,"</p><p>") + '</p>';   
    codeTask += '<button class="control green" onclick="var el= this.parentNode; hide(el); show(\'task' + taskList[0] + '\');">Hide this message and show the first task</button>';
  codeTask += '</div>';
    
  
 
  //print list of tasks
  codeTask += '<div id="task-wrapper">';
  
 
  for (j=0; j<maxTasks; j++ ){
   
    colIndex = taskList[j];
     
    cellVal = sheet.getRange(1, colIndex, 1,1).getValue();
   
   
   
    //print a task only if there is a task description (or else the feedback column will create trouble)
    if (cellVal.length > 1) {
      solVal = sheet.getRange(2, colIndex, 1,1).getValue();
     
      codeTask += '<div class="task hidden" id="task' + colIndex + '">';
   
        //print the task description and start button
        codeTask += '<div id="taskText' + colIndex + '" class="boxMessage">';  
          codeTask += '<p id="p' + colIndex + '">' + cellVal + '</p>'; //task description
          codeTask += '<button id="start' + colIndex + '" class="control red" onclick="startTask(' + colIndex + ',\'' + solVal + '\') ">Start</button>'; 
        codeTask += '</div>';
     
     
        //print task reminder and the "I give up" button (start as hidden)
        codeTask += '<div id="taskBtn' + colIndex + '" class="hidden taskbtn-wrapper">';  
          codeTask += '<p>Task: '+ cellVal + '</p>'; // add reminder text above the "I give up" button
          codeTask += '<button id="out' + colIndex + '" class="control red giveup" onclick="stopTask(); setSolution(\'I give up\'); ">I give up!</button>';
        codeTask += '</div>';
     
     
        //print the input fields to store the results
        codeTask += '<input class="hidden" type="text" id="storeStart' + colIndex + '" name="storeStart' + colIndex + '" value="" />';
        codeTask += '<input class="hidden" type="text" id="storeStop' + colIndex + '" name="storeStop' + colIndex + '" value="" />';
        codeTask += '<input class="hidden" type="text" id="storeClicks' + colIndex + '" name="storeClicks' + colIndex + '" value="" />';
        codeTask += '<input class="hidden" type="textarea" id="storeMouse' + colIndex + '" name="storeMouse' + colIndex + '" />';
   
      codeTask += '</div>';  
    }
   
  }

  codeTask += '</div>';

  
  
  
  //print feedback textarea, thank you message, and "Submit" button
  codeTask += '<div id="thankyou" class="hidden boxMessage thankyou">';
    codeTask += '<textarea id="feedback" name="feedback" value="" placeholder="Any feedback? Type it here!"></textarea>';
    codeTask += '<p class="message">Click on "Submit", and you are done!</p>';  
    codeTask += '<button class="control red" onclick="var outV = outArray(' + lastCol + '); google.script.run.printResult(outV, \'Tasks\'); hide(\'thankyou\'); show(\'restart\')">Submit</button>';
  codeTask += '</div>';
  
  
  //print restart code
  codeTask += '<div id="restart"  class="hidden" >';
  codeTask += '<a  class="control red" href="' + scriptURL + '?action=measure" style="display:inline-block;">Do it again?</a>';
  codeTask += '<p class="message">The list of tasks is randomly generated every time</p>';
  codeTask += '</div>';
  
  
  return codeTask; 
}





// ******************************************************************************************************
// Function to print a new row of items (pas as the array "outArr") in the spreadsheet (identified by name, passed as "targetSheet")
// ******************************************************************************************************
function printResult(outArr, targetSheet){
  var ss = SpreadsheetApp.openById(sheetID);
  var sheet = ss.getSheetByName(targetSheet);
 
  //lock to avoid concurrent writes 
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
  
  // get column where to print the data
  var rowNum = sheet.getLastRow()+1; 
    
  //print the timestamp
  var timestap = new Date();
  sheet.getRange(rowNum, 1, 1, 1).setValue(timestap);
  
  //print all collected values.
  //it is more efficient to print a single array that to set each value individually
  sheet.getRange(rowNum,2,1,outArr.length).setValues([outArr]);
  
  
  lock.releaseLock();
  
}



// ******************************************************************************************************
// Function to print all the values in the spreadsheet
// ******************************************************************************************************
function submitTask(outArr){
  var ss = SpreadsheetApp.openById(sheetID);
  var sheet = ss.getSheetByName("Tasks");
  
  //lock to avoid concurrent writes 
  var lock = LockService.getPublicLock();
  
  // get column where to print the data
  var lastCol = sheet.getLastColumn(); 
  sheet.insertColumnAfter(lastCol-1); //add a column before the last (where "feedback" is stored)
  
  //print the task description and solution, and set the status as "Submitted"
  var oArr = [];
  sheet.getRange(1,lastCol,1,1).setValue(outArr[0]);
  sheet.getRange(2,lastCol,1,1).setValue(outArr[1]);
  sheet.getRange(3,lastCol,1,1).setValue('Submitted');
    
  //add comment to the first cell
  var noteText = 'Audience : ' + outArr[2];
  if (outArr[3]) {noteText += ' Note : ' + outArr[3];}
  sheet.getRange(1,lastCol,1,1).setNote(noteText);
  
  lock.releaseLock();
}




// ******************************************************************************************************
// Function to print all the menu items as a collapsible list of options when creating a new task
// ******************************************************************************************************
function printOptions(){
  var ss = SpreadsheetApp.openById(sheetID);
  var sSettings = ss.getSheetByName(settingsName);
  var sheets = ss.getSheets();
  var sheetName ='';
  var cellVal = '';
  var outC = '<div id="taskSelect" class="select"><ul>';
  
  //loop over the sheets, and the HTML to be used by the JSTree plugin
  for(var h=0; h<ss.getSheets().length; h++){   
    var sheet = ss.getSheets()[h]; // look at every sheet in spreadsheet   
    
    //check that the sheet is part of the menu structure
    cellVal = sSettings.getRange(h+2, 2, 1,1).getValue();
    if (cellVal == "Primary" || h == "Secondary") {
      var lastRow = sheet.getLastRow(); 
      var lastCol = sheet.getLastColumn();
      sheetName = sheet.getName();
      
      outC += '<li class="lvl1">' + sheetName + '<ul>';
    
      //loop over the columns in the sheet
      for ( var j=1; j<lastCol+1; j++ ){
        outC += '<li class="lvl2">' + sheet.getRange(1, j, 1,1).getValue() + '<ul>';
        
        //loop over the rows
        for ( var i=2; i<lastRow+1; i++ ){
	      cellVal = sheet.getRange(i, j, 1,1).getValue();
          if (cellVal != '') {
  	        outC += '<li class="lvl3">' + cellVal + '</li>';
	      }
	    }
        
        outC += '</ul></li>';
      }
    
      outC += '</ul></li>';
    }  
  }
  
  outC += '</ul></div>';
  
  
  
  //initialise jsTree
  outC += '<script> $(function() { $("#taskSelect").jstree( {';
    outC += '"core" : { "multiple" : false, "themes" : { "dots" : false} },';
    outC += '"plugins" : ["search","wholerow"]';
  outC += '}); }); </script>';

  
  //add CSS
  var primaryColor = sSettings.getRange(4, 7, 1,1).getValue();
  
  outC += '<style>';
    outC += 'body .jstree-default .jstree-wholerow-clicked{ background: ' + primaryColor +'; }';
  outC += '</style>';
  
  return outC;
}



// ******************************************************************************************************
// Function to print all the audience groups as a collapsible list of options when creating a new task
// ******************************************************************************************************
function printAudience(){
  var ss = SpreadsheetApp.openById(sheetID);
  var sSettings = ss.getSheetByName(settingsName);
  var cellVal = '';
  
  //print list (start with "everyone")
  var outC = '<div id="audienceSelect" class="select"><ul>';
  outC += '<li class="lvl1">Everyone<ul>';
  
  var lastRow = sSettings.getLastRow();
  //loop over the rows
  for ( var i=2; i<lastRow+1; i++ ){
    cellVal = sSettings.getRange(i, 4, 1,1).getValue();
    if (cellVal != '') {
      outC += '<li class="lvl2">' + cellVal + '</li>';
    }
  }
  outC += '</ul></li></ul></div>';
  
  
  
  //initialise jsTree
  outC += '<script> $(function() { $("#audienceSelect").jstree( {';
    outC += '"core" : { "multiple" : true, "themes" : { "dots" : false} },';
    outC += '"plugins" : ["wholerow","checkbox"]';
  outC += '}); }); </script>';

  
  return outC;
}



// ******************************************************************************************************
// Function to randomise an array from https://stackoverflow.com/questions/2450954/how-to-randomize-shuffle-a-javascript-array
// use it like: arr = shuffle(arr);
// ******************************************************************************************************
function shuffle(array) {
  var currentIndex = array.length, temporaryValue, randomIndex;

  // While there remain elements to shuffle...
  while (0 !== currentIndex) {

    // Pick a remaining element...
    randomIndex = Math.floor(Math.random() * currentIndex);
    currentIndex -= 1;

    // And swap it with the current element.
    temporaryValue = array[currentIndex];
    array[currentIndex] = array[randomIndex];
    array[randomIndex] = temporaryValue;
  }

  return array;
}