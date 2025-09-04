


/** 
code snippet saved 10/11
Dropdown picker that applies two filters
One that filters by column 6 (frequency by name)
and one that filters by column 8 (if freqency by number is < days past since last complete)

function dropdownpicker(){

 var ss = SpreadsheetApp.getActiveSpreadsheet(); //gets spreadsheet
 var dataSheet = ss.getActiveSheet(); //gets current worksheet
 var currentview = dataSheet.getActiveCell().getValue(); // sucessfully gets value of the cell

 if (currentview != 'Task Filter'){
  var toFilter = dataSheet.getDataRange(); //finds all the cells with data in them
  var filter = toFilter.createFilter();   // assigns the found data to create a filter
  var fcb = SpreadsheetApp.newFilterCriteria();    // variable to use FilterCriteria methods to make a new filter//
   fcb.whenTextContains("TRUE");
   filter.setColumnFilterCriteria(8, fcb.build());
   fcb.whenTextContains(currentview); 
   filter.setColumnFilterCriteria(6, fcb.build());  
 }


 * //code saved 9/22 fully functional

function DailyClearTodo() { //triggers daily
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var rowcount = sheet.getLastRow();
  var filter = sheet.getFilter();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Today\'s Tasks'), true);
  spreadsheet.getRange('A2:B').activate();
  spreadsheet.getActiveRange().setValue('FALSE');
  spreadsheet.getRange('C1:C1').activate();
  spreadsheet.getActiveRange().setValue('Task Filter')
  sheet.showRows(1, rowcount);
  if (filter){
   filter.remove();
  }
};


function onEdit(e) { 
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheetidchecker = ss.getSheetId().toString();
  var columnchecker = e.range.getColumn();
  var rowchecker = e.range.getRow();
  var filter = ss.getFilter();
if (sheetidchecker == '239661122' && columnchecker < 3){
 rowhider(e);
}
if (sheetidchecker == '239661122' && columnchecker == 3 && rowchecker == 1 && filter){
  filter.remove();
  dropdownpicker();
}
else if (sheetidchecker == '239661122' && columnchecker == 3 && rowchecker == 1){ //pops if both of the above are false IDK how else to do this since variables can't be != checked 
  dropdownpicker();
}
}

function rowhider(e) {
  var ss = SpreadsheetApp.getActiveSheet();
  var columnchecker = e.range.getColumn();
  if (e.value != "TRUE") return;
  if (columnchecker == 1){ // COMPLETE! if column 1 do this
       ss.getCurrentCell().offset(0,4).activate(); //simply targets the cell 3 cells to the right of the active cell (the checkbox) so adjust if you move columns.
       ss.getActiveCell().setValue(new Date());
       }
  if (columnchecker == 2){ // SKIP! if column 2 do this ... nothing for now :D
       }
  ss.hideRows(e.range.rowStart);
};

function dropdownpicker(){

 var ss = SpreadsheetApp.getActiveSpreadsheet(); //gets spreadsheet
 var dataSheet = ss.getActiveSheet(); //gets current worksheet
 var currentview = dataSheet.getActiveCell().getValue(); // sucessfully gets value of the cell

 if (currentview != 'Task Filter'){
  var toFilter = dataSheet.getDataRange(); //finds all the cells with data in them
  var filter = toFilter.createFilter();   // assigns the found data to create a filter
  var fcb = SpreadsheetApp.newFilterCriteria();    // variable to use FilterCriteria methods to make a new filter//
   fcb.whenTextContains(currentview); // uses currentview to tell the filter criteria which text it should look for (assigned by editing dropdown)
   filter.setColumnFilterCriteria(6, fcb.build());  // Filters the range based on the 6th column
 }
}
*/


/** Code Saved 9/20/23
function onEdit(e) { 
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheetidchecker = ss.getSheetId().toString();
  var columnchecker = e.range.getColumn();
  var rowchecker = e.range.getRow();
  var filter = ss.getFilter();
if (sheetidchecker == '239661122' && columnchecker < 3){
 rowhider(e);
}
if (sheetidchecker == '239661122' && columnchecker == 3 && rowchecker == 1 && filter){
  filter.remove();
  dropdownpicker();
}
else if (sheetidchecker == '239661122' && columnchecker == 3 && rowchecker == 1){ //pops if both of the above are false IDK how else to do this since filter can't be checked as != ?? 
  dropdownpicker();
}
}

function rowhider(e) {
  var ss = SpreadsheetApp.getActiveSheet();
  var columnchecker = e.range.getColumn();
  if (e.value != "TRUE") return;
  if (columnchecker == 1){
       }
  if (columnchecker == 2){
       }
  ss.hideRows(e.range.rowStart);
};

function dropdownpicker(){ // cannot remove filter... might need a WAIT?

 var ss = SpreadsheetApp.getActiveSpreadsheet(); //gets spreadsheet
 var dataSheet = ss.getActiveSheet(); //gets current worksheet
 var currentview = dataSheet.getActiveCell().getValue(); // sucessfully gets value of the cell

 if (currentview != 'Task Filter'){
  var toFilter = dataSheet.getDataRange(); //finds all the cells with data in them
  var filter = toFilter.createFilter();   // assigns the found data to create a filter
  var fcb = SpreadsheetApp.newFilterCriteria();    // variable to use FilterCriteria methods to make a new filter//
   fcb.whenTextContains(currentview); // uses currentview to tell the filter criteria which text it should look for (assigned by editing dropdown)
   filter.setColumnFilterCriteria(6, fcb.build());  // Filters the range based on the 6th column
 }
}

*/

//______________________________________________________________
// Code Saved 9/13/23
//______________________________________________________________
//⭐function onEdit(e){ //implemented 9/13/2023 with no edits
//⭐  var ss = SpreadsheetApp.getActiveSheet();
//⭐  if (e.range.columnStart >2 || e.value != "TRUE") return;
//⭐  if (e.range.columnStart !=2 && ss.getSheetId().toString() == "239661122")
//⭐   { 
//   ss.getCurrentCell().setValue("FALSE") //sets checkbox to false immediately upon checking instead of waiting for daily reset
//⭐   ss.getCurrentCell().offset(0,3).activate(); //simply targets the cell 3 cells to the right of the active cell (the checkbox) so adjust if you move columns.
//⭐   ss.getActiveCell().setValue(new Date());
//⭐   ss.hideRows(e.range.rowStart);
//⭐   }
//⭐  else if(e.rangecolumnStart !=1){
//   ss.getCurrentCell().setValue("FALSE") //sets checkbox to false immediately upon checking instead of waiting for daily reset
//⭐   ss.hideRows(e.range.rowStart);
//⭐  }
//⭐  return;
//⭐};


//⭐function DailyClearTodoNEW() { //implemented 9/13/2023 with edit to name to match original trigger
//⭐  var spreadsheet = SpreadsheetApp.getActive();
//⭐  var sheet = spreadsheet.getActiveSheet();
//⭐  var rowcount = sheet.getLastRow();
//⭐  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Today\'s Tasks'), true);
//⭐  spreadsheet.getRange('A2:B').activate();
//⭐  spreadsheet.getActiveRange().setValue('FALSE');
//⭐  sheet.showRows(1, rowcount);
//⭐};




//______________________________________________________________
//code saved ??? (OLD)
//______________________________________________________________
//function DailyClearTodo() { // clears all checkboxes on the Today's Task list -- added to code.gs as a daily occurence
//  var spreadsheet = SpreadsheetApp.getActive();
//  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Today\'s Tasks'), true);
//  spreadsheet.getRange('A2:A1000').activate();
//  spreadsheet.getActiveRange().setValue('FALSE')
//};  

//___________________  Functions beautifully commented out while working on new code, is fully implemented in working sheet

//⭐function onEdit(e){
//⭐  var ss = SpreadsheetApp.getActiveSheet();
//⭐  if (e.range.columnStart !=1 || e.value != "TRUE") return;
//⭐  ss.hideRows(e.range.rowStart);
//⭐  ss.getCurrentCell().offset(0,3).activate(); //simply targets the cell 3 cell to the right of the active cell (the checkbox) so adjust if you move columns.
//⭐  ss.getActiveCell().setValue(new Date());
/////// explanation
///// title: function onEdit is actually a trigger (event trigger) that bypasses being a function and instead acts as a listener of sorts
///// line 1: function checks when an edit happens
///// line 2: if the edit was within a specific range (the first column) AND the value = TRUE (checked) hold on to the information (return)
///// line 3: pull up the spreadsheet you're in, hide the row that was marked true
///// line 4: checks the active cell (which will always be the checkbox) and looks 2 cells to the right and activates THAT cell
///// line 5: inserts the current date into that cell
//⭐};



//-------------------------------------------------------------------------------------



// function onEdit(e){
//   var ss = SpreadsheetApp.getActiveSheet();
//   if (e.range.columnStart !=1 || e.value != "TRUE") return;
//   ss.hideRows(e.range.rowStart);
//   //⭐ss.getCurrentCell().offset(0,3).activate(); //original code that targets the offset cell, same as above formula.
//   ss.getCurrentCell().offset(0,2).copyto.name
//   ss.getCurrentCell().offset(0,3).setValue(name); //this doesn't do anything I guess. I was trying to copy to 
  //above is trying to target C column on "Task Database" sheet with a dynamic row based on the name of the row is checked on "Today's Tasks" sheet 
 //this and above unhide rows when falsed


 /////// explanation
 ///// title: function onEdit is actually a trigger (event trigger) that bypasses being a function and instead acts as a listener of sorts
 ///// line 1: function checks when an edit happens
 ///// line 2: if the edit was within a specific range (the first column) AND the value = TRUE (checked) hold on to the information (return)
 ///// line 3: pull up the spreadsheet you're in, hide the row that was marked true
 ///// line 4: checks the active cell (which will always be the checkbox) and looks 2 cells to the right and activates THAT cell
 ///// line 5: inserts the current date into that cell
 
//};

// function onChange(e){
//  var ss = SpreadsheetApp.getActiveSheet();
//  if (e.range.columnStart !=1 || e.value != "FALSE") return;
//  ss.showRows(e.range.rowStart);

// };
