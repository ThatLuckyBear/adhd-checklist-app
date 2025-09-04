//⭐ means remove comment to re-enable code highlight the CTRL+D to capture all the stars
//  Browser.msgBox('blah', Browser.Buttons.OK_CANCEL); //hacky code to help check values -- replace 'blah' with whatever
// to limit a function to a single sheet use >> ss.getSheetId().toString() != "239661122" << as part of an if statement -- replace the numbers with whatever sheet ID you need to target

//-------------------------------------------------------------------------------------
//✨New Quest~ 
//-------------------------------------------------------------------------------------

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
   fcb.whenTextContains("TRUE");
   filter.setColumnFilterCriteria(8, fcb.build());
   fcb.whenTextContains(currentview); 
   filter.setColumnFilterCriteria(6, fcb.build());  
 }

}

/** functions but does nothing different. The getG and getD vars are getting the values as they should.
 function dropdownpicker(){

 var ss = SpreadsheetApp.getActiveSpreadsheet(); //gets spreadsheet
 var dataSheet = ss.getActiveSheet(); //gets current worksheet
 var currentview = dataSheet.getActiveCell().getValue(); // sucessfully gets value of the cell

 if (currentview != 'Task Filter'){
  var getG = dataSheet.getRange('G2').getValue();
  var getD = dataSheet.getRange('D2').getValue();
  var formula = "="+getG+"<"+getD //getting around string restrictions for variables using + -- this works to make a flexible formula
  var toFilter = dataSheet.getDataRange(); //finds all the cells with data in them
  var filter = toFilter.createFilter();   // assigns the found data to create a filter
  var fcb = SpreadsheetApp.newFilterCriteria();    // variable to use FilterCriteria methods to make a new filter//
   fcb.whenTextContains(currentview); // uses currentview to tell the filter criteria which text it should look for (assigned by editing dropdown)
   filter.setColumnFilterCriteria(6, fcb.build());  // Filters the range based on the 6th column

 }
}
  DOESNT WORK

 if (currentview != 'Task Filter'){
  var getG = dataSheet.getRange('G2').getValue();
  var getD = dataSheet.getRange('D2').getValue();
  var formula = "="+getG+">"+getD
  //   Browser.msgBox(formula, Browser.Buttons.OK_CANCEL);
  var toFilter = dataSheet.getDataRange(); //finds all the cells with data in them
  var filter = toFilter.createFilter();   // assigns the found data to create a filter
  //⭐ disabling these 3 lines in order to build a simpler filter that uses less than equal to before trying to chain them together
  //⭐var fcb = SpreadsheetApp.newFilterCriteria();    // variable to use FilterCriteria methods to make a new filter//
  //⭐ fcb.whenTextContains(currentview); // uses currentview to tell the filter criteria which text it should look for (assigned by editing dropdown)
  //⭐ filter.setColumnFilterCriteria(6, fcb.build());  // Filters the range based on the 6th column
  var fcb1 = SpreadsheetApp.newFilterCriteria();    // variable to use FilterCriteria methods to make a new filter//
   fcb.whenTextContains("TRUE");
  var fcb2 = SpreadsheetApp.newFilterCriteria();    // variable to use FilterCriteria methods to make a new filter//
   fcb.whenTextContains(currentview).build();
  filter.setColumnFilterCriteria(8, fcb1).build();  // Filters the range based on the 6th column


 */




/** "from scratch" backup for dropdownpicker function since we're editing it a lot.
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

//_____________ this is the current workspace. Use the above checkbox function if this one breaks. Plan is to change the current cell so it targets a specific column for a row in
// a separate sheet. The initial plan is to try targeting by creating an ID# for each task OR by using the name as that should be identical. I'd rather use the name but idk how to proceed.
// OH and I have to set this up so it targets a specific sheet no matter what because it's fucking stuff up if there's a checkbox in the first column lol.











