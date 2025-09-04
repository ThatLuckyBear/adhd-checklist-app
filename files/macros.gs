function DailyClearTodoOG() { // clears all checkboxes on the Today's Task list -- added to code.gs as a daily occurence
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Today\'s Tasks'), true);
  spreadsheet.getRange('A2:B1000').activate();
  spreadsheet.getActiveRange().setValue('FALSE')
};


function DailyClearTodoNEW() { 
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var rowcount = sheet.getLastRow();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Today\'s Tasks'), true);
  spreadsheet.getRange('A2:B').activate();
  spreadsheet.getActiveRange().setValue('FALSE');
  sheet.showRows(1, rowcount);
};


//___________________________________________________________________


//___________________________________________________________________

//⭐function ClearTodoTEST() { // clears all checkboxes on the Today's Task list -- added to code.gs as a daily occurence
//⭐  var spreadsheet = SpreadsheetApp.getActive();
//⭐  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Today\'s Tasks'), true);
//⭐  spreadsheet.getRange('A2:A100').activate();
//⭐  spreadsheet.getActiveRange().setValue('FALSE')
//⭐  spreadsheet.getActiveSheet().showRows(0, 100);
//⭐};  


//⭐function PopulateTodoUNFINISHED() {
//⭐ var spreadsheet = SpreadsheetApp.getActive();
//⭐ spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Tomorrow\'s Tasks'), true);
//⭐ spreadsheet.getRange('C2:C1000').activate();
//⭐ if(spreadsheet.getActiveRangeList=('TRUE'));
//⭐ debugger;
//⭐};

//function DailyClearTodo() { // clears all checkboxes on the Today's Task list -- added to code.gs as a daily occurence
//  var spreadsheet = SpreadsheetApp.getActive();
//  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Today\'s Tasks'), true);
//  spreadsheet.getRange('A2:A1000').activate();
//  spreadsheet.getActiveRange().setValue('FALSE')
//};  


// function onOpen() {
//   // set up spreadsheet and sheet
//   var ss = SpreadsheetApp.getActiveSpreadsheet(), sheets = ss.getSheets();

//   for(var i = 0, iLen = sheets.length; i < iLen; i++) {
//     // get sheet
//     var sh = sheets[i];

//     // unhide columns
//     var rCols = sh.getRange("1:1");
//     sh.unhideColumn(rCols);

//     // unhide rows
//     var rRows = sh.getRange("A:A");
//     sh.unhideRow(rRows);
//   }
// }



//⭐function ShowAllRowsMACRO() {
//⭐  var spreadsheet = SpreadsheetApp.getActive();
//⭐  var getall = spreadsheet.getRange("A:A").activate();
//⭐  getall.unhideRows(0, 1000);
//⭐};
//___________________________________________________________________
//Recorded Macro
// absolute kicks back out of bounds error
// relative starts unhiding stuff but then stops might
// can we check TRUE/FALSE on "there are hidden rows?"
//___________________________________________________________________


function absolutemacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('2:2').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getActiveRange().getDataRegion().activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveSheet().showRows(2, 1000);
};

function relativemacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(spreadsheet.getCurrentCell().getRow() - 17, 1, 1, sheet.getMaxColumns()).activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getActiveRange().getDataRegion().activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveSheet().showRows(0, 1000);
};













