function test1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B5:E5').activate();
  spreadsheet.getCurrentCell().setValue('test');
  spreadsheet.getRange('J7').activate();
};

function headers_fill() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:K2').activate();
  spreadsheet.getActiveRangeList().setBackground('#434343');
};