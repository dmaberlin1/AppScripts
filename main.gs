


function sendMap() {
    const sh=SpreadsheetApp.getActiveSheet();
    const address=sh.getRange('A1').getValue()
    const map=Maps.newStaticMap().addMarker(address);
    GmailApp.sendEmail('dmaberlin77@gmail.com','Map','see bottom',{
        attachments:[map],
    })
}


function func1() {
    
}

function func2() {
    
}
function func3() {
    
}




function Myfunc() {
    var spreadsheet=SpreadsheetApp.getActive()
    var sheet=spreadsheet.getActiveSheet();
    sheet.getRange(spreadsheet.getCurrentCell().getRow(),1,1,sheet.getMaxColumns()).activate()
    spreadsheet.getActiveRangeList().setBackground('#4c1130')
    .setBorder('#ffffff')
    .setFontWeight('bold');
    spreadsheet.getActiveSheet().setFrozenRows(1);

}



function addToDatabase0() {
    var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1gsPEw8zoP5nuWnq4EAkaY62Ympy1BGb069BP2VhGBXg/edit?usp=sharing');
    var sh_db_movies=ss.getSheetByName('dbmovie');
    var values=sh_db_movies.getRange(2,1,11,11).getValues()
      Logger.log(values)
}

function addToDatabase(){
    var ss=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dbmovie')
    var ss_active=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('db_active')
    var values=sh.getRange(2,1,11,11).getValues()
    
    Logger.log(values)
    
}

function addHeaders(){
    var ss_db_movies=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dbmovie')
    var ss_active=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('db_active')
    var headers=ss_db_movies.getSheetValues(1,1,1,11)
    Logger.log(headers)
}



//**********************UI ******************/

function onOpen() {
    var ui=SpreadsheetApp.getUi();
    ui.createMenu('Custom Menu')
    .addItem('btn1','func1')
    .addSeparator()
    .addItem('btn2','func2')
    .addSeparator()
    .addItem('btn2','func3')
    .addSeparator()
    .addItem('btn2','func4')
    .addToUi();
}