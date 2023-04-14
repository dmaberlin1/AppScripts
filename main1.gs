function createDoc() {
  const doc=DocumentApp.create(VAL);
}


function doGet(){
  return HtmlService.createHtmlOutputFromFile('index.html')
}


function doGet1(){
  return HtmlService.createHtmlOutput('hello')
}