///Depreciated smartsheetmaker script
/*function smartsheetmaker1() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  //create a second sheet
  if(spreadsheet.getSheets().length<2){
    spreadsheet.insertSheet();
  }

  //A# through K#
  //max 20 to start
  var url = "https://docs.google.com/spreadsheets/d/1HwZWvyBOLKrbTWFwUnA8vQwKhC5VRi-Srp6Ii95ymFY/edit#gid=3166507";
  var rang = "Sheet1!A";
  var start_row = 3;
  var column = 'A';

  //SpreadsheetApp.setActiveSheet(spreadsheet.getSheets()[1]);

  for(var i = 1; i<getClassSize(1)+1; i++){
    spreadsheet.getRange('A'+i).activate();
    spreadsheet.getCurrentCell().setFormula("=importrange(\" "+url+ "\",\"" + rang +start_row+ ":K" + start_row +"\")")
    start_row++;
  }
}*/

function smartsheetmaker2(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var url = sheet.getUrl();
  var urlFull = "\""+url+"\"";
  var first_sheet = sheet.setActiveSheet(sheet.getSheets()[0]).getSheetName();
  var rang = "\""+first_sheet+"!A";

  //Number of sheets we need to make
  var classSize = getClassSize();
  var length = 0;
  var index = 1;
  //note, index +2 is the start of the table
  //make a second sheet
  if(sheet.getSheets().length<2){
    sheet.insertSheet();
  }else{
    sheet.insertSheet();
  }
  //Create the First Smartsheet 
  sheet.setActiveSheet(sheet.getSheets()[index]);
  sheet.getRange('A1').activate();
  sheet.getCurrentCell().setFormula("=importrange("+urlFull+",\""+first_sheet+"!A2:L2\")");
  sheet.getRange('A2').activate();
  sheet.getCurrentCell().setFormula("=importrange(" +urlFull + "," + rang +(index+2)+ ":L" + (index+2) +"\")");
  sheet.getActiveSheet().setName(getName(sheet, index));
  sheet.deleteColumns(13,14);
  sheet.deleteRows(3,998);
  index++;
    

  var total_size = getClassSize()+1+length;
  sheet.setActiveSheet(sheet.getSheets()[index-1]);

  //Create subsequent sheets
  for(index; index<total_size; index++){
    sheet.duplicateActiveSheet();
    sheet.setActiveSheet(sheet.getSheets()[index]);
    sheet.getRange('A2').activate();
    sheet.getCurrentCell().setFormula("=importrange(" + urlFull + "," + rang +(index+2)+ ":L" + (index+2) +"\")");
    sheet.getActiveSheet().setName(getName(sheet,index));
  }
  sheet.setActiveSheet(sheet.getSheets()[0]);
}

function getClassSize(a=0){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(sheet.getSheets()[0]);
  var retVal = sheet.getRange('C1').activate().getValue();
  if(a == 1){
  SpreadsheetApp.setActiveSheet(sheet.getSheets()[1]);
  }
  return retVal; 
}

function getName(sheet, index){
  sheet.setActiveSheet(sheet.getSheets()[0]);
  var name = sheet.getRange('D' + (index+2)).activate().getValue();
  sheet.setActiveSheet(sheet.getSheets()[index]);
  return name;
}

