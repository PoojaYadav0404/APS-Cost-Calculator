var Cal = "Calculator"
var DD = "Dropdown"
var Shiprocket = "Shiprocket (with Gst)"
var Indiapost = "India Post"
var MailInno = "Mail Innovation"
var SS = SpreadsheetApp.getActiveSpreadsheet();
var CalSheet = SS.getSheetByName(Cal);
var DDSheet = SS.getSheetByName(DD);
var ShiprocketSheet = SS.getSheetByName(Shiprocket);
var IndiapostSheet = SS.getSheetByName(Indiapost);
var MailInnoSheet = SS.getSheetByName(MailInno);
var ActiveSheet = SS.getActiveSheet();
var Row = ActiveSheet.getActiveCell().getRow();
var Col = ActiveSheet.getActiveCell().getColumn();

function onOpen(){

  var Ui = SpreadsheetApp.getUi();
  Ui.createMenu("NEW MENU")
  .addItem("Address Finder", "addressFinder")
  .addToUi();


  var HiddenSheets = [DD,Shiprocket, Indiapost, MailInno];
  var Sheets = SS.getSheets();
  Sheets.forEach(function (s){
    if(HiddenSheets.includes(s.getName())){
      s.hideSheet();
    }
  })
  
  
}

function onEdit(){

  itemWeightAndPP();
  shiprocketZoneChrg();
  indiapostZoneChrg();
  mailInnovation();
  dropdowns();

}


 
