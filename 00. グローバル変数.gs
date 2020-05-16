var book = SpreadsheetApp.getActive();
var samarySheet = book.getSheetByName("総合")
var mastersSheet = book.getSheetByName("Master");
var memberNum = 7;

//全日付
var dates = [];
var startDatesRow = 5;
var datesCol = 2;

for(var i = startDatesRow; i <= samarySheet.getLastRow(); i++){
  var date = samarySheet.getRange(i, datesCol).getValue();
  if(!date){
    break;
  }
  dates.push(samarySheet.getRange(i, datesCol).getValue());
}
  