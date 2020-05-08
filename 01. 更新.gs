function update() {
  var book = SpreadsheetApp.getActive();
  var sheet = book.getSheetByName("総合")
  var dates = [];
  var datesCol = 1;
  var startDatesRow = 5;
  
  //全日付取得
  for(var i = startDatesRow; i <= sheet.getLastRow(); i++){
    var date = sheet.getRange(i, datesCol).getValue();
    if(!date){
      break;
    }
    dates.push(sheet.getRange(i, datesCol).getValue());
  }
  

  copySheet(dates);
  arranegFormula(dates);
  updateDailyWinRate();
  updateGameWinRate();
}