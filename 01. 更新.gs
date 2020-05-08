function update() {
  var dates = [];
  var datesCol = 1;
  var startDatesRow = 5;
  
  //全日付取得
  for(var i = startDatesRow; i <= samarySheet.getLastRow(); i++){
    var date = samarySheet.getRange(i, datesCol).getValue();
    if(!date){
      break;
    }
    dates.push(samarySheet.getRange(i, datesCol).getValue());
  }
  
  copySheet(dates);
  arranegFormula(dates);
  updateDailyWinRate();
  updateGameWinRate();
}