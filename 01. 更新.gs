function update() {
  var dates = [];
  var datesCol = 1;
  var startDatesRow = 5;
  var dateRow = 20;
  var dateCol = 19;
  var dateCell = samarySheet.getRange(dateRow, dateCol);//更新日時
  
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
  
  //更新日付入力
  var now = new Date(); 
  var year = now.getFullYear();
  var month = now.getMonth()+1;
  var date = now.getDate();
  var hour = now.getHours();
  var min = now.getMinutes();
  var sec = now.getSeconds();
  
  dateCell.setValue("最終更新日時 : " + year + "/" + month + "/" + date + " " + hour + ":" + min + ":" + sec);

}