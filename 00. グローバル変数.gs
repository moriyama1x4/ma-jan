var book = SpreadsheetApp.getActive();
var samarySheet = book.getSheetByName("総合")
var mastersSheet = book.getSheetByName("Master");
var activeSheet = book.getActiveSheet();
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
  
//全シーズン取得
var seasons = [];
var startSeasonsRow = startDatesRow;
var seasonsCol = 1;

for(var i = startSeasonsRow; i <= samarySheet.getLastRow(); i++){
  var season = samarySheet.getRange(i, seasonsCol).getValue();
  if(!season){
    break;
  }
  //被ってたらskip
  var skipFlag = false;
  for(var j = 0; j <= seasons.length; j++){
    if(season == seasons[j]){
      skipFlag = true;
    }
  }
  if(skipFlag){
    continue;
  }
  seasons.push(samarySheet.getRange(i, seasonsCol).getValue());
}

//現在日時取得
var now = new Date();
var year = now.getFullYear();
var month = now.getMonth()+1;
var date = now.getDate();
var hour = now.getHours();
var min = now.getMinutes();
var sec = now.getSeconds();
var updateTime = "最終更新日時 : " + year + "/" + month + "/" + date + " " + hour + ":" + min + ":" + sec;
