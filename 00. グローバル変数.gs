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

