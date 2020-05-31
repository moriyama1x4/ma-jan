function prepareGame(){
  copySheet(dates);
  arranegFormula(dates);
}

/******
日毎シート複製
*******/
function copySheet(dates) {

  dates.forEach(function(value, index){
    if(!book.getSheetByName(value)){//ない時のみ複製
    
    var copySheet = mastersSheet.copyTo(book);
    copySheet.setName(value);
    
    //順番
    SpreadsheetApp.setActiveSheet(copySheet);
    SpreadsheetApp.getActive().moveActiveSheet(samarySheet.getIndex() + 1);
    
    }
  });
  
  //参照がバグるので、こぴって直す
  var targetCol = 4;
  
  var targetRange = samarySheet.getRange(1, targetCol, samarySheet.getLastRow(), 1);
  targetRange.copyTo(targetRange);
}



/******
順位数式複製
*******/
function arranegFormula() {
  var height = 4;
  var width = 7;
  var samaryStartTargetRow = 5;
  var samaryStartTargetCol = 21;
  var samaryTargetCell = samarySheet.getRange(samaryStartTargetRow, samaryStartTargetCol);
  var samaryTargetRange = samarySheet.getRange(samaryStartTargetRow, samaryStartTargetCol, height, width);
  var samaryFormulaKeyCell = "T5"
  
  
  for(var h = 0; h < seasons.length; h++){
    var seasonSheet = book.getSheetByName(seasons[h]); 
    var seasonStartTargetRow = samaryStartTargetRow;
    var seasonStartTargetCol = 20;
    var seasonTargetCell = seasonSheet.getRange(seasonStartTargetRow, seasonStartTargetCol);
    var seasonTargetRange = seasonSheet.getRange(seasonStartTargetRow, seasonStartTargetCol, height, width);
    
    //シーズン対象日付取得
    var seasonDates = [];
    var seasonStartDatesRow = 5;
    var seasonDatesCol = 1;
    for(var i = seasonStartDatesRow; i <= seasonSheet.getLastRow(); i++){
      var date = seasonSheet.getRange(i, seasonDatesCol).getValue();
      if(!date){
        break;
      }
      seasonDates.push(seasonSheet.getRange(i, seasonDatesCol).getValue());
    }
    
    //シーズン数式生成
    var seasonFormulaKeyCell = "T3"
    var seasonFormula = "'" + seasonDates[0] + "'!" + seasonFormulaKeyCell;
    for(var i = 1; i < seasonDates.length; i++){
      seasonFormula += ("+'" + seasonDates[i] + "'!" + seasonFormulaKeyCell);
    }
    
    //シーズン数式セット
    seasonTargetCell.setFormula(seasonFormula);
    seasonTargetCell.copyTo(seasonTargetRange); 
  }
  
  //総合数式生成
  var samaryFormulaKeyCell = "T5"
  var samaryFormula = "'" + seasons[0] + "'!" + samaryFormulaKeyCell;
  for(var i = 1; i < seasons.length; i++){
    samaryFormula += ("+'" + seasons[i] + "'!" + samaryFormulaKeyCell);
  }
  
  //総合数式セット
  samaryTargetCell.setFormula(samaryFormula);
  samaryTargetCell.copyTo(samaryTargetRange); 
}