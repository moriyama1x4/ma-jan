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
function arranegFormula(dates) {
  var startRow = 5;
  var startCol = 21;
  var startCell = samarySheet.getRange(startRow, startCol);
  var height = 4;
  var width = 7;
  var formulaKeyCell = "T3"
  var newFormula = "'" + dates[0] + "'!" + formulaKeyCell;
  
  //数式生成
  for(var i = 1; i < dates.length; i++){
    newFormula += ("+'" + dates[i] + "'!" + formulaKeyCell);
  }
  
  //数式セット
  startCell.setFormula(newFormula);
  startCell.copyTo(samarySheet.getRange(startRow, startCol, height, width)); 
}