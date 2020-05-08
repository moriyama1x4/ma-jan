function copySheet(dates) {
  var mastersSheet = book.getSheetByName("Master");

  //シート複製
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
  var targetCol = 3;
  
  var targetRange = samarySheet.getRange(1, targetCol, samarySheet.getLastRow(), 1);
  targetRange.copyTo(targetRange);
}