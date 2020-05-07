function copySheet(dates) {
  var book = SpreadsheetApp.getActive();
  var mastersSheet = book.getSheetByName("Master");
  var totalSheet = book.getSheetByName("総合")
  
  //シート複製
  dates.forEach(function(value, index){
    if(!book.getSheetByName(value)){//ない時のみ複製
    
    var copySheet = mastersSheet.copyTo(book);
    copySheet.setName(value);
    
    //順番
    SpreadsheetApp.setActiveSheet(copySheet);
    SpreadsheetApp.getActive().moveActiveSheet(totalSheet.getIndex() + 1);
    
    }
  });
  
  //参照がバグるので、こぴって直す
  var targetCol = 3;
  
  var targetRange = totalSheet.getRange(1, targetCol, totalSheet.getLastRow(), 1);
  targetRange.copyTo(targetRange);
}