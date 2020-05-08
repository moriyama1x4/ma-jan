function arranegFormula(dates) {
  
  var book = SpreadsheetApp.getActive();
  var sheet = book.getSheetByName("総合")
  var startRow = 5;
  var startCol = 12;
  var startCell = sheet.getRange(startRow, startCol);
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
  startCell.copyTo(sheet.getRange(startRow, startCol, height, width));
  
//  var originFormula = sheet.getRange(startRow, startCol).getFormula();
//  Logger.log(originFormula);
}