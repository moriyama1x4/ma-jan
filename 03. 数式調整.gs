function arranegFormula(dates) {
  var startRow = 5;
  var startCol = 20;
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