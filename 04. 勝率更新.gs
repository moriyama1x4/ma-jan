function inputWinRate() {
  
  var book = SpreadsheetApp.getActive();
  var sheet = book.getSheetByName("総合")
  var startRow = 11;
  var startCol = 12;
  var startCell = sheet.getRange(startRow, startCol);
  var memberNum = 7
  var height = memberNum;
  var width = memberNum;
  var sliceFormula = "SPARKLINE({1,0})" //斜線
  var winLoseArray = [];
  var winRateFormulaArray = [];
  
  //array作成
  
  for(var i = 0; i < memberNum; i++){
    winLoseArray.push([]);
    winRateFormulaArray.push([]);
    
    for(var j = 0; j < memberNum; j++){
    winLoseArray[i].push([0, 0]);
    }
  }
 
  
  //実績取得
  var resultStartRow = 4;
  var resultStartCol = 3; 
  var resultData = sheet.getRange(resultStartRow, resultStartCol, sheet.getLastRow(), memberNum).getValues();
  
  //勝利数入力
  for(var i = 0; true; i++){
    if(resultData[i][0] == ""){
      break;
    }
      
    for(var j = 0; j < memberNum; j++){
      if(resultData[i][j] == "-"){
        continue;
      }
      var selfScore = resultData[i][j];
      
      for(var k = j + 1; k < memberNum; k++){
        if(resultData[i][k] == "-"){
          continue;
        }
        var taegetScore = resultData[i][k];
        
        if(selfScore > taegetScore){
          winLoseArray[j][k][0] ++;
          winLoseArray[k][j][1] ++;
        }
        else if(selfScore == taegetScore){
          winLoseArray[j][k][0] += 0.5;
          winLoseArray[j][k][1] += 0.5;
          winLoseArray[k][j][0] += 0.5;
          winLoseArray[k][j][1] += 0.5;
        }else if(selfScore < taegetScore){
          winLoseArray[j][k][1] ++;
          winLoseArray[k][j][0] ++;
        }
        
      }
    }
  }
  
  //勝率変換
  for(var i = 0; i < memberNum; i++){
    for(var j = 0; j < memberNum; j++){
      if(i == j){
        winRateFormulaArray[i].push(sliceFormula);
      }else{
        winRateFormulaArray[i].push("IFERROR(" + winLoseArray[i][j][0] + "/" + (winLoseArray[i][j][0] + winLoseArray[i][j][1]) + ', "-")')
      }
    }
  }
  
  
  //勝率入力
  sheet.getRange(startRow, startCol, height, width).setFormulas(winRateFormulaArray);
  
  
//  Logger.log(resultData);
//  Logger.log(winLoseArray);
//  Logger.log(winRateFormulaArray);
}