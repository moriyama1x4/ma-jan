function undateWinRate(){
  var dateRow = 20;
  var dateCol = 20;
  var dateCell = samarySheet.getRange(dateRow, dateCol);//更新日時
  
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

/******
日毎勝率更新
*******/
function updateDailyWinRate() {
  var startRow = 13;
  var startCol = 39;
  var startCell = samarySheet.getRange(startRow, startCol);
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
  var resultStartRow = 5;
  var resultStartCol = 4; 
  var resultData = samarySheet.getRange(resultStartRow, resultStartCol, samarySheet.getLastRow(), memberNum).getValues();
  
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
  samarySheet.getRange(startRow, startCol, height, width).setFormulas(winRateFormulaArray);
  
}



/******
半荘毎勝率更新
*******/
function updateGameWinRate() {
  var startRow = 13;
  var startCol = 21;
  var startCell = samarySheet.getRange(startRow, startCol);
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
  var sheets = book.getSheets();
  var excludeNames = ["Master", "総合"];
  
  //除外シートにシーズン追加
  for(var i = 0; i < seasons.length; i++){
    excludeNames.push(seasons[i]);
  }
  
  //除外シートを除外
  for(var i = 0; i < excludeNames.length; i++){
    for(var j = 0; j < sheets.length; j++){
      if(sheets[j].getSheetName() == excludeNames[i]){
        sheets.splice(j, 1);
        break;
      }
    } 
  }
  
  for(var h = 0; h < sheets.length; h++){  
    var resultStartRow = 3;
    var resultStartCol = 2; 
    var resultData = sheets[h].getRange(resultStartRow, resultStartCol, sheets[h].getLastRow(), memberNum).getValues();
    
    //勝利数入力
    for(var i = 0; true; i++){
      var endFlag = true
      for(var j = 0; j < memberNum; j++){
        if(!(resultData[i][j] === "" || resultData[i][j] === "-")){
          endFlag = false;
        }
      }
      
      if(endFlag){
        break;
      }
      
      for(var j = 0; j < memberNum; j++){
        if(resultData[i][j] === "" || resultData[i][j] === "-"){
          continue;
        }
        var selfScore = resultData[i][j];
        
        for(var k = j + 1; k < memberNum; k++){
          if(resultData[i][k] === "" || resultData[i][k] === "-"){
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
  samarySheet.getRange(startRow, startCol, height, width).setFormulas(winRateFormulaArray);
  
}