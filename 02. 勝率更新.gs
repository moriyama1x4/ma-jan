function updateWinRate(){
  updateDailyWinRate();
  updateGameWinRate();
}




/******
日毎勝率更新
*******/
function updateDailyWinRate() {
  var height = memberNum;
  var width = memberNum;
  var samaryStartTargetRow = 13;
  var samaryStartTargetCol = 39;
  var samaryTargetRange = samarySheet.getRange(samaryStartTargetRow, samaryStartTargetCol, height, width);
  var samaryTimeRow = 20;
  var samaryTimeCol = 20;
  var samaryTimeCell = samarySheet.getRange(samaryTimeRow, samaryTimeCol);
  var sliceFormula = "SPARKLINE({1,0})" //斜線
  var samaryWinLoseArray = [];
  var samaryWinRateFormulaArray = [];
  
  //総合勝敗array作成
  for(var i = 0; i < memberNum; i++){
    samaryWinLoseArray.push([]);
    samaryWinRateFormulaArray.push([]);
    
    for(var j = 0; j < memberNum; j++){
      samaryWinLoseArray[i].push([0, 0]);
    }
  }
  
  //シーズンシート更新
  for(var h = 0; h < seasons.length; h++){
    var seasonSheet = book.getSheetByName(seasons[h]); 
    var seasonStartTargetRow = samaryStartTargetRow;
    var seasonStartTargetCol = 38;
    var seasonTargetRange = seasonSheet.getRange(seasonStartTargetRow, seasonStartTargetCol, height, width);
    var seasonTimeRow = 20;
    var seasonTimeCol = 19;
    var seasonTimeCell = seasonSheet.getRange(seasonTimeRow, seasonTimeCol);
    var seasonWinLoseArray = [];
    var seasonWinRateFormulaArray = [];
    
    //シーズン勝敗array作成
    for(var i = 0; i < memberNum; i++){
      seasonWinLoseArray.push([]);
      seasonWinRateFormulaArray.push([]);
      
      for(var j = 0; j < memberNum; j++){
        seasonWinLoseArray[i].push([0, 0]);
      }
    }
    
    //シーズン実績取得
    var seasonResultStartRow = 5;
    var seasonResultStartCol = 3;
    var seasonResultData = seasonSheet.getRange(seasonResultStartRow, seasonResultStartCol, seasonSheet.getLastRow(), memberNum).getValues();
    
    //勝利数入力
    for(var i = 0; true; i++){
      if(seasonResultData[i][0] == ""){
        break;
      }
      
      for(var j = 0; j < memberNum; j++){
        if(seasonResultData[i][j] == "-"){
          continue;
        }
        var selfScore = seasonResultData[i][j];
        
        for(var k = j + 1; k < memberNum; k++){
          if(seasonResultData[i][k] == "-"){
            continue;
          }
          var taegetScore = seasonResultData[i][k];
          
          if(selfScore > taegetScore){
            seasonWinLoseArray[j][k][0] ++;
            seasonWinLoseArray[k][j][1] ++;
            samaryWinLoseArray[j][k][0] ++;
            samaryWinLoseArray[k][j][1] ++;
          }
          else if(selfScore == taegetScore){
            seasonWinLoseArray[j][k][0] += 0.5;
            seasonWinLoseArray[j][k][1] += 0.5;
            seasonWinLoseArray[k][j][0] += 0.5;
            seasonWinLoseArray[k][j][1] += 0.5;
            samaryWinLoseArray[j][k][0] += 0.5;
            samaryWinLoseArray[j][k][1] += 0.5;
            samaryWinLoseArray[k][j][0] += 0.5;
            samaryWinLoseArray[k][j][1] += 0.5;
          }else if(selfScore < taegetScore){
            seasonWinLoseArray[j][k][1] ++;
            seasonWinLoseArray[k][j][0] ++;
            samaryWinLoseArray[j][k][1] ++;
            samaryWinLoseArray[k][j][0] ++;
          }
          
        }
      }
    }
    
    //シーズン勝率変換
    for(var i = 0; i < memberNum; i++){
      for(var j = 0; j < memberNum; j++){
        if(i == j){
          seasonWinRateFormulaArray[i].push(sliceFormula);
        }else{
          seasonWinRateFormulaArray[i].push("IFERROR(" + seasonWinLoseArray[i][j][0] + "/" + (seasonWinLoseArray[i][j][0] + seasonWinLoseArray[i][j][1]) + ', "-")')
        }
      }
    }
    
    //シーズン勝率入力
    seasonTargetRange.setFormulas(seasonWinRateFormulaArray);
    
    //シーズン更新日時入力
    seasonTimeCell.setValue(updateTime);
  }
  
  //総合勝率変換
  for(var i = 0; i < memberNum; i++){
    for(var j = 0; j < memberNum; j++){
      if(i == j){
        samaryWinRateFormulaArray[i].push(sliceFormula);
      }else{
        samaryWinRateFormulaArray[i].push("IFERROR(" + samaryWinLoseArray[i][j][0] + "/" + (samaryWinLoseArray[i][j][0] + samaryWinLoseArray[i][j][1]) + ', "-")')
      }
    }
  }
  
  //総合勝率入力
  samaryTargetRange.setFormulas(samaryWinRateFormulaArray);
  
  //更新日時入力
  samaryTimeCell.setValue(updateTime);
}




/******
半荘毎勝率更新
*******/
function updateGameWinRate() {
  
  var height = memberNum;
  var width = memberNum;
  var samaryStartTargetRow = 13;
  var samaryStartTargetCol = 21;
  var samaryTargetRange = samarySheet.getRange(samaryStartTargetRow, samaryStartTargetCol, height, width);
  var sliceFormula = "SPARKLINE({1,0})" //斜線
  var samaryWinLoseArray = [];
  var samaryWinRateFormulaArray = [];

  //総合勝敗array作成
  for(var i = 0; i < memberNum; i++){
    samaryWinLoseArray.push([]);
    samaryWinRateFormulaArray.push([]);

    for(var j = 0; j < memberNum; j++){
      samaryWinLoseArray[i].push([0, 0]);
    }
  }

  //シーズンシート更新
  for(var g = 0; g < seasons.length; g++){
    var seasonSheet = book.getSheetByName(seasons[g]); 
    var seasonStartTargetRow = samaryStartTargetRow;
    var seasonStartTargetCol = 20;
    var seasonTargetRange = seasonSheet.getRange(seasonStartTargetRow, seasonStartTargetCol, height, width);
    var seasonWinLoseArray = [];
    var seasonWinRateFormulaArray = [];
    
    //シーズン勝敗array作成
    for(var i = 0; i < memberNum; i++){
      seasonWinLoseArray.push([]);
      seasonWinRateFormulaArray.push([]);
      
      for(var j = 0; j < memberNum; j++){
        seasonWinLoseArray[i].push([0, 0]);
      }
    }
    
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
    
    //シーズン対象シート取得
    var sheets = [];
    for(var i = 0; i < seasonDates.length; i++){
      sheets.push(book.getSheetByName(seasonDates[i]));
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
              seasonWinLoseArray[j][k][0] ++;
              seasonWinLoseArray[k][j][1] ++;
              samaryWinLoseArray[j][k][0] ++;
              samaryWinLoseArray[k][j][1] ++;
            }
            else if(selfScore == taegetScore){
              seasonWinLoseArray[j][k][0] += 0.5;
              seasonWinLoseArray[j][k][1] += 0.5;
              seasonWinLoseArray[k][j][0] += 0.5;
              seasonWinLoseArray[k][j][1] += 0.5;
              samaryWinLoseArray[j][k][0] += 0.5;
              samaryWinLoseArray[j][k][1] += 0.5;
              samaryWinLoseArray[k][j][0] += 0.5;
              samaryWinLoseArray[k][j][1] += 0.5;
            }else if(selfScore < taegetScore){
              seasonWinLoseArray[j][k][1] ++;
              seasonWinLoseArray[k][j][0] ++;
              samaryWinLoseArray[j][k][1] ++;
              samaryWinLoseArray[k][j][0] ++;
            }
            
          }
        }
      }
    }
    
    //シーズン勝率変換
    for(var i = 0; i < memberNum; i++){
      for(var j = 0; j < memberNum; j++){
        if(i == j){
          seasonWinRateFormulaArray[i].push(sliceFormula);
        }else{
          seasonWinRateFormulaArray[i].push("IFERROR(" + seasonWinLoseArray[i][j][0] + "/" + (seasonWinLoseArray[i][j][0] + seasonWinLoseArray[i][j][1]) + ', "-")')
        }
      }
    }
    
    //シーズン勝率入力
    seasonTargetRange.setFormulas(seasonWinRateFormulaArray);
  }
  
    //総合勝率変換
  for(var i = 0; i < memberNum; i++){
    for(var j = 0; j < memberNum; j++){
      if(i == j){
        samaryWinRateFormulaArray[i].push(sliceFormula);
      }else{
        samaryWinRateFormulaArray[i].push("IFERROR(" + samaryWinLoseArray[i][j][0] + "/" + (samaryWinLoseArray[i][j][0] + samaryWinLoseArray[i][j][1]) + ', "-")')
      }
    }
  }
  
  //総合勝率入力
  samaryTargetRange.setFormulas(samaryWinRateFormulaArray);
  
}