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
  var height = memberNum;
  var width = memberNum;
  var samaryStartTargetRow = 13;
  var samaryStartTargetCol = 39;
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
  for(var h = 0; h < seasons.length; h++){
    var seasonSheet = book.getSheetByName(seasons[h]); 
    var seasonStartTargetRow = samaryStartTargetRow;
    var seasonStartTargetCol = 38;
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




/******
半荘毎勝率更新
*******/
function updateGameWinRate() {
  var samaryStartTargetRow = 13;
  var samaryStartTargetCol = 21;
  var samaryStartTargetCell = samarySheet.getRange(samaryStartTargetRow, samaryStartTargetCol);
  var seasonStartTargetRow = samaryStartTargetRow;
  var seasonStartTargetCol = 20;
  var seasonStartTargetCell = samarySheet.getRange(seasonStartTargetRow, seasonStartTargetCol);
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
