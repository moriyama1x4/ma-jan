function outputNetRate(){
  var targetCol = 29;
  var startTargetRow = 3;
  var startGamesCol = 2;
  var startGamesRow = 3;
  var gamesData = activeSheet.getRange(startGamesRow, startGamesCol, activeSheet.getLastRow(), memberNum).getValues();
  
  //ゲーム数取得
  var gameNum = 0;
  
  for(var i = 0; true; i++){
    var brakeFlag = true;
    for(var j = 0; j < memberNum; j++){
      if(gamesData[i][j]) {
        brakeFlag = false;
      }
    }
    
    if(brakeFlag){
      break;
    }else{
      gameNum++; 
    }  
  }
  
  
  //入力済みの時アラート
  
  
  
  
  //レート出力
  var targetRange = activeSheet.getRange(startTargetRow, targetCol, gameNum, 1);
  var ratesArray = []; //並び替え前のレート
  var randomRatesArray = []; //ランダムら並び替え後
  var lowest = -1;
  var highest = 3;
  var difference = highest - lowest;
  
  ////並び替え前レート作成
  for(var i = 0; i < gameNum; i++){
    ratesArray.push([lowest + (difference * (i / (gameNum - 1)))]);
  }
  
  
  ////並び替え
  for(var i = 0; i < gameNum; i++){
    var spliceIndex = Math.floor((Math.random() * (i + 1)));
    randomRatesArray.splice(spliceIndex, 0, ratesArray[i]);
  }
  
  ////出力
  targetRange.setValues(randomRatesArray);

}