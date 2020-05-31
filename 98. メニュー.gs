function onOpen() {
  var entries = [
    {
      name : "ゲーム準備",
      functionName : "prepareGame"
    },
    {
      name : "勝率更新",
      functionName : "updateWinRate"
    }
//    ,
//    {
//      name : "レート出力",
//      functionName : "outputNetRate"
//    }
  ];
  book.addMenu("スクリプト実行", entries);
};
