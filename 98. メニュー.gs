function onOpen() {
  var entries = [
    {
      name : "ゲーム準備",
      functionName : "prepareGame"
    },
    {
      name : "勝率更新",
      functionName : "undateWinRate"
    }
  ];
  book.addMenu("スクリプト実行", entries);
};
