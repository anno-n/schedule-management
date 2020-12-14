/**
 * スプレッドシート表示の際に呼出し
 */
function onOpen() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //スプレッドシートのメニューにカスタムメニュー「計画取込 > 実行」を作成
  var subMenus = [];
  subMenus.push({
    name: "Migrate Plan| 計画取込",
    functionName: "createSchedule"  //実行で呼び出す関数を指定
  });
  subMenus.push({
    name: "Export Sche.| 予定取得",
    functionName: "exportSchedule"  //実行で呼び出す関数を指定
  });
  ss.addMenu("Plan & Schedule| 計画と予定", subMenus);
}
