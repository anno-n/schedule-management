/**
 * 予定を作成する
 */
function createSchedule() {

  // 連携するアカウント
  const gAccount = CalendarApp.getDefaultCalendar().getId();  //デフォルトのカレンダーを取得
  Logger.log("===gAccount:"+gAccount);
  
  // 読み取り範囲（表の始まり行と終わり列）
  const topRow = 5;
  const lastCol = 12;

  // 0始まりで列を指定しておく
  const statusCellNum = 4;
  const dayCellNum = 5;
  const startCellNum = 6;
  const endCellNum = 7;
  const titleCellNum = 10;
  const descriptionCellNum = 11;

  // シートを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Plan");
  Logger.log("===sheet:"+sheet);

  // 予定の最終行を取得
  var lastRow = sheet.getLastRow();
  
  //予定の一覧を取得
  var contents = sheet.getRange(topRow, 1, sheet.getLastRow(), lastCol).getValues();

  // googleカレンダーの取得
  var calender = CalendarApp.getCalendarById(gAccount);

  //順に予定を作成（今回は正しい値が来ることを想定）
  for (i = 0; i <= lastRow - topRow; i++) {

    //「済」っぽいのか、空の場合は飛ばす
    var status = contents[i][statusCellNum];
    if (
      status == "Done" ||
      contents[i][dayCellNum] == ""
    ) {
      continue;
    }

    // 値をセット 日時はフォーマットして保持
    var day = new Date(contents[i][dayCellNum]);
    var startTime = contents[i][startCellNum];
    var endTime = contents[i][endCellNum];
    var title = contents[i][titleCellNum];
    
    // 詳細をセット
    var options = {description: contents[i][descriptionCellNum]};
    Logger.log("===day:"+day+" startTime:"+startTime+" endTime:"+endTime+" title:"+title+" options:"+options);
    
    try {
      // 開始終了が無ければ終日で設定
      if (startTime == '' || endTime == '') {
        //予定を作成
        calender.createAllDayEvent(
          title,
          new Date(day),
          options
        );
        
      // 開始終了時間があれば範囲で設定
      } else {
        // 開始日時をフォーマット
        var startDate = new Date(day);
        startDate.setHours(startTime.getHours())
        startDate.setMinutes(startTime.getMinutes());
        // 終了日時をフォーマット
        var endDate = new Date(day);
        endDate.setHours(endTime.getHours())
        endDate.setMinutes(endTime.getMinutes());
        // 予定を作成
        calender.createEvent(
          title,
          startDate,
          endDate,
          options
        );
      }

      //無事に予定が作成されたら「済」にする
      sheet.getRange(topRow + i, statusCellNum+1).setValue("Done");

    // エラーの場合（今回はログ出力のみ）
    } catch(e) {
      Logger.log(e);
      Browser.msgBox("予定作成失敗"+e);
    }
    
  }
  // ブラウザへ完了通知
  Browser.msgBox("完了");
}
