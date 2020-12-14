/**
 * 予定を出力する
 */
function exportSchedule() {
  var checkSheet = getCheckSheet();
  var salesforceCalendars = getSalesforceCalendars();
  
  // 予定取得実行
  getEvent(checkSheet, salesforceCalendars);
}

function getCheckSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Check");
  return sheet;
}

function getSalesforceCalendars() {
  var calendars = CalendarApp.getCalendarsByName('Salesforce');
  Logger.log('Found %s matching calendars.', calendars.length);
  return calendars[0];// とりあえず1行のみ返す
}

function getEvent(sheet, calendar) {
  // 対象期間を取得
  var settingTopRow = 4; 
  var startSetting = sheet.getRange(settingTopRow, 6).getValues();
  var endSetting = sheet.getRange(++settingTopRow, 6).getValues();
  Logger.log("===startSetting:"+startSetting + " endSetting:"+endSetting);
  var startDate = new Date(startSetting);
  var endDate = new Date(endSetting);
  Logger.log("===targetRange:"+startDate + " endDate:"+endDate);
  
  // 予定の取得
  var events = calendar.getEvents(startDate, endDate);
  Logger.log('Found %s matching events.', events.length);
  
  // 予定一覧の作成
  var topRow = 8; // 出力の開始行
  sheet.getRange(topRow, 5).setValue("No");
  sheet.getRange(topRow, 6).setValue("人 | Member");
  sheet.getRange(topRow, 7).setValue("タイトル | Title");
  sheet.getRange(topRow, 8).setValue("開始日時｜Start");
  sheet.getRange(topRow, 9).setValue("終了日時 | End");
  sheet.getRange(topRow, 10).setValue("合計時間 | Diff Time(min)");
  sheet.getRange(topRow, 11).setValue("所有カレンダID | Default Calendar");
  topRow++;
  var memberName = calendar.getName();

  events.forEach(function(event) {
    var diffDateMinutes = diffDate(event.getStartTime(), event.getEndTime());

    sheet.getRange(topRow, 5).setValue(topRow-8);
    sheet.getRange(topRow, 6).setValue(memberName);
    sheet.getRange(topRow, 7).setValue(event.getTitle());
    sheet.getRange(topRow, 8).setValue(event.getStartTime());
    sheet.getRange(topRow, 9).setValue(event.getEndTime());
    sheet.getRange(topRow, 10).setValue(diffDateMinutes);
    var originalCalendar = CalendarApp.getCalendarById(event.getOriginalCalendarId());
    sheet.getRange(topRow, 11).setValue(originalCalendar.getName());
    topRow++;
  }, memberName)

}

function diffDate(dateStart, dateEnd){
  var diffTime = (dateEnd - dateStart)/(60*1000);// ミリ秒なので分に修正
  console.log(diffTime);
  return diffTime;
}

function formatHHmm(diffDateMinutes){
  return Utilities.formatString("%sh %sm", (Math.floor(diffDateMinutes/60)), (diffDateMinutes % 60) );
}
