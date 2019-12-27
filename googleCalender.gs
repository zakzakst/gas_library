//==============================
// ■シートの内容をカレンダーに登録
//==============================
function createCalendarEvent() {
  var calendar = CalendarApp.getDefaultCalendar();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Googleカレンダー連携');
  var values = sheet.getDataRange().getValues();
  for(var i = 1; i < values.length; i++){
    if(sheet.getRange(i + 1, 7).getValue() == '') {
      var title = values[i][0];
      var startTime = addCalendarTime(new Date(values[i][1]), values[i][2]);
      var endTime = addCalendarTime(new Date(values[i][1]), startTime, values[i][3]);
      var option = {
        description: values[i][4],
        location: values[i][5]
      }
      calendar.createEvent(title, startTime, endTime, option);
      sheet.getRange(i + 1, 7).setValue('登録済');
    }
  }
}


//==============================
// ■日付と時間を組み合わせる
//==============================
function addCalendarTime(date, time, howLong){
  howLong = howLong || new Date(0,0,0,0,0,0);
  date.setHours(time.getHours() + howLong.getHours());
  date.setMinutes(time.getMinutes() + howLong.getMinutes());
  return date;
}


//==============================
// ■シートの入力データをクリア
//==============================
function clearCalendarSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Googleカレンダー連携');
  var nums = getlastRowColumNum('Googleカレンダー連携');
  sheet.getRange(2, 1, nums.row, nums.column).clear();
}


//==============================
// ■カレンダーシートのバックアップを作成
//==============================
function backUpCalendarSheet() {
  backUpSheet('Googleカレンダー連携');
}


//==============================
// ■カレンダー入力用のシートを作成
//==============================
function initCalendarSheet() {
  // 登録情報用のシートを作成
  var initSheet = insertSheet('Googleカレンダー連携', 'confirm');

  // 一行目のセルのスタイルを変更
  initSheet.getRange('A1:G1')
    .setBackground('#666')
    .setFontColor('#fff')
    .setFontWeight('bold');

  // 一行目に見出しを入力
  initSheet.getRange('A1').setValue('イベント名')
    .offset(0, 1).setValue('日付')
    .offset(0, 1).setValue('開始時間')
    .offset(0, 1).setValue('期間')
    .offset(0, 1).setValue('場所')
    .offset(0, 1).setValue('概要')
    .offset(0, 1).setValue('登録状況');

  // 二行目以降の書式を設定
  initSheet.getRange('B2:B').setNumberFormat('yyyy/MM/dd');
  initSheet.getRange('C2:C').setNumberFormat('h":"mm');
  initSheet.getRange('D2:D').setNumberFormat('h":"mm');
}
