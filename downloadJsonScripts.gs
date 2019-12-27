//==============================
// ■アクティブシート全体のデータをJSONでダウンロード
//==============================
function downloadJsonActive() {
  // アクティブシートからデータを取得
  var sheetData = getActiveSheetData();

  // 先頭行をキーとしてオブジェクトを作成
  var dataObj = convertObject(sheetData.getValues());

  // データをJSONに変更
  var dataJson = JSON.stringify(dataObj);

  // データを渡してダイアログを表示
  var data = {
    name: SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName(),
    json: dataJson
  };
  showDialog(data, 'downloadJson', 'JSONデータをダウンロード');
}


//==============================
// ■選択範囲のデータをJSONでダウンロード
//==============================
function downloadJsonSelect() {
  // 選択範囲からデータを取得
  var selectData = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange();

  // 先頭行をキーとしてオブジェクトを作成
  var dataObj = convertObject(selectData.getValues());

  // データをJSONに変更
  var dataJson = JSON.stringify(dataObj);

  // データを渡してダイアログを表示
  var data = {
    name: SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName(),
    json: dataJson
  };
  showDialog(data, 'downloadJson', 'JSONデータをダウンロード');
}
