//==============================
// ■指定したデータ・HTMLより、ダイアログを表示
//==============================
function showDialog(data, html, title) {
  // HTMLテンプレートを作成
  var dialog = HtmlService.createTemplateFromFile(html);

  // テンプレートにデータを渡す
  dialog.data = data;

  // ダイアログを表示
  SpreadsheetApp.getUi().showModalDialog(dialog.evaluate(), title);
}


//==============================
// ■指定したデータ・HTMLより、サイドバーを表示
//==============================
function showSideBar(data, html, title) {
  // HTMLテンプレートを作成
  var sideBar = HtmlService.createTemplateFromFile(html);

  // テンプレートにデータを渡す
  sideBar.data = data;

  // ダイアログを表示
  SpreadsheetApp.getUi().showSidebar(sideBar.evaluate().setTitle(title));
}


//==============================
// ■アクティブシートの全範囲のデータを取得
//==============================
function getActiveSheetData() {
  // 設定情報からデータのあるスプレッドシートを読み取る
  var spSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // データが入力されている範囲を取得
  var dataRange = spSheet.getDataRange();

  return dataRange
}


//==============================
// ■指定したシートの全範囲のデータを取得
//==============================
function getSheetData(spID, sheetName) {
  // 設定情報からデータのあるスプレッドシートを読み取る
  var spApp = SpreadsheetApp.openById(spID);
  var spSheet = spApp.getSheetByName(sheetName);

  // データが入力されている範囲を取得
  var dataRange = spSheet.getDataRange();

  return dataRange
}


//==============================
// ■データの先頭行をキーとしてオブジェクトを作成
//==============================
function convertObject(data) {
  // 先頭行の配列を取得
  var dataKey = data[0];

  // 二行目以降の配列を取得
  data.shift();
  var dataProps = data;

  // オブジェクトを作成
  var dataObject = [];
  for(var i = 0; i < dataProps.length; i++) {
    var dataTemp = {};
    for(var j = 0; j < dataKey.length; j++) {
      dataTemp[dataKey[j]] = dataProps[i][j];
    }
    dataObject.push(dataTemp);
  }

  return dataObject
}


//==============================
// ■指定したフォルダにファイルを作成
//（※fileType参考：https://developers.google.com/apps-script/reference/base/mime-type）
//==============================
function outputFile(conf) {
  // confの各値を代入
  var folderId = conf.folderId;
  var fileName = conf.fileName;
  var fileContent = conf.fileContent;
  var fileType = conf.fileType;
  var fileDate = conf.fileDate;
  Logger.log('fileDate:' + fileDate);

  // 日付設定がtureの場合、ファイル名末尾に日付を追加
  if(fileDate) {
    var d = new Date();
    var dateStr = '_' + d.getYear() + ('00' + (d.getMonth() + 1)).slice(-2) + ('00' + d.getDate()).slice(-2) + ('00' + d.getHours()).slice(-2) + ('00' + d.getMinutes()).slice(-2);
    fileName += dateStr;
  }

  // 指定したフォルダにファイルを作成
  var folder = DriveApp.getFolderById(folderId);
  folder.createFile(fileName, fileContent, MimeType[fileType])
}


//==============================
// ■名前を指定してシートを挿入
//==============================
function insertSheet(name, insertType) {
  // 指定した名前のシートがすでに存在した場合の処理
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if(targetSheet) {
    if(insertType == 'always') {
      // 常に新規シートを挿入
      deleteSheet(name);
    }
    else if(insertType == 'confirm') {
      // 確認して「OK」の場合に新規シートを挿入
      var confirmRes = Browser.msgBox('シート「' + name + '」はすでに存在します。上書きしますか？', Browser.Buttons.OK_CANCEL);
      Logger.log(confirmRes);
      if(confirmRes == 'cancel') { return; }
      deleteSheet(name);
    }
    else {
      // 常に挿入をキャンセル
      return;
    }
  }

  // 指定した名前でシートを挿入
  var insertItem = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name, 0);
  return insertItem;
}


//==============================
// ■指定した名前のシートを削除
//==============================
function deleteSheet(name) {
  var spSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spSheet.getSheetByName(name);
  if(sheet) {
    spSheet.deleteSheet(sheet);
  }
}


//==============================
// ■指定した名前のシートのコピーを作成
//==============================
function copySheet(name, copyName) {
  var spSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spSheet.getSheetByName(name);
  var targetSheet = spSheet.getSheetByName(copyName);
  if(!sheet || targetSheet) { return; }
  spSheet.insertSheet(copyName, 0, {template: sheet});
}


//==============================
// ■指定した名前のシートのバックアップを作成
//==============================
function backUpSheet(name) {
  var copyName = name + '_BK' + getDateStr().substring(0, 8);
  copySheet(name, copyName);
}


//==============================
// ■アクティブシートのバックアップを作成
//==============================
function backUpSheetActive() {
  var name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName();
  var copyName = name + '_BK' + getDateStr().substring(0, 8);
  copySheet(name, copyName);
}


//==============================
// ■現在の日時を取得
//==============================
function getDateStr() {
  var d = new Date();
  var fullDate = d.getYear()
    + ('00' + (d.getMonth() + 1)).slice(-2)
    + ('00' + d.getDate()).slice(-2)
    + ('00' + d.getHours()).slice(-2)
    + ('00' + d.getMinutes()).slice(-2);
  return fullDate;
}


//==============================
// ■指定したシートの最終行、最終列を取得
//==============================
function getlastRowColumNum(name) {
  var dataRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getDataRange();
  var nums = {
    row: dataRange.getLastRow(),
    column: dataRange.getLastColumn()
  };
  return nums;
}
