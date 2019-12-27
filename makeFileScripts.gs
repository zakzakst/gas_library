//==============================
// ■アクティブシート全体のデータからファイルを作成
//==============================
function makeFileActive() {
  // データを渡してサイドバーを表示
  var data = {
    name: SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName(),
    dataRange: 'all'
  };
  showSideBar(data, 'makeFile', 'ファイル作成（全体）');
}


//==============================
// ■選択範囲のデータからファイルを作成
//==============================
function makeFileSelect() {
  // データを渡してサイドバーを表示
  var data = {
    name: SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName(),
    dataRange: 'select'
  };
  showSideBar(data, 'makeFile', 'ファイル作成（選択範囲）');
}


//==============================
// ■テンプレートにデータを渡してファイルを作成
//==============================
function makeFile(folderId, fileType, templateStr, dataRange, fileName, fileDate) {
  // 使用するデータを取得
  var data;
  var dataObj;
  if(dataRange == 'select') {
    // 選択範囲のデータを使用する場合
    data = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange();
    dataObj = convertObject(data.getValues());
  } else {
    // アクティブシート全体のデータを使用する場合
    data = getActiveSheetData();
    dataObj = convertObject(data.getValues());
  }

  // 引数用のオブジェクトに値を代入
  var conf = {};
  conf.folderId = folderId;
  conf.fileType = fileType;
  conf.fileDate = fileDate;
  var templateStr = templateStr;

  // 各行ごとの値でファイルを作成
  for(var i = 0; i < dataObj.length; i++) {
    // 各値を初期化
    conf.fileName = '';
    conf.fileContent = '';

    // ファイル名を代入
    if(fileName) {
      conf.fileName = dataObj[i]['fileName'];
    } else {
      conf.fileName = 'ファイル-' + ('00' + (i + 1)).slice(-2);
    }

    // テンプレート文字列の対象箇所を置換
    var str = templateStr;
    for(var n in dataObj[i]) {
      str = str.replace(new RegExp('{{{' + n + '}}}', 'g'), dataObj[i][n]);
    }
    conf.fileContent = str;

    // ファイルを作成
    outputFile(conf);
  }
}
