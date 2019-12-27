//==============================
// ■シートに記載されているURLより各情報を読み取る
//==============================
function seoCheck() {
  var spSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SEOチェック');
  if(!spSheet) {Browser.msgBox('SEOチェック用のシートが見つかりません'); return;}
  var nums = getlastRowColumNum('SEOチェック');
  for(var i = 1; i < nums.row; i++) {
    if(spSheet.getRange(i + 1, 6).getValue() == '') {
      var url = spSheet.getRange(i + 1, 1).getValue();
      spSheet.getRange(i + 1, 2).setFormula('=IFERROR(TEXTJOIN(" ／ ", TRUE, IMPORTXML("' + url + '","//title")), "-----")')
        .offset(0, 1).setFormula('=IFERROR(TEXTJOIN(" ／ ", TRUE, IMPORTXML("' + url + '","//h1")), "-----")')
        .offset(0, 1).setFormula('=IFERROR(TEXTJOIN(" ／ ", TRUE, IMPORTXML("' + url + '","//meta[@name=\'description\']/@content")), "-----")')
        .offset(0, 1).setFormula('=IFERROR(TEXTJOIN(" ／ ", TRUE, IMPORTXML("' + url + '","//meta[@name=\'keywords\']/@content")), "-----")')
        .offset(0, 1).setValue('チェック済');
      // IMPORTXMLだと読み込みが遅いため、値のみペーストする（※何故かLogger.logを挟まないと空白がコピーされたため、記述している、IMPORTXMLの取得が終わらないうちにコピーしてしまうからか？）
      Logger.log(spSheet.getRange(i + 1, 2, 1, 4).getValues());
      spSheet.getRange(i + 1, 2, 1, 4).copyTo(spSheet.getRange(i + 1, 2, 1, 4), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    }
  }
}


//==============================
// ■SEOチェック用のシートを作成
//==============================
function initSeoSheet() {
  // 登録情報用のシートを作成
  var initSheet = insertSheet('SEOチェック', 'confirm');
  if(!initSheet) {return;}

  // 一行目のセルのスタイルを変更
  initSheet.getRange('A1:F1')
    .setBackground('#666')
    .setFontColor('#fff')
    .setFontWeight('bold');

  // 一行目に見出しを入力
  initSheet.getRange('A1').setValue('URL')
    .offset(0, 1).setValue('タイトル')
    .offset(0, 1).setValue('H1')
    .offset(0, 1).setValue('ディスクリプション')
    .offset(0, 1).setValue('キーワード')
    .offset(0, 1).setValue('チェック状況');
}
