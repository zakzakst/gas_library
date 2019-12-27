//==============================
// ■ファイル作成実行メニューを追加
//（※「MyScripts」としてライブラリを登録している場合のみ使用）
//==============================
function addMenuInit(subMenuList) {
  // 関数呼び出し用に各キーのプロパティとして関数を代入
  var menuList = {
    'makeFile': addMenuMakeFile,
    'download': addMenuDownload,
    'calendar': addMenuCalendar,
    'search': addMenuSearch,
    'seo': addMenuSeo,
    'rank': addMenuRank
  };
  // 各サブメニューを追加
  var menu = SpreadsheetApp.getUi().createMenu('追加メニュー');
  for(var i in subMenuList) {
    menuList[subMenuList[i]](menu);
  }
}


//==============================
// ■ファイル作成実行メニューを追加
//==============================
function addMenuMakeFile(menu) {
  var ui = SpreadsheetApp.getUi();
  menu.addSubMenu(ui.createMenu("ファイル作成")
      .addItem("全体", "MyScripts.makeFileActive")
      .addItem("選択範囲", "MyScripts.makeFileSelect")
  ).addToUi();
}


//==============================
// ■データダウンロード実行メニューを追加
//==============================
function addMenuDownload(menu) {
  var ui = SpreadsheetApp.getUi();
  menu.addSubMenu(ui.createMenu("JSONダウンロード")
      .addItem("全体", "MyScripts.downloadJsonActive")
      .addItem("選択範囲", "MyScripts.downloadJsonSelect")
  ).addToUi();
}


//==============================
// ■Googleカレンダー連携メニューを追加
//==============================
function addMenuCalendar(menu) {
  var ui = SpreadsheetApp.getUi();
  menu.addSubMenu(ui.createMenu("Googleカレンダー連携")
      .addItem("イベントを登録", "MyScripts.createCalendarEvent")
      .addItem("データクリア", "MyScripts.clearCalendarSheet")
      .addItem("バックアップを作成", "MyScripts.backUpCalendarSheet")
      .addSeparator()
      .addItem("シートを初期化", "MyScripts.initCalendarSheet")
  ).addToUi();
}


//==============================
// ■Googleカスタム検索メニューを追加
//==============================
function addMenuSearch(menu) {
  var ui = SpreadsheetApp.getUi();
  menu.addSubMenu(ui.createMenu("Googleカスタム検索")
      .addItem("設定した文言を出力", "setSearchResultRun")
      .addItem("文言を入力して出力", "setSearchResultInputRun")
      .addSeparator()
      .addItem("設定した文言のシートを初期化", "initSeachSheetRun")
      .addItem("文言を入力してシートを初期化", "MyScripts.initSeachSheetInput")
  ).addToUi();
}


//==============================
// ■SEOチェックメニューを追加
//==============================
function addMenuSeo(menu) {
  var ui = SpreadsheetApp.getUi();
  menu.addSubMenu(ui.createMenu("SEOチェック")
      .addItem("URLのSEO情報を出力", "MyScripts.seoCheck")
      .addSeparator()
      .addItem("シートを初期化", "MyScripts.initSeoSheet")
  ).addToUi();
}


//==============================
// ■ランク調査メニューを追加
//==============================
function addMenuRank(menu) {
  var ui = SpreadsheetApp.getUi();
  menu.addSubMenu(ui.createMenu("ランク調査")
      .addItem("ランク調査を出力", "seoRankRun")
      .addSeparator()
      .addItem("シートを初期化", "initRankSheetRun")
  ).addToUi();
}
