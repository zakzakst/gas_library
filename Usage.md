##マニフェストファイル（appsscript.json）に以下を追記
``` json
"oauthScopes": [
  "https://www.googleapis.com/auth/spreadsheets.currentonly",
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/script.container.ui",
  "https://www.googleapis.com/auth/drive.readonly",
  "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/calendar",
  "https://www.googleapis.com/auth/calendar.readonly",
  "https://www.google.com/calendar/feeds",
  "https://www.googleapis.com/auth/script.external_request"
]
```

##スクリプトファイルに以下を記載
``` javascript
//==============================
// ■共通設定
//==============================
var searchSheetId = '';
var searchKey = '';
var searchCx = '';
var searchParam = "&pws=0&tbm=nws"; // 参考：http://www13.plala.or.jp/bigdata/google.html
var searchCount = 2; // 10倍した数の検索結果が表示される


//==============================
// ■メニュー追加
//==============================
function onOpen() {
//  MyScripts.addMenuInit(['download', 'calendar', 'makeFile', 'search']);
  MyScripts.addMenuInit(['search', 'seo', 'rank']);
}


//==============================
// ■ファイル作成の実行
//==============================
function runMakeFile(folderId, fileType, templateStr, dataObj, fileName, fileDate) {
  MyScripts.makeFile(folderId, fileType, templateStr, dataObj, fileName, fileDate);
}


//==============================
// ■カスタム検索の設定・実行
//==============================
var queryListSearch = [
  '検索文言',
  'search text'
];
function setSearchResultRun() {
  MyScripts.setSearchResult(searchSheetId, searchKey, searchCx, searchParam, queryListSearch, searchCount);
}
function setSearchResultInputRun() {
  MyScripts.setSearchResultInput(searchKey, searchCx, searchParam, searchCount);
}
function initSeachSheetRun() {
  MyScripts.initSeachSheet(queryListSearch);
}


//==============================
// ■ランクチェックの設定・実行
//==============================
var queryListRank = [
  'おもちゃ',
  'プレゼント 子供'
];
function seoRankRun() {
  MyScripts.seoRank(searchSheetId, searchKey, searchCx, searchParam, queryListRank, searchCount);
}
function initRankSheetRun() {
  MyScripts.initRankSheet(queryListRank);
}
```
