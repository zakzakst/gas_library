//==============================
// ■各検索クエリのシートにランク調査を出力
//==============================
function seoRank(searchSheetId, searchKey, searchCx, searchParam, queryList, searchCount) {
  var baseUrl = 'https://www.googleapis.com/customsearch/v1' + '?key=' + searchKey + '&cx=' + searchCx + searchParam;
  for(var i = 0; i < queryList.length; i++) {
    var spSheet = SpreadsheetApp.openById(searchSheetId).getSheetByName('ランク調査_' + queryList[i]);
    if(!spSheet) {Browser.msgBox('ランク調査_' + queryList[i] + 'シートが見つかりません'); return;}
    var nums = getlastRowColumNum('ランク調査_' + queryList[i]);

    // ドメインのリストを代入
    var domainList = spSheet.getRange(2, 2, 1, nums.column - 1).getValues()[0];
    var domainResult = {};
    for(var l = 0; l < domainList.length; l++) {
      domainResult[domainList[l]] = [];
    }

    // 調査した日付をA列に出力
    var searchDate = getDateStr();
    spSheet.getRange(nums.row + 1, 1).setValue(searchDate);

    // 検索結果を受け取れるのが一度に10件までのため、ループ処理で対応
    for(var j = 0; j < searchCount; j++) {
      // 検索URLより情報を取得
      var searchUrl = baseUrl + '&q=' + queryList[i] + '&start=' + (j * 10 + 1);
      var searchResponse = UrlFetchApp.fetch(searchUrl);
      var searchJson = JSON.parse(searchResponse.getContentText());

      // 検索結果をシートの末尾に出力
      for(var n = 0; n < 10; n++) {
        var link = searchJson['items'][n]['link'];
        for(var k = 0; k < domainList.length; k++) {
          if(link == domainList[k]) {
          // URLの完全一致ではなく「含まれる」としたい場合は下記のif文を使用（※ドメインでの調査をしたい時など）
          // if(link.indexOf(domainList[k]) != -1) {
            domainResult[domainList[k]].push(j * 10 + n + 1);
          }
        }
      }
    }
    for(var m = 0; m < domainList.length; m++) {
      spSheet.getRange(nums.row + 1, 1).offset(0, 1 + m).setValue(domainResult[domainList[m]].join('、') || '--');
    }
  }
}


//==============================
// ■ランク調査出力用のシートを作成
//==============================
function initRankSheet(queryList) {
  for(var i = 0; i < queryList.length; i++) {
    // 登録情報用のシートを作成
    var initSheet = insertSheet('ランク調査_' + queryList[i], 'confirm');
    if(!initSheet) {continue;}

    // 一行目のセルのスタイルを変更
    initSheet.getRange('A1:C2')
      .setBackground('#666')
      .setFontColor('#fff')
      .setFontWeight('bold');

    // 一行目に見出しを入力
    initSheet.getRange('A1').setValue(queryList[i])
      .offset(1, 0).setValue('URL');
  }
}
