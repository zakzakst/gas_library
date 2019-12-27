//==============================
// ■各検索クエリのシートに検索結果を出力
//==============================
function setSearchResult(searchSheetId, searchKey, searchCx, searchParam, queryList, searchCount) {
  var baseUrl = 'https://www.googleapis.com/customsearch/v1' + '?key=' + searchKey + '&cx=' + searchCx + searchParam;
  for(var i = 0; i < queryList.length; i++) {
    // 「設定した文言用」と「文言を指定する用」のスクリプトで条件分岐
    if(searchSheetId == 'active') {
      var spSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var nums = getlastRowColumNum(spSheet.getSheetName());
    } else {
      var spSheet = SpreadsheetApp.openById(searchSheetId).getSheetByName('検索結果_' + queryList[i]);
      var nums = getlastRowColumNum('検索結果_' + queryList[i]);
    }
    // 検索結果を受け取れるのが一度に10件までのため、ループ処理で対応
    for(var j = 0; j < searchCount; j++) {
      // 検索URLより情報を取得
      var searchUrl = baseUrl + '&q=' + queryList[i] + '&start=' + (j * 10 + 1);
      var searchResponse = UrlFetchApp.fetch(searchUrl);
      var searchJson = JSON.parse(searchResponse.getContentText());

      // 検索結果をシートの末尾に出力
      var searchDate = getDateStr();
      for(var n = 0; n < 10; n++) {
        // OGディスクリプションが存在しない場合が多かったため、下記でエラー対応
        var ogDescription;
        try {
          ogDescription = searchJson["items"][n]['pagemap']['metatags'][0]['og:description'];
        } catch(error) {
          ogDescription = '-----';
        }
        spSheet.getRange(nums.row + j * 10 + n + 1, 1).setValue(searchDate)
          .offset(0, 1).setValue(queryList[i])
          .offset(0, 1).setValue(j * 10 + n + 1)
          .offset(0, 1).setValue(searchJson['items'][n]['title'] || '-----')
          .offset(0, 1).setValue(searchJson['items'][n]['link'] || '-----')
          .offset(0, 1).setValue(searchJson['items'][n]['snippet'] || '-----')
          .offset(0, 1).setValue(ogDescription || '-----');
      }
    }
  }
}


//==============================
// ■検索クエリを入力して、現在のシートに検索結果を出力
//==============================
function setSearchResultInput(searchKey, searchCx, searchParam, searchCount) {
  var inputQuery = [Browser.inputBox('検索クエリを入力してください')];
  if(inputQuery == 'cancel') {return;}
  setSearchResult('active', searchKey, searchCx, searchParam, inputQuery, searchCount);
}


//==============================
// ■カスタム検索出力用のシートを作成
//==============================
function initSeachSheet(queryList) {
  for(var i = 0; i < queryList.length; i++) {
    // 登録情報用のシートを作成
    var initSheet = insertSheet('検索結果_' + queryList[i], 'confirm');
    if(!initSheet) {continue;}

    // 一行目のセルのスタイルを変更
    initSheet.getRange('A1:G1')
      .setBackground('#666')
      .setFontColor('#fff')
      .setFontWeight('bold');

    // 一行目に見出しを入力
    initSheet.getRange('A1').setValue('日付')
      .offset(0, 1).setValue('検索クエリ')
      .offset(0, 1).setValue('順位')
      .offset(0, 1).setValue('タイトル')
      .offset(0, 1).setValue('リンク')
      .offset(0, 1).setValue('スニペット')
      .offset(0, 1).setValue('OGディスクリプション');
  }
}


//==============================
// ■シート名を入力して、カスタム検索出力用のシートを作成
//==============================
function initSeachSheetInput() {
  var inputQuery = [Browser.inputBox('シート名を入力してください')];
  if(inputQuery == 'cancel') {return;}
  initSeachSheet(inputQuery);
}
