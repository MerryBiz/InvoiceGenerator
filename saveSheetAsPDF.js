function saveSheetAsPDF(pdfFileName,newSheetName) {
  var folderID = "1Re1KsYTIpoIDsz37wHfLxtq9n-pLQhdO"; // 請求書を保存するフォルダのID
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // 勤務実績表
  var ssid = ss.getId(); //勤務実績表のスプレッドシートID
  var sheetid = ss.getSheetByName(newSheetName).getSheetId();   // 請求書シートのシートIDを取得
  Logger.log(sheetid);
  
  createPDF(folderID, ssid, sheetid, pdfFileName);

}

function createPDF(folderID, ssid, sheetid, filename) {
  var saveFolder = DriveApp.getFolderById(folderID);// 保存するフォルダを取得
  var rootFolder = DriveApp.getRootFolder(); // マイドライブ直下を取得
  var url = "https://docs.google.com/spreadsheets/d/SSID/export?".replace("SSID", ssid);
  // PDF作成のオプションを指定
  var opts = {
    exportFormat: "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
    format:       "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
    size:         "A4",     // 用紙サイズの指定 legal / letter / A4
    portrait:     "true",   // true → 縦向き、false → 横向き
    fitw:         "true",   // 幅を用紙に合わせるか
    sheetnames:   "false",  // シート名をPDF上部に表示するか
    printtitle:   "false",  // スプレッドシート名をPDF上部に表示するか
    pagenumbers:  "false",  // ページ番号の有無
    gridlines:    "false",  // グリッドラインの表示有無
    fzr:          "false",  // 固定行の表示有無
    gid:          sheetid   // シートIDを指定 sheetidは引数で取得
  };

  var url_ext = [];
  
  // 上記のoptsのオプション名と値を「=」で繋げて配列url_extに格納
  for( optName in opts ){
    url_ext.push( optName + "=" + opts[optName] );
  }

  // url_extの各要素を「&」で繋げる
  var options = url_ext.join("&");

  // API使用のためのOAuth認証
  var token = ScriptApp.getOAuthToken();

    // PDF作成
    var response = UrlFetchApp.fetch(url + options, {
      headers: {
        'Authorization': 'Bearer ' +  token
      }
    });

    // 
    var blob = response.getBlob().setName(filename + '.pdf');

  //}

  //　PDFを指定したフォルダに保存
  saveFolder.createFile(blob);
  rootFolder.createFile(blob);

  //暫定的に下書き保存ロジックに追加
  createDraft(filename,blob);
}