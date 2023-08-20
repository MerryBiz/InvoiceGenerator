function saveAndDraftSheetAsPDF(pdfFileName, newSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // 勤務実績表
  var ssid = ss.getId(); //勤務実績表のスプレッドシートID
  var sheetid = ss.getSheetByName(newSheetName).getSheetId(); // 請求書シートのシートIDを取得
  Logger.log(sheetid);

  var pdfBlob = createPDF(ssid, sheetid, pdfFileName); // PDF作成し、Blobを取得

  savePDF(pdfBlob); // PDFを指定のフォルダおよびマイドライブのルートに保存

  createDraft(pdfFileName, pdfBlob);
}

function createPDF(ssid, sheetid, filename) {
  var url = "https://docs.google.com/spreadsheets/d/SSID/export?".replace("SSID", ssid);
  var opts = {
    exportFormat: "pdf",
    format: "pdf",
    size: "A4",
    portrait: "true",
    fitw: "true",
    sheetnames: "false",
    printtitle: "false",
    pagenumbers: "false",
    gridlines: "false",
    fzr: "false",
    gid: sheetid
  };

  var url_ext = [];
  for (optName in opts) {
    url_ext.push(optName + "=" + opts[optName]);
  }

  var options = url_ext.join("&");
  var token = ScriptApp.getOAuthToken();

  var response = UrlFetchApp.fetch(url + options, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });

  return response.getBlob().setName(filename + '.pdf');
}

function savePDF(pdfBlob) {
  var folderID = "1Re1KsYTIpoIDsz37wHfLxtq9n-pLQhdO"; // 請求書を保存するフォルダのID

  var saveFolder = DriveApp.getFolderById(folderID); // 保存するフォルダを取得
  saveFolder.createFile(pdfBlob);

  var rootFolder = DriveApp.getRootFolder(); // マイドライブ直下を取得
  rootFolder.createFile(pdfBlob);
}

function createDraft(fileName,pdfFile) {
    var subject = "【請求書】"+fileName+"";
  var body = "※ 請求書作成ツールで自動作成されたメールです。\n\n今月の請求書を送付します。\nファイル名："+fileName+"\n\nよろしくお願いします。";
  
  // Gmailの下書きを作成
  GmailApp.createDraft(
    'recipient@example.com', // 宛先のメールアドレス
    subject,
    body,
    {
      attachments: [pdfFile.getAs(MimeType.PDF)]
    }
  );
}
