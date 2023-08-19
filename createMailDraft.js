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
