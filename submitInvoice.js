function submitInvoice() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const invoiceSheet = SpreadsheetApp.getActiveSheet();
  const invoiceSheetName = invoiceSheet.getName();

  const staffID = invoiceSheet.getRange("H9").getValue();
  const accountNumber = invoiceSheet.getRange("H2").getValue();
  const staffName = invoiceSheet.getRange("H10").getValue();
  const staffNameKana = invoiceSheet.getRange("B9").getValue();
  const invoiceNo = invoiceSheet.getRange("H11").getValue();
  const billingDate = invoiceSheet.getRange("H3").getValue();
  const address1 = invoiceSheet.getRange("H7").getValue();
  const address2 = invoiceSheet.getRange("H8").getValue();
  const subtotal = invoiceSheet.getRange("G28").getValue();
  const taxRate = invoiceSheet.getRange("G29").getValue();
  const tax = invoiceSheet.getRange("G30").getValue();
  const description = invoiceSheet.getRange("B6").getValue();
  const paymentDeadline = invoiceSheet.getRange("B7").getValue();
  const bankInfo = invoiceSheet.getRange("B8").getValue();

  if (!invoiceNo) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '確認',
      '登録番号が未記入です。免税事業者として請求を実施しますか？',
      ui.ButtonSet.OK_CANCEL
    );

    // "Cancel"が選択された場合は処理を中断
    if (response === ui.Button.CANCEL) {
      return;
    }
  }

  var pdfFileName = staffID + "_" + staffName + "_" + invoiceSheetName;
  saveAndDraftSheetAsPDF(pdfFileName, invoiceSheetName);

  const attendanceSheetName = invoiceSheetName.replace("_請求書", "");
  console.log(accountNumber, staffID, attendanceSheetName, staffName, staffNameKana, invoiceSheet.getRange("B11").getValue(), spreadsheet.getUrl(), billingDate, invoiceNo, address1, address2, subtotal, taxRate, tax, description, paymentDeadline, bankInfo)
  addOrUpdateRecord(accountNumber, staffID, attendanceSheetName, staffName, staffNameKana, invoiceSheet.getRange("B11").getValue(), spreadsheet.getUrl(), billingDate, invoiceNo, address1, address2, subtotal, taxRate, tax, description, paymentDeadline, bankInfo);

  // invoiceSheetを保護する
  const protection = invoiceSheet.protect();
  protection.removeEditors(protection.getEditors());
  protection.addEditor(spreadsheet.getOwner());
  protection.setDescription('発行済みのため保護されたシート');
}
