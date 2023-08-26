function submitInvoice() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const invoiceSheet = SpreadsheetApp.getActiveSheet();
  const invoiceSheetName = invoiceSheet.getName();

  const staffID = invoiceSheet.getRange("H9").getValue();
  const staffName = invoiceSheet.getRange("H10").getValue();
  const invoiceNo = invoiceSheet.getRange("H11").getValue();

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
  addOrUpdateRecord(staffID, attendanceSheetName, staffName, invoiceSheet.getRange("B11").getValue(), spreadsheet.getUrl());

  // invoiceSheetを保護する
  const protection = invoiceSheet.protect();
  protection.removeEditors(protection.getEditors());
  protection.addEditor(spreadsheet.getOwner());
  protection.setDescription('発行済みのため保護されたシート');
}
