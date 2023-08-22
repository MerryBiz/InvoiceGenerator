function submitInvoice() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const invoiceSheet = SpreadsheetApp.getActiveSheet();
  const invoiceSheetName = invoiceSheet.getName();

  const staffID = invoiceSheet.getRange("H9").getValue();
  const staffName = invoiceSheet.getRange("H10").getValue();

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