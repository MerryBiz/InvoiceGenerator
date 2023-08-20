function copyInvoiceSheet() {
  
  const attendanceSheet = SpreadsheetApp.getActiveSheet();  // 今月の勤務実績表シート
  const spreadsheet = attendanceSheet.getParent();
  const invoiceInfoSheet = spreadsheet.getSheetByName("各種情報");
const basicInfoSheet = spreadsheet.getSheetByName("基本情報");
const invoiceTemplateSheet = spreadsheet.getSheetByName("請求書テンプレート");
const attendanceSheetName = attendanceSheet.getName();  // 今月の勤務実績表のシート名
const invoiceSheetName = attendanceSheetName + "_請求書"; // 今月の請求書シートのシート名


const invoiceBodyRng = "B6:B9"; //請求書の内容の範囲
 
  if (!invoiceTemplateSheet) {
    throw new Error("請求書テンプレートシートが見つかりませんでした。");
  }

  // 同名のシートが存在する場合はエラーを出力する
  if (spreadsheet.getSheetByName(attendanceSheetName + "_請求書")) {
    throw new Error("同名のシートがすでに存在しています。");
  }
  var invoiceSheet = invoiceTemplateSheet.copyTo(spreadsheet);
  invoiceSheet.setName(invoiceSheetName);

  // 勤務実績表から情報をがさっと取ってくる
  const data = attendanceSheet.getDataRange().getValues();
  let startRow = 15;

  const accountName = invoiceInfoSheet.getRange("C5").getValue();
  const staffID = basicInfoSheet.getRange("A2").getValue();
  const accountDescription = invoiceInfoSheet.getRange("C1:C4").getValues().flat().join(" ");
  const invoiceNo = invoiceInfoSheet.getRange("C11").getValue();
  const companyName = invoiceInfoSheet.getRange("C10").getValue();
  const staffName = basicInfoSheet.getRange("B2").getValue();
  const zipCode = invoiceInfoSheet.getRange("C7").getValue();
  const staffAddress1 = invoiceInfoSheet.getRange("C8").getValue();
  const staffAddress2 = invoiceInfoSheet.getRange("C9").getValue();

  var year = new Date().getFullYear();
  var month = new Date().getMonth() + 1;
  if (month === 13) {
    year++;
    month = 1;
  }
  var lastDay = new Date(year, month, 0);

  var invoiceDay = new Date(year, month-1, 0);

  const valuesToSetAccountInfo = [    [attendanceSheetName + "リモートスタッフ稼働分"],
    [lastDay],
    [accountDescription],
    [accountName],
  ];

  invoiceSheet.getRange(invoiceBodyRng).setValues(valuesToSetAccountInfo);

  const valuesToSetCompanyInfo = [[companyName]];
  invoiceSheet.getRange("H7").setValues(valuesToSetCompanyInfo);

  const valuesToSetInvoiceInfo = [[invoiceDay],
  [""],
  [companyName],
  [zipCode],
  [staffAddress1],
  [staffAddress2],
  [staffID],
  [staffName],
  [invoiceNo]];
  invoiceSheet.getRange("H3:H11").setValues(valuesToSetInvoiceInfo);


  const dataToSet = [];
  const rowsToCopy = [];
  
  for (let row = 6; row < data.length; row++) {
    if (!data[row][2] && !data[row][3] && !data[row][4] && !data[row][6] && !data[row][7]) {
      break;
    }

    rowsToCopy.push(row);
  }

  for (let i = 0; i < rowsToCopy.length; i++) {
    const row = rowsToCopy[i];
    const rowData = [      data[row][2],
      "",
      "",
      data[row][3],
      data[row][4],
      data[row][6],
      data[row][7],
    ];
    dataToSet.push(rowData);
  }

  const numRowsToSet = dataToSet.length;
  const numColsToSet = dataToSet[0].length;

  invoiceSheet.getRange(startRow, 1, numRowsToSet, numColsToSet).setValues(dataToSet);
  //最終rangeに書き込みがされていない場合、無限ループ
  var rangechk = invoiceSheet.getRange("A15").getValue();
  while (rangechk == ""){
    rangechk = invoiceSheet.getRange("A15").getValue();
  }
  
  // invoiceNoが空欄の場合、A28からD32を削除
  if (invoiceNo == "") {
  invoiceSheet.getRange("A28:D32").clear();
  invoiceSheet.getRange("H29").setValue(0);
  Utilities.sleep(1000);
  }

  var pdfFileName = staffID + "_" + staffName + "_" + invoiceSheetName;
  saveAndDraftSheetAsPDF(pdfFileName, invoiceSheetName);

  addOrUpdateRecord(staffID, attendanceSheetName, staffName, invoiceSheet.getRange("B11").getValue(), spreadsheet.getUrl());

  // invoiceSheetを保護する
  const protection = invoiceSheet.protect();
  protection.removeEditors(protection.getEditors());
  protection.addEditor(spreadsheet.getOwner());
  protection.setDescription('発行済みのため保護されたシート');

  // 発行済みのリストに登録する
}
