function copyInvoiceSheet() {
  const attendanceSheet = SpreadsheetApp.getActiveSheet();  // 今月の勤務実績表シート

  const protections = attendanceSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  let isProtected = false;
  for (const protection of protections) {
    if (protection.getDescription() === '勤務実績表確定による保護') {
      isProtected = true;
      break;
    }
  }

  if (!isProtected) {
    console.error("勤務実績が確定していないので請求書の作成を中止");
    return;
  }

  deleteProtectedInvoiceSheets();

  const spreadsheet = attendanceSheet.getParent();
  const invoiceInfoSheet = spreadsheet.getSheetByName("各種情報");
  const invoiceTemplateSpreadsheet = SpreadsheetApp.openById("1VEqArPvib0sIHaDDmVFu1wsnMBp5nFmQwXjx_kK9TCI");
  const invoiceTemplateSheet = invoiceTemplateSpreadsheet.getSheetByName("請求書テンプレート");
  const attendanceSheetName = attendanceSheet.getName();  // 今月の勤務実績表のシート名
  const invoiceSheetName = attendanceSheetName + "_請求書"; // 今月の請求書シートのシート名


  const invoiceBodyRng = "B6:B9"; //請求書の内容の範囲

  if (!invoiceTemplateSheet) {
    throw new Error("請求書テンプレートシートが見つかりませんでした。");
  }

  // 同名のシートが存在する場合はエラーを出力する
  if (spreadsheet.getSheetByName(attendanceSheetName + "_請求書")) {
    SpreadsheetApp.getUi().alert('エラー', '既に該当月の請求書が作成されています', SpreadsheetApp.getUi().ButtonSet.OK);
    throw new Error("同名のシートがすでに存在しています。");
  }

  const accountName = invoiceInfoSheet.getRange("C12").getValue();
  const staffID = invoiceInfoSheet.getRange("C2").getValue();
  const accountDescription = invoiceInfoSheet.getRange("C8:C11").getValues().flat().join(" ");
  const invoiceNo = invoiceInfoSheet.getRange("C18").getValue();
  const companyName = invoiceInfoSheet.getRange("C17").getValue();
  const staffName = invoiceInfoSheet.getRange("C3").getValue();
  const zipCode = invoiceInfoSheet.getRange("C14").getValue();
  const staffAddress1 = invoiceInfoSheet.getRange("C15").getValue();
  const staffAddress2 = invoiceInfoSheet.getRange("C16").getValue();

  //基本情報バリデーション
  if (!accountName || !staffID || !accountDescription || !companyName || !staffName || !zipCode || !staffAddress1 || !staffAddress2) {
    SpreadsheetApp.getUi().alert('エラー', '各種情報の必須項目に空欄があります。各種情報シートの更新を行ってください', SpreadsheetApp.getUi().ButtonSet.OK);
    throw new Error("各種情報シートの情報取得エラー");
  }

  //海外スタッフ処理
  if (invoiceNo == "対象外") {
    return;
  } else if (invoiceNo == "登録しない") {
    invoiceNo = "";
  }

  var invoiceSheet = invoiceTemplateSheet.copyTo(spreadsheet);
  invoiceSheet.setName(invoiceSheetName);

  // 勤務実績表から情報をがさっと取ってくる
  const data = attendanceSheet.getDataRange().getValues();
  let startRow = 15;


  var year = new Date().getFullYear();
  var month = new Date().getMonth() + 1;
  if (month === 13) {
    year++;
    month = 1;
  }
  var lastDay = new Date(year, month, 0);

  var invoiceDay = new Date(year, month - 1, 0);

  var replaceAttedanceSheetName = attendanceSheetName.replace("年", "").replace("月", "");
  var accountNumber = staffID + "-" + replaceAttedanceSheetName;

  const valuesToSetAccountInfo = [[attendanceSheetName + "リモートスタッフ稼働分"],
  [lastDay],
  [accountDescription],
  [accountName],
  ];

  invoiceSheet.getRange(invoiceBodyRng).setValues(valuesToSetAccountInfo);

  const valuesToSetCompanyInfo = [[companyName]];
  invoiceSheet.getRange("H7").setValues(valuesToSetCompanyInfo);

  const valuesToSetInvoiceInfo = [[accountNumber],
  [invoiceDay],
  [""],
  [companyName],
  [zipCode],
  [staffAddress1],
  [staffAddress2],
  [staffID],
  [staffName],
  [invoiceNo]];
  invoiceSheet.getRange("H2:H11").setValues(valuesToSetInvoiceInfo);


  const dataToSet = [];
  const rowsToCopy = [];

  for (let row = 6; row < data.length; row++) {
    if (!data[row][2] && !data[row][3] && !data[row][4] && !data[row][6] && !data[row][7]) {
      break;
    }

    if (data[row][3] === "時間単価") {
      // 時分を時間に変える
      console.log("変換前" + data[row][6]);
      const timeSplit = Utilities.formatDate(data[row][6], "JST", "HH:mm");
      console.log("変換後" + data[row][6]);
      const timeSplits = String(timeSplit).split(/:/,2);
      console.log(timeSplits[0] + timeSplits[1]);
      const vHour = Number(timeSplits[0]);
      const vMinutes = Number(timeSplits[1] / 60);
      const valueTime = Math.floor((vHour + vMinutes)*10) / 10;
      data[row][6] = [valueTime];
    console.log(data[row][6]+ "." +valueTime);
    } else if (data[row][3] === "件数") {
    
    } else if (data[row][3] === "月額固定") {
    
    data[row][7] = 1;
    
    }

    rowsToCopy.push(row);
    console.log(data[row][4]);
  }

  for (let i = 0; i < rowsToCopy.length; i++) {
    const row = rowsToCopy[i];
    if (data[row][3] === "時間単価") {
      const rowData = [data[row][2]+"("+data[row][3]+")",
      "",
      "",
      "",
    data[row][4],
    data[row][6],
    ];
    dataToSet.push(rowData);
    console.log(rowData);
    } else if (data[row][3] === "件数") {
      const rowData = [data[row][2]+"("+data[row][3]+")",
      "",
      "",
      "",
    data[row][4],
    data[row][7],
    ];
    dataToSet.push(rowData);
    console.log(rowData);
    } else if (data[row][3] === "月額固定") {
      const rowData = [data[row][2]+"("+data[row][3]+")",
      "",
      "",
      "",
    data[row][4],
    data[row][7],
    ];
    dataToSet.push(rowData);
    console.log(rowData);
    }
    
  }

  const numRowsToSet = dataToSet.length;
  const numColsToSet = dataToSet[0].length;

  invoiceSheet.getRange(startRow, 1, numRowsToSet, numColsToSet).setValues(dataToSet);
  //最終rangeに書き込みがされていない場合、無限ループ
  var rangechk = invoiceSheet.getRange("A15").getValue();
  while (rangechk == "") {
    rangechk = invoiceSheet.getRange("A15").getValue();
  }

  // invoiceNoが空欄の場合、A28からD32を削除
  if (invoiceNo == "") {
    invoiceSheet.getRange("A28:D32").clear();
    invoiceSheet.getRange("H29").setValue(0);
    Utilities.sleep(1000);
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('請求書を作成しました', '確定した月の請求書シートに移動しますか？', ui.ButtonSet.OK_CANCEL);

  if (response === ui.Button.OK) {
    spreadsheet.setActiveSheet(invoiceSheet);
  }
}

// 請求書シート作成時に過去に発行済みのシートがあったら削除する
function deleteProtectedInvoiceSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    if (sheetName.includes('_請求書')) {
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (const protection of protections) {
        const description = protection.getDescription();
        if (description === '発行済みのため保護されたシート') {
          if (protection.canEdit()) {
            protection.remove();
            console.log(`シート${sheetName}は保護されており提出済みのためシートを削除します`)
            ss.deleteSheet(sheet);
          } else {
            console.log(`シート${sheetName}は保護されており、削除できません。`);
          }
          break;
        }
      }
    }
  }
}


