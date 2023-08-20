function addOrUpdateRecord(staffId, sheetName, fullName, amount, url) {
  const sheet = SpreadsheetApp.openById("1gAFrlO0lbVfJWprw0-9JQFuKxrD9ayqfXFIMpx68kCA").getSheetByName("発行履歴");
  const dataRange = sheet.getDataRange().getDisplayValues();

  let rowIndex = -1;

  // スタッフIDと対象シート名が重複する行を探す
  for (let i = 0; i < dataRange.length; i++) {
    if (dataRange[i][0] === staffId && dataRange[i][1] === sheetName) {
      rowIndex = i;
      break;
    }
  }

  // 重複する行が見つかった場合、その行を削除
  if (rowIndex !== -1) {
    sheet.deleteRow(rowIndex + 1);
  }

  // 新しいデータを追加
  const timestamp = new Date(); // 現在の日付と時刻を取得
  sheet.appendRow([staffId, sheetName, fullName, amount, timestamp, url]);
}
