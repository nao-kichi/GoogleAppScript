// 期限過ぎたものを削除できるシート
function myFunctionWMib() {
  let ss = SpreadsheetApp.getActiveSpreadsheet()

  // 現在のシートを取得
  let sheet = ss.getActiveSheet()

  // シートのA列の最終行の値を取得
  let lastRowNum = sheet.getRange("A:A").getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow()

  // 実行時の日付
  let nowDate = new Date()
  // 2ヶ月前の日付
  let towMonthsAgoDate = new Date()
  towMonthsAgoDate.setMonth(nowDate.getMonth() - 2)

  // 削除カウント
  let deleteRowCount = 0

  // シートのA列の入力値を取得し、配列化
  let values = sheet.getRange(`A1:A${lastRowNum}`).getValues().flat()

  // A列データの繰り返し
  values.forEach(function(v, index) {
    if (v < towMonthsAgoDate) { // ２ヶ月前の行の場合
      // 行を削除し、カウントを増やす
      sheet.deleteRow(index + 1 - deleteRowCount)
      deleteRowCount++
    }
  })

  // 実行時の日付確認用
  sheet.getRange(`C${lastRowNum + 2 - deleteRowCount}`).setValue(nowDate.toISOString())
  sheet.getRange(`C${lastRowNum + 3 - deleteRowCount}`).setValue(towMonthsAgoDate.toISOString())
}