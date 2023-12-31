function importCSVtoSheet() {
  // フォルダとファイルの情報
  var folderName = "MoneyForwardMeDataInput";

  // フォルダの取得
  var folder = DriveApp.getFoldersByName(folderName).next();

  var files = folder.getFiles();

  // 正規表現パターン
  var pattern = /^収入・支出詳細_(\d{4})\.csv$/;

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    var match = fileName.match(pattern);

    if (match) {
      // スプレッドシートの取得
      var fileNameWithoutExtension = fileName.replace(/\.csv$/, "");

      var spreadsheetId = "## your spread sheet ID ##";
      var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      // シートの取得（存在しない場合は新しいシートを挿入）
      var sheet = spreadsheet.getSheetByName(fileNameWithoutExtension);
      if (!sheet) {
        sheet = spreadsheet.insertSheet(fileNameWithoutExtension);
      }
      // シートをクリア
      sheet.clear();

      // CSVファイルの内容を取得してシートに書き込み
      var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
      sheet
        .getRange(1, 1, csvData.length, csvData[0].length)
        .setValues(csvData);

      // Logger.log("データのインポートが完了しました。");

      file.setTrashed(true);
    }
  }
}
