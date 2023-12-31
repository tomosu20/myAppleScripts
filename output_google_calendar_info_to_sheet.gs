/**
 * メニューバーに項目を追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("スクリプト実行");
  menu.addItem("カレンダーの予定を出力する", "exportCalendarSchedule");
  menu.addToUi();
}

const MAIL_ADDRESS = "##your email address##";

/**
 * スプレッドシートから指定した日付のシートを取得します。
 * @param {Date} date - シートを取得するための日付。
 * @return {GoogleAppsScript.Spreadsheet.Sheet} - 取得されたシート。
 */
function getSheet(date) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const year = date.getFullYear();
  const sheetName = "GoogleCalendar" + year;
  let inputDataSheet = spreadsheet.getSheetByName(sheetName);
  if (!inputDataSheet) {
    inputDataSheet = spreadsheet.insertSheet(sheetName);
  }
  return inputDataSheet;
}

/**
 * 指定された期間内のカレンダーイベントを取得します。
 * @param {Date} date_start - 期間の開始日。
 * @param {Date} date_end - 期間の終了日。
 * @return {GoogleAppsScript.Calendar.CalendarEvent[]} - 取得されたカレンダーイベントの配列。
 */
function getCalendarEvents(date_start, date_end) {
  const calendar = CalendarApp.getCalendarById(MAIL_ADDRESS);
  return calendar.getEvents(date_start, date_end);
}

/**
 * イベントが除外条件を満たすかどうかを判定します。
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event - 判定するイベント。
 * @return {boolean} - イベントが除外される場合は true、それ以外は false。
 */
function isExclude(event) {
  if (
    event.isAllDayEvent() || // 終日のイベントを除く
    event.getGuestByEmail(MAIL_ADDRESS)?.getStatus() == "no" // 出席していないMTGを除く
  ) {
    return true;
  }

  return false;
}

/**
 * イベントの作業タイプを取得します。
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event - 取得するイベント。
 * @return {string} - イベントの作業タイプ。
 */
function getWorkType(event) {
  const colorNum = event.getColor();
  if (colorNum == "") {
    const title = event.getTitle();
    if (title.match("【外出】") || title.match("休】") || title.match("公休")) {
      return "休み";
    }
    return "MTG";
  }

  switch (colorNum) {
    case "1":
      return "開発";
    case "3":
      return "共通";
    case "4":
      return "個人";
    case "5":
      return "休み";
    case "6":
      return "その他";
    case "7":
      return "保守";
    default:
      return colorNum;
  }
}

/**
 * GoogleCalendarの情報をSpreadSheetに出力
 */
function exportCalendarSchedule() {
  const now = new Date();
  const sheet = getSheet(now);

  const result = [
    ["年月", "日付", "分類", "予定のタイトル", "開始", "終了"], // 項目名を配列の先頭に追加する
  ];

  const date_start = new Date(now.getFullYear(), 0, 1); //年始
  const events = getCalendarEvents(date_start, now);

  events.forEach((event) => {
    if (isExclude(event)) {
      return;
    }

    let begin = event.getStartTime();
    result.push(
      // 結果用の配列にまとめて追加する
      [
        Utilities.formatDate(begin, "JST", "yyyy/MM"), //年月
        Utilities.formatDate(begin, "JST", "yyyy/MM/dd"), //日付
        getWorkType(event), //分類
        event.getTitle(), //予定のタイトル
        begin, //予定の開始日時
        event.getEndTime(), //予定の終了日時
      ]
    );
  });

  const range = sheet.getRange(1, 1, result.length, result[0].length);
  range.setValues(result); // シートに予定を書き込む
}
