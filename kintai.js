function doPost(e) {
  const triggerWord = e.parameter.trigger_word;
  const userName = e.parameter.user_name;
  switch (triggerWord) {
    case "おはよう":
      recordAttendance(userName);
      break;
    case "おやすみ":
      recordLeaving(userName);
      break;
    default:
      return;
  }
}

const columnNumber = new Map([
  ["日付", 1],
  ["出勤時間", 3],
  ["退勤時間", 4],
]);

function recordAttendance(userName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(userName);

  // 現在時刻から日付セルと開始時刻セルに記録する文字列を生成
  const date = new Date();
  const dateString = `${date.getFullYear()}/${
    date.getMonth() + 1
  }/${date.getDate()}`;
  Logger.log(`dateString: ${dateString}`);
  const timeString = `${date.getHours().toString().padStart(2, "0")}:${date
    .getMinutes()
    .toString()
    .padStart(2, "0")}`;
  Logger.log(`timeString: ${timeString}`);

  // 日付が記録されている最新のセルを特定
  const topCellOfDate = sheet.getRange(1, columnNumber.get("日付"));
  Logger.log(`topCellOfDate: ${topCellOfDate.getA1Notation()}`);
  // 指定位置の先頭行から下方向に検索して最終行を取得。
  const lastCellOfDate = topCellOfDate.getNextDataCell(
    SpreadsheetApp.Direction.DOWN
  );
  Logger.log(`lastCellOfDate: ${lastCellOfDate.getA1Notation()}`);

  // 2重出勤チェック→2回目の出勤は無視する
  const lastCellOfDateValue = lastCellOfDate.getValue();
  const lastCellOfDateValueString = `${lastCellOfDateValue.getFullYear()}/${
    lastCellOfDateValue.getMonth() + 1
  }/${lastCellOfDateValue.getDate()}`;
  Logger.log(`lastCellOfDateValueString: ${lastCellOfDateValueString}`);
  if (dateString === lastCellOfDateValueString) {
    Logger.log("出勤2回押してル");
    return;
  }

  // 1行下の(空の)日付セルと開始時刻セルを特定 offsetは指定された行と列だけオフセットされた(進んだ)範囲を返す。
  const newCellOfDate = lastCellOfDate.offset(1, 0);
  Logger.log(`newCellOfDate: ${newCellOfDate.getA1Notation()}`);
  const newCellOfAttendanceTime = newCellOfDate.offset(
    0,
    columnNumber.get("出勤時間") - columnNumber.get("日付")
  );
  Logger.log(
    `newCellOfAttendanceTime: ${newCellOfAttendanceTime.getA1Notation()}`
  );

  newCellOfDate.setValue(dateString);
  newCellOfAttendanceTime.setValue(timeString);
}

function recordLeaving(userName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(userName);

  // 開始時刻が記録されている最新のセルを特定
  const topCellOfAttendanceTime = sheet.getRange(
    1,
    columnNumber.get("出勤時間")
  );
  const lastCellOfAttendanceTime = topCellOfAttendanceTime.getNextDataCell(
    SpreadsheetApp.Direction.DOWN
  );

  // 終了時刻セルを特定
  const newCellOfLeavingTime = lastCellOfAttendanceTime.offset(
    0,
    columnNumber.get("退勤時間") - columnNumber.get("出勤時間")
  );

  const date = new Date();
  const timeString = `${date.getHours().toString().padStart(2, "0")}:${date
    .getMinutes()
    .toString()
    .padStart(2, "0")}`;

  newCellOfLeavingTime.setValue(timeString);
}
