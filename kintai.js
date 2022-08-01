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
    case "休憩":
      recordBreaking(userName);
      break;
    case "再開":
      recordResuming(userName);
    default:
      return;
  }
}

const columnNumber = new Map([
  ["日付", 1],
  ["出勤時間", 3],
  ["退勤時間", 4],
  ["休憩時間", 5],
  ["最後の休憩時間", 9],
]);

// 現在時刻から日付の文字列を生成
function createDateString(date) {
  return `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`;
}

// 現在時刻から時刻の文字列を生成
function createTimeString(date) {
  return `${date.getHours().toString().padStart(2, "0")}:${date
    .getMinutes()
    .toString()
    .padStart(2, "0")}`;
}

// 「おはよう」の前の「休憩」,「再開」,「おやすみ」の打刻判定
function isBeforeAttendance(lastCellOfDate, dateString) {
  // 当日の出勤時間が記録されているかチェック
  const lastCellOfDateValue = lastCellOfDate.getValue();
  Logger.log(`lastCellOfDate: ${lastCellOfDateValue}`);
  if (lastCellOfDateValue === "" || lastCellOfDateValue === "自動") {
    return true;
  }

  // 出勤と退勤打刻の日付が一致しているかチェック
  const lastCellOfDateValueString = `${lastCellOfDateValue.getFullYear()}/${
    lastCellOfDateValue.getMonth() + 1
  }/${lastCellOfDateValue.getDate()}`;
  Logger.log(`lastCellOfDateValueString: ${lastCellOfDateValueString}`);
  if (dateString !== lastCellOfDateValueString) {
    Logger.log("退勤を打刻する前に出勤してください。");
    return true;
  }
}

function recordAttendance(userName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(userName);

  // 現在時刻から日付セルと開始時刻セルに記録する文字列を生成
  const date = new Date();
  const dateString = createDateString(date);
  Logger.log(`dateString: ${dateString}`);
  const timeString = createTimeString(date);
  Logger.log(`timeString: ${timeString}`);

  // 日付が記録されている最新のセルを特定
  const topCellOfDate = sheet.getRange(1, columnNumber.get("日付"));
  Logger.log(`topCellOfDate: ${topCellOfDate.getA1Notation()}`);
  // 指定位置の先頭行から下方向に検索して最終行を取得。
  const lastCellOfDate = topCellOfDate.getNextDataCell(
    SpreadsheetApp.Direction.DOWN
  );
  Logger.log(`lastCellOfDate: ${lastCellOfDate.getA1Notation()}`);

  // 2重出勤チェック→2回目の出勤打刻は無視する
  const lastCellOfDateValue = lastCellOfDate.getValue();
  Logger.log(`lastCellOfDate: ${lastCellOfDateValue}`);
  if (!(lastCellOfDateValue === "" || lastCellOfDateValue === "自動")) {
    const lastCellOfDateValueString = `${lastCellOfDateValue.getFullYear()}/${
      lastCellOfDateValue.getMonth() + 1
    }/${lastCellOfDateValue.getDate()}`;
    Logger.log(`lastCellOfDateValueString: ${lastCellOfDateValueString}`);
    if (dateString === lastCellOfDateValueString) {
      Logger.log("出勤2回押してル");
      return;
    }
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

  // 現在時刻から日付セルと開始時刻セルに記録する文字列を生成
  const date = new Date();
  const dateString = createDateString(date);
  Logger.log(`dateString: ${dateString}`);
  const timeString = createTimeString(date);
  Logger.log(`timeString: ${timeString}`);

  // 日付が記録されている最新のセルを特定
  const topCellOfDate = sheet.getRange(1, columnNumber.get("日付"));
  Logger.log(`topCellOfDate: ${topCellOfDate.getA1Notation()}`);
  // 指定位置の先頭行から下方向に検索して最終行を取得。
  const lastCellOfDate = topCellOfDate.getNextDataCell(
    SpreadsheetApp.Direction.DOWN
  );
  Logger.log(`lastCellOfDate: ${lastCellOfDate.getA1Notation()}`);

  // 「おはよう」の前に「おやすみ」が打刻されているかチェック
  if (isBeforeAttendance(lastCellOfDate, dateString)) {
    return;
  }

  // 終了時刻セルを特定
  const newCellOfLeavingTime = lastCellOfDate.offset(
    0,
    columnNumber.get("退勤時間") - columnNumber.get("日付")
  );

  newCellOfLeavingTime.setValue(timeString);
}

function recordBreaking(userName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(userName);

  // 現在時刻から日付セルと開始時刻セルに記録する文字列を生成
  const date = new Date();
  const dateString = createDateString(date);
  Logger.log(`dateString: ${dateString}`);
  const timeString = createTimeString(date);
  Logger.log(`timeString: ${timeString}`);

  // 日付が記録されている最新のセルを特定
  const topCellOfDate = sheet.getRange(1, columnNumber.get("日付"));
  Logger.log(`topCellOfDate: ${topCellOfDate.getA1Notation()}`);
  // 指定位置の先頭行から下方向に検索して最終行を取得。
  const lastCellOfDate = topCellOfDate.getNextDataCell(
    SpreadsheetApp.Direction.DOWN
  );
  Logger.log(`lastCellOfDate: ${lastCellOfDate.getA1Notation()}`);

  // 「おはよう」の前に「休憩」が打刻されているかチェック
  if (isBeforeAttendance(lastCellOfDate, dateString)) {
    return;
  }

  // 最後の休憩時間セルを特定
  const lastCellOfBreakingTime = lastCellOfDate.offset(
    0,
    columnNumber.get("最後の休憩時間") - columnNumber.get("日付")
  );
  const lastCellOfBreakingTimeValue = lastCellOfBreakingTime.getValue();
  Logger.log(
    `lastCellOfBreakingTimeValue.getValue: ${lastCellOfBreakingTimeValue}`
  );
  // 初回の休憩
  if (
    lastCellOfBreakingTimeValue === "" ||
    lastCellOfBreakingTimeValue === "最後の休憩時間"
  ) {
    lastCellOfBreakingTime.setValue(timeString);
    Logger.log("今日初回の休憩です。");
    return;
  }
  lastCellOfBreakingTimeValue.setFullYear(date.getFullYear());
  lastCellOfBreakingTimeValue.setMonth(date.getMonth());
  lastCellOfBreakingTimeValue.setDate(date.getDate());

  Logger.log(
    `lastCellOfBreakingTimeValue.getValue: ${lastCellOfBreakingTimeValue}`
  );

  lastCellOfBreakingTime.setValue(timeString);
}

function recordResuming(userName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(userName);
  // 現在時刻から日付セルと開始時刻セルに記録する文字列を生成
  const date = new Date();
  const dateString = createDateString(date);
  Logger.log(`dateString: ${dateString}`);
  const timeString = createTimeString(date);
  Logger.log(`timeString: ${timeString}`);

  // 日付が記録されている最新のセルを特定
  const topCellOfDate = sheet.getRange(1, columnNumber.get("日付"));
  Logger.log(`topCellOfDate: ${topCellOfDate.getA1Notation()}`);
  // 指定位置の先頭行から下方向に検索して最終行を取得。
  const lastCellOfDate = topCellOfDate.getNextDataCell(
    SpreadsheetApp.Direction.DOWN
  );
  Logger.log(`lastCellOfDate: ${lastCellOfDate.getA1Notation()}`);

  // 「おはよう」の前に「再開」が打刻されているかチェック
  if (isBeforeAttendance(lastCellOfDate, dateString)) {
    return;
  }

  // 最後の休憩時間セルを特定
  const lastCellOfBreakingTime = lastCellOfDate.offset(
    0,
    columnNumber.get("最後の休憩時間") - columnNumber.get("日付")
  );
  const lastCellOfBreakingTimeValue = lastCellOfBreakingTime.getValue();
  if (
    lastCellOfBreakingTimeValue === "" ||
    lastCellOfBreakingTimeValue === "最後の休憩時間"
  ) {
    return;
  }
  lastCellOfBreakingTimeValue.setFullYear(date.getFullYear());
  lastCellOfBreakingTimeValue.setMonth(date.getMonth());
  lastCellOfBreakingTimeValue.setDate(date.getDate());

  Logger.log(
    `lastCellOfBreakingTimeValue.getValue: ${lastCellOfBreakingTimeValue}`
  );

  const lastCellOfBreakingTimeValueString = createTimeString(
    lastCellOfBreakingTimeValue
  );
  Logger.log(
    `lastCellOfBreakingTimeValueString: ${lastCellOfBreakingTimeValueString}`
  );

  // 休憩時間のセルを特定
  const newCellOfBreakingTime = lastCellOfDate.offset(
    0,
    columnNumber.get("休憩時間") - columnNumber.get("日付")
  );
  // 休憩時間セルの値を取得
  const newCellOfBreakingTimeValue = newCellOfBreakingTime.getValue();
  Logger.log(`date.getTime(): ${date.getTime()}`);
  Logger.log(
    `lastCellOfBreakingTimeValue.getTime(): ${lastCellOfBreakingTimeValue.getTime()}`
  );
  Logger.log(`date.getTime(): ${date.getTime()}`);
  // 休憩時間セルに累計休憩時間を記録
  newCellOfBreakingTime.setValue(
    newCellOfBreakingTimeValue +
      Math.floor(
        ((date.getTime() - lastCellOfBreakingTimeValue.getTime()) /
          (1000 * 60 * 60)) *
          100
      ) /
        100
  );
}
