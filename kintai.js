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

function recordBreaking(userName) {
  // const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(userName);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("amahaya0831");

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

// 当日の出勤時間が記録されているかチェック
  const lastCellOfDateValue = lastCellOfDate.getValue();
  Logger.log(`lastCellOfDate: ${lastCellOfDateValue}`);
  if (
    lastCellOfDateValue === "" ||
    lastCellOfDateValue === "自動"
  ) {
    return;
  }

  // 出勤と休憩の日付が一致しているかチェック
  const lastCellOfDateValueString = `${lastCellOfDateValue.getFullYear()}/${
    lastCellOfDateValue.getMonth() + 1
  }/${lastCellOfDateValue.getDate()}`;
    Logger.log(`lastCellOfDateValueString: ${lastCellOfDateValueString}`);
  if (dateString !== lastCellOfDateValueString) {
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
  if (lastCellOfBreakingTimeValue === "" || lastCellOfBreakingTimeValue === "最後の休憩時間") {
    lastCellOfBreakingTime.setValue(timeString);
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
  const dateString = `${date.getFullYear()}/${
    date.getMonth() + 1
  }/${date.getDate()}`;
  Logger.log(`dateString: ${dateString}`);
  const timeString = `${date.getHours().toString().padStart(2, "0")}:${date
    .getMinutes()
    .toString()
    .padStart(2, "0")}`;
  Logger.log(`timeString: ${timeString}`);

  // 開始時刻が記録されている最新のセルを特定
  const topCellOfAttendanceTime = sheet.getRange(
    1,
    columnNumber.get("出勤時間")
  );
  const lastCellOfAttendanceTime = topCellOfAttendanceTime.getNextDataCell(
    SpreadsheetApp.Direction.DOWN
  );

  // 最後の休憩時間セルを特定
  const lastCellOfBreakingTime = lastCellOfAttendanceTime.offset(
    0,
    columnNumber.get("最後の休憩時間") - columnNumber.get("出勤時間")
  );
  const lastCellOfBreakingTimeValue = lastCellOfBreakingTime.getValue();
  lastCellOfBreakingTimeValue.setFullYear(date.getFullYear());
  lastCellOfBreakingTimeValue.setMonth(date.getMonth());
  lastCellOfBreakingTimeValue.setDate(date.getDate());

  Logger.log(
    `lastCellOfBreakingTimeValue.getValue: ${lastCellOfBreakingTimeValue}`
  );

  // 最後の休憩時間が空の場合、何もしない。
  if (lastCellOfBreakingTimeValue === "") {
    return;
  } else {
    const lastCellOfBreakingTimeValueString = `${lastCellOfBreakingTimeValue
      .getHours()
      .toString()
      .padStart(2, "0")}:${lastCellOfBreakingTimeValue
      .getMinutes()
      .toString()
      .padStart(2, "0")}`;
    Logger.log(
      `lastCellOfBreakingTimeValueString: ${lastCellOfBreakingTimeValueString}`
    );

    // 休憩時間のセルを特定
    const newCellOfBreakingTime = lastCellOfAttendanceTime.offset(
      0,
      columnNumber.get("休憩時間") - columnNumber.get("出勤時間")
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
}
