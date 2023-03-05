function importPBI() {
  const jiraConfigSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設定値");
  const templateSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("テンプレート");
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const jiraConfig = new JIRAConfig(jiraConfigSheet);
  const jira = new JIRA(jiraConfig);
  const config = new Config();

  const response = jira
    .importPBI()
    .issues.map((issue) => issue.key + "\n" + issue.fields.summary);
  sheets
    .filter((sheet) => {
      const sheetName = sheet.getSheetName();
      return !config.excludeSheets.includes(sheetName);
    })
    .forEach((sheet) => {
      sheet.deleteRows(1, 12);
      templateSheet.getRange("A1:B12").copyTo(sheet.getRange("A1:B12"));
      response.forEach((v, i) => {
        sheet
          .getRange(1, i + 3)
          .setValue(v)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment("top");
      });
    });
}

function calculateStoryPoint() {
  const resuletSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("推定結果");
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const config = new Config();

  const users = sheets
    .filter((sheet) => {
      const sheetName = sheet.getSheetName();
      return !config.excludeSheets.includes(sheetName);
    })
    .map((sheet) => ({
      name: sheet.getSheetName(),
      values: sheet.getDataRange().getValues(),
    }));
  const storyGroups = users
    .map((data) => ({
      // ユーザー単位で倍率ごとに仕分ける
      name: data.name,
      point0_1: StoryProc.toTrimStorys(data, 1),
      point0_3: StoryProc.toTrimStorys(data, 2),
      point0_6: StoryProc.toTrimStorys(data, 3),
      point1: StoryProc.toTrimStorys(data, 4),
      point1_5: StoryProc.toTrimStorys(data, 5),
      point2: StoryProc.toTrimStorys(data, 6),
      point4: StoryProc.toTrimStorys(data, 7),
      point6: StoryProc.toTrimStorys(data, 8),
      point10: StoryProc.toTrimStorys(data, 9),
      point20: StoryProc.toTrimStorys(data, 10),
      pointQ: StoryProc.toTrimStorys(data, 11),
    }))
    .flatMap((data) =>
      // 倍率のポイント対応
      [].concat(
        data.point0_1.map(StoryProc.toStorysWithPoint(data, 1)),
        data.point0_3.map(StoryProc.toStorysWithPoint(data, 2)),
        data.point0_6.map(StoryProc.toStorysWithPoint(data, 3)),
        data.point1.map(StoryProc.toStorysWithPoint(data, 5)),
        data.point1_5.map(StoryProc.toStorysWithPoint(data, 8)),
        data.point2.map(StoryProc.toStorysWithPoint(data, 13)),
        data.point4.map(StoryProc.toStorysWithPoint(data, 21)),
        data.point6.map(StoryProc.toStorysWithPoint(data, 34)),
        data.point10.map(StoryProc.toStorysWithPoint(data, 50)),
        data.point20.map(StoryProc.toStorysWithPoint(data, 100)),
        data.pointQ.map(StoryProc.toStorysWithPoint(data, 250))
      )
    )
    .reduce(StoryProc.groupingByKey, {});

  const result = Object.values(storyGroups).map((storys) => {
    const first = storys[0];
    return {
      key: first.key,
      value: first.value,
      point: StoryProc.optimizePoint(
        storys.reduce((c, s) => c + s.point, 0) / storys.length
      ),
    };
  });

  [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].forEach((i) =>
    resuletSheet
      .getRange(i, 2, i, resuletSheet.getLastColumn())
      .deleteCells(SpreadsheetApp.Dimension.COLUMNS)
  );

  const initCount = 2;
  const counts = {
    _1: initCount,
    _2: initCount,
    _3: initCount,
    _5: initCount,
    _8: initCount,
    _13: initCount,
    _21: initCount,
    _34: initCount,
    _50: initCount,
    _100: initCount,
    _Q: initCount,
  };
  result.forEach((v) => {
    var [row, column] = StoryProc.calcRange(v, counts);
    StoryProc.outputResultCell(resuletSheet, row, column, v);
  });
}
