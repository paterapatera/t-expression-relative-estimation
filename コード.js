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
    .map((sheet) => {
      return {
        name: sheet.getSheetName(),
        values: sheet.getDataRange().getValues(),
      };
    });
  const storyGroups = users
    .map((data) => {
      // ポイントごとにストーリーを振り分け
      // toTrimStorys
      const point1 =
        data.values[1].length < 3
          ? []
          : data.values[1].slice(2).filter((x) => x !== "");
      const point2 =
        data.values[2].length < 3
          ? []
          : data.values[2].slice(2).filter((x) => x !== "");
      const point3 =
        data.values[3].length < 3
          ? []
          : data.values[3].slice(2).filter((x) => x !== "");
      const point5 =
        data.values[4].length < 3
          ? []
          : data.values[4].slice(2).filter((x) => x !== "");
      const point8 =
        data.values[5].length < 3
          ? []
          : data.values[5].slice(2).filter((x) => x !== "");
      const point13 =
        data.values[6].length < 3
          ? []
          : data.values[6].slice(2).filter((x) => x !== "");
      const point21 =
        data.values[7].length < 3
          ? []
          : data.values[7].slice(2).filter((x) => x !== "");
      const point34 =
        data.values[8].length < 3
          ? []
          : data.values[8].slice(2).filter((x) => x !== "");
      const point50 =
        data.values[9].length < 3
          ? []
          : data.values[9].slice(2).filter((x) => x !== "");
      const point100 =
        data.values[10].length < 3
          ? []
          : data.values[10].slice(2).filter((x) => x !== "");
      const pointQ =
        data.values[11].length < 3
          ? []
          : data.values[11].slice(2).filter((x) => x !== "");

      return {
        name: data.name,
        point1,
        point2,
        point3,
        point5,
        point8,
        point13,
        point21,
        point34,
        point50,
        point100,
        pointQ,
      };
    })
    .flatMap((data) => {
      // ストーリーに名前とポイントを持たせて1つの配列にする
      // toStorysWithPoint
      const point1 = data.point1.map((p) => ({
        name: data.name,
        value: p.split("\n")[1],
        key: p.split("\n")[0],
        point: 1,
      }));
      const point2 = data.point2.map((p) => ({
        name: data.name,
        value: p.split("\n")[1],
        key: p.split("\n")[0],
        point: 2,
      }));
      const point3 = data.point3.map((p) => ({
        name: data.name,
        value: p.split("\n")[1],
        key: p.split("\n")[0],
        point: 3,
      }));
      const point5 = data.point5.map((p) => ({
        name: data.name,
        value: p.split("\n")[1],
        key: p.split("\n")[0],
        point: 5,
      }));
      const point8 = data.point8.map((p) => ({
        name: data.name,
        value: p.split("\n")[1],
        key: p.split("\n")[0],
        point: 8,
      }));
      const point13 = data.point13.map((p) => ({
        name: data.name,
        value: p.split("\n")[1],
        key: p.split("\n")[0],
        point: 13,
      }));
      const point21 = data.point21.map((p) => ({
        name: data.name,
        value: p.split("\n")[1],
        key: p.split("\n")[0],
        point: 21,
      }));
      const point34 = data.point34.map((p) => ({
        name: data.name,
        value: p.split("\n")[1],
        key: p.split("\n")[0],
        point: 34,
      }));
      const point50 = data.point50.map((p) => ({
        name: data.name,
        value: p.split("\n")[1],
        key: p.split("\n")[0],
        point: 50,
      }));
      const point100 = data.point100.map((p) => ({
        name: data.name,
        value: p.split("\n")[1],
        key: p.split("\n")[0],
        point: 100,
      }));
      const pointQ = data.pointQ.map((p) => ({
        name: data.name,
        value: p.split("\n")[1],
        key: p.split("\n")[0],
        point: 250,
      }));

      return [].concat(
        point1,
        point2,
        point3,
        point5,
        point8,
        point13,
        point21,
        point34,
        point50,
        point100,
        pointQ
      );
    })
    // groupingByKey
    .reduce((calcValue, current) => {
      // ストーリーのキーごとにグルーピング
      if (calcValue[current.key]) {
        calcValue[current.key].push(current);
      } else {
        calcValue[current.key] = [current];
      }
      return calcValue;
    }, {});

  // TODO 後でちゃんと作る
  const optimizePoint = (v) => {
    return Math.round(v);
  };

  const result = Object.values(storyGroups).map((storys) => {
    const first = storys[0];
    return {
      key: first.key,
      value: first.value,
      point: optimizePoint(
        storys.reduce((c, s) => c + s.point, 0) / storys.length
      ),
    };
  });
  // each
  resuletSheet
    .getRange(2, 2, 2, resuletSheet.getLastColumn())
    .deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  resuletSheet
    .getRange(3, 2, 3, resuletSheet.getLastColumn())
    .deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  resuletSheet
    .getRange(4, 2, 4, resuletSheet.getLastColumn())
    .deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  resuletSheet
    .getRange(5, 2, 5, resuletSheet.getLastColumn())
    .deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  resuletSheet
    .getRange(6, 2, 6, resuletSheet.getLastColumn())
    .deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  resuletSheet
    .getRange(7, 2, 7, resuletSheet.getLastColumn())
    .deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  resuletSheet
    .getRange(8, 2, 8, resuletSheet.getLastColumn())
    .deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  resuletSheet
    .getRange(9, 2, 9, resuletSheet.getLastColumn())
    .deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  resuletSheet
    .getRange(10, 2, 10, resuletSheet.getLastColumn())
    .deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  resuletSheet
    .getRange(11, 2, 11, resuletSheet.getLastColumn())
    .deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  resuletSheet
    .getRange(12, 2, 12, resuletSheet.getLastColumn())
    .deleteCells(SpreadsheetApp.Dimension.COLUMNS);

  // objectで管理
  var count1 = 1;
  var count2 = 1;
  var count3 = 1;
  var count5 = 1;
  var count8 = 1;
  var count13 = 1;
  var count21 = 1;
  var count34 = 1;
  var count50 = 1;
  var count100 = 1;
  var countQ = 1;
  result.forEach((v, i) => {
    var row = 1;
    switch (v.point) {
      case 1:
        // outputResultCell
        row = 2;
        count1++;
        resuletSheet
          .getRange(row, count1)
          .setValue(v.key + ":" + v.point + "\n" + v.value)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment("top");
        break;
      case 2:
        row = 3;
        count2++;
        resuletSheet
          .getRange(row, count2)
          .setValue(v.key + ":" + v.point + "\n" + v.value)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment("top");
        break;
      case 3:
        row = 4;
        count3++;
        resuletSheet
          .getRange(row, count3)
          .setValue(v.key + ":" + v.point + "\n" + v.value)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment("top");
        break;
      case 5:
        row = 5;
        count5++;
        resuletSheet
          .getRange(row, count5)
          .setValue(v.key + ":" + v.point + "\n" + v.value)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment("top");
        break;
      case 8:
        row = 6;
        count8++;
        resuletSheet
          .getRange(row, count8)
          .setValue(v.key + ":" + v.point + "\n" + v.value)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment("top");
        break;
      case 13:
        row = 7;
        count13++;
        resuletSheet
          .getRange(row, count13)
          .setValue(v.key + ":" + v.point + "\n" + v.value)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment("top");
        break;
      case 21:
        row = 8;
        count21++;
        resuletSheet
          .getRange(row, count21)
          .setValue(v.key + ":" + v.point + "\n" + v.value)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment("top");
        break;
      case 34:
        row = 9;
        count34++;
        resuletSheet
          .getRange(row, count34)
          .setValue(v.key + ":" + v.point + "\n" + v.value)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment("top");
        break;
      case 50:
        row = 10;
        count50++;
        resuletSheet
          .getRange(row, count50)
          .setValue(v.key + ":" + v.point + "\n" + v.value)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment("top");
        break;
      case 100:
        row = 11;
        count100++;
        resuletSheet
          .getRange(row, count100)
          .setValue(v.key + ":" + v.point + "\n" + v.value)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment("top");
        break;
      default:
        row = 12;
        countQ++;
        resuletSheet
          .getRange(row, countQ)
          .setValue(v.key + ":" + v.point + "\n" + v.value)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment("top");
    }
  });
}
