class StoryProc {}

StoryProc.toTrimStorys = (data, i) => {
  return data.values[i].length < 3
    ? []
    : data.values[i].slice(2).filter((x) => x !== "");
};

StoryProc.toStorysWithPoint = (data, i) => {
  return (p) => ({
    name: data.name,
    value: p.split("\n")[1],
    key: p.split("\n")[0],
    point: i,
  });
};

StoryProc.groupingByKey = (calcValue, current) => {
  if (calcValue[current.key]) {
    calcValue[current.key].push(current);
  } else {
    calcValue[current.key] = [current];
  }
  return calcValue;
};

StoryProc.optimizePoint = (point) => {
  condition = (v1, v2) => point < (v2 - v1) / 3 + v1;
  if (condition(1, 2)) return 1;
  else if (condition(2, 3)) return 2;
  else if (condition(3, 5)) return 3;
  else if (condition(5, 8)) return 5;
  else if (condition(8, 13)) return 8;
  else if (condition(13, 21)) return 13;
  else if (condition(21, 34)) return 21;
  else if (condition(34, 50)) return 34;
  else if (condition(50, 100)) return 50;
  else if (condition(100, 250)) return 100;
  else return 250;
};

StoryProc.outputResultCell = (resuletSheet, row, count, story) => {
  resuletSheet
    .getRange(row, count)
    .setValue(story.key + "\n" + story.value)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setVerticalAlignment("top");
};

StoryProc.calcRange = (story, counts) => {
  switch (story.point) {
    case 1:
      return [2, counts._1++];
    case 2:
      return [3, counts._2++];
    case 3:
      return [4, counts._3++];
    case 5:
      return [5, counts._5++];
    case 8:
      return [6, counts._8++];
    case 13:
      return [7, counts._13++];
    case 21:
      return [8, counts._21++];
    case 34:
      return [9, counts._34++];
    case 50:
      return [10, counts._50++];
    case 100:
      return [11, counts._100++];
    default:
      return [12, counts._Q++];
  }
};
