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

// TODO 後でちゃんと作る
StoryProc.optimizePoint = (v) => {
  return Math.round(v);
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
