const SHEET_NAME = "classes";

function doGet(e) {
  const callback = e.parameter.callback || "callback";
  const action = e.parameter.action || "list";

  try {
    const sheet = getSheet();
    if (action === "save") {
      const data = JSON.parse(e.parameter.payload || "{}");
      saveClass(sheet, data);
      return jsonp(callback, { ok: true });
    }
    if (action === "delete") {
      deleteClass(sheet, e.parameter.className || "");
      return jsonp(callback, { ok: true });
    }
    return jsonp(callback, { ok: true, classes: listClasses(sheet) });
  } catch (error) {
    return jsonp(callback, { ok: false, error: String(error.message || error) });
  }
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.appendRow(["className", "title", "classAverage", "courseAverage", "scores", "averages", "distributions", "highestScore", "revealMode", "updatedAt"]);
  }
  return sheet;
}

function listClasses(sheet) {
  const values = sheet.getDataRange().getValues();
  const classes = {};
  for (let row = 1; row < values.length; row += 1) {
    const [className, title, classAverage, courseAverage, scores, averagesJson, distributionsJson, highestScore, revealMode] = values[row];
    if (!className) continue;
    const averages = parseJsonArray(averagesJson) || [
      { label: "クラス平均", value: String(classAverage || "") },
      { label: "コース平均", value: String(courseAverage || "") }
    ];
    const distributions = parseJsonArray(distributionsJson) || [
      { label: "クラス度数分布", scores: String(scores || "") }
    ];
    classes[className] = {
      className: String(className),
      title: String(title || ""),
      classAverage: String(classAverage || ""),
      courseAverage: String(courseAverage || ""),
      scores: String(scores || ""),
      highestScore: String(highestScore || ""),
      revealMode: String(revealMode || "line"),
      averages,
      distributions
    };
  }
  return classes;
}

function saveClass(sheet, data) {
  if (!data.className) throw new Error("className is required");
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const row = findClassRow(sheet, data.className);
    const values = [
      data.className,
      data.title || "",
      data.classAverage || "",
      data.courseAverage || "",
      data.scores || "",
      JSON.stringify(data.averages || []),
      JSON.stringify(data.distributions || []),
      data.highestScore || "",
      data.revealMode || "line",
      new Date()
    ];
    if (row) {
      sheet.getRange(row, 1, 1, values.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
  } finally {
    lock.releaseLock();
  }
}

function parseJsonArray(value) {
  if (!value) return null;
  try {
    const parsed = JSON.parse(String(value));
    return Array.isArray(parsed) ? parsed : null;
  } catch (error) {
    return null;
  }
}

function deleteClass(sheet, className) {
  if (!className) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const row = findClassRow(sheet, className);
    if (row) sheet.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
}

function findClassRow(sheet, className) {
  const values = sheet.getDataRange().getValues();
  for (let row = 1; row < values.length; row += 1) {
    if (String(values[row][0]) === String(className)) return row + 1;
  }
  return null;
}

function jsonp(callback, data) {
  const safeCallback = String(callback).replace(/[^\w.$]/g, "");
  return ContentService
    .createTextOutput(`${safeCallback}(${JSON.stringify(data)});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
