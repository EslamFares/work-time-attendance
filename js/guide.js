export const APPS_SCRIPT = `var SIGNATURE = "WORKTIME_ATTENDANCE_V1";

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (String(params.ping) === "1") return json_({ ok: true, app: SIGNATURE });
  if (String(params.pull) === "1") {
    try {
      return json_({ ok: true, app: SIGNATURE, records: pullRecords_(), settings: pullSettings_() });
    } catch (err) {
      return json_({ ok: false, error: String(err) });
    }
  }
  return json_({ ok: true, app: SIGNATURE });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (!body || body.signature !== SIGNATURE) return json_({ ok: false, error: "bad signature" });
    var records = body.records || [];
    var sheet = getSheet_();
    records.forEach(function (rec) { upsert_(sheet, rec); });
    if (body.settings && body.settings.length) writeSettings_(body.settings);
    return json_({ ok: true, app: SIGNATURE, count: records.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("سجل الحضور");
  if (!sheet) {
    sheet = ss.insertSheet("سجل الحضور");
    sheet.appendRow(["date", "status", "checkIn", "checkOut", "breakMinutes", "permissionMinutes", "permissionType", "grossMinutes", "netMinutes", "targetMinutes", "overtimeMinutes", "updatedAt", "checkedOut"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function pad_(n) {
  n = String(n);
  return n.length < 2 ? "0" + n : n;
}

function westernDigits_(text) {
  var map = {"٠":"0","١":"1","٢":"2","٣":"3","٤":"4","٥":"5","٦":"6","٧":"7","٨":"8","٩":"9","۰":"0","۱":"1","۲":"2","۳":"3","۴":"4","۵":"5","۶":"6","۷":"7","۸":"8","۹":"9"};
  return String(text == null ? "" : text).replace(/[٠-٩۰-۹]/g, function (ch) { return map[ch] || ch; });
}

function dateKey_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  if (typeof value === "number" && isFinite(value) && value > 20000 && value < 80000) {
    var utc = new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000);
    return Utilities.formatDate(utc, "UTC", "yyyy-MM-dd");
  }
  var text = westernDigits_(value).replace(/\\s+/g, " ").trim();
  if (!text) return "";
  var iso = text.match(/(\\d{4})-(\\d{1,2})-(\\d{1,2})/);
  if (iso) return iso[1] + "-" + pad_(iso[2]) + "-" + pad_(iso[3]);
  var ymd = text.match(/(\\d{4})[\\/.](\\d{1,2})[\\/.](\\d{1,2})/);
  if (ymd) return ymd[1] + "-" + pad_(ymd[2]) + "-" + pad_(ymd[3]);
  var dmy = text.match(/(\\d{1,2})[\\/.\\-](\\d{1,2})[\\/.\\-](\\d{4})/);
  if (dmy) {
    var a = Number(dmy[1]);
    var b = Number(dmy[2]);
    var day = a;
    var month = b;
    if (a <= 12 && b > 12) { month = a; day = b; }
    return dmy[3] + "-" + pad_(month) + "-" + pad_(day);
  }
  var parsed = new Date(text);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return "";
}

function timeText_(value) {
  if (value == null || value === "") return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  }
  if (typeof value === "number" && isFinite(value)) {
    if (Math.abs(value) > 1.5) return "";
    var frac = value % 1;
    if (frac < 0) frac += 1;
    var mins = Math.round(frac * 1440) % 1440;
    return pad_(Math.floor(mins / 60)) + ":" + pad_(mins % 60);
  }
  var text = westernDigits_(value).trim();
  var match = text.match(/(\\d{1,2}):(\\d{2})/);
  if (!match) return "";
  var hours = Number(match[1]);
  var minutes = Number(match[2]);
  if (hours > 23 || minutes > 59) return "";
  return pad_(hours) + ":" + pad_(minutes);
}

function cellDateKey_(raw, shown) {
  var fromRaw = dateKey_(raw);
  if (/^\\d{4}-\\d{2}-\\d{2}$/.test(fromRaw)) return fromRaw;
  var fromShown = dateKey_(shown);
  if (/^\\d{4}-\\d{2}-\\d{2}$/.test(fromShown)) return fromShown;
  return fromRaw || fromShown || "";
}

function writeRow_(sheet, rowIndex, rec) {
  var row = row_(rec);
  var date = dateKey_(rec.date);
  row[0] = date;
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  var dateCell = sheet.getRange(rowIndex, 1);
  dateCell.setNumberFormat("@");
  dateCell.setValue(date);
}

function upsert_(sheet, rec) {
  var date = dateKey_(rec.date);
  if (!date) return;
  var last = sheet.getLastRow();
  var matchRow = 0;
  var dupes = [];
  if (last >= 2) {
    var range = sheet.getRange(2, 1, last - 1, 1);
    var values = range.getValues();
    var shown = range.getDisplayValues();
    for (var i = 0; i < values.length; i++) {
      if (cellDateKey_(values[i][0], shown[i][0]) !== date) continue;
      if (!matchRow) matchRow = i + 2;
      else dupes.push(i + 2);
    }
  }
  if (matchRow) writeRow_(sheet, matchRow, rec);
  else {
    sheet.appendRow(row_(rec));
    writeRow_(sheet, sheet.getLastRow(), rec);
  }
  for (var d = dupes.length - 1; d >= 0; d--) sheet.deleteRow(dupes[d]);
}

function row_(rec) {
  return [
    rec.date || "",
    rec.status || "",
    rec.checkIn || "",
    rec.checkOut || "",
    Number(rec.breakMinutes) || 0,
    Number(rec.permissionMinutes) || 0,
    rec.permissionType || "",
    Number(rec.grossMinutes) || 0,
    Number(rec.netMinutes) || 0,
    Number(rec.targetMinutes) || 0,
    Number(rec.overtimeMinutes) || 0,
    new Date().toISOString(),
    rec.checkedOut ? "1" : ""
  ];
}

function pullRecords_() {
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var range = sheet.getRange(2, 1, last - 1, 13);
  var values = range.getValues();
  var shown = range.getDisplayValues();
  var records = [];
  for (var i = 0; i < values.length; i++) {
    var date = cellDateKey_(values[i][0], shown[i][0]);
    if (!date) continue;
    var checkIn = timeText_(values[i][2]) || timeText_(shown[i][2]);
    var checkOut = timeText_(values[i][3]) || timeText_(shown[i][3]);
    records.push({
      date: date,
      status: String(shown[i][1] || values[i][1] || ""),
      checkIn: checkIn,
      checkOut: checkOut,
      breakMinutes: Number(values[i][4]) || 0,
      permissionMinutes: Number(values[i][5]) || 0,
      permissionType: String(values[i][6] || ""),
      grossMinutes: Number(values[i][7]) || 0,
      netMinutes: Number(values[i][8]) || 0,
      targetMinutes: Number(values[i][9]) || 0,
      overtimeMinutes: Number(values[i][10]) || 0,
      checkedOut: String(values[i][12] || shown[i][12] || "") === "1" ? "1" : ""
    });
  }
  return records;
}

function getSettingsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("الإعدادات");
  if (!sheet) {
    sheet = ss.insertSheet("الإعدادات");
    sheet.getRange(1, 1, 1, 3).setValues([["البيان", "القيمة", "key"]]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
    return sheet;
  }
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, 3).setValues([["البيان", "القيمة", "key"]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function writeSettings_(rows) {
  var sheet = getSettingsSheet_();
  var map = {};
  var last = sheet.getLastRow();
  if (last >= 2) {
    var existing = sheet.getRange(2, 3, last - 1, 1).getDisplayValues();
    for (var i = 0; i < existing.length; i++) {
      var found = String(existing[i][0] || "").trim();
      if (found) map[found] = i + 2;
    }
  }
  for (var r = 0; r < rows.length; r++) {
    var item = rows[r] || {};
    var key = String(item.key || "").trim();
    if (!key || key === "key") continue;
    var value = item.value == null ? "" : String(item.value);
    var label = item.label == null ? "" : String(item.label);
    var rowIndex = map[key];
    if (!rowIndex) {
      sheet.appendRow([label, "", key]);
      rowIndex = sheet.getLastRow();
      map[key] = rowIndex;
    } else if (label) {
      sheet.getRange(rowIndex, 1).setValue(label);
    }
    var keyCell = sheet.getRange(rowIndex, 3);
    keyCell.setNumberFormat("@");
    keyCell.setValue(key);
    var valueCell = sheet.getRange(rowIndex, 2);
    valueCell.setNumberFormat("@");
    valueCell.setValue(value);
  }
}

function pullSettings_() {
  var sheet = getSettingsSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var shown = sheet.getRange(2, 1, last - 1, 3).getDisplayValues();
  var rows = [];
  for (var i = 0; i < shown.length; i++) {
    var key = String(shown[i][2] || "").trim();
    if (!key || key === "key") continue;
    rows.push({ key: key, value: shown[i][1] == null ? "" : String(shown[i][1]) });
  }
  return rows;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
`;
