export const APPS_SCRIPT = `var SIGNATURE = "WORKTIME_ATTENDANCE_V1";

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (String(params.ping) === "1") return json_({ ok: true, app: SIGNATURE });
  return json_({ ok: true, app: SIGNATURE });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (!body || body.signature !== SIGNATURE) return json_({ ok: false, error: "bad signature" });
    var records = body.records || [];
    var sheet = getSheet_();
    records.forEach(function (rec) { upsert_(sheet, rec); });
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
    sheet.appendRow(["date", "status", "checkIn", "checkOut", "breakMinutes", "permissionMinutes", "permissionType", "grossMinutes", "netMinutes", "targetMinutes", "overtimeMinutes", "updatedAt"]);
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
    new Date().toISOString()
  ];
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
`;
