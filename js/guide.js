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

function dateKey_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  if (typeof value === "number" && isFinite(value)) {
    var utc = new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000);
    return Utilities.formatDate(utc, "UTC", "yyyy-MM-dd");
  }
  var text = String(value == null ? "" : value).trim();
  var iso = text.match(/(\\d{4})-(\\d{2})-(\\d{2})/);
  if (iso) return iso[1] + "-" + iso[2] + "-" + iso[3];
  return text;
}

function writeRow_(sheet, rowIndex, rec) {
  var row = row_(rec);
  sheet.getRange(rowIndex, 1).setNumberFormat("@");
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  sheet.getRange(rowIndex, 1).setNumberFormat("@");
  sheet.getRange(rowIndex, 1).setValue(String(rec.date || ""));
}

function upsert_(sheet, rec) {
  var date = String(rec.date || "");
  if (!date) return;
  var last = sheet.getLastRow();
  var matchRow = 0;
  var dupes = [];
  if (last >= 2) {
    var dates = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < dates.length; i++) {
      if (dateKey_(dates[i][0]) !== date) continue;
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
