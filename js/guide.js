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

function upsert_(sheet, rec) {
  var last = sheet.getLastRow();
  var date = String(rec.date || "");
  if (last >= 2) {
    var dates = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < dates.length; i++) {
      if (String(dates[i][0]) === date) {
        sheet.getRange(i + 2, 1, 1, 12).setValues([row_(rec)]);
        return;
      }
    }
  }
  sheet.appendRow(row_(rec));
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
