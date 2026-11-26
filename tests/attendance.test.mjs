import assert from "node:assert/strict";
import test from "node:test";
import XLSX from "xlsx";
import {
  APP_SIGNATURE,
  applyImport,
  applySheetRecords,
  blankWorkInput,
  computeDay,
  defaultSettings,
  effectiveOvertimeMinutes,
  defaultState,
  dueAlerts,
  resolveDayStatus,
  fillMissingDays,
  settleToday,
} from "../js/calc.js";
import { buildMarkdown, buildWorkbook, normalizeTimeCell, overlaySettings, parseSheetRows, payloadFromWorkbook, readMarkdown } from "../js/excel.js";
import { APPS_SCRIPT, linkIsBehindScript, scriptChangedAt, stampWebhook } from "../js/guide.js";
import { APP_VERSION } from "../app.config.js";
import { DICT, formatClockTime, formatDayCount, formatLength, formatMinuteStamp, translate } from "../js/i18n.js";

test("morning permission at 11:00 credits 09:00 and leaves at 05:00", () => {
  const day = computeDay(blankWorkInput({ checkIn: "11:00", nowMin: 17 * 60, permissionMinutes: 120 }));
  assert.equal(day.permission.type, "morning");
  assert.equal(day.creditedStart, 9 * 60);
  assert.equal(day.expectedCheckout, 17 * 60);
  assert.equal(day.morningCredit, 120);
  assert.equal(day.breakAllowance, 60);
  assert.equal(Math.round(day.breakConsumed), 60);
  assert.equal(Math.round(day.gross), 480);
  assert.equal(Math.round(day.net), 420);
  assert.equal(formatClockTime(day.creditedStart, "ar"), "09:00 ص");
  assert.equal(formatClockTime(day.expectedCheckout, "en"), "05:00 PM");
  assert.equal(formatLength(5 * 60 + 40), "\u20665:40\u2069");
  assert.equal(formatLength(8 * 60), "\u20668:00\u2069");
  assert.equal(formatLength(7), "\u20660:07\u2069");
  assert.equal(formatLength(-90), "\u2066−1:30\u2069");
});

test("arrival at 09:31 clamps the morning credit to the official start", () => {
  const day = computeDay(blankWorkInput({ checkIn: "09:31", nowMin: 12 * 60 }));
  assert.equal(day.permission.type, "morning");
  assert.equal(day.creditedStart, 8 * 60 + 30);
  assert.equal(day.expectedCheckout, 16 * 60 + 30);
  assert.equal(day.morningCredit, 61);
});

test("eight hours start from the check-in minute", () => {
  const arrived = computeDay(blankWorkInput({ checkIn: "08:58", nowMin: 10 * 60 }));
  assert.equal(arrived.permission.active, false);
  assert.equal(arrived.creditedStart, 8 * 60 + 58);
  assert.equal(arrived.expectedCheckout, 16 * 60 + 58);
  const grace = computeDay(blankWorkInput({ checkIn: "09:15", nowMin: 10 * 60 }));
  assert.equal(grace.permission.active, false);
  assert.equal(grace.creditedStart, 9 * 60 + 15);
  assert.equal(grace.expectedCheckout, 17 * 60 + 15);
  const early = computeDay(blankWorkInput({ checkIn: "08:00", nowMin: 9 * 60 }));
  assert.equal(early.creditedStart, 8 * 60);
  assert.equal(early.expectedCheckout, 16 * 60);
});

test("evening permission moves leave earlier and skips a later break", () => {
  const day = computeDay(blankWorkInput({
    checkIn: "08:30",
    nowMin: 12 * 60,
    permissionType: "evening",
    permissionMinutes: 120,
  }));
  assert.equal(day.expectedCheckout, 14 * 60 + 30);
  assert.equal(day.breaks.asr.applied, false);
  assert.equal(day.breakAllowance, 45);
});

test("mid-day permission keeps the same leave time", () => {
  const day = computeDay(blankWorkInput({
    checkIn: "08:30",
    nowMin: 12 * 60,
    permissionType: "midday",
    permissionMinutes: 120,
  }));
  assert.equal(day.expectedCheckout, 16 * 60 + 30);
  assert.equal(day.middayExcused, 120);
});

test("turning permission off keeps a late arrival on the actual check-in", () => {
  const day = computeDay(blankWorkInput({
    checkIn: "11:00",
    nowMin: 12 * 60,
    permissionTouched: true,
    permissionOn: false,
  }));
  assert.equal(day.permission.active, false);
  assert.equal(day.creditedStart, 11 * 60);
  assert.equal(day.expectedCheckout, 19 * 60);
});

test("break minutes count one minute at a time and sum 45, 15, 60, plus manual", () => {
  const partial = computeDay(blankWorkInput({ checkIn: "08:30", nowMin: 12 * 60 }));
  assert.equal(Math.round(partial.breaks.dhuhr.elapsed), 10);
  assert.equal(partial.breakAllowance, 60);
  assert.equal(Math.round(partial.breakConsumed), 10);

  const dhuhrOnly = computeDay(blankWorkInput({ checkIn: "08:30", nowMin: 18 * 60, asrEnabled: false }));
  assert.equal(dhuhrOnly.breakAllowance, 45);
  const asrOnly = computeDay(blankWorkInput({ checkIn: "13:00", nowMin: 18 * 60 }));
  assert.equal(asrOnly.breaks.dhuhr.reason, "late");
  assert.equal(asrOnly.breakAllowance, 15);
  const both = computeDay(blankWorkInput({
    checkIn: "08:30",
    nowMin: 18 * 60,
    manualMinutes: 20,
    manualStart: "10:00",
  }));
  assert.equal(both.breakAllowance, 80);
  assert.equal(Math.round(both.breakConsumed), 80);
});

test("net hours stay flat while a break is running", () => {
  const before = computeDay(blankWorkInput({ checkIn: "08:30", nowMin: 11 * 60 + 50, asrEnabled: false }));
  const after = computeDay(blankWorkInput({ checkIn: "08:30", nowMin: 12 * 60 + 35, asrEnabled: false }));
  assert.equal(Math.round(before.net), 200);
  assert.equal(Math.round(after.net), 200);
  assert.equal(Math.round(after.breakConsumed), 45);
});

test("friday and saturday work counts the whole day as overtime", () => {
  const friday = computeDay(blankWorkInput({
    date: "2026-11-27",
    statusOverride: "work",
    checkIn: "09:00",
    checkOut: "17:00",
    nowMin: 17 * 60,
  }));
  assert.equal(friday.weekendWork, true);
  assert.equal(friday.overtime, 8 * 60);
  const saturday = computeDay(blankWorkInput({
    date: "2026-11-28",
    statusOverride: "work",
    checkIn: "09:00",
    checkOut: "17:00",
    nowMin: 18 * 60,
  }));
  assert.equal(saturday.overtime, 8 * 60);
  const longer = computeDay(blankWorkInput({
    date: "2026-11-27",
    statusOverride: "work",
    checkIn: "09:00",
    nowMin: 12 * 60,
  }));
  assert.equal(longer.overtime, 3 * 60);
  const weekday = computeDay(blankWorkInput({
    date: "2026-11-25",
    checkIn: "09:00",
    checkOut: "17:00",
    nowMin: 17 * 60,
  }));
  assert.equal(weekday.weekendWork, false);
  assert.equal(weekday.overtime, 0);
  assert.equal(effectiveOvertimeMinutes({
    date: "2026-11-27",
    status: "work",
    checkIn: "09:00",
    checkOut: "17:00",
    overtimeMinutes: 0,
  }), 8 * 60);
  assert.equal(effectiveOvertimeMinutes({
    date: "2026-11-27",
    status: "weekend",
    checkIn: "",
    checkOut: "",
    overtimeMinutes: 0,
  }), 0);
  assert.equal(effectiveOvertimeMinutes({
    date: "2026-11-25",
    status: "work",
    checkIn: "09:00",
    checkOut: "18:00",
    overtimeMinutes: 60,
  }), 60);
});

test("friday and saturday are weekly off unless overridden", () => {
  assert.equal(resolveDayStatus("2026-11-27", null), "weekend");
  assert.equal(resolveDayStatus("2026-11-28", null), "weekend");
  assert.equal(resolveDayStatus("2026-11-27", "work"), "work");
  assert.equal(resolveDayStatus("2026-11-25", "personal"), "personal");
});

test("opening the app on a new day archives the previous one", () => {
  const state = defaultState("2026-11-24");
  state.today.checkIn = "08:30";
  const result = settleToday(state, "2026-11-25");
  assert.equal(result.archived.date, "2026-11-24");
  assert.equal(result.archived.status, "work");
  assert.equal(result.state.today.date, "2026-11-25");
  assert.equal(result.state.today.checkIn, null);
  assert.equal(result.state.log.length, 1);
  assert.equal(settleToday(result.state, "2026-11-25").archived, null);
});

test("alerts fire inside their windows and a late arrival skips the missed break", () => {
  const model = computeDay(blankWorkInput({ checkIn: "08:30", nowMin: 11 * 60 + 40 }));
  const alerts = dueAlerts({ nowMin: 11 * 60 + 40, model, notifications: defaultSettings().notifications, firedKeys: [] });
  assert.ok(alerts.some((alert) => alert.id === "pre-dhuhr"));
  const late = computeDay(blankWorkInput({ checkIn: "13:00", nowMin: 11 * 60 + 45 }));
  const skipped = dueAlerts({ nowMin: 11 * 60 + 45, model: late, notifications: defaultSettings().notifications, firedKeys: [] });
  assert.equal(skipped.some((alert) => alert.id === "pre-dhuhr"), false);
});

test("invalid import does not replace current data", () => {
  const state = defaultState("2026-11-25");
  state.today.checkIn = "08:45";
  const result = applyImport(state, { signature: "OTHER" });
  assert.equal(result.ok, false);
  assert.equal(result.error, "فشل الاستيراد");
  assert.equal(result.state, state);
  assert.equal(state.today.checkIn, "08:45");
});

test("excel time fractions and a real workbook round-trip", () => {
  assert.equal(normalizeTimeCell(8.5 / 24), "08:30");
  const state = defaultState("2026-11-25");
  state.today.checkIn = "09:10";
  state.log.push({
    date: "2026-11-24",
    status: "work",
    checkIn: "08:30",
    checkOut: "16:30",
    breakMinutes: 60,
    permissionMinutes: 0,
    permissionType: "",
    grossMinutes: 480,
    netMinutes: 420,
    targetMinutes: 480,
    overtimeMinutes: 0,
  });
  const parsed = payloadFromWorkbook(XLSX, buildWorkbook(XLSX, state));
  assert.equal(parsed.ok, true);
  assert.equal(parsed.payload.today.checkIn, "09:10");
  assert.equal(parsed.payload.log[0].status, "work");
  assert.equal(parsed.payload.log[0].netMinutes, 420);

  state.settings.lang = "en";
  state.settings.accent = "ocean";
  state.settings.dailyTargetMinutes = 510;
  state.settings.alertSeconds = 8;
  state.settings.dhuhr.duration = 30;
  const english = payloadFromWorkbook(XLSX, buildWorkbook(XLSX, state));
  assert.equal(english.ok, true);
  assert.equal(english.payload.log[0].checkIn, "08:30");
  assert.equal(english.payload.settings.accent, "ocean");
  assert.equal(english.payload.settings.dailyTargetMinutes, 510);
  assert.equal(english.payload.settings.alertSeconds, 8);
  assert.equal(defaultSettings().alertSeconds, 5);
  assert.equal(english.payload.settings.dhuhr.duration, 30);
  assert.equal(english.payload.settings.fileSnapshot, false);
  assert.equal(english.payload.settings.notifications.preBreak.minutesBefore, 10);
  assert.ok(buildWorkbook(XLSX, state).SheetNames.includes("Settings"));

  const bad = parseSheetRows([["signature", "NOPE"]], [["Date"], ["date"]]);
  assert.equal(bad.ok, false);
  assert.equal(bad.error, "فشل الاستيراد");
});

test("markdown backup keeps settings and rejects a foreign file", () => {
  const state = defaultState("2026-11-26");
  state.settings.accent = "violet";
  state.settings.officialCheckIn = "09:00";
  state.settings.dhuhr.duration = 30;
  state.settings.theme = "light";
  state.settings.sheetsWebhook = "https://script.google.com/macros/s/abc/exec";
  state.settings.fileSnapshot = true;
  state.today.checkIn = "09:05";
  state.log.push({
    date: "2026-11-25",
    status: "work",
    checkIn: "08:30",
    checkOut: "16:30",
    breakMinutes: 60,
    permissionMinutes: 0,
    permissionType: "",
    grossMinutes: 480,
    netMinutes: 420,
    targetMinutes: 480,
    overtimeMinutes: 0,
  });
  const text = buildMarkdown(state);
  assert.match(text, new RegExp(APP_SIGNATURE));
  const parsed = readMarkdown(text);
  assert.equal(parsed.ok, true);
  assert.equal(parsed.payload.settings.accent, "violet");
  assert.equal(parsed.payload.settings.officialCheckIn, "09:00");
  assert.equal(parsed.payload.settings.dhuhr.duration, 30);
  assert.equal(parsed.payload.settings.theme, "light");
  assert.equal(parsed.payload.settings.alertSeconds, 5);
  assert.equal(parsed.payload.settings.sheetsWebhook, state.settings.sheetsWebhook);
  assert.equal(parsed.payload.settings.fileSnapshot, false);
  assert.match(text, /notifications\.preBreak\.minutesBefore/);
  const edited = text.replace(
    "| دقائق التنبيه قبل البريك | 10 | notifications.preBreak.minutesBefore |",
    "| دقائق التنبيه قبل البريك | 22 | notifications.preBreak.minutesBefore |",
  );
  const fromTable = readMarkdown(edited);
  assert.equal(fromTable.ok, true);
  assert.equal(fromTable.payload.settings.notifications.preBreak.minutesBefore, 22);
  assert.equal(fromTable.payload.settings.dhuhr.duration, 30);
  assert.equal(parsed.payload.today.checkIn, "09:05");
  const applied = applyImport(defaultState("2026-11-26"), parsed.payload);
  assert.equal(applied.ok, true);
  assert.equal(applied.state.settings.accent, "violet");
  assert.equal(applied.state.settings.officialCheckIn, "09:00");
  assert.equal(applied.state.settings.dhuhr.duration, 30);
  assert.equal(applied.state.today.checkIn, "09:05");
  assert.equal(readMarkdown("# notes\n\nnot our file").ok, false);
  assert.equal(readMarkdown(text.replaceAll(APP_SIGNATURE, "OTHER")).ok, false);
});

test("settings sheet overrides the hidden copy and still keeps the device link", () => {
  const state = defaultState("2026-11-26");
  state.settings.notifications.preBreak.minutesBefore = 10;
  state.settings.sheetsWebhook = "https://script.google.com/macros/s/local/exec";
  state.settings.sheetsConnected = true;
  const workbook = buildWorkbook(XLSX, state);
  const grid = XLSX.utils.sheet_to_json(workbook.Sheets.Settings, { header: 1, defval: "" });
  const row = grid.find((item) => item[2] === "notifications.preBreak.minutesBefore");
  assert.equal(String(row[1]), "10");
  row[1] = "25";
  workbook.Sheets.Settings = XLSX.utils.aoa_to_sheet(grid);
  const parsed = payloadFromWorkbook(XLSX, workbook);
  assert.equal(parsed.ok, true);
  assert.equal(parsed.payload.settings.notifications.preBreak.minutesBefore, 25);
  assert.equal(parsed.payload.settings.notifications.preBreak.enabled, true);

  delete workbook.Sheets.Settings;
  workbook.SheetNames = workbook.SheetNames.filter((name) => name !== "Settings");
  const legacy = payloadFromWorkbook(XLSX, workbook);
  assert.equal(legacy.ok, true);
  assert.equal(legacy.payload.settings.notifications.preBreak.minutesBefore, 10);

  const synced = overlaySettings(state.settings, [
    { key: "notifications.preBreak.minutesBefore", value: "15" },
    { key: "dhuhr.enabled", value: "false" },
    { key: "sheetsWebhook", value: "https://script.google.com/macros/s/remote/exec" },
  ], { keepLink: true });
  assert.equal(synced.notifications.preBreak.minutesBefore, 15);
  assert.equal(synced.dhuhr.enabled, false);
  assert.equal(synced.dhuhr.start, state.settings.dhuhr.start);
  assert.equal(synced.sheetsWebhook, state.settings.sheetsWebhook);
  assert.equal(synced.sheetsConnected, true);
});

test("the sheet script reads and writes a settings sheet", () => {
  assert.match(APPS_SCRIPT, /الإعدادات/);
  assert.match(APPS_SCRIPT, /function pullSettings_/);
  assert.match(APPS_SCRIPT, /function writeSettings_/);
  assert.match(APPS_SCRIPT, /settings: pullSettings_\(\)/);
  assert.match(APPS_SCRIPT, /رصيد الإجازات/);
  assert.match(APPS_SCRIPT, /function pullLeave_/);
  assert.match(APPS_SCRIPT, /function writeLeave_/);
  assert.match(APPS_SCRIPT, /leave: pullLeave_\(\)/);
  assert.match(APPS_SCRIPT, /var CHANGED_AT = "2026-11-26T11:26:00.000Z"/);
  assert.match(APPS_SCRIPT, /source: c === "hours" \? "hours" : "manual"/);
  assert.doesNotThrow(() => new Function(APPS_SCRIPT));
});

test("an older sheet link is behind the script's last change", () => {
  const scriptAt = scriptChangedAt();
  assert.equal(scriptAt, "2026-11-26T11:26:00.000Z");
  const webhook = "https://script.google.com/macros/s/abc/exec";
  assert.equal(linkIsBehindScript({ sheetsWebhook: webhook }), true);
  assert.equal(linkIsBehindScript({ sheetsWebhook: webhook, sheetsWebhookAt: "2026-11-25T09:00:00.000Z" }), true);
  assert.equal(linkIsBehindScript({ sheetsWebhook: webhook, sheetsWebhookAt: scriptAt }), false);
  assert.equal(linkIsBehindScript({ sheetsWebhook: webhook, sheetsWebhookAt: "2026-11-26T08:40:45.000Z" }), true);
  assert.equal(linkIsBehindScript({ sheetsWebhook: webhook, sheetsWebhookAt: "2026-11-26T11:25:59.000Z" }), true);
  assert.equal(linkIsBehindScript({ sheetsWebhook: webhook, sheetsWebhookAt: "2026-11-26T11:26:30.000Z" }), false);
  assert.equal(linkIsBehindScript({ sheetsWebhook: webhook, sheetsWebhookAt: "2026-11-26T13:00:30.000Z" }), false);
  assert.equal(linkIsBehindScript({ sheetsWebhook: webhook, sheetsWebhookAt: "2026-11-26T10:00:00.000Z" }), true);
  assert.equal(linkIsBehindScript({ sheetsWebhook: "", sheetsWebhookAt: "" }), false);
  const now = "2026-11-26T10:05:00.000Z";
  assert.equal(stampWebhook(webhook, webhook, "", now), "");
  assert.equal(stampWebhook(webhook, webhook, "2026-11-25T09:00:00.000Z", now), "2026-11-25T09:00:00.000Z");
  assert.equal(stampWebhook(webhook, "https://script.google.com/macros/s/new/exec", "2026-11-25T09:00:00.000Z", now), now);
  assert.equal(stampWebhook(webhook, "", "2026-11-25T09:00:00.000Z", now), "");
  const shown = formatMinuteStamp("2026-11-26T08:40:45.000Z", "en");
  assert.match(shown, /2026/);
  assert.match(shown, /\d{2}:\d{2}/);
  assert.equal(shown.includes(":45"), false);
});

test("sheet import replaces local days and restores today from the spreadsheet", () => {
  const state = defaultState("2026-11-26");
  state.settings.sheetsWebhook = "https://script.google.com/macros/s/abc/exec";
  state.today.checkIn = "09:00";
  state.log = [
    { date: "2026-11-24", status: "personal", checkIn: "", checkOut: "", breakMinutes: 0, permissionMinutes: 0, permissionType: "", grossMinutes: 0, netMinutes: 0, targetMinutes: 0, overtimeMinutes: 0 },
    { date: "2026-11-25", status: "personal", checkIn: "", checkOut: "", breakMinutes: 0, permissionMinutes: 0, permissionType: "", grossMinutes: 0, netMinutes: 0, targetMinutes: 0, overtimeMinutes: 0 },
  ];
  const result = applySheetRecords(state, [
    { date: "2026-11-24", status: "عمل", checkIn: "08:30", checkOut: "16:30", breakMinutes: 60, permissionMinutes: 0, permissionType: "", grossMinutes: 480, netMinutes: 420, targetMinutes: 480, overtimeMinutes: 0 },
    { date: "2026-11-25", status: "Work", checkIn: "08:45", checkOut: "16:45", breakMinutes: 60, permissionMinutes: 120, permissionType: "morning", grossMinutes: 480, netMinutes: 420, targetMinutes: 480, overtimeMinutes: 0 },
    { date: "2026-11-26", status: "عمل", checkIn: "08:40", checkOut: "16:40", checkedOut: "", breakMinutes: 15, permissionMinutes: 0, permissionType: "", grossMinutes: 80, netMinutes: 65, targetMinutes: 480, overtimeMinutes: 0 },
  ], "2026-11-26");
  assert.equal(result.imported, 3);
  assert.deepEqual(result.state.log.map((row) => row.date), ["2026-11-24", "2026-11-25"]);
  assert.equal(result.state.log[0].checkIn, "08:30");
  assert.equal(result.state.log[0].status, "work");
  assert.equal(result.state.log[1].permissionType, "morning");
  assert.equal(result.state.today.checkIn, "08:40");
  assert.equal(result.state.today.checkOut, null);
  assert.equal(result.state.today.permissionTouched, true);
  assert.equal(result.state.settings.sheetsWebhook, state.settings.sheetsWebhook);

  const checkedOut = applySheetRecords(defaultState("2026-11-26"), [
    { date: "2026-11-26", status: "work", checkIn: "08:30", checkOut: "17:10", checkedOut: "1", breakMinutes: 60, permissionMinutes: 0, permissionType: "", grossMinutes: 480, netMinutes: 420, targetMinutes: 480, overtimeMinutes: 0 },
  ], "2026-11-26");
  assert.equal(checkedOut.state.today.checkOut, "17:10");
  assert.equal(checkedOut.state.log.length, 0);
  assert.equal(applySheetRecords(state, [], "2026-11-26").state, state);
});

test("an empty log starts at today and does not invent earlier leave", () => {
  const state = defaultState("2026-11-26");
  const filled = fillMissingDays(state);
  assert.equal(filled, state);
  assert.equal(filled.log.length, 0);
});

test("gaps after the first logged day become leave, and a dismissed day stays out", () => {
  const state = defaultState("2026-11-30");
  state.log = [{
    date: "2026-11-26",
    status: "work",
    checkIn: "08:30",
    checkOut: "16:30",
    breakMinutes: 60,
    permissionMinutes: 0,
    permissionType: "",
    grossMinutes: 480,
    netMinutes: 420,
    targetMinutes: 480,
    overtimeMinutes: 0,
  }];
  const filled = fillMissingDays(state);
  const byDate = Object.fromEntries(filled.log.map((row) => [row.date, row.status]));
  assert.equal(byDate["2026-11-25"], undefined);
  assert.equal(byDate["2026-11-26"], "work");
  assert.equal(byDate["2026-11-27"], "weekend");
  assert.equal(byDate["2026-11-28"], "weekend");
  assert.equal(byDate["2026-11-29"], "official");
  assert.equal(byDate["2026-11-30"], undefined);
  filled.dismissedDates = ["2026-11-28"];
  filled.log = filled.log.filter((row) => row.date !== "2026-11-28");
  const again = fillMissingDays(filled);
  assert.equal(again.log.some((row) => row.date === "2026-11-28"), false);
  assert.equal(again.log.some((row) => row.date === "2026-11-20"), false);
});

test("personal leave day counts use arabic and english wording", () => {
  assert.equal(formatDayCount(0, "ar"), "0 يوم");
  assert.equal(formatDayCount(1, "ar"), "يوم واحد");
  assert.equal(formatDayCount(2, "ar"), "يومان");
  assert.equal(formatDayCount(5, "ar"), "5 أيام");
  assert.equal(formatDayCount(11, "ar"), "11 يوم");
  assert.equal(formatDayCount(1, "en"), "1 day");
  assert.equal(formatDayCount(4, "en"), "4 days");
});

test("the release version is edited in the app config page", () => {
  assert.equal(APP_VERSION, "1.0.2");
  assert.equal(translate("ar", "ui.version", { version: APP_VERSION }), `النسخة ${APP_VERSION}`);
  assert.equal(translate("en", "ui.version", { version: APP_VERSION }), `Version ${APP_VERSION}`);
});

test("arabic and english dictionaries cover the same phrases", () => {
  assert.deepEqual(Object.keys(DICT.ar).sort(), Object.keys(DICT.en).sort());
  assert.equal(translate("ar", "checkin.now"), "سجل حضوري الآن");
  assert.equal(translate("en", "checkin.now"), "Check in now");
  assert.equal(translate("ar", "hero.kicker"), "المفروض أمشي الساعة كام");
  assert.equal(translate("en", "nav.report"), "Monthly log");
});
