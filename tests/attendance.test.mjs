import assert from "node:assert/strict";
import test from "node:test";
import XLSX from "xlsx";
import {
  applyImport,
  blankWorkInput,
  computeDay,
  defaultSettings,
  defaultState,
  dueAlerts,
  resolveDayStatus,
  fillMissingDays,
  settleToday,
} from "../js/calc.js";
import { buildWorkbook, normalizeTimeCell, parseSheetRows, payloadFromWorkbook } from "../js/excel.js";
import { DICT, formatClockTime, translate } from "../js/i18n.js";

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
  const english = payloadFromWorkbook(XLSX, buildWorkbook(XLSX, state));
  assert.equal(english.ok, true);
  assert.equal(english.payload.log[0].checkIn, "08:30");

  const bad = parseSheetRows([["signature", "NOPE"]], [["Date"], ["date"]]);
  assert.equal(bad.ok, false);
  assert.equal(bad.error, "فشل الاستيراد");
});

test("missing days before today become leave, and the last friday and saturday are weekly off", () => {
  const state = defaultState("2026-11-29");
  const filled = fillMissingDays(state);
  const byDate = Object.fromEntries(filled.log.map((row) => [row.date, row.status]));
  assert.equal(byDate["2026-11-27"], "weekend");
  assert.equal(byDate["2026-11-28"], "weekend");
  assert.equal(byDate["2026-11-29"], undefined);
  filled.dismissedDates = ["2026-11-28"];
  filled.log = filled.log.filter((row) => row.date !== "2026-11-28");
  const again = fillMissingDays(filled);
  assert.equal(again.log.some((row) => row.date === "2026-11-28"), false);
});

test("arabic and english dictionaries cover the same phrases", () => {
  assert.deepEqual(Object.keys(DICT.ar).sort(), Object.keys(DICT.en).sort());
  assert.equal(translate("ar", "checkin.now"), "سجل حضوري الآن");
  assert.equal(translate("en", "checkin.now"), "Check in now");
  assert.equal(translate("ar", "hero.kicker"), "المفروض أمشي الساعة كام");
  assert.equal(translate("en", "nav.report"), "Monthly log");
});
