import assert from "node:assert/strict";
import test from "node:test";
import XLSX from "xlsx";
import { APP_SIGNATURE, applyImport, applySheetRecords, defaultState, sanitizeStored } from "../js/calc.js";
import { buildLeaveWorkbook, buildMarkdown, buildWorkbook, leaveFromWorkbook, payloadFromWorkbook, readMarkdown } from "../js/excel.js";
import {
  addLeaveEntries,
  buildLeaveMarkdown,
  buildTimeline,
  coverLeaveHorizon,
  defaultLeave,
  entriesBeforeStart,
  expandLeaveHorizon,
  leaveFromMarkdown,
  leaveFromRows,
  removeLeaveEntry,
  resetLeaveSettings,
  setLeaveMonthMode,
  setLeaveMonthUsed,
  leaveDaysByMonthInYear,
  leaveDaysInMonth,
  leaveMonthsInYear,
  summarizeLeave,
} from "../js/leave.js";

test("accrual adds whole days and carries the fraction forward", () => {
  const timeline = buildTimeline(defaultLeave());
  const cycle = [0, 0.75, 0.5, 0.25];
  assert.equal(timeline.months.length, 24);
  assert.equal(timeline.months[0].month, "2026-08");
  assert.equal(timeline.months.at(-1).month, "2028-07");
  assert.equal(timeline.months[0].carryIn, 0.25);
  assert.equal(timeline.months[0].accrual, 1.75);
  assert.equal(timeline.months[0].cumulative, 2);
  assert.equal(timeline.months[0].added, 2);
  assert.equal(timeline.months[0].availableStart, 13);
  assert.equal(timeline.months[0].finalBalance, 13);
  assert.equal(timeline.months[3].month, "2026-11");
  assert.equal(timeline.months[3].availableStart, 18);
  assert.equal(timeline.months[4].month, "2026-12");
  assert.equal(timeline.months[4].finalBalance, 20);
  timeline.months.forEach((row, index) => {
    assert.equal(row.carryOut, cycle[index % 4]);
    assert.equal(row.added, Math.floor(row.cumulative));
    assert.equal(row.carryOut, Math.round((row.cumulative - row.added) * 10000) / 10000);
  });
  const summary = summarizeLeave(timeline, "2026-11");
  assert.equal(summary.available, 18);
  assert.equal(summary.usedMonth, 0);
  assert.equal(summary.remaining, 18);
  assert.equal(summary.carryOut, 0.25);
});

test("a date range splits across months and the same day is counted once", () => {
  const first = addLeaveEntries(defaultLeave(), { start: "2026-12-30", end: "2027-01-03", note: "trip" });
  assert.equal(first.error, null);
  assert.deepEqual(first.added, ["2026-12-30", "2026-12-31", "2027-01-01", "2027-01-02", "2027-01-03"]);
  const again = addLeaveEntries(first.leave, { start: "2026-12-31", end: "2027-01-01", note: "updated" });
  assert.deepEqual(again.added, []);
  assert.deepEqual(again.skipped, ["2026-12-31", "2027-01-01"]);
  assert.equal(again.leave.entries.length, 5);
  assert.equal(again.leave.entries.find((entry) => entry.date === "2026-12-31").note, "updated");
  const timeline = buildTimeline(again.leave);
  const december = summarizeLeave(timeline, "2026-12");
  assert.equal(december.usedMonth, 2);
  assert.equal(december.remaining, 18);
  assert.equal(timeline.months.find((row) => row.month === "2026-12").used, 2);
  assert.deepEqual(leaveDaysInMonth(again.leave, "2026-12").map((day) => day.date), ["2026-12-30", "2026-12-31"]);
  assert.deepEqual(leaveDaysByMonthInYear(again.leave, "2026"), [{
    month: "2026-12",
    days: [
      { date: "2026-12-30", note: "trip" },
      { date: "2026-12-31", note: "updated" },
    ],
  }]);
  assert.equal(leaveDaysByMonthInYear(again.leave, "2027")[0].days.length, 3);
  assert.deepEqual(leaveMonthsInYear(timeline, "2026"), [{ month: "2026-12", days: 2 }]);
  assert.deepEqual(leaveMonthsInYear(timeline, "2027"), [{ month: "2027-01", days: 3 }]);
  assert.equal(timeline.months.find((row) => row.month === "2027-01").used, 3);
  const removed = removeLeaveEntry(again.leave, "2027-01-02");
  assert.equal(buildTimeline(removed).months.find((row) => row.month === "2027-01").used, 2);
});

test("manual months override the date count and a deficit stays visible", () => {
  const dated = addLeaveEntries(defaultLeave(), { start: "2026-08-10", end: "2026-08-11" }).leave;
  const manual = setLeaveMonthUsed(dated, "2026-08", 30);
  const august = buildTimeline(manual).months[0];
  assert.equal(august.mode, "manual");
  assert.equal(august.dates.length, 2);
  assert.equal(august.used, 30);
  assert.equal(august.availableStart, 13);
  assert.equal(august.finalBalance, -17);
  assert.equal(august.deficit, true);
  const back = setLeaveMonthMode(manual, "2026-08", "auto");
  assert.equal(buildTimeline(back).months[0].used, 2);
  assert.equal(buildTimeline(back).months[0].mode, "auto");
});

test("the horizon grows to cover today and recorded days without dropping history", () => {
  const short = coverLeaveHorizon(defaultLeave(), "2026-11");
  short.settings.horizonMonths = 2;
  const covered = coverLeaveHorizon(short, "2026-11");
  assert.equal(covered.settings.horizonMonths, 4);
  const far = addLeaveEntries(defaultLeave(), { start: "2028-02-01", end: "2028-02-01" }).leave;
  far.settings.horizonMonths = 3;
  assert.ok(coverLeaveHorizon(far, "2026-11").settings.horizonMonths >= 19);
  const extended = expandLeaveHorizon(defaultLeave(), 12);
  assert.equal(extended.settings.horizonMonths, 36);
  const early = addLeaveEntries(defaultLeave(), { start: "2026-07-15", note: "before" }).leave;
  assert.equal(entriesBeforeStart(early).length, 1);
  assert.equal(buildTimeline(early).months[0].used, 0);
  const reset = resetLeaveSettings({ ...early, settings: { ...early.settings, monthlyAccrual: 3, startMonth: "2027-01" } });
  assert.equal(reset.settings.monthlyAccrual, 1.75);
  assert.equal(reset.settings.startMonth, "2026-08");
  assert.equal(reset.entries.length, 1);
});

test("old working-hours data keeps its log and gains the default leave balance", () => {
  const stored = sanitizeStored({
    settings: { officialCheckIn: "09:00", graceEnd: "10:00", dailyTargetMinutes: 480, dhuhr: { start: "12:00" }, asr: { start: "15:00" } },
    today: { date: "2026-11-26", checkIn: "09:10" },
    log: [{ date: "2026-11-25", status: "work", checkIn: "08:30", checkOut: "16:30", breakMinutes: 60, permissionMinutes: 0, permissionType: "", grossMinutes: 480, netMinutes: 420, targetMinutes: 480, overtimeMinutes: 0 }],
  }, "2026-11-26");
  assert.equal(stored.today.checkIn, "09:10");
  assert.equal(stored.log.length, 1);
  assert.equal(stored.leave.settings.startMonth, "2026-08");
  assert.equal(stored.leave.settings.initialBalance, 11);
  assert.equal(stored.leave.settings.initialCarryOver, 0.25);
  assert.equal(stored.leave.settings.monthlyAccrual, 1.75);
  assert.equal(stored.leave.settings.horizonMonths, 24);
  assert.deepEqual(stored.leave.entries, []);
});

test("import without a leave section keeps the leave already on the device", () => {
  const state = defaultState("2026-11-26");
  state.leave = addLeaveEntries(state.leave, { start: "2026-11-02", note: "doctor" }).leave;
  const applied = applyImport(state, {
    signature: APP_SIGNATURE,
    settings: state.settings,
    today: state.today,
    log: [],
  });
  assert.equal(applied.ok, true);
  assert.equal(applied.state.log.length, 0);
  assert.equal(applied.state.leave.entries[0].date, "2026-11-02");
  assert.equal(applied.state.leave.entries[0].note, "doctor");
  const replaced = applyImport(applied.state, {
    signature: APP_SIGNATURE,
    settings: state.settings,
    today: state.today,
    log: [],
    leave: { settings: { startMonth: "2026-08", horizonMonths: 24 }, entries: [{ date: "2026-08-04", note: "from file" }], months: [] },
  });
  assert.equal(replaced.state.leave.entries.length, 1);
  assert.equal(replaced.state.leave.entries[0].note, "from file");
  assert.equal(replaced.state.leave.settings.horizonMonths, 24);
  assert.equal(replaced.state.today.checkIn, state.today.checkIn);
});

test("sheet sync replaces attendance days and leaves the leave balance in place", () => {
  const state = defaultState("2026-11-26");
  state.leave = addLeaveEntries(state.leave, { start: "2026-11-05", note: "keep" }).leave;
  const result = applySheetRecords(state, [{
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
    checkedOut: "1",
  }], "2026-11-26");
  assert.equal(result.imported, 1);
  assert.equal(result.state.log[0].date, "2026-11-24");
  assert.equal(result.state.leave.entries[0].note, "keep");
});

test("excel and markdown round-trip the leave sheet without touching the log", () => {
  const state = defaultState("2026-11-26");
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
  state.leave = setLeaveMonthUsed(
    addLeaveEntries(state.leave, { start: "2026-12-30", end: "2027-01-02", note: "trip" }).leave,
    "2026-09",
    1.5,
  );
  const workbook = buildWorkbook(XLSX, state);
  assert.ok(workbook.SheetNames.includes("Leave Balance"));
  assert.ok(workbook.SheetNames.includes("Log"));
  const parsed = payloadFromWorkbook(XLSX, workbook);
  assert.equal(parsed.ok, true);
  assert.equal(parsed.payload.today.checkIn, "09:05");
  assert.equal(parsed.payload.log[0].netMinutes, 420);
  assert.equal(parsed.payload.leave.settings.monthlyAccrual, 1.75);
  assert.deepEqual(parsed.payload.leave.entries.map((entry) => entry.date), ["2026-12-30", "2026-12-31", "2027-01-01", "2027-01-02"]);
  assert.equal(parsed.payload.leave.entries[0].note, "trip");
  assert.equal(parsed.payload.leave.months["2026-09"].used, 1.5);

  const grid = XLSX.utils.sheet_to_json(workbook.Sheets["Leave Balance"], { header: 1, defval: "" });
  const accrual = grid.find((row) => row[0] === "monthlyAccrual");
  accrual[1] = "2";
  workbook.Sheets["Leave Balance"] = XLSX.utils.aoa_to_sheet(grid);
  const edited = payloadFromWorkbook(XLSX, workbook);
  assert.equal(edited.payload.leave.settings.monthlyAccrual, 2);
  assert.equal(edited.payload.settings.dailyTargetMinutes, state.settings.dailyTargetMinutes);

  delete workbook.Sheets["Leave Balance"];
  workbook.SheetNames = workbook.SheetNames.filter((name) => name !== "Leave Balance");
  const fromMeta = payloadFromWorkbook(XLSX, workbook);
  assert.equal(fromMeta.ok, true);
  assert.equal(fromMeta.payload.leave.entries.length, 4);

  const text = buildMarkdown(state);
  assert.match(text, /## رصيد الإجازات/);
  assert.match(text, /2026-12-30/);
  const fromMarkdown = readMarkdown(text);
  assert.equal(fromMarkdown.ok, true);
  assert.equal(fromMarkdown.payload.log[0].date, "2026-11-25");
  assert.equal(fromMarkdown.payload.leave.months["2026-09"].mode, "manual");
  assert.equal(fromMarkdown.payload.leave.entries.length, 4);
  const changed = readMarkdown(text.replace("| monthlyAccrual | 1.75 |", "| monthlyAccrual | 2.25 |"));
  assert.equal(changed.payload.leave.settings.monthlyAccrual, 2.25);
  assert.equal(changed.payload.today.checkIn, "09:05");
  assert.equal(leaveFromRows([["nope", "x"]]), null);
  assert.equal(parsed.payload.leave.entries[0].source, "manual");
});

test("leave days from work hours keep their source through a leave-only file", () => {
  const leave = addLeaveEntries(defaultLeave(), { start: "2026-11-03", end: "2026-11-03", source: "hours" }).leave;
  const workbook = buildLeaveWorkbook(XLSX, leave);
  assert.deepEqual(workbook.SheetNames, ["Leave Balance"]);
  const fromSheet = leaveFromWorkbook(XLSX, workbook);
  assert.equal(fromSheet.entries[0].source, "hours");
  const fromMarkdown = leaveFromMarkdown(buildLeaveMarkdown(leave, false));
  assert.equal(fromMarkdown.entries[0].date, "2026-11-03");
  assert.equal(fromMarkdown.entries[0].source, "hours");
  const oldSheet = leaveFromRows([
    ["signature", "WORKTIME_LEAVE_V1"],
    ["date", "note", ""],
    ["2026-11-01", "note", ""],
  ]);
  assert.equal(oldSheet.entries[0].source, "manual");
});
