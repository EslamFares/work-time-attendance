export const LEAVE_SIGNATURE = "WORKTIME_LEAVE_V1";
export const LEAVE_SHEET_NAME = "Leave Balance";
export const LEAVE_SHEET_ALIASES = [LEAVE_SHEET_NAME, "رصيد الإجازات"];

export const DEFAULT_LEAVE_SETTINGS = {
  startMonth: "2026-08",
  initialBalance: 11,
  initialCarryOver: 0.25,
  monthlyAccrual: 1.75,
  horizonMonths: 24,
};

const MAX_HORIZON = 600;
const MAX_RANGE_DAYS = 4000;
const MAX_ENTRIES = 20000;

export function round4(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  const rounded = Math.round((n + Number.EPSILON) * 10000) / 10000;
  return Object.is(rounded, -0) ? 0 : rounded;
}

export function formatLeaveAmount(value) {
  const n = round4(value);
  if (!Number.isFinite(n)) return "0";
  const text = Math.abs(n).toFixed(4).replace(/0+$/, "").replace(/\.$/, "");
  return n < 0 ? `-${text}` : text;
}

function clampNum(value, min, max, fallback) {
  const n = typeof value === "number" ? value : Number(String(value ?? "").trim());
  if (!Number.isFinite(n)) return fallback;
  return round4(Math.min(max, Math.max(min, n)));
}

function leaveSource(value) {
  return value === "hours" ? "hours" : "manual";
}

function clampInt(value, min, max, fallback) {
  const n = Number(String(value ?? "").trim());
  if (!Number.isFinite(n)) return fallback;
  return Math.min(max, Math.max(min, Math.round(n)));
}

function cleanNote(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim().slice(0, 500);
}

export function parseMonth(value) {
  const match = /^(\d{4})-(\d{1,2})$/.exec(String(value ?? "").trim());
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  if (year < 1900 || year > 2500 || month < 1 || month > 12) return null;
  return { year, month, key: `${match[1]}-${String(month).padStart(2, "0")}` };
}

export function monthDiff(start, key) {
  const from = parseMonth(start);
  const to = parseMonth(key);
  if (!from || !to) return null;
  return (to.year * 12 + (to.month - 1)) - (from.year * 12 + (from.month - 1));
}

function monthFromIndex(index) {
  const year = Math.floor(index / 12);
  const month = index - year * 12;
  return `${String(year).padStart(4, "0")}-${String(month + 1).padStart(2, "0")}`;
}

export function addMonths(key, delta) {
  const parsed = parseMonth(key);
  if (!parsed) return "";
  return monthFromIndex(parsed.year * 12 + (parsed.month - 1) + delta);
}

function isIsoDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ""))) return false;
  const [year, month, day] = String(value).split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

function asIsoDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    const iso = `${year}-${month}-${day}`;
    return isIsoDate(iso) ? iso : "";
  }
  if (typeof value === "number" && value > 20000 && value < 80000) {
    const utc = new Date(Math.round((value - 25569) * 86400 * 1000));
    const year = utc.getUTCFullYear();
    const month = String(utc.getUTCMonth() + 1).padStart(2, "0");
    const day = String(utc.getUTCDate()).padStart(2, "0");
    const iso = `${year}-${month}-${day}`;
    return isIsoDate(iso) ? iso : "";
  }
  const text = String(value ?? "").trim();
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(text);
  if (!match) return "";
  const iso = `${match[1]}-${match[2]}-${match[3]}`;
  return isIsoDate(iso) ? iso : "";
}

function addDaysIso(iso, days) {
  const [year, month, day] = iso.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day + days));
  const nextYear = date.getUTCFullYear();
  const nextMonth = String(date.getUTCMonth() + 1).padStart(2, "0");
  const nextDay = String(date.getUTCDate()).padStart(2, "0");
  return `${nextYear}-${nextMonth}-${nextDay}`;
}

export function datesInRange(start, end) {
  const from = asIsoDate(start);
  const to = asIsoDate(end || start);
  if (!from || !to) return { dates: [], error: "invalid-date" };
  let lo = from;
  let hi = to;
  if (hi < lo) {
    const swap = lo;
    lo = hi;
    hi = swap;
  }
  const dates = [];
  for (let cursor = lo; cursor <= hi; cursor = addDaysIso(cursor, 1)) {
    dates.push(cursor);
    if (dates.length > MAX_RANGE_DAYS) return { dates: [], error: "range" };
  }
  return { dates, error: null };
}

export function defaultLeave() {
  return {
    settings: { ...DEFAULT_LEAVE_SETTINGS },
    entries: [],
    months: {},
  };
}

export function sanitizeLeaveSettings(raw) {
  const source = raw && typeof raw === "object" && !Array.isArray(raw) ? raw : {};
  const start = parseMonth(source.startMonth);
  return {
    startMonth: start ? start.key : DEFAULT_LEAVE_SETTINGS.startMonth,
    initialBalance: clampNum(source.initialBalance ?? DEFAULT_LEAVE_SETTINGS.initialBalance, 0, 100000, DEFAULT_LEAVE_SETTINGS.initialBalance),
    initialCarryOver: clampNum(source.initialCarryOver ?? DEFAULT_LEAVE_SETTINGS.initialCarryOver, 0, 100000, DEFAULT_LEAVE_SETTINGS.initialCarryOver),
    monthlyAccrual: clampNum(source.monthlyAccrual ?? DEFAULT_LEAVE_SETTINGS.monthlyAccrual, 0, 100, DEFAULT_LEAVE_SETTINGS.monthlyAccrual),
    horizonMonths: clampInt(source.horizonMonths ?? DEFAULT_LEAVE_SETTINGS.horizonMonths, 1, MAX_HORIZON, DEFAULT_LEAVE_SETTINGS.horizonMonths),
  };
}

function monthRecords(raw) {
  const months = {};
  const write = (key, mode, used) => {
    const parsed = parseMonth(key);
    if (!parsed || mode !== "manual") return;
    months[parsed.key] = { mode: "manual", used: clampNum(used, 0, 10000, 0) };
  };
  if (Array.isArray(raw)) {
    raw.forEach((item) => {
      if (!item || typeof item !== "object") return;
      write(item.month, item.mode === "manual" ? "manual" : "", item.used);
    });
    return months;
  }
  if (!raw || typeof raw !== "object") return months;
  Object.keys(raw).forEach((key) => {
    const item = raw[key];
    if (!item || typeof item !== "object") return;
    write(key, item.mode === "manual" ? "manual" : "", item.used);
  });
  return months;
}

export function sanitizeLeave(raw) {
  const source = raw && typeof raw === "object" && !Array.isArray(raw) ? raw : {};
  const settings = sanitizeLeaveSettings(source.settings);
  const seen = new Map();
  const list = Array.isArray(source.entries) ? source.entries : [];
  list.forEach((item) => {
    if (seen.size >= MAX_ENTRIES) return;
    const date = asIsoDate(item?.date ?? item);
    if (!date || seen.has(date)) return;
    seen.set(date, { date, note: cleanNote(item?.note), source: leaveSource(item?.source) });
  });
  const entries = [...seen.values()].sort((a, b) => a.date.localeCompare(b.date));
  return { settings, entries, months: monthRecords(source.months) };
}

export function requiredHorizon(leave, todayMonth) {
  const clean = sanitizeLeave(leave);
  let need = clean.settings.horizonMonths;
  const consider = (key) => {
    const index = monthDiff(clean.settings.startMonth, key);
    if (index != null && index + 1 > need) need = index + 1;
  };
  consider(todayMonth);
  clean.entries.forEach((entry) => consider(entry.date.slice(0, 7)));
  Object.keys(clean.months).forEach(consider);
  return Math.min(MAX_HORIZON, Math.max(1, need));
}

export function coverLeaveHorizon(leave, todayMonth) {
  const clean = sanitizeLeave(leave);
  const need = requiredHorizon(clean, todayMonth);
  if (need === clean.settings.horizonMonths) return clean;
  return { ...clean, settings: { ...clean.settings, horizonMonths: need } };
}

export function expandLeaveHorizon(leave, extra = 12) {
  const clean = sanitizeLeave(leave);
  const step = Math.max(1, Math.round(Number(extra) || 12));
  const horizonMonths = Math.min(MAX_HORIZON, clean.settings.horizonMonths + step);
  return { ...clean, settings: { ...clean.settings, horizonMonths } };
}

export function updateLeaveSettings(leave, patch) {
  const clean = sanitizeLeave(leave);
  return sanitizeLeave({ ...clean, settings: { ...clean.settings, ...(patch || {}) } });
}

export function resetLeaveSettings(leave) {
  const clean = sanitizeLeave(leave);
  return { ...clean, settings: { ...DEFAULT_LEAVE_SETTINGS } };
}

export function buildTimeline(leave) {
  const clean = sanitizeLeave(leave);
  const settings = clean.settings;
  const byMonth = new Map();
  clean.entries.forEach((entry) => {
    const key = entry.date.slice(0, 7);
    const list = byMonth.get(key) || [];
    list.push(entry);
    byMonth.set(key, list);
  });
  const months = [];
  let previous = null;
  for (let index = 0; index < settings.horizonMonths; index += 1) {
    const month = addMonths(settings.startMonth, index);
    const carryIn = index === 0 ? round4(settings.initialCarryOver) : previous.carryOut;
    const accrual = round4(settings.monthlyAccrual);
    const cumulative = round4(carryIn + accrual);
    const added = Math.floor(cumulative);
    const carryOut = round4(cumulative - added);
    const availableStart = index === 0
      ? round4(settings.initialBalance + added)
      : round4(previous.finalBalance + added);
    const dates = byMonth.get(month) || [];
    const manual = clean.months[month];
    const mode = manual?.mode === "manual" ? "manual" : "auto";
    const used = mode === "manual" ? round4(manual.used) : dates.length;
    const finalBalance = round4(availableStart - used);
    const row = {
      month,
      index,
      carryIn,
      accrual,
      cumulative,
      added,
      carryOut,
      availableStart,
      used,
      finalBalance,
      mode,
      dates,
      deficit: finalBalance < 0,
    };
    months.push(row);
    previous = row;
  }
  return { settings, months };
}

export function summarizeLeave(timeline, todayMonth) {
  const settings = timeline.settings;
  const current = timeline.months.find((row) => row.month === todayMonth) || null;
  const beforeStart = monthDiff(settings.startMonth, todayMonth);
  if (!current) {
    const untouched = beforeStart != null && beforeStart < 0;
    return {
      available: settings.initialBalance,
      carryOut: settings.initialCarryOver,
      usedMonth: 0,
      remaining: untouched ? settings.initialBalance : null,
      current: null,
    };
  }
  return {
    available: current.availableStart,
    carryOut: current.carryOut,
    usedMonth: current.used,
    remaining: current.finalBalance,
    current,
  };
}

export function entriesBeforeStart(leave) {
  const clean = sanitizeLeave(leave);
  return clean.entries.filter((entry) => entry.date.slice(0, 7) < clean.settings.startMonth);
}

export function leaveDaysInMonth(leave, month) {
  const key = String(month || "");
  return sanitizeLeave(leave).entries
    .filter((entry) => entry.date.slice(0, 7) === key)
    .map((entry) => ({ date: entry.date, note: entry.note }));
}

export function leaveMonthsInYear(timeline, year) {
  const prefix = `${year}-`;
  return (timeline?.months || [])
    .filter((row) => row.month.startsWith(prefix) && row.used > 0)
    .map((row) => ({ month: row.month, days: row.used }));
}

export function leaveDaysByMonthInYear(leave, year) {
  const prefix = `${year}-`;
  const groups = new Map();
  sanitizeLeave(leave).entries.forEach((entry) => {
    if (!entry.date.startsWith(prefix)) return;
    const month = entry.date.slice(0, 7);
    const list = groups.get(month) || [];
    list.push({ date: entry.date, note: entry.note });
    groups.set(month, list);
  });
  return [...groups.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([month, days]) => ({
      month,
      days: days.sort((a, b) => a.date.localeCompare(b.date)),
    }));
}

export function addLeaveEntries(leave, { start, end, note, source } = {}) {
  const clean = sanitizeLeave(leave);
  const range = datesInRange(start, end || start);
  if (range.error) return { leave: clean, added: [], skipped: [], error: range.error };
  const byDate = new Map(clean.entries.map((entry) => [entry.date, { ...entry }]));
  const added = [];
  const skipped = [];
  const text = cleanNote(note);
  const origin = leaveSource(source);
  range.dates.forEach((date) => {
    if (byDate.size >= MAX_ENTRIES && !byDate.has(date)) return;
    const current = byDate.get(date);
    if (current) {
      if (text) current.note = text;
      skipped.push(date);
      return;
    }
    byDate.set(date, { date, note: text, source: origin });
    added.push(date);
  });
  const entries = [...byDate.values()].sort((a, b) => a.date.localeCompare(b.date));
  return { leave: { ...clean, entries }, added, skipped, error: null };
}

export function removeLeaveEntry(leave, date) {
  const clean = sanitizeLeave(leave);
  const target = asIsoDate(date);
  if (!target) return clean;
  return { ...clean, entries: clean.entries.filter((entry) => entry.date !== target) };
}

export function setLeaveMonthMode(leave, month, mode) {
  const clean = sanitizeLeave(leave);
  const key = parseMonth(month)?.key;
  if (!key) return clean;
  const months = { ...clean.months };
  if (mode !== "manual") {
    delete months[key];
    return { ...clean, months };
  }
  if (months[key]?.mode === "manual") return clean;
  const timeline = buildTimeline({ ...clean, months });
  const row = timeline.months.find((item) => item.month === key);
  months[key] = { mode: "manual", used: row ? row.used : 0 };
  return { ...clean, months };
}

export function setLeaveMonthUsed(leave, month, used) {
  const clean = sanitizeLeave(leave);
  const key = parseMonth(month)?.key;
  if (!key) return clean;
  return {
    ...clean,
    months: { ...clean.months, [key]: { mode: "manual", used: clampNum(used, 0, 10000, 0) } },
  };
}

export function leaveSyncPayload(leave) {
  const clean = sanitizeLeave(leave);
  return {
    settings: { ...clean.settings },
    entries: clean.entries.map((entry) => ({ date: entry.date, note: entry.note, source: entry.source })),
    months: Object.keys(clean.months).sort().map((month) => ({
      month,
      mode: "manual",
      used: clean.months[month].used,
    })),
  };
}

export function leaveSheetRows(leave) {
  const payload = leaveSyncPayload(leave);
  const rows = [
    ["signature", LEAVE_SIGNATURE, ""],
    ["key", "value", ""],
    ["startMonth", payload.settings.startMonth, ""],
    ["initialBalance", String(payload.settings.initialBalance), ""],
    ["initialCarryOver", String(payload.settings.initialCarryOver), ""],
    ["monthlyAccrual", String(payload.settings.monthlyAccrual), ""],
    ["horizonMonths", String(payload.settings.horizonMonths), ""],
    ["", "", ""],
    ["date", "note", "source"],
  ];
  payload.entries.forEach((entry) => rows.push([entry.date, entry.note, entry.source]));
  rows.push(["", "", ""]);
  rows.push(["month", "mode", "used"]);
  payload.months.forEach((item) => rows.push([item.month, item.mode, String(item.used)]));
  return rows;
}

function cellText(value) {
  return String(value ?? "").trim();
}

export function leaveFromRows(rows) {
  const list = Array.isArray(rows) ? rows : [];
  let signature = false;
  let section = "settings";
  const settings = {};
  const entries = [];
  const months = [];
  list.forEach((row) => {
    const cells = Array.isArray(row) ? row : [];
    const key = cellText(cells[0]);
    const second = cellText(cells[1]);
    const third = cellText(cells[2]);
    if (!key && !second && !third) return;
    if (key === "signature" && second === LEAVE_SIGNATURE) {
      signature = true;
      return;
    }
    if (!signature) return;
    if (key === "key" && (second === "value" || second === "القيمة")) {
      section = "settings";
      return;
    }
    if (key === "date" && (second === "note" || second === "ملاحظة")) {
      section = "entries";
      return;
    }
    if (key === "month" && (second === "mode" || second === "الوضع")) {
      section = "months";
      return;
    }
    if (section === "settings") settings[key] = cells[1];
    else if (section === "entries") entries.push({ date: cells[0], note: second, source: third });
    else if (section === "months") months.push({ month: cells[0], mode: second, used: cells[2] });
  });
  if (!signature) return null;
  return sanitizeLeave({ settings, entries, months });
}

function markdownTables(source) {
  const tables = [];
  let current = null;
  String(source || "").split(/\r?\n/).forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed.startsWith("|")) {
      if (current) tables.push(current);
      current = null;
      return;
    }
    const cells = trimmed.split("|").slice(1, -1).map((cell) => cell.trim());
    if (cells.every((cell) => /^:?-{3,}:?$/.test(cell))) return;
    if (!current) current = [];
    current.push(cells);
  });
  if (current) tables.push(current);
  return tables;
}

export function leaveMarkdownSection(leave, english) {
  const payload = leaveSyncPayload(leave);
  const title = english ? "## Leave Balance" : "## رصيد الإجازات";
  const lines = [
    title,
    "",
    "| key | value |",
    "| --- | --- |",
    `| startMonth | ${payload.settings.startMonth} |`,
    `| initialBalance | ${payload.settings.initialBalance} |`,
    `| initialCarryOver | ${payload.settings.initialCarryOver} |`,
    `| monthlyAccrual | ${payload.settings.monthlyAccrual} |`,
    `| horizonMonths | ${payload.settings.horizonMonths} |`,
    "",
    "| date | note | source |",
    "| --- | --- | --- |",
  ];
  payload.entries.forEach((entry) => {
    const note = String(entry.note || "").replace(/\|/g, "/");
    lines.push(`| ${entry.date} | ${note} | ${entry.source} |`);
  });
  lines.push("", "| month | mode | used |", "| --- | --- | --- |");
  payload.months.forEach((item) => {
    lines.push(`| ${item.month} | ${item.mode} | ${item.used} |`);
  });
  lines.push("");
  return lines;
}

export function buildLeaveMarkdown(leave, english) {
  const title = english ? "# Leave balance" : "# رصيد الإجازات";
  return [
    title,
    "",
    `signature: ${LEAVE_SIGNATURE}`,
    "",
    ...leaveMarkdownSection(leave, english),
  ].join("\n");
}

export function leaveFromMarkdown(source) {
  const text = String(source || "");
  const match = /^## (?:رصيد الإجازات|Leave Balance)\s*$/m.exec(text);
  if (!match) return null;
  const fence = text.indexOf("```", match.index);
  const slice = text.slice(match.index, fence === -1 ? text.length : fence);
  const settings = {};
  const entries = [];
  const months = [];
  markdownTables(slice).forEach((table) => {
    const header = (table[0] || []).map((cell) => cell.toLowerCase());
    const body = table.slice(1);
    if (header.includes("key") || header.includes("المفتاح")) {
      const keyAt = Math.max(header.indexOf("key"), header.indexOf("المفتاح"));
      const valueAt = Math.max(header.indexOf("value"), header.indexOf("القيمة"));
      body.forEach((row) => {
        const key = cellText(row[keyAt]);
        if (key) settings[key] = row[valueAt === -1 ? 1 : valueAt];
      });
      return;
    }
    if (header[0] === "date") {
      const sourceAt = header.indexOf("source");
      body.forEach((row) => entries.push({
        date: row[0],
        note: row[1] || "",
        source: sourceAt >= 0 ? row[sourceAt] : "",
      }));
      return;
    }
    if (header[0] === "month") {
      body.forEach((row) => months.push({ month: row[0], mode: row[1], used: row[2] }));
    }
  });
  return sanitizeLeave({ settings, entries, months });
}
