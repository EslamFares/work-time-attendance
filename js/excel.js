import {
  APP_SIGNATURE,
  LABEL_TO_STATUS,
  LOG_HEADERS_AR,
  LOG_HEADERS_EN,
  LOG_KEYS,
  STATUS_LABELS,
  STATUS_LABELS_EN,
  defaultSettings,
  hhmmFromMinutes,
  minutesFromHHMM,
  sanitizeSettings,
  validateImportPayload,
} from "./calc.js";
import {
  LEAVE_SHEET_ALIASES,
  buildLeaveMarkdown,
  leaveFromMarkdown,
  leaveFromRows,
  leaveMarkdownSection,
  leaveSheetRows,
  leaveSyncPayload,
  sanitizeLeave,
} from "./leave.js";

function sheetFromRows(XLSX, rows, textCols) {
  const ws = {};
  let maxC = 0;
  rows.forEach((row, r) => {
    row.forEach((value, c) => {
      maxC = Math.max(maxC, c);
      const addr = XLSX.utils.encode_cell({ r, c });
      const forceText = textCols === "all" || (Array.isArray(textCols) && textCols.includes(c));
      if (!forceText && typeof value === "number") ws[addr] = { t: "n", v: value };
      else ws[addr] = { t: "s", v: value == null ? "" : String(value) };
    });
  });
  ws["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: Math.max(0, rows.length - 1), c: maxC },
  });
  return ws;
}

function settingsForExport(settings) {
  const copy = { ...(settings || {}) };
  delete copy.fileSnapshot;
  return copy;
}

const SETTINGS_FIELDS = [
  { key: "officialCheckIn", kind: "text", ar: "الميعاد الرسمي للحضور", en: "Official check-in", read: (s) => s.officialCheckIn },
  { key: "graceEnd", kind: "text", ar: "نهاية فترة السماح", en: "Grace ends", read: (s) => s.graceEnd },
  { key: "dailyTargetMinutes", kind: "number", ar: "ساعات اليوم المطلوبة (دقائق)", en: "Daily target (minutes)", read: (s) => s.dailyTargetMinutes },
  { key: "defaultPermissionMinutes", kind: "number", ar: "مدة الإذن الافتراضية (دقائق)", en: "Default permission (minutes)", read: (s) => s.defaultPermissionMinutes },
  { key: "weekendDays", kind: "days", ar: "أيام الإجازة الأسبوعية (0=أحد … 6=سبت)", en: "Weekend days (0=Sun … 6=Sat)", read: (s) => s.weekendDays },
  { key: "dhuhr.enabled", kind: "bool", ar: "تفعيل بريك الظهر تلقائيًا", en: "Dhuhr break on by default", read: (s) => s.dhuhr?.enabled },
  { key: "dhuhr.start", kind: "text", ar: "بداية بريك الظهر", en: "Dhuhr break starts", read: (s) => s.dhuhr?.start },
  { key: "dhuhr.duration", kind: "number", ar: "مدة بريك الظهر (دقائق)", en: "Dhuhr break minutes", read: (s) => s.dhuhr?.duration },
  { key: "asr.enabled", kind: "bool", ar: "تفعيل بريك العصر تلقائيًا", en: "Asr break on by default", read: (s) => s.asr?.enabled },
  { key: "asr.start", kind: "text", ar: "بداية بريك العصر", en: "Asr break starts", read: (s) => s.asr?.start },
  { key: "asr.duration", kind: "number", ar: "مدة بريك العصر (دقائق)", en: "Asr break minutes", read: (s) => s.asr?.duration },
  { key: "notifications.preBreak.enabled", kind: "bool", ar: "تنبيه قبل البريك", en: "Alert before a break", read: (s) => s.notifications?.preBreak?.enabled },
  { key: "notifications.preBreak.minutesBefore", kind: "number", ar: "دقائق التنبيه قبل البريك", en: "Minutes before a break", read: (s) => s.notifications?.preBreak?.minutesBefore },
  { key: "notifications.preBreak.blink", kind: "bool", ar: "وميض التابة قبل البريك", en: "Blink tab before a break", read: (s) => s.notifications?.preBreak?.blink },
  { key: "notifications.breakEnd.enabled", kind: "bool", ar: "تنبيه انتهاء البريك", en: "Alert when a break ends", read: (s) => s.notifications?.breakEnd?.enabled },
  { key: "notifications.breakEnd.blink", kind: "bool", ar: "وميض التابة عند انتهاء البريك", en: "Blink tab when a break ends", read: (s) => s.notifications?.breakEnd?.blink },
  { key: "notifications.shiftEnd.enabled", kind: "bool", ar: "تنبيه وقت الانصراف", en: "Alert at check-out", read: (s) => s.notifications?.shiftEnd?.enabled },
  { key: "notifications.shiftEnd.minutesBefore", kind: "number", ar: "دقائق التنبيه قبل الانصراف", en: "Minutes before check-out", read: (s) => s.notifications?.shiftEnd?.minutesBefore },
  { key: "notifications.shiftEnd.blink", kind: "bool", ar: "وميض التابة عند الانصراف", en: "Blink tab at check-out", read: (s) => s.notifications?.shiftEnd?.blink },
  { key: "notifications.permission.enabled", kind: "bool", ar: "تنبيه وقت الإذن", en: "Alert at permission time", read: (s) => s.notifications?.permission?.enabled },
  { key: "notifications.permission.blink", kind: "bool", ar: "وميض التابة عند الإذن", en: "Blink tab at permission time", read: (s) => s.notifications?.permission?.blink },
  { key: "timePicker", kind: "text", ar: "شكل اختيار الوقت", en: "Time picker style", read: (s) => s.timePicker },
  { key: "showHeaderTarget", kind: "bool", ar: "إظهار التارجت اليومي في الأعلى", en: "Show the daily target in the header", read: (s) => s.showHeaderTarget },
  { key: "progressShimmer", kind: "bool", ar: "وميض شريط التقدم", en: "Progress bar shimmer", read: (s) => s.progressShimmer },
  { key: "theme", kind: "text", ar: "المظهر", en: "Theme", read: (s) => s.theme },
  { key: "accent", kind: "text", ar: "لون الواجهة", en: "Accent color", read: (s) => s.accent },
  { key: "lang", kind: "text", ar: "اللغة", en: "Language", read: (s) => s.lang },
  { key: "soundEnabled", kind: "bool", ar: "الصوت", en: "Sound", read: (s) => s.soundEnabled },
  { key: "tabBlinkEnabled", kind: "bool", ar: "وميض عنوان التابة", en: "Tab title blink", read: (s) => s.tabBlinkEnabled },
  { key: "alertSeconds", kind: "number", ar: "مدة تنبيه الرسالة السريعة (ثواني)", en: "Quick message duration (seconds)", read: (s) => s.alertSeconds },
  { key: "sheetsWebhook", kind: "text", ar: "رابط مزامنة Google Sheets", en: "Google Sheets sync link", read: (s) => s.sheetsWebhook },
];

const FIELD_BY_KEY = new Map(SETTINGS_FIELDS.map((field) => [field.key, field]));

function formatSettingValue(field, raw) {
  if (field.kind === "bool") return raw ? "true" : "false";
  if (field.kind === "number") return String(Math.round(Number(raw) || 0));
  if (field.kind === "days") return (Array.isArray(raw) ? raw : []).join(",");
  return raw == null ? "" : String(raw);
}

export function settingsPairs(settings) {
  const english = settings?.lang === "en";
  return SETTINGS_FIELDS.map((field) => ({
    key: field.key,
    value: formatSettingValue(field, field.read(settings || {})),
    label: english ? field.en : field.ar,
  }));
}

export function settingsSheetRows(settings) {
  const english = settings?.lang === "en";
  return [
    [english ? "label" : "البيان", english ? "value" : "القيمة", "key"],
    ...settingsPairs(settings).map((row) => [row.label, row.value, row.key]),
  ];
}

function coerceSetting(key, value) {
  const field = FIELD_BY_KEY.get(key);
  if (!field) return value;
  if (field.kind === "bool") {
    if (typeof value === "boolean") return value;
    if (typeof value === "number") return value !== 0;
    const text = String(value).trim().toLowerCase();
    return text === "true" || text === "1" || text === "yes" || text === "نعم";
  }
  if (field.kind === "number") {
    const number = typeof value === "number" ? value : Number(String(value).trim());
    return Number.isFinite(number) ? number : value;
  }
  if (field.kind === "days") {
    const parts = Array.isArray(value) ? value : String(value).split(/[,،\s]+/);
    return parts.map(Number).filter((n) => Number.isInteger(n) && n >= 0 && n <= 6);
  }
  return value == null ? "" : String(value).trim();
}

export function pairsFromRows(rows) {
  const list = Array.isArray(rows) ? rows : [];
  const pairs = [];
  let keyIndex = 0;
  let valueIndex = 1;
  list.forEach((row, index) => {
    if (row && typeof row === "object" && !Array.isArray(row)) {
      const key = String(row.key ?? "").trim();
      if (key && key !== "key") pairs.push({ key, value: row.value });
      return;
    }
    if (!Array.isArray(row)) return;
    const cells = row.map((cell) => String(cell ?? "").trim());
    if (index === 0) {
      const keyAt = cells.findIndex((cell) => cell === "key" || cell === "المفتاح");
      const valueAt = cells.findIndex((cell) => cell === "value" || cell === "القيمة");
      if (keyAt >= 0) {
        keyIndex = keyAt;
        if (valueAt >= 0) valueIndex = valueAt;
        return;
      }
    }
    const key = cells[keyIndex];
    if (!key || key === "key" || key === "المفتاح") return;
    pairs.push({ key, value: row[valueIndex] });
  });
  return pairs;
}

function nestSettings(pairs) {
  const target = {};
  pairs.forEach((pair) => {
    const key = String(pair.key || "").trim();
    if (!FIELD_BY_KEY.has(key)) return;
    const parts = key.split(".");
    let node = target;
    for (let i = 0; i < parts.length - 1; i += 1) {
      if (!node[parts[i]] || typeof node[parts[i]] !== "object") node[parts[i]] = {};
      node = node[parts[i]];
    }
    node[parts[parts.length - 1]] = coerceSetting(key, pair.value);
  });
  return target;
}

function mergeSettings(base, patch) {
  const next = { ...(base || {}), ...patch };
  if (patch.dhuhr) next.dhuhr = { ...(base?.dhuhr || {}), ...patch.dhuhr };
  if (patch.asr) next.asr = { ...(base?.asr || {}), ...patch.asr };
  if (patch.notifications) {
    next.notifications = { ...(base?.notifications || {}) };
    Object.keys(patch.notifications).forEach((key) => {
      next.notifications[key] = { ...(base?.notifications?.[key] || {}), ...patch.notifications[key] };
    });
  }
  return next;
}

export function overlaySettings(base, rows, { keepLink = false } = {}) {
  const nested = nestSettings(pairsFromRows(rows));
  if (!Object.keys(nested).length) return base;
  const fallback = sanitizeSettings(base, defaultSettings());
  const settings = sanitizeSettings(mergeSettings(fallback, nested), fallback);
  if (keepLink && base) {
    const localWebhook = typeof base.sheetsWebhook === "string" ? base.sheetsWebhook.trim() : "";
    settings.sheetsWebhook = localWebhook || settings.sheetsWebhook;
    settings.sheetsConnected = base.sheetsConnected === true;
    settings.sheetsImport = base.sheetsImport === "keep" || base.sheetsImport === "replace" ? base.sheetsImport : "";
    settings.fileSnapshot = base.fileSnapshot === true;
  }
  return settings;
}

function markdownTableRows(source) {
  const rows = [];
  const lines = String(source || "").split(/\r?\n/);
  let started = false;
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed.startsWith("|")) {
      if (started) break;
      continue;
    }
    const cells = trimmed.split("|").slice(1, -1).map((cell) => cell.trim());
    if (!started) {
      if (!cells.includes("key")) continue;
      started = true;
    }
    if (cells.every((cell) => /^:?-{3,}:?$/.test(cell))) continue;
    rows.push(cells);
  }
  return rows;
}

export function buildSheets(state) {
  const english = state.settings?.lang === "en";
  const metaRows = [
    ["signature", APP_SIGNATURE],
    ["settings", JSON.stringify(settingsForExport(state.settings))],
    ["today", JSON.stringify(state.today)],
    ["leave", JSON.stringify(leaveSyncPayload(state.leave))],
  ];
  const logRows = [
    english ? LOG_HEADERS_EN : LOG_HEADERS_AR,
    LOG_KEYS,
    ...state.log.map((row) => LOG_KEYS.map((key) => {
      if (key === "status") {
        const labels = english ? STATUS_LABELS_EN : STATUS_LABELS;
        return labels[row.status] || row.status || "";
      }
      const value = row[key];
      if (["breakMinutes", "permissionMinutes", "grossMinutes", "netMinutes", "targetMinutes", "overtimeMinutes"].includes(key)) {
        return Number(value) || 0;
      }
      return value ?? "";
    })),
  ];
  return { metaRows, logRows };
}

export function normalizeTimeCell(value) {
  if (value == null || value === "") return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return hhmmFromMinutes(value.getHours() * 60 + value.getMinutes());
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    let frac = value;
    if (Math.abs(frac) > 1.5) return "";
    frac %= 1;
    if (frac < 0) frac += 1;
    return hhmmFromMinutes(Math.round(frac * 1440) % 1440);
  }
  const text = String(value).trim();
  const match = /^(\d{1,2}):(\d{2})(?::\d{2})?(?:\s*([صم]))?$/.exec(text);
  if (!match) return text;
  let hours = Number(match[1]);
  const mins = Number(match[2]);
  if (match[3] === "م" && hours < 12) hours += 12;
  if (match[3] === "ص" && hours === 12) hours = 0;
  if (hours > 23 || mins > 59) return text;
  return hhmmFromMinutes(hours * 60 + mins);
}

function normalizeDateCell(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (typeof value === "number" && value > 20000 && value < 80000) {
    const utc = new Date(Math.round((value - 25569) * 86400 * 1000));
    const y = utc.getUTCFullYear();
    const m = String(utc.getUTCMonth() + 1).padStart(2, "0");
    const d = String(utc.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  const text = String(value ?? "").trim();
  const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(text);
  return iso ? `${iso[1]}-${iso[2]}-${iso[3]}` : text;
}

function cell(row, index) {
  if (!row) return "";
  return row[index] == null ? "" : row[index];
}

export function parseSheetRows(metaRows, logRows) {
  const fail = () => ({ ok: false, error: "فشل الاستيراد", payload: null });
  if (!Array.isArray(metaRows) || !Array.isArray(logRows)) return fail();
  const meta = new Map();
  for (const row of metaRows) {
    if (!row || row[0] == null || row[0] === "") continue;
    meta.set(String(row[0]).trim(), row[1]);
  }
  if (meta.get("signature") !== APP_SIGNATURE) return fail();

  let settings;
  let today;
  try {
    settings = JSON.parse(String(meta.get("settings") ?? ""));
    today = JSON.parse(String(meta.get("today") ?? ""));
  } catch {
    return fail();
  }

  const keyRow = (logRows[1] || []).map((value) => String(value ?? "").trim());
  if (LOG_KEYS.some((key, index) => keyRow[index] !== key)) return fail();

  const log = [];
  for (let i = 2; i < logRows.length; i += 1) {
    const row = logRows[i];
    if (!row || row.every((value) => value == null || String(value).trim() === "")) continue;
    const record = {};
    LOG_KEYS.forEach((key, index) => {
      record[key] = cell(row, index);
    });
    record.date = normalizeDateCell(record.date);
    record.checkIn = normalizeTimeCell(record.checkIn);
    record.checkOut = normalizeTimeCell(record.checkOut);
    if (LABEL_TO_STATUS[record.status]) record.status = LABEL_TO_STATUS[record.status];
    log.push(record);
  }

  const payload = { signature: APP_SIGNATURE, settings, today, log };
  if (meta.has("leave")) {
    try {
      const leave = JSON.parse(String(meta.get("leave") ?? ""));
      if (leave && typeof leave === "object" && !Array.isArray(leave)) payload.leave = leave;
    } catch {
      /* a broken leave copy still imports the working-hours log */
    }
  }
  if (!validateImportPayload(payload).ok) return fail();
  return { ok: true, error: null, payload };
}

export function buildWorkbook(XLSX, state) {
  const { metaRows, logRows } = buildSheets(state);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheetFromRows(XLSX, logRows, [0, 1, 2, 3, 6]), "Log");
  XLSX.utils.book_append_sheet(wb, sheetFromRows(XLSX, settingsSheetRows(state.settings), "all"), "Settings");
  XLSX.utils.book_append_sheet(wb, sheetFromRows(XLSX, leaveSheetRows(state.leave), "all"), "Leave Balance");
  XLSX.utils.book_append_sheet(wb, sheetFromRows(XLSX, metaRows, "all"), "Meta");
  return wb;
}

export function payloadFromWorkbook(XLSX, workbook) {
  if (!workbook?.Sheets?.Meta || !workbook?.Sheets?.Log) return { ok: false, error: "فشل الاستيراد", payload: null };
  const metaRows = XLSX.utils.sheet_to_json(workbook.Sheets.Meta, { header: 1, raw: true, defval: "" });
  const logRows = XLSX.utils.sheet_to_json(workbook.Sheets.Log, { header: 1, raw: true, defval: "" });
  const parsed = parseSheetRows(metaRows, logRows);
  const settingsSheet = workbook.Sheets?.Settings || workbook.Sheets?.["الإعدادات"];
  if (parsed.ok && settingsSheet) {
    const settingsRows = XLSX.utils.sheet_to_json(settingsSheet, { header: 1, raw: true, defval: "" });
    parsed.payload.settings = overlaySettings(parsed.payload.settings, settingsRows);
  }
  if (!parsed.ok) return parsed;
  const leaveSheet = LEAVE_SHEET_ALIASES.map((name) => workbook.Sheets?.[name]).find(Boolean);
  if (leaveSheet) {
    const leaveRows = XLSX.utils.sheet_to_json(leaveSheet, { header: 1, raw: true, defval: "" });
    const leave = leaveFromRows(leaveRows);
    if (leave) parsed.payload.leave = leave;
  }
  return parsed;
}

export function buildLeaveWorkbook(XLSX, leave) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheetFromRows(XLSX, leaveSheetRows(leave), "all"), "Leave Balance");
  return wb;
}

export function leaveFromWorkbook(XLSX, workbook) {
  const sheet = LEAVE_SHEET_ALIASES.map((name) => workbook?.Sheets?.[name]).find(Boolean);
  if (!sheet) return null;
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
  return leaveFromRows(rows);
}

export function downloadLeaveWorkbook(state) {
  const XLSX = globalThis.XLSX;
  if (!XLSX) throw new Error("مكتبة إكسل غير متاحة");
  const english = state.settings?.lang === "en";
  const name = english
    ? `leave-balance-${state.today.date}.xlsx`
    : `رصيد-الإجازات-${state.today.date}.xlsx`;
  XLSX.writeFile(buildLeaveWorkbook(XLSX, state.leave), name);
}

export function downloadLeaveMarkdown(state) {
  const english = state.settings?.lang === "en";
  const name = english
    ? `leave-balance-${state.today.date}.md`
    : `رصيد-الإجازات-${state.today.date}.md`;
  const blob = new Blob([buildLeaveMarkdown(state.leave, english)], { type: "text/markdown;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

export function downloadWorkbook(state) {
  const XLSX = globalThis.XLSX;
  if (!XLSX) throw new Error("مكتبة إكسل غير متاحة");
  const name = state.settings?.lang === "en"
    ? `worktime-${state.today.date}.xlsx`
    : `سجل-ساعات-العمل-${state.today.date}.xlsx`;
  XLSX.writeFile(buildWorkbook(XLSX, state), name);
}

export function readWorkbook(arrayBuffer) {
  const XLSX = globalThis.XLSX;
  if (!XLSX) return { ok: false, error: "فشل الاستيراد", payload: null };
  try {
    return payloadFromWorkbook(XLSX, XLSX.read(arrayBuffer, { type: "array", cellDates: true }));
  } catch {
    return { ok: false, error: "فشل الاستيراد", payload: null };
  }
}

export function minutesOrBlank(value) {
  const mins = minutesFromHHMM(value);
  return mins == null ? "" : hhmmFromMinutes(mins);
}

function backupPayload(state) {
  return {
    signature: APP_SIGNATURE,
    settings: settingsForExport(state.settings),
    today: state.today,
    log: state.log,
    leave: leaveSyncPayload(state.leave),
  };
}

function cellText(value) {
  return String(value ?? "").replace(/\|/g, "/").replace(/\s+/g, " ").trim();
}

export function buildMarkdown(state) {
  const english = state.settings?.lang === "en";
  const headers = english ? LOG_HEADERS_EN : LOG_HEADERS_AR;
  const labels = english ? STATUS_LABELS_EN : STATUS_LABELS;
  const log = Array.isArray(state.log) ? state.log : [];
  const rows = log.map((row) => `| ${LOG_KEYS.map((key) => {
    if (key === "status") return cellText(labels[row.status] || row.status || "");
    return cellText(row[key]);
  }).join(" | ")} |`);
  const settingsTable = settingsSheetRows(state.settings);
  const lines = [
    english ? "# Work hours" : "# ساعات العمل",
    "",
    `signature: ${APP_SIGNATURE}`,
    "",
    english
      ? "The settings table is saved in this file. Sync copies the same values to every device."
      : "جدول الإعدادات جوه الملف. المزامنة تنقل نفس القيم لكل الأجهزة.",
    "",
    english ? "## Settings" : "## الإعدادات",
    "",
    `| ${settingsTable[0].join(" | ")} |`,
    `| ${settingsTable[0].map(() => "---").join(" | ")} |`,
    ...settingsTable.slice(1).map((row) => `| ${row.map((cell) => cellText(cell)).join(" | ")} |`),
    "",
    english ? "## Log" : "## السجل",
    "",
    `| ${headers.join(" | ")} |`,
    `| ${headers.map(() => "---").join(" | ")} |`,
    ...rows,
    "",
    ...leaveMarkdownSection(state.leave, english),
    "```worktime",
    JSON.stringify(backupPayload(state)),
    "```",
    "",
  ];
  return lines.join("\n");
}

export function readMarkdown(text) {
  const fail = () => ({ ok: false, error: "فشل الاستيراد", payload: null });
  const source = String(text ?? "").replace(/^\uFEFF/, "");
  const match = /```worktime\s*([\s\S]*?)```/.exec(source);
  if (!match) return fail();
  if (!source.includes(APP_SIGNATURE)) return fail();
  let payload;
  try {
    payload = JSON.parse(match[1]);
  } catch {
    return fail();
  }
  if (!validateImportPayload(payload).ok) return fail();
  const tableRows = markdownTableRows(source);
  if (tableRows.length) {
    const overlaid = overlaySettings(payload.settings, tableRows);
    const candidate = { ...payload, settings: overlaid };
    if (validateImportPayload(candidate).ok) payload = candidate;
  }
  const leaveSection = leaveFromMarkdown(source);
  if (leaveSection) payload.leave = leaveSection;
  return { ok: true, error: null, payload };
}

export function downloadMarkdown(state) {
  const name = state.settings?.lang === "en"
    ? `worktime-${state.today.date}.md`
    : `سجل-ساعات-العمل-${state.today.date}.md`;
  const blob = new Blob([buildMarkdown(state)], { type: "text/markdown;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

function looksLikeZip(buffer) {
  const bytes = new Uint8Array(buffer.slice(0, 2));
  return bytes[0] === 0x50 && bytes[1] === 0x4b;
}

export async function readLeaveFile(file) {
  const fail = () => ({ ok: false, error: "no-leave", leave: null });
  const accept = (leave) => (leave ? { ok: true, error: null, leave: sanitizeLeave(leave) } : fail());
  const name = String(file?.name || "").toLowerCase();
  if (name.endsWith(".md") || name.endsWith(".markdown")) return accept(leaveFromMarkdown(await file.text()));
  const buffer = await file.arrayBuffer();
  if (!looksLikeZip(buffer)) return accept(leaveFromMarkdown(new TextDecoder("utf-8").decode(buffer)));
  const XLSX = globalThis.XLSX;
  if (!XLSX) return fail();
  try {
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    const fromSheet = leaveFromWorkbook(XLSX, workbook);
    if (fromSheet) return accept(fromSheet);
    const parsed = payloadFromWorkbook(XLSX, workbook);
    if (parsed.ok && parsed.payload.leave) return accept(parsed.payload.leave);
  } catch {
    return fail();
  }
  return fail();
}

export async function readBackupFile(file) {
  const name = String(file?.name || "").toLowerCase();
  if (name.endsWith(".md") || name.endsWith(".markdown")) return readMarkdown(await file.text());
  const buffer = await file.arrayBuffer();
  if (!looksLikeZip(buffer)) return readMarkdown(new TextDecoder("utf-8").decode(buffer));
  return readWorkbook(buffer);
}
