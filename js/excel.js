import {
  APP_SIGNATURE,
  LABEL_TO_STATUS,
  LOG_HEADERS_AR,
  LOG_HEADERS_EN,
  LOG_KEYS,
  STATUS_LABELS,
  STATUS_LABELS_EN,
  hhmmFromMinutes,
  minutesFromHHMM,
  validateImportPayload,
} from "./calc.js";

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

export function buildSheets(state) {
  const english = state.settings?.lang === "en";
  const metaRows = [
    ["signature", APP_SIGNATURE],
    ["settings", JSON.stringify(state.settings)],
    ["today", JSON.stringify(state.today)],
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
  if (!validateImportPayload(payload).ok) return fail();
  return { ok: true, error: null, payload };
}

export function buildWorkbook(XLSX, state) {
  const { metaRows, logRows } = buildSheets(state);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheetFromRows(XLSX, logRows, [0, 1, 2, 3, 6]), "Log");
  XLSX.utils.book_append_sheet(wb, sheetFromRows(XLSX, metaRows, "all"), "Meta");
  return wb;
}

export function payloadFromWorkbook(XLSX, workbook) {
  if (!workbook?.Sheets?.Meta || !workbook?.Sheets?.Log) return { ok: false, error: "فشل الاستيراد", payload: null };
  const metaRows = XLSX.utils.sheet_to_json(workbook.Sheets.Meta, { header: 1, raw: true, defval: "" });
  const logRows = XLSX.utils.sheet_to_json(workbook.Sheets.Log, { header: 1, raw: true, defval: "" });
  return parseSheetRows(metaRows, logRows);
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
