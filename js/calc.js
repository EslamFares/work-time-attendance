/**
 * Pure attendance calculations.
 * Times are minutes from local midnight. UI strings stay outside this module
 * except for stable Arabic status labels stored in the monthly log.
 */

export const APP_SIGNATURE = "WORKTIME_ATTENDANCE_V1";

export const STATUS = {
  work: "work",
  weekend: "weekend",
  official: "official",
  personal: "personal",
};

export const STATUS_LABELS = {
  work: "عمل",
  weekend: "إجازة أسبوعية",
  official: "إجازة رسمية",
  personal: "إجازة شخصية",
};

export const STATUS_LABELS_EN = {
  work: "Work",
  weekend: "Weekly off",
  official: "Official holiday",
  personal: "Personal leave",
};

export const LABEL_TO_STATUS = Object.fromEntries([
  ...Object.entries(STATUS_LABELS).map(([key, label]) => [label, key]),
  ...Object.entries(STATUS_LABELS_EN).map(([key, label]) => [label, key]),
]);

export const WEEKDAYS_AR = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"];
export const MONTHS_AR = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"];

export const LOG_KEYS = [
  "date",
  "status",
  "checkIn",
  "checkOut",
  "breakMinutes",
  "permissionMinutes",
  "permissionType",
  "grossMinutes",
  "netMinutes",
  "targetMinutes",
  "overtimeMinutes",
];

export const LOG_HEADERS_EN = [
  "Date",
  "Status",
  "Check-in",
  "Check-out",
  "Break minutes",
  "Permission minutes",
  "Permission type",
  "Gross",
  "Net",
  "Target",
  "Overtime",
];

export const LOG_HEADERS_AR = [
  "التاريخ",
  "الحالة",
  "الحضور",
  "الانصراف",
  "دقائق البريك",
  "دقائق الإذن",
  "نوع الإذن",
  "الإجمالي",
  "الصافي",
  "الهدف",
  "الإضافي",
];

const PERMISSION_TYPES = new Set(["auto", "morning", "evening", "midday"]);

export function clamp(value, min, max) {
  const n = Number(value);
  if (!Number.isFinite(n)) return min;
  return Math.min(max, Math.max(min, n));
}

export function clampInt(value, min, max) {
  return Math.round(clamp(value, min, max));
}

export function isoDateLocal(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

export function shiftIso(isoDate, days) {
  const [y, m, d] = String(isoDate).split("-").map(Number);
  return isoDateLocal(new Date(y, m - 1, d + days));
}

export function weekdayIndex(isoDate) {
  const [y, m, d] = String(isoDate).split("-").map(Number);
  return new Date(y, m - 1, d).getDay();
}

export function leaveEntry(isoDate, weekendDays = [5, 6]) {
  const weekend = weekendDays.includes(weekdayIndex(isoDate));
  return {
    date: isoDate,
    status: weekend ? STATUS.weekend : STATUS.official,
    checkIn: "",
    checkOut: "",
    breakMinutes: 0,
    permissionMinutes: 0,
    permissionType: "",
    grossMinutes: 0,
    netMinutes: 0,
    targetMinutes: 0,
    overtimeMinutes: 0,
  };
}

export function fillMissingDays(state) {
  const today = state?.today?.date;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(today || "")) return state;
  const weekendDays = state.settings?.weekendDays?.length ? state.settings.weekendDays : [5, 6];
  const dismissed = new Set(state.dismissedDates || []);
  const have = new Set((state.log || []).map((row) => row.date));
  let start = null;
  for (const date of have) {
    if (date < today && (!start || date < start)) start = date;
  }
  if (!start || start >= today) return state;
  const log = (state.log || []).slice();
  let changed = false;
  for (let date = start; date < today; date = shiftIso(date, 1)) {
    if (have.has(date) || dismissed.has(date)) continue;
    log.push(leaveEntry(date, weekendDays));
    have.add(date);
    changed = true;
  }
  if (!changed) return state;
  log.sort((a, b) => a.date.localeCompare(b.date));
  return { ...state, log };
}

export function weekdayName(isoDate) {
  return WEEKDAYS_AR[weekdayIndex(isoDate)] || "";
}

export function minutesFromHHMM(value) {
  if (typeof value !== "string") return null;
  const match = /^(\d{1,2}):(\d{2})$/.exec(value.trim());
  if (!match) return null;
  const hours = Number(match[1]);
  const mins = Number(match[2]);
  if (hours < 0 || hours > 23 || mins < 0 || mins > 59) return null;
  return hours * 60 + mins;
}

export function hhmmFromMinutes(totalMinutes) {
  if (totalMinutes == null || !Number.isFinite(totalMinutes)) return "";
  let mins = Math.round(totalMinutes);
  mins = ((mins % 1440) + 1440) % 1440;
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

export function timeParts(totalMinutes) {
  if (totalMinutes == null || !Number.isFinite(totalMinutes)) {
    return { display: "—", period: "", extra: 0 };
  }
  let extra = 0;
  let mins = Math.round(totalMinutes);
  if (mins < 0) {
    while (mins < 0) {
      mins += 1440;
      extra -= 1;
    }
  }
  extra += Math.floor(mins / 1440);
  mins %= 1440;
  const h24 = Math.floor(mins / 60);
  const m = mins % 60;
  let h12 = h24 % 12;
  if (h12 === 0) h12 = 12;
  return {
    display: `${String(h12).padStart(2, "0")}:${String(m).padStart(2, "0")}`,
    period: h24 < 12 ? "ص" : "م",
    extra,
  };
}

export function formatTime12(totalMinutes) {
  const parts = timeParts(totalMinutes);
  if (parts.display === "—") return "—";
  const extra = parts.extra > 0 ? ` +${parts.extra}` : parts.extra < 0 ? ` ${parts.extra}` : "";
  return `${parts.display} ${parts.period}${extra}`;
}

export function formatDuration(mins) {
  if (mins == null || !Number.isFinite(mins)) return "—";
  const sign = mins < 0 ? "−" : "";
  const total = Math.abs(Math.round(mins));
  const h = Math.floor(total / 60);
  const m = total % 60;
  if (h === 0) return `${sign}${m} د`;
  if (m === 0) return `${sign}${h} س`;
  return `${sign}${h} س ${m} د`;
}

export function formatClock(date) {
  let h24 = date.getHours();
  const m = date.getMinutes();
  const s = date.getSeconds();
  const period = h24 < 12 ? "ص" : "م";
  let h12 = h24 % 12;
  if (h12 === 0) h12 = 12;
  const display = [h12, m, s].map((n, i) => (i === 0 ? String(n).padStart(2, "0") : String(n).padStart(2, "0"))).join(":");
  return { display, period };
}

export function formatLongDate(isoDate) {
  const [y, m, d] = isoDate.split("-").map(Number);
  return `${weekdayName(isoDate)}، ${d} ${MONTHS_AR[m - 1]} ${y}`;
}

export function defaultStatusForDate(isoDate, weekendDays = [5, 6]) {
  return weekendDays.includes(weekdayIndex(isoDate)) ? STATUS.weekend : STATUS.work;
}

export function resolveDayStatus(isoDate, override, weekendDays = [5, 6]) {
  if (override === STATUS.work || override === STATUS.weekend || override === STATUS.official || override === STATUS.personal) {
    return override;
  }
  return defaultStatusForDate(isoDate, weekendDays);
}

export function defaultSettings() {
  return {
    officialCheckIn: "08:30",
    graceEnd: "09:30",
    dailyTargetMinutes: 480,
    defaultPermissionMinutes: 120,
    theme: "dark",
    accent: "emerald",
    lang: "ar",
    timePicker: "clock",
    progressShimmer: true,
    showHeaderTarget: true,
    soundEnabled: true,
    tabBlinkEnabled: true,
    weekendDays: [5, 6],
    dhuhr: { enabled: true, start: "11:50", duration: 45 },
    asr: { enabled: true, start: "14:40", duration: 15 },
    notifications: {
      preBreak: { enabled: true, minutesBefore: 10, blink: true },
      breakEnd: { enabled: true, blink: true },
      shiftEnd: { enabled: true, minutesBefore: 0, blink: true },
      permission: { enabled: true, blink: false },
    },
    sheetsWebhook: "",
    sheetsConnected: false,
    sheetsImport: "",
    fileSnapshot: false,
  };
}

export function createToday(dateISO, defaults = {}) {
  return {
    date: dateISO,
    statusOverride: null,
    workplace: "office",
    checkIn: null,
    checkOut: null,
    targetMinutes: null,
    permissionOn: false,
    permissionTouched: false,
    permissionMinutes: defaults.permissionMinutes ?? 120,
    permissionType: "auto",
    permissionTime: null,
    manualBreakMinutes: 0,
    manualBreakStart: null,
    dhuhrEnabled: null,
    asrEnabled: null,
    lastSeenAt: null,
  };
}

export function defaultState(dateISO) {
  const settings = defaultSettings();
  return {
    version: 1,
    settings,
    today: createToday(dateISO, { permissionMinutes: settings.defaultPermissionMinutes }),
    log: [],
    dismissedDates: [],
    notifiedKeys: [],
    lastSyncAt: null,
  };
}

/**
 * Permission is active when the user arms it, picks a concrete type,
 * sets a permission time, or arrives after the grace window while still on automatic.
 * Turning the switch off after a touch suppresses even the late-arrival permission.
 */
export function resolvePermission({
  checkInMin,
  graceEndMin,
  permissionType = "auto",
  permissionOn = false,
  permissionTouched = false,
  permissionTime = null,
}) {
  if (permissionType === "off") return { active: false, type: "none", automatic: false };
  const late = checkInMin != null && graceEndMin != null && checkInMin > graceEndMin;
  const concrete = permissionType === "morning" || permissionType === "evening" || permissionType === "midday";
  let active = false;
  if (permissionTouched) active = !!permissionOn;
  else if (concrete) active = true;
  else if (permissionTime) active = true;
  else active = late;

  if (!active) return { active: false, type: "none", automatic: permissionType === "auto", late };
  let type = permissionType;
  const automatic = !concrete;
  if (automatic) type = late ? "morning" : "evening";
  return { active: true, type, automatic, late };
}

export function creditedStartMinutes({ checkInMin, officialMin, graceEndMin, permission }) {
  if (checkInMin == null) return null;
  const morning = permission?.active && permission.type === "morning" && checkInMin > graceEndMin;
  if (morning) return Math.max(officialMin, checkInMin - (permission.minutes || 0));
  return checkInMin;
}

export function expectedCheckoutMinutes({ creditedStart, targetMinutes, permission }) {
  if (creditedStart == null) return null;
  let end = creditedStart + targetMinutes;
  if (permission?.active && permission.type === "evening") end -= permission.minutes || 0;
  if (end < creditedStart) end = creditedStart;
  return end;
}

export function evaluateScheduledBreak({
  id,
  enabled,
  start,
  duration,
  checkInMin,
  nowMin,
  checkOutMin,
  expectedCheckout,
  isWorkDay,
}) {
  const startMin = minutesFromHHMM(start);
  const durationMin = clampInt(duration, 0, 600);
  const base = {
    id,
    enabled: !!enabled,
    startMin,
    durationMin,
    elapsed: 0,
    remaining: durationMin,
    progress: 0,
    applied: false,
    phase: "off",
    reason: "disabled",
  };
  if (!isWorkDay) return { ...base, reason: "not-work" };
  if (!enabled || startMin == null) return base;
  if (checkInMin == null) return { ...base, phase: "idle", reason: "no-checkin" };
  if (checkInMin > startMin) return { ...base, phase: "skipped", reason: "late", remaining: 0 };
  if (checkOutMin != null && checkOutMin <= startMin) {
    return { ...base, phase: "skipped", reason: "left-before", remaining: 0 };
  }
  if (checkOutMin == null && nowMin < startMin && expectedCheckout != null && expectedCheckout <= startMin) {
    return { ...base, phase: "skipped", reason: "shift-ends-before", remaining: 0 };
  }

  const evalMin = checkOutMin != null ? checkOutMin : nowMin;
  if (evalMin <= startMin) {
    return { ...base, applied: true, phase: "upcoming", reason: "scheduled", remaining: durationMin };
  }
  const elapsed = Math.min(durationMin, Math.max(0, evalMin - startMin));
  const remaining = Math.max(0, durationMin - elapsed);
  const progress = durationMin === 0 ? 1 : Math.min(1, elapsed / durationMin);
  const phase = remaining <= 0 ? "done" : "active";
  return { ...base, applied: true, phase, reason: phase, elapsed, remaining, progress };
}

export function evaluateManualBreak({ minutes, start, checkOutMin, nowMin, isWorkDay }) {
  const durationMin = clampInt(minutes || 0, 0, 600);
  const startMin = typeof start === "string" ? minutesFromHHMM(start) : null;
  const base = {
    id: "manual",
    enabled: durationMin > 0 && startMin != null,
    startMin,
    durationMin,
    elapsed: 0,
    remaining: durationMin,
    progress: 0,
    applied: false,
    phase: "idle",
    reason: "idle",
  };
  if (!isWorkDay || durationMin <= 0 || startMin == null) return base;
  if (checkOutMin != null && checkOutMin <= startMin) {
    return { ...base, phase: "skipped", reason: "left-before", remaining: 0 };
  }
  const evalMin = checkOutMin != null ? checkOutMin : nowMin;
  if (evalMin <= startMin) {
    return { ...base, applied: true, phase: "upcoming", reason: "scheduled", remaining: durationMin };
  }
  const elapsed = Math.min(durationMin, Math.max(0, evalMin - startMin));
  const remaining = Math.max(0, durationMin - elapsed);
  const progress = durationMin === 0 ? 1 : Math.min(1, elapsed / durationMin);
  const phase = remaining <= 0 ? "done" : "active";
  return { ...base, applied: true, phase, reason: phase, elapsed, remaining, progress };
}

export function workProgress({ creditedStart, expectedCheckout, nowMin }) {
  if (creditedStart == null || expectedCheckout == null) {
    return { ratio: 0, remaining: null, overtime: 0, started: false, complete: false, span: 0 };
  }
  const span = Math.max(0, expectedCheckout - creditedStart);
  if (nowMin <= creditedStart) {
    return { ratio: 0, remaining: span, overtime: 0, started: false, complete: false, span };
  }
  const elapsed = nowMin - creditedStart;
  if (span === 0) {
    const overtime = Math.max(0, nowMin - expectedCheckout);
    return { ratio: 1, remaining: 0, overtime, started: true, complete: true, span };
  }
  const remaining = expectedCheckout - nowMin;
  return {
    ratio: elapsed / span,
    remaining: Math.max(0, remaining),
    overtime: remaining < 0 ? -remaining : 0,
    started: true,
    complete: remaining <= 0,
    span,
  };
}

export function inputFromState(state, nowMin) {
  const today = state.today;
  const settings = state.settings;
  return {
    date: today.date,
    statusOverride: today.statusOverride,
    weekendDays: settings.weekendDays,
    checkIn: today.checkIn,
    checkOut: today.checkOut,
    nowMin,
    officialCheckIn: settings.officialCheckIn,
    graceEnd: settings.graceEnd,
    targetMinutes: today.targetMinutes ?? settings.dailyTargetMinutes,
    permissionOn: today.permissionOn,
    permissionTouched: today.permissionTouched,
    permissionMinutes: today.permissionMinutes,
    permissionType: today.permissionType,
    permissionTime: today.permissionTime,
    dhuhrEnabled: today.dhuhrEnabled ?? settings.dhuhr.enabled,
    dhuhrStart: settings.dhuhr.start,
    dhuhrDuration: settings.dhuhr.duration,
    asrEnabled: today.asrEnabled ?? settings.asr.enabled,
    asrStart: settings.asr.start,
    asrDuration: settings.asr.duration,
    manualMinutes: today.manualBreakMinutes,
    manualStart: today.manualBreakStart,
  };
}

export function computeDay(input) {
  const officialMin = minutesFromHHMM(input.officialCheckIn) ?? 8 * 60 + 30;
  const graceEndMin = minutesFromHHMM(input.graceEnd) ?? 9 * 60 + 30;
  const checkInMin = input.checkIn ? minutesFromHHMM(input.checkIn) : null;
  const checkOutMin = input.checkOut ? minutesFromHHMM(input.checkOut) : null;
  const nowMin = Number.isFinite(input.nowMin) ? input.nowMin : 0;
  const targetMinutes = clampInt(input.targetMinutes ?? 480, 30, 16 * 60);
  const permissionMinutes = clampInt(input.permissionMinutes ?? 0, 0, 12 * 60);
  const status = resolveDayStatus(input.date, input.statusOverride, input.weekendDays || [5, 6]);
  const isWorkDay = status === STATUS.work;
  const permissionTimeMin = input.permissionTime ? minutesFromHHMM(input.permissionTime) : null;

  const resolved = resolvePermission({
    checkInMin,
    graceEndMin,
    permissionType: input.permissionType,
    permissionOn: input.permissionOn,
    permissionTouched: input.permissionTouched,
    permissionTime: input.permissionTime || null,
  });
  const permission = {
    ...resolved,
    minutes: resolved.active ? permissionMinutes : 0,
    timeMin: permissionTimeMin,
  };

  const creditedStart = isWorkDay
    ? creditedStartMinutes({
        checkInMin,
        officialMin,
        graceEndMin,
        permission: { ...permission, minutes: permissionMinutes },
      })
    : null;

  let expectedCheckout = null;
  let checkoutClamped = false;
  if (creditedStart != null) {
    const raw = creditedStart + targetMinutes - (permission.active && permission.type === "evening" ? permissionMinutes : 0);
    if (raw < creditedStart) {
      expectedCheckout = creditedStart;
      checkoutClamped = true;
    } else {
      expectedCheckout = raw;
    }
  }

  const dhuhr = evaluateScheduledBreak({
    id: "dhuhr",
    enabled: !!input.dhuhrEnabled,
    start: input.dhuhrStart,
    duration: input.dhuhrDuration,
    checkInMin,
    nowMin,
    checkOutMin,
    expectedCheckout,
    isWorkDay,
  });
  const asr = evaluateScheduledBreak({
    id: "asr",
    enabled: !!input.asrEnabled,
    start: input.asrStart,
    duration: input.asrDuration,
    checkInMin,
    nowMin,
    checkOutMin,
    expectedCheckout,
    isWorkDay,
  });
  const manual = evaluateManualBreak({
    minutes: input.manualMinutes,
    start: input.manualStart,
    checkOutMin,
    nowMin,
    isWorkDay,
  });

  const breakList = [dhuhr, asr, manual];
  const breakConsumed = breakList.reduce((sum, item) => sum + (item.applied ? item.elapsed : 0), 0);
  const breakAllowance = breakList.reduce((sum, item) => sum + (item.applied ? item.durationMin : 0), 0);

  const evalMin = !isWorkDay ? null : checkOutMin != null ? checkOutMin : nowMin;
  let gross = 0;
  let net = 0;
  if (creditedStart != null && evalMin != null) {
    gross = Math.max(0, evalMin - creditedStart);
    net = Math.max(0, gross - breakConsumed);
  }

  const progressNow = checkOutMin != null ? checkOutMin : nowMin;
  let progress = isWorkDay
    ? workProgress({ creditedStart, expectedCheckout, nowMin: progressNow })
    : { ratio: 0, remaining: null, overtime: 0, started: false, complete: false, span: 0 };
  const weekendWork = isWorkDay && (input.weekendDays || [5, 6]).includes(weekdayIndex(input.date));
  let overtime = progress.overtime || 0;
  if (weekendWork && checkInMin != null && evalMin != null) {
    overtime = Math.max(0, evalMin - checkInMin);
    progress = { ...progress, overtime };
  }

  const morningCredit =
    permission.active && permission.type === "morning" && checkInMin != null && creditedStart != null
      ? Math.max(0, checkInMin - creditedStart)
      : 0;

  return {
    status,
    isWorkDay,
    officialMin,
    graceEndMin,
    checkInMin,
    checkOutMin,
    nowMin,
    targetMinutes,
    permission,
    permissionMinutes,
    creditedStart,
    expectedCheckout,
    checkoutClamped,
    morningCredit,
    eveningReduction: permission.active && permission.type === "evening" ? permissionMinutes : 0,
    middayExcused: permission.active && permission.type === "midday" ? permissionMinutes : 0,
    breaks: { dhuhr, asr, manual, consumed: breakConsumed, allowance: breakAllowance },
    breakConsumed,
    breakAllowance,
    gross,
    net,
    progress,
    weekendWork,
    overtime,
    evalMin,
  };
}

export function effectiveOvertimeMinutes(row, weekendDays = [5, 6]) {
  const stored = Math.max(0, Math.round(Number(row?.overtimeMinutes) || 0));
  if (!row || row.status !== STATUS.work) return stored;
  if (!(weekendDays || [5, 6]).includes(weekdayIndex(row.date))) return stored;
  const start = minutesFromHHMM(String(row.checkIn || ""));
  const end = minutesFromHHMM(String(row.checkOut || ""));
  if (start == null || end == null || end < start) return stored;
  return Math.max(stored, end - start);
}

export function toLogEntry(day, computed) {
  const checkout =
    day.checkOut ||
    (computed.isWorkDay && computed.expectedCheckout != null ? hhmmFromMinutes(computed.expectedCheckout) : "");
  return {
    date: day.date,
    status: computed.status,
    checkIn: day.checkIn || "",
    checkOut: checkout || "",
    breakMinutes: Math.round(computed.breakConsumed || 0),
    permissionMinutes: computed.permission?.active ? Math.round(computed.permission.minutes || 0) : 0,
    permissionType: computed.permission?.active ? computed.permission.type : "",
    grossMinutes: Math.round(computed.gross || 0),
    netMinutes: Math.round(computed.net || 0),
    targetMinutes: computed.targetMinutes,
    overtimeMinutes: Math.round(computed.overtime || 0),
  };
}

export function upsertLog(log, entry) {
  const next = Array.isArray(log) ? log.slice() : [];
  const index = next.findIndex((row) => row && row.date === entry.date);
  if (index === -1) next.push(entry);
  else next[index] = entry;
  next.sort((a, b) => String(a.date).localeCompare(String(b.date)));
  return next;
}

export function chooseArchiveEval(state) {
  const day = state.today;
  if (day.checkOut) {
    const mins = minutesFromHHMM(day.checkOut);
    if (mins != null) return mins;
  }
  if (day.lastSeenAt) {
    const seen = new Date(day.lastSeenAt);
    if (!Number.isNaN(seen.getTime()) && isoDateLocal(seen) === day.date) {
      return seen.getHours() * 60 + seen.getMinutes() + seen.getSeconds() / 60;
    }
  }
  const preview = computeDay(inputFromState(state, 12 * 60));
  if (preview.expectedCheckout != null) return preview.expectedCheckout;
  return 23 * 60 + 59;
}

export function settleToday(state, todayISO) {
  if (!state?.today?.date || state.today.date === todayISO) return { state, archived: null };
  if (state.today.date < todayISO) return archiveAndStartDay(state, todayISO, chooseArchiveEval(state));
  return {
    archived: null,
    state: {
      ...state,
      today: createToday(todayISO, { permissionMinutes: state.settings.defaultPermissionMinutes }),
      notifiedKeys: [],
    },
  };
}

export function archiveAndStartDay(state, todayISO, evalMin) {
  if (!state?.today?.date || state.today.date === todayISO) return { state, archived: null };
  if (state.today.date > todayISO) return { state, archived: null };
  const computed = computeDay(inputFromState(state, evalMin));
  const entry = toLogEntry(state.today, computed);
  return {
    archived: entry,
    state: {
      ...state,
      log: upsertLog(state.log, entry),
      today: createToday(todayISO, { permissionMinutes: state.settings.defaultPermissionMinutes }),
      notifiedKeys: [],
    },
  };
}

export function dueAlerts({ nowMin, model, notifications, firedKeys }) {
  const fired = new Set(firedKeys || []);
  const alerts = [];
  const now = nowMin;
  const push = (alert) => {
    if (fired.has(alert.id) || alerts.some((item) => item.id === alert.id)) return;
    alerts.push(alert);
  };
  if (!model?.isWorkDay) return alerts;

  const names = { dhuhr: "الظهر", asr: "العصر", manual: "اليدوي" };
  const breaks = [model.breaks.dhuhr, model.breaks.asr, model.breaks.manual];

  if (notifications?.preBreak?.enabled) {
    const lead = clampInt(notifications.preBreak.minutesBefore ?? 10, 0, 180);
    for (const item of breaks) {
      if (!item.enabled || item.startMin == null) continue;
      const end = item.startMin + item.durationMin;
      const running = now >= item.startMin && now < end && item.reason !== "late" && item.reason !== "left-before" && item.reason !== "shift-ends-before";
      if (running) {
        push({
          id: `start-${item.id}-${item.startMin}`,
          kind: "break-start",
          title: "بدأ البريك",
          body: `بريك ${names[item.id] || ""} شغّال. المتبقي ${formatDuration(Math.max(0, end - now))}.`.trim(),
        });
      }
      if (!item.applied) continue;
      const event = item.startMin - lead;
      const latest = lead === 0 ? item.startMin + 1 : item.startMin;
      if (now >= event && now < latest) {
        push({
          id: `pre-${item.id}`,
          kind: "pre-break",
          title: "اقترب البريك",
          body: lead === 0
            ? `بدأ بريك ${names[item.id] || ""}`.trim()
            : `بريك ${names[item.id] || ""} يبدأ خلال ${formatDuration(lead)}`.trim(),
        });
      }
    }
  }

  if (notifications?.breakEnd?.enabled) {
    for (const item of breaks) {
      if (!item.applied || item.startMin == null) continue;
      const event = item.startMin + item.durationMin;
      if (now >= event && now < event + 5) {
        push({
          id: `end-${item.id}`,
          kind: "break-end",
          title: "انتهى البريك",
          body: `انتهى بريك ${names[item.id] || ""}، حان الرجوع للعمل`.trim(),
        });
      }
    }
  }

  if (notifications?.shiftEnd?.enabled && model.expectedCheckout != null && model.checkOutMin == null) {
    const lead = clampInt(notifications.shiftEnd.minutesBefore ?? 0, 0, 180);
    const event = model.expectedCheckout - lead;
    const latest = lead === 0 ? model.expectedCheckout + 10 : model.expectedCheckout;
    if (now >= event && now < latest) {
      push({
        id: "shift-end",
        kind: "shift-end",
        title: "موعد الانصراف",
        body: "اكتملت ساعات الدوام. المفروض تمشي الآن.",
      });
    }
  }

  if (notifications?.permission?.enabled && model.permission?.timeMin != null) {
    const event = model.permission.timeMin;
    if (now >= event && now < event + 10) {
      push({
        id: "permission-time",
        kind: "permission",
        title: "موعد الإذن",
        body: "حان الوقت الذي حددته للإذن.",
      });
    }
  }

  const priority = { "shift-end": 0, "break-end": 1, permission: 2, "pre-break": 3 };
  alerts.sort((a, b) => (priority[a.kind] ?? 9) - (priority[b.kind] ?? 9));
  return alerts;
}

export function monthKey(isoDate) {
  return String(isoDate).slice(0, 7);
}

export function filterLogByMonth(log, key) {
  return (log || []).filter((row) => row && String(row.date).startsWith(key));
}

export function sumNetMinutes(rows) {
  return (rows || []).reduce((sum, row) => {
    if (!row || row.status !== STATUS.work) return sum;
    return sum + (Number(row.netMinutes) || 0);
  }, 0);
}

function validTimeString(value) {
  return typeof value === "string" && minutesFromHHMM(value) != null && /^(\d{2}):(\d{2})$/.test(value);
}

function acceptableTime(value) {
  if (value == null || value === "") return true;
  return minutesFromHHMM(String(value).trim()) != null;
}

export function validateImportPayload(payload) {
  const fail = () => ({ ok: false, error: "فشل الاستيراد" });
  if (!payload || typeof payload !== "object") return fail();
  if (payload.signature !== APP_SIGNATURE) return fail();
  if (!payload.settings || typeof payload.settings !== "object" || Array.isArray(payload.settings)) return fail();
  if (!payload.today || typeof payload.today !== "object" || Array.isArray(payload.today)) return fail();
  if (!Array.isArray(payload.log)) return fail();
  if (!validTimeString(payload.settings.officialCheckIn) || !validTimeString(payload.settings.graceEnd)) return fail();
  if (!Number.isFinite(Number(payload.settings.dailyTargetMinutes))) return fail();
  if (!payload.settings.dhuhr || !payload.settings.asr) return fail();
  if (!validTimeString(payload.settings.dhuhr.start) || !validTimeString(payload.settings.asr.start)) return fail();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(payload.today.date || ""))) return fail();
  for (const key of ["checkIn", "checkOut", "permissionTime", "manualBreakStart"]) {
    if (!acceptableTime(payload.today[key])) return fail();
  }

  for (const row of payload.log) {
    if (!row || typeof row !== "object") return fail();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(row.date || ""))) return fail();
    if (row.status && !STATUS_LABELS[row.status] && !LABEL_TO_STATUS[row.status]) return fail();
    if (!acceptableTime(row.checkIn) || !acceptableTime(row.checkOut)) return fail();
    for (const key of ["breakMinutes", "permissionMinutes", "grossMinutes", "netMinutes", "targetMinutes", "overtimeMinutes"]) {
      if (row[key] != null && row[key] !== "" && !Number.isFinite(Number(row[key]))) return fail();
    }
  }
  return { ok: true, error: null };
}

export function applyImport(state, payload) {
  const check = validateImportPayload(payload);
  if (!check.ok) return { ok: false, error: "فشل الاستيراد", state };
  const next = sanitizeImported(payload);
  return { ok: true, error: null, state: next };
}

function sanitizeBreak(raw, fallback) {
  const source = raw && typeof raw === "object" ? raw : {};
  const start = validTimeString(source.start) ? source.start : minutesFromHHMM(source.start) != null ? hhmmFromMinutes(minutesFromHHMM(source.start)) : fallback.start;
  return {
    enabled: typeof source.enabled === "boolean" ? source.enabled : fallback.enabled,
    start: validTimeString(start) ? start : fallback.start,
    duration: clampInt(source.duration ?? fallback.duration, 0, 600),
  };
}

const ACCENT_ALIAS = { blue: "ocean", green: "emerald", brass: "amber", teal: "emerald", rose: "violet" };
const ACCENT_IDS = ["emerald", "ocean", "amber", "violet"];

function normalizeAccent(value) {
  const id = ACCENT_ALIAS[value] || value;
  return ACCENT_IDS.includes(id) ? id : "emerald";
}

function sanitizeSettings(raw, fallback) {
  const source = raw && typeof raw === "object" ? raw : {};
  const notifications = source.notifications && typeof source.notifications === "object" ? source.notifications : {};
  const weekend = Array.isArray(source.weekendDays) && source.weekendDays.length
    ? source.weekendDays.map(Number).filter((n) => n >= 0 && n <= 6)
    : fallback.weekendDays;
  return {
    officialCheckIn: validTimeString(source.officialCheckIn) ? source.officialCheckIn : fallback.officialCheckIn,
    graceEnd: validTimeString(source.graceEnd) ? source.graceEnd : fallback.graceEnd,
    dailyTargetMinutes: clampInt(source.dailyTargetMinutes ?? fallback.dailyTargetMinutes, 30, 16 * 60),
    defaultPermissionMinutes: clampInt(source.defaultPermissionMinutes ?? fallback.defaultPermissionMinutes, 0, 12 * 60),
    theme: source.theme === "dark" ? "dark" : "light",
    accent: normalizeAccent(source.accent),
    lang: source.lang === "en" ? "en" : "ar",
    timePicker: source.timePicker === "scroll" ? "scroll" : "clock",
    progressShimmer: source.progressShimmer === false ? false : true,
    showHeaderTarget: source.showHeaderTarget === false ? false : true,
    soundEnabled: typeof source.soundEnabled === "boolean" ? source.soundEnabled : fallback.soundEnabled,
    tabBlinkEnabled: typeof source.tabBlinkEnabled === "boolean" ? source.tabBlinkEnabled : fallback.tabBlinkEnabled,
    weekendDays: weekend.length ? weekend : fallback.weekendDays,
    dhuhr: sanitizeBreak(source.dhuhr, fallback.dhuhr),
    asr: sanitizeBreak(source.asr, fallback.asr),
    notifications: {
      preBreak: {
        enabled: typeof notifications.preBreak?.enabled === "boolean" ? notifications.preBreak.enabled : fallback.notifications.preBreak.enabled,
        minutesBefore: clampInt(notifications.preBreak?.minutesBefore ?? fallback.notifications.preBreak.minutesBefore, 0, 180),
        blink: typeof notifications.preBreak?.blink === "boolean" ? notifications.preBreak.blink : fallback.notifications.preBreak.blink,
      },
      breakEnd: {
        enabled: typeof notifications.breakEnd?.enabled === "boolean" ? notifications.breakEnd.enabled : fallback.notifications.breakEnd.enabled,
        blink: typeof notifications.breakEnd?.blink === "boolean" ? notifications.breakEnd.blink : fallback.notifications.breakEnd.blink,
      },
      shiftEnd: {
        enabled: typeof notifications.shiftEnd?.enabled === "boolean" ? notifications.shiftEnd.enabled : fallback.notifications.shiftEnd.enabled,
        minutesBefore: clampInt(notifications.shiftEnd?.minutesBefore ?? fallback.notifications.shiftEnd.minutesBefore ?? 0, 0, 180),
        blink: typeof notifications.shiftEnd?.blink === "boolean" ? notifications.shiftEnd.blink : fallback.notifications.shiftEnd.blink,
      },
      permission: {
        enabled: typeof notifications.permission?.enabled === "boolean" ? notifications.permission.enabled : fallback.notifications.permission.enabled,
        blink: typeof notifications.permission?.blink === "boolean" ? notifications.permission.blink : fallback.notifications.permission.blink,
      },
    },
    sheetsWebhook: typeof source.sheetsWebhook === "string" ? source.sheetsWebhook.trim() : "",
    sheetsConnected: source.sheetsConnected === true,
    sheetsImport: source.sheetsImport === "keep" || source.sheetsImport === "replace" ? source.sheetsImport : "",
    fileSnapshot: source.fileSnapshot === true,
  };
}

function sanitizeToday(raw, settings, fallbackDate) {
  const source = raw && typeof raw === "object" ? raw : {};
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(source.date || "")) ? source.date : fallbackDate;
  const override = [STATUS.work, STATUS.weekend, STATUS.official, STATUS.personal].includes(source.statusOverride)
    ? source.statusOverride
    : null;
  const type = PERMISSION_TYPES.has(source.permissionType) ? source.permissionType : "auto";
  return {
    date,
    statusOverride: override,
    workplace: source.workplace === "remote" ? "remote" : "office",
    checkIn: validTimeString(source.checkIn) ? source.checkIn : null,
    checkOut: validTimeString(source.checkOut) ? source.checkOut : null,
    targetMinutes: source.targetMinutes == null || source.targetMinutes === "" ? null : clampInt(source.targetMinutes, 30, 16 * 60),
    permissionOn: source.permissionOn === true,
    permissionTouched: source.permissionTouched === true,
    permissionMinutes: clampInt(source.permissionMinutes ?? settings.defaultPermissionMinutes, 0, 12 * 60),
    permissionType: type,
    permissionTime: validTimeString(source.permissionTime) ? source.permissionTime : null,
    manualBreakMinutes: clampInt(source.manualBreakMinutes ?? 0, 0, 600),
    manualBreakStart: validTimeString(source.manualBreakStart) ? source.manualBreakStart : null,
    dhuhrEnabled: typeof source.dhuhrEnabled === "boolean" ? source.dhuhrEnabled : null,
    asrEnabled: typeof source.asrEnabled === "boolean" ? source.asrEnabled : null,
    lastSeenAt: typeof source.lastSeenAt === "string" ? source.lastSeenAt : null,
  };
}

function normalizeStatus(value) {
  if (STATUS_LABELS[value]) return value;
  if (LABEL_TO_STATUS[value]) return LABEL_TO_STATUS[value];
  return "";
}

export function sanitizeLog(raw) {
  if (!Array.isArray(raw)) return [];
  const rows = [];
  for (const row of raw) {
    if (!row || typeof row !== "object") continue;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(row.date || ""))) continue;
    const status = normalizeStatus(row.status) || STATUS.work;
    rows.push({
      date: row.date,
      status,
      checkIn: validTimeString(row.checkIn) ? row.checkIn : row.checkIn ? hhmmFromMinutes(minutesFromHHMM(String(row.checkIn))) : "",
      checkOut: validTimeString(row.checkOut) ? row.checkOut : row.checkOut ? hhmmFromMinutes(minutesFromHHMM(String(row.checkOut))) : "",
      breakMinutes: clampInt(row.breakMinutes ?? 0, 0, 24 * 60),
      permissionMinutes: clampInt(row.permissionMinutes ?? 0, 0, 24 * 60),
      permissionType: ["morning", "evening", "midday"].includes(row.permissionType) ? row.permissionType : "",
      grossMinutes: clampInt(row.grossMinutes ?? 0, 0, 24 * 60),
      netMinutes: clampInt(row.netMinutes ?? 0, 0, 24 * 60),
      targetMinutes: clampInt(row.targetMinutes ?? 480, 0, 24 * 60),
      overtimeMinutes: clampInt(row.overtimeMinutes ?? 0, 0, 24 * 60),
    });
  }
  rows.sort((a, b) => a.date.localeCompare(b.date));
  return rows;
}

function sheetFlag(value) {
  return value === true || value === 1 || value === "1";
}

export function todayFromSheetRecord(row, settings) {
  const entry = sanitizeLog([row])[0];
  const fallback = createToday(entry?.date || isoDateLocal(new Date()), {
    permissionMinutes: settings?.defaultPermissionMinutes,
  });
  if (!entry) return fallback;
  const active = entry.permissionMinutes > 0 && !!entry.permissionType;
  return {
    ...fallback,
    date: entry.date,
    statusOverride: entry.status,
    workplace: "office",
    checkIn: entry.checkIn || null,
    checkOut: sheetFlag(row?.checkedOut) && entry.checkOut ? entry.checkOut : null,
    targetMinutes: entry.targetMinutes > 0 ? entry.targetMinutes : null,
    permissionOn: active,
    permissionTouched: true,
    permissionMinutes: active ? entry.permissionMinutes : fallback.permissionMinutes,
    permissionType: active ? entry.permissionType : "auto",
  };
}

export function applySheetRecords(state, records, todayISO) {
  const incoming = Array.isArray(records) ? records : [];
  const byDate = new Map();
  for (const row of incoming) {
    const entry = sanitizeLog([row])[0];
    if (!entry || byDate.has(entry.date)) continue;
    byDate.set(entry.date, { ...entry, checkedOut: sheetFlag(row?.checkedOut) });
  }
  if (!byDate.size) return { ok: true, imported: 0, state };
  const todayRow = byDate.get(todayISO);
  const log = [...byDate.values()]
    .filter((row) => row.date !== todayISO)
    .map(({ checkedOut, ...entry }) => entry)
    .sort((a, b) => a.date.localeCompare(b.date));
  const today = todayRow
    ? todayFromSheetRecord(todayRow, state.settings)
    : state.today?.date === todayISO
      ? state.today
      : createToday(todayISO, { permissionMinutes: state.settings?.defaultPermissionMinutes });
  return {
    ok: true,
    imported: byDate.size,
    state: { ...state, today, log },
  };
}

export function sanitizeStored(raw, fallbackDate) {
  const defaults = defaultState(fallbackDate);
  if (!raw || typeof raw !== "object") return defaults;
  const settings = sanitizeSettings(raw.settings, defaults.settings);
  return {
    version: 1,
    settings,
    today: sanitizeToday(raw.today, settings, fallbackDate),
    log: sanitizeLog(raw.log),
    dismissedDates: Array.isArray(raw.dismissedDates) ? raw.dismissedDates.filter((date) => /^\d{4}-\d{2}-\d{2}$/.test(String(date))) : [],
    notifiedKeys: Array.isArray(raw.notifiedKeys) ? raw.notifiedKeys.filter((key) => typeof key === "string") : [],
    lastSyncAt: typeof raw.lastSyncAt === "string" ? raw.lastSyncAt : null,
  };
}

function sanitizeImported(payload) {
  const base = sanitizeStored(
    {
      version: 1,
      settings: payload.settings,
      today: payload.today,
      log: payload.log,
      notifiedKeys: [],
      lastSyncAt: null,
    },
    payload.today.date,
  );
  base.settings.sheetsConnected = payload.settings.sheetsConnected === true && !!base.settings.sheetsWebhook;
  return base;
}

export function blankWorkInput(overrides = {}) {
  return {
    date: "2026-11-25",
    statusOverride: null,
    weekendDays: [5, 6],
    checkIn: "08:30",
    checkOut: null,
    nowMin: 16 * 60 + 30,
    officialCheckIn: "08:30",
    graceEnd: "09:30",
    targetMinutes: 480,
    permissionOn: false,
    permissionTouched: false,
    permissionMinutes: 120,
    permissionType: "auto",
    permissionTime: null,
    dhuhrEnabled: true,
    dhuhrStart: "11:50",
    dhuhrDuration: 45,
    asrEnabled: true,
    asrStart: "14:40",
    asrDuration: 15,
    manualMinutes: 0,
    manualStart: null,
    ...overrides,
  };
}
