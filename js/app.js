import {
  APP_SIGNATURE,
  STATUS,
  STATUS_LABELS,
  STATUS_LABELS_EN,
  applyImport,
  applySheetRecords,
  clampInt,
  computeDay,
  createToday,
  dueAlerts,
  fillMissingDays,
  effectiveOvertimeMinutes,
  filterLogByMonth,
  hhmmFromMinutes,
  inputFromState,
  isoDateLocal,
  leaveEntry,
  minutesFromHHMM,
  monthKey,
  sanitizeStored,
  settleToday,
  sumNetMinutes,
  timeParts,
  toLogEntry,
  upsertLog,
} from "./calc.js";
import { downloadLeaveMarkdown, downloadLeaveWorkbook, downloadMarkdown, downloadWorkbook, overlaySettings, readBackupFile, readLeaveFile, settingsPairs } from "./excel.js";
import { APPS_SCRIPT, linkIsBehindScript, scriptChangedAt, stampWebhook } from "./guide.js";
import { formatClockTime, formatDateLine, formatDayCount, formatHourMinute, formatLength, formatMinuteStamp, formatMonthTitle, translate, weekdayLabel } from "./i18n.js";
import {
  addLeaveEntries,
  buildTimeline,
  coverLeaveHorizon,
  entriesBeforeStart,
  expandLeaveHorizon,
  formatLeaveAmount,
  leaveDaysByMonthInYear,
  leaveDaysInMonth,
  removeLeaveEntry,
  resetLeaveSettings,
  round4,
  sanitizeLeave,
  setLeaveMonthMode,
  setLeaveMonthUsed,
  summarizeLeave,
  updateLeaveSettings,
  leaveSyncPayload,
} from "./leave.js";
import { playChime, setBaseTitle, startTabBlink, stopTabBlink, warmAudio, watchFocusStop } from "./notify.js";
import { pingWebhook, pullWebhook, pushWebhook } from "./sheets.js";
import { loadRaw, saveState } from "./storage.js";
import { APP_VERSION } from "../app.config.js";

const ACCENTS = ["emerald", "ocean", "amber", "violet"];
const ACCENT_ALIAS = { blue: "ocean", green: "emerald", brass: "amber", teal: "emerald", rose: "violet" };
const UI_TO_TYPE = { auto: "auto", morning: "morning", evening: "evening", half_day: "midday" };
const TYPE_TO_UI = { auto: "auto", morning: "morning", evening: "evening", midday: "half_day" };
const ITEM_HEIGHT = 42;
const MODE_TO_STATUS = {
  office: STATUS.work,
  remote: STATUS.work,
  weekend: STATUS.weekend,
  personal_leave: STATUS.personal,
  official_holiday: STATUS.official,
};

const ui = { lastMinute: null, syncTimer: null, syncing: false, pullReady: false, pulling: false, webhookEditing: false, pendingWebhook: "", logSig: "", logMonth: "", monthSig: "", picker: null, feedbackTimer: null, view: "hours", leaveYear: "", leaveStatus: "all", leaveSpan: "single", leaveSig: "", leaveYearSig: "", leaveHoldSettings: false, leaveEditing: false };
const stored = loadRaw();
const storedForeign = Boolean(stored && stored.signature && stored.signature !== APP_SIGNATURE);
let state = sanitizeStored(storedForeign ? null : stored, isoDateLocal(new Date()));

const lang = () => (state.settings.lang === "en" ? "en" : "ar");
const t = (key, vars) => translate(lang(), key, vars);
const len = (mins) => formatLength(mins, lang());
const clockText = (mins) => (mins == null ? "--:--" : formatClockTime(mins, lang()));
const hhmmNow = (date) => hhmmFromMinutes(date.getHours() * 60 + date.getMinutes());
const nowFloat = (date) => date.getHours() * 60 + date.getMinutes() + date.getSeconds() / 60;
const modelNow = (date = new Date()) => computeDay(inputFromState(state, nowFloat(date)));

function normalizeAccent(value) {
  const id = ACCENT_ALIAS[value] || value;
  return ACCENTS.includes(id) ? id : "emerald";
}

function $(id) { return document.getElementById(id); }

function persist({ sync = false } = {}) {
  try { saveState({ ...state, signature: APP_SIGNATURE }); } catch { /* storage may be full */ }
  if (sync) scheduleSync();
}

function hideFeedback() {
  clearTimeout(ui.feedbackTimer);
  const box = $("ioFeedbackBox");
  if (box) box.className = "feedback-banner hidden";
}

function feedbackKind(tone) {
  if (tone === true || tone === "error") return "error";
  if (tone === "warn") return "warn";
  return "success";
}

function showFeedback(message, tone = false) {
  const box = $("ioFeedbackBox");
  if (!box) return;
  const kind = feedbackKind(tone);
  const text = $("ioFeedbackText");
  const bar = box.querySelector(".feedback-timer");
  const already = text
    && text.textContent === message
    && box.classList.contains(kind)
    && !box.classList.contains("hidden");
  if (already) return;
  if (text) text.textContent = message;
  box.className = `feedback-banner ${kind}`;
  const seconds = Math.min(60, Math.max(1, Math.round(Number(state.settings.alertSeconds) || 5)));
  if (bar) {
    bar.style.animation = "none";
    void bar.offsetWidth;
    bar.style.animation = `feedbackLife ${seconds}s linear forwards`;
  }
  clearTimeout(ui.feedbackTimer);
  ui.feedbackTimer = setTimeout(hideFeedback, seconds * 1000);
}

function applyTheme() {
  state.settings.accent = normalizeAccent(state.settings.accent);
  const dark = state.settings.theme === "dark";
  document.documentElement.dataset.theme = dark ? "dark" : "light";
  document.documentElement.dataset.accent = state.settings.accent;
  const meta = document.querySelector('meta[name="theme-color"]');
  if (meta) meta.content = dark ? "#0b1120" : "#f1f5f9";
  document.querySelectorAll("[data-accent-btn]").forEach((btn) => {
    const on = btn.dataset.accentBtn === state.settings.accent;
    btn.classList.toggle("active", on);
    btn.setAttribute("aria-pressed", on ? "true" : "false");
  });
  const themeLabel = dark ? t("ui.themeDay") : t("ui.themeNight");
  const soundLabel = state.settings.soundEnabled ? t("ui.soundOn") : t("ui.soundOff");
  const langLabel = t("ui.lang");
  document.querySelectorAll(".theme-toggle").forEach((btn) => { btn.textContent = themeLabel; });
  document.querySelectorAll(".sound-toggle").forEach((btn) => { btn.textContent = soundLabel; });
  document.querySelectorAll(".lang-toggle").forEach((btn) => { btn.textContent = langLabel; });
}

function applyLang() {
  const english = lang() === "en";
  document.documentElement.lang = english ? "en" : "ar";
  document.documentElement.dir = english ? "ltr" : "rtl";
  setBaseTitle(t("tab.title"));
  document.querySelectorAll("[data-i18n]").forEach((el) => { el.textContent = t(el.dataset.i18n); });
  document.querySelectorAll("[data-i18n-aria]").forEach((el) => el.setAttribute("aria-label", t(el.dataset.i18nAria)));
  document.querySelectorAll("[data-i18n-title]").forEach((el) => { el.title = t(el.dataset.i18nTitle); });
  const version = $("appVersion");
  if (version) version.textContent = t("ui.version").replace("{version}", APP_VERSION);
  applyTheme();
}

function setChecked(id, value) {
  const el = $(id);
  if (!el || document.activeElement === el) return;
  el.checked = !!value;
}

function setNumber(id, value) {
  const el = $(id);
  if (!el || document.activeElement === el) return;
  const next = String(value);
  if (el.value !== next) el.value = next;
}

function clockParts(totalMinutes) {
  const parts = timeParts(totalMinutes);
  const [hourText, minuteText] = parts.display.split(":");
  return {
    hour12: Number(hourText) || 12,
    minute: Number(minuteText) || 0,
    pm: parts.period === "م",
    display: parts.display,
    period: parts.period,
  };
}

function liveClock(date) {
  const parts = clockParts(date.getHours() * 60 + date.getMinutes());
  const second = String(date.getSeconds()).padStart(2, "0");
  const period = lang() === "en" ? (parts.pm ? "PM" : "AM") : parts.period;
  return `${parts.display}:${second} ${period}`;
}

function modeKey(model) {
  if (model.status === STATUS.weekend) return "weekend";
  if (model.status === STATUS.official) return "official_holiday";
  if (model.status === STATUS.personal) return "personal_leave";
  return state.today.workplace === "remote" ? "remote" : "office";
}

function statusText(status) {
  const labels = lang() === "en" ? STATUS_LABELS_EN : STATUS_LABELS;
  return labels[status] || status;
}

function permLabel(model) {
  if (!model.permission?.active) return lang() === "en" ? "None" : "بدون إذن";
  const key = { morning: "ui.typeMorning", evening: "ui.typeEvening", midday: "ui.typeMid" }[model.permission.type];
  return key ? t(key) : t("ui.typeAuto");
}

function activeBreak(model) {
  return [model.breaks.dhuhr, model.breaks.asr, model.breaks.manual].find((item) => item.phase === "active") || null;
}

function liveBreak(model, date) {
  const items = [model.breaks.dhuhr, model.breaks.asr, model.breaks.manual];
  const nowSec = date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds();
  const inside = (item) => item.enabled && item.startMin != null && item.durationMin > 0
    && nowSec >= item.startMin * 60
    && nowSec < (item.startMin + item.durationMin) * 60
    && item.reason !== "late"
    && item.reason !== "left-before"
    && item.reason !== "shift-ends-before";
  const running = items.find((item) => item.phase === "active") || items.find(inside);
  if (running) {
    return {
      item: running,
      until: running.startMin + running.durationMin,
      mode: running.reason === "no-checkin" ? "need-checkin" : "active",
    };
  }
  const upcoming = items.find((item) => item.phase === "upcoming");
  if (!upcoming) return null;
  return { item: upcoming, until: upcoming.startMin, mode: "upcoming" };
}

function breakName(id) {
  if (id === "dhuhr") return lang() === "en" ? "Dhuhr" : "الظهر";
  if (id === "asr") return lang() === "en" ? "Asr" : "العصر";
  return lang() === "en" ? "Manual" : "اليدوي";
}

function paint(date = new Date()) {
  const model = modelNow(date);
  const target = state.today.targetMinutes ?? state.settings.dailyTargetMinutes;
  const onLeave = ui.view === "leave";
  $("pageTitle").textContent = onLeave ? t("leave.title") : t("brand.title");
  $("dateSubtitle").textContent = onLeave
    ? t("leave.subtitle")
    : `${formatDateLine(state.today.date, lang())} • ${statusText(model.status)}`;
  $("liveCurrentClock").textContent = liveClock(date);
  $("headerTargetHours").textContent = `${(target / 60).toFixed(1)} ${lang() === "en" ? "h" : "س"}`;
  if (document.activeElement !== $("dayStatusSelect")) $("dayStatusSelect").value = modeKey(model);

  $("checkinDisplay").textContent = state.today.checkIn ? clockText(model.checkInMin) : "--:--";
  $("checkinStatusNote").textContent = !state.today.checkIn
    ? (lang() === "en" ? "Not checked in yet" : "لم يُسجَّل الحضور بعد")
    : model.permission?.active && model.permission.type === "morning"
      ? `${t("ui.typeMorning")} • ${clockText(model.creditedStart)}`
      : `${clockText(model.officialMin)} – ${clockText(model.graceEndMin)}`;

  $("expectedCheckoutDisplay").textContent = model.isWorkDay && state.today.checkIn ? clockText(model.expectedCheckout) : (model.isWorkDay ? "--:--" : (lang() === "en" ? "Off" : "إجازة"));
  const remain = model.progress?.remaining;
  const weekendShift = model.weekendWork && state.today.checkIn;
  $("checkoutNote").textContent = weekendShift
    ? `${lang() === "en" ? "Overtime" : "إضافي"} ${len(model.overtime)}`
    : !model.isWorkDay
      ? statusText(model.status)
      : !state.today.checkIn
        ? (lang() === "en" ? "Check in to calculate your leave time" : "سجّل الحضور ليظهر موعد الانصراف")
        : model.progress?.complete
          ? (model.progress.overtime > 0 ? `${lang() === "en" ? "Overtime" : "إضافي"} ${len(model.progress.overtime)}` : (lang() === "en" ? "Shift complete" : "اكتمل الدوام"))
          : `${lang() === "en" ? "Left" : "متبقي"} ${remain == null ? "—" : len(remain)}`;

  $("checkoutDisplay").textContent = state.today.checkOut ? clockText(model.checkOutMin) : "--:--";
  $("checkoutStatusNote").textContent = state.today.checkOut
    ? (lang() === "en" ? "Checked out" : "تم تسجيل الانصراف")
    : state.today.checkIn && model.expectedCheckout != null
      ? t("ui.checkoutDefault")
        .replace("{time}", clockText(model.expectedCheckout))
        .replace("{hours}", String(Math.round((target / 60) * 10) / 10))
      : (lang() === "en" ? "Tap to record overtime" : "اضغط لو هتزود أوفر تايم");
  $("clearCheckoutBtn").hidden = !state.today.checkOut;
  const checkedIn = Boolean(state.today.checkIn);
  $("checkInNowBtn").hidden = checkedIn;
  $("checkOutNowBtn").hidden = !checkedIn;
  $("checkOutNowBtn").classList.toggle("btn-primary", checkedIn);
  $("checkOutNowBtn").classList.toggle("btn-secondary", !checkedIn);

  const ratio = model.progress?.ratio || 0;
  const percent = Math.min(100, Math.max(0, Math.round(ratio * 100)));
  const fill = $("workProgressFill");
  fill.style.width = `${percent}%`;
  fill.classList.toggle("completed", percent >= 100);
  $("workProgressTrack").setAttribute("aria-valuenow", String(percent));
  $("workProgressPercent").textContent = `${percent}%`;
  $("workElapsedText").textContent = `${lang() === "en" ? "Elapsed" : "المنقضي"}: ${len(model.gross)}`;
  $("workRemainingText").textContent = weekendShift
    ? `${lang() === "en" ? "Overtime" : "إضافي"}: ${len(model.overtime)}`
    : model.isWorkDay
      ? (model.progress?.complete
        ? `${lang() === "en" ? "Overtime" : "إضافي"}: ${len(model.progress.overtime)}`
        : `${lang() === "en" ? "Until leave" : "المتبقي للانصراف"}: ${remain == null ? "—" : len(remain)}`)
      : (lang() === "en" ? "No shift today" : "لا يوجد دوام اليوم");

  $("workedWithBreakVal").textContent = len(model.gross);
  $("workedWithoutBreakVal").textContent = len(model.net);
  $("totalBreakVal").textContent = len(model.breakConsumed);
  $("permissionSummaryVal").textContent = model.permission?.active ? `${len(model.permission.minutes)} (${permLabel(model)})` : len(0);
  $("detectedPermBadge").textContent = `${lang() === "en" ? "Type" : "النوع"}: ${permLabel(model)}`;

  setChecked("dhuhrBreakEnabled", state.today.dhuhrEnabled ?? state.settings.dhuhr.enabled);
  setChecked("asrBreakEnabled", state.today.asrEnabled ?? state.settings.asr.enabled);
  setChecked("dhuhrAutoEnable", state.settings.dhuhr.enabled);
  setChecked("asrAutoEnable", state.settings.asr.enabled);
  $("dhuhrStartText").textContent = clockText(minutesFromHHMM(state.settings.dhuhr.start));
  $("asrStartText").textContent = clockText(minutesFromHHMM(state.settings.asr.start));
  setNumber("dhuhrDurationInput", state.settings.dhuhr.duration);
  setNumber("asrDurationInput", state.settings.asr.duration);
  const minuteWord = t("ui.minuteShort");
  $("dhuhrBreakLabel").textContent = `${t("ui.dhuhr")} (${state.settings.dhuhr.duration} ${minuteWord})`;
  $("asrBreakLabel").textContent = `${t("ui.asr")} (${state.settings.asr.duration} ${minuteWord})`;
  setNumber("manualBreakInput", state.today.manualBreakMinutes || 0);
  $("breakMinutesSumBadge").textContent = len(model.breakAllowance);

  const live = liveBreak(model, date);
  const badge = $("activeBreakBadge");
  if (live) {
    const left = secondsUntil(live.until, date);
    const span = Math.max(1, live.item.durationMin * 60);
    const gone = Math.min(span, Math.max(0, span - left));
    badge.textContent = `${breakName(live.item.id)} ${live.mode === "upcoming" ? (lang() === "en" ? "soon" : "قريب") : (lang() === "en" ? "now" : "جارٍ الآن")}`;
    badge.className = live.mode === "upcoming" ? "status-pill" : "status-pill active-break";
    const action = live.mode === "upcoming"
      ? (lang() === "en" ? "starts in" : "يبدأ خلال")
      : live.mode === "need-checkin"
        ? (lang() === "en" ? "left • check in to count it" : "متبقي • سجّل الحضور ليتحسب")
        : (lang() === "en" ? "left" : "متبقي");
    $("breakCountdownLabel").textContent = `${breakName(live.item.id)} • ${action}`;
    $("breakRemainingText").textContent = formatCountdown(left);
    $("breakProgressFill").style.width = live.mode === "upcoming" ? "0%" : `${Math.round((gone / span) * 100)}%`;
  } else {
    badge.textContent = t("ui.noBreak");
    badge.className = "status-pill";
    $("breakCountdownLabel").textContent = t("ui.breakLive");
    $("breakRemainingText").textContent = len(model.breakAllowance);
    $("breakProgressFill").style.width = "0%";
  }

  const permMins = state.today.permissionMinutes || 0;
  $("permDurationDisplay").textContent = `${String(Math.floor(permMins / 60)).padStart(2, "0")}:${String(permMins % 60).padStart(2, "0")}`;
  $("permTimeDisplay").textContent = state.today.permissionTime ? clockText(minutesFromHHMM(state.today.permissionTime)) : "--:--";
  if (document.activeElement !== $("permTypeSelect")) {
    $("permTypeSelect").value = TYPE_TO_UI[state.today.permissionType] || "auto";
  }

  const notes = state.settings.notifications;
  setChecked("alertPreBreakEnabled", notes.preBreak.enabled);
  setNumber("preBreakMinutesInput", notes.preBreak.minutesBefore);
  setChecked("blinkPreBreakToggle", notes.preBreak.blink);
  setChecked("alertBreakEndEnabled", notes.breakEnd.enabled);
  setChecked("blinkBreakEndToggle", notes.breakEnd.blink);
  setChecked("alertCheckoutEnabled", notes.shiftEnd.enabled);
  setNumber("preCheckoutMinutesInput", notes.shiftEnd.minutesBefore || 0);
  setChecked("blinkCheckoutToggle", notes.shiftEnd.blink);
  setChecked("alertPermissionEnabled", notes.permission.enabled);
  setChecked("blinkPermissionToggle", notes.permission.blink);

  setNumber("targetHoursInput", (target / 60).toFixed(1));
  setNumber("alertSecondsInput", state.settings.alertSeconds ?? 5);
  $("officialStartText").textContent = clockText(minutesFromHHMM(state.settings.officialCheckIn));
  $("graceEndText").textContent = clockText(minutesFromHHMM(state.settings.graceEnd));
  const webhookInput = $("webhookUrlInput");
  if (webhookInput && !ui.webhookEditing && document.activeElement !== webhookInput) {
    webhookInput.value = state.settings.sheetsWebhook || "";
  }
  if (webhookInput) webhookInput.readOnly = Boolean(state.settings.sheetsWebhook) && !ui.webhookEditing;
  const editWebhookBtn = $("editWebhookBtn");
  if (editWebhookBtn) editWebhookBtn.hidden = !webhookInput?.readOnly;
  const pickerStyle = state.settings.timePicker === "scroll" ? "scroll" : "clock";
  const pickerSwitch = $("pickerStyleSwitch");
  if (pickerSwitch) {
    pickerSwitch.dataset.value = pickerStyle;
    pickerSwitch.querySelectorAll("[data-picker-style]").forEach((btn) => {
      const on = btn.dataset.pickerStyle === pickerStyle;
      btn.setAttribute("aria-checked", on ? "true" : "false");
    });
  }

  document.documentElement.classList.toggle("no-shimmer", state.settings.progressShimmer === false);
  document.documentElement.classList.toggle("hide-header-target", state.settings.showHeaderTarget === false);
  const shimmerToggle = $("progressShimmerToggle");
  if (shimmerToggle && document.activeElement !== shimmerToggle) shimmerToggle.checked = state.settings.progressShimmer !== false;
  const headerTargetToggle = $("showHeaderTargetToggle");
  if (headerTargetToggle && document.activeElement !== headerTargetToggle) headerTargetToggle.checked = state.settings.showHeaderTarget !== false;
  const sheetBadge = $("sheetStatusBadge");
  sheetBadge.textContent = ui.pulling ? t("ui.sheetPulling") : state.settings.sheetsConnected ? t("ui.sheetsOn") : t("ui.sheetsOff");
  sheetBadge.className = ui.pulling
    ? "status-pill pulling"
    : state.settings.sheetsConnected
      ? "status-pill connected"
      : "status-pill";
  const staleNote = $("scriptStaleText");
  if (staleNote) staleNote.classList.toggle("hidden", !linkIsBehindScript(state.settings));
  const added = $("webhookAddedText");
  if (added) {
    const when = state.settings.sheetsWebhook && state.settings.sheetsWebhookAt
      ? formatMinuteStamp(state.settings.sheetsWebhookAt, lang())
      : "";
    added.classList.toggle("hidden", !when);
    added.textContent = when ? t("ui.linkAdded").replace("{when}", when) : "";
  }
  paintScriptDates();
  const syncNote = $("lastSyncText");
  if (syncNote) {
    syncNote.classList.toggle("pulling", ui.pulling);
    if (ui.pulling) syncNote.textContent = t("ui.sheetPulling");
    else if (!state.lastSyncAt) syncNote.textContent = t("ui.syncNever");
    else {
      const when = new Date(state.lastSyncAt);
      const stamp = Number.isNaN(when.getTime()) ? state.lastSyncAt : clockText(when.getHours() * 60 + when.getMinutes());
      syncNote.textContent = `${t("ui.synced")} • ${t("ui.lastSync")} ${stamp}`;
    }
  }
  paintLog(model);
  paintLeave(date);
}

function paintScriptDates() {
  const latest = $("scriptLatestText");
  if (!latest) return;
  const when = formatMinuteStamp(scriptChangedAt(), lang());
  latest.textContent = t("guide.scriptLatest").replace("{when}", when);
}

function logMonthKeys() {
  const keys = new Set();
  for (const row of state.log || []) {
    const key = monthKey(row && row.date);
    if (/^\d{4}-\d{2}$/.test(key)) keys.add(key);
  }
  const current = monthKey(state.today.date);
  if (/^\d{4}-\d{2}$/.test(current)) keys.add(current);
  return [...keys].sort().reverse();
}

function selectedLogMonth(keys) {
  const current = monthKey(state.today.date);
  const month = keys.includes(ui.logMonth) ? ui.logMonth : current;
  const select = $("logMonthSelect");
  if (!select) return month;
  const sig = `${lang()}|${keys.join(",")}`;
  if (sig !== ui.monthSig) {
    ui.monthSig = sig;
    select.innerHTML = keys.map((key) => `<option value="${key}">${formatMonthTitle(key, lang())}</option>`).join("");
  }
  if (select.value !== month) select.value = month;
  return month;
}

function monthRows(model) {
  const keys = logMonthKeys();
  const month = selectedLogMonth(keys);
  const todayEntry = toLogEntry(state.today, model);
  let rows = filterLogByMonth(state.log, month);
  if (monthKey(todayEntry.date) === month) rows = upsertLog(rows, todayEntry);
  return { month, rows };
}

function shownOvertime(row) {
  const openToday = row.date === state.today.date && !state.today.checkOut;
  if (openToday) return Math.max(0, Math.round(Number(row.overtimeMinutes) || 0));
  return effectiveOvertimeMinutes(row, state.settings.weekendDays || [5, 6]);
}

let overtimeDraft = null;

function minutesFromDuration(value) {
  const match = /^(\d{1,3}):(\d{1,2})$/.exec(String(value || "").trim());
  if (!match) return null;
  const minutes = Number(match[2]);
  if (minutes > 59) return null;
  return Number(match[1]) * 60 + minutes;
}

function monthDateBounds(month) {
  const [year, monthNumber] = String(month).split("-").map(Number);
  const last = new Date(year, monthNumber, 0).getDate();
  return { min: `${month}-01`, max: `${month}-${String(last).padStart(2, "0")}` };
}

function overtimeCopyText(month, days, total) {
  const title = t("ui.overtimeTitle", { month: formatMonthTitle(month, lang()) });
  const head = [t("ui.colDay"), t("ui.colDate"), t("ui.colOvertime")].join("\t");
  const lines = days.map((day) => [weekdayLabel(day.date, lang()), day.date, formatHourMinute(day.minutes)].join("\t"));
  const body = lines.length ? [head, ...lines].join("\n") : t("ui.overtimeEmpty");
  return `${title}\n\n${body}\n\n${t("ui.overtimeTotal")}: ${formatHourMinute(total)}`;
}

function syncOvertimeSummary() {
  if (!overtimeDraft) return;
  const counted = overtimeDraft.days.filter((day) => day.minutes > 0);
  const total = counted.reduce((sum, day) => sum + day.minutes, 0);
  $("overtimeReportTotal").textContent = formatHourMinute(total);
  $("overtimeCopySource").textContent = overtimeCopyText(overtimeDraft.month, counted, total);
}

function paintOvertimeDraft() {
  if (!overtimeDraft) return;
  overtimeDraft.days.sort((a, b) => a.date.localeCompare(b.date));
  $("overtimeModalTitle").textContent = `⏱️ ${t("ui.overtimeTitle", { month: formatMonthTitle(overtimeDraft.month, lang()) })}`;
  $("overtimeReportBody").innerHTML = overtimeDraft.days.length
    ? overtimeDraft.days.map((day) => `
      <tr>
        <td>${weekdayLabel(day.date, lang())}${day.fromLog ? "" : ` <span class="ot-badge">${t("ui.overtimeAdded")}</span>`}</td>
        <td class="ltr-num">${day.date}</td>
        <td><input class="text-input ltr-num ot-duration" data-ot-minutes="${day.date}" value="${formatHourMinute(day.minutes)}" inputmode="numeric" aria-label="${t("ui.colOvertime")}"></td>
        <td><button type="button" class="chip-btn" data-ot-remove="${day.date}">${t("ui.overtimeRemove")}</button></td>
      </tr>
    `).join("")
    : `<tr><td colspan="4">${t("ui.overtimeEmpty")}</td></tr>`;
  syncOvertimeSummary();
}

function openOvertimeReport() {
  const { month, rows } = monthRows(modelNow());
  const bounds = monthDateBounds(month);
  overtimeDraft = {
    month,
    days: rows.filter((row) => shownOvertime(row) > 0).map((row) => ({
      date: row.date,
      minutes: shownOvertime(row),
      fromLog: true,
    })),
  };
  const dateInput = $("overtimeAddDate");
  dateInput.min = bounds.min;
  dateInput.max = bounds.max;
  dateInput.value = month === monthKey(state.today.date) ? state.today.date : bounds.min;
  $("overtimeAddHours").value = "0";
  $("overtimeAddMinutes").value = "0";
  paintOvertimeDraft();
  $("overtimeModal").classList.remove("hidden");
  syncPageScroll();
}

function addOvertimeDraftDay(event) {
  event.preventDefault();
  if (!overtimeDraft) return;
  const date = $("overtimeAddDate").value;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    showFeedback(t("ui.overtimeNeedDate"), true);
    return;
  }
  const minutes = clampInt($("overtimeAddHours").value, 0, 48) * 60 + clampInt($("overtimeAddMinutes").value, 0, 59);
  if (minutes <= 0) {
    showFeedback(t("ui.overtimeNeedTime"), true);
    return;
  }
  const existing = overtimeDraft.days.find((day) => day.date === date);
  if (existing) existing.minutes = minutes;
  else overtimeDraft.days.push({ date, minutes, fromLog: false });
  $("overtimeAddHours").value = "0";
  $("overtimeAddMinutes").value = "0";
  paintOvertimeDraft();
}

function closeOvertimeReport() {
  overtimeDraft = null;
  $("overtimeModal").classList.add("hidden");
  syncPageScroll();
}

let personalDraft = null;

function personalCopyText(month, days) {
  const title = t("ui.personalTitle", { month: formatMonthTitle(month, lang()) });
  const head = [t("ui.colDay"), t("ui.colDate")].join("\t");
  const lines = days.map((day) => [weekdayLabel(day.date, lang()), day.date].join("\t"));
  const body = lines.length ? [head, ...lines].join("\n") : t("ui.personalEmpty");
  return `${title}\n\n${body}\n\n${t("ui.personalTotal")}: ${formatDayCount(days.length, lang())}`;
}

function syncPersonalSummary() {
  if (!personalDraft) return;
  $("personalReportTotal").textContent = formatDayCount(personalDraft.days.length, lang());
  $("personalCopySource").textContent = personalCopyText(personalDraft.month, personalDraft.days);
}

function paintPersonalDraft() {
  if (!personalDraft) return;
  personalDraft.days.sort((a, b) => a.date.localeCompare(b.date));
  $("personalModalTitle").textContent = `🌴 ${t("ui.personalTitle", { month: formatMonthTitle(personalDraft.month, lang()) })}`;
  $("personalReportBody").innerHTML = personalDraft.days.length
    ? personalDraft.days.map((day) => `
      <tr>
        <td>${weekdayLabel(day.date, lang())}${day.fromLog ? "" : ` <span class="ot-badge">${t("ui.overtimeAdded")}</span>`}</td>
        <td class="ltr-num">${day.date}</td>
        <td><button type="button" class="chip-btn" data-personal-remove="${day.date}">${t("ui.overtimeRemove")}</button></td>
      </tr>
    `).join("")
    : `<tr><td colspan="3">${t("ui.personalEmpty")}</td></tr>`;
  syncPersonalSummary();
}

function openPersonalReport() {
  const { month, rows } = monthRows(modelNow());
  const bounds = monthDateBounds(month);
  personalDraft = {
    month,
    days: rows.filter((row) => row.status === STATUS.personal).map((row) => ({
      date: row.date,
      fromLog: true,
    })),
  };
  const dateInput = $("personalAddDate");
  dateInput.min = bounds.min;
  dateInput.max = bounds.max;
  dateInput.value = month === monthKey(state.today.date) ? state.today.date : bounds.min;
  paintPersonalDraft();
  $("personalModal").classList.remove("hidden");
  syncPageScroll();
}

function addPersonalDraftDay(event) {
  event.preventDefault();
  if (!personalDraft) return;
  const date = $("personalAddDate").value;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || monthKey(date) !== personalDraft.month) {
    showFeedback(t("ui.overtimeNeedDate"), true);
    return;
  }
  if (personalDraft.days.some((day) => day.date === date)) {
    showFeedback(t("ui.personalDuplicate"), true);
    return;
  }
  personalDraft.days.push({ date, fromLog: false });
  paintPersonalDraft();
}

function closePersonalReport() {
  personalDraft = null;
  $("personalModal").classList.add("hidden");
  syncPageScroll();
}

let leaveMonthDraft = null;
let leaveYearDraft = null;

function reportLeaveYear() {
  const value = $("leaveYearSelect")?.value;
  if (value && value !== "all") return value;
  return isoDateLocal(new Date()).slice(0, 4);
}

function leaveMonthCopyText(draft) {
  const title = t("leave.monthTitle", { month: formatMonthTitle(draft.month, lang()) });
  const head = [t("ui.colDay"), t("ui.colDate"), t("leave.reportNote")].join("\t");
  const lines = draft.days.map((day) => [weekdayLabel(day.date, lang()), day.date, day.note || ""].join("\t"));
  const body = lines.length ? [head, ...lines].join("\n") : t("leave.monthEmpty");
  return `${title}\n\n${body}\n\n${t("ui.overtimeTotal")}: ${formatLeaveAmount(draft.days.length)}`;
}

function leaveYearDayCount(draft) {
  return draft.months.reduce((sum, row) => sum + row.days.length, 0);
}

function leaveYearCopyText(draft) {
  const title = t("leave.yearTitle", { year: draft.year });
  const blocks = draft.months.map((row) => {
    const head = `${formatMonthTitle(row.month, lang())} (${formatLeaveAmount(row.days.length)})`;
    const lines = row.days.map((day) => [weekdayLabel(day.date, lang()), day.date].join("\t"));
    return [head, ...lines].join("\n");
  });
  const body = blocks.length ? blocks.join("\n\n") : t("leave.yearEmpty");
  return `${title}\n\n${body}\n\n${t("ui.overtimeTotal")}: ${formatLeaveAmount(leaveYearDayCount(draft))}`;
}

function syncLeaveMonthSummary() {
  if (!leaveMonthDraft) return;
  $("leaveMonthReportTotal").textContent = formatLeaveAmount(leaveMonthDraft.days.length);
  $("leaveMonthCopySource").textContent = leaveMonthCopyText(leaveMonthDraft);
}

function syncLeaveYearSummary() {
  if (!leaveYearDraft) return;
  $("leaveYearReportTotal").textContent = formatLeaveAmount(leaveYearDayCount(leaveYearDraft));
  $("leaveYearCopySource").textContent = leaveYearCopyText(leaveYearDraft);
}

function paintLeaveMonthDraft() {
  if (!leaveMonthDraft) return;
  leaveMonthDraft.days.sort((a, b) => a.date.localeCompare(b.date));
  $("leaveMonthModalTitle").textContent = `🌴 ${t("leave.monthTitle", { month: formatMonthTitle(leaveMonthDraft.month, lang()) })}`;
  $("leaveMonthReportBody").innerHTML = leaveMonthDraft.days.length
    ? leaveMonthDraft.days.map((day) => `
      <tr>
        <td>${weekdayLabel(day.date, lang())}${day.fromLog ? "" : ` <span class="ot-badge">${t("ui.overtimeAdded")}</span>`}</td>
        <td class="ltr-num">${day.date}</td>
        <td><input class="text-input" data-leave-month-note="${day.date}" value="${escapeHtml(day.note || "")}" maxlength="500" aria-label="${t("leave.reportNote")}"></td>
        <td><button type="button" class="chip-btn" data-leave-month-remove="${day.date}">${t("ui.overtimeRemove")}</button></td>
      </tr>
    `).join("")
    : `<tr><td colspan="4">${t("leave.monthEmpty")}</td></tr>`;
  syncLeaveMonthSummary();
}

function paintLeaveYearDraft() {
  if (!leaveYearDraft) return;
  leaveYearDraft.months.sort((a, b) => a.month.localeCompare(b.month));
  leaveYearDraft.months.forEach((row) => row.days.sort((a, b) => a.date.localeCompare(b.date)));
  $("leaveYearModalTitle").textContent = `📆 ${t("leave.yearTitle", { year: leaveYearDraft.year })}`;
  $("leaveYearReportBody").innerHTML = leaveYearDraft.months.length
    ? leaveYearDraft.months.map((row) => row.days.map((day, index) => `
      <tr>
        ${index === 0 ? `<td class="leave-year-month" rowspan="${row.days.length}">${formatMonthTitle(row.month, lang())}${row.fromLog ? "" : ` <span class="ot-badge">${t("ui.overtimeAdded")}</span>`}</td>` : ""}
        <td>${weekdayLabel(day.date, lang())}${day.fromLog ? "" : ` <span class="ot-badge">${t("ui.overtimeAdded")}</span>`}</td>
        <td class="ltr-num">${day.date}</td>
        ${index === 0 ? `<td class="ltr-num leave-year-count" rowspan="${row.days.length}">${formatLeaveAmount(row.days.length)}</td>` : ""}
        <td><button type="button" class="chip-btn" data-leave-year-day="${day.date}">${t("ui.overtimeRemove")}</button></td>
      </tr>
    `).join("")).join("")
    : `<tr><td colspan="5">${t("leave.yearEmpty")}</td></tr>`;
  syncLeaveYearSummary();
}

function openLeaveMonthReport() {
  closeLeaveYearReport();
  const month = isoDateLocal(new Date()).slice(0, 7);
  const bounds = monthDateBounds(month);
  leaveMonthDraft = {
    month,
    days: leaveDaysInMonth(state.leave, month).map((day) => ({ ...day, fromLog: true })),
  };
  const dateInput = $("leaveMonthAddDate");
  dateInput.min = bounds.min;
  dateInput.max = bounds.max;
  dateInput.value = month === monthKey(state.today.date) ? state.today.date : bounds.min;
  $("leaveMonthAddNote").value = "";
  paintLeaveMonthDraft();
  $("leaveMonthModal").classList.remove("hidden");
  syncPageScroll();
}

function openLeaveYearReport() {
  closeLeaveMonthReport();
  const year = reportLeaveYear();
  leaveYearDraft = {
    year,
    months: leaveDaysByMonthInYear(state.leave, year).map((row) => ({
      month: row.month,
      fromLog: true,
      days: row.days.map((day) => ({ ...day, fromLog: true })),
    })),
  };
  const dateInput = $("leaveYearAddDate");
  const today = isoDateLocal(new Date());
  dateInput.min = `${year}-01-01`;
  dateInput.max = `${year}-12-31`;
  dateInput.value = today.startsWith(`${year}-`) ? today : `${year}-01-01`;
  paintLeaveYearDraft();
  $("leaveYearModal").classList.remove("hidden");
  syncPageScroll();
}

function addLeaveMonthDraftDay(event) {
  event.preventDefault();
  if (!leaveMonthDraft) return;
  const date = $("leaveMonthAddDate").value;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || monthKey(date) !== leaveMonthDraft.month) {
    showFeedback(t("ui.overtimeNeedDate"), true);
    return;
  }
  if (leaveMonthDraft.days.some((day) => day.date === date)) {
    showFeedback(t("leave.reportDuplicate"), true);
    return;
  }
  leaveMonthDraft.days.push({ date, note: $("leaveMonthAddNote").value.trim().slice(0, 500), fromLog: false });
  $("leaveMonthAddNote").value = "";
  paintLeaveMonthDraft();
}

function addLeaveYearDraftDay(event) {
  event.preventDefault();
  if (!leaveYearDraft) return;
  const date = $("leaveYearAddDate").value;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !date.startsWith(`${leaveYearDraft.year}-`)) {
    showFeedback(t("ui.overtimeNeedDate"), true);
    return;
  }
  if (leaveYearDraft.months.some((row) => row.days.some((day) => day.date === date))) {
    showFeedback(t("leave.reportDuplicate"), true);
    return;
  }
  const month = date.slice(0, 7);
  const existing = leaveYearDraft.months.find((row) => row.month === month);
  if (existing) existing.days.push({ date, note: "", fromLog: false });
  else leaveYearDraft.months.push({ month, fromLog: false, days: [{ date, note: "", fromLog: false }] });
  paintLeaveYearDraft();
}

function closeLeaveMonthReport() {
  leaveMonthDraft = null;
  $("leaveMonthModal")?.classList.add("hidden");
  syncPageScroll();
}

function closeLeaveYearReport() {
  leaveYearDraft = null;
  $("leaveYearModal")?.classList.add("hidden");
  syncPageScroll();
}

async function copyReportText(id) {
  const text = $(id)?.textContent || "";
  if (!text) {
    showFeedback(t("toast.copyFail"), true);
    return;
  }
  if (copyWithCommand(text)) {
    showFeedback(t("toast.copiedReport"));
    return;
  }
  try {
    await navigator.clipboard.writeText(text);
    showFeedback(t("toast.copiedReport"));
  } catch {
    showFeedback(t("toast.copyFail"), true);
  }
}

async function copyPersonalReport() {
  const text = $("personalCopySource")?.textContent || "";
  if (!text) {
    showFeedback(t("toast.copyFail"), true);
    return;
  }
  if (copyWithCommand(text)) {
    showFeedback(t("toast.copiedReport"));
    return;
  }
  try {
    await navigator.clipboard.writeText(text);
    showFeedback(t("toast.copiedReport"));
  } catch {
    showFeedback(t("toast.copyFail"), true);
  }
}

async function copyOvertimeReport() {
  const text = $("overtimeCopySource")?.textContent || "";
  if (!text) {
    showFeedback(t("toast.copyFail"), true);
    return;
  }
  if (copyWithCommand(text)) {
    showFeedback(t("toast.copiedReport"));
    return;
  }
  try {
    await navigator.clipboard.writeText(text);
    showFeedback(t("toast.copiedReport"));
  } catch {
    showFeedback(t("toast.copyFail"), true);
  }
}

function paintLog(model) {
  const { month, rows } = monthRows(model);
  const today = state.today.date;
  const sig = `${lang()}|${month}|${rows.map((row) => `${row.date}:${row.netMinutes}:${shownOvertime(row)}:${row.status}:${row.checkIn}:${row.checkOut}:${row.date === today ? 1 : 0}`).join(",")}`;
  const net = sumNetMinutes(rows);
  const office = rows.filter((row) => row.status === STATUS.work).length;
  $("monthlyStatsSummary").textContent = lang() === "en"
    ? `Work days: ${office} • Net: ${len(net)}`
    : `أيام العمل: ${office} • إجمالي الصافي: ${len(net)}`;
  if (sig === ui.logSig) return;
  ui.logSig = sig;
  $("monthlyLogBody").innerHTML = rows.map((row) => `
    <tr class="${row.date === today ? "is-today" : ""}">
      <td class="ltr-num">${row.date}</td>
      <td>${weekdayLabel(row.date, lang())}</td>
      <td>${statusText(row.status)}</td>
      <td class="ltr-num">${row.checkIn ? clockText(minutesFromHHMM(row.checkIn)) : "--:--"}</td>
      <td class="ltr-num">${row.checkOut ? clockText(minutesFromHHMM(row.checkOut)) : "--:--"}</td>
      <td class="ltr-num">${len(row.breakMinutes)}</td>
      <td class="ltr-num">${len(row.permissionMinutes)}</td>
      <td class="ltr-num">${len(row.netMinutes)}</td>
      <td class="ltr-num overtime-cell">${formatHourMinute(shownOvertime(row))}</td>
      <td>
        <button type="button" class="chip-btn" data-edit-date="${row.date}">${t("ui.edit")}</button>
        <button type="button" class="chip-btn" data-delete-date="${row.date}">${t("ui.delete")}</button>
      </td>
    </tr>
  `).join("");
}

function armPermission(minutes, type) {
  state.today.permissionMinutes = clampInt(minutes, 0, 12 * 60);
  if (type) state.today.permissionType = type;
  state.today.permissionTouched = true;
  state.today.permissionOn = state.today.permissionMinutes > 0 || state.today.permissionType !== "auto";
}

function setManual(minutes) {
  state.today.manualBreakMinutes = clampInt(minutes, 0, 480);
  if (state.today.manualBreakMinutes > 0 && !state.today.manualBreakStart) {
    state.today.manualBreakStart = state.today.checkIn || hhmmNow(new Date());
  }
  if (state.today.manualBreakMinutes === 0) state.today.manualBreakStart = null;
}

function rememberToday() {
  state.log = upsertLog(state.log, toLogEntry(state.today, modelNow()));
  state.dismissedDates = (state.dismissedDates || []).filter((date) => date !== state.today.date);
  ui.logSig = "";
}

function recordsForSync(model) {
  const labels = lang() === "en" ? STATUS_LABELS_EN : STATUS_LABELS;
  return upsertLog(state.log, toLogEntry(state.today, model)).map((row) => ({
    ...row,
    status: labels[row.status] || row.status,
    checkedOut: row.date === state.today.date ? Boolean(state.today.checkOut) : "1",
  }));
}

function scheduleSync() {
  if (!ui.pullReady || !state.settings.sheetsConnected || !state.settings.sheetsWebhook) return;
  clearTimeout(ui.syncTimer);
  ui.syncTimer = setTimeout(() => { void pushSync(false); }, 2500);
}

function applyPulled(records) {
  const webhook = state.settings.sheetsWebhook;
  const result = applySheetRecords(state, records, isoDateLocal(new Date()));
  if (!result.imported) return 0;
  state = result.state;
  state.settings.sheetsWebhook = webhook;
  state.settings.sheetsConnected = true;
  ui.logSig = "";
  return result.imported;
}

function clearLocalLog() {
  const today = isoDateLocal(new Date());
  state.today = createToday(today, { permissionMinutes: state.settings.defaultPermissionMinutes });
  state.log = [];
  state.dismissedDates = [];
  state.notifiedKeys = [];
  ui.logSig = "";
}

async function sheetOwnership(url) {
  try {
    await pingWebhook(url);
    return "ours";
  } catch (error) {
    return error && error.message === "webhook" ? "foreign" : "unreachable";
  }
}

async function confirmSheetOnOpen(shouldPull) {
  const verdict = await sheetOwnership(state.settings.sheetsWebhook);
  if (verdict === "foreign") {
    state.settings.sheetsConnected = false;
    ui.pulling = false;
    ui.pullReady = false;
    if (shouldPull) state = fillMissingDays(state);
    persist();
    paint();
    showFeedback(t("toast.notOurs"), true);
    return;
  }
  if (verdict === "unreachable") {
    ui.pulling = false;
    if (shouldPull) {
      state = fillMissingDays(state);
      persist();
      paint();
      showFeedback(t("toast.syncFail"), true);
    }
    return;
  }
  if (!shouldPull) {
    if (state.settings.fileSnapshot && state.settings.sheetsImport !== "keep") {
      state.settings.sheetsConnected = true;
      ui.pullReady = true;
      persist({ sync: true });
    }
    return;
  }
  const ok = await pullSync({ announce: true });
  if (ok) return;
  state = fillMissingDays(state);
  ui.logSig = "";
  persist();
  paint();
}

async function pullSync({ announce = false, replaceEmpty = false } = {}) {
  const url = state.settings.sheetsWebhook;
  if (!url) {
    if (announce) showFeedback(t("toast.syncNeed"), true);
    return false;
  }
  if (ui.syncing) return false;
  clearTimeout(ui.syncTimer);
  ui.syncing = true;
  ui.pulling = true;
  ui.pullReady = false;
  paint();
  if (announce) showFeedback(t("toast.sheetPulling"), "warn");
  try {
    const data = await pullWebhook(url);
    if (!Array.isArray(data.records)) {
      state.settings.sheetsConnected = false;
      ui.pullReady = false;
      persist();
      if (announce) showFeedback(t("toast.scriptUpdate"), true);
      return false;
    }
    const remoteSettings = Array.isArray(data.settings) ? data.settings : null;
    state.settings.sheetsConnected = true;
    if (remoteSettings?.length) {
      state.settings = overlaySettings(state.settings, remoteSettings, { keepLink: true });
      applyLang();
    }
    const imported = applyPulled(data.records);
    if (data.leave && typeof data.leave === "object" && !Array.isArray(data.leave)) {
      state.leave = sanitizeLeave(data.leave);
      ui.leaveSig = "";
    }
    if (!imported && replaceEmpty) clearLocalLog();
    state.lastSyncAt = new Date().toISOString();
    ui.pullReady = true;
    ui.logSig = "";
    persist({ sync: remoteSettings?.length === 0 });
    if (announce && (linkIsBehindScript(state.settings) || !remoteSettings)) showFeedback(t("toast.scriptUpdate"), true);
    else if (announce) showFeedback(imported || replaceEmpty ? t("toast.sheetImported") : t("toast.connected"));
    return true;
  } catch {
    ui.pullReady = false;
    if (announce) showFeedback(t("toast.syncFail"), true);
    return false;
  } finally {
    ui.pulling = false;
    ui.syncing = false;
    paint();
  }
}

async function pushSync(manual) {
  if (!ui.pullReady || !state.settings.sheetsWebhook || !state.settings.sheetsConnected) {
    if (manual) showFeedback(t("toast.syncNeed"), true);
    return;
  }
  if (ui.syncing) return;
  ui.syncing = true;
  try {
    await pushWebhook(state.settings.sheetsWebhook, {
      signature: APP_SIGNATURE,
      sentAt: new Date().toISOString(),
      records: recordsForSync(modelNow()),
      settings: settingsPairs(state.settings),
      leave: leaveSyncPayload(state.leave),
    });
    state.lastSyncAt = new Date().toISOString();
    if (state.settings.fileSnapshot) state.settings.fileSnapshot = false;
    persist();
    if (manual) showFeedback(t("toast.syncOk"));
  } catch {
    if (manual) showFeedback(t("toast.syncFail"), true);
  } finally {
    ui.syncing = false;
  }
}

function rememberWebhook(url) {
  const next = String(url || "").trim();
  state.settings.sheetsWebhookAt = stampWebhook(
    state.settings.sheetsWebhook,
    next,
    state.settings.sheetsWebhookAt,
    new Date().toISOString(),
  );
  state.settings.sheetsWebhook = next;
}

async function connectAndPull() {
  const url = $("webhookUrlInput").value.trim();
  rememberWebhook(url);
  ui.pullReady = false;
  if (!url) {
    state.settings.sheetsConnected = false;
    ui.pulling = false;
    persist();
    paint();
    showFeedback(t("toast.url"), true);
    return;
  }
  ui.pulling = true;
  paint();
  showFeedback(t("toast.sheetPulling"), "warn");
  try {
    await pingWebhook(url);
  } catch {
    ui.pulling = false;
    state.settings.sheetsConnected = false;
    persist();
    paint();
    showFeedback(t("toast.connectFail"), true);
    return;
  }
  await pullSync({ announce: true, replaceEmpty: true });
}

function openEditWebhook() {
  if (!state.settings.sheetsWebhook) return;
  $("editWebhookModal").classList.remove("hidden");
  syncPageScroll();
}

function closeEditWebhook() {
  $("editWebhookModal").classList.add("hidden");
  syncPageScroll();
}

function confirmEditWebhook() {
  closeEditWebhook();
  ui.webhookEditing = true;
  paint();
  const input = $("webhookUrlInput");
  input.focus();
  input.select();
}

function openImportChoice(url) {
  ui.pendingWebhook = url;
  $("importChoiceModal").classList.remove("hidden");
  syncPageScroll();
}

function hideImportChoice() {
  $("importChoiceModal").classList.add("hidden");
  ui.pendingWebhook = "";
  syncPageScroll();
}

function cancelImportChoice() {
  if ($("importChoiceModal").classList.contains("hidden")) return;
  const typed = ui.pendingWebhook;
  hideImportChoice();
  if (state.settings.sheetsWebhook) {
    ui.webhookEditing = false;
    $("webhookUrlInput").value = state.settings.sheetsWebhook;
  } else if (typed) {
    $("webhookUrlInput").value = typed;
  }
  paint();
}

async function keepLocalLink() {
  ui.pullReady = false;
  try {
    await pingWebhook(state.settings.sheetsWebhook);
    state.settings.sheetsConnected = true;
    state.settings.sheetsImport = "keep";
    persist();
    showFeedback(t("toast.sheetKept"));
  } catch {
    state.settings.sheetsConnected = false;
    persist();
    showFeedback(t("toast.connectFail"), true);
  } finally {
    paint();
  }
}

async function chooseSheetImport(replace) {
  const url = (ui.pendingWebhook || "").trim();
  hideImportChoice();
  if (!url) return;
  rememberWebhook(url);
  state.settings.sheetsImport = replace ? "replace" : "keep";
  ui.webhookEditing = false;
  ui.pullReady = false;
  $("webhookUrlInput").value = url;
  if (replace) {
    await connectAndPull();
    return;
  }
  await keepLocalLink();
}

function commitWebhook() {
  const url = $("webhookUrlInput").value.trim();
  const previous = state.settings.sheetsWebhook || "";
  if (!url) {
    rememberWebhook("");
    state.settings.sheetsConnected = false;
    state.settings.sheetsImport = "";
    ui.pullReady = false;
    ui.pulling = false;
    ui.webhookEditing = false;
    persist();
    paint();
    showFeedback(t("toast.url"), true);
    return;
  }
  if (url !== previous) {
    openImportChoice(url);
    return;
  }
  ui.webhookEditing = false;
  paint();
  if (state.settings.sheetsImport === "keep") return;
  void pullSync({ announce: true });
}

function blinkFor(kind) {
  const notes = state.settings.notifications;
  const flag = {
    "pre-break": notes.preBreak.blink,
    "break-start": notes.preBreak.blink,
    "break-end": notes.breakEnd.blink,
    "shift-end": notes.shiftEnd.blink,
    permission: notes.permission.blink,
  }[kind];
  return state.settings.tabBlinkEnabled && flag !== false;
}

function runAlerts(minute) {
  const alerts = dueAlerts({
    nowMin: minute,
    model: computeDay(inputFromState(state, minute)),
    notifications: state.settings.notifications,
    firedKeys: state.notifiedKeys,
  });
  if (!alerts.length) return;
  state.notifiedKeys.push(...alerts.map((alert) => alert.id));
  persist();
  if (state.settings.soundEnabled) playChime();
  for (const alert of alerts) {
    if ("Notification" in window && Notification.permission === "granted") {
      try { new Notification(alert.title, { body: alert.body }); } catch { /* the browser can reject a notification */ }
    }
  }
  const blink = alerts.find((alert) => blinkFor(alert.kind));
  if (blink) startTabBlink(blink.title);
  showFeedback(alerts[0].body || alerts[0].title);
}

function pingAlerts(date = new Date()) {
  const minute = date.getHours() * 60 + date.getMinutes();
  ui.lastMinute = minute;
  runAlerts(minute);
}

function secondsUntil(targetMin, date) {
  const now = date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds();
  return Math.max(0, targetMin * 60 - now);
}

function formatCountdown(totalSeconds) {
  const safe = Math.max(0, Math.floor(totalSeconds));
  const hours = Math.floor(safe / 3600);
  const minutes = Math.floor((safe % 3600) / 60);
  const seconds = safe % 60;
  return [hours, minutes, seconds].map((part) => String(part).padStart(2, "0")).join(":");
}

const logEdit = { date: null, status: "personal", checkIn: "", checkOut: "" };

function paintLogEdit() {
  $("logEditTitle").textContent = logEdit.date || "";
  $("logEditStatus").value = logEdit.status;
  $("logEditCheckIn").textContent = logEdit.checkIn ? clockText(minutesFromHHMM(logEdit.checkIn)) : "--:--";
  $("logEditCheckOut").textContent = logEdit.checkOut ? clockText(minutesFromHHMM(logEdit.checkOut)) : "--:--";
}

function openLogEdit(date) {
  const stored = state.log.find((row) => row.date === date);
  const live = date === state.today.date ? toLogEntry(state.today, modelNow()) : null;
  const row = stored || live;
  if (!row) return;
  logEdit.date = date;
  logEdit.status = ["work", "weekend", "personal", "official"].includes(row.status) ? row.status : "personal";
  logEdit.checkIn = row.checkIn || "";
  logEdit.checkOut = row.checkOut || "";
  paintLogEdit();
  $("logEditModal").classList.remove("hidden");
  syncPageScroll();
}

function closeLogEdit() {
  $("logEditModal").classList.add("hidden");
  logEdit.date = null;
  syncPageScroll();
}

function saveLogEdit() {
  if (!logEdit.date) return;
  const date = logEdit.date;
  const before = date === state.today.date
    ? modelNow().status
    : (state.log.find((row) => row.date === date)?.status || "");
  logEdit.status = $("logEditStatus").value;
  if (logEdit.date === state.today.date) {
    state.today.statusOverride = logEdit.status;
    state.today.workplace = "office";
    state.today.checkIn = logEdit.checkIn || null;
    state.today.checkOut = logEdit.checkOut || null;
    rememberToday();
  } else {
    const current = state.log.find((row) => row.date === logEdit.date);
    const next = {
      ...(current || leaveEntry(logEdit.date, state.settings.weekendDays)),
      status: logEdit.status,
      checkIn: logEdit.checkIn || "",
      checkOut: logEdit.checkOut || "",
    };
    state.log = upsertLog(state.log, next);
    state.dismissedDates = (state.dismissedDates || []).filter((date) => date !== logEdit.date);
    ui.logSig = "";
  }
  closeLogEdit();
  persist({ sync: true });
  paint();
  offerLeaveLink(date, before, logEdit.status);
}

let pendingLeaveLink = null;

function leaveDateLabel(date) {
  return `${date} • ${weekdayLabel(date, lang())}`;
}

function openLeaveLink(mode, date) {
  pendingLeaveLink = { mode, date };
  const label = leaveDateLabel(date);
  const remove = mode === "remove";
  $("leaveLinkTitle").textContent = t(remove ? "leave.linkKeepTitle" : "leave.linkAddTitle");
  $("leaveLinkLabel").textContent = label;
  $("leaveLinkText").textContent = t(remove ? "leave.linkKeep" : "leave.linkAdd", { date: label });
  $("leaveLinkRejectBtn").textContent = t(remove ? "leave.linkKeepYes" : "leave.linkAddNo");
  const accept = $("leaveLinkAcceptBtn");
  accept.textContent = t(remove ? "leave.linkDrop" : "leave.linkAddYes");
  accept.classList.toggle("btn-danger", remove);
  accept.classList.toggle("btn-primary", !remove);
  $("leaveLinkModal").classList.remove("hidden");
  syncPageScroll();
}

function closeLeaveLink() {
  $("leaveLinkModal").classList.add("hidden");
  pendingLeaveLink = null;
  syncPageScroll();
}

function offerLeaveLink(date, before, after) {
  const wasPersonal = before === STATUS.personal;
  const isPersonal = after === STATUS.personal;
  const exists = (state.leave?.entries || []).some((entry) => entry.date === date);
  if (!wasPersonal && isPersonal && !exists) {
    openLeaveLink("add", date);
    return;
  }
  if (wasPersonal && !isPersonal && exists) {
    openLeaveLink("remove", date);
    return;
  }
  if (pendingLeaveLink?.date === date) closeLeaveLink();
}

function confirmLeaveLink() {
  const pending = pendingLeaveLink;
  if (!pending) return;
  closeLeaveLink();
  const todayMonth = isoDateLocal(new Date()).slice(0, 7);
  if (pending.mode === "add") {
    const result = addLeaveEntries(state.leave, { start: pending.date, end: pending.date, source: "hours" });
    state.leave = coverLeaveHorizon(result.leave, todayMonth);
    showFeedback(t("leave.addedBalance", { date: pending.date }));
  } else {
    state.leave = removeLeaveEntry(state.leave, pending.date);
    showFeedback(t("leave.removedBalance", { date: pending.date }));
  }
  ui.leaveSig = "";
  persist({ sync: true });
  paint();
}

function rejectLeaveLink() {
  const pending = pendingLeaveLink;
  closeLeaveLink();
  if (pending?.mode === "remove") showFeedback(t("leave.keptBalance", { date: pending.date }));
}

const picker = { target: null, duration: false, hour: 8, minute: 30, pm: false, hand: "hour", entry: "dial" };

function buildWheel(wheel, labels, index, onSelect) {
  wheel.innerHTML = "";
  let suppress = true;
  labels.forEach((label, idx) => {
    const item = document.createElement("div");
    item.className = "drum-item";
    item.textContent = label;
    item.addEventListener("click", () => {
      suppress = true;
      wheel.scrollTop = idx * ITEM_HEIGHT;
      markWheel(wheel, idx);
      onSelect(idx);
      previewPicker();
      setTimeout(() => { suppress = false; }, 60);
    });
    wheel.appendChild(item);
  });
  wheel.scrollTop = index * ITEM_HEIGHT;
  markWheel(wheel, index);
  wheel.onscroll = () => {
    if (suppress) return;
    const idx = Math.max(0, Math.min(labels.length - 1, Math.round(wheel.scrollTop / ITEM_HEIGHT)));
    markWheel(wheel, idx);
    onSelect(idx);
    previewPicker();
  };
  setTimeout(() => { suppress = false; }, 80);
}

function markWheel(wheel, index) {
  wheel.querySelectorAll(".drum-item").forEach((el, idx) => el.classList.toggle("selected", idx === index));
}

function previewPicker() {
  const period = picker.duration ? "" : ` ${picker.pm ? (lang() === "en" ? "PM" : "م") : (lang() === "en" ? "AM" : "ص")}`;
  $("pickerLivePreview").textContent = `${String(picker.hour).padStart(2, "0")}:${String(picker.minute).padStart(2, "0")}${period}`;
  const hourPart = $("clockHourPart");
  const minutePart = $("clockMinutePart");
  if (hourPart && minutePart) {
    hourPart.textContent = String(picker.hour).padStart(2, "0");
    minutePart.textContent = String(picker.minute).padStart(2, "0");
    hourPart.classList.toggle("active", picker.hand !== "minute");
    minutePart.classList.toggle("active", picker.hand === "minute");
  }
  const hourInput = $("clockHourInput");
  const minuteInput = $("clockMinuteInput");
  if (hourInput && document.activeElement !== hourInput) hourInput.value = String(picker.hour).padStart(2, "0");
  if (minuteInput && document.activeElement !== minuteInput) minuteInput.value = String(picker.minute).padStart(2, "0");
  document.querySelectorAll("[data-clock-period]").forEach((btn) => {
    const pm = btn.dataset.clockPeriod === "pm";
    btn.textContent = lang() === "en" ? (pm ? "PM" : "AM") : (pm ? "م" : "ص");
    btn.setAttribute("aria-pressed", ((!pm && !picker.pm) || (pm && picker.pm)) ? "true" : "false");
  });
}

function minuteFromPointer(event) {
  const face = $("clockFace");
  const rect = face.getBoundingClientRect();
  const x = event.clientX - (rect.left + rect.width / 2);
  const y = event.clientY - (rect.top + rect.height / 2);
  let angle = Math.atan2(y, x) + Math.PI / 2;
  if (angle < 0) angle += Math.PI * 2;
  return Math.round((angle / (Math.PI * 2)) * 60) % 60;
}

function hourFromPointer(event) {
  const face = $("clockFace");
  const rect = face.getBoundingClientRect();
  const x = event.clientX - (rect.left + rect.width / 2);
  const y = event.clientY - (rect.top + rect.height / 2);
  let angle = Math.atan2(y, x) + Math.PI / 2;
  if (angle < 0) angle += Math.PI * 2;
  const step = Math.round((angle / (Math.PI * 2)) * 12) % 12;
  if (picker.duration) return step;
  return step === 0 ? 12 : step;
}

function renderClock() {
  const face = $("clockFace");
  if (!face) return;
  face.innerHTML = "";
  const hourMode = picker.hand !== "minute";
  if (!hourMode) {
    for (let minute = 0; minute < 60; minute += 1) {
      const angle = (minute / 60) * Math.PI * 2 - Math.PI / 2;
      const tick = document.createElement("span");
      tick.className = "clock-tick";
      tick.style.left = `${50 + Math.cos(angle) * 46}%`;
      tick.style.top = `${50 + Math.sin(angle) * 46}%`;
      face.appendChild(tick);
    }
  }
  for (let i = 0; i < 12; i += 1) {
    const angle = (i / 12) * Math.PI * 2 - Math.PI / 2;
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "clock-num";
    if (hourMode) btn.textContent = picker.duration ? String(i) : String(i === 0 ? 12 : i);
    else btn.textContent = String(i * 5).padStart(2, "0");
    const selected = hourMode
      ? (picker.duration ? picker.hour === i : picker.hour % 12 === i)
      : picker.minute === i * 5;
    if (selected) btn.classList.add("selected");
    btn.style.left = `${50 + Math.cos(angle) * 38}%`;
    btn.style.top = `${50 + Math.sin(angle) * 38}%`;
    btn.addEventListener("click", () => {
      if (hourMode) {
        picker.hour = picker.duration ? i : (i === 0 ? 12 : i);
        picker.hand = "minute";
      } else {
        picker.minute = i * 5;
      }
      syncPickerChrome();
    });
    face.appendChild(btn);
  }
  if (!hourMode && picker.minute % 5 !== 0) {
    const angle = (picker.minute / 60) * Math.PI * 2 - Math.PI / 2;
    const bubble = document.createElement("div");
    bubble.className = "clock-minute-bubble";
    bubble.textContent = String(picker.minute).padStart(2, "0");
    bubble.style.left = `${50 + Math.cos(angle) * 38}%`;
    bubble.style.top = `${50 + Math.sin(angle) * 38}%`;
    face.appendChild(bubble);
  }
  const hand = document.createElement("div");
  hand.className = "clock-hand";
  const deg = hourMode ? (picker.hour % 12) * 30 : (picker.minute / 60) * 360;
  hand.style.transform = `rotate(${deg}deg)`;
  face.appendChild(hand);
  const dot = document.createElement("div");
  dot.className = "clock-center";
  face.appendChild(dot);
}

function syncPickerChrome() {
  const clock = state.settings.timePicker !== "scroll";
  $("drumContainer").classList.toggle("hidden", clock);
  $("clockPicker").classList.toggle("hidden", !clock);
  $("clockPeriodRow").classList.toggle("hidden", picker.duration);
  const dial = picker.entry !== "type";
  $("clockFaceWrap").classList.toggle("hidden", !clock || !dial);
  $("clockTypeWrap").classList.toggle("hidden", !clock || dial);
  $("clockDialBtn").setAttribute("aria-pressed", dial ? "true" : "false");
  $("clockTypeBtn").setAttribute("aria-pressed", dial ? "false" : "true");
  if (clock && dial) renderClock();
  previewPicker();
}

function openPicker(target) {
  picker.target = target;
  picker.hand = "hour";
  picker.entry = "dial";
  const titles = {
    checkin: t("ui.checkin"),
    checkout: t("ui.checkoutPick"),
    logCheckIn: t("ui.colIn"),
    logCheckOut: t("ui.colOut"),
    dhuhrStart: t("ui.dhuhr"),
    asrStart: t("ui.asr"),
    permTime: t("ui.permTime"),
    officialStart: t("ui.officialStart"),
    graceEnd: t("ui.grace"),
    permDuration: t("ui.permDuration"),
  };
  $("pickerModalTitle").textContent = titles[target] || t("ui.checkin");
  if (target === "permDuration") {
    picker.duration = true;
    $("drumPeriodCol").classList.add("hidden");
    const total = state.today.permissionMinutes || 0;
    picker.hour = Math.min(12, Math.floor(total / 60));
    picker.minute = total % 60;
    buildWheel($("drumHourWheel"), Array.from({ length: 13 }, (_, i) => String(i).padStart(2, "0")), picker.hour, (idx) => { picker.hour = idx; });
    buildWheel($("drumMinuteWheel"), Array.from({ length: 60 }, (_, i) => String(i).padStart(2, "0")), picker.minute, (idx) => { picker.minute = idx; });
  } else {
    picker.duration = false;
    $("drumPeriodCol").classList.remove("hidden");
    const source = {
      checkin: state.today.checkIn || state.settings.officialCheckIn,
      checkout: state.today.checkOut || hhmmFromMinutes(modelNow().expectedCheckout || 17 * 60),
      logCheckIn: logEdit.checkIn || state.settings.officialCheckIn,
      logCheckOut: logEdit.checkOut || hhmmFromMinutes(modelNow().expectedCheckout || 17 * 60),
      dhuhrStart: state.settings.dhuhr.start,
      asrStart: state.settings.asr.start,
      permTime: state.today.permissionTime || "13:00",
      officialStart: state.settings.officialCheckIn,
      graceEnd: state.settings.graceEnd,
    }[target] || "08:30";
    const mins = minutesFromHHMM(source) ?? 8 * 60 + 30;
    const parts = clockParts(mins);
    picker.hour = parts.hour12;
    picker.minute = parts.minute;
    picker.pm = parts.pm;
    const hours = Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, "0"));
    const minutes = Array.from({ length: 60 }, (_, i) => String(i).padStart(2, "0"));
    buildWheel($("drumHourWheel"), hours, picker.hour - 1, (idx) => { picker.hour = idx + 1; });
    buildWheel($("drumMinuteWheel"), minutes, picker.minute, (idx) => { picker.minute = idx; });
    buildWheel($("drumPeriodWheel"), lang() === "en" ? ["AM", "PM"] : ["ص", "م"], picker.pm ? 1 : 0, (idx) => { picker.pm = idx === 1; });
  }
  syncPickerChrome();
  $("drumPickerModal").classList.remove("hidden");
  syncPageScroll();
}

function closePicker() {
  $("drumPickerModal").classList.add("hidden");
  picker.target = null;
  syncPageScroll();
}

function closeGuide() {
  $("gasGuideModal").classList.add("hidden");
  syncPageScroll();
}

function copyWithCommand(text) {
  const area = document.createElement("textarea");
  area.value = text;
  area.setAttribute("readonly", "");
  area.style.position = "fixed";
  area.style.top = "0";
  area.style.insetInlineStart = "0";
  area.style.opacity = "0";
  document.body.appendChild(area);
  area.focus();
  area.select();
  try { area.setSelectionRange(0, text.length); } catch { /* ignore */ }
  let ok = false;
  try { ok = document.execCommand("copy"); } catch { ok = false; }
  area.remove();
  return ok;
}

async function copyGasScript() {
  const text = $("gasScript")?.textContent || "";
  if (!text) {
    showFeedback(t("toast.copyFail"), true);
    return;
  }
  if (copyWithCommand(text)) {
    showFeedback(t("toast.copied"));
    return;
  }
  try {
    await navigator.clipboard.writeText(text);
    showFeedback(t("toast.copied"));
  } catch {
    showFeedback(t("toast.copyFail"), true);
  }
}

let pendingDeleteDate = null;
let pendingDeleteKind = "log";
let pendingLeaveSettings = null;

function openDeleteConfirm(date, kind = "log") {
  pendingDeleteDate = date;
  pendingDeleteKind = kind === "leave" ? "leave" : "log";
  const label = `${date} • ${weekdayLabel(date, lang())}`;
  $("deleteConfirmLabel").textContent = label;
  $("deleteConfirmText").textContent = pendingDeleteKind === "leave"
    ? t("leave.deleteConfirm", { date: label })
    : t("ui.deleteConfirm").replace("{date}", label);
  $("deleteConfirmModal").classList.remove("hidden");
  syncPageScroll();
}

function closeDeleteConfirm() {
  $("deleteConfirmModal").classList.add("hidden");
  pendingDeleteDate = null;
  pendingDeleteKind = "log";
  syncPageScroll();
}

function confirmDelete() {
  const date = pendingDeleteDate;
  const kind = pendingDeleteKind;
  if (!date) return;
  closeDeleteConfirm();
  if (kind === "leave") {
    removeLeaveDate(date);
    showFeedback(t("leave.deleted", { date }));
    return;
  }
  state.log = state.log.filter((row) => row.date !== date);
  if (date !== state.today.date && !(state.dismissedDates || []).includes(date)) {
    state.dismissedDates = [...(state.dismissedDates || []), date];
  }
  ui.logSig = "";
  persist({ sync: true });
  paint();
}

function syncPageScroll() {
  const open = document.querySelector(".modal-backdrop:not(.hidden)");
  document.body.classList.toggle("modal-open", !!open);
}

function confirmPicker() {
  const target = picker.target;
  if (!target) return;
  if (target === "permDuration") {
    armPermission(picker.hour * 60 + picker.minute);
  } else {
    let hour = picker.hour % 12;
    if (picker.pm) hour += 12;
    const value = `${String(hour).padStart(2, "0")}:${String(picker.minute).padStart(2, "0")}`;
    if (target === "checkin") state.today.checkIn = value;
    else if (target === "checkout") state.today.checkOut = value;
    else if (target === "logCheckIn") logEdit.checkIn = value;
    else if (target === "logCheckOut") logEdit.checkOut = value;
    else if (target === "dhuhrStart") state.settings.dhuhr.start = value;
    else if (target === "asrStart") state.settings.asr.start = value;
    else if (target === "permTime") state.today.permissionTime = value;
    else if (target === "officialStart") state.settings.officialCheckIn = value;
    else if (target === "graceEnd") state.settings.graceEnd = value;
  }
  closePicker();
  if (target === "logCheckIn" || target === "logCheckOut") {
    paintLogEdit();
    return;
  }
  if (target === "checkin" || target === "checkout") rememberToday();
  persist({ sync: true });
  paint();
  if (target === "dhuhrStart" || target === "asrStart" || target === "permDuration") pingAlerts();
}

function onClick(event) {
  const target = event.target instanceof Element ? event.target : event.target.parentElement;
  if (!target) return;
  if (target.closest("#checkInNowBtn") || target.closest("#checkOutNowBtn")) return;
  if (target.closest("#clearCheckoutBtn")) {
    state.today.checkOut = null;
    state.log = state.log.map((row) => row.date === state.today.date ? { ...row, checkOut: "" } : row);
    ui.logSig = "";
    persist({ sync: true });
    paint();
    showFeedback(lang() === "en" ? "Check-out removed. The day is still open." : "تم حذف الانصراف. اليوم لسه مفتوح.");
    return;
  }
  const pickerBtn = target.closest("[data-open-picker]");
  if (pickerBtn) { openPicker(pickerBtn.dataset.openPicker); return; }
  const accent = target.closest("[data-accent-btn]");
  if (accent) {
    state.settings.accent = normalizeAccent(accent.dataset.accentBtn);
    applyTheme();
    persist({ sync: true });
    paint();
    return;
  }
  const preset = target.closest("[data-perm-preset]");
  if (preset) {
    armPermission(Number(preset.dataset.permPreset) || 0);
    persist({ sync: true });
    paint();
    return;
  }
  const edit = target.closest("[data-edit-date]");
  if (edit) {
    openLogEdit(edit.dataset.editDate);
    return;
  }
  const del = target.closest("[data-delete-date]");
  if (del) {
    openDeleteConfirm(del.dataset.deleteDate);
    return;
  }
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function setLeaveSpan(span) {
  ui.leaveSpan = span === "range" ? "range" : "single";
  const end = $("leaveEndWrap");
  if (end) end.classList.toggle("hidden", ui.leaveSpan !== "range");
  document.querySelectorAll("[data-leave-span]").forEach((btn) => {
    const on = btn.dataset.leaveSpan === ui.leaveSpan;
    btn.classList.toggle("is-active", on);
    btn.setAttribute("aria-pressed", on ? "true" : "false");
  });
}

function setView(view) {
  ui.view = view === "leave" ? "leave" : "hours";
  $("hoursView")?.classList.toggle("hidden", ui.view !== "hours");
  $("leaveView")?.classList.toggle("hidden", ui.view !== "leave");
  document.querySelectorAll(".view-tab").forEach((btn) => {
    const on = btn.dataset.view === ui.view;
    btn.classList.toggle("is-active", on);
    btn.setAttribute("aria-selected", on ? "true" : "false");
  });
  paint();
}

function leaveChipHtml(entry) {
  const fromHours = entry.source === "hours";
  const note = entry.note ? ` — ${entry.note}` : "";
  const origin = t(fromHours ? "leave.legendHours" : "leave.legendManual");
  const label = `${formatDateLine(entry.date, lang())}${note} — ${origin}`;
  return `<button type="button" class="leave-chip ${fromHours ? "leave-chip-hours" : "leave-chip-manual"}" data-leave-remove="${escapeHtml(entry.date)}" title="${escapeHtml(label)}" aria-label="${escapeHtml(t("leave.removeDate", { date: entry.date }))}"><span class="ltr-num">${escapeHtml(entry.date)}</span><span class="leave-chip-x" aria-hidden="true">×</span></button>`;
}

function selectedLeaveYear(years, todayMonth) {
  if (ui.leaveYear === "all") return "all";
  if (ui.leaveYear && years.includes(ui.leaveYear)) return ui.leaveYear;
  const current = String(todayMonth).slice(0, 4);
  if (years.includes(current)) return current;
  return years[0] || current;
}

function setLeaveAmount(id, value, negative) {
  const el = $(id);
  if (!el) return;
  el.textContent = value == null ? "—" : formatLeaveAmount(value);
  el.classList.toggle("is-negative", !!negative);
}

function paintLeave(date = new Date()) {
  if (!$("leaveTableBody")) return;
  const todayMonth = isoDateLocal(date).slice(0, 7);
  const covered = coverLeaveHorizon(state.leave, todayMonth);
  if (JSON.stringify(covered) !== JSON.stringify(state.leave)) {
    state.leave = covered;
    persist({ sync: true });
  }
  const timeline = buildTimeline(state.leave);
  const summary = summarizeLeave(timeline, todayMonth);
  setLeaveAmount("leaveKpiAvailable", summary.available, summary.available < 0);
  setLeaveAmount("leaveKpiUsed", summary.usedMonth, false);
  setLeaveAmount("leaveKpiYearEnd", summary.remaining, summary.remaining != null && summary.remaining < 0);
  setLeaveAmount("leaveKpiCarry", summary.carryOut, false);
  const note = $("leaveNote");
  if (note && document.activeElement !== note) note.placeholder = t("leave.notePh");
  const settings = state.leave.settings;
  syncLeaveSettingsLock();
  if (!ui.leaveHoldSettings) {
    const startInput = $("leaveStartMonth");
    if (startInput && document.activeElement !== startInput && startInput.value !== settings.startMonth) startInput.value = settings.startMonth;
    setNumber("leaveInitialBalance", formatLeaveAmount(settings.initialBalance));
    setNumber("leaveInitialCarry", formatLeaveAmount(settings.initialCarryOver));
    setNumber("leaveMonthlyAccrual", formatLeaveAmount(settings.monthlyAccrual));
    setNumber("leaveHorizon", settings.horizonMonths);
  }

  const years = [...new Set(timeline.months.map((row) => row.month.slice(0, 4)))];
  const year = selectedLeaveYear(years, todayMonth);
  const yearSelect = $("leaveYearSelect");
  const yearSig = `${years.join(",")}|${year}|${lang()}`;
  if (yearSelect && document.activeElement !== yearSelect && ui.leaveYearSig !== yearSig) {
    ui.leaveYearSig = yearSig;
    yearSelect.innerHTML = `<option value="all">${escapeHtml(t("leave.allYears"))}</option>${years.map((item) => `<option value="${item}">${item}</option>`).join("")}`;
    yearSelect.value = year;
  }
  const statusSelect = $("leaveStatusSelect");
  if (statusSelect && document.activeElement !== statusSelect && statusSelect.value !== ui.leaveStatus) statusSelect.value = ui.leaveStatus;

  const deficits = timeline.months.filter((row) => row.deficit);
  const alert = $("leaveDeficitAlert");
  if (alert) {
    alert.classList.toggle("hidden", deficits.length === 0);
    if (deficits.length) {
      alert.textContent = t("leave.deficitAlert", {
        month: formatMonthTitle(deficits[0].month, lang()),
        count: String(deficits.length),
      });
    }
  }

  const status = ui.leaveStatus || "all";
  const sig = JSON.stringify({ leave: state.leave, year, status, language: lang(), todayMonth });
  if (sig === ui.leaveSig) return;
  ui.leaveSig = sig;
  const orphans = entriesBeforeStart(state.leave);
  const orphanBox = $("leaveOrphanBox");
  const orphanList = $("leaveOrphanList");
  if (orphanBox && orphanList) {
    orphanBox.classList.toggle("hidden", orphans.length === 0);
    orphanList.innerHTML = orphans.map(leaveChipHtml).join("");
  }
  const visible = timeline.months.filter((row) => {
    if (year !== "all" && !row.month.startsWith(year)) return false;
    const kind = row.month === todayMonth ? "current" : row.month < todayMonth ? "past" : "upcoming";
    if ((status === "past" || status === "current" || status === "upcoming") && kind !== status) return false;
    if (status === "used" && !(row.used > 0 || row.dates.length)) return false;
    return true;
  });
  const body = $("leaveTableBody");
  if (!visible.length) {
    body.innerHTML = `<tr><td colspan="10">${escapeHtml(t("leave.empty"))}</td></tr>`;
    return;
  }
  body.innerHTML = visible.map((row) => {
    const kind = row.month === todayMonth ? "current" : row.month < todayMonth ? "past" : "upcoming";
    const usedCell = row.mode === "manual"
      ? `<input type="number" class="num-input ltr-num leave-used-input" min="0" step="0.25" data-leave-used="${row.month}" value="${formatLeaveAmount(row.used)}">`
      : `<span class="ltr-num">${formatLeaveAmount(row.used)}</span>`;
    return `<tr class="${kind === "current" ? "is-current" : ""}">
      <td><div class="leave-month-cell"><strong>${escapeHtml(formatMonthTitle(row.month, lang()))}</strong>${kind === "current" ? `<span class="badge-chip">${escapeHtml(t("leave.currentBadge"))}</span>` : ""}</div></td>
      <td><select class="select-input" data-leave-mode="${row.month}"><option value="auto"${row.mode === "auto" ? " selected" : ""}>${escapeHtml(t("leave.modeAuto"))}</option><option value="manual"${row.mode === "manual" ? " selected" : ""}>${escapeHtml(t("leave.modeManual"))}</option></select></td>
      <td class="ltr-num">${formatLeaveAmount(row.accrual)}</td>
      <td class="ltr-num">${formatLeaveAmount(row.carryIn)}</td>
      <td class="ltr-num">${formatLeaveAmount(row.cumulative)}</td>
      <td class="ltr-num">${formatLeaveAmount(row.added)}</td>
      <td class="ltr-num">${formatLeaveAmount(row.carryOut)}</td>
      <td>${usedCell}</td>
      <td class="${row.deficit ? "leave-balance-bad ltr-num" : "ltr-num"}">${formatLeaveAmount(row.finalBalance)}</td>
      <td class="leave-dates">${row.dates.map(leaveChipHtml).join("")}</td>
    </tr>`;
  }).join("");
}

const LEAVE_SETTING_IDS = ["leaveStartMonth", "leaveInitialBalance", "leaveInitialCarry", "leaveMonthlyAccrual", "leaveHorizon"];

function syncLeaveSettingsLock() {
  const locked = !ui.leaveEditing;
  LEAVE_SETTING_IDS.forEach((id) => {
    const field = $(id);
    if (field) field.disabled = locked;
  });
  $("leaveEditBtn")?.classList.toggle("hidden", ui.leaveEditing);
  $("leaveSaveBtn")?.classList.toggle("hidden", !ui.leaveEditing);
  $("leaveCancelBtn")?.classList.toggle("hidden", !ui.leaveEditing);
}

function startLeaveSettingsEdit() {
  ui.leaveEditing = true;
  syncLeaveSettingsLock();
  $("leaveStartMonth")?.focus();
}

function cancelLeaveSettingsEdit() {
  ui.leaveEditing = false;
  ui.leaveHoldSettings = false;
  paint();
}

function leaveSettingsDraft() {
  return {
    startMonth: $("leaveStartMonth").value,
    initialBalance: $("leaveInitialBalance").value,
    initialCarryOver: $("leaveInitialCarry").value,
    monthlyAccrual: $("leaveMonthlyAccrual").value,
    horizonMonths: $("leaveHorizon").value,
  };
}

function leaveConfirmRows(items) {
  return items.map((item) => `<li><span>${escapeHtml(item.label)}</span><strong class="ltr-num${item.negative ? " is-negative" : ""}">${escapeHtml(item.value)}</strong></li>`).join("");
}

function requestLeaveSettingsSave() {
  const todayMonth = isoDateLocal(new Date()).slice(0, 7);
  const next = coverLeaveHorizon(updateLeaveSettings(state.leave, leaveSettingsDraft()), todayMonth);
  if (JSON.stringify(next.settings) === JSON.stringify(state.leave.settings)) {
    ui.leaveEditing = false;
    ui.leaveHoldSettings = false;
    paint();
    showFeedback(t("leave.saveSame"), "warn");
    return;
  }
  pendingLeaveSettings = next;
  const settings = next.settings;
  $("leaveSettingsSummary").innerHTML = leaveConfirmRows([
    { label: t("leave.startMonth"), value: settings.startMonth },
    { label: t("leave.initialBalance"), value: formatLeaveAmount(settings.initialBalance) },
    { label: t("leave.initialCarry"), value: formatLeaveAmount(settings.initialCarryOver) },
    { label: t("leave.accrual"), value: formatLeaveAmount(settings.monthlyAccrual) },
    { label: t("leave.horizon"), value: String(settings.horizonMonths) },
  ]);
  $("leaveSettingsModal").classList.remove("hidden");
  syncPageScroll();
}

function closeLeaveSettingsConfirm() {
  $("leaveSettingsModal").classList.add("hidden");
  pendingLeaveSettings = null;
  syncPageScroll();
}

function confirmLeaveSettings() {
  if (!pendingLeaveSettings) return;
  state.leave = pendingLeaveSettings;
  pendingLeaveSettings = null;
  $("leaveSettingsModal").classList.add("hidden");
  ui.leaveEditing = false;
  ui.leaveHoldSettings = false;
  ui.leaveSig = "";
  persist({ sync: true });
  paint();
  const summary = summarizeLeave(buildTimeline(state.leave), isoDateLocal(new Date()).slice(0, 7));
  const remaining = summary.remaining == null ? "—" : formatLeaveAmount(summary.remaining);
  $("leaveSavedSummary").innerHTML = leaveConfirmRows([
    { label: t("leave.available"), value: formatLeaveAmount(summary.available), negative: summary.available < 0 },
    { label: t("leave.usedTotal"), value: formatLeaveAmount(summary.usedMonth) },
    { label: t("leave.yearEnd"), value: remaining, negative: summary.remaining != null && summary.remaining < 0 },
    { label: t("leave.carry"), value: formatLeaveAmount(summary.carryOut) },
  ]);
  $("leaveSavedModal").classList.remove("hidden");
  syncPageScroll();
}

function closeLeaveSaved() {
  $("leaveSavedModal").classList.add("hidden");
  syncPageScroll();
}

function addLeaveFromForm(event) {
  event.preventDefault();
  const start = $("leaveStartDate").value;
  const end = ui.leaveSpan === "range" ? ($("leaveEndDate").value || start) : start;
  const result = addLeaveEntries(state.leave, { start, end, note: $("leaveNote").value });
  if (result.error === "invalid-date") {
    showFeedback(t("leave.needDate"), true);
    return;
  }
  if (result.error === "range") {
    showFeedback(t("leave.rangeLimit"), true);
    return;
  }
  state.leave = coverLeaveHorizon(result.leave, isoDateLocal(new Date()).slice(0, 7));
  if (result.added.length) {
    const years = [...new Set(result.added.map((date) => date.slice(0, 4)))];
    ui.leaveYear = years.length > 1 ? "all" : years[0];
    ui.leaveStatus = "used";
    ui.leaveYearSig = "";
  }
  ui.leaveSig = "";
  persist({ sync: true });
  paint();
  if (!result.added.length) showFeedback(t("leave.duplicate"), true);
  else showFeedback(t("leave.added", { count: String(result.added.length) }));
  $("leaveNote").value = "";
}

function removeLeaveDate(date) {
  state.leave = removeLeaveEntry(state.leave, date);
  ui.leaveSig = "";
  persist({ sync: true });
  paint();
}

function bind() {
  document.addEventListener("click", onClick);
  document.querySelectorAll(".view-tab").forEach((btn) => {
    btn.addEventListener("click", () => setView(btn.dataset.view));
  });
  $("leaveView").addEventListener("click", (event) => {
    const span = event.target.closest("[data-leave-span]");
    if (span) {
      setLeaveSpan(span.dataset.leaveSpan);
      return;
    }
    const remove = event.target.closest("[data-leave-remove]");
    if (remove) openDeleteConfirm(remove.dataset.leaveRemove, "leave");
  });
  $("leaveView").addEventListener("change", (event) => {
    if (event.target.id === "leaveYearSelect") {
      ui.leaveYear = event.target.value;
      ui.leaveSig = "";
      paint();
      return;
    }
    if (event.target.id === "leaveStatusSelect") {
      ui.leaveStatus = event.target.value;
      ui.leaveSig = "";
      paint();
      return;
    }
    if (LEAVE_SETTING_IDS.includes(event.target.id)) {
      ui.leaveHoldSettings = true;
      return;
    }
    const mode = event.target.closest("[data-leave-mode]");
    if (mode) {
      state.leave = setLeaveMonthMode(state.leave, mode.dataset.leaveMode, mode.value);
      ui.leaveSig = "";
      persist({ sync: true });
      paint();
      return;
    }
    const used = event.target.closest("[data-leave-used]");
    if (!used) return;
    state.leave = setLeaveMonthUsed(state.leave, used.dataset.leaveUsed, used.value);
    ui.leaveSig = "";
    persist({ sync: true });
    paint();
  });
  $("leaveAddForm").addEventListener("submit", addLeaveFromForm);
  $("leaveEditBtn").addEventListener("click", startLeaveSettingsEdit);
  $("leaveCancelBtn").addEventListener("click", cancelLeaveSettingsEdit);
  $("leaveSaveBtn").addEventListener("click", requestLeaveSettingsSave);
  LEAVE_SETTING_IDS.forEach((id) => {
    $(id).addEventListener("input", () => {
      if (!ui.leaveEditing) return;
      ui.leaveHoldSettings = true;
    });
  });
  $("leaveSettingsRejectBtn").addEventListener("click", closeLeaveSettingsConfirm);
  $("leaveSettingsAcceptBtn").addEventListener("click", confirmLeaveSettings);
  $("leaveSettingsModal").addEventListener("click", (event) => {
    if (event.target === $("leaveSettingsModal")) closeLeaveSettingsConfirm();
  });
  $("leaveLinkRejectBtn").addEventListener("click", rejectLeaveLink);
  $("leaveLinkAcceptBtn").addEventListener("click", confirmLeaveLink);
  $("leaveLinkModal").addEventListener("click", (event) => {
    if (event.target === $("leaveLinkModal")) rejectLeaveLink();
  });
  $("leaveSavedOkBtn").addEventListener("click", closeLeaveSaved);
  $("leaveSavedModal").addEventListener("click", (event) => {
    if (event.target === $("leaveSavedModal")) closeLeaveSaved();
  });
  let pendingLeaveAction = null;
  function openLeaveAction(kind) {
    pendingLeaveAction = kind === "extend" ? "extend" : "reset";
    const extend = pendingLeaveAction === "extend";
    $("leaveActionTitle").textContent = t(extend ? "leave.extendTitle" : "leave.resetTitle");
    $("leaveActionText").textContent = t(extend ? "leave.extendConfirm" : "leave.resetConfirm");
    $("leaveActionAcceptBtn").textContent = t(extend ? "leave.extend" : "leave.reset");
    $("leaveActionModal").classList.remove("hidden");
    syncPageScroll();
  }
  function closeLeaveAction() {
    $("leaveActionModal").classList.add("hidden");
    pendingLeaveAction = null;
    syncPageScroll();
  }
  function confirmLeaveAction() {
    const kind = pendingLeaveAction;
    if (!kind) return;
    closeLeaveAction();
    const todayMonth = isoDateLocal(new Date()).slice(0, 7);
    state.leave = kind === "extend"
      ? expandLeaveHorizon(state.leave, 12)
      : coverLeaveHorizon(resetLeaveSettings(state.leave), todayMonth);
    ui.leaveEditing = false;
    ui.leaveHoldSettings = false;
    ui.leaveSig = "";
    persist({ sync: true });
    paint();
    showFeedback(t(kind === "extend" ? "leave.extended" : "leave.resetOk"));
  }
  $("leaveResetBtn").addEventListener("click", () => openLeaveAction("reset"));
  $("leaveExtendBtn").addEventListener("click", () => openLeaveAction("extend"));
  $("leaveActionRejectBtn").addEventListener("click", closeLeaveAction);
  $("leaveActionAcceptBtn").addEventListener("click", confirmLeaveAction);
  $("leaveActionModal").addEventListener("click", (event) => {
    if (event.target === $("leaveActionModal")) closeLeaveAction();
  });
  $("feedbackDismissBtn").addEventListener("click", () => {
    hideFeedback();
    stopTabBlink();
  });
  $("pickerCancelBtn").addEventListener("click", closePicker);
  $("pickerConfirmBtn").addEventListener("click", confirmPicker);
  $("drumPickerModal").addEventListener("click", (event) => {
    if (event.target === $("drumPickerModal")) closePicker();
  });
  document.querySelectorAll(".theme-toggle").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.settings.theme = state.settings.theme === "dark" ? "light" : "dark";
      applyTheme();
      persist({ sync: true });
    });
  });
  document.querySelectorAll(".sound-toggle").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.settings.soundEnabled = !state.settings.soundEnabled;
      applyTheme();
      persist({ sync: true });
    });
  });
  document.querySelectorAll(".lang-toggle").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.settings.lang = lang() === "en" ? "ar" : "en";
      ui.logSig = "";
      applyLang();
      persist({ sync: true });
      paint();
    });
  });
  $("checkInNowBtn").addEventListener("click", () => {
    if (!modelNow().isWorkDay) {
      state.today.statusOverride = STATUS.work;
      state.today.workplace = "office";
    }
    state.today.checkIn = hhmmNow(new Date());
    rememberToday();
    persist({ sync: true });
    paint();
  });
  $("checkOutNowBtn").addEventListener("click", () => {
    if (!state.today.checkIn) {
      showFeedback(lang() === "en" ? "Check in first." : "سجّل الحضور أولًا.", true);
      return;
    }
    state.today.checkOut = hhmmNow(new Date());
    rememberToday();
    persist({ sync: true });
    paint();
  });
  const syncFromCloud = () => {
    if (!state.settings.sheetsWebhook) {
      showFeedback(t("toast.syncNeed"), true);
      return;
    }
    if (state.settings.sheetsImport === "keep") {
      openImportChoice(state.settings.sheetsWebhook);
      return;
    }
    void pullSync({ announce: true });
  };
  $("syncSheetsBtn").addEventListener("click", syncFromCloud);
  $("leaveSyncBtn").addEventListener("click", syncFromCloud);
  $("testConnectSheetsBtn").addEventListener("click", () => { commitWebhook(); });
  $("editWebhookBtn").addEventListener("click", openEditWebhook);
  $("editWebhookRejectBtn").addEventListener("click", closeEditWebhook);
  $("editWebhookAcceptBtn").addEventListener("click", confirmEditWebhook);
  $("editWebhookModal").addEventListener("click", (event) => {
    if (event.target === $("editWebhookModal")) closeEditWebhook();
  });
  $("importKeepBtn").addEventListener("click", () => { void chooseSheetImport(false); });
  $("importReplaceBtn").addEventListener("click", () => { void chooseSheetImport(true); });
  $("importChoiceModal").addEventListener("click", (event) => {
    if (event.target === $("importChoiceModal")) cancelImportChoice();
  });
  $("toggleGasGuideBtn").addEventListener("click", () => {
    $("gasGuideModal").classList.remove("hidden");
    syncPageScroll();
  });
  $("guideCloseBtn").addEventListener("click", closeGuide);
  $("copyGasScriptBtn").addEventListener("click", () => { void copyGasScript(); });
  $("deleteRejectBtn").addEventListener("click", closeDeleteConfirm);
  $("deleteAcceptBtn").addEventListener("click", confirmDelete);
  $("deleteConfirmModal").addEventListener("click", (event) => {
    if (event.target === $("deleteConfirmModal")) closeDeleteConfirm();
  });
  $("logEditCancel").addEventListener("click", closeLogEdit);
  $("logEditSave").addEventListener("click", saveLogEdit);
  $("logMonthSelect").addEventListener("change", (event) => {
    ui.logMonth = event.target.value;
    ui.logSig = "";
    paint();
  });
  $("overtimeReportBtn").addEventListener("click", openOvertimeReport);
  $("overtimeCloseBtn").addEventListener("click", closeOvertimeReport);
  $("overtimeCopyBtn").addEventListener("click", () => { void copyOvertimeReport(); });
  $("overtimeModal").addEventListener("click", (event) => {
    if (event.target === $("overtimeModal")) closeOvertimeReport();
    const remove = event.target.closest("[data-ot-remove]");
    if (!remove || !overtimeDraft) return;
    overtimeDraft.days = overtimeDraft.days.filter((day) => day.date !== remove.dataset.otRemove);
    paintOvertimeDraft();
  });
  $("overtimeReportBody").addEventListener("input", (event) => {
    const input = event.target.closest("[data-ot-minutes]");
    if (!input || !overtimeDraft) return;
    const minutes = minutesFromDuration(input.value);
    if (minutes == null) return;
    const day = overtimeDraft.days.find((item) => item.date === input.dataset.otMinutes);
    if (!day) return;
    day.minutes = minutes;
    syncOvertimeSummary();
  });
  $("overtimeReportBody").addEventListener("focusout", (event) => {
    const input = event.target.closest("[data-ot-minutes]");
    if (!input || !overtimeDraft) return;
    const day = overtimeDraft.days.find((item) => item.date === input.dataset.otMinutes);
    if (!day) return;
    input.value = formatHourMinute(day.minutes);
  });
  $("overtimeAddForm").addEventListener("submit", addOvertimeDraftDay);
  $("personalReportBtn").addEventListener("click", openPersonalReport);
  $("personalCloseBtn").addEventListener("click", closePersonalReport);
  $("personalCopyBtn").addEventListener("click", () => { void copyPersonalReport(); });
  $("personalModal").addEventListener("click", (event) => {
    if (event.target === $("personalModal")) closePersonalReport();
    const remove = event.target.closest("[data-personal-remove]");
    if (!remove || !personalDraft) return;
    personalDraft.days = personalDraft.days.filter((day) => day.date !== remove.dataset.personalRemove);
    paintPersonalDraft();
  });
  $("personalAddForm").addEventListener("submit", addPersonalDraftDay);
  $("leaveMonthReportBtn").addEventListener("click", openLeaveMonthReport);
  $("leaveMonthCloseBtn").addEventListener("click", closeLeaveMonthReport);
  $("leaveMonthCopyBtn").addEventListener("click", () => { void copyReportText("leaveMonthCopySource"); });
  $("leaveMonthModal").addEventListener("click", (event) => {
    if (event.target === $("leaveMonthModal")) closeLeaveMonthReport();
    const remove = event.target.closest("[data-leave-month-remove]");
    if (!remove || !leaveMonthDraft) return;
    leaveMonthDraft.days = leaveMonthDraft.days.filter((day) => day.date !== remove.dataset.leaveMonthRemove);
    paintLeaveMonthDraft();
  });
  $("leaveMonthReportBody").addEventListener("input", (event) => {
    const input = event.target.closest("[data-leave-month-note]");
    if (!input || !leaveMonthDraft) return;
    const day = leaveMonthDraft.days.find((item) => item.date === input.dataset.leaveMonthNote);
    if (!day) return;
    day.note = input.value.slice(0, 500);
    syncLeaveMonthSummary();
  });
  $("leaveMonthAddForm").addEventListener("submit", addLeaveMonthDraftDay);
  $("leaveYearReportBtn").addEventListener("click", openLeaveYearReport);
  $("leaveYearCloseBtn").addEventListener("click", closeLeaveYearReport);
  $("leaveYearCopyBtn").addEventListener("click", () => { void copyReportText("leaveYearCopySource"); });
  $("leaveYearModal").addEventListener("click", (event) => {
    if (event.target === $("leaveYearModal")) closeLeaveYearReport();
    if (!leaveYearDraft) return;
    const dayRemove = event.target.closest("[data-leave-year-day]");
    if (dayRemove) {
      const date = dayRemove.dataset.leaveYearDay;
      leaveYearDraft.months.forEach((row) => {
        row.days = row.days.filter((day) => day.date !== date);
      });
      leaveYearDraft.months = leaveYearDraft.months.filter((row) => row.days.length);
      paintLeaveYearDraft();
      return;
    }
    const remove = event.target.closest("[data-leave-year-remove]");
    if (!remove) return;
    leaveYearDraft.months = leaveYearDraft.months.filter((row) => row.month !== remove.dataset.leaveYearRemove);
    paintLeaveYearDraft();
  });
  $("leaveYearAddForm").addEventListener("submit", addLeaveYearDraftDay);
  $("logEditStatus").addEventListener("change", (event) => { logEdit.status = event.target.value; });
  $("logEditModal").addEventListener("click", (event) => {
    if (event.target === $("logEditModal")) closeLogEdit();
  });
  $("gasGuideModal").addEventListener("click", (event) => {
    if (event.target === $("gasGuideModal")) closeGuide();
  });
  $("clearPermTimeBtn").addEventListener("click", () => {
    state.today.permissionTime = null;
    persist({ sync: true });
    paint();
  });
  $("dhuhrBreakEnabled").addEventListener("change", (event) => {
    state.today.dhuhrEnabled = event.target.checked;
    persist({ sync: true });
    paint();
  });
  $("asrBreakEnabled").addEventListener("change", (event) => {
    state.today.asrEnabled = event.target.checked;
    persist({ sync: true });
    paint();
  });
  $("dhuhrAutoEnable").addEventListener("change", (event) => {
    state.settings.dhuhr.enabled = event.target.checked;
    persist({ sync: true });
    paint();
  });
  $("asrAutoEnable").addEventListener("change", (event) => {
    state.settings.asr.enabled = event.target.checked;
    persist({ sync: true });
    paint();
  });
  $("progressShimmerToggle").addEventListener("change", (event) => {
    state.settings.progressShimmer = event.target.checked;
    persist({ sync: true });
    paint();
  });
  $("showHeaderTargetToggle").addEventListener("change", (event) => {
    state.settings.showHeaderTarget = event.target.checked;
    persist({ sync: true });
    paint();
  });
  $("manualBreakInput").addEventListener("input", (event) => {
    setManual(Number(event.target.value) || 0);
    persist({ sync: true });
    paint();
  });
  $("permTypeSelect").addEventListener("change", (event) => {
    armPermission(state.today.permissionMinutes, UI_TO_TYPE[event.target.value] || "auto");
    persist({ sync: true });
    paint();
  });
  $("dayStatusSelect").addEventListener("change", (event) => {
    const before = modelNow().status;
    const key = event.target.value;
    const after = MODE_TO_STATUS[key] || STATUS.work;
    state.today.statusOverride = after;
    state.today.workplace = key === "remote" ? "remote" : "office";
    persist({ sync: true });
    paint();
    offerLeaveLink(state.today.date, before, after);
  });
  const pickerSwitch = $("pickerStyleSwitch");
  const setPickerStyle = (value) => {
    state.settings.timePicker = value === "scroll" ? "scroll" : "clock";
    persist({ sync: true });
    paint();
  };
  let pickerSwipe = null;
  let ignorePickerClick = false;
  pickerSwitch.addEventListener("pointerdown", (event) => {
    pickerSwipe = { x: event.clientX, id: event.pointerId, moved: false };
  });
  pickerSwitch.addEventListener("pointermove", (event) => {
    if (!pickerSwipe || event.pointerId !== pickerSwipe.id) return;
    if (Math.abs(event.clientX - pickerSwipe.x) > 10) pickerSwipe.moved = true;
  });
  pickerSwitch.addEventListener("pointerup", (event) => {
    if (!pickerSwipe || event.pointerId !== pickerSwipe.id) return;
    const dx = event.clientX - pickerSwipe.x;
    const moved = pickerSwipe.moved;
    pickerSwipe = null;
    if (!moved || Math.abs(dx) < 28) return;
    ignorePickerClick = true;
    const rtl = document.documentElement.dir === "rtl";
    const toScroll = rtl ? dx > 0 : dx < 0;
    setPickerStyle(toScroll ? "scroll" : "clock");
  });
  pickerSwitch.addEventListener("click", (event) => {
    if (ignorePickerClick) {
      ignorePickerClick = false;
      return;
    }
    const btn = event.target.closest("[data-picker-style]");
    if (!btn) return;
    setPickerStyle(btn.dataset.pickerStyle);
  });
  $("clockHourPart").addEventListener("click", () => { picker.hand = "hour"; syncPickerChrome(); });
  $("clockMinutePart").addEventListener("click", () => { picker.hand = "minute"; syncPickerChrome(); });
  const face = $("clockFace");
  const applyDial = (event) => {
    if (picker.entry === "type" || state.settings.timePicker === "scroll") return false;
    if (picker.hand === "hour") picker.hour = hourFromPointer(event);
    else picker.minute = minuteFromPointer(event);
    return true;
  };
  face.addEventListener("selectstart", (event) => event.preventDefault());
  face.addEventListener("pointerdown", (event) => {
    if (!applyDial(event)) return;
    event.preventDefault();
    face.setPointerCapture(event.pointerId);
    syncPickerChrome();
  });
  face.addEventListener("pointermove", (event) => {
    if (!face.hasPointerCapture(event.pointerId)) return;
    const before = picker.hand === "hour" ? picker.hour : picker.minute;
    if (!applyDial(event)) return;
    const after = picker.hand === "hour" ? picker.hour : picker.minute;
    if (before === after) return;
    syncPickerChrome();
  });
  face.addEventListener("pointerup", (event) => {
    if (picker.hand !== "hour" || picker.entry === "type" || state.settings.timePicker === "scroll") return;
    if (!face.hasPointerCapture(event.pointerId) && event.target !== face && !face.contains(event.target)) return;
    picker.hand = "minute";
    syncPickerChrome();
  });
  const writeBreakDuration = (key, value) => {
    const number = Number(value);
    if (!Number.isFinite(number)) return;
    state.settings[key].duration = clampInt(number, 1, 180);
    persist({ sync: true });
    paint();
    pingAlerts();
  };
  $("dhuhrDurationInput").addEventListener("change", (event) => writeBreakDuration("dhuhr", event.target.value));
  $("asrDurationInput").addEventListener("change", (event) => writeBreakDuration("asr", event.target.value));
  $("clockDialBtn").addEventListener("click", () => { picker.entry = "dial"; syncPickerChrome(); });
  $("clockTypeBtn").addEventListener("click", () => { picker.entry = "type"; syncPickerChrome(); });
  document.querySelectorAll("[data-clock-period]").forEach((btn) => {
    btn.addEventListener("click", () => {
      picker.pm = btn.dataset.clockPeriod === "pm";
      previewPicker();
    });
  });
  const clampClockField = (input, kind) => {
    input.addEventListener("input", () => {
      const raw = input.value.replace(/\D/g, "").slice(0, 2);
      input.value = raw;
      if (raw === "") return;
      const number = Number(raw);
      if (kind === "hour") {
        const max = picker.duration ? 12 : 12;
        const min = picker.duration ? 0 : 1;
        picker.hour = Math.min(max, Math.max(min, number));
      } else {
        picker.minute = Math.min(59, Math.max(0, number));
      }
      previewPicker();
    });
  };
  clampClockField($("clockHourInput"), "hour");
  clampClockField($("clockMinuteInput"), "minute");
  $("targetHoursInput").addEventListener("change", (event) => {
    const hours = Math.min(16, Math.max(0.5, Number(event.target.value) || 8));
    state.settings.dailyTargetMinutes = Math.round(hours * 60);
    state.today.targetMinutes = null;
    persist({ sync: true });
    paint();
  });
  $("alertSecondsInput").addEventListener("change", (event) => {
    state.settings.alertSeconds = Math.min(60, Math.max(1, Math.round(Number(event.target.value) || 5)));
    persist({ sync: true });
    paint();
  });
  $("webhookUrlInput").addEventListener("keydown", (event) => {
    if (event.key !== "Enter" || event.target.readOnly) return;
    event.preventDefault();
    commitWebhook();
  });
  const wires = [
    ["alertPreBreakEnabled", "change", (v) => { state.settings.notifications.preBreak.enabled = v; }],
    ["preBreakMinutesInput", "change", (v) => { state.settings.notifications.preBreak.minutesBefore = clampInt(v, 0, 180); }],
    ["blinkPreBreakToggle", "change", (v) => { state.settings.notifications.preBreak.blink = v; }],
    ["alertBreakEndEnabled", "change", (v) => { state.settings.notifications.breakEnd.enabled = v; }],
    ["blinkBreakEndToggle", "change", (v) => { state.settings.notifications.breakEnd.blink = v; }],
    ["alertCheckoutEnabled", "change", (v) => { state.settings.notifications.shiftEnd.enabled = v; }],
    ["preCheckoutMinutesInput", "change", (v) => { state.settings.notifications.shiftEnd.minutesBefore = clampInt(v, 0, 180); }],
    ["blinkCheckoutToggle", "change", (v) => { state.settings.notifications.shiftEnd.blink = v; }],
    ["alertPermissionEnabled", "change", (v) => { state.settings.notifications.permission.enabled = v; }],
    ["blinkPermissionToggle", "change", (v) => { state.settings.notifications.permission.blink = v; }],
  ];
  wires.forEach(([id, type, write]) => {
    $(id).addEventListener(type, (event) => {
      const value = event.target.type === "checkbox" ? event.target.checked : Number(event.target.value);
      write(value);
      persist({ sync: true });
    });
  });
  $("requestBrowserNotifBtn").addEventListener("click", async () => {
    if (!("Notification" in window)) return;
    const result = await Notification.requestPermission();
    showFeedback(result === "granted" ? t("toast.connected") : t("toast.connectFail"), result !== "granted");
  });
  const closeExportMenu = () => {
    $("exportMenu").classList.add("hidden");
    $("exportMenuBtn").setAttribute("aria-expanded", "false");
  };
  $("exportMenuBtn").addEventListener("click", (event) => {
    event.stopPropagation();
    const menu = $("exportMenu");
    const willOpen = menu.classList.contains("hidden");
    menu.classList.toggle("hidden", !willOpen);
    $("exportMenuBtn").setAttribute("aria-expanded", willOpen ? "true" : "false");
  });
  const closeLeaveExportMenu = () => {
    $("leaveExportMenu").classList.add("hidden");
    $("leaveExportMenuBtn").setAttribute("aria-expanded", "false");
  };
  $("leaveExportMenuBtn").addEventListener("click", (event) => {
    event.stopPropagation();
    const menu = $("leaveExportMenu");
    const willOpen = menu.classList.contains("hidden");
    menu.classList.toggle("hidden", !willOpen);
    $("leaveExportMenuBtn").setAttribute("aria-expanded", willOpen ? "true" : "false");
  });
  document.addEventListener("click", (event) => {
    if (!$("exportMenu").classList.contains("hidden") && !$("exportMenu").contains(event.target)) closeExportMenu();
    if ($("leaveExportMenu").classList.contains("hidden")) return;
    if ($("leaveExportMenu").contains(event.target)) return;
    closeLeaveExportMenu();
  });
  const exportBackup = (kind) => {
    closeExportMenu();
    try {
      state.log = upsertLog(state.log, toLogEntry(state.today, modelNow()));
      ui.logSig = "";
      persist();
      if (kind === "md") downloadMarkdown(state);
      else downloadWorkbook(state);
      showFeedback(t("toast.exportOk"));
      paint();
    } catch {
      showFeedback(t("toast.exportFail"), true);
    }
  };
  const exportLeaveBackup = (kind) => {
    closeLeaveExportMenu();
    try {
      if (kind === "md") downloadLeaveMarkdown(state);
      else downloadLeaveWorkbook(state);
      showFeedback(t("toast.exportOk"));
    } catch {
      showFeedback(t("toast.exportFail"), true);
    }
  };
  $("exportExcelBtn").addEventListener("click", () => exportBackup("xlsx"));
  $("exportMdBtn").addEventListener("click", () => exportBackup("md"));
  $("leaveExportExcelBtn").addEventListener("click", () => exportLeaveBackup("xlsx"));
  $("leaveExportMdBtn").addEventListener("click", () => exportLeaveBackup("md"));
  $("importExcelBtn").addEventListener("click", () => $("excelFileInput").click());
  $("leaveImportBtn").addEventListener("click", () => $("leaveFileInput").click());
  $("leaveFileInput").addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    const parsed = await readLeaveFile(file);
    if (!parsed.ok) {
      showFeedback(t("leave.importEmpty"), true);
      return;
    }
    state.leave = coverLeaveHorizon(parsed.leave, isoDateLocal(new Date()).slice(0, 7));
    ui.leaveEditing = false;
    ui.leaveHoldSettings = false;
    ui.leaveSig = "";
    persist({ sync: true });
    paint();
    showFeedback(t("leave.importOk"));
  });
  $("excelFileInput").addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    const parsed = await readBackupFile(file);
    if (!parsed.ok) {
      showFeedback(t("toast.importFail"), true);
      return;
    }
    const result = applyImport(state, parsed.payload);
    if (!result.ok) {
      showFeedback(result.error || t("toast.importFail"), true);
      return;
    }
    state = result.state;
    state.settings.fileSnapshot = true;
    ui.pullReady = false;
    ui.logSig = "";
    applyLang();
    persist();
    paint();
    showFeedback(t("toast.importOk"));
    if (state.settings.sheetsWebhook) void confirmSheetOnOpen(false);
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      closePicker();
      closeGuide();
      closeOvertimeReport();
      closePersonalReport();
      closeLeaveMonthReport();
      closeLeaveYearReport();
      closeDeleteConfirm();
      closeLeaveSettingsConfirm();
      closeLeaveSaved();
      closeLeaveLink();
      closeLeaveAction();
      $("leaveExportMenu").classList.add("hidden");
      $("leaveExportMenuBtn").setAttribute("aria-expanded", "false");
      closeEditWebhook();
      cancelImportChoice();
      $("exportMenu").classList.add("hidden");
      $("exportMenuBtn").setAttribute("aria-expanded", "false");
    }
  });
}

function boot() {
  const landed = settleToday(state, isoDateLocal(new Date()));
  const hasLink = Boolean(landed.state.settings.sheetsWebhook);
  const keepLocal = landed.state.settings.sheetsImport === "keep";
  const fileSnapshot = landed.state.settings.fileSnapshot === true;
  const shouldPull = hasLink && !keepLocal && !fileSnapshot;
  state = shouldPull ? landed.state : fillMissingDays(landed.state);
  state.settings.accent = normalizeAccent(state.settings.accent);
  if (shouldPull) ui.pulling = true;
  $("gasScript").textContent = APPS_SCRIPT;
  applyLang();
  persist();
  paint();
  if (storedForeign) showFeedback(t("toast.notOurs"), true);
  else if (shouldPull) showFeedback(t("toast.sheetPulling"), "warn");
  else if (linkIsBehindScript(state.settings)) showFeedback(t("toast.scriptUpdate"), true);
  else if (landed.archived) showFeedback(t("toast.rollover"));
  bind();
  watchFocusStop();
  document.addEventListener("pointerdown", () => warmAudio(), { once: true });
  pingAlerts();
  if (hasLink && !storedForeign) void confirmSheetOnOpen(shouldPull);
  window.setInterval(() => {
    const now = new Date();
    const landedNow = settleToday(state, isoDateLocal(now));
    if (landedNow.archived || landedNow.state !== state) {
      state = fillMissingDays(landedNow.state);
      ui.logSig = "";
      if (landedNow.archived) showFeedback(t("toast.rollover"));
    }
    paint(now);
    const minute = now.getHours() * 60 + now.getMinutes();
    if (minute === ui.lastMinute) return;
    ui.lastMinute = minute;
    state.today.lastSeenAt = now.toISOString();
    rememberToday();
    persist({ sync: true });
    runAlerts(minute);
  }, 1000);
}

boot();
