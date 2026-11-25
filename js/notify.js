let baseTitle = "ساعات العمل";
let audioCtx = null;
let blinkTimer = null;

export function setBaseTitle(title) {
  if (title) baseTitle = title;
  if (!blinkTimer) document.title = baseTitle;
}

export function warmAudio() {
  const AudioCtx = window.AudioContext || window.webkitAudioContext;
  if (!AudioCtx) return null;
  if (!audioCtx) audioCtx = new AudioCtx();
  if (audioCtx.state === "suspended") audioCtx.resume();
  return audioCtx;
}

export function playChime() {
  const ctx = warmAudio();
  if (!ctx) return;
  const now = ctx.currentTime;
  [523.25, 659.25, 783.99].forEach((freq, index) => {
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = "sine";
    osc.frequency.value = freq;
    const start = now + index * 0.11;
    gain.gain.setValueAtTime(0.0001, start);
    gain.gain.exponentialRampToValueAtTime(0.08, start + 0.03);
    gain.gain.exponentialRampToValueAtTime(0.0001, start + 0.55);
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.start(start);
    osc.stop(start + 0.6);
  });
}

export function startTabBlink(message) {
  stopTabBlink();
  let on = false;
  blinkTimer = window.setInterval(() => {
    on = !on;
    document.title = on ? `● ${message}` : baseTitle;
  }, 900);
}

export function stopTabBlink() {
  if (blinkTimer) window.clearInterval(blinkTimer);
  blinkTimer = null;
  document.title = baseTitle;
}

export function watchFocusStop() {
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) stopTabBlink();
  });
  window.addEventListener("focus", () => stopTabBlink());
}
