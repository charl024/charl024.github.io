'use strict';

const STORAGE_KEY = 'wheel-of-anime:v1';
const MIN_DURATION = 1;
const MAX_DURATION = 30;
const MERGE = '__all__';
const TAU = Math.PI * 2;
const LABEL_FONT = '"Comic Relief", "Comic Sans MS", cursive';
const NAME_HEADER = /^(names?|titles?|shows?|series)$/i;
const LINK_HEADER = /\b(link|url|mal)\b/i;
const URL_RE = /^https?:\/\/\S+$/i;

const $ = (id) => document.getElementById(id);
const el = {
  wheelWrap: $('wheelWrap'), wheel: $('wheel'), emptyHint: $('emptyHint'), ticker: $('ticker'),
  spinBtn: $('spinBtn'), muteBtn: $('muteBtn'), zoom: $('zoom'), zoomOut: $('zoomOut'),
  pasteBox: $('pasteBox'), pasteText: $('pasteText'), pasteBtn: $('pasteBtn'),
  drop: $('drop'), file: $('file'), fileName: $('fileName'), sheet: $('sheet'), dedupe: $('dedupe'),
  shuffleBtn: $('shuffleBtn'), resetBtn: $('resetBtn'), resetAllBtn: $('resetAllBtn'), count: $('count'), itemList: $('itemList'),
  durationRange: $('durationRange'), duration: $('duration'), seed: $('seed'), seedHint: $('seedHint'),
  autoZoom: $('autoZoom'), history: $('history'), historyEmpty: $('historyEmpty'), clearHistory: $('clearHistory'),
  modal: $('winnerModal'), winnerName: $('winnerName'), winnerLink: $('winnerLink'),
  removeBtn: $('removeBtn'), keepBtn: $('keepBtn'), confetti: $('confetti'),
};
const ctx = el.wheel.getContext('2d');
const reducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

/* ---------------------------------------------------------------- state */

function defaultState() {
  return {
    fileName: '',
    sheets: [],          // [{ name, items: [{ name, url }] }] as parsed from the file
    sheetChoice: '0',    // sheet index as a string, or MERGE
    items: [],           // entries currently on the wheel
    history: [],         // [{ name, url, at, removed }], newest first
    spinCount: 0,        // spins since the seed or list was last set; feeds the seeded RNG
    settings: { duration: 5, seed: '', autoZoom: true, dedupe: true, muted: false },
  };
}

function loadState() {
  const base = defaultState();
  try {
    const saved = JSON.parse(sessionStorage.getItem(STORAGE_KEY));
    if (!saved || typeof saved !== 'object') return base;
    return { ...base, ...saved, settings: { ...base.settings, ...saved.settings } };
  } catch {
    return base;
  }
}

function save() {
  try {
    sessionStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {
    // Storage full or blocked: the app keeps working, it just won't survive a reload.
  }
}

const state = loadState();

/* ---------------------------------------------------------- seeded RNG */

// cyrb128 string hash -> four 32-bit seeds for sfc32.
function cyrb128(str) {
  let h1 = 1779033703, h2 = 3144134277, h3 = 1013904242, h4 = 2773480762;
  for (let i = 0; i < str.length; i++) {
    const k = str.charCodeAt(i);
    h1 = h2 ^ Math.imul(h1 ^ k, 597399067);
    h2 = h3 ^ Math.imul(h2 ^ k, 2869860233);
    h3 = h4 ^ Math.imul(h3 ^ k, 951274213);
    h4 = h1 ^ Math.imul(h4 ^ k, 2716044179);
  }
  h1 = Math.imul(h3 ^ (h1 >>> 18), 597399067);
  h2 = Math.imul(h4 ^ (h2 >>> 22), 2869860233);
  h3 = Math.imul(h1 ^ (h3 >>> 17), 951274213);
  h4 = Math.imul(h2 ^ (h4 >>> 19), 2716044179);
  h1 ^= h2 ^ h3 ^ h4; h2 ^= h1; h3 ^= h1; h4 ^= h1;
  return [h1 >>> 0, h2 >>> 0, h3 >>> 0, h4 >>> 0];
}

function sfc32(a, b, c, d) {
  return function () {
    a |= 0; b |= 0; c |= 0; d |= 0;
    const t = (((a + b) | 0) + d) | 0;
    d = (d + 1) | 0;
    a = b ^ (b >>> 9);
    b = (c + (c << 3)) | 0;
    c = (c << 21) | (c >>> 11);
    c = (c + t) | 0;
    return (t >>> 0) / 4294967296;
  };
}

// Spin k with seed S always draws the same numbers, so the same seed + same list replays the same results.
function makeRng(seed, spinIndex) {
  if (!seed) return Math.random;
  return sfc32(...cyrb128(`${seed}\u0000${spinIndex}`));
}

/* ------------------------------------------------------------- parsing */

const clean = (v) => String(v ?? '').replace(/\s+/g, ' ').trim();
const safeUrl = (v) => (URL_RE.test(clean(v)) ? clean(v) : '');

function extractItems(ws) {
  if (!ws || !ws['!ref']) return [];
  const range = XLSX.utils.decode_range(ws['!ref']);
  const cell = (r, c) => ws[XLSX.utils.encode_cell({ r, c })];
  const text = (r, c) => {
    const x = cell(r, c);
    return x ? clean(XLSX.utils.format_cell(x)) : '';
  };
  const rowTexts = (r) => {
    const out = [];
    for (let c = range.s.c; c <= range.e.c; c++) out.push(text(r, c));
    return out;
  };

  let first = range.s.r;
  while (first <= range.e.r && rowTexts(first).every((t) => !t)) first++;
  if (first > range.e.r) return [];

  // A header row names its columns ("Name", "MAL link"); otherwise the first row is already data.
  const head = rowTexts(first);
  let nameCol = head.findIndex((t) => NAME_HEADER.test(t));
  let linkCol = -1;
  let start = first + 1;
  if (nameCol >= 0) {
    linkCol = head.findIndex((t, i) => i !== nameCol && LINK_HEADER.test(t));
  } else {
    nameCol = head.findIndex((t) => t && !URL_RE.test(t));
    linkCol = head.findIndex((t) => URL_RE.test(t));
    start = first;
    if (nameCol < 0) return [];
  }

  const items = [];
  for (let r = start; r <= range.e.r; r++) {
    const name = text(r, range.s.c + nameCol);
    if (!name || URL_RE.test(name)) continue;
    const linkCell = linkCol >= 0 ? cell(r, range.s.c + linkCol) : null;
    const url =
      safeUrl(linkCol >= 0 ? text(r, range.s.c + linkCol) : '') ||
      safeUrl(linkCell?.l?.Target) ||
      safeUrl(cell(r, range.s.c + nameCol)?.l?.Target);
    items.push({ name, url });
  }
  return items;
}

async function readFile(file) {
  if (!window.XLSX) {
    showFileMessage('The spreadsheet library failed to load. Check your connection and reload.', true);
    return;
  }
  let wb;
  try {
    wb = /\.csv$/i.test(file.name)
      ? XLSX.read(await file.text(), { type: 'string' })
      : XLSX.read(await file.arrayBuffer(), { type: 'array' });
  } catch (err) {
    showFileMessage(`Couldn't read "${file.name}": ${err.message}`, true);
    return;
  }
  state.fileName = file.name;
  state.sheets = wb.SheetNames.map((name) => ({ name, items: extractItems(wb.Sheets[name]) }));
  state.sheetChoice = defaultSheetChoice();
  loadItemsFromSheets();
}

// One entry per line; a URL anywhere on the line becomes that entry's link.
function parsePasted(text) {
  const items = [];
  for (const line of text.split(/\r?\n/)) {
    const match = line.match(/https?:\/\/\S+/i);
    const name = clean(match ? line.replace(match[0], ' ') : line)
      .replace(/^[-*•]\s+/, '')        // list bullets
      .replace(/[\s|,;:–—-]+$/, '');   // separators left over from "Name - link"
    if (name) items.push({ name, url: match ? safeUrl(match[0]) : '' });
  }
  return items;
}

// A pasted list is treated as a one-sheet file, so reset/restore/dedupe all work the same.
function usePastedList() {
  const items = parsePasted(el.pasteText.value);
  if (!items.length) {
    showFileMessage('Nothing to add: the pasted text has no entries.', true);
    return;
  }
  state.fileName = 'pasted list';
  state.sheets = [{ name: 'Pasted list', items }];
  state.sheetChoice = '0';
  loadItemsFromSheets();
  el.pasteBox.open = false;
}

// Default to the first sheet, unless it's empty and a later one isn't.
function defaultSheetChoice() {
  return String(Math.max(0, state.sheets.findIndex((s) => s.items.length)));
}

// Back to the state right after an upload: default sheet, full list, no history,
// wheel unrotated. Settings (spin time, seed, duplicates, auto-zoom) are kept.
function resetAll() {
  if (spinning || !state.sheets.length) return;
  if (!confirm('Reset the wheel? This restores every entry, returns to the default sheet, and clears the history.')) return;
  state.sheetChoice = defaultSheetChoice();
  state.history = [];
  rot = 0;
  zoomAnim = null;
  setZoom(1);
  loadItemsFromSheets();
  renderHistory();
  renderSeedHint();
}

function buildSource() {
  const picked = state.sheetChoice === MERGE ? state.sheets : [state.sheets[Number(state.sheetChoice)]].filter(Boolean);
  let items = picked.flatMap((s) => s.items);
  if (state.settings.dedupe) {
    const seen = new Map();
    for (const it of items) {
      const key = it.name.toLocaleLowerCase();
      const prev = seen.get(key);
      if (!prev) seen.set(key, { ...it });
      else if (!prev.url && it.url) prev.url = it.url;
    }
    items = [...seen.values()];
  }
  return items.map((it) => ({ ...it }));
}

function loadItemsFromSheets() {
  state.items = buildSource();
  state.spinCount = 0;
  save();
  onItemsChanged();
  renderSheetPicker();
  renderFileName();
}

/* --------------------------------------------------------------- wheel */

const view = { w: 0, h: 0, dpr: 1 };
let rot = 0;              // wheel rotation in radians; the pointer sits at angle 0 (3 o'clock)
let zoom = 1;
let zoomBeforeSpin = 1;
let spinAnim = null;
let zoomAnim = null;
let spinning = false;
let dirty = true;
let pendingWinner = -1;
let labelCache = new Map();
let lastTickerIndex = -1;

const mod = (a, n) => ((a % n) + n) % n;
const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));
const wrapAngle = (a) => mod(a + Math.PI, TAU) - Math.PI;
const easeOutQuart = (t) => 1 - (1 - t) ** 4;
const easeInOutCubic = (t) => (t < 0.5 ? 4 * t * t * t : 1 - (-2 * t + 2) ** 3 / 2);

function geometry(z) {
  const R0 = Math.min(view.w, view.h) / 2 - 18;
  const R = R0 * z;
  const px = view.w / 2 + R0; // pointer x stays fixed on the rim while zooming
  return { R0, R, cx: px - R, cy: view.h / 2, px };
}

// Zoom at which each slice is ~26px wide near the rim, enough for readable labels.
function autoZoomTarget() {
  const n = state.items.length;
  if (!n || !view.w) return 1;
  const { R0 } = geometry(1);
  return clamp((26 * n) / (TAU * R0 * 0.8), 1, 100);
}

const maxZoom = () => Math.max(2, Math.min(100, Math.ceil(autoZoomTarget() * 2)));

function sliceColor(i) {
  const hue = (i * 137.508) % 360; // golden angle keeps neighbours distinct at any count
  return `hsl(${hue.toFixed(1)} 70% ${i % 2 ? 46 : 56}%)`;
}

function fittedLabel(i, fs, maxW) {
  const key = `${i}|${fs.toFixed(1)}|${Math.round(maxW / 8)}`;
  let label = labelCache.get(key);
  if (label !== undefined) return label;
  label = state.items[i].name;
  if (ctx.measureText(label).width > maxW) {
    let lo = 0, hi = label.length;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (ctx.measureText(label.slice(0, mid) + '…').width <= maxW) lo = mid;
      else hi = mid - 1;
    }
    label = label.slice(0, lo).trimEnd() + '…';
  }
  if (labelCache.size > 20000) labelCache.clear();
  labelCache.set(key, label);
  return label;
}

function draw() {
  const { w, h, dpr } = view;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, w, h);
  const n = state.items.length;
  const g = geometry(zoom);

  if (!n) {
    ctx.beginPath();
    ctx.arc(g.cx, g.cy, g.R, 0, TAU);
    ctx.fillStyle = 'rgba(255,255,255,0.04)';
    ctx.fill();
    ctx.lineWidth = 2;
    ctx.strokeStyle = 'rgba(255,255,255,0.15)';
    ctx.stroke();
    return;
  }

  const s = TAU / n;
  // When zoomed, only slices near the pointer are on screen; skip the rest.
  const half = g.R <= Math.max(w, h) ? Math.PI : Math.asin(Math.min(1, (h / 2 + 10) / g.R));
  const fs = Math.min(20, g.R * s * 0.55);
  const maxW = Math.min(g.R * 0.8, w * 0.8);
  const seam = Math.min(s * 0.05, 0.004);

  ctx.save();
  ctx.translate(g.cx, g.cy);
  for (let i = 0; i < n; i++) {
    const a0 = rot + i * s;
    if (Math.abs(wrapAngle(a0 + s / 2)) > half + s) continue;
    ctx.beginPath();
    ctx.moveTo(0, 0);
    ctx.arc(0, 0, g.R, a0, a0 + s + (n > 1 ? seam : 0));
    ctx.closePath();
    ctx.fillStyle = sliceColor(i);
    ctx.fill();
  }

  if (fs >= 6) {
    ctx.font = `600 ${fs.toFixed(1)}px ${LABEL_FONT}`;
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    ctx.lineJoin = 'round';
    ctx.lineWidth = Math.max(2, fs / 5);
    ctx.strokeStyle = 'rgba(15, 10, 30, 0.6)';
    ctx.fillStyle = '#fff';
    for (let i = 0; i < n; i++) {
      const mid = rot + (i + 0.5) * s;
      if (Math.abs(wrapAngle(mid)) > half + s) continue;
      const label = fittedLabel(i, fs, maxW);
      ctx.save();
      ctx.rotate(mid);
      ctx.strokeText(label, g.R - 14, 0);
      ctx.fillText(label, g.R - 14, 0);
      ctx.restore();
    }
  }

  // Rim and hub.
  ctx.beginPath();
  ctx.arc(0, 0, g.R, 0, TAU);
  ctx.lineWidth = 6;
  ctx.strokeStyle = '#ffc94d';
  ctx.stroke();
  if (half === Math.PI) {
    const hub = Math.max(14, g.R * 0.07);
    ctx.beginPath();
    ctx.arc(0, 0, hub, 0, TAU);
    ctx.fillStyle = '#1f1b33';
    ctx.fill();
    ctx.lineWidth = 4;
    ctx.stroke();
  }
  ctx.restore();

  // Pointer, plus a guide line across the readable band when zoomed in.
  if (zoom > 1.5) {
    ctx.fillStyle = 'rgba(255,255,255,0.35)';
    ctx.fillRect(g.px - Math.min(maxW + 30, g.R), g.cy - 1, Math.min(maxW + 30, g.R), 2);
  }
  ctx.beginPath();
  ctx.moveTo(g.px - 24, g.cy);
  ctx.lineTo(g.px + 12, g.cy - 15);
  ctx.lineTo(g.px + 12, g.cy + 15);
  ctx.closePath();
  ctx.fillStyle = '#ff4f8b';
  ctx.fill();
  ctx.lineWidth = 2.5;
  ctx.strokeStyle = '#fff';
  ctx.stroke();
}

function pointerIndex() {
  const n = state.items.length;
  if (!n) return -1;
  return Math.floor(mod(-rot, TAU) / (TAU / n)) % n;
}

function updateTicker() {
  const i = pointerIndex();
  if (i === lastTickerIndex) return;
  lastTickerIndex = i;
  el.ticker.textContent = i < 0 ? '—' : state.items[i].name;
  if (spinAnim) playTick();
}

function resize() {
  const size = Math.floor(el.wheelWrap.getBoundingClientRect().width);
  if (!size) return;
  view.w = view.h = size;
  view.dpr = window.devicePixelRatio || 1;
  el.wheel.width = Math.round(size * view.dpr);
  el.wheel.height = Math.round(size * view.dpr);
  labelCache.clear();
  updateZoomRange();
  dirty = true;
}

function frame(now) {
  if (spinAnim) stepSpin(now);
  if (zoomAnim) stepZoom(now);
  if (dirty) {
    draw();
    updateTicker();
    dirty = false;
  }
  requestAnimationFrame(frame);
}

/* ---------------------------------------------------------------- zoom */

function setZoom(z) {
  zoom = clamp(z, 1, maxZoom());
  el.zoom.value = zoom.toFixed(1);
  el.zoomOut.textContent = `${zoom < 10 ? zoom.toFixed(1) : Math.round(zoom)}×`;
  dirty = true;
}

function updateZoomRange() {
  el.zoom.max = String(maxZoom());
  setZoom(zoom);
}

function animateZoom(to, ms) {
  if (reducedMotion) {
    zoomAnim = null;
    setZoom(to);
    return;
  }
  zoomAnim = { from: zoom, to: clamp(to, 1, maxZoom()), start: performance.now(), ms };
}

function stepZoom(now) {
  const t = Math.min(1, (now - zoomAnim.start) / zoomAnim.ms);
  // Interpolate geometrically so zooming feels even at every scale.
  setZoom(zoomAnim.from * (zoomAnim.to / zoomAnim.from) ** easeInOutCubic(t));
  if (t >= 1) zoomAnim = null;
}

function userZoom(z) {
  zoomAnim = null;
  setZoom(z);
}

/* ---------------------------------------------------------------- spin */

function spin() {
  const n = state.items.length;
  if (spinning || !n) return;
  const rng = makeRng(state.settings.seed, state.spinCount);
  let winner = Math.floor(rng() * n);
  // Easter egg: the exact seed "OnePiece" (no space) always lands on One Piece when it's on the wheel.
  if (state.settings.seed === 'OnePiece') {
    const onePiece = state.items.findIndex((it) => it.name.trim().toLowerCase() === 'one piece');
    if (onePiece >= 0) winner = onePiece;
  }
  const offset = 0.15 + rng() * 0.7; // land inside the slice, never on an edge
  const duration = clamp(Number(state.settings.duration) || 5, MIN_DURATION, MAX_DURATION) * 1000;
  const s = TAU / n;
  const turns = Math.max(3, Math.round((duration / 1000) * 1.2));
  const delta = mod(-(winner + offset) * s - rot, TAU) + turns * TAU;

  state.spinCount++;
  save();
  renderSeedHint();
  unlockAudio(); // browsers only allow audio to start from a click or key press
  spinning = true;
  setControlsDisabled(true);
  spinAnim = { start: performance.now(), from: rot, delta, duration, winner };
  zoomBeforeSpin = zoom;
  if (state.settings.autoZoom) animateZoom(autoZoomTarget(), clamp(duration * 0.15, 600, 1500));
}

function stepSpin(now) {
  const a = spinAnim;
  const t = Math.min(1, (now - a.start) / a.duration);
  rot = a.from + a.delta * easeOutQuart(t);
  dirty = true;
  if (t >= 1) {
    rot = mod(rot, TAU);
    spinAnim = null;
    finishSpin(a.winner);
  }
}

function finishSpin(index) {
  const item = state.items[index];
  state.history.unshift({ name: item.name, url: item.url || '', at: Date.now(), removed: false });
  save();
  renderHistory();
  pendingWinner = index;
  showWinner(item);
}

function showWinner(item) {
  el.winnerName.textContent = item.name;
  el.winnerLink.hidden = !item.url;
  if (item.url) {
    el.winnerLink.href = item.url;
    el.winnerLink.textContent = linkLabel(item.url);
  } else {
    el.winnerLink.removeAttribute('href');
  }
  el.modal.hidden = false;
  el.keepBtn.focus();
  burstConfetti();
  playFanfare();
}

function linkLabel(url) {
  try {
    const host = new URL(url).hostname.replace(/^www\./, '');
    return host === 'myanimelist.net' ? 'Open on MyAnimeList ↗' : `Open on ${host} ↗`;
  } catch {
    return 'Open link ↗';
  }
}

function closeWinner(remove) {
  if (el.modal.hidden) return;
  el.modal.hidden = true;
  if (remove && pendingWinner >= 0) {
    state.items.splice(pendingWinner, 1);
    state.history[0].removed = true;
    save();
    renderHistory();
    onItemsChanged();
  }
  pendingWinner = -1;
  spinning = false;
  setControlsDisabled(false);
  if (state.settings.autoZoom) animateZoom(zoomBeforeSpin, 700);
  el.spinBtn.focus();
}

/* --------------------------------------------------------------- sound */

const TICK_MIN_GAP_MS = 13; // ~40 ticks/s max; one tick per frame is also a hard limit (60/s on most screens)
let audio = null;
let lastTickAt = 0;

function unlockAudio() {
  const AC = window.AudioContext || window.webkitAudioContext;
  if (!AC) return;
  audio ??= new AC();
  if (audio.state === 'suspended') audio.resume();
}

// A short square-wave blip with a fast decay: the 8-bit sound that fits the pixel look.
function blip(freq, at, length, volume) {
  const osc = audio.createOscillator();
  const gain = audio.createGain();
  osc.type = 'square';
  osc.frequency.setValueAtTime(freq, at);
  gain.gain.setValueAtTime(volume, at);
  gain.gain.exponentialRampToValueAtTime(0.0001, at + length);
  osc.connect(gain).connect(audio.destination);
  osc.start(at);
  osc.stop(at + length + 0.02);
}

function playTick() {
  if (state.settings.muted || !audio) return;
  // At full speed hundreds of slices pass per second; cap it so it stays a clicking sound, not a buzz.
  const now = performance.now();
  if (now - lastTickAt < TICK_MIN_GAP_MS) return;
  lastTickAt = now;
  blip(1150 + Math.random() * 150, audio.currentTime, 0.018, 0.05);
}

function playFanfare() {
  if (state.settings.muted || !audio) return;
  const t = audio.currentTime;
  [523.25, 659.25, 783.99, 1046.5].forEach((f, i) => blip(f, t + i * 0.09, i === 3 ? 0.45 : 0.12, 0.07));
}

function renderMute() {
  const muted = state.settings.muted;
  el.muteBtn.textContent = muted ? '🔇' : '🔊';
  el.muteBtn.setAttribute('aria-pressed', String(muted));
  el.muteBtn.setAttribute('aria-label', muted ? 'Unmute sound' : 'Mute sound');
  el.muteBtn.title = muted ? 'Unmute sound' : 'Mute sound';
}

/* ------------------------------------------------------------ confetti */

function burstConfetti() {
  if (reducedMotion) return;
  const c = el.confetti;
  const cx = c.getContext('2d');
  const dpr = window.devicePixelRatio || 1;
  const w = window.innerWidth;
  const h = window.innerHeight;
  c.width = w * dpr;
  c.height = h * dpr;
  cx.setTransform(dpr, 0, 0, dpr, 0, 0);
  const colors = ['#ff4f8b', '#ffc94d', '#6ee7ff', '#a78bfa', '#4ade80', '#ffffff'];
  const parts = Array.from({ length: 220 }, (_, i) => {
    const fromLeft = i % 2 === 0;
    return {
      x: fromLeft ? 0 : w,
      y: h * 0.7,
      vx: (fromLeft ? 1 : -1) * (6 + Math.random() * 12),
      vy: -(10 + Math.random() * 14),
      size: 6 + Math.random() * 6,
      rot: Math.random() * TAU,
      vr: (Math.random() - 0.5) * 0.4,
      flip: Math.random() * TAU,
      color: colors[i % colors.length],
    };
  });
  const start = performance.now();
  (function tick(now) {
    cx.clearRect(0, 0, w, h);
    let alive = false;
    for (const p of parts) {
      p.vy += 0.35;
      p.vx *= 0.985;
      p.x += p.vx;
      p.y += p.vy;
      p.rot += p.vr;
      p.flip += 0.15;
      if (p.y < h + 20) alive = true;
      cx.save();
      cx.translate(p.x, p.y);
      cx.rotate(p.rot);
      cx.scale(1, Math.cos(p.flip));
      cx.fillStyle = p.color;
      cx.fillRect(-p.size / 2, -p.size / 3, p.size, p.size * 0.66);
      cx.restore();
    }
    if (alive && now - start < 5000) requestAnimationFrame(tick);
    else cx.clearRect(0, 0, w, h);
  })(start);
}

/* ------------------------------------------------------------ rendering */

function onItemsChanged() {
  labelCache.clear();
  lastTickerIndex = null;
  if (!state.items.length) el.ticker.textContent = '—';
  updateZoomRange();
  renderItems();
  dirty = true;
}

function renderItems() {
  const n = state.items.length;
  el.count.textContent = String(n);
  el.emptyHint.hidden = n > 0;
  el.itemList.replaceChildren(
    ...state.items.map((it) => {
      const li = document.createElement('li');
      li.textContent = it.name;
      return li;
    }),
  );
  setControlsDisabled(spinning);
}

function renderSheetPicker() {
  el.sheet.replaceChildren();
  for (const [i, s] of state.sheets.entries()) {
    el.sheet.add(new Option(`${s.name} (${s.items.length})`, String(i)));
  }
  if (state.sheets.length > 1) {
    const total = state.sheets.reduce((sum, s) => sum + s.items.length, 0);
    el.sheet.add(new Option(`All sheets merged (${total})`, MERGE));
  }
  el.sheet.value = state.sheetChoice;
  el.sheet.disabled = spinning || state.sheets.length < 2;
}

function renderFileName() {
  showFileMessage(state.fileName ? `Loaded: ${state.fileName}` : '', false);
}

function showFileMessage(msg, isError) {
  el.fileName.textContent = msg;
  el.fileName.classList.toggle('error', isError);
}

function renderHistory() {
  el.historyEmpty.hidden = state.history.length > 0;
  el.clearHistory.disabled = state.history.length === 0;
  el.history.replaceChildren(
    ...state.history.map((h, i) => {
      const li = document.createElement('li');
      const name = document.createElement(h.url ? 'a' : 'span');
      name.className = 'name';
      name.textContent = h.name;
      if (h.url) {
        name.href = h.url;
        name.target = '_blank';
        name.rel = 'noopener noreferrer';
      }
      const meta = document.createElement('span');
      meta.className = 'meta';
      meta.textContent = `#${state.history.length - i} · ${new Date(h.at).toLocaleTimeString()}`;
      const badge = document.createElement('span');
      badge.className = `badge ${h.removed ? 'removed' : 'kept'}`;
      badge.textContent = h.removed ? 'removed' : 'kept';
      li.append(name, badge, meta);
      return li;
    }),
  );
}

function renderSeedHint() {
  el.seedHint.textContent = state.settings.seed
    ? `Seeded: next is spin #${state.spinCount + 1} for this seed. Same seed and same list give the same results.`
    : 'No seed: every spin is fully random.';
}

function renderSettings() {
  el.duration.value = el.durationRange.value = String(state.settings.duration);
  el.seed.value = state.settings.seed;
  el.autoZoom.checked = state.settings.autoZoom;
  el.dedupe.checked = state.settings.dedupe;
  renderSeedHint();
  renderMute();
}

function setControlsDisabled(disabled) {
  const n = state.items.length;
  el.spinBtn.disabled = disabled || n === 0;
  el.shuffleBtn.disabled = disabled || n < 2;
  el.resetBtn.disabled = disabled || state.sheets.length === 0;
  el.resetAllBtn.disabled = disabled || state.sheets.length === 0;
  el.sheet.disabled = disabled || state.sheets.length < 2;
  // Settings are locked mid-spin too: editing the seed would reset its spin counter halfway through.
  for (const input of [el.dedupe, el.durationRange, el.duration, el.seed, el.autoZoom, el.pasteText]) {
    input.disabled = disabled;
  }
  el.pasteBtn.disabled = disabled || !el.pasteText.value.trim();
  el.file.disabled = disabled;
  el.drop.classList.toggle('disabled', disabled);
}

/* -------------------------------------------------------------- events */

function bindEvents() {
  el.spinBtn.addEventListener('click', spin);
  el.muteBtn.addEventListener('click', () => {
    state.settings.muted = !state.settings.muted;
    if (!state.settings.muted) unlockAudio();
    save();
    renderMute();
  });

  el.pasteText.addEventListener('input', () => {
    el.pasteBtn.disabled = spinning || !el.pasteText.value.trim();
  });
  el.pasteBtn.addEventListener('click', usePastedList);

  el.file.addEventListener('change', () => {
    if (el.file.files[0]) readFile(el.file.files[0]);
    el.file.value = '';
  });

  // Accept a drop anywhere on the page so a near-miss doesn't open the file in the browser.
  let dragDepth = 0;
  window.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dragDepth++;
    el.drop.classList.add('over');
  });
  window.addEventListener('dragleave', () => {
    if (--dragDepth <= 0) {
      dragDepth = 0;
      el.drop.classList.remove('over');
    }
  });
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => {
    e.preventDefault();
    dragDepth = 0;
    el.drop.classList.remove('over');
    const file = e.dataTransfer?.files?.[0];
    if (file && !spinning) readFile(file);
  });

  el.sheet.addEventListener('change', () => {
    state.sheetChoice = el.sheet.value;
    loadItemsFromSheets();
  });
  el.dedupe.addEventListener('change', () => {
    state.settings.dedupe = el.dedupe.checked;
    if (state.sheets.length) loadItemsFromSheets();
    else save();
  });

  el.shuffleBtn.addEventListener('click', () => {
    const a = state.items;
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1)); // shuffle deliberately ignores the seed
      [a[i], a[j]] = [a[j], a[i]];
    }
    save();
    onItemsChanged();
  });
  el.resetBtn.addEventListener('click', loadItemsFromSheets);
  el.resetAllBtn.addEventListener('click', resetAll);

  const setDuration = (v, commit) => {
    const num = Number(v);
    if (!Number.isFinite(num)) return;
    const d = clamp(num, MIN_DURATION, MAX_DURATION);
    state.settings.duration = d;
    el.durationRange.value = String(d);
    if (commit) el.duration.value = String(d);
    save();
  };
  el.durationRange.addEventListener('input', () => {
    el.duration.value = el.durationRange.value;
    setDuration(el.durationRange.value, true);
  });
  el.duration.addEventListener('input', () => setDuration(el.duration.value, false));
  el.duration.addEventListener('change', () => setDuration(el.duration.value, true));

  el.seed.addEventListener('input', () => {
    state.settings.seed = el.seed.value;
    state.spinCount = 0;
    save();
    renderSeedHint();
  });
  el.autoZoom.addEventListener('change', () => {
    state.settings.autoZoom = el.autoZoom.checked;
    save();
  });

  el.zoom.addEventListener('input', () => userZoom(Number(el.zoom.value)));
  el.wheel.addEventListener(
    'wheel',
    (e) => {
      // Plain scrolling moves the page; Ctrl/⌘ + scroll (also what a trackpad pinch sends) zooms.
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      userZoom(zoom * Math.exp(-e.deltaY * 0.0015));
    },
    { passive: false },
  );
  el.wheel.addEventListener('dblclick', () => animateZoom(1, 400));

  el.clearHistory.addEventListener('click', () => {
    state.history = [];
    save();
    renderHistory();
  });

  el.removeBtn.addEventListener('click', () => closeWinner(true));
  el.keepBtn.addEventListener('click', () => closeWinner(false));
  el.modal.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeWinner(false);
  });

  document.addEventListener('keydown', (e) => {
    if (e.code !== 'Space' || !el.modal.hidden) return;
    if (e.target.closest('input, select, textarea, button, a, summary')) return;
    e.preventDefault();
    spin();
  });

  new ResizeObserver(resize).observe(el.wheelWrap);
}

/* ---------------------------------------------------------------- init */

renderSettings();
renderSheetPicker();
renderFileName();
renderItems();
renderHistory();
bindEvents();
resize();
requestAnimationFrame(frame);

// The canvas can't swap fonts on its own; redraw the labels once the web font arrives.
document.fonts?.load(`600 16px ${LABEL_FONT}`).then(() => {
  labelCache.clear();
  dirty = true;
});
