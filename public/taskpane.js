
// ----------------------------
// ProofWrite taskpane.js (University Standard Edition)
// Version: 2025-12-02_v4_UNI  (timestamps + notify + safe-ui fixes + anchor-first export)
// ----------------------------

/* ============ CONFIG ============ */
const APP_VERSION = "2025-12-02_v4_UNI";
const THRESHOLD_LARGE_PASTE_WORDS = 50;
const THRESHOLD_SPEED_SPIKE_WORDS = 20;
const THRESHOLD_AUTO_START_JUMP = 50;
const SUSTAINED_SPEED_WPM = 120;
const SUSTAINED_SPEED_PERIOD_MS = 60000;

const MONITOR_POLL_MS = 1500;
const UI_TIMER_MS = 1500;
const AUTO_STOP_THRESHOLD_MS = 300000; // 5 minutes

const PENALTY_LARGE_PASTE_MAX = 50;
const PENALTY_SPEED_SPIKE_MAX = 30;
const PENALTY_SUSTAINED_SPEED = 20;

const DOC_PROP_KEY = "ProofWriteSessions";
const RUNTIME_STORAGE_KEY = "ProofWrite_sessions_backup";
const DIAG_KEY = "ProofWrite_diagnostics";
const TOKEN_STORAGE_KEY = "ProofWrite_verification_token";

/* ============ GLOBAL STATE ============ */
let verificationToken = null;
let isSessionActive = false;
let sessionStartTime = 0;
let initialWordCount = 0;
let uiTimerId = null;
let sessions = [];
let lastWordCount = 0;
let monitoringTimerId = null;
let currentSession = null;
let lastActivityTime = 0;
let exportJsonRunning = false;
let isExporting = false;

let recentWordCounts = [];

let sessionTimeEl, wordCountEl, charCountEl, lastReportEl, autoBannerEl, flaggedSummaryEl, exportBtnEl;

/* ============ HELPERS: Notifications + Timestamp formatting ============ */
function notify(message, level = "info", duration = 3000) {
  try {
    const id = "proofwrite-notify";
    let el = document.getElementById(id);
    if (!el) {
      el = document.createElement("div");
      el.id = id;
      el.style.position = "fixed";
      el.style.right = "18px";
      el.style.bottom = "18px";
      el.style.zIndex = 99999;
      el.style.minWidth = "180px";
      el.style.maxWidth = "420px";
      el.style.padding = "12px 14px";
      el.style.borderRadius = "10px";
      el.style.boxShadow = "0 6px 20px rgba(0,0,0,0.12)";
      el.style.fontSize = "13px";
      el.style.fontFamily = "inherit";
      el.style.color = "#fff";
      el.style.display = "none";
      document.body.appendChild(el);
    }

    switch (level) {
      case "success": el.style.background = "#188038"; break;
      case "warn":    el.style.background = "#d97706"; break;
      case "error":   el.style.background = "#b91c1c"; break;
      default:        el.style.background = "#0b69c7"; break;
    }

    el.textContent = message;
    el.style.display = "block";
    if (el._timeout) clearTimeout(el._timeout);
    el._timeout = setTimeout(() => { el.style.display = "none"; el._timeout = null; }, duration);
    return Promise.resolve(true);
  } catch (err) {
    console.warn("notify fallback:", message, err);
    return Promise.resolve(false);
  }
}

function formatTimestamp(input) {
  try {
    if (!input) return "(N/A)";
    const d = (input instanceof Date) ? input : (typeof input === "number" ? new Date(input) : new Date(String(input)));
    if (isNaN(d.getTime())) return "(Invalid date)";
    return d.toLocaleString(undefined, {
      year: "numeric", month: "short", day: "2-digit",
      hour: "2-digit", minute: "2-digit", second: "2-digit"
    });
  } catch (err) {
    console.warn("formatTimestamp failed", err);
    return String(input);
  }
}

/* ============ Graph helpers (calls into graphs.js if present) ============ */
function safeInitGraphs() {
  if (typeof window !== "undefined" && typeof window.initGraphs === "function") {
    try { window.initGraphs(); return true; } catch (e) { console.warn("initGraphs threw", e); }
  }
  return false;
}
function safeUpdateLiveGraphs(events) {
  // events: [{ t: <ms timestamp>, c: <chars> }, ...]
  if (typeof window !== "undefined") {
    if (typeof window.updateLiveGraphs === "function") {
      try { window.updateLiveGraphs(events); return true; } catch (e) { console.warn("updateLiveGraphs threw", e); }
    }
    if (typeof window.updateGraphs === "function") {
      try { window.updateGraphs(events); return true; } catch (e) { console.warn("updateGraphs threw", e); }
    }
  }
  return false;
}
function safeRenderSessionGraphs(session) {
  if (!session) return false;
  if (typeof window !== "undefined") {
    if (typeof window.renderSessionGraphs === "function") {
      try { window.renderSessionGraphs(session); return true; } catch (e) { console.warn("renderSessionGraphs threw", e); }
    }
    if (typeof window.updateGraphs === "function") {
      try { window.updateGraphs(session.events || []); return true; } catch (e) { console.warn("updateGraphs threw", e); }
    }
  }
  return false;
}

/* ============ Initialization ============ */
Office.onReady(async () => {
  console.log("Office.js is ready");

  // try to initialise graphs (graphs.js is loaded in your HTML before this script)
  safeInitGraphs();

  try {
    sessionTimeEl = document.getElementById("sessionTime");
    charCountEl = document.getElementById("charCount");
    lastReportEl = document.getElementById("lastReport");
    autoBannerEl = document.getElementById("autoBanner");
    flaggedSummaryEl = document.getElementById("flaggedSummary");
    exportBtnEl = document.getElementById("exportBtn");

    if (exportBtnEl) exportBtnEl.addEventListener("click", handleExportButton);

    const historyHeader = document.getElementById("historyHeader");
    const historyContainer = document.getElementById("pastSessionsContainer");
    const historyChevron = document.getElementById("historyChevron");
    if (historyHeader && historyContainer) {
      historyHeader.addEventListener("click", () => {
        const isHidden = historyContainer.style.display === "none";
        historyContainer.style.display = isHidden ? "block" : "none";
        if (historyChevron) {
          historyChevron.textContent = isHidden ? "▲ (Click to Hide)" : "▼ (Click to Show)";
        }
      });
    }

    // attach click handler to the flaggedSummary container (event delegation)
    if (flaggedSummaryEl) {
      flaggedSummaryEl.addEventListener("click", (ev) => {
        const btn = ev.target.closest && ev.target.closest(".session-item");
        if (!btn) return;
        const idx = btn.getAttribute("data-idx");
        if (idx == null) return;
        const index = parseInt(idx, 10);
        if (Number.isNaN(index) || index < 0 || index >= sessions.length) {
          console.warn("session click index OOB", idx);
          return;
        }
        // render graphs for clicked session
        const session = sessions[index];
        if (!session) return;
        // call graphs renderer (safe)
        const ok = safeRenderSessionGraphs(session);
        if (!ok) {
          notify("No graph renderer available (graphs.js must expose renderSessionGraphs or updateGraphs).", "warn", 2500);
        }
      });
    }

    if (typeof initUI === "function") {
      try { initUI(); } catch (e) { console.warn("initUI failed", e); }
    }

    await loadSessionsFromStorage();
    updateFlaggedSummary();

    try {
      verificationToken = await loadVerificationToken();
    } catch (e) {
      console.warn("loadVerificationToken failed", e);
      verificationToken = null;
    }
    updateVerificationTokenUI();

    // copy token button
    const copyTokenBtnEl = document.getElementById("copyTokenBtn");
    if (copyTokenBtnEl) {
      copyTokenBtnEl.addEventListener("click", async () => {
        if (!verificationToken) {
          await notify("No verification token yet. Export first.", "warn");
          return;
        }
        const ok = await copyToClipboardSafe(verificationToken);
        if (ok) {
          copyTokenBtnEl.textContent = "Copied ✓";
          copyTokenBtnEl.style.background = "#0a8a38";
          setTimeout(() => {
            copyTokenBtnEl.textContent = "Copy Token";
            copyTokenBtnEl.style.background = "";
          }, 1400);
        } else {
          await notify("Copy failed. Please copy token manually.", "error");
        }
      });
    }

    await startMonitoring();
    console.info("ProofWrite ready, version:", APP_VERSION);
  } catch (err) {
    await logDiagnostic("Office.onReady failed", err);
    await notify("Initialization error — check diagnostics console.", "error");
  }
});

/* ============ LOGGING ============ */
async function logDiagnostic(message, extra = null) {
  const entry = { ts: new Date().toISOString(), msg: String(message), extra, version: APP_VERSION };
  console.warn("ProofWrite diag:", entry);
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      const prev = await OfficeRuntime.storage.getItem(DIAG_KEY) || "[]";
      const arr = JSON.parse(prev);
      arr.push(entry);
      while (arr.length > 200) arr.shift();
      await OfficeRuntime.storage.setItem(DIAG_KEY, JSON.stringify(arr));
    }
  } catch (err) { console.warn("diag storage failed", err); }
}

/* ============ OFFICE WRAPPERS ============ */
async function safeWordRun(fn) {
  if (typeof Word === "undefined" || typeof Word.run !== "function") {
    await logDiagnostic("Word.run unavailable", { typeofWord: typeof Word });
    return await fn({ fallback: true });
  }
  try { return await Word.run(async (ctx) => await fn(ctx)); }
  catch (err) { await logDiagnostic("Word.run threw", err); throw err; }
}

async function getDocumentText() {
  try {
    return await safeWordRun(async (ctx) => {
      const body = ctx.document.body;
      body.load("text");
      await ctx.sync();
      return body.text || "";
    });
  } catch (err) { await logDiagnostic("getDocumentText failed", err); return null; }
}

async function getDocumentWordCount() {
  const txt = await getDocumentText();
  if (txt === null) return lastWordCount;
  const trimmed = txt.trim();
  if (!trimmed) return 0;
  return trimmed.split(/\s+/).filter(w => w.length > 0).length;
}

/* ============ getDocumentHash (robust via slices) ============ */
async function getDocumentHash() {
  let file = null;
  try {
    file = await new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 64 * 1024 },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
          else reject(new Error("getFileAsync failed: " + (result.error && result.error.message)));
        }
      );
    });

    if (!file) {
      await logDiagnostic("getDocumentHash: no file");
      return null;
    }

    const sliceCount = file.sliceCount || 0;
    if (sliceCount === 0) {
      try { file.closeAsync(); } catch (e) {}
      await logDiagnostic("getDocumentHash: sliceCount 0");
      return null;
    }

    const views = [];
    for (let i = 0; i < sliceCount; i++) {
      /* eslint-disable no-await-in-loop */
      const slice = await new Promise((resolve, reject) => {
        file.getSliceAsync(i, (sliceResult) => {
          if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(sliceResult.value.data);
          } else {
            reject(new Error("getSliceAsync failed at index " + i + ": " + (sliceResult.error && sliceResult.error.message)));
          }
        });
      });

      if (slice instanceof ArrayBuffer) views.push(new Uint8Array(slice));
      else if (ArrayBuffer.isView(slice)) views.push(new Uint8Array(slice.buffer, slice.byteOffset, slice.byteLength));
      else {
        try { views.push(new Uint8Array(slice)); } catch (coerceErr) {
          try { file.closeAsync(); } catch (e) {}
          throw new Error("Unsupported slice type at index " + i);
        }
      }
    }

    try { file.closeAsync(); } catch (e) {}

    const totalLength = views.reduce((acc, v) => acc + v.length, 0);
    if (totalLength === 0) {
      await logDiagnostic("getDocumentHash: totalLength 0");
      return null;
    }

    const combined = new Uint8Array(totalLength);
    let offset = 0;
    for (const v of views) { combined.set(v, offset); offset += v.length; }

    const hashBuffer = await crypto.subtle.digest("SHA-256", combined);
    const hashHex = Array.from(new Uint8Array(hashBuffer)).map(b => b.toString(16).padStart(2, "0")).join("");
    await logDiagnostic("getDocumentHash: computed", { hash: hashHex, bytes: totalLength });
    return hashHex;
  } catch (err) {
    await logDiagnostic("getDocumentHash failed", String(err));
    try { if (file && typeof file.closeAsync === "function") file.closeAsync(); } catch (e) {}
    return null;
  }
}

/* ============ PERSISTENCE ============ */
async function storeSessionsLocally() {
  const payload = { sessions, version: APP_VERSION, savedAt: new Date().toISOString() };
  const json = JSON.stringify(payload);
  try {
    if (OfficeRuntime?.storage) { await OfficeRuntime.storage.setItem(RUNTIME_STORAGE_KEY, json); return true; }
  } catch (err) { await logDiagnostic("OfficeRuntime.storage set failed", err); }
  try {
    await safeWordRun(async (ctx) => {
      const props = ctx.document.properties.customProperties;
      props.load("items");
      await ctx.sync();
      const existing = props.items.find(p => p.key === DOC_PROP_KEY);
      if (existing) existing.delete();
      props.add(DOC_PROP_KEY, json);
      await ctx.sync();
    });
    return true;
  } catch (err) { await logDiagnostic("customProperties save failed", err); return false; }
}

async function loadSessionsFromStorage() {
  try {
    if (OfficeRuntime?.storage) {
      const raw = await OfficeRuntime.storage.getItem(RUNTIME_STORAGE_KEY);
      if (raw) { const parsed = JSON.parse(raw); if (Array.isArray(parsed.sessions)) { sessions = parsed.sessions; return true; } }
    }
  } catch (err) { await logDiagnostic("OfficeRuntime.storage read failed", err); }
  try {
    await safeWordRun(async (ctx) => {
      const props = ctx.document.properties.customProperties;
      props.load("items");
      await ctx.sync();
      const existing = props.items.find(p => p.key === DOC_PROP_KEY);
      if (existing?.value) { const parsed = JSON.parse(existing.value); if (Array.isArray(parsed.sessions)) sessions = parsed.sessions; }
    });
    return true;
  } catch (err) { await logDiagnostic("customProperties read failed", err); return false; }
}

/* ============ Verification token persistence ============ */
async function saveVerificationToken(token) {
  try {
    if (OfficeRuntime?.storage) {
      await OfficeRuntime.storage.setItem(TOKEN_STORAGE_KEY, token);
      return true;
    } else if (typeof localStorage !== "undefined") {
      localStorage.setItem(TOKEN_STORAGE_KEY, token);
      return true;
    }
  } catch (err) { await logDiagnostic("saveVerificationToken failed", err); }
  return false;
}
async function loadVerificationToken() {
  try {
    if (OfficeRuntime?.storage) {
      return await OfficeRuntime.storage.getItem(TOKEN_STORAGE_KEY) || null;
    } else if (typeof localStorage !== "undefined") {
      return localStorage.getItem(TOKEN_STORAGE_KEY) || null;
    }
  } catch (err) { await logDiagnostic("loadVerificationToken failed", err); }
  return null;
}
async function clearVerificationToken() {
  if (!verificationToken) return;
  verificationToken = null;
  updateVerificationTokenUI();
  try {
    if (OfficeRuntime?.storage) await OfficeRuntime.storage.removeItem(TOKEN_STORAGE_KEY);
    else if (typeof localStorage !== "undefined") localStorage.removeItem(TOKEN_STORAGE_KEY);
    await logDiagnostic("Verification token cleared due to document activity.");
  } catch (err) { await logDiagnostic("clearVerificationToken storage failed", err); }
}

/* ============ CRYPTO + TOKEN BUILD ============ */
async function sha256(data) {
  let buffer;
  if (typeof data === 'string') buffer = new TextEncoder().encode(data);
  else if (data instanceof ArrayBuffer) buffer = data;
  else throw new Error("sha256 input must be string or ArrayBuffer.");
  const hashBuffer = await crypto.subtle.digest("SHA-256", buffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, "0")).join("");
}
function generateToken() {
  return "HV-" + [...crypto.getRandomValues(new Uint8Array(4))].map(x => x.toString(16).padStart(2, "0")).join("").toUpperCase();
}
async function buildVerificationPackage(sessionData) {
  const jsonString = JSON.stringify(sessionData);
  const hash = await sha256(jsonString);
  const token = generateToken();
  return { token, hash, session: sessionData };
}

function updateVerificationTokenUI() {
  const container = document.getElementById("verificationContainer");
  const tokenEl = document.getElementById("verificationToken");
  const copyBtn = document.getElementById("copyTokenBtn");
  if (!tokenEl || !container) return;
  if (verificationToken) {
    tokenEl.innerHTML = `Token: <b>${verificationToken}</b>`;
    container.style.display = "flex";
    if (copyBtn) copyBtn.disabled = false;
  } else {
    tokenEl.textContent = "Token: —";
    container.style.display = "none";
    if (copyBtn) copyBtn.disabled = true;
  }
}

function updateDocumentHashUI(hash) {
  const el = document.getElementById("documentHashDisplay");
  if (!el) return;
  el.textContent = hash ? hash : "(Unavailable)";
}

/* ============ EXPORT LOGIC (ANCHOR-FIRST) ============ */
async function exportJson(payload, filename) {
  if (exportJsonRunning) {
    await logDiagnostic("exportJson blocked (already running)");
    return { success: false, method: "double_block" };
  }
  exportJsonRunning = true;
  try {
    const totalActiveMs = sessions.reduce((sum, s) => {
      const start = new Date(s.startTime).getTime();
      const end = new Date(s.endTime || Date.now()).getTime();
      return sum + (end - start);
    }, 0);
    const totalCharacters = sessions.reduce((sum, s) => sum + (s.charactersTyped || 0), 0);
    const totalMinutes = Math.floor(totalActiveMs / 60000);
    const totalSeconds = Math.floor((totalActiveMs % 60000) / 1000);

    const docHash = await getDocumentHash();

    const summary = {
      totalSessions: sessions.length,
      totalActiveTime: `${totalMinutes} min ${totalSeconds} sec`,
      totalCharactersTyped: totalCharacters,
      documentIntegrityHash: docHash || "N/A - Failed to compute",
    };

    const detailedSessions = sessions.map((s, i) => {
      const start = new Date(s.startTime);
      const end = new Date(s.endTime || Date.now());
      const activeMs = end - start;
      const activeMin = Math.floor(activeMs / 60000);
      const activeSec = Math.floor((activeMs % 60000) / 1000);
      const cpm = activeMs > 0 ? Math.round((s.charactersTyped || 0) / (activeMs / 60000)) : 0;
      return {
        sessionNumber: i + 1,
        startTimeISO: start.toISOString(),
        endTimeISO: end.toISOString(),
        startTimeFormatted: formatTimestamp(start),
        endTimeFormatted: formatTimestamp(end),
        activeWritingTime: `${activeMin} min ${activeSec} sec`,
        charactersTyped: s.charactersTyped || 0,
        charactersPerMinute: cpm,
        flags: s.flags || {},
        edits: s.edits || []
      };
    });

    const exportPayload = {
      summary,
      sessions: detailedSessions,
      exportedAt: new Date().toISOString(),
      exportedAtFormatted: formatTimestamp(new Date()),
      exporterVersion: APP_VERSION
    };

    const verificationPackage = await buildVerificationPackage(exportPayload);
    if (docHash) {
      verificationPackage.documentHash = docHash;
      updateDocumentHashUI(docHash);
    }

    verificationToken = verificationPackage.token;
    await saveVerificationToken(verificationToken);
    updateVerificationTokenUI();

    const json = JSON.stringify(verificationPackage, null, 2);

    // PRIMARY: Anchor download (user device).
    try {
      const blob = new Blob([json], { type: "application/json;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      a.style.display = "none";
      document.body.appendChild(a);

      const evt = document.createEvent("MouseEvents");
      evt.initEvent("click", true, true);
      a.dispatchEvent(evt);

      a.remove();
      URL.revokeObjectURL(url);

      await notify("Export downloaded to your device.", "success");
      return { success: true, method: "anchor_download" };
    } catch (err) {
      await logDiagnostic("Anchor download failed", err);
    }

    // FALLBACK 1: OfficeRuntime storage
    try {
      if (OfficeRuntime?.storage) {
        await OfficeRuntime.storage.setItem("ProofWrite_export_" + filename, json);
        await notify("Export saved to Office runtime storage (fallback).", "warn");
        return { success: true, method: "runtime_storage" };
      }
    } catch (err) { await logDiagnostic("Export runtime_storage failed", err); }

    // FALLBACK 2: Render JSON into the UI panel
    try {
      if (lastReportEl) {
        lastReportEl.textContent = json;
        await notify("Export JSON displayed in the panel (fallback).", "warn");
        return { success: true, method: "ui_copy" };
      }
    } catch (err) { await logDiagnostic("Final export fallback failed", err); }

    await notify("Export failed completely.", "error");
    return { success: false, method: "none" };

  } catch (err) {
    await logDiagnostic("exportJson failed", err);
    await notify("Export failed. Check diagnostics.", "error");
    return { success: false, method: "error" };
  } finally {
    setTimeout(() => { exportJsonRunning = false; }, 350);
  }
}

async function handleExportButton(e) {
  if (e) {
    try { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); } catch (_) {}
  }
  if (isExporting || exportJsonRunning) {
    await notify("Export already running.", "warn");
    await logDiagnostic("Blocked duplicate export call");
    return;
  }

  isExporting = true;
  if (exportBtnEl) {
    exportBtnEl.disabled = true;
    exportBtnEl.textContent = "Exporting...";
  }

  try {
    if (isSessionActive) {
      await logDiagnostic("Ending active session before export");
      await endSession();
    }

    if (!sessions?.length) {
      await notify("No sessions to export.", "warn");
      return;
    }

    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    const filename = `proofwrite_sessions_${timestamp}.json`;

    const result = await exportJson(sessions, filename);
    if (!result.success) {
      await notify("Export failed. JSON shown in panel.", "error");
    }
  } catch (err) {
    await logDiagnostic("handleExportButton failed", err);
    await notify("Export error. See diagnostics.", "error");
  } finally {
    isExporting = false;
    if (exportBtnEl) {
      exportBtnEl.disabled = false;
      exportBtnEl.textContent = "Export Sessions & Generate Token";
    }
  }
}

/* ============ HUMAN SCORE ============ */
function applyLargePastePenalty(session, wordsAdded) {
  const penalty = Math.min(PENALTY_LARGE_PASTE_MAX, Math.floor((wordsAdded / 100) * PENALTY_LARGE_PASTE_MAX));
  session.humanScore = Math.max(0, (session.humanScore || 100) - penalty);
  session._penalties = session._penalties || [];
  session._penalties.push({ type: "largePaste", wordsAdded, penalty });
}
function applySpeedSpikePenalty(session, wordsAdded) {
  const penalty = Math.min(PENALTY_SPEED_SPIKE_MAX, Math.floor((wordsAdded / 50) * PENALTY_SPEED_SPIKE_MAX));
  session.humanScore = Math.max(0, (session.humanScore || 100) - penalty);
  session._penalties = session._penalties || [];
  session._penalties.push({ type: "speedSpike", wordsAdded, penalty });
}
function applySustainedSpeedPenalty(session, wpm) {
  session.humanScore = Math.max(0, (session.humanScore || 100) - PENALTY_SUSTAINED_SPEED);
  session._penalties = session._penalties || [];
  session._penalties.push({ type: "sustainedSpeed", wpm, penalty: PENALTY_SUSTAINED_SPEED });
}

/* ============ UI HELPERS ============ */
function updateUiCounts(elapsedSec, charactersTyped) {
  if (sessionTimeEl) sessionTimeEl.textContent = `Session Time: ${elapsedSec}s`;
  if (charCountEl) charCountEl.textContent = `Characters Typed: ${charactersTyped}`;
  else if (wordCountEl) wordCountEl.textContent = `Characters Typed: ${charactersTyped}`;
}
function showAutoStartBanner() {
  if (!autoBannerEl) return;
  autoBannerEl.style.display = "block";
  setTimeout(() => { try { autoBannerEl.style.display = "none"; } catch (e) {} }, 3000);
}
function sanitizeSessionsList(arr) {
  return (Array.isArray(arr) ? arr.filter(s => s != null) : []);
}

/* ============ FLAGGED SUMMARY / CLICKABLE SESSIONS ============ */
function updateFlaggedSummary() {
  if (!flaggedSummaryEl) return;
  sessions = sanitizeSessionsList(sessions);
  if (!sessions.length) {
    flaggedSummaryEl.innerHTML = "<i>No sessions yet.</i>";
    return;
  }

  const reversedSessions = [...sessions].reverse();

  flaggedSummaryEl.innerHTML = reversedSessions.map((s, i) => {
    // compute original index in sessions array
    const originalIndex = sessions.length - 1 - i;
    const start = new Date(s.startTime);
    const end = new Date(s.endTime || Date.now());
    const activeMs = end - start;
    const activeMin = Math.floor(activeMs / 60000);
    const activeSec = Math.floor((activeMs % 60000) / 1000);
    const flags = [];
    if (s.flags?.largePaste) flags.push("Large Paste");
    if (s.flags?.speedSpike) flags.push("Speed Spike");
    if (s.flags?.sustainedSpeed) flags.push("Sustained Speed");

    const flagHtml = flags.length > 0
      ? `<span style="color:#d93025; font-weight:bold;">${flags.join(", ")}</span>`
      : `<span style="color:#188038; font-weight:600;">Clean</span>`;

    return `
      <div class="session-item-wrapper" style="margin-bottom:10px; padding:10px; border-left: 4px solid ${flags.length > 0 ? '#d93025' : '#188038'}; background:white; box-shadow:0 1px 3px rgba(0,0,0,0.1); border-radius:4px;">
        <div style="display:flex; justify-content:space-between; margin-bottom:4px;">
          <b style="font-size:14px;">Session ${originalIndex + 1}</b>
          <span style="font-size:12px; color:#666;">${formatTimestamp(start)}</span>
        </div>
        <div style="font-size:13px; color:#444; line-height:1.5;">
          Duration: ${activeMin}m ${activeSec}s <br>
          Characters: ${s.charactersTyped || 0} <br>
          Flags: ${flagHtml}
        </div>
        <div style="margin-top:8px; display:flex; gap:8px;">
          <button class="session-item" data-idx="${originalIndex}" style="padding:6px 10px; border-radius:6px; border:none; background:#0078d7; color:#fff; cursor:pointer;">Show graphs</button>
          <button class="session-download" data-idx="${originalIndex}" style="padding:6px 10px; border-radius:6px; border:1px solid #0078d7; background:#fff; color:#0078d7; cursor:pointer;">Download JSON</button>
        </div>
      </div>
    `;
  }).join("");

  // attach handler for download buttons (delegation will capture them via flaggedSummaryEl listener)
  // we rely on the delegated click listener set up in Office.onReady
}

/* ============ SESSION CONTROL ============ */
async function updateSessionMetrics(wcNow, charsNow) {
  if (!currentSession) return;
  const now = Date.now();
  const elapsed = Math.floor((now - sessionStartTime) / 1000);
  const diff = wcNow - lastWordCount;
  const totalDiff = wcNow - initialWordCount;
  const charDiff = charsNow - (currentSession.charactersTyped || 0);

  if (diff !== 0 || charDiff !== 0) lastActivityTime = now;

  if ((now - lastActivityTime) >= AUTO_STOP_THRESHOLD_MS) {
    await logDiagnostic("Auto-stop triggered due to inactivity", { inactivityMs: now - lastActivityTime });
    await endSession();
    return;
  }

  currentSession.charactersTyped = charsNow;
  updateUiCounts(elapsed, charsNow);

  // store event (timestamp + char count) — used by graphs
  currentSession.events = currentSession.events || [];
  currentSession.events.push({ t: now, c: charsNow });

  

  recentWordCounts.push({ timestamp: now, count: wcNow });
  const thresholdTime = now - SUSTAINED_SPEED_PERIOD_MS;
  recentWordCounts = recentWordCounts.filter(entry => entry.timestamp >= thresholdTime);

  if (recentWordCounts.length > 1) {
    const oldestEntry = recentWordCounts[0];
    const timeDeltaMs = now - oldestEntry.timestamp;
    const wordDelta = wcNow - oldestEntry.count;
    if (timeDeltaMs > 10000 && wordDelta > 0) {
      const activeWPM = (wordDelta / timeDeltaMs) * 60000;
      if (activeWPM >= SUSTAINED_SPEED_WPM && !currentSession.flags.sustainedSpeed) {
        currentSession.flags.sustainedSpeed = true;
        applySustainedSpeedPenalty(currentSession, Math.round(activeWPM));
        currentSession.edits.push({
          timestamp: new Date().toISOString(),
          wordCount: wcNow,
          wpm: Math.round(activeWPM),
          flag: "sustainedSpeed"
        });
        await logDiagnostic("Detected sustainedSpeed", { activeWPM, wordDelta, timeDeltaMs });
      }
    }
  }

  if (totalDiff >= THRESHOLD_LARGE_PASTE_WORDS && !currentSession.flags.largePaste) {
    currentSession.flags.largePaste = true;
    const words = Math.min(totalDiff, 10000);
    applyLargePastePenalty(currentSession, words);
    currentSession.edits.push({ timestamp: new Date().toISOString(), wordCount: wcNow, wordsAdded: words, flag: "largePaste" });
    await logDiagnostic("Detected largePaste", { totalDiff, wcNow });
  }

  if (diff >= THRESHOLD_SPEED_SPIKE_WORDS) {
    currentSession.flags.speedSpike = true;
    const words = Math.min(diff, 10000);
    applySpeedSpikePenalty(currentSession, words);
    currentSession.edits.push({ timestamp: new Date().toISOString(), wordCount: wcNow, wordsAdded: words, flag: "speedSpike" });
    await logDiagnostic("Detected speedSpike", { diff, wcNow });
  }

  lastWordCount = wcNow;
}

async function endSession() {
  if (!isSessionActive || !currentSession) return;
  try {
    const endTime = Date.now();
    currentSession.endTime = new Date(endTime).toISOString();
    currentSession.endTimestamp = endTime;
    currentSession.finalWordCount = await getDocumentWordCount();
    const txt = await getDocumentText();
    currentSession.charactersTyped = txt ? txt.length : (currentSession.charactersTyped || 0);

    // ensure event for end
    currentSession.events = currentSession.events || [];
    currentSession.events.push({ t: Date.now(), c: currentSession.charactersTyped });

    sessions.push(currentSession);

    try { if (lastReportEl) lastReportEl.textContent = JSON.stringify(currentSession, null, 2); } catch (e) { /* ignore */ }

    updateFlaggedSummary();
    updateUiCounts(0, currentSession.charactersTyped || 0);
    await storeSessionsLocally();

    clearInterval(uiTimerId);
    uiTimerId = null;
    isSessionActive = false;
    currentSession = null;
    recentWordCounts = [];

    await logDiagnostic("Session ended", { sessionsCount: sessions.length });
  } catch (err) { await logDiagnostic("endSession failed", err); }
}

async function startSession(auto = false, forcedInitialCount = null) {
  if (isSessionActive) return;
  try {
    const wc = forcedInitialCount ?? await getDocumentWordCount();
    initialWordCount = wc;
    lastWordCount = wc;
    sessionStartTime = Date.now();
    isSessionActive = true;
    lastActivityTime = Date.now();
    recentWordCounts = [{ timestamp: sessionStartTime, count: wc }];

    currentSession = {
      startTime: new Date(sessionStartTime).toISOString(),
      startTimestamp: sessionStartTime,
      initialWordCount: wc,
      edits: [],
      flags: { largePaste: false, speedSpike: false, sustainedSpeed: false },
      humanScore: 100,
      charactersTyped: 0,
      events: [{ t: sessionStartTime, c: 0 }] // start event
    };

    if (auto) showAutoStartBanner();
    updateUiCounts(0, 0);

    uiTimerId = setInterval(async () => {
      if (!isSessionActive) return;
      try {
        const txt = await getDocumentText();
        const wcNow = txt === null ? lastWordCount : (txt.trim() ? txt.trim().split(/\s+/).filter(w => w.length > 0).length : 0);
        const charsNow = txt === null ? (currentSession?.charactersTyped || 0) : txt.length;
        await updateSessionMetrics(wcNow, charsNow);
      } catch (err) {
        await logDiagnostic("uiTimer interval failed", err);
      }
    }, UI_TIMER_MS);
  } catch (err) { await logDiagnostic("startSession failed", err); }
}

/* ============ AUTO MONITOR ============ */
async function startMonitoring() {
  if (monitoringTimerId) return;
  try {
    await loadSessionsFromStorage();
    updateFlaggedSummary();

    let prevCount = await getDocumentWordCount();
    lastWordCount = prevCount;

    monitoringTimerId = setInterval(async () => {
      try {
        const wc = await getDocumentWordCount();
        const diff = wc - prevCount;

        if (verificationToken && diff > 0) {
          await clearVerificationToken();
        }

        if (!isSessionActive && diff > 0) {
          await logDiagnostic("Auto-start typing detected", { prevCount, wc, diff });
          await startSession(true, prevCount);
          prevCount = await getDocumentWordCount();
          lastWordCount = prevCount;
          return;
        }

        if (!isSessionActive && diff >= THRESHOLD_AUTO_START_JUMP) {
          await logDiagnostic("Auto-start large jump", { prevCount, wc, diff });
          await startSession(true, prevCount);
          prevCount = await getDocumentWordCount();
          lastWordCount = prevCount;
          return;
        }

        prevCount = wc;
        lastWordCount = wc;
      } catch (err) { await logDiagnostic("monitoring interval error", err); }
    }, MONITOR_POLL_MS);
  } catch (err) { await logDiagnostic("startMonitoring failed", err); }
}

/* ============ Clipboard helper (used by copy token) ============ */
async function copyToClipboardSafe(text) {
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch (err) {
    await logDiagnostic("navigator.clipboard failed", err);
  }
  try {
    const temp = document.createElement("textarea");
    temp.value = text;
    temp.style.position = "fixed";
    temp.style.opacity = "0";
    document.body.appendChild(temp);
    temp.select();
    const success = document.execCommand("copy");
    temp.remove();
    if (success) return true;
  } catch (err) {
    await logDiagnostic("execCommand copy failed", err);
  }
  return false;
}