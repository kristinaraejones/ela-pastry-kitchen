// The ELA Pastry Kitchen — frontend
// Ports the validated mockup's rendering/interaction logic against the
// Apps Script + Google Sheet API instead of localStorage. See
// apps-script/Code.gs for the API contract this file talks to.

const SUBJECT_ORDER = ["vocab", "spelling", "grammar", "reading", "writing"];
const CHILD_META = {
  kenley: { name: "Kenley", subtitle: "GRADE 7 · ELA · FALL TERM" },
  adelyn: { name: "Adelyn", subtitle: "GRADE 4 · ELA · FALL TERM" }
};
const TAG_ABBREV = {
  "Prepositional": "PREP", "Appositive": "APP",
  "Subject": "S", "Verb": "V", "Direct Object": "DO", "Indirect Object": "IO", "Subject Complement": "SC",
  "Action Verb": "AV", "Linking Verb": "LV"
};

let API_URL = "";
let currentChild = "kenley";
let currentView = "kenley";
let openStation = null;
let posPopupOpen = null;
let conceptPopupOpen = null;
let phraseRangeStart = null;
let pendingPhraseRange = null;
let parentNavWeek = null;      // parent-only week browser; null = not yet landed on current week
let exportControlsReady = false; // guards one-time default-fill of the export range inputs

let DATA = null;   // current child's { subjectKey: {name, tag, tasks:[...]} }
let state = null;  // current child's { subjectKey: { tasks: { taskId: {...} } } }
let settings = { weeks: { kenley: 1, adelyn: 1 }, termFinalsUnlocked: false, monthlyTestOverride: null };

const childrenCache = {};   // { kenley: {DATA, state} }
const reviewPoolCache = {}; // { kenley: [ {word,timesMissed,...} ] }
const answerLogCache = {}; // { kenley: [ {timestamp,game,subject,word,question,givenAnswer,correctAnswer,correct} ] }
const markersCache = {};    // { kenley: {vocab: 3, ...} }
const burnLogCache = {};    // { kenley: [ {...} ] }
const banksCache = {};      // { kenley: {vocab:[...], spelling:[...], ...} }

// ---------- API layer ----------

function resolveApiUrl() {
  const stored = localStorage.getItem("elaApiUrl");
  if (stored) return stored;
  if (typeof DEFAULT_API_URL !== "undefined" && DEFAULT_API_URL) {
    // Persist it on first successful use so every load after this one is
    // independent of config.js loading/executing correctly at all — a
    // device that saw it once can never hit the "connect the kitchen"
    // screen again, regardless of what caused config.js to come back
    // empty that one time.
    try { localStorage.setItem("elaApiUrl", DEFAULT_API_URL); } catch (e) {}
    return DEFAULT_API_URL;
  }
  return "";
}

const BOOTSTRAP_CACHE_PREFIX = "elaBootstrapCache_";

async function apiGetBootstrap(student) {
  try {
    const res = await fetch(`${API_URL}?action=bootstrap&student=${encodeURIComponent(student)}`);
    const json = await res.json();
    if (!json.ok) throw new Error(json.error || "Bootstrap failed");
    try { localStorage.setItem(BOOTSTRAP_CACHE_PREFIX + student, JSON.stringify(json)); } catch (e) {}
    return json;
  } catch (err) {
    const cached = localStorage.getItem(BOOTSTRAP_CACHE_PREFIX + student);
    if (!cached) throw err;
    const parsed = JSON.parse(cached);
    parsed._offline = true;
    applyQueuedWritesToBootstrap(parsed, student);
    return parsed;
  }
}

// Reconciles the last-known-good bootstrap snapshot with any writes still
// sitting in the offline queue, so reopening the app offline (after already
// having worked offline once) doesn't show stale/reverted task state.
function applyQueuedWritesToBootstrap(resp, student) {
  const queue = loadWriteQueue();
  if (queue.length === 0) return;
  const subByTaskId = {};
  (resp.submissions || []).forEach(s => { subByTaskId[s.task_id] = s; });
  resp.settings = resp.settings || {};
  queue.forEach(item => {
    if (item.action === "saveSubmission" && item.payload.student === student) {
      subByTaskId[item.payload.task_id] = { ...subByTaskId[item.payload.task_id], ...item.payload };
    } else if (item.action === "saveSetting") {
      resp.settings[item.payload.key] = item.payload.value;
    }
  });
  resp.submissions = Object.values(subByTaskId);
}

// ---------- Offline write queue ----------
// apiPost's callers already update local `state`/`settings` synchronously
// before calling it — every call site is fire-and-forget (`.catch(() => {})`).
// That means the UI is already optimistic, so going offline just means
// queuing the persistence call instead of losing it.

const WRITE_QUEUE_KEY = "elaWriteQueue";
let flushInFlight = false;

function loadWriteQueue() {
  try { return JSON.parse(localStorage.getItem(WRITE_QUEUE_KEY)) || []; }
  catch (e) { return []; }
}
function saveWriteQueue(q) {
  localStorage.setItem(WRITE_QUEUE_KEY, JSON.stringify(q));
}
function queueWrite(action, payload) {
  const q = loadWriteQueue();
  q.push({ action, payload, queuedAt: Date.now() });
  saveWriteQueue(q);
  updateOfflineBanner();
}

async function flushWriteQueue() {
  if (flushInFlight || !navigator.onLine) { updateOfflineBanner(); return; }
  flushInFlight = true;
  try {
    while (true) {
      const q = loadWriteQueue();
      if (q.length === 0) break;
      const item = q[0];
      try {
        const res = await fetch(API_URL, { method: "POST", body: JSON.stringify({ action: item.action, ...item.payload }) });
        const json = await res.json();
        if (!json.ok) console.error("Dropped queued write (server rejected):", item, json.error);
      } catch (err) {
        break; // still unreachable — stop draining, leave the rest queued
      }
      const remaining = loadWriteQueue();
      remaining.shift();
      saveWriteQueue(remaining);
    }
  } finally {
    flushInFlight = false;
    updateOfflineBanner();
  }
}

function updateOfflineBanner() {
  const el = document.getElementById("syncBanner");
  if (!el) return;
  const q = loadWriteQueue();
  if (!navigator.onLine) {
    el.textContent = q.length > 0
      ? `📴 Offline — working from saved lesson data. ${q.length} change${q.length === 1 ? "" : "s"} will sync once you're back online.`
      : "📴 Offline — working from saved lesson data.";
    el.style.display = "block";
  } else if (q.length > 0) {
    el.textContent = `🔄 Syncing ${q.length} saved change${q.length === 1 ? "" : "s"}…`;
    el.style.display = "block";
  } else {
    el.style.display = "none";
  }
}

window.addEventListener("online", flushWriteQueue);
window.addEventListener("offline", updateOfflineBanner);

async function apiPost(action, payload) {
  if (!navigator.onLine) {
    queueWrite(action, payload);
    return { ok: true, queued: true };
  }
  try {
    const res = await fetch(API_URL, { method: "POST", body: JSON.stringify({ action, ...payload }) });
    const json = await res.json();
    if (!json.ok) throw new Error(json.error || "Request failed");
    clearSyncError();
    return json;
  } catch (err) {
    if (err instanceof TypeError) {
      // fetch couldn't even complete (offline/unreachable) — queue instead of losing it
      queueWrite(action, payload);
      return { ok: true, queued: true };
    }
    console.error("apiPost failed:", action, err);
    showSyncError("Couldn't save your last update — check the connection and try again.");
    throw err;
  }
}

function showSyncError(msg) {
  const el = document.getElementById("syncBanner");
  if (!el) return;
  el.textContent = "⚠️ " + msg;
  el.style.display = "block";
}
function clearSyncError() {
  updateOfflineBanner();
}

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("sw.js").catch(err => console.error("SW registration failed:", err));
  });
}

// ---------- Setup / connect screen ----------

function showSetupOverlay(errorMsg) {
  document.getElementById("loadingNote").style.display = "none";
  document.getElementById("board").style.display = "none";
  const overlay = document.getElementById("setupOverlay");
  overlay.style.display = "flex";
  const errEl = document.getElementById("setupError");
  if (errorMsg) {
    errEl.textContent = errorMsg;
    errEl.style.display = "block";
  } else {
    errEl.style.display = "none";
  }
}

function submitSetupUrl() {
  const val = document.getElementById("setupUrlInput").value.trim();
  if (!val.startsWith("http")) {
    document.getElementById("setupError").textContent = "That doesn't look like a URL — paste the full Web App URL from Apps Script.";
    document.getElementById("setupError").style.display = "block";
    return;
  }
  localStorage.setItem("elaApiUrl", val);
  document.getElementById("setupOverlay").style.display = "none";
  document.getElementById("loadingNote").style.display = "block";
  init();
}

// ---------- Boot ----------

async function init() {
  API_URL = resolveApiUrl();
  if (!API_URL) {
    showSetupOverlay();
    return;
  }
  try {
    await loadChild("kenley");
    settings = parseSettings(childrenCache._settings);
    currentChild = "kenley";
    DATA = childrenCache.kenley.DATA;
    state = childrenCache.kenley.state;
    document.getElementById("loadingNote").style.display = "none";
    document.getElementById("board").style.display = "block";
    render();
    updateOfflineBanner();
    flushWriteQueue();
  } catch (err) {
    console.error(err);
    showSetupOverlay("Couldn't reach that URL: " + err.message + ". Double-check it and try again.");
  }
}

async function loadChild(student) {
  const resp = await apiGetBootstrap(student);
  childrenCache._settings = resp.settings; // shared/global, same on every bootstrap call
  const built = buildChildFromBootstrap(resp);
  childrenCache[student] = built;
  reviewPoolCache[student] = resp.reviewPool;
  answerLogCache[student] = resp.answerLog || [];
  markersCache[student] = resp.markers;
  burnLogCache[student] = resp.burnLog;
  banksCache[student] = resp.banks;
}

function parseSettings(raw) {
  raw = raw || {};
  const override = raw.monthlyTestOverride === "true" ? true : raw.monthlyTestOverride === "false" ? false : null;
  const weeks = {};
  Object.keys(CHILD_META).forEach(id => { weeks[id] = Number(raw[`${id}_current_week`]) || 1; });
  return {
    weeks,
    termFinalsUnlocked: raw.termFinalsUnlocked === "true" || raw.termFinalsUnlocked === true,
    monthlyTestOverride: override
  };
}

function buildChildFromBootstrap(resp) {
  const DATA = {};
  SUBJECT_ORDER.forEach(subjectKey => {
    const tasksForSubject = resp.schedule.filter(t => t.subject_key === subjectKey);
    if (tasksForSubject.length === 0) return;
    const tagsByWeek = {};
    tasksForSubject.forEach(t => { tagsByWeek[Number(t.week_number) || 1] = t.subject_tag; });
    DATA[subjectKey] = {
      name: tasksForSubject[0].subject_name,
      tag: tasksForSubject[0].subject_tag, // fallback for weeks with no dedicated tag (e.g. Adelyn's single-week placeholders)
      tagsByWeek,
      tasks: tasksForSubject.map(t => Object.assign(
        { id: t.id, type: t.type, label: t.label, dynamic: t.dynamic, termFinal: t.termFinal, monthlyTest: t.monthlyTest, week_number: Number(t.week_number) || 1 },
        t.content || {}
      ))
    };
  });

  const subMap = {};
  resp.submissions.forEach(s => { subMap[s.task_id] = s; });

  const state = {};
  Object.keys(DATA).forEach(key => {
    state[key] = { tasks: {} };
    DATA[key].tasks.forEach(t => {
      const sub = subMap[t.id];
      const a = (sub && sub.answers) || {};
      const base = {
        open: false,
        done: sub ? ["complete", "needs_review", "reviewed"].includes(sub.status) : false,
        needsReview: sub ? sub.status === "needs_review" : false,
        reviewed: sub ? sub.status === "reviewed" : false,
        sentBack: sub ? sub.status === "sent_back" : false,
        answers: a.answers || {},
        score: (sub && sub.score) || null,
        parentComment: (sub && sub.parent_comment) || null,
        results: a.results || null
      };
      if (t.type === "pos-tagger") base.labels = a.labels || new Array(t.sentence.length).fill(null);
      if (t.type === "phrase-tagger") base.selections = a.selections || [];
      if (t.type === "concept-check") base.labels = a.labels || {};
      state[key].tasks[t.id] = base;
    });
  });
  return { DATA, state };
}

// A blank, not-started state for one task (right shape for its type).
function freshTaskState(t) {
  const base = { open: false, done: false, needsReview: false, reviewed: false, sentBack: false, answers: {}, score: null, parentComment: null, results: null };
  if (t.type === "pos-tagger") base.labels = new Array(t.sentence.length).fill(null);
  if (t.type === "phrase-tagger") base.selections = [];
  if (t.type === "concept-check") base.labels = {};
  return base;
}

function deriveStatus(s) {
  if (s.sentBack && !s.done) return "sent_back";
  if (s.reviewed) return "reviewed";
  if (s.needsReview) return "needs_review";
  if (s.done) return "complete";
  return "not_started";
}

// Human-readable version of deriveStatus, shared by the past-week report and the PDF export.
function statusLabel(s) {
  if (s.sentBack && !s.done) return "Sent back — awaiting resubmission";
  if (s.reviewed) return "Reviewed & approved";
  if (s.needsReview) return "Submitted — awaiting review";
  if (s.done) return "Complete";
  return "Not started";
}

function persistTask(key, id) {
  const s = state[key].tasks[id];
  apiPost("saveSubmission", {
    student: currentChild,
    task_id: id,
    status: deriveStatus(s),
    score: s.score || "",
    parent_comment: s.parentComment || "",
    answers: { answers: s.answers, labels: s.labels, selections: s.selections, results: s.results }
  }).catch(() => {});
}

async function switchChild(id) {
  if (id === currentChild) return;
  currentChild = id;
  if (!childrenCache[id]) {
    document.getElementById("loadingNote").style.display = "block";
    document.getElementById("board").style.display = "none";
    try {
      await loadChild(id);
    } catch (err) {
      showSetupOverlay("Couldn't load " + CHILD_META[id].name + "'s kitchen: " + err.message);
      return;
    }
    document.getElementById("loadingNote").style.display = "none";
    document.getElementById("board").style.display = "block";
  }
  DATA = childrenCache[id].DATA;
  state = childrenCache[id].state;
  openStation = null; posPopupOpen = null; conceptPopupOpen = null; phraseRangeStart = null; pendingPhraseRange = null;
  parentNavWeek = null; exportControlsReady = false;
  render();
}

// ---------- Settings: per-student current_week (completion-based, no dates) ----------

function currentWeek() { return settings.weeks[currentChild] || 1; }
// The subject's tag for the CURRENT active week (e.g. "AAS Level 7 · Step 8"
// once she's on week 2) — falls back to whatever tag the subject's first
// row carries, for subjects/weeks with no dedicated per-week tag.
function subjectTag(key) { return DATA[key].tagsByWeek[currentWeek()] || DATA[key].tag; }
function isMonthlyTestWeek() { return currentWeek() % 4 === 0; }
function nextMonthlyTestWeek() {
  const w = currentWeek();
  return isMonthlyTestWeek() ? w : w + (4 - (w % 4));
}
function pendingReviewCount() {
  let n = 0;
  Object.keys(DATA).forEach(key => DATA[key].tasks.forEach(t => {
    const s = state[key].tasks[t.id];
    if (s.needsReview && !s.reviewed) n++;
  }));
  return n;
}
function advanceAnyway() {
  if (!confirm(`Week ${currentWeek()} isn't fully finished. Advance ${CHILD_META[currentChild].name} to Week ${currentWeek() + 1} anyway?`)) return;
  advanceWeek();
}
function advanceWeek() {
  settings.weeks[currentChild] = currentWeek() + 1;
  apiPost("saveSetting", { key: `${currentChild}_current_week`, value: String(settings.weeks[currentChild]) }).catch(() => {});
  render();
}
function toggleTermFinals() {
  settings.termFinalsUnlocked = !settings.termFinalsUnlocked;
  apiPost("saveSetting", { key: "termFinalsUnlocked", value: String(settings.termFinalsUnlocked) }).catch(() => {});
  render();
}
function toggleMonthlyTests() {
  if (settings.monthlyTestOverride === null) settings.monthlyTestOverride = false;
  else if (settings.monthlyTestOverride === false) settings.monthlyTestOverride = true;
  else settings.monthlyTestOverride = null;
  apiPost("saveSetting", { key: "monthlyTestOverride", value: settings.monthlyTestOverride === null ? "" : String(settings.monthlyTestOverride) }).catch(() => {});
  render();
}
function isTaskLocked(t) {
  if (t.termFinal) return !settings.termFinalsUnlocked;
  if (t.monthlyTest) {
    if (settings.monthlyTestOverride !== null) return settings.monthlyTestOverride;
    return !isMonthlyTestWeek();
  }
  return false;
}

// ---------- Station status ----------

// Tasks belonging to the currently-active week for this student, plus any
// "unlocked whenever" bank-driven task (monthly tests, term finals, and the
// review-pool dictation drill) that isn't tied to any one week — matching
// the data model's own "one row per week, or per 'unlocked whenever' for
// banks" description. Past/future week content just isn't part of this set
// at all — it's not "locked," it's simply not this week's work.
function activeTasks(key) {
  const week = currentWeek();
  return DATA[key].tasks.filter(t => t.dynamic || t.week_number === week);
}
function unlockedActiveTasks(key) {
  return activeTasks(key).filter(t => !isTaskLocked(t));
}
function stationScorePct(key) {
  let sumX = 0, sumY = 0;
  unlockedActiveTasks(key).forEach(t => {
    const s = state[key].tasks[t.id];
    if (s.done && s.score) {
      const parts = s.score.split("/").map(Number);
      sumX += parts[0]; sumY += parts[1];
    }
  });
  return sumY > 0 ? sumX / sumY : null;
}
function stationDone(key) {
  const u = unlockedActiveTasks(key);
  return u.length > 0 && u.every(t => state[key].tasks[t.id].done);
}
// "empty" = nothing loaded for this student's current week yet (distinct
// from "progress" so it never counts as done and never blocks/enables
// Advance-to-next-week by accident — see allSubjectsServed()).
function stationStatus(key) {
  const u = unlockedActiveTasks(key);
  if (u.length === 0) return "empty";
  if (u.every(t => state[key].tasks[t.id].done)) {
    const pct = stationScorePct(key);
    if (pct !== null && pct < 0.70) return "burning";
    return "served";
  }
  return "progress";
}
function allSubjectsServed() {
  return Object.keys(DATA).every(key => stationStatus(key) === "served");
}

// ---------- Parent-only week navigator (past/current/upcoming) ----------

// Bank-driven "unlocked whenever" tasks (monthly tests, term finals, the
// review-pool drill) are excluded from week-scoped views: they're not tied
// to one week, and Submissions only ever keeps the latest attempt (no
// per-attempt history) — so attributing a retaken one's current score to
// whichever week it happened to be authored under would misrepresent that
// week's actual record. Their live status still shows normally in the
// current-week dashboard, unchanged.
function weekScopedTasks(key, week) {
  return DATA[key] ? DATA[key].tasks.filter(t => t.week_number === week && !t.dynamic) : [];
}
function maxAuthoredWeek() {
  let max = currentWeek();
  Object.keys(DATA).forEach(key => {
    DATA[key].tasks.forEach(t => {
      if (!t.dynamic && t.week_number > max) max = t.week_number;
    });
  });
  return max;
}
function jumpToReview(key, id, week) {
  parentNavWeek = week;
  render();
  const el = document.getElementById(`review-${key}-${id}`);
  if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
}
function navPrevWeek() {
  parentNavWeek = Math.max(1, parentNavWeek - 1);
  render();
}
function navNextWeek() {
  parentNavWeek = Math.min(maxAuthoredWeek(), parentNavWeek + 1);
  render();
}
function redoStation(key) {
  const items = activeTasks(key).map(t => {
    const s = state[key].tasks[t.id];
    if (s.done && s.score) return `${t.label}: scored ${s.score}`;
    if (s.done && s.needsReview) return `${t.label}: submitted, was awaiting your review`;
    if (s.done) return `${t.label}: completed`;
    return `${t.label}: not completed`;
  });
  const record = {
    station: DATA[key].name, tag: subjectTag(key),
    date: new Date().toLocaleDateString(),
    reason: `Scored below 70% (${Math.round((stationScorePct(key) || 0) * 100)}%)`,
    items
  };
  (burnLogCache[currentChild] = burnLogCache[currentChild] || []).unshift(record);
  apiPost("addBurnLog", { student: currentChild, ...record }).catch(() => {});
  activeTasks(key).forEach(t => {
    state[key].tasks[t.id] = freshTaskState(t);
    persistTask(key, t.id);
  });
  render();
}

// ---------- Reset / send back ONE section (task), not the whole subject ----------

// Reflections keep the written text so it can be edited rather than retyped.
function resetOneTask(key, id, by, comment) {
  const t = DATA[key].tasks.find(x => x.id === id);
  const old = state[key].tasks[id];
  const was = old.score ? `scored ${old.score}` : old.needsReview ? "submitted, awaiting review" : old.done ? "completed" : "not finished";
  const record = {
    station: DATA[key].name, tag: subjectTag(key), date: new Date().toLocaleDateString(),
    reason: by === "parent" ? "Section sent back by parent" : "Section redone by student to improve it",
    items: [`${t.label}: was ${was}${comment ? ` — note: ${comment}` : ""}`]
  };
  (burnLogCache[currentChild] = burnLogCache[currentChild] || []).unshift(record);
  apiPost("addBurnLog", { student: currentChild, ...record }).catch(() => {});
  const fresh = freshTaskState(t);
  if (t.type === "reflection" && old.answers) fresh.answers = old.answers;
  fresh.open = by !== "parent";
  if (by === "parent") { fresh.sentBack = true; fresh.parentComment = comment || "Please take another look at this part and resubmit."; }
  state[key].tasks[id] = fresh;
  persistTask(key, id);
  render();
}
function sendBackSection(key, id) {
  const box = document.getElementById(`sendback-${key}-${id}`);
  resetOneTask(key, id, "parent", box ? box.value.trim() : "");
}
function redoSection(key, id) {
  if (!confirm("Start this section over? Your current answers here will be cleared so you can try again.")) return;
  resetOneTask(key, id, "student", "");
}
// Students can redo any finished section to improve it, except read-only lesson
// steps and sampled tests (a redo would just re-roll the questions).
function canStudentRedo(t, s) {
  return currentView !== "parent" && s.done && !t.dynamic && !t.termFinal && !t.monthlyTest && t.type !== "read" && t.type !== "external";
}

function toggleView() {
  currentView = currentView === "kenley" ? "parent" : "kenley";
  if (currentView !== "parent") { parentNavWeek = null; exportControlsReady = false; } // re-land on current week next time parent view opens
  render();
}
function openStationFn(key) {
  openStation = openStation === key ? null : key;
  render();
}
function toggleTask(key, id) {
  const t = DATA[key].tasks.find(x => x.id === id);
  if (isTaskLocked(t)) return;
  state[key].tasks[id].open = !state[key].tasks[id].open;
  render();
}

// ---------- Task actions ----------

function markRead(key, id) {
  state[key].tasks[id].done = true;
  persistTask(key, id);
  render();
}
function markExternal(key, id, checked) {
  state[key].tasks[id].done = checked;
  persistTask(key, id);
  render();
}
function submitReflection(key, id) {
  const val = document.getElementById(`ta-${key}-${id}`).value.trim();
  if (!val) return;
  state[key].tasks[id].answers.text = val;
  state[key].tasks[id].done = true;
  state[key].tasks[id].needsReview = true;
  state[key].tasks[id].sentBack = false;
  persistTask(key, id);
  render();
}

// ---------- Review pool (missed-word bank) ----------

function loadPool() { return reviewPoolCache[currentChild] || []; }
function getReviewWords(n) {
  const pool = loadPool().filter(p => p.status === "active");
  pool.sort((a, b) => b.timesMissed - a.timesMissed);
  return pool.slice(0, n).map(p => ({ answer: p.word, kind: "word", context: p.context || null }));
}
function persistReviewWord(entry) {
  apiPost("saveReviewWord", {
    student: currentChild, word: entry.word, timesMissed: entry.timesMissed,
    timesCorrect: entry.timesCorrect, lastSeen: entry.lastSeen, status: entry.status, context: entry.context || ""
  }).catch(() => {});
}
function logMiss(word, source, context) {
  const pool = loadPool();
  const existing = pool.find(p => p.word.toLowerCase() === word.toLowerCase());
  if (existing) {
    existing.timesMissed++; existing.timesCorrect = 0; existing.lastSeen = source; existing.status = "active";
    if (context) existing.context = context;
    persistReviewWord(existing);
  } else {
    const entry = { word, timesMissed: 1, timesCorrect: 0, lastSeen: source, status: "active", context: context || null };
    pool.push(entry);
    persistReviewWord(entry);
  }
}
function logReviewResult(word, correct) {
  const pool = loadPool();
  const existing = pool.find(p => p.word.toLowerCase() === word.toLowerCase());
  if (!existing) return;
  if (correct) { existing.timesCorrect++; if (existing.timesCorrect >= 2) existing.status = "mastered"; }
  else { existing.timesMissed++; existing.timesCorrect = 0; }
  persistReviewWord(existing);
}

// Logs one graded item from ANY subject/task type (MC, fill-in, pos-tagger,
// phrase-tagger, concept-check, dictation) into the same AnswerLog sheet the
// external vocab games (Case Files / Word Bakery) write to, so the parent
// view's answer history covers the whole curriculum, not just those two
// games. Best-effort — a failed/queued write never blocks grading.
function logGradedAnswer(key, { word, question, given, correct, wasCorrect }) {
  apiPost("logAnswer", {
    student: currentChild,
    game: "ELA Pastry Kitchen",
    subject: (DATA[key] && DATA[key].name) || key,
    word: word || "",
    question: question || "",
    given_answer: given == null ? "(no answer)" : given,
    correct_answer: correct,
    correct: wasCorrect
  }).catch(() => {});
}

// ---------- Monthly/term banks ----------

function getNewSinceLastTest(bank, subjectKey) {
  const marker = (markersCache[currentChild] && markersCache[currentChild][subjectKey]) || 0;
  const fresh = bank.slice(marker);
  return fresh.length > 0 ? fresh : bank;
}
function markTested(bank, subjectKey) {
  if (!markersCache[currentChild]) markersCache[currentChild] = {};
  markersCache[currentChild][subjectKey] = bank.length;
  apiPost("saveMarker", { student: currentChild, subject_key: subjectKey, bank_position: bank.length }).catch(() => {});
}
function sampleVocabQuestions(bank, n) {
  const pool = getNewSinceLastTest(bank, "vocab");
  return [...pool].sort(() => Math.random() - 0.5).slice(0, n).map(item => {
    const opts = Math.random() < 0.5 ? [item.correct, item.wrong] : [item.wrong, item.correct];
    return { q: `Which sentence uses "${item.word}" correctly?`, options: opts, correct: opts.indexOf(item.correct) };
  });
}
function sampleSpellingWords(bank, n) {
  const pool = getNewSinceLastTest(bank, "spelling");
  return [...pool].sort(() => Math.random() - 0.5).slice(0, n).map(w => ({ answer: w.word, kind: "word", context: w.context || null }));
}
function sampleGrammarQuestions(bank, n) {
  return [...getNewSinceLastTest(bank, "grammar")].sort(() => Math.random() - 0.5).slice(0, n);
}
function sampleReadingQuestions(bank, n) {
  return [...getNewSinceLastTest(bank, "reading")].sort(() => Math.random() - 0.5).slice(0, n);
}
function sampleFull(bank, n) { return [...bank].sort(() => Math.random() - 0.5).slice(0, n); }
function sampleExamVocabQuestions(bank, n) {
  return sampleFull(bank, n).map(item => {
    const opts = Math.random() < 0.5 ? [item.correct, item.wrong] : [item.wrong, item.correct];
    return { q: `Which sentence uses "${item.word}" correctly?`, options: opts, correct: opts.indexOf(item.correct) };
  });
}
function sampleExamGrammarQuestions(bank, n) { return sampleFull(bank, n); }
function sampleExamReadingQuestions(bank, n) { return sampleFull(bank, n); }
function sampleExamSpellingWords(bank, n) {
  return sampleFull(bank, n).map(w => ({ answer: w.word, kind: "word", context: w.context || null }));
}
function getDynamicBankConfig(dynamicKey) {
  const b = banksCache[currentChild] || {};
  const CONFIGS = {
    vocabMonthBank: { bank: b.vocab || [], subjectKey: "vocab", sample: sampleVocabQuestions, size: 10 },
    grammarMonthBank: { bank: b.grammar || [], subjectKey: "grammar", sample: sampleGrammarQuestions, size: 6 },
    readingMonthBank: { bank: b.reading || [], subjectKey: "reading", sample: sampleReadingQuestions, size: 6 },
    examVocabBank: { bank: b.vocab || [], sample: sampleExamVocabQuestions, size: 20, noMarker: true },
    examGrammarBank: { bank: b.grammar || [], sample: sampleExamGrammarQuestions, size: 12, noMarker: true },
    examReadingBank: { bank: b.reading || [], sample: sampleExamReadingQuestions, size: 12, noMarker: true }
  };
  return CONFIGS[dynamicKey];
}
function getTaskQuestions(key, id) {
  const t = DATA[key].tasks.find(x => x.id === id);
  if (t.dynamic) {
    const cfg = getDynamicBankConfig(t.dynamic);
    if (cfg) {
      if (!state[key].tasks[id]._questions) state[key].tasks[id]._questions = cfg.sample(cfg.bank, Math.min(cfg.size, cfg.bank.length));
      return state[key].tasks[id]._questions;
    }
  }
  return t.questions;
}

// ---------- Speech synthesis (dictation) ----------

// Prefers warmer-sounding built-in voices (Apple's Samantha on iPad/Mac,
// Google's default on Android/Chrome) over flatter ones, and otherwise
// falls back to any offline-capable (localService) English voice — never a
// network-only voice, since that would silently fail to read anything with
// no connection, defeating the point of the offline mode above.
let cachedVoice = null;
const PREFERRED_VOICE_NAMES = [
  "Samantha", "Google US English", "Microsoft Zira - English (United States)",
  "Karen", "Moira", "Tessa", "Google UK English Female"
];
function pickFriendlyVoice() {
  if (cachedVoice) return cachedVoice;
  const voices = window.speechSynthesis.getVoices();
  if (!voices.length) return null;
  for (const name of PREFERRED_VOICE_NAMES) {
    const v = voices.find(v => v.name === name);
    if (v) { cachedVoice = v; return v; }
  }
  const localEn = voices.find(v => v.localService && v.lang.startsWith("en"));
  if (localEn) { cachedVoice = localEn; return localEn; }
  return voices[0] || null;
}
if ("speechSynthesis" in window) {
  window.speechSynthesis.onvoiceschanged = () => { cachedVoice = null; };
}

function speakWord(word) {
  if (!("speechSynthesis" in window)) return;
  const u = new SpeechSynthesisUtterance(word);
  u.voice = pickFriendlyVoice();
  u.pitch = 1.05;
  u.rate = 0.85;
  window.speechSynthesis.cancel();
  window.speechSynthesis.speak(u);
}
function speakSequence(parts) {
  if (!("speechSynthesis" in window)) return;
  window.speechSynthesis.cancel();
  parts.forEach(text => {
    const u = new SpeechSynthesisUtterance(text);
    u.voice = pickFriendlyVoice();
    u.pitch = 1.05;
    u.rate = 0.85;
    window.speechSynthesis.speak(u);
  });
}
function speakWordInContext(word, context) { speakSequence([word, context, word]); }

// Reads aloud whatever text is currently in the DOM element with this id —
// used for "read this to me" buttons on lesson content and instructions,
// where the text is long/HTML-bearing and not safe to inline into an
// onclick attribute the way the short dictation words above are.
function speakElementText(elId) {
  if (!("speechSynthesis" in window)) return;
  const el = document.getElementById(elId);
  if (!el) return;
  const text = (el.innerText || el.textContent || "").trim();
  if (!text) return;
  const u = new SpeechSynthesisUtterance(text);
  u.voice = pickFriendlyVoice();
  u.pitch = 1.05;
  u.rate = 0.9;
  window.speechSynthesis.cancel();
  window.speechSynthesis.speak(u);
}
function readAloudButton(elId, label) {
  return `<button class="btn read-aloud-btn" onclick="speakElementText('${elId}')">🔊 ${label || "Read this to me"}</button>`;
}

// Aligns correct/typed word lists via LCS instead of raw position, so one
// extra/missing/split word (e.g. typing "milk man" for "milkman") doesn't
// shift every later word out of alignment and get them all flagged wrong.
function alignWords(correctWords, typedWords) {
  const n = correctWords.length, m = typedWords.length;
  const eq = (a, b) => a.toLowerCase() === b.toLowerCase();
  const dp = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] = eq(correctWords[i], typedWords[j]) ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }
  const ops = [];
  let i = 0, j = 0;
  while (i < n && j < m) {
    if (eq(correctWords[i], typedWords[j])) { ops.push({ correctWord: correctWords[i], typedWord: typedWords[j], match: true }); i++; j++; }
    else if (dp[i + 1][j] >= dp[i][j + 1]) { ops.push({ correctWord: correctWords[i], typedWord: "", match: false }); i++; }
    else { ops.push({ correctWord: "", typedWord: typedWords[j], match: false }); j++; }
  }
  while (i < n) { ops.push({ correctWord: correctWords[i], typedWord: "", match: false }); i++; }
  while (j < m) { ops.push({ correctWord: "", typedWord: typedWords[j], match: false }); j++; }
  // Merge an adjacent delete+insert (in either order) into one substitution,
  // so a plain misspelling still shows as a single chip like before.
  const merged = [];
  for (let k = 0; k < ops.length; k++) {
    const cur = ops[k], next = ops[k + 1];
    if (!cur.match && next && !next.match && (!cur.correctWord || !cur.typedWord) && (!next.correctWord || !next.typedWord)
      && (cur.correctWord || next.correctWord) && (cur.typedWord || next.typedWord) && !(cur.correctWord && next.correctWord) && !(cur.typedWord && next.typedWord)) {
      merged.push({ correctWord: cur.correctWord || next.correctWord, typedWord: cur.typedWord || next.typedWord, match: false });
      k++;
    } else {
      merged.push(cur);
    }
  }
  return merged;
}

const GRADE_VERSION = 2; // bump when grading rules change; older saved results get re-graded

// Spacing is not graded yet: a word typed as two ("news paper") or two words
// run together ("cannot" for "can not") counts as right. Looks for a short run
// of adjacent wrong ops whose letters, joined, match exactly, and marks them right.
function forgiveSpacing(ops) {
  const out = [];
  let i = 0;
  while (i < ops.length) {
    if (!ops[i].match) {
      let merged = false;
      for (let len = Math.min(4, ops.length - i); len >= 2 && !merged; len--) {
        const run = ops.slice(i, i + len);
        if (run.some(o => o.match)) continue;
        const c = run.map(o => o.correctWord).join("").toLowerCase();
        const t = run.map(o => o.typedWord).join("").toLowerCase();
        if (c && c === t) {
          out.push({ correctWord: run.map(o => o.correctWord).filter(Boolean).join(" "), typedWord: run.map(o => o.typedWord).filter(Boolean).join(" "), match: true });
          i += len; merged = true;
        }
      }
      if (merged) continue;
    }
    out.push(ops[i]); i++;
  }
  return out;
}

function gradeSentence(correctText, typedText) {
  const tokenize = s => s.match(/[A-Za-z']+/g) || [];
  const correctWords = tokenize(correctText);
  const typedWords = tokenize(typedText);
  const wordResults = forgiveSpacing(alignWords(correctWords, typedWords)).map(op => ({
    correctWord: op.correctWord, typedWord: op.typedWord, spellingCorrect: op.match
  }));
  const correctTrim = correctText.trim(), typedTrim = typedText.trim();
  const startsCapCorrect = /^[A-Z]/.test(correctTrim);
  const startsCapTyped = /^[A-Z]/.test(typedTrim);
  const endPunctCorrect = (correctTrim.match(/[.!?]$/) || [])[0] || "(none)";
  const endPunctTyped = (typedTrim.match(/[.!?]$/) || [])[0] || "(none)";
  // The dictation voice doesn't stress anything, so a child can't hear whether
  // a statement ends in "." or "!" — treat those two as interchangeable. A
  // question mark still has to match a question mark.
  const isStop = p => p === "." || p === "!";
  const punctPass = endPunctCorrect === endPunctTyped || (isStop(endPunctCorrect) && isStop(endPunctTyped));
  const punctLabel = isStop(endPunctCorrect) ? "Ending punctuation (. or !)" : `Ending punctuation (${endPunctCorrect})`;
  const grammar = [
    { label: "Capital letter to start the sentence", pass: !startsCapCorrect || startsCapTyped },
    { label: punctLabel, pass: punctPass }
  ];
  // Each sentence is worth 10 points; every wrong item (a misspelled, missing,
  // or extra word, the capital, the ending punctuation) costs 1 point, down to 0.
  const wrongItems = wordResults.filter(w => !w.spellingCorrect).length + grammar.filter(g => !g.pass).length;
  return {
    v: GRADE_VERSION, wordResults, grammar,
    points: Math.max(0, SENTENCE_POINTS - wrongItems),
    allSpellingCorrect: wordResults.every(w => w.spellingCorrect),
    allGrammarPass: grammar.every(g => g.pass)
  };
}
const SENTENCE_POINTS = 10;

// One-time migration: dictation sentences finished under the old all-or-nothing
// grader are re-graded with the 10-point rules from what she actually typed.
// Retype-to-fix progress is kept (word positions don't change), and the new
// score is saved back so every device shows it.
function regradeLegacyDictation() {
  if (!DATA || !state) return;
  Object.keys(DATA).forEach(key => {
    DATA[key].tasks.forEach(t => {
      if (t.type === "phrase-tagger") {
        // Re-score finished tagging tasks under the current matching rules (e.g. "flashlight" alone now counts as a direct object).
        const ps = state[key].tasks[t.id];
        if (!ps || !ps.done || !ps.selections || !t.phrases) return;
        const sc = phraseScore(t, ps.selections);
        const fresh = `${sc.ok}/${sc.total}`;
        if (ps.score !== fresh) { ps.score = fresh; persistTask(key, t.id); }
        return;
      }
      if (t.type !== "graded-dictation") return;
      const st = state[key].tasks[t.id];
      if (!st || !st.done || !st.results) return;
      if (!st.results.some(r => r.kind === "sentence" && r.grade && r.grade.v !== GRADE_VERSION)) return;
      let got = 0, possible = 0;
      st.results = st.results.map(r => {
        if (r.kind === "sentence") {
          const grade = gradeSentence(r.answer, r.typed || "");
          got += grade.points; possible += SENTENCE_POINTS;
          // Word positions can shift when spacing is forgiven, so carry retype progress over by word.
          let corrections = r.corrections;
          if (corrections && r.grade && r.grade.wordResults) {
            const byWord = {};
            Object.keys(corrections).forEach(k => {
              const old = r.grade.wordResults[Number(k)];
              if (old && old.correctWord) byWord[old.correctWord.toLowerCase()] = corrections[k];
            });
            corrections = {};
            grade.wordResults.forEach((wr, wi) => {
              if (!wr.spellingCorrect && wr.correctWord && byWord[wr.correctWord.toLowerCase()]) corrections[String(wi)] = byWord[wr.correctWord.toLowerCase()];
            });
          }
          return Object.assign({}, r, { grade }, corrections ? { corrections } : {});
        }
        got += r.correct ? 1 : 0; possible += 1;
        return r;
      });
      st.score = `${got}/${possible}`;
      persistTask(key, t.id);
    });
  });
}
// Points for one stored sentence result (older saved results had no points: all-or-nothing).
function sentencePoints(r) {
  if (typeof r.grade.points === "number") return r.grade.points;
  return (r.grade.allSpellingCorrect && r.grade.allGrammarPass) ? SENTENCE_POINTS : 0;
}

function checkDictation(key, id) {
  const t = DATA[key].tasks.find(x => x.id === id);
  const words = t.dynamic === "reviewPool" ? state[key].tasks[id]._reviewWords : t.words;
  let correctCount = 0, possible = 0;
  const results = [];
  words.forEach((w, i) => {
    if (w.kind === "sentence") {
      const typed = document.getElementById(`dict-${key}-${id}-${i}`).value.trim();
      const grade = gradeSentence(w.answer, typed);
      correctCount += grade.points;
      possible += SENTENCE_POINTS;
      results.push({ kind: "sentence", typed, answer: w.answer, grade });
      logGradedAnswer(key, {
        word: w.answer, question: "Dictation sentence", given: typed, correct: w.answer,
        wasCorrect: grade.points === SENTENCE_POINTS
      });
      grade.wordResults.forEach(wr => {
        if (!wr.spellingCorrect && wr.correctWord) logMiss(wr.correctWord, subjectTag(key) + " (in a sentence)", null);
      });
    } else {
      const typed = document.getElementById(`dict-${key}-${id}-${i}`).value.trim();
      const isRight = typed.toLowerCase() === w.answer.toLowerCase();
      if (isRight) correctCount++;
      possible += 1;
      results.push({ kind: "word", typed, correct: isRight, answer: w.answer, context: w.context || null });
      logGradedAnswer(key, { word: w.answer, question: "Dictation word", given: typed, correct: w.answer, wasCorrect: isRight });
      if (t.dynamic === "reviewPool") {
        logReviewResult(w.answer, isRight);
      } else if (!isRight) {
        const source = t.dynamic === "spellingMonthBank" ? "Monthly Test" : t.dynamic === "examSpellingBank" ? "Term Exam" : subjectTag(key);
        logMiss(w.answer, source, w.context);
      }
    }
  });
  state[key].tasks[id].results = results;
  state[key].tasks[id].score = `${correctCount}/${possible}`;
  state[key].tasks[id].done = true;
  if (t.dynamic === "spellingMonthBank") markTested(banksCache[currentChild].spelling || [], "spelling");
  persistTask(key, id);
  render();
}
function tagAbbrev(type) { return TAG_ABBREV[type] || type.slice(0, 4).toUpperCase(); }

// ---------- Retype-to-confirm for missed dictation words ----------
// Requires the child to retype each missed word correctly before it counts
// as reviewed, so a parent can confirm the word was actually learned, not
// just shown the answer. wi is null for a standalone missed word, or the
// index into a sentence's wordResults for a missed word inside a sentence.
function correctionKey(wi) { return wi === null || wi === undefined ? "w" : `${wi}`; }
function correctionBox(key, id, ri, wi, correctWord, r) {
  const ck = correctionKey(wi);
  const entry = (r.corrections && r.corrections[ck]) || null;
  const inputId = `correction-${key}-${id}-${ri}-${ck}`;
  if (entry && entry.confirmed) {
    return `<div class="correction-box confirmed">✓ Retyped correctly — nice work!</div>`;
  }
  return `<div class="correction-box${entry && !entry.confirmed ? " wrong" : ""}">
    <input type="text" id="${inputId}" placeholder="Retype it correctly..." value="${entry ? entry.typed : ""}">
    <button class="btn" onclick="submitCorrection('${key}','${id}',${ri},${wi === null ? "null" : wi},'${correctWord.replace(/'/g, "\\'")}')">Check</button>
    ${entry && !entry.confirmed ? `<div class="correction-msg">Not quite — try again.</div>` : ``}
  </div>`;
}
function dictationPassed(s) {
  const parts = (s.score || "0/1").split("/").map(Number);
  return parts[0] === parts[1] || parts[0] / parts[1] >= 0.7;
}
function allCorrectionsConfirmed(results) {
  return results.every(r => {
    if (r.kind === "sentence") {
      return r.grade.wordResults.every((wr, wi) => {
        if (wr.spellingCorrect || !wr.correctWord) return true;
        const entry = r.corrections && r.corrections[correctionKey(wi)];
        return entry && entry.confirmed;
      });
    }
    if (r.correct) return true;
    const entry = r.corrections && r.corrections[correctionKey(null)];
    return entry && entry.confirmed;
  });
}
function submitCorrection(key, id, ri, wi, correctWord) {
  const ck = correctionKey(wi);
  const inputId = `correction-${key}-${id}-${ri}-${ck}`;
  const typed = document.getElementById(inputId).value.trim();
  const r = state[key].tasks[id].results[ri];
  if (!r.corrections) r.corrections = {};
  r.corrections[ck] = { typed, confirmed: typed.toLowerCase() === correctWord.toLowerCase() };
  persistTask(key, id);
  render();
}

// ---------- Reading fluency check (parent listens, marks each word) ----------

function markFluencyWord(key, id, idx, correct) {
  const s = state[key].tasks[id];
  if (!s.answers.fluency) s.answers.fluency = {};
  s.answers.fluency[idx] = correct;
  render();
}
function checkFluency(key, id) {
  const t = DATA[key].tasks.find(x => x.id === id);
  const s = state[key].tasks[id];
  const marks = s.answers.fluency || {};
  const correctCount = t.words.filter((w, i) => marks[i] === true).length;
  s.score = `${correctCount}/${t.words.length}`;
  s.done = true;
  persistTask(key, id);
  render();
}

function openPosPopup(key, id, idx) {
  posPopupOpen = (posPopupOpen && posPopupOpen.key === key && posPopupOpen.id === id && posPopupOpen.idx === idx) ? null : { key, id, idx };
  render();
}
function selectPos(key, id, idx, option) {
  state[key].tasks[id].labels[idx] = option;
  posPopupOpen = null;
  render();
}
function checkPosTagging(key, id) {
  const t = DATA[key].tasks.find(x => x.id === id);
  const s = state[key].tasks[id];
  let correct = 0;
  s.labels.forEach((l, i) => {
    const wasCorrect = l === t.answers[i];
    if (wasCorrect) correct++;
    logGradedAnswer(key, { word: t.sentence[i], question: "Part of speech", given: l, correct: t.answers[i], wasCorrect });
  });
  s.score = `${correct}/${t.answers.length}`;
  s.done = true;
  persistTask(key, id);
  render();
}
function openConceptPopup(key, id, idx) {
  conceptPopupOpen = (conceptPopupOpen && conceptPopupOpen.key === key && conceptPopupOpen.id === id && conceptPopupOpen.idx === idx) ? null : { key, id, idx };
  render();
}
// A concept-check target's answer is a string, or an array when a word is
// legitimately two things at once (e.g. courtyard = common AND compound).
// Stored/compared as one string, in option order, joined by " + ".
function conceptAnswerText(target, options) {
  let ans = target.answer;
  // Older seeded rows stored just "Compound"/"Collective"; those nouns are also common.
  if (!Array.isArray(ans) && options.includes("Common") && (ans === "Compound" || ans === "Collective")) ans = ["Common", ans];
  if (!Array.isArray(ans)) return ans;
  return options.filter(o => ans.includes(o)).join(" + ");
}
// Noun-type and pronoun-role checks let you pick MORE than one answer (a word can
// be common AND compound; an indefinite pronoun can also be the subject). Verb
// checks (one role per verb) stay single-pick.
function conceptIsMulti(t) {
  return !!t.multi || t.targets.some(tg => Array.isArray(tg.answer)) || t.options.includes("Common") || t.options.includes("Indefinite");
}
function closeConceptPopup() { conceptPopupOpen = null; render(); }
function selectConceptAnswer(key, id, idx, option, multi) {
  const labels = state[key].tasks[id].labels;
  if (multi) {
    const t = DATA[key].tasks.find(x => x.id === id);
    const cur = (labels[idx] || "").split(" + ").filter(Boolean);
    const next = cur.includes(option) ? cur.filter(o => o !== option) : cur.concat(option);
    labels[idx] = next.length ? t.options.filter(o => next.includes(o)).join(" + ") : null;
    render();
    return;
  }
  labels[idx] = option;
  conceptPopupOpen = null;
  render();
}
function checkConceptCheck(key, id) {
  const t = DATA[key].tasks.find(x => x.id === id);
  const s = state[key].tasks[id];
  let correct = 0;
  t.targets.forEach(tg => {
    const given = s.labels[tg.index];
    const expected = conceptAnswerText(tg, t.options);
    const wasCorrect = given === expected;
    if (wasCorrect) correct++;
    logGradedAnswer(key, { word: t.sentence[tg.index], question: "Concept check", given, correct: expected, wasCorrect });
  });
  s.score = `${correct}/${t.targets.length}`;
  s.done = true;
  persistTask(key, id);
  render();
}
function selectPhraseWord(key, id, idx) {
  if (!phraseRangeStart || phraseRangeStart.key !== key || phraseRangeStart.id !== id) {
    phraseRangeStart = { key, id, idx };
    render();
    return;
  }
  const start = Math.min(phraseRangeStart.idx, idx);
  const end = Math.max(phraseRangeStart.idx, idx);
  phraseRangeStart = null;
  pendingPhraseRange = { key, id, start, end };
  render();
}
function choosePhraseType(type) {
  if (!pendingPhraseRange) return;
  const { key, id, start, end } = pendingPhraseRange;
  state[key].tasks[id].selections.push({ start, end, type });
  pendingPhraseRange = null;
  render();
}
function cancelPhraseRange() { pendingPhraseRange = null; render(); }
function clearPhraseSelections(key, id) { state[key].tasks[id].selections = []; render(); }
// A direct/indirect object may be tagged as the whole noun phrase ("the flashlight")
// or as just the noun ("flashlight") — both are correct grammar, so both count.
function spanHit(p, sel) {
  if (p.start === sel.start && p.end === sel.end) return true;
  return (p.type === "Direct Object" || p.type === "Indirect Object") && p.end > p.start && sel.start === p.end && sel.end === p.end;
}
function phraseHit(p, sel) { return p.type === sel.type && spanHit(p, sel); }

// Score = right chunks out of (chunks to find + chunks she made up that aren't
// any answer). Marking something that isn't one of the parts (e.g. a clause where
// phrases were asked for) costs a point; a right chunk with the wrong label only
// loses that one item, not two.
function phraseScore(t, selections) {
  const seen = new Set();
  const sels = (selections || []).filter(sel => { const k = sel.start + "-" + sel.end + "-" + sel.type; if (seen.has(k)) return false; seen.add(k); return true; });
  const ok = sels.filter(sel => t.phrases.some(p => phraseHit(p, sel))).length;
  const extra = sels.filter(sel => !t.phrases.some(p => spanHit(p, sel))).length;
  return { ok, extra, total: t.phrases.length + extra };
}

// ---------- "Why" explanations for parts-of-speech and noun/pronoun/verb-type misses ----------

const POS_DEFS = {
  Noun: "names a person, place, thing, or idea",
  Pronoun: "takes the place of a noun (she, they, everyone, mine)",
  Verb: "shows an action or a state of being, or helps another verb",
  Adjective: "describes a noun or pronoun — it tells which one, what kind, or how many (a, an, and the are adjectives called articles)",
  Adverb: "describes a verb, an adjective, or another adverb — it tells how, when, where, or how much",
  Preposition: "shows how a noun or pronoun relates to another word (where, when, which way), and starts a phrase like \"in the courtyard\"",
  Conjunction: "joins words, phrases, or clauses (and, but, or, because)",
  Interjection: "shows sudden feeling and stands apart from the sentence (Wow! Oh!)",
  Gerund: "is an -ing verb form that works as a noun",
  Participle: "is a verb form (-ing or -ed) that works as an adjective",
  Infinitive: "is \"to\" plus a verb (to try, to relax)"
};
// Tips for the mix-ups kids actually make: [said][actual]
const POS_CONTRAST = {
  Noun: {
    Adjective: "Watch for a noun doing an adjective's job: when a word describes another noun (like \"gymnastics\" in \"gymnastics team\"), it's working as an adjective in that sentence. Ask: is it naming something, or describing something?",
    Verb: "Ask: can the subject DO this word, or is it a thing? If it's the action or the state of being, it's a verb.",
    Pronoun: "A pronoun stands in for a noun. Try swapping in a name — if the sentence still makes sense, it's a pronoun."
  },
  Adjective: {
    Noun: "An adjective only describes. If the word names the person, place, or thing itself, it's a noun.",
    Adverb: "Adjectives describe nouns (which one? what kind?). Adverbs describe verbs, adjectives, or other adverbs (how? when? where? how much?).",
    Pronoun: "If the word stands in for a noun, it's a pronoun. If it comes right before a noun and describes it, it's an adjective."
  },
  Adverb: {
    Adjective: "Ask what the word is describing. If it describes a noun, it's an adjective. If it describes a verb, adjective, or adverb (how? when? where? how much?), it's an adverb.",
    Preposition: "A preposition needs an object right after it (\"in the courtyard\"). An adverb stands alone and answers how, when, or where.",
    Conjunction: "A conjunction joins two parts of a sentence. An adverb describes one word."
  },
  Verb: {
    Noun: "Ask: does it show an action or a state of being? Then it's a verb. Nouns name things.",
    Adjective: "Verbs tell what someone does or is. Adjectives describe nouns.",
    Preposition: "Watch \"to\": before a verb (to try) it starts an infinitive; before a noun (to the market) it's a preposition."
  },
  Preposition: {
    Conjunction: "A preposition is followed by a noun or pronoun (in the courtyard). A conjunction joins words, phrases, or clauses (and, but, because).",
    Adverb: "A preposition always has an object after it. If nothing follows it, it's probably an adverb.",
    Verb: "Prepositions relate a noun to another word; they don't show action."
  },
  Conjunction: {
    Preposition: "A conjunction connects two things (and, but, or, because). A preposition starts a phrase with an object (in, on, through).",
    Adverb: "Conjunctions join; adverbs describe."
  },
  Pronoun: {
    Noun: "Pronouns take the place of nouns (she, they, everyone). If it names something itself, it's a noun.",
    Adjective: "A possessive word by itself (mine, theirs) is a pronoun. Before a noun (her room) it describes the noun."
  },
  Gerund: {
    Verb: "It ends in -ing but is doing a NOUN's job here (a subject or an object), so it's a gerund, not a verb.",
    Noun: "It's a verb form (-ing) working as a noun, which makes it a gerund.",
    Participle: "A participle describes a noun. A gerund IS the noun (the thing being done)."
  },
  Participle: {
    Verb: "It looks like a verb but describes a noun here, so it's a participle.",
    Adjective: "It describes a noun, but it's made from a verb (-ing/-ed), so it's a participle.",
    Gerund: "A gerund acts as a noun. A participle describes one."
  },
  Infinitive: {
    Preposition: "\"To\" before a VERB (to try) starts an infinitive. \"To\" before a NOUN (to the market) is a preposition.",
    Verb: "\"To\" plus a verb make an infinitive together."
  }
};
function posWhyHTML(word, given, correct, note) {
  const parts = [];
  const noteLc = (note || "").trim().toLowerCase();
  const isArticle = noteLc === "article";
  // Skip notes that just repeat the part of speech's name (the definition below covers it).
  if (note && !isArticle && noteLc !== correct.toLowerCase()) parts.push(note.charAt(0).toUpperCase() + note.slice(1) + (/[.!?]$/.test(note) ? "" : "."));
  if (isArticle) parts.push(`<b>"${word}"</b> is an <b>article</b> (a, an, the). Articles are a special kind of adjective: they point to a noun and tell which one or how many.`);
  else if (POS_DEFS[correct]) parts.push(`<b>${correct}</b> means the word ${POS_DEFS[correct]}.`);
  const tip = !isArticle && given && POS_CONTRAST[given] && POS_CONTRAST[given][correct];
  if (tip) parts.push(tip);
  else if (given && POS_DEFS[given] && given !== correct) parts.push(`(${given} is a word that ${POS_DEFS[given]} — not the job this word does here.)`);
  return parts.join(" ");
}

const CONCEPT_DEFS = {
  Common: "names any person, place, or thing, not one specific one",
  Proper: "names one specific person, place, or thing, so it starts with a capital",
  Collective: "names a group acting as one unit, like team, family, or class",
  Compound: "is two words joined into one noun idea, like handstand or courtyard",
  Subject: "does the action in its clause",
  Object: "receives the action or comes after a preposition",
  Possessive: "shows who owns something",
  Indefinite: "points to no specific person or thing (everyone, someone, nobody)",
  Action: "shows something the subject does",
  Linking: "connects the subject to a word that describes or renames it (is, seems, tasted)",
  Helping: "works with the main verb to show tense or mood (was stirring, has practiced)"
};
function conceptWhyHTML(word, given, tg, options) {
  const expected = conceptAnswerText(tg, options);
  const rightList = expected.split(" + ");
  const bits = [];
  const gotSome = (given || "").split(" + ").filter(g => rightList.includes(g));
  if (rightList.length > 1) bits.push(`${gotSome.length ? `Partly right — <b>${gotSome.join(" + ")}</b> fits. ` : ""}This word is <b>both</b> ${rightList.map(a => `<b>${a}</b>, which ${CONCEPT_DEFS[a] || ""}`).join(", and ")}. A word can fit more than one kind, so tap every one that fits.`);
  else if (CONCEPT_DEFS[expected]) bits.push(`<b>${expected}</b> means the word ${CONCEPT_DEFS[expected]}.`);
  const givenList = (given || "").split(" + ").filter(Boolean).filter(g => !rightList.includes(g));
  givenList.forEach(g => { if (CONCEPT_DEFS[g]) bits.push(`<b>${g}</b> would mean it ${CONCEPT_DEFS[g]} — that's not what this word does here.`); });
  return bits.join(" ");
}

// Explains WHY a tagged chunk isn't one of the answers, not just that it's wrong.
// Uses the other tagging tasks on the same sentence (subjects, verbs, clauses)
// to tell a clause (has its own subject + verb) from a phrase (doesn't).
function explainOffKeyTag(key, t, sel) {
  const words = t.sentence.slice(sel.start, sel.end + 1);
  const text = words.join(" ").replace(/[,.;:]+$/, "");
  const sentenceKey = t.sentence.join(" ");
  const sibs = DATA[key].tasks.filter(x => x.type === "phrase-tagger" && x.sentence && x.sentence.join(" ") === sentenceKey);
  const all = sibs.flatMap(x => x.phrases || []);
  const inside = list => list.filter(p => p.start >= sel.start && p.end <= sel.end);
  const clean = p => t.sentence.slice(p.start, p.end + 1).join(" ").replace(/[,.;:]+$/, "");
  const subjects = inside(all.filter(p => p.type === "Subject"));
  const verbs = inside(all.filter(p => /Verb Predicate/.test(p.type)));
  const hasSV = subjects.length > 0 && verbs.length > 0;
  const opts = t.options || [];
  const isPhraseLevel = opts.includes("Prepositional") || opts.includes("Appositive");
  const isClauseLevel = opts.includes("Independent Clause");
  if (isPhraseLevel && hasSV) {
    const keyList = (t.phrases || []).map(p => `"${clean(p)}" (${p.type})`).join(" and ");
    return `<b>"${text}"</b> — that's a <b>clause</b>, not a phrase. It has its own subject (<b>${subjects.map(clean).join(", ")}</b>) and its own verb (<b>${verbs.map(clean).join(", ")}</b>). A <b>phrase</b> is a group of words that does <i>not</i> have a subject and verb working together; a <b>clause</b> does. The phrases to mark here are ${keyList || "the ones that work as a single part"}.`;
  }
  if (isClauseLevel && !hasSV) {
    return `<b>"${text}"</b> — that's a <b>phrase</b>, not a clause. It doesn't have its own subject <i>and</i> verb. A clause always needs both, like "Reynie carried…". Look for the chunks that could each be read as their own little statement.`;
  }
  const overlap = (t.phrases || []).find(p => sel.start <= p.end && sel.end >= p.start);
  if (overlap) {
    return `<b>"${text}"</b> — close, but not the exact chunk. The <b>${overlap.type}</b> here is <b>"${clean(overlap)}"</b>. ${overlap.explanation || ""}`;
  }
  return `<b>"${text}"</b> — that's not one of the parts we're looking for here.`;
}

function checkPhraseTagging(key, id) {
  const t = DATA[key].tasks.find(x => x.id === id);
  const s = state[key].tasks[id];
  const sc = phraseScore(t, s.selections);
  s.selections.forEach(sel => {
    if (!t.phrases.some(p => spanHit(p, sel))) {
      logGradedAnswer(key, { word: t.sentence.slice(sel.start, sel.end + 1).join(" "), question: "Phrase/clause tagging (extra chunk)", given: sel.type, correct: "(not one of the parts)", wasCorrect: false });
    }
  });
  t.phrases.forEach(p => {
    const phraseText = t.sentence.slice(p.start, p.end + 1).join(" ");
    const match = s.selections.find(sel => spanHit(p, sel));
    const wasCorrect = !!match && match.type === p.type;
    logGradedAnswer(key, { word: phraseText, question: "Phrase/clause tagging", given: match ? match.type : "(missed)", correct: p.type, wasCorrect });
  });
  s.score = `${sc.ok}/${sc.total}`;
  s.done = true;
  persistTask(key, id);
  render();
}
function checkFill(key, id) {
  const t = DATA[key].tasks.find(x => x.id === id);
  let correct = 0;
  t.words.forEach((w, i) => {
    const typed = document.getElementById(`fill-${key}-${id}-${i}`).value.trim();
    const wasCorrect = typed.toLowerCase() === w.answer.toLowerCase();
    if (wasCorrect) correct++;
    logGradedAnswer(key, { word: w.answer, question: "Fill in the blank", given: typed, correct: w.answer, wasCorrect });
  });
  state[key].tasks[id].score = `${correct}/${t.words.length}`;
  state[key].tasks[id].done = true;
  persistTask(key, id);
  render();
}
function selectMC(key, id, qIdx, optIdx) {
  if (state[key].tasks[id].score !== null) return;
  if (!state[key].tasks[id].answers.mc) state[key].tasks[id].answers.mc = {};
  state[key].tasks[id].answers.mc[qIdx] = optIdx;
  render();
}
function submitMC(key, id) {
  const t = DATA[key].tasks.find(x => x.id === id);
  const questions = getTaskQuestions(key, id);
  const answers = state[key].tasks[id].answers.mc || {};
  let correct = 0;
  questions.forEach((q, i) => {
    const wasCorrect = answers[i] === q.correct;
    if (wasCorrect) correct++;
    logGradedAnswer(key, {
      word: q.q, question: q.q,
      given: answers[i] != null ? q.options[answers[i]] : null,
      correct: q.options[q.correct], wasCorrect
    });
  });
  state[key].tasks[id].score = `${correct}/${questions.length}`;
  state[key].tasks[id].done = true;
  if (t.dynamic) {
    const cfg = getDynamicBankConfig(t.dynamic);
    if (cfg && !cfg.noMarker) markTested(cfg.bank, cfg.subjectKey);
  }
  persistTask(key, id);
  render();
}
function approveReflection(key, id) {
  const comment = document.getElementById(`comment-${key}-${id}`).value.trim();
  const s = state[key].tasks[id];
  s.reviewed = true;
  s.parentComment = comment || null;
  s.sentBack = false;
  persistTask(key, id);
  render();
}
function sendBackReflection(key, id) {
  const comment = document.getElementById(`comment-${key}-${id}`).value.trim();
  const s = state[key].tasks[id];
  s.parentComment = comment || "Please take another look and resubmit.";
  s.done = false;
  s.needsReview = false;
  s.reviewed = false;
  s.sentBack = true;
  persistTask(key, id);
  render();
}

// ---------- Rendering ----------

function taskHeadHTML(key, t) {
  const s = state[key].tasks[t.id];
  if (isTaskLocked(t)) {
    const msg = t.termFinal ? "🔒 Unlocks at end of term" : `🔒 Locked until Week ${nextMonthlyTestWeek()}`;
    return `<div class="task-head locked-task">
      <div class="dot"></div><span class="label">${t.label}</span>
      <span class="status-text">${msg}</span>
    </div>`;
  }
  let cls = "", statusText = "Not started";
  if (s.sentBack && !s.done) { cls = "review"; statusText = "Refired — needs another pass"; }
  else if (s.done && s.needsReview && !s.reviewed) { cls = "review"; statusText = "Awaiting review"; }
  else if (s.done) { cls = "done"; statusText = s.score ? `Scored ${s.score}` : (s.reviewed ? "Approved" : "Complete"); }
  return `<div class="task-head ${cls}" onclick="toggleTask('${key}','${t.id}')">
    <div class="dot"></div><span class="label">${t.label}</span>
    <span class="status-text">${statusText}</span>
  </div>`;
}

function taskBodyHTML(key, t) {
  const s = state[key].tasks[t.id];
  let inner = "";
  // Kenley doesn't want the whole session/instructions narrated to her (she
  // just needs single-word audio, handled separately per task type) — but
  // this function renders both kids' tasks, so every read-aloud button here
  // must stay gated on currentChild, not removed outright, or it disappears
  // for Adelyn too.
  const trimReadAloud = currentChild === "kenley";
  if (t.type === "read") {
    const contentId = `read-content-${key}-${t.id}`;
    const parentNotesHtml = (t.parentNotes && currentView === "parent")
      ? `<div class="parent-notes"><b>👪 Notes for you</b>${t.parentNotes}</div>` : "";
    inner = `${trimReadAloud ? "" : readAloudButton(contentId, "Read this to me")}
      <div id="${contentId}">${t.content}</div>
      ${parentNotesHtml}
      ${s.done ? `` : `<button class="btn primary" onclick="markRead('${key}','${t.id}')">Mark as read</button>`}`;
  } else if (t.type === "external") {
    const noteId = `ext-note-${key}-${t.id}`;
    inner = `${trimReadAloud ? "" : readAloudButton(noteId)}
      <a class="ext-link" href="${t.url || '#'}" target="_blank" rel="noopener">${t.linkText} ↗</a>
      <div class="lesson-text" id="${noteId}" style="opacity:.75;font-size:0.78rem;">${t.note}</div>
      <label style="font-size:0.82rem;display:flex;align-items:center;gap:8px;">
        <input type="checkbox" ${s.done ? "checked" : ""} onchange="markExternal('${key}','${t.id}',this.checked)"> Mark complete
      </label>`;
  } else if (t.type === "reflection") {
    const promptId = `refl-prompt-${key}-${t.id}`;
    const feedbackNote = s.parentComment ? `<div class="parent-feedback">📝 ${s.reviewed ? "Feedback from parent:" : "Refired — please revise:"} ${s.parentComment}</div>` : "";
    inner = `${trimReadAloud && key === "reading" ? "" : readAloudButton(promptId, "Read the question to me")}
      <div class="lesson-text" id="${promptId}"><p>${t.prompt}</p></div>
      ${feedbackNote}
      <textarea id="ta-${key}-${t.id}" placeholder="Type your answer here..." ${s.done ? "disabled" : ""}>${s.answers.text || ""}</textarea>
      ${s.done ? `<div class="graded-note">Submitted — waiting on parent review.</div>` : `<button class="btn primary" onclick="submitReflection('${key}','${t.id}')">${s.sentBack ? "Refire" : "Submit"}</button>`}`;
  } else if (t.type === "graded-dictation") {
    let words = t.words;
    const banks = banksCache[currentChild] || {};
    if (t.dynamic === "reviewPool") {
      words = getReviewWords(2);
      state[key].tasks[t.id]._reviewWords = words;
    } else if (t.dynamic === "spellingMonthBank") {
      if (!s.done) { words = sampleSpellingWords(banks.spelling || [], Math.min(15, (banks.spelling || []).length)); state[key].tasks[t.id]._reviewWords = words; }
      else words = state[key].tasks[t.id]._reviewWords || [];
    } else if (t.dynamic === "examSpellingBank") {
      if (!s.done) { words = sampleExamSpellingWords(banks.spelling || [], Math.min(30, (banks.spelling || []).length)); state[key].tasks[t.id]._reviewWords = words; }
      else words = state[key].tasks[t.id]._reviewWords || [];
    }
    const dictPromptId = `dict-prompt-${key}-${t.id}`;
    inner = `${trimReadAloud ? "" : readAloudButton(dictPromptId, "Read the instructions to me")}<div class="lesson-text" id="${dictPromptId}"><p>${t.prompt}</p></div>`;
    if (s.done && s.results) {
      inner += s.results.map((r, ri) => {
        if (r.kind === "sentence") {
          const chips = r.grade.wordResults.map(wr => `<span class="word-chip ${wr.spellingCorrect ? "chip-correct" : "chip-incorrect"}">${wr.typedWord || "—"}</span>`).join(" ");
          const missedWords = r.grade.wordResults.filter(wr => !wr.spellingCorrect && wr.correctWord).map(wr => wr.correctWord);
          const pts = sentencePoints(r);
          const overallClass = pts === SENTENCE_POINTS ? "correct" : "incorrect";
          const missedBoxes = r.grade.wordResults.map((wr, wi) => (!wr.spellingCorrect && wr.correctWord) ? correctionBox(key, t.id, ri, wi, wr.correctWord, r) : "").join("");
          return `<div class="dict-result ${overallClass}">
            <div class="dict-result-typed" style="margin-bottom:6px;">Spelling, word by word: <b>${pts}/${SENTENCE_POINTS} points</b></div>
            <div style="margin-bottom:8px;">${chips}</div>
            ${missedWords.length ? `<div class="dict-result-answer">Correct spelling for missed word${missedWords.length > 1 ? "s" : ""}: <b>${missedWords.join(", ")}</b></div>` : ``}
            ${missedBoxes}
            <div class="grammar-check">
              ${r.grade.grammar.map(g => `<div class="grammar-line ${g.pass ? "pass" : "fail"}">${g.pass ? "✓" : "✗"} ${g.label}</div>`).join("")}
            </div>
          </div>`;
        } else {
          const icon = r.correct ? "✓" : "✗";
          return `<div class="dict-result ${r.correct ? "correct" : "incorrect"}">
            <div class="dict-result-typed"><span class="dict-icon">${icon}</span> You wrote: <b>${r.typed || "(blank)"}</b></div>
            ${!r.correct ? `<div class="dict-result-answer">Correct spelling: <b>${r.answer}</b></div>${correctionBox(key, t.id, ri, null, r.answer, r)}` : ``}
          </div>`;
        }
      }).join("");
      const allConfirmed = allCorrectionsConfirmed(s.results);
      inner += `<div class="score-result ${dictationPassed(s) ? "pass" : "retry"}">Scored automatically: ${s.score}${t.dynamic === "reviewPool" ? " — pool updated, mastered words drop out automatically" : ""}. Any missed word — even outside this week's list — has been added to the review file.</div>`;
      if (!allConfirmed) {
        inner += `<div class="score-result retry">✏️ Before moving on: retype each missed word above until it's spelled correctly.</div>`;
      }
    } else if (words.length === 0) {
      inner += `<div class="empty-note">No review words yet — once ${CHILD_META[currentChild].name} misses a dictation word anywhere in the app, it'll show up here automatically.</div>`;
    } else {
      words.forEach((w, i) => {
        const isSentence = w.kind === "sentence";
        const hasContext = !!w.context;
        const playFn = hasContext
          ? `speakWordInContext('${w.answer.replace(/'/g, "\\'")}','${w.context.replace(/'/g, "\\'")}')`
          : `speakWord('${w.answer.replace(/'/g, "\\'")}')`;
        const btnLabel = isSentence ? "Play sentence" : (hasContext ? `Play word ${i + 1} (in a sentence)` : `Play word ${i + 1}`);
        inner += `<div class="word-row" style="align-items:${isSentence ? "flex-start" : "center"};">
          <button class="btn" style="margin-top:0;flex-shrink:0;" onclick="${playFn}">🔊 ${btnLabel}</button>
          ${isSentence
            ? `<textarea id="dict-${key}-${t.id}-${i}" style="min-height:44px;" placeholder="Type the sentence you hear"></textarea>`
            : `<input type="text" id="dict-${key}-${t.id}-${i}" placeholder="Type what you hear">`}
        </div>`;
      });
      inner += `<button class="btn primary" onclick="checkDictation('${key}','${t.id}')">Check my spelling</button>`;
    }
  } else if (t.type === "fluency-read") {
    const fluencyPromptId = `fluency-prompt-${key}-${t.id}`;
    inner = `${trimReadAloud ? "" : readAloudButton(fluencyPromptId, "Read the instructions to me")}<div class="lesson-text" id="${fluencyPromptId}"><p>${t.prompt}</p></div>`;
    if (s.done) {
      inner += t.words.map((w, i) => {
        const mark = s.answers.fluency ? s.answers.fluency[i] : false;
        return `<div class="fluency-row ${mark ? "correct" : "incorrect"}">
          <span class="fluency-word">${w.answer}</span>
          <span class="fluency-icon">${mark ? "✓" : "✗"}</span>
        </div>`;
      }).join("");
      inner += `<div class="score-result ${(s.answers.fluency ? Object.values(s.answers.fluency).filter(Boolean).length : 0) / t.words.length >= 0.7 ? "pass" : "retry"}">Marked by a grown-up: ${s.score}</div>`;
    } else {
      inner += `<div class="lesson-text" style="opacity:.75;font-size:0.78rem;">Read each word out loud. Not sure of one? Tap 🔊 to hear it. When she's read them all, a grown-up marks each one and taps Save.</div>`;
      inner += t.words.map((w, i) => {
        const mark = s.answers.fluency ? s.answers.fluency[i] : undefined;
        const playFn = `speakWord('${w.answer.replace(/'/g, "\\'")}')`;
        return `<div class="fluency-row">
          <span class="fluency-word">${w.answer}</span>
          <button class="btn" onclick="${playFn}">🔊 Hear it</button>
          <button class="btn ${mark === true ? "done" : ""}" onclick="markFluencyWord('${key}','${t.id}',${i},true)">✓ Read it right</button>
          <button class="btn ${mark === false ? "done" : ""}" onclick="markFluencyWord('${key}','${t.id}',${i},false)">Needs practice</button>
        </div>`;
      }).join("");
      inner += `<button class="btn primary" onclick="checkFluency('${key}','${t.id}')">Save results</button>`;
    }
  } else if (t.type === "pos-tagger") {
    const posSentenceId = `pos-sentence-${key}-${t.id}`;
    inner = `<div class="lesson-text"><p>Tap a word, then tap its part of speech.</p></div>
      ${trimReadAloud ? "" : readAloudButton(posSentenceId, "Read the sentence to me")}
      <div id="${posSentenceId}" style="opacity:.75;font-size:0.82rem;">${t.sentence.join(" ")}</div>
      <div class="pos-row">`;
    let openTray = "";
    t.sentence.forEach((word, i) => {
      const label = s.labels[i];
      let cls = "", shownLabel = label;
      if (s.done) {
        cls = label === t.answers[i] ? "correct" : "incorrect";
        if (label !== t.answers[i]) shownLabel = `${label || "—"} → ${t.answers[i]}`;
      }
      const popupOpen = !s.done && posPopupOpen && posPopupOpen.key === key && posPopupOpen.id === t.id && posPopupOpen.idx === i;
      inner += `<div class="word-slot ${popupOpen ? "active-word" : ""}" onclick="${s.done ? "" : `openPosPopup('${key}','${t.id}',${i})`}">
        <div class="word-text">${word}</div>
        <div class="word-label ${cls}">${shownLabel || "+ tag"}</div>
      </div>`;
      if (popupOpen) openTray = `<div class="pos-tray"><div class="pos-tray-title">What is "<b>${word}</b>"?</div><div class="pos-tray-options">${t.options.map(o => `<button onclick="selectPos('${key}','${t.id}',${i},'${o}')">${o}</button>`).join("")}</div></div>`;
    });
    inner += `</div>${openTray}`;
    if (s.done) {
      const misses = t.sentence.map((w, i) => i).filter(i => s.labels[i] !== t.answers[i]);
      if (misses.length) {
        inner += `<div class="tag-review">${misses.map(i => `<div class="tag-review-item"><b>${t.sentence[i]}</b> — you said ${s.labels[i] || "nothing"}, it's actually <b>${t.answers[i]}</b>. ${posWhyHTML(t.sentence[i], s.labels[i], t.answers[i], t.explanations[i])}</div>`).join("")}</div>`;
      }
      inner += `<div class="score-result ${s.score.split("/")[0] === s.score.split("/")[1] ? "pass" : "retry"}">Scored automatically: ${s.score}</div>`;
    } else {
      const allLabeled = s.labels.every(l => l !== null);
      inner += `<button class="btn primary" ${allLabeled ? "" : "disabled"} onclick="checkPosTagging('${key}','${t.id}')">Check my tagging</button>`;
    }
  } else if (t.type === "phrase-tagger") {
    const phraseSentenceId = `phrase-sentence-${key}-${t.id}`;
    inner = `<div class="lesson-text"><p>Tap the first word, then the last word of a chunk, then choose what it is. (Tap the same word twice for a single-word chunk.)</p></div>
      ${trimReadAloud ? "" : readAloudButton(phraseSentenceId, "Read the sentence to me")}
      <div id="${phraseSentenceId}" style="opacity:.75;font-size:0.82rem;">${t.sentence.join(" ")}</div>
      <div class="pos-row">`;
    t.sentence.forEach((word, i) => {
      const sel = s.selections.find(sel => i >= sel.start && i <= sel.end);
      let cls = "", tagLabel = "";
      if (sel) {
        tagLabel = tagAbbrev(sel.type);
        if (s.done) {
          const match = t.phrases.some(p => phraseHit(p, sel));
          cls = match ? "correct" : "incorrect";
        }
      }
      const isPendingStart = !s.done && phraseRangeStart && phraseRangeStart.key === key && phraseRangeStart.id === t.id && phraseRangeStart.idx === i;
      inner += `<div class="word-slot phrase-word ${sel ? "in-phrase" : ""} ${isPendingStart ? "pending-start" : ""}" onclick="${s.done ? "" : `selectPhraseWord('${key}','${t.id}',${i})`}">
        <div class="word-text">${word}</div>
        ${tagLabel ? `<div class="word-label ${cls}">${tagLabel}</div>` : ""}
      </div>`;
    });
    inner += `</div>`;
    if (pendingPhraseRange && pendingPhraseRange.key === key && pendingPhraseRange.id === t.id) {
      const phraseText = t.sentence.slice(pendingPhraseRange.start, pendingPhraseRange.end + 1).join(" ");
      inner += `<div class="phrase-type-picker">
        <div>What is "${phraseText}"?</div>
        <div style="display:flex;gap:6px;flex-wrap:wrap;">
          ${t.options.map(o => `<button class="btn" style="margin-top:0;" onclick="choosePhraseType('${o}')">${o}</button>`).join("")}
          <button class="btn" style="margin-top:0;" onclick="cancelPhraseRange()">Cancel</button>
        </div>
      </div>`;
    }
    if (s.done) {
      const reviewLines = [];
      s.selections.forEach(sel => {
        const exact = t.phrases.find(p => phraseHit(p, sel));
        if (!exact) {
          const text = t.sentence.slice(sel.start, sel.end + 1).join(" ");
          const sameSpan = t.phrases.find(p => spanHit(p, sel));
          if (sameSpan) reviewLines.push(`<b>"${text}"</b> — you tagged it ${sel.type}, but it's actually <b>${sameSpan.type}</b>. ${sameSpan.explanation}`);
          else reviewLines.push(explainOffKeyTag(key, t, sel));
        }
      });
      t.phrases.forEach(p => {
        const found = s.selections.some(sel => spanHit(p, sel));
        if (!found) {
          const text = t.sentence.slice(p.start, p.end + 1).join(" ");
          reviewLines.push(`<b>"${text}"</b> — you didn't tag this one. It's the <b>${p.type}</b>. ${p.explanation}`);
        }
      });
      if (reviewLines.length) inner += `<div class="tag-review">${reviewLines.map(l => `<div class="tag-review-item">${l}</div>`).join("")}</div>`;
      inner += `<div class="score-result ${s.score.split("/")[0] === s.score.split("/")[1] ? "pass" : "retry"}">Scored automatically: ${s.score}</div>`;
    } else {
      inner += `<div style="display:flex;gap:8px;margin-top:10px;">
        <button class="btn" onclick="clearPhraseSelections('${key}','${t.id}')">Clear</button>
        <button class="btn primary" onclick="checkPhraseTagging('${key}','${t.id}')">Check my answers</button>
      </div>`;
    }
  } else if (t.type === "concept-check") {
    // Follow-up to a pos-tagger sentence: only the words the lesson is
    // actually teaching (t.targets) are tappable; everything else is plain
    // context so the sentence still reads naturally.
    const conceptPromptId = `concept-prompt-${key}-${t.id}`;
    inner = `${trimReadAloud ? "" : readAloudButton(conceptPromptId)}<div class="lesson-text" id="${conceptPromptId}"><p>${t.prompt}</p></div><div class="pos-row">`;
    let openConceptTray = "";
    t.sentence.forEach((word, i) => {
      const target = t.targets.find(tg => tg.index === i);
      if (!target) {
        inner += `<div class="word-slot plain-word"><div class="word-text">${word}</div></div>`;
        return;
      }
      const label = s.labels[i];
      let cls = "", shownLabel = label;
      if (s.done) {
        cls = label === conceptAnswerText(target, t.options) ? "correct" : "incorrect";
        if (cls === "incorrect") shownLabel = `${label || "—"} → ${conceptAnswerText(target, t.options)}`;
      }
      const popupOpen = !s.done && conceptPopupOpen && conceptPopupOpen.key === key && conceptPopupOpen.id === t.id && conceptPopupOpen.idx === i;
      inner += `<div class="word-slot ${popupOpen ? "active-word" : ""}" onclick="${s.done ? "" : `openConceptPopup('${key}','${t.id}',${i})`}">
        <div class="word-text">${word}</div>
        <div class="word-label ${cls}">${shownLabel || "+ tag"}</div>
      </div>`;
      if (popupOpen) {
        const multi = conceptIsMulti(t);
        const picked = (label || "").split(" + ");
        openConceptTray = `<div class="pos-tray"><div class="pos-tray-title">${multi ? `What kind of noun is "<b>${word}</b>"? It can be more than one — tap every one that fits, then Done.` : `What kind of noun is "<b>${word}</b>"?`}</div><div class="pos-tray-options">${t.options.map(o => `<button class="${multi && picked.includes(o) ? "picked" : ""}" onclick="selectConceptAnswer('${key}','${t.id}',${i},'${o}',${multi})">${o}</button>`).join("")}${multi ? `<button class="tray-done" onclick="closeConceptPopup()">Done</button>` : ""}</div></div>`;
      }
    });
    inner += `</div>${openConceptTray}`;
    if (s.done) {
      const misses = t.targets.filter(tg => s.labels[tg.index] !== conceptAnswerText(tg, t.options));
      if (misses.length) {
        inner += `<div class="tag-review">${misses.map(tg => `<div class="tag-review-item"><b>${t.sentence[tg.index]}</b> — you said ${s.labels[tg.index] || "nothing"}, it's actually <b>${conceptAnswerText(tg, t.options)}</b>. ${conceptWhyHTML(t.sentence[tg.index], s.labels[tg.index], tg, t.options) || tg.explanation || ""}</div>`).join("")}</div>`;
      }
      inner += `<div class="score-result ${s.score.split("/")[0] === s.score.split("/")[1] ? "pass" : "retry"}">Scored automatically: ${s.score}</div>`;
    } else {
      const allLabeled = t.targets.every(tg => s.labels[tg.index] !== undefined && s.labels[tg.index] !== null);
      inner += `<button class="btn primary" ${allLabeled ? "" : "disabled"} onclick="checkConceptCheck('${key}','${t.id}')">Check my answers</button>`;
    }
  } else if (t.type === "graded-mc") {
    const questions = getTaskQuestions(key, t.id);
    if (t.dynamic) {
      const hint = {
        vocabMonthBank: "Up to 10 words, sampled from everything taught since the last monthly test (falls back to the whole set if nothing new yet).",
        grammarMonthBank: "Concepts covered since the last monthly test.",
        readingMonthBank: "Comprehension check on everything discussed since the last monthly test.",
        examVocabBank: "Up to 20 questions — double a monthly test — pulled from every vocab word taught all term, not just since the last test.",
        examGrammarBank: "Up to 12 questions pulled from every grammar concept taught all term.",
        examReadingBank: "Up to 12 comprehension questions pulled from everything read and discussed all term."
      }[t.dynamic];
      inner += `<div class="lesson-text" style="opacity:.75;font-size:0.78rem;">${hint || ""}</div>`;
    }
    questions.forEach((q, qi) => {
      const qId = `mc-q-${key}-${t.id}-${qi}`;
      inner += `<div id="${qId}"><div class="lesson-text"><p>${qi + 1}. ${q.q}</p></div><div class="mc-options">`;
      q.options.forEach((opt, oi) => {
        const sel = s.answers.mc && s.answers.mc[qi] === oi;
        let cls = sel ? "selected" : "";
        if (s.done) {
          if (oi === q.correct) cls = "correct";
          else if (sel && oi !== q.correct) cls = "incorrect";
        }
        inner += `<div class="mc-option ${cls}" onclick="selectMC('${key}','${t.id}',${qi},${oi})">${opt}</div>`;
      });
      inner += `</div></div>${trimReadAloud ? "" : readAloudButton(qId, "Read this question and choices to me")}`;
      if (s.done && q.explanation) inner += `<div class="tag-review-item" style="margin-top:6px;">${q.explanation}</div>`;
    });
    if (s.done) {
      const parts = s.score.split("/");
      inner += `<div class="score-result ${parts[0] === parts[1] || (parseInt(parts[0]) / parseInt(parts[1])) >= 0.7 ? "pass" : "retry"}">Scored automatically: ${s.score}</div>`;
    } else {
      inner += `<button class="btn primary" onclick="submitMC('${key}','${t.id}')">Submit answers</button>`;
    }
  }
  if (s.sentBack && !s.done && s.parentComment && t.type !== "reflection") {
    inner = `<div class="parent-feedback">📝 Sent back — please redo this part: ${s.parentComment}</div>` + inner;
  }
  if (canStudentRedo(t, s)) {
    inner += `<button class="btn redo-btn" onclick="redoSection('${key}','${t.id}')">🔄 Redo this section to improve it</button>`;
  }
  return `<div class="task-body ${s.open ? "open" : ""}">${inner}</div>`;
}

// ---------- Past-week report / upcoming-week preview (read-only, no handlers) ----------

// Renders the sentence as word tiles with each word's correct part-of-speech tag underneath —
// reuses the same tile look as the live interactive pos-tagger, just non-clickable and pre-filled.
function renderPosTaggerPreview(t) {
  const tiles = t.sentence.map((w, i) =>
    `<div class="word-slot" style="cursor:default;"><span class="word-text">${w}</span><span class="word-label key">${t.answers[i]}</span></div>`
  ).join("");
  return `<div class="pos-row">${tiles}</div>`;
}

// Shows the full sentence with only the taught-concept target words tiled
// with their correct subtype answer; other words render as plain text.
function renderConceptCheckPreview(t) {
  const tiles = t.sentence.map((w, i) => {
    const target = t.targets.find(tg => tg.index === i);
    if (!target) return `<div class="word-slot plain-word"><span class="word-text">${w}</span></div>`;
    return `<div class="word-slot" style="cursor:default;"><span class="word-text">${w}</span><span class="word-label key">${target.answer}</span></div>`;
  }).join("");
  return `<div class="lesson-text"><p>${t.prompt}</p></div><div class="pos-row">${tiles}</div>`;
}

// Renders the sentence plus each authored phrase/clause with its type and explanation.
function renderPhraseTaggerPreview(t) {
  const sentence = t.sentence.join(" ");
  const items = (t.phrases || []).map(p => {
    const text = t.sentence.slice(p.start, p.end + 1).join(" ");
    return `<div class="preview-key-item">"${text}" — <b>${p.type}</b>. ${p.explanation || ""}</div>`;
  }).join("");
  return `<div class="preview-sentence">${sentence}</div><div class="preview-key">${items}</div>`;
}

// Lists each authored word/sentence with its correct spelling/answer for a dictation set.
function renderDictationPreview(t) {
  const promptHtml = t.prompt ? `<p>${t.prompt}</p>` : "";
  if (!t.words || t.words.length === 0) {
    return `<div class="lesson-text">${promptHtml}<p style="opacity:.75;">(pulled dynamically — not fixed content to preview)</p></div>`;
  }
  const items = t.words.map(w => `<div class="preview-key-item">${w.answer}${w.context ? ` <span class="tag">(context: ${w.context})</span>` : ""}</div>`).join("");
  return `<div class="lesson-text">${promptHtml}</div><div class="preview-key">${items}</div>`;
}

// Lists each question with its options, marking the correct one and showing any explanation.
function renderGradedMcPreview(t) {
  if (!t.questions || t.questions.length === 0) {
    return `<div class="lesson-text"><p style="opacity:.75;">(pulled dynamically — not fixed content to preview)</p></div>`;
  }
  return t.questions.map(q => {
    const opts = (q.options || []).map((o, i) =>
      `<div class="${i === q.correct ? "preview-key-item" : "preview-key-item-dim"}">${i === q.correct ? "✓ " : ""}${o}</div>`
    ).join("");
    const explanation = q.explanation ? `<div class="sample-answer"><b>Why:</b> ${q.explanation}</div>` : "";
    return `<div class="lesson-text"><p>${q.q}</p></div><div class="preview-key">${opts}</div>${explanation}`;
  }).join(`<hr style="border:none;border-top:1px solid var(--card-border);margin:10px 0;">`);
}

// Shared content renderer used by both the current/past-week report and the upcoming-week
// preview — the lesson material itself doesn't depend on whether she's done it yet.
function renderTaskContent(t) {
  if (t.type === "read") return t.content + (t.parentNotes ? `<div class="parent-notes"><b>👪 Notes for you</b>${t.parentNotes}</div>` : "");
  if (t.type === "reflection") {
    return `<div class="lesson-text"><p>${t.prompt}</p></div>${t.sampleAnswer ? `<div class="sample-answer"><b>Sample answer:</b> ${t.sampleAnswer}</div>` : ""}`;
  }
  if (t.type === "external") return `<div class="lesson-text">${t.note || ""}</div>`;
  if (t.type === "graded-mc") return renderGradedMcPreview(t);
  if (t.type === "graded-dictation") return renderDictationPreview(t);
  if (t.type === "fluency-read") return renderDictationPreview(t);
  if (t.type === "pos-tagger") return renderPosTaggerPreview(t);
  if (t.type === "phrase-tagger") return renderPhraseTaggerPreview(t);
  if (t.type === "concept-check") return renderConceptCheckPreview(t);
  return `<div class="lesson-text">(interactive exercise — nothing to preview yet)</div>`;
}

// Task types with real per-item auto-grading, where taskBodyHTML's "done"
// branch already colors each item green/pink by her actual answer — reuse
// that directly instead of the generic blue reference-key preview once
// she's actually completed the task. (selectMC/selectPos/etc. all no-op
// once s.done is true, so the reused markup is effectively read-only.)
const AUTO_GRADED_TYPES = ["graded-mc", "graded-dictation", "fluency-read", "pos-tagger", "phrase-tagger", "concept-check"];

// ---------- Case Files vocab report (parent view) ----------

function escHtml(str) {
  return String(str == null ? "" : str).replace(/[&<>"']/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]));
}
function isCaseFilesTask(t) {
  return t.type === "external" && /case-files/i.test(t.url || "");
}
// Which vocab words she has played in Case Files, and which she missed, from
// the answer log the game pushes to the Sheet (vocabulary questions only).
function caseFilesVocabReportHTML() {
  const rows = (answerLogCache[currentChild] || []).filter(r =>
    /case files/i.test(r.game || "") && /vocab/i.test(r.subject || "") && !/^TEST/i.test(r.word || ""));
  if (rows.length === 0) {
    return `<div class="cf-report"><b>📚 Case Files vocabulary</b><div class="cf-empty">No vocabulary answers from Case Files have come in yet. Once ${CHILD_META[currentChild].name} plays a round, the words she practiced and any she missed will show here.</div></div>`;
  }
  const byWord = {};
  rows.forEach(r => {
    const w = byWord[r.word] || (byWord[r.word] = { word: r.word, tries: 0, right: 0, misses: [], last: r.timestamp });
    w.tries++;
    if (r.correct) w.right++; else w.misses.push(r);
    if (r.timestamp > w.last) w.last = r.timestamp;
  });
  const words = Object.values(byWord).sort((x, y) => x.word.localeCompare(y.word));
  const missed = words.filter(w => w.misses.length > 0);
  const chips = words.map(w => `<span class="word-chip ${w.misses.length ? "chip-incorrect" : "chip-correct"}">${escHtml(w.word)}${w.misses.length ? ` ✗${w.misses.length}` : " ✓"}</span>`).join(" ");
  const missLines = missed.map(w => w.misses.map(m =>
    `<div class="cf-miss"><b>${escHtml(w.word)}</b> — she chose "${escHtml(m.givenAnswer)}"; the right answer is "${escHtml(m.correctAnswer)}".</div>`).join("")).join("");
  const lastPlay = new Date(rows.reduce((mx, r) => (r.timestamp > mx ? r.timestamp : mx), rows[0].timestamp)).toLocaleDateString();
  return `<div class="cf-report"><b>📚 Case Files vocabulary</b>
    <div class="cf-summary">${words.length} word${words.length === 1 ? "" : "s"} practiced · ${missed.length} with a miss · last played ${lastPlay}</div>
    <div>${chips}</div>
    ${missed.length ? `<div class="cf-miss-title">Missed:</div>${missLines}` : `<div class="cf-summary">No misses — every word she's answered so far was right. 🎉</div>`}
    <div class="cf-note">Shows her most recent answers on record, so very old play may drop off.</div>
  </div>`;
}

function renderPastTaskReport(key, t, s) {
  // taskBodyHTML's wrapper only shows its content when the task row has
  // been toggled open (s.open) — force it open here without mutating the
  // real state, since this read-only report always shows content expanded.
  let extra = (s.done && AUTO_GRADED_TYPES.includes(t.type))
    ? taskBodyHTML(key, t).replace('class="task-body ', 'class="task-body open ')
    : renderTaskContent(t);
  if (isCaseFilesTask(t) && currentChild === "kenley") extra += caseFilesVocabReportHTML();
  const awaiting = t.type === "reflection" && s.needsReview && !s.reviewed;
  if (t.type === "reflection" && s.answers && s.answers.text) {
    extra += `<div class="submitted-text student-answer">${s.answers.text}</div>`;
    if (s.parentComment) extra += `<div class="parent-feedback">📝 ${s.parentComment}</div>`;
  }
  if (awaiting) {
    // Review happens right here, next to the prompt and sample answer it responds to.
    extra += `<textarea id="comment-${key}-${t.id}" placeholder="Optional feedback (shown either way — required reading if you send it back)" style="min-height:50px;margin-top:6px;"></textarea>
      <div class="review-actions">
        <button onclick="sendBackReflection('${key}','${t.id}')">Refire (send back)</button>
        <button class="approve" onclick="approveReflection('${key}','${t.id}')">Approve</button>
      </div>`;
  }
  if (s.done && !awaiting) {
    extra += `<details class="sendback-box"><summary>↩ Send back just this section</summary>
      <textarea id="sendback-${key}-${t.id}" placeholder="Optional note about what to fix (she'll see it)" style="min-height:44px;"></textarea>
      <div class="review-actions"><button onclick="sendBackSection('${key}','${t.id}')">Send back &amp; reset this section</button></div>
    </details>`;
  }
  return `<div class="review-item${awaiting ? " needs-attention" : ""}" id="review-${key}-${t.id}">
    <strong>${t.label}</strong>
    <div class="meta">${awaiting ? "Submitted by " + CHILD_META[currentChild].name + " — awaiting your review" : statusLabel(s)}${s.score ? " · Scored " + s.score : ""}</div>
    ${extra}
  </div>`;
}

function renderUpcomingTaskPreview(t) {
  return `<div class="review-item"><strong>${t.label}</strong><div class="meta">${t.type}</div>${renderTaskContent(t)}</div>`;
}

function renderWeekReportPanel() {
  const panel = document.getElementById("weekReportPanel");
  const week = parentNavWeek;
  const isPast = week <= currentWeek();
  const sections = [];
  let anyContent = false;
  SUBJECT_ORDER.forEach(key => {
    const tasks = weekScopedTasks(key, week);
    if (tasks.length === 0) return;
    anyContent = true;
    sections.push(`<div class="week-report-subject">${DATA[key].name}</div>`);
    tasks.forEach(t => {
      const s = state[key].tasks[t.id];
      sections.push(isPast ? renderPastTaskReport(key, t, s) : renderUpcomingTaskPreview(t));
    });
  });
  const titleSuffix = week < currentWeek() ? "Completed Record" : week === currentWeek() ? "In Progress" : "Preview";
  panel.innerHTML = `
    <div class="week-report-title">Week ${week} — ${titleSuffix}</div>
    ${anyContent ? sections.join("") : `<div class="empty-note">Nothing planned yet for Week ${week}.</div>`}
  `;
}

// ---------- PDF export ----------

function exportWeeksPdf() {
  const from = Math.max(1, parseInt(document.getElementById("exportFromWeek").value, 10) || 1);
  const to = Math.max(from, parseInt(document.getElementById("exportToWeek").value, 10) || currentWeek());
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "letter" });
  const marginX = 48, maxWidth = 612 - marginX * 2, pageBottom = 792 - 48;
  let y = 56;

  function writeLine(text, opts = {}) {
    const size = opts.size || 10;
    doc.setFontSize(size);
    doc.setFont(undefined, opts.bold ? "bold" : "normal");
    const lines = doc.splitTextToSize(String(text), maxWidth);
    if (y + lines.length * size * 1.35 > pageBottom) { doc.addPage(); y = 56; }
    doc.text(lines, marginX, y);
    y += lines.length * size * 1.35;
  }

  writeLine(`The ELA Pastry Kitchen — ${CHILD_META[currentChild].name}`, { size: 18, bold: true });
  writeLine(`Weeks ${from}–${to} · exported ${new Date().toLocaleDateString()}`, { size: 10 });
  y += 8;

  for (let week = from; week <= to; week++) {
    writeLine(`Week ${week}`, { size: 14, bold: true });
    let anyForWeek = false;
    SUBJECT_ORDER.forEach(key => {
      const tasks = weekScopedTasks(key, week);
      if (tasks.length === 0) return;
      anyForWeek = true;
      writeLine(DATA[key].name, { size: 12, bold: true });
      tasks.forEach(t => {
        const s = state[key].tasks[t.id];
        let line = `${t.label} (${t.type}) — ${statusLabel(s)}`;
        if (s.score) line += ` — Scored ${s.score}`;
        writeLine(line, { size: 10 });
        if (t.type === "reflection" && s.answers && s.answers.text) {
          writeLine(s.answers.text, { size: 9 });
          if (s.parentComment) writeLine(`Parent feedback: ${s.parentComment}`, { size: 9 });
        }
      });
      y += 4;
    });
    if (!anyForWeek) writeLine("Nothing recorded for this week.", { size: 10 });
    y += 10;
  }

  doc.save(`ela-pastry-kitchen-${currentChild}-weeks-${from}-${to}.pdf`);
}

const stampedStations = new Set();
const stampSeeded = {}; // per child: false until the first render has recorded already-served plates

function render() {
  regradeLegacyDictation();
  document.getElementById("childSwitcher").innerHTML = Object.keys(CHILD_META).map(id =>
    `<button class="child-pill kid-${id} ${currentChild === id ? "active" : ""}" onclick="switchChild('${id}')">${currentChild === id ? "✓ " : ""}${CHILD_META[id].name}</button>`
  ).join("");
  document.getElementById("board").dataset.child = currentChild;
  document.body.dataset.child = currentChild;
  document.body.dataset.view = currentView;
  document.getElementById("boardSub").textContent = CHILD_META[currentChild].subtitle;

  document.getElementById("viewToggle").classList.toggle("parent", currentView === "parent");
  document.getElementById("toggleKnob").textContent = currentView === "parent" ? "PARENT" : CHILD_META[currentChild].name.toUpperCase();
  document.getElementById("reviewQueue").style.display = currentView === "parent" ? "block" : "none";
  if (currentView !== "parent") document.getElementById("waitingOnYou").style.display = "none";

  if (currentView === "parent" && parentNavWeek === null) parentNavWeek = currentWeek();
  const showingWeekReport = currentView === "parent";

  document.getElementById("studentWeekControl").style.display = currentView === "parent" ? "none" : "flex";
  const navCtrl = document.getElementById("weekNavControl");
  const exportCtrl = document.getElementById("exportControl");
  navCtrl.style.display = currentView === "parent" ? "flex" : "none";
  exportCtrl.style.display = currentView === "parent" ? "flex" : "none";
  if (currentView === "parent") {
    const week = parentNavWeek;
    const qualifier = week === currentWeek() ? " (current)" : week < currentWeek() ? " (past — read only)" : " (upcoming — preview)";
    document.getElementById("weekNavLabel").textContent = `Week ${week}${qualifier}`;
    document.getElementById("weekNavPrev").disabled = week <= 1;
    document.getElementById("weekNavNext").disabled = week >= maxAuthoredWeek();
    if (!exportControlsReady) {
      document.getElementById("exportFromWeek").value = 1;
      document.getElementById("exportToWeek").value = currentWeek();
      exportControlsReady = true;
    }
  }
  document.getElementById("weekReportPanel").style.display = showingWeekReport ? "block" : "none";
  if (showingWeekReport) renderWeekReportPanel();

  const grid = document.getElementById("stationsGrid");
  grid.innerHTML = "";
  let doneCount = 0;
  const keys = Object.keys(DATA);
  keys.forEach(key => {
    const status = stationStatus(key);
    if (stationDone(key)) doneCount++; // only count actually-completed stations, not merely-locked/empty ones
    const needsReview = DATA[key].tasks.some(t => state[key].tasks[t.id].needsReview && !state[key].tasks[t.id].reviewed);
    const sentBackTasks = activeTasks(key).filter(t => state[key].tasks[t.id].sentBack && !state[key].tasks[t.id].done);
    const card = document.createElement("div");
    card.className = "station" + (status === "served" ? " done" : "") + (status === "burning" ? " burning" : "") + (openStation === key ? " active" : "");
    // Stamp pop-in only the moment a plate is newly served (not on every re-render or first load).
    if (status === "served") {
      const stampId = `${currentChild}:${key}:${currentWeek()}`;
      if (!stampedStations.has(stampId)) {
        stampedStations.add(stampId);
        if (stampSeeded[currentChild]) card.classList.add("stamp-new");
      }
    }
    card.onclick = () => openStationFn(key);
    const active = activeTasks(key);
    const doneN = active.filter(t => state[key].tasks[t.id].done).length;
    const unlockedTasks = active.filter(t => !isTaskLocked(t));
    const lockedCount = active.length - unlockedTasks.length;
    const stampText = status === "served" ? "SERVED ✓" : status === "burning" ? "BURNING 🔥" : "";
    card.innerHTML = `
      ${sentBackTasks.length > 0 ? `<div class="sentback-badge">🔥 Try again (${sentBackTasks.length})</div>` : (currentView === "parent" && needsReview ? '<div class="review-badge">Review</div>' : "")}
      <div class="station-tag">${subjectTag(key)}</div>
      <div class="station-title">${DATA[key].name}</div>
      <div class="station-lesson">${active.length === 0 ? "Nothing loaded for this week yet" : `${unlockedTasks.length} item${unlockedTasks.length !== 1 ? "s" : ""} available${lockedCount > 0 ? ` · +${lockedCount} locked` : ""}`}</div>
      <div class="station-footer">
        <span class="task-count">${active.length === 0 ? "—" : `${doneN}/${unlockedTasks.length} done`}</span>
        <span class="stamp">${stampText}</span>
      </div>`;
    if (!showingWeekReport) grid.appendChild(card);
  });

  document.getElementById("weekNumberText").innerHTML = `<span class="due-status ${isMonthlyTestWeek() ? "ok" : ""}">Week ${currentWeek()}${isMonthlyTestWeek() ? " — test week!" : ""}</span>`;

  const advanceBanner = document.getElementById("advanceBanner");
  // Only the parent advances the week. Kids see a "waiting" message once everything is served.
  const served = allSubjectsServed();
  if (currentView !== "parent") {
    advanceBanner.innerHTML = served
      ? `<div class="advance-banner">
          <div><div class="advance-banner-title">🎉 All done with Week ${currentWeek()}!</div>
          <div class="advance-banner-sub">Nice work! Waiting for Mom to check everything, then you'll move on to Week ${currentWeek() + 1}.</div></div>
        </div>`
      : "";
  } else {
    const nextWk = currentWeek() + 1, name = CHILD_META[currentChild].name;
    if (served) {
      const pending = pendingReviewCount();
      advanceBanner.innerHTML = `<div class="advance-banner">
        ${pending > 0
          ? `<div><div class="advance-banner-title">✅ ${name} finished every section of Week ${currentWeek()}</div><div class="advance-banner-sub">${pending} written answer${pending > 1 ? "s are" : " is"} still waiting on your review (see "Waiting on You" above), then you can advance.</div></div>
             <button class="btn" disabled>Advance ${name} to Week ${nextWk}</button>`
          : `<div class="advance-banner-title">✅ ${name} finished every section of Week ${currentWeek()}</div><button class="btn primary" onclick="advanceWeek()">Advance ${name} to Week ${nextWk}</button>`}
      </div>`;
    } else {
      advanceBanner.innerHTML = `<div class="advance-banner advance-banner-quiet">
        <div class="advance-banner-sub">Week ${currentWeek()} isn't fully finished yet.</div>
        <button class="btn" onclick="advanceAnyway()">Advance ${name} to Week ${nextWk} anyway</button>
      </div>`;
    }
  }

  const termCtrl = document.getElementById("termFinalControl");
  termCtrl.style.display = currentView === "parent" ? "flex" : "none";
  document.getElementById("termFinalBtn").textContent = settings.termFinalsUnlocked ? "Re-lock term finals" : "Unlock term finals";

  const monthCtrl = document.getElementById("monthlyTestControl");
  monthCtrl.style.display = currentView === "parent" ? "flex" : "none";
  const modeText = document.getElementById("monthlyTestModeText");
  const modeBtn = document.getElementById("monthlyTestBtn");
  if (settings.monthlyTestOverride === null) {
    modeText.textContent = `Monthly Tests: automatic — unlocked every 4th week (next: Week ${nextMonthlyTestWeek()})`;
    modeBtn.textContent = "Override schedule";
  } else if (settings.monthlyTestOverride === false) {
    modeText.textContent = "Monthly Tests: forced available";
    modeBtn.textContent = "Force locked instead";
  } else {
    modeText.textContent = "Monthly Tests: forced locked";
    modeBtn.textContent = "Reset to automatic schedule";
  }
  stampSeeded[currentChild] = true;
  document.getElementById("progressFill").style.width = (doneCount / keys.length * 100) + "%";
  document.getElementById("progressLabel").textContent = `${doneCount} / ${keys.length} plates served`;

  const sentBackAll = [];
  keys.forEach(key => {
    activeTasks(key).forEach(t => {
      if (state[key].tasks[t.id].sentBack && !state[key].tasks[t.id].done) sentBackAll.push({ key, label: t.label, subject: DATA[key].name });
    });
  });
  const banner = document.getElementById("sentBackBanner");
  banner.innerHTML = sentBackAll.length > 0 ? `<div class="sentback-banner">
      <div class="sentback-banner-title">🔥 ${sentBackAll.length} plate${sentBackAll.length > 1 ? "s need" : " needs"} another try! Tap to fix ${sentBackAll.length > 1 ? "them" : "it"} up:</div>
      <div class="sentback-chips">
        ${sentBackAll.map(item => `<span class="sentback-chip" onclick="openStationFn('${item.key}')">${item.subject}: ${item.label}</span>`).join("")}
      </div>
    </div>` : "";

  const panel = document.getElementById("detailPanel");
  if (showingWeekReport) {
    panel.className = "detail"; panel.innerHTML = "";
  } else if (openStation) {
    const d = DATA[openStation];
    const status = stationStatus(openStation);
    const active = activeTasks(openStation);
    let rows = active.length === 0
      ? `<div class="empty-note">Nothing loaded for Week ${currentWeek()} yet — check back once new content is added.</div>`
      : active.map(t => `<div class="task-row">${taskHeadHTML(openStation, t)}${isTaskLocked(t) ? "" : taskBodyHTML(openStation, t)}</div>`).join("");
    let note = "";
    if (status === "burning") note = `<div class="burn-note">🔥 Burning — scored below 70%. Redoing will reset this section and log the original scores for your review.</div>`;
    panel.className = "detail open";
    panel.innerHTML = `<div class="detail-head">
        <div><div class="detail-tag">${subjectTag(openStation)}</div><div class="detail-title">${d.name}</div>${note}</div>
        <div style="display:flex;gap:8px;align-items:flex-start;">
          ${status === "burning" ? `<button class="btn" style="margin-top:0;background:var(--saffron);border-color:var(--saffron);color:#5B4636;" onclick="redoStation('${openStation}')">Redo this section</button>` : ""}
          <button class="detail-close" onclick="openStationFn('${openStation}')">✕ close</button>
        </div>
      </div>${rows}`;
  } else { panel.className = "detail"; panel.innerHTML = ""; }

  if (currentView === "parent") {
    const pool = loadPool();
    const bank = document.getElementById("reviewBank");
    const active = pool.filter(p => p.status === "active");
    const mastered = pool.filter(p => p.status === "mastered");
    bank.innerHTML = pool.length === 0
      ? `<div class="empty-note">No missed words logged yet.</div>`
      : `<div class="lesson-text" style="font-size:0.82rem;">
        <b>Active (${active.length}):</b> ${active.map(p => `${p.word} (missed ${p.timesMissed}×, from ${p.lastSeen})`).join(", ") || "none"}<br>
        <b>Mastered (${mastered.length}):</b> ${mastered.map(p => p.word).join(", ") || "none"}
      </div>`;

    const answerLog = (answerLogCache[currentChild] || []).slice(0, 40);
    const answerLogList = document.getElementById("answerLogList");
    answerLogList.innerHTML = answerLog.length === 0
      ? `<div class="empty-note">No graded answers logged yet.</div>`
      : answerLog.map(a => `
        <div class="review-item answer-log-item ${a.correct ? "answer-log-correct" : "answer-log-wrong"}">
          <strong>${a.correct ? "✅" : "❌"} ${a.word}</strong>
          <div class="meta">${a.game || ""}${a.subject ? " · " + a.subject : ""} · ${a.timestamp ? new Date(a.timestamp).toLocaleString() : ""}</div>
          ${a.question ? `<div class="submitted-text">${a.question}</div>` : ""}
          ${!a.correct ? `<div class="submitted-text">Answered: "${a.givenAnswer}" — Correct answer: "${a.correctAnswer}"</div>` : ""}
        </div>`).join("");

    const burnLog = burnLogCache[currentChild] || [];
    const burnList = document.getElementById("burnLogList");
    burnList.innerHTML = burnLog.length === 0
      ? `<div class="empty-note">No sections have needed a redo yet.</div>`
      : burnLog.map(rec => `
        <div class="review-item">
          <strong>${rec.station} — ${rec.tag}</strong>
          <div class="meta">${rec.date} · ${rec.reason}</div>
          <div class="submitted-text">${rec.items.join("\n")}</div>
        </div>`).join("");

    let items = [];
    Object.keys(DATA).forEach(key => {
      DATA[key].tasks.forEach(t => {
        const s = state[key].tasks[t.id];
        if (s.needsReview && !s.reviewed) items.push({ key, t, s });
      });
    });
    const list = document.getElementById("reviewList");
    document.getElementById("waitingOnYou").style.display = items.length === 0 ? "none" : "block"; // only takes space at the top when something needs you
    list.innerHTML = items.length === 0
      ? ""
      : `<div class="lesson-text" style="font-size:0.82rem;">Review each one in the week outline below, right next to the lesson it answers (marked in blue). Tap to jump to it:</div>
        <div class="sentback-chips">${items.map(({ key, t }) => `<span class="review-jump-chip" onclick="jumpToReview('${key}','${t.id}',${Number(t.week_number) || 1})">Week ${t.week_number} · ${DATA[key].name}: ${t.label}</span>`).join("")}</div>`;
  }

}

// Keeps any open tap-to-tag popup (pos-tagger) fully on-screen, regardless
// of where its word sits in the wrapped sentence layout — a fixed-position
// popup placed with getBoundingClientRect math instead of pure CSS
// anchoring, since CSS-anchored popups near a screen edge were getting
// clipped/hidden.
function positionOpenPopup() {
  const popup = document.querySelector(".pos-popup");
  if (!popup) return;
  const anchor = popup.parentElement;
  const wordRect = anchor.getBoundingClientRect();
  popup.style.position = "fixed";
  popup.style.margin = "0";
  const popupRect = popup.getBoundingClientRect();
  const margin = 8;
  let left = wordRect.left + wordRect.width / 2 - popupRect.width / 2;
  left = Math.max(margin, Math.min(left, window.innerWidth - popupRect.width - margin));
  let top = wordRect.bottom + 4;
  if (top + popupRect.height + margin > window.innerHeight) {
    top = wordRect.top - popupRect.height - 4;
    if (top < margin) top = margin;
  }
  popup.style.left = `${left}px`;
  popup.style.top = `${top}px`;
}

init();
