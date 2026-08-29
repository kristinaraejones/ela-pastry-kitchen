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
let phraseRangeStart = null;
let pendingPhraseRange = null;

let DATA = null;   // current child's { subjectKey: {name, tag, tasks:[...]} }
let state = null;  // current child's { subjectKey: { tasks: { taskId: {...} } } }
let settings = { currentWeekNumber: 1, dueDate: new Date(), termFinalsUnlocked: false, monthlyTestOverride: null };

const childrenCache = {};   // { kenley: {DATA, state} }
const reviewPoolCache = {}; // { kenley: [ {word,timesMissed,...} ] }
const markersCache = {};    // { kenley: {vocab: 3, ...} }
const burnLogCache = {};    // { kenley: [ {...} ] }
const banksCache = {};      // { kenley: {vocab:[...], spelling:[...], ...} }

// ---------- API layer ----------

function resolveApiUrl() {
  return localStorage.getItem("elaApiUrl") || (typeof DEFAULT_API_URL !== "undefined" ? DEFAULT_API_URL : "") || "";
}

async function apiGetBootstrap(student) {
  const res = await fetch(`${API_URL}?action=bootstrap&student=${encodeURIComponent(student)}`);
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || "Bootstrap failed");
  return json;
}

async function apiPost(action, payload) {
  try {
    const res = await fetch(API_URL, { method: "POST", body: JSON.stringify({ action, ...payload }) });
    const json = await res.json();
    if (!json.ok) throw new Error(json.error || "Request failed");
    clearSyncError();
    return json;
  } catch (err) {
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
  const el = document.getElementById("syncBanner");
  if (el) el.style.display = "none";
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
  markersCache[student] = resp.markers;
  burnLogCache[student] = resp.burnLog;
  banksCache[student] = resp.banks;
}

function parseSettings(raw) {
  raw = raw || {};
  const override = raw.monthlyTestOverride === "true" ? true : raw.monthlyTestOverride === "false" ? false : null;
  // Expect a plain "yyyy-MM-dd" string, but tolerate a full ISO timestamp
  // too (e.g. an older backend that hadn't normalized a Sheets-auto-converted
  // Date cell yet) so we never silently produce an Invalid Date/NaN.
  let dueDate;
  if (raw.dueDate) {
    dueDate = new Date(raw.dueDate.includes("T") ? raw.dueDate : raw.dueDate + "T00:00:00");
    if (isNaN(dueDate)) { dueDate = new Date(); dueDate.setDate(dueDate.getDate() + 4); }
  } else {
    dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + 4);
  }
  return {
    currentWeekNumber: Number(raw.currentWeekNumber) || 1,
    dueDate,
    termFinalsUnlocked: raw.termFinalsUnlocked === "true" || raw.termFinalsUnlocked === true,
    monthlyTestOverride: override
  };
}

function isoDateOnly(d) {
  return d.toISOString().slice(0, 10);
}

function buildChildFromBootstrap(resp) {
  const DATA = {};
  SUBJECT_ORDER.forEach(subjectKey => {
    const tasksForSubject = resp.schedule.filter(t => t.subject_key === subjectKey);
    if (tasksForSubject.length === 0) return;
    DATA[subjectKey] = {
      name: tasksForSubject[0].subject_name,
      tag: tasksForSubject[0].subject_tag,
      tasks: tasksForSubject.map(t => Object.assign(
        { id: t.id, type: t.type, label: t.label, dynamic: t.dynamic, termFinal: t.termFinal, monthlyTest: t.monthlyTest },
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
      state[key].tasks[t.id] = base;
    });
  });
  return { DATA, state };
}

function deriveStatus(s) {
  if (s.sentBack) return "sent_back";
  if (s.reviewed) return "reviewed";
  if (s.needsReview) return "needs_review";
  if (s.done) return "complete";
  return "not_started";
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
  openStation = null; posPopupOpen = null; phraseRangeStart = null; pendingPhraseRange = null;
  render();
}

// ---------- Settings (shared, global) ----------

function daysUntilDue() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return Math.round((settings.dueDate - today) / 86400000);
}
function adjustDueDate(delta) {
  settings.dueDate.setDate(settings.dueDate.getDate() + delta);
  apiPost("saveSetting", { key: "dueDate", value: isoDateOnly(settings.dueDate) }).catch(() => {});
  render();
}
function isMonthlyTestWeek() { return settings.currentWeekNumber % 4 === 0; }
function nextMonthlyTestWeek() {
  return isMonthlyTestWeek() ? settings.currentWeekNumber : settings.currentWeekNumber + (4 - (settings.currentWeekNumber % 4));
}
function adjustWeekNumber(delta) {
  settings.currentWeekNumber = Math.max(1, settings.currentWeekNumber + delta);
  apiPost("saveSetting", { key: "currentWeekNumber", value: String(settings.currentWeekNumber) }).catch(() => {});
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

function stationScorePct(key) {
  let sumX = 0, sumY = 0;
  DATA[key].tasks.forEach(t => {
    const s = state[key].tasks[t.id];
    if (s.done && s.score) {
      const parts = s.score.split("/").map(Number);
      sumX += parts[0]; sumY += parts[1];
    }
  });
  return sumY > 0 ? sumX / sumY : null;
}
function stationDone(key) {
  return DATA[key].tasks.filter(t => !isTaskLocked(t)).every(t => state[key].tasks[t.id].done);
}
function stationStatus(key) {
  if (stationDone(key)) {
    const pct = stationScorePct(key);
    if (pct !== null && pct < 0.70) return "burning";
    return "served";
  }
  const days = daysUntilDue();
  if (days < 0) return "burnt";
  if (days <= 3) return "burning";
  return "progress";
}
function redoStation(key) {
  const items = DATA[key].tasks.map(t => {
    const s = state[key].tasks[t.id];
    if (s.done && s.score) return `${t.label}: scored ${s.score}`;
    if (s.done && s.needsReview) return `${t.label}: submitted, was awaiting your review`;
    if (s.done) return `${t.label}: completed`;
    return `${t.label}: not completed`;
  });
  const record = {
    station: DATA[key].name, tag: DATA[key].tag,
    date: new Date().toLocaleDateString(),
    reason: `Scored below 70% (${Math.round((stationScorePct(key) || 0) * 100)}%)`,
    items
  };
  burnLogCache[currentChild].unshift(record);
  apiPost("addBurnLog", { student: currentChild, ...record }).catch(() => {});
  DATA[key].tasks.forEach(t => {
    state[key].tasks[t.id] = { open: false, done: false, needsReview: false, reviewed: false, answers: {}, score: null, results: null };
    persistTask(key, t.id);
  });
  render();
}

function toggleView() {
  currentView = currentView === "kenley" ? "parent" : "kenley";
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

function speakWord(word) {
  if (!("speechSynthesis" in window)) return;
  const u = new SpeechSynthesisUtterance(word);
  u.rate = 0.85;
  window.speechSynthesis.cancel();
  window.speechSynthesis.speak(u);
}
function speakSequence(parts) {
  if (!("speechSynthesis" in window)) return;
  window.speechSynthesis.cancel();
  parts.forEach(text => {
    const u = new SpeechSynthesisUtterance(text);
    u.rate = 0.85;
    window.speechSynthesis.speak(u);
  });
}
function speakWordInContext(word, context) { speakSequence([word, context, word]); }

function gradeSentence(correctText, typedText) {
  const tokenize = s => s.match(/[A-Za-z']+/g) || [];
  const correctWords = tokenize(correctText);
  const typedWords = tokenize(typedText);
  const maxLen = Math.max(correctWords.length, typedWords.length);
  const wordResults = [];
  for (let i = 0; i < maxLen; i++) {
    const cw = correctWords[i] || "";
    const tw = typedWords[i] || "";
    wordResults.push({ correctWord: cw, typedWord: tw, spellingCorrect: cw !== "" && cw.toLowerCase() === tw.toLowerCase() });
  }
  const correctTrim = correctText.trim(), typedTrim = typedText.trim();
  const startsCapCorrect = /^[A-Z]/.test(correctTrim);
  const startsCapTyped = /^[A-Z]/.test(typedTrim);
  const endPunctCorrect = (correctTrim.match(/[.!?]$/) || [])[0] || "(none)";
  const endPunctTyped = (typedTrim.match(/[.!?]$/) || [])[0] || "(none)";
  const grammar = [
    { label: "Capital letter to start the sentence", pass: !startsCapCorrect || startsCapTyped },
    { label: `Ending punctuation (${endPunctCorrect})`, pass: endPunctCorrect === endPunctTyped }
  ];
  return {
    wordResults, grammar,
    allSpellingCorrect: wordResults.every(w => w.spellingCorrect),
    allGrammarPass: grammar.every(g => g.pass)
  };
}

function checkDictation(key, id) {
  const t = DATA[key].tasks.find(x => x.id === id);
  const words = t.dynamic === "reviewPool" ? state[key].tasks[id]._reviewWords : t.words;
  let correctCount = 0;
  const results = [];
  words.forEach((w, i) => {
    if (w.kind === "sentence") {
      const typed = document.getElementById(`dict-${key}-${id}-${i}`).value.trim();
      const grade = gradeSentence(w.answer, typed);
      const isRight = grade.allSpellingCorrect && grade.allGrammarPass;
      if (isRight) correctCount++;
      results.push({ kind: "sentence", typed, answer: w.answer, grade });
      grade.wordResults.forEach(wr => {
        if (!wr.spellingCorrect && wr.correctWord) logMiss(wr.correctWord, DATA[key].tag + " (in a sentence)", null);
      });
    } else {
      const typed = document.getElementById(`dict-${key}-${id}-${i}`).value.trim();
      const isRight = typed.toLowerCase() === w.answer.toLowerCase();
      if (isRight) correctCount++;
      results.push({ kind: "word", typed, correct: isRight, answer: w.answer, context: w.context || null });
      if (t.dynamic === "reviewPool") {
        logReviewResult(w.answer, isRight);
      } else if (!isRight) {
        const source = t.dynamic === "spellingMonthBank" ? "Monthly Test" : t.dynamic === "examSpellingBank" ? "Term Exam" : DATA[key].tag;
        logMiss(w.answer, source, w.context);
      }
    }
  });
  state[key].tasks[id].results = results;
  state[key].tasks[id].score = `${correctCount}/${words.length}`;
  state[key].tasks[id].done = true;
  if (t.dynamic === "spellingMonthBank") markTested(banksCache[currentChild].spelling || [], "spelling");
  persistTask(key, id);
  render();
}
function tagAbbrev(type) { return TAG_ABBREV[type] || type.slice(0, 4).toUpperCase(); }

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
  s.labels.forEach((l, i) => { if (l === t.answers[i]) correct++; });
  s.score = `${correct}/${t.answers.length}`;
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
function checkPhraseTagging(key, id) {
  const t = DATA[key].tasks.find(x => x.id === id);
  const s = state[key].tasks[id];
  let correct = 0;
  s.selections.forEach(sel => { if (t.phrases.some(p => p.start === sel.start && p.end === sel.end && p.type === sel.type)) correct++; });
  s.score = `${correct}/${t.phrases.length}`;
  s.done = true;
  persistTask(key, id);
  render();
}
function checkFill(key, id) {
  const t = DATA[key].tasks.find(x => x.id === id);
  let correct = 0;
  t.words.forEach((w, i) => {
    const v = document.getElementById(`fill-${key}-${id}-${i}`).value.trim().toLowerCase();
    if (v === w.answer.toLowerCase()) correct++;
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
  questions.forEach((q, i) => { if (answers[i] === q.correct) correct++; });
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
  if (t.type === "read") {
    inner = t.content + (s.done ? `` : `<button class="btn primary" onclick="markRead('${key}','${t.id}')">Mark as read</button>`);
  } else if (t.type === "external") {
    inner = `<a class="ext-link" href="#" onclick="return false;">${t.linkText} ↗</a>
      <div class="lesson-text" style="opacity:.75;font-size:0.78rem;">${t.note}</div>
      <label style="font-size:0.82rem;display:flex;align-items:center;gap:8px;">
        <input type="checkbox" ${s.done ? "checked" : ""} onchange="markExternal('${key}','${t.id}',this.checked)"> Mark complete
      </label>`;
  } else if (t.type === "reflection") {
    const feedbackNote = s.parentComment ? `<div class="parent-feedback">📝 ${s.reviewed ? "Feedback from parent:" : "Refired — please revise:"} ${s.parentComment}</div>` : "";
    inner = `<div class="lesson-text"><p>${t.prompt}</p></div>
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
    inner = `<div class="lesson-text"><p>${t.prompt}</p></div>`;
    if (s.done && s.results) {
      inner += s.results.map(r => {
        if (r.kind === "sentence") {
          const chips = r.grade.wordResults.map(wr => `<span class="word-chip ${wr.spellingCorrect ? "chip-correct" : "chip-incorrect"}">${wr.typedWord || "—"}</span>`).join(" ");
          const missedWords = r.grade.wordResults.filter(wr => !wr.spellingCorrect && wr.correctWord).map(wr => wr.correctWord);
          const overallClass = (r.grade.allSpellingCorrect && r.grade.allGrammarPass) ? "correct" : "incorrect";
          return `<div class="dict-result ${overallClass}">
            <div class="dict-result-typed" style="margin-bottom:6px;">Spelling, word by word:</div>
            <div style="margin-bottom:8px;">${chips}</div>
            ${missedWords.length ? `<div class="dict-result-answer">Correct spelling for missed word${missedWords.length > 1 ? "s" : ""}: <b>${missedWords.join(", ")}</b></div>` : ``}
            <div class="grammar-check">
              ${r.grade.grammar.map(g => `<div class="grammar-line ${g.pass ? "pass" : "fail"}">${g.pass ? "✓" : "✗"} ${g.label}</div>`).join("")}
            </div>
          </div>`;
        } else {
          const icon = r.correct ? "✓" : "✗";
          return `<div class="dict-result ${r.correct ? "correct" : "incorrect"}">
            <div class="dict-result-typed"><span class="dict-icon">${icon}</span> You wrote: <b>${r.typed || "(blank)"}</b></div>
            ${!r.correct ? `<div class="dict-result-answer">Correct spelling: <b>${r.answer}</b></div>` : ``}
          </div>`;
        }
      }).join("");
      inner += `<div class="score-result ${s.results.every(r => r.kind === "sentence" ? (r.grade.allSpellingCorrect && r.grade.allGrammarPass) : r.correct) ? "pass" : "retry"}">Scored automatically: ${s.score}${t.dynamic === "reviewPool" ? " — pool updated, mastered words drop out automatically" : ""}. Any missed word — even outside this week's list — has been added to the review file.</div>`;
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
  } else if (t.type === "pos-tagger") {
    inner = `<div class="lesson-text"><p>Tap a word, then tap its part of speech.</p></div><div class="pos-row">`;
    t.sentence.forEach((word, i) => {
      const label = s.labels[i];
      let cls = "", shownLabel = label;
      if (s.done) {
        cls = label === t.answers[i] ? "correct" : "incorrect";
        if (label !== t.answers[i]) shownLabel = `${label || "—"} → ${t.answers[i]}`;
      }
      const popupOpen = !s.done && posPopupOpen && posPopupOpen.key === key && posPopupOpen.id === t.id && posPopupOpen.idx === i;
      inner += `<div class="word-slot" onclick="${s.done ? "" : `openPosPopup('${key}','${t.id}',${i})`}">
        <div class="word-text">${word}</div>
        <div class="word-label ${cls}">${shownLabel || "+ tag"}</div>
        ${popupOpen ? `<div class="pos-popup">${t.options.map(o => `<button onclick="event.stopPropagation();selectPos('${key}','${t.id}',${i},'${o}')">${o}</button>`).join("")}</div>` : ""}
      </div>`;
    });
    inner += `</div>`;
    if (s.done) {
      const misses = t.sentence.map((w, i) => i).filter(i => s.labels[i] !== t.answers[i]);
      if (misses.length) {
        inner += `<div class="tag-review">${misses.map(i => `<div class="tag-review-item"><b>${t.sentence[i]}</b> — you said ${s.labels[i] || "nothing"}, it's actually <b>${t.answers[i]}</b>. ${t.explanations[i]}</div>`).join("")}</div>`;
      }
      inner += `<div class="score-result ${s.score.split("/")[0] === s.score.split("/")[1] ? "pass" : "retry"}">Scored automatically: ${s.score}</div>`;
    } else {
      const allLabeled = s.labels.every(l => l !== null);
      inner += `<button class="btn primary" ${allLabeled ? "" : "disabled"} onclick="checkPosTagging('${key}','${t.id}')">Check my tagging</button>`;
    }
  } else if (t.type === "phrase-tagger") {
    inner = `<div class="lesson-text"><p>Tap the first word, then the last word of a chunk, then choose what it is. (Tap the same word twice for a single-word chunk.)</p></div><div class="pos-row">`;
    t.sentence.forEach((word, i) => {
      const sel = s.selections.find(sel => i >= sel.start && i <= sel.end);
      let cls = "", tagLabel = "";
      if (sel) {
        tagLabel = tagAbbrev(sel.type);
        if (s.done) {
          const match = t.phrases.some(p => p.start === sel.start && p.end === sel.end && p.type === sel.type);
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
        const exact = t.phrases.find(p => p.start === sel.start && p.end === sel.end && p.type === sel.type);
        if (!exact) {
          const text = t.sentence.slice(sel.start, sel.end + 1).join(" ");
          const sameSpan = t.phrases.find(p => p.start === sel.start && p.end === sel.end);
          if (sameSpan) reviewLines.push(`<b>"${text}"</b> — you tagged it ${sel.type}, but it's actually <b>${sameSpan.type}</b>. ${sameSpan.explanation}`);
          else reviewLines.push(`<b>"${text}"</b> — that's not one of the parts we're looking for here.`);
        }
      });
      t.phrases.forEach(p => {
        const found = s.selections.some(sel => sel.start === p.start && sel.end === p.end);
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
      inner += `<div class="lesson-text"><p>${qi + 1}. ${q.q}</p></div><div class="mc-options">`;
      q.options.forEach((opt, oi) => {
        const sel = s.answers.mc && s.answers.mc[qi] === oi;
        let cls = sel ? "selected" : "";
        if (s.done) {
          if (oi === q.correct) cls = "correct";
          else if (sel && oi !== q.correct) cls = "incorrect";
        }
        inner += `<div class="mc-option ${cls}" onclick="selectMC('${key}','${t.id}',${qi},${oi})">${opt}</div>`;
      });
      inner += `</div>`;
      if (s.done && q.explanation) inner += `<div class="tag-review-item" style="margin-top:6px;">${q.explanation}</div>`;
    });
    if (s.done) {
      const parts = s.score.split("/");
      inner += `<div class="score-result ${parts[0] === parts[1] || (parseInt(parts[0]) / parseInt(parts[1])) >= 0.7 ? "pass" : "retry"}">Scored automatically: ${s.score}</div>`;
    } else {
      inner += `<button class="btn primary" onclick="submitMC('${key}','${t.id}')">Submit answers</button>`;
    }
  }
  return `<div class="task-body ${s.open ? "open" : ""}">${inner}</div>`;
}

function render() {
  document.getElementById("childSwitcher").innerHTML = Object.keys(CHILD_META).map(id =>
    `<button class="child-pill ${currentChild === id ? "active" : ""}" onclick="switchChild('${id}')">${CHILD_META[id].name}</button>`
  ).join("");
  document.getElementById("boardSub").textContent = CHILD_META[currentChild].subtitle;

  document.getElementById("viewToggle").classList.toggle("parent", currentView === "parent");
  document.getElementById("toggleKnob").textContent = currentView === "parent" ? "PARENT" : CHILD_META[currentChild].name.toUpperCase();
  document.getElementById("reviewQueue").style.display = currentView === "parent" ? "block" : "none";

  const grid = document.getElementById("stationsGrid");
  grid.innerHTML = "";
  let doneCount = 0;
  const keys = Object.keys(DATA);
  keys.forEach(key => {
    const status = stationStatus(key);
    if (stationDone(key)) doneCount++; // only count actually-completed stations, not ones merely "burning" from being incomplete-and-due-soon
    const needsReview = DATA[key].tasks.some(t => state[key].tasks[t.id].needsReview && !state[key].tasks[t.id].reviewed);
    const sentBackTasks = DATA[key].tasks.filter(t => state[key].tasks[t.id].sentBack && !state[key].tasks[t.id].done);
    const card = document.createElement("div");
    card.className = "station" + (status === "served" ? " done" : "") + (status === "burning" ? " burning" : "") + (status === "burnt" ? " burnt" : "") + (openStation === key ? " active" : "");
    card.onclick = () => openStationFn(key);
    const doneN = DATA[key].tasks.filter(t => state[key].tasks[t.id].done).length;
    const unlockedTasks = DATA[key].tasks.filter(t => !isTaskLocked(t));
    const lockedCount = DATA[key].tasks.length - unlockedTasks.length;
    const stampText = status === "served" ? "SERVED ✓" : status === "burning" ? "BURNING 🔥" : status === "burnt" ? "BURNT ⚫" : "";
    card.innerHTML = `
      ${sentBackTasks.length > 0 ? `<div class="sentback-badge">🔁 Refire (${sentBackTasks.length})</div>` : (currentView === "parent" && needsReview ? '<div class="review-badge">Review</div>' : "")}
      <div class="station-tag">${DATA[key].tag}</div>
      <div class="station-title">${DATA[key].name}</div>
      <div class="station-lesson">${unlockedTasks.length} item${unlockedTasks.length > 1 ? "s" : ""} available${lockedCount > 0 ? ` · +${lockedCount} locked` : ""}</div>
      <div class="station-footer">
        <span class="task-count">${doneN}/${unlockedTasks.length} done</span>
        <span class="stamp">${stampText}</span>
      </div>`;
    grid.appendChild(card);
  });

  const days = daysUntilDue();
  const dueEl = document.getElementById("dueStatusText");
  if (days < 0) dueEl.innerHTML = `<span class="due-status late">${Math.abs(days)} day(s) PAST due</span>`;
  else if (days <= 3) dueEl.innerHTML = `<span class="due-status warn">${days} day(s) until due</span>`;
  else dueEl.innerHTML = `<span class="due-status ok">${days} day(s) until due</span>`;

  document.getElementById("weekNumberText").innerHTML = `<span class="due-status ${isMonthlyTestWeek() ? "ok" : ""}">Week ${settings.currentWeekNumber}${isMonthlyTestWeek() ? " — test week!" : ""}</span>`;

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
  document.getElementById("progressFill").style.width = (doneCount / keys.length * 100) + "%";
  document.getElementById("progressLabel").textContent = `${doneCount} / ${keys.length} plates served`;

  const sentBackAll = [];
  keys.forEach(key => {
    DATA[key].tasks.forEach(t => {
      if (state[key].tasks[t.id].sentBack && !state[key].tasks[t.id].done) sentBackAll.push({ key, label: t.label, subject: DATA[key].name });
    });
  });
  const banner = document.getElementById("sentBackBanner");
  banner.innerHTML = sentBackAll.length > 0 ? `<div class="sentback-banner">
      <div class="sentback-banner-title">🔁 ${sentBackAll.length} plate${sentBackAll.length > 1 ? "s" : ""} sent back to refire — review before moving on:</div>
      <div class="sentback-chips">
        ${sentBackAll.map(item => `<span class="sentback-chip" onclick="openStationFn('${item.key}')">${item.subject}: ${item.label}</span>`).join("")}
      </div>
    </div>` : "";

  const panel = document.getElementById("detailPanel");
  if (openStation) {
    const d = DATA[openStation];
    const status = stationStatus(openStation);
    let rows = "";
    d.tasks.forEach(t => { rows += `<div class="task-row">${taskHeadHTML(openStation, t)}${isTaskLocked(t) ? "" : taskBodyHTML(openStation, t)}</div>`; });
    let note = "";
    const days2 = daysUntilDue();
    if (status === "burning" && stationDone(openStation)) note = `<div class="burn-note">🔥 Burning — scored below 70%. Redoing will reset this section and log the original scores for your review.</div>`;
    else if (status === "burning") note = `<div class="burn-note" style="color:var(--saffron);">🔥 Burning — due in ${days2} day(s) and not yet complete.</div>`;
    else if (status === "burnt") note = `<div class="burn-note">⚫ Burnt — this was due ${Math.abs(days2)} day(s) ago and is still incomplete. Flagged for you in Parent view.</div>`;
    panel.className = "detail open";
    panel.innerHTML = `<div class="detail-head">
        <div><div class="detail-tag">${d.tag}</div><div class="detail-title">${d.name}</div>${note}</div>
        <div style="display:flex;gap:8px;align-items:flex-start;">
          ${status === "burning" && stationDone(openStation) ? `<button class="btn" style="margin-top:0;background:var(--saffron);border-color:var(--saffron);color:#5B4636;" onclick="redoStation('${openStation}')">Redo this section</button>` : ""}
          <button class="detail-close" onclick="openStationFn('${openStation}')">✕ close</button>
        </div>
      </div>${rows}`;
  } else { panel.className = "detail"; panel.innerHTML = ""; }

  if (currentView === "parent") {
    const overdueList = document.getElementById("overdueList");
    const burntStations = Object.keys(DATA).filter(k => stationStatus(k) === "burnt");
    const days3 = daysUntilDue();
    overdueList.innerHTML = burntStations.length === 0
      ? `<div class="empty-note">Nothing overdue right now.</div>`
      : burntStations.map(key => {
          const incomplete = DATA[key].tasks.filter(t => !state[key].tasks[t.id].done).map(t => t.label);
          return `<div class="review-item">
            <strong>${DATA[key].name} — ${DATA[key].tag}</strong>
            <div class="meta">${Math.abs(days3)} day(s) past due · still missing:</div>
            <div class="submitted-text">${incomplete.join("\n") || "(nothing recorded)"}</div>
          </div>`;
        }).join("");

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
    list.innerHTML = items.length === 0
      ? `<div class="empty-note">Nothing waiting on you right now.</div>`
      : items.map(({ key, t, s }) => `
        <div class="review-item">
          <strong>${DATA[key].name} — ${t.label}</strong>
          <div class="meta">Submitted by ${CHILD_META[currentChild].name}, awaiting review</div>
          <div class="submitted-text">${s.answers.text}</div>
          <textarea id="comment-${key}-${t.id}" placeholder="Optional feedback (shown either way — required reading if you send it back)" style="min-height:50px;margin-top:6px;"></textarea>
          <div class="review-actions">
            <button onclick="sendBackReflection('${key}','${t.id}')">Refire (send back)</button>
            <button class="approve" onclick="approveReflection('${key}','${t.id}')">Approve</button>
          </div>
        </div>`).join("");
  }
}

init();
