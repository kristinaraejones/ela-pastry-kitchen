// The ELA Pastry Kitchen — design refresh layer (Stage 1).
// Loaded AFTER app.js on beta.html only. It replaces how things are drawn
// (render, task headers, the parent week report) and leaves all data, saving,
// grading and the offline queue in app.js untouched.

// ---------- Themes ----------
const BK_THEME = {
  kenley: { emblem: "croissant", accent: "#2F7A74", ink: "#1E504C", soft: "#DDF1EA", bar: "#4F9E91", tiles: ["#DDF1EA", "#DCEDF9"] },
  adelyn: { emblem: "macaron", accent: "#7A56A8", ink: "#553A7E", soft: "#F5E5F3", bar: "#9B7BC8", tiles: ["#FCE6EF", "#EEE6FA", "#E0F0FB"] }
};
function bkTheme(id) { return BK_THEME[id || currentChild] || BK_THEME.kenley; }
const BK_SUBJECT_PASTRY = { vocab: "macaron", spelling: "cookie", grammar: "cupcake", reading: "croissant", writing: "donut" };
function bkPastryFor(key) { return BK_SUBJECT_PASTRY[key] || "cookie"; }

// ---------- Pastry drawings (same art as the design canvas) ----------
const BK_PASTRY_COLORS = {
  kenley: { mac: ["#A8DCCB", "#5FA592"], wrap: ["#CFE9D8", "#7DB594"], frost: ["#BFE0F5", "#6EA6CF"], cherry: ["#7FAEDD", "#4E80B5"], icing: ["#A9D3EE", "#6AA0C9"], spr: ["#FFFFFF", "#4F9E91", "#F6D77C"] },
  adelyn: { mac: ["#F7BCD3", "#D97FA4"], wrap: ["#BFE0F5", "#7DAFD6"], frost: ["#D8C5F2", "#A283D1"], cherry: ["#F28CAB", "#C9627F"], icing: ["#F7BCD3", "#D97FA4"], spr: ["#FFFFFF", "#9B7BC8", "#6FB3E0"] }
};
function bkPastry(kind, child, size, extraAttr) {
  const c = BK_PASTRY_COLORS[child] || BK_PASTRY_COLORS.kenley;
  const sw = size <= 40 ? 3 : 2.5;
  const open = `<svg width="${size}" height="${size}" viewBox="0 0 64 64" aria-hidden="true" ${extraAttr || ""}>`;
  let body = "";
  if (kind === "croissant") {
    body = `<path d="M6 40C6 23 18 13 32 13S58 23 58 40C54 45 47 46 43 42C40 48 36 50 32 50S24 48 21 42C17 46 10 45 6 40Z" fill="#F6C98A" stroke="#C98B45" stroke-width="${sw}" stroke-linejoin="round"/><path d="M23 16L25 44M41 16L39 44${size > 40 ? "M12 25L18 42M52 25L46 42" : ""}" stroke="#C98B45" stroke-width="${sw}" stroke-linecap="round" fill="none"/>${size > 40 ? `<ellipse cx="30" cy="21" rx="4" ry="1.8" fill="#FCE3BC"/>` : ""}`;
  } else if (kind === "macaron") {
    body = `<path d="M9 31C9 20 19 14 32 14S55 20 55 31Z" fill="${c.mac[0]}" stroke="${c.mac[1]}" stroke-width="${sw}" stroke-linejoin="round"/><rect x="10" y="31" width="44" height="7" rx="3.5" fill="#FFF7EC" stroke="${c.mac[1]}" stroke-width="${sw}"/><path d="M9 38H55C55 46 45 51 32 51S9 46 9 38Z" fill="${c.mac[0]}" stroke="${c.mac[1]}" stroke-width="${sw}" stroke-linejoin="round"/>${size > 40 ? `<ellipse cx="24" cy="21" rx="5" ry="2.2" fill="#FFFFFF" opacity="0.6"/>` : ""}`;
  } else if (kind === "cookie") {
    body = `<circle cx="32" cy="32" r="21" fill="#EDC48E" stroke="#C0904F" stroke-width="${sw}"/><ellipse cx="24" cy="25" rx="3.2" ry="2.6" fill="#7A5236"/><ellipse cx="38" cy="23" rx="2.8" ry="2.3" fill="#7A5236"/><ellipse cx="42" cy="37" rx="3.2" ry="2.6" fill="#7A5236"/><ellipse cx="27" cy="40" rx="3" ry="2.4" fill="#7A5236"/>${size > 40 ? `<ellipse cx="33" cy="32" rx="2" ry="1.7" fill="#7A5236"/><ellipse cx="19" cy="34" rx="2" ry="1.7" fill="#7A5236"/>` : ""}`;
  } else if (kind === "cupcake") {
    body = `<path d="M15 35H49L45 56H19Z" fill="${c.wrap[0]}" stroke="${c.wrap[1]}" stroke-width="${sw}" stroke-linejoin="round"/>${size > 40 ? `<path d="M24 35L26 56M32 35V56M40 35L38 56" stroke="${c.wrap[1]}" stroke-width="2" fill="none"/>` : ""}<path d="M12 36C8 30 13 24 19 25C19 17 28 13 34 17C40 13 50 18 47 26C53 26 56 32 52 36Z" fill="${c.frost[0]}" stroke="${c.frost[1]}" stroke-width="${sw}" stroke-linejoin="round"/><circle cx="33" cy="12" r="4.5" fill="${c.cherry[0]}" stroke="${c.cherry[1]}" stroke-width="2"/>${size > 40 ? `<path d="M21 30l3-1.5M30 24l2.5 1.5M40 30l3 1M34 31l-1 2.5" stroke="#FFFFFF" stroke-width="2.2" stroke-linecap="round"/>` : ""}`;
  } else if (kind === "donut") {
    body = `<circle cx="32" cy="32" r="21" fill="#EFC593" stroke="#C69455" stroke-width="${sw}"/><circle cx="32" cy="32" r="16" fill="${c.icing[0]}" stroke="${c.icing[1]}" stroke-width="2.2"/><circle cx="32" cy="32" r="6.5" fill="#FFFFFF" stroke="#C69455" stroke-width="${sw}"/>${size > 40 ? `<path d="M24 22l3 1M44 32l2.5 1.5" stroke="${c.spr[0]}" stroke-width="2.4" stroke-linecap="round"/><path d="M39 21l-1 3M25 41l-2-2.5" stroke="${c.spr[1]}" stroke-width="2.4" stroke-linecap="round"/><path d="M38 42l3-1M20 30l-1 3" stroke="${c.spr[2]}" stroke-width="2.4" stroke-linecap="round"/>` : ""}`;
  }
  return open + body + "</svg>";
}

// ---------- Line icons (replace every emoji) ----------
const BK_ICON_PATHS = {
  check: `<path d="M5 12.5l4.5 4.5L19 7.5"/>`,
  timer: `<circle cx="12" cy="13.5" r="7.5"/><path d="M12 13.5V9.5M10 2.5h4"/>`,
  lock: `<rect x="5" y="11" width="14" height="9" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/>`,
  arrow: `<path d="M5 12h14M13 6l6 6-6 6"/>`,
  back: `<path d="M19 12H5M11 6l-6 6 6 6"/>`,
  close: `<path d="M6 6l12 12M18 6L6 18"/>`,
  chart: `<path d="M4 20V10M10 20V4M16 20v-7M21 20H3"/>`,
  download: `<path d="M12 4v11M7 10l5 5 5-5M5 20h14"/>`,
  speaker: `<path d="M4 9v6h4l5 4V5L8 9z"/><path d="M16 9a4 4 0 0 1 0 6M18.5 6.5a8 8 0 0 1 0 11"/>`,
  note: `<path d="M4 5h16v11H9l-5 4z"/>`,
  bulb: `<path d="M9 18h6M10 21h4"/><path d="M12 3a6 6 0 0 0-3.5 10.9c.6.5 1 1.2 1 2.1h5c0-.9.4-1.6 1-2.1A6 6 0 0 0 12 3z"/>`,
  key: `<circle cx="8" cy="15" r="4"/><path d="M11 12l9-9M17 6l3 3M14 9l2 2"/>`,
  eye: `<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>`,
  eyeOff: `<path d="M3 3l18 18"/><path d="M10.6 5.1A10 10 0 0 1 12 5c6.5 0 10 7 10 7a17 17 0 0 1-3.2 4.2M6.6 6.6A17 17 0 0 0 2 12s3.5 7 10 7a9.8 9.8 0 0 0 5.4-1.6"/>`,
  history: `<path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5M12 7v5l3 2"/>`,
  book: `<path d="M12 6.5C10 5 7 4.5 3.5 5v13c3.5-.5 6.5 0 8.5 1.5 2-1.5 5-2 8.5-1.5V5C17 4.5 14 5 12 6.5z"/><path d="M12 6.5v13"/>`,
  pencil: `<path d="M4 20l1-4L16 5l3 3L8 19z"/><path d="M14 7l3 3"/>`,
  bolt: `<path d="M13 3L5 14h6l-1 7 8-11h-6z"/>`,
  printer: `<path d="M7 9V3h10v6M7 17H4v-7h16v7h-3M7 14h10v7H7z"/>`,
  wifiOff: `<path d="M3 3l18 18M8.5 16.5a5 5 0 0 1 7 0M5 13a10 10 0 0 1 5-2.8M19 13a10 10 0 0 0-2.3-1.7M2 9.5a15 15 0 0 1 4.5-2.9M22 9.5A15 15 0 0 0 12 6c-.9 0-1.8.1-2.6.2"/><circle cx="12" cy="20" r=".8"/>`,
  sync: `<path d="M20 11a8 8 0 0 0-14.3-4.9L4 8M4 13a8 8 0 0 0 14.3 4.9L20 16"/><path d="M4 4v4h4M20 20v-4h-4"/>`,
  alert: `<path d="M12 3l10 18H2z"/><path d="M12 10v5M12 18v.5"/>`,
  hat: `<path d="M7 17h10v-4a3.6 3.6 0 0 0 .8-7.1A4.6 4.6 0 0 0 12 3.5a4.6 4.6 0 0 0-5.8 2.4A3.6 3.6 0 0 0 7 13z"/><path d="M7 20.5h10"/>`,
  sendBack: `<path d="M9 14L4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-3"/>`,
  plus: `<path d="M12 5v14M5 12h14"/>`,
  circle: `<circle cx="12" cy="12" r="8"/>`
};
function bkIcon(name, size, extraClass) {
  const s = size || 18;
  return `<svg class="bk-ico ${extraClass || ""}" width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${BK_ICON_PATHS[name] || ""}</svg>`;
}

// Emoji → icon. Everything app.js still prints with an emoji passes through here.
const BK_EMOJI = [
  ["🔊", "speaker"], ["📝", "note"], ["📜", "history"], ["🔒", "lock"], ["🔥", "timer"], ["✅", "check"],
  ["📊", "chart"], ["📚", "book"], ["🖨️", "printer"], ["🖨", "printer"], ["✏️", "pencil"], ["✏", "pencil"],
  ["👪", "bulb"], ["💡", "bulb"], ["📴", "wifiOff"], ["🔄", "sync"], ["⚠️", "alert"], ["⚠", "alert"], ["↩", "sendBack"], ["＋", "plus"]
];
const BK_EMOJI_DROP = /[\u{1F389}\u{1F950}\u{1F370}\u{1F9C1}\u{1F36A}\u{1F967}\u{2728}\u{1F31F}\u{2B50}\u{1F44F}\u{1F44D}\u{1F4AA}]️?/gu;
function bkDeEmoji(html) {
  if (!html) return html;
  let out = String(html);
  BK_EMOJI.forEach(([e, ico]) => { if (out.includes(e)) out = out.split(e).join(bkIcon(ico, 16, "inline")); });
  return out.replace(BK_EMOJI_DROP, "").replace(/️/g, "");
}
function bkStripEmojiText(text) {
  let out = String(text || "");
  BK_EMOJI.forEach(([e]) => { out = out.split(e).join(""); });
  return out.replace(BK_EMOJI_DROP, "").replace(/️/g, "").trim();
}

// ---------- Small helpers ----------
const BK_NUM_WORDS = ["Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"];
function bkNumWord(n) { return BK_NUM_WORDS[n] || String(n); }
function bkPlural(n, one, many) { return n === 1 ? one : many; }
function bkSentBackTasks(key) { return activeTasks(key).filter(t => state[key].tasks[t.id].sentBack && !state[key].tasks[t.id].done); }
// Header rule: every word capitalized ("Read Chapters 1–8"). Only raises first letters; never lowercases (keeps AAS, MCT).
function bkTitle(str) { return String(str == null ? "" : str).replace(/(^|[\s(“"\/\-])([a-z])/g, (m, a, b) => a + b.toUpperCase()); }
function bkAttr(str) { return String(str == null ? "" : str).replace(/'/g, "\\'"); }

function bkPill(kind, text, icon) {
  return `<span class="bk-pill bk-pill-${kind}">${icon ? bkIcon(icon, 15) : ""}${text}</span>`;
}

// ---------- Open a specific task from the Chef's notes ----------
function bkOpenTask(key, id) {
  openStation = key;
  if (state[key] && state[key].tasks[id]) state[key].tasks[id].open = true;
  render();
  const panel = document.getElementById("detailPanel");
  if (panel) panel.scrollIntoView({ behavior: "smooth", block: "start" });
}

// ---------- Locked quiz / test chips on each station ----------
function bkUnlockHint(t) {
  if (t.termFinal) return "end of term";
  if (t.monthlyTest) return `Week ${nextMonthlyTestWeek()}`;
  return "later";
}
function bkLockChips(locked) {
  if (!locked.length) return "";
  return `<div class="bk-st-locks">${locked.map(t =>
    `<span class="bk-lock-chip">${bkIcon("lock", 14)}${bkTitle(t.label)} <span class="bk-lock-when">· ${bkUnlockHint(t)}</span></span>`).join("")}</div>`;
}

// ---------- Kid hero ----------
function bkPlateHTML(key) {
  const kid = currentChild, th = bkTheme();
  const status = stationStatus(key);
  const sentBack = bkSentBackTasks(key).length;
  const active = unlockedActiveTasks(key);
  const doneN = active.filter(t => state[key].tasks[t.id].done).length;
  let badge = "", faded = false, label = `${DATA[key].name}: `;
  if (status === "served") { badge = `<span class="bk-plate-badge done">${bkIcon("check", 12)}</span>`; label += "served"; }
  else if (sentBack > 0 || status === "burning") { badge = `<span class="bk-plate-badge note">${bkIcon("timer", 12)}</span>`; label += "needs a touch-up"; }
  else if (doneN > 0) { label += "in progress"; }
  else { faded = true; label += status === "empty" ? "nothing yet" : "not started"; }
  return `<button class="bk-plate${faded ? " faded" : ""}" onclick="openStationFn('${key}')" aria-label="${label}" title="${label}">${bkPastry(bkPastryFor(key), kid, 36)}${badge}</button>`;
}
function bkKidHeroHTML(keys, doneCount, sentBackCount) {
  const kid = currentChild, name = CHILD_META[kid].name, th = bkTheme();
  const remaining = keys.length - doneCount;
  let line;
  if (sentBackCount > 0) line = `${bkNumWord(sentBackCount)} ${bkPlural(sentBackCount, "plate needs", "plates need")} a little touch-up. Start there, then keep baking!`;
  else if (remaining === 0 && keys.length) line = "Every plate is served this week. Beautiful work!";
  else line = `${bkNumWord(remaining)} ${bkPlural(remaining, "plate", "plates")} to go this week. You’ve got this!`;
  const testWeek = isMonthlyTestWeek() ? " · TEST WEEK" : "";
  return `<div class="bk-hero">
    <span class="bk-sprinkle s1"></span><span class="bk-sprinkle s2"></span><span class="bk-sprinkle s3"></span><span class="bk-sprinkle s4"></span><span class="bk-sprinkle s5"></span><span class="bk-sprinkle s6"></span>
    <div class="bk-emblem">${bkPastry(th.emblem, kid, 92)}</div>
    <div class="bk-hero-text">
      <div class="bk-eyebrow">${CHILD_META[kid].subtitle} · WEEK ${currentWeek()}${testWeek}</div>
      <h1>${name}’s Kitchen</h1>
      <p>${line}</p>
    </div>
    <div class="bk-plates">
      <div class="bk-plate-row">${keys.map(bkPlateHTML).join("")}</div>
      <div class="bk-plate-label">${doneCount} of ${keys.length} plates served</div>
    </div>
  </div>`;
}

// ---------- Parent header ----------
function bkNeedsGradeItems() {
  const items = [];
  Object.keys(DATA).forEach(key => DATA[key].tasks.forEach(t => {
    const s = state[key].tasks[t.id];
    if (s.needsReview && !s.reviewed) items.push({ key, t, s, why: "Needs your review" });
    else if (t.type === "reflection" && s.reviewed && !s.score) items.push({ key, t, s, why: "Approved · needs a grade" });
  }));
  return items;
}
function bkSentBackItems() {
  const items = [];
  Object.keys(DATA).forEach(key => activeTasks(key).forEach(t => {
    const s = state[key].tasks[t.id];
    if (s.sentBack && !s.done) items.push({ key, t, s });
  }));
  return items;
}
function bkAutoScoredItems(week) {
  const items = [];
  Object.keys(DATA).forEach(key => weekScopedTasks(key, week).forEach(t => {
    const s = state[key].tasks[t.id];
    if (s.done && s.score && AUTO_GRADED_TYPES.includes(t.type)) items.push({ key, t, s });
  }));
  return items;
}
function bkParentHeroHTML(keys, doneCount) {
  const kid = currentChild, name = CHILD_META[kid].name, th = bkTheme();
  const week = parentNavWeek;
  const qualifier = week === currentWeek() ? " (current)" : week < currentWeek() ? " (past)" : " (upcoming)";
  const needs = bkNeedsGradeItems().length, sent = bkSentBackItems().length;
  return `<div class="bk-phead">
    <span class="bk-phead-badge">${bkPastry(th.emblem, kid, 50)}</span>
    <div class="bk-phead-name"><div class="bk-phead-title">${name}</div><div class="bk-phead-sub">${CHILD_META[kid].subtitle.split(" · ").map(p => p === "ELA" ? p : p.charAt(0) + p.slice(1).toLowerCase()).join(" · ")}</div></div>
    <div class="bk-weeknav">
      <button onclick="navPrevWeek()" ${week <= 1 ? "disabled" : ""} aria-label="Previous week">‹</button>
      <span>Week ${week}${qualifier}</span>
      <button onclick="navNextWeek()" ${week >= maxAuthoredWeek() ? "disabled" : ""} aria-label="Next week">›</button>
    </div>
    <div class="bk-stats">
      <div class="bk-stat"><b>${doneCount} / ${keys.length}</b><span>plates served</span></div>
      <div class="bk-stat"><b>${sent}</b><span>sent back</span></div>
      <div class="bk-stat"><b>${needs}</b><span>waiting on you</span></div>
      <div class="bk-stat"><b>Week ${currentWeek()}</b><span>she’s working on</span></div>
    </div>
  </div>`;
}

function bkLegendHTML() {
  const name = CHILD_META[currentChild].name;
  return `<span class="bk-legend-title">How to read this page</span>
    <span class="bk-key bk-key-teach">${bkIcon("bulb", 15)}Teaching note · for you</span>
    <span class="bk-key bk-key-sample">${bkIcon("key", 15)}Sample answer · for you</span>
    <span class="bk-key bk-key-answer">${name}’s answer</span>
    <span class="bk-key bk-key-note">${bkIcon("note", 15)}Your note · she sees it</span>
    ${bkPill("ungraded", "Not graded yet", "circle")}
    ${bkPill("graded", "Graded", "check")}`;
}

function bkQueueHTML() {
  const name = CHILD_META[currentChild].name;
  const needs = bkNeedsGradeItems(), sent = bkSentBackItems(), auto = bkAutoScoredItems(parentNavWeek);
  const row = (it, icon, extra) => `<button class="bk-q-item" onclick="jumpToReview('${it.key}','${it.t.id}',${Number(it.t.week_number) || 1})">
      ${icon}<span class="bk-q-text"><span class="bk-q-subj">${DATA[it.key].name}${it.why ? ` · ${it.why}` : ""}</span><span class="bk-q-label">${bkTitle(it.t.label)}</span></span>${extra || ""}</button>`;
  const doneIco = `<span class="bk-q-dot solid">${bkIcon("check", 11)}</span>`;
  return `<h2>Waiting On You</h2>
    <div class="bk-q-group"><div class="bk-q-head">Needs your review or grade · ${needs.length}</div>
      ${needs.length ? needs.map(it => row(it, `<span class="bk-q-dot hollow"></span>`)).join("") : `<div class="bk-q-empty">${bkIcon("check", 16)}All caught up</div>`}</div>
    <div class="bk-q-group"><div class="bk-q-head">Sent back · waiting on ${name} · ${sent.length}</div>
      ${sent.length ? sent.map(it => row(it, `<span class="bk-q-ico note">${bkIcon("timer", 17)}</span>`)).join("") : `<div class="bk-q-empty">${bkIcon("check", 16)}Nothing sent back</div>`}</div>
    <div class="bk-q-group"><div class="bk-q-head">Scored in Week ${parentNavWeek}</div>
      ${auto.length ? auto.slice(0, 4).map(it => row(it, doneIco, `<span class="bk-q-score">${it.s.score}</span>`)).join("") : `<div class="bk-q-empty">Nothing scored yet</div>`}
      ${auto.length > 4 ? `<details class="bk-q-more"><summary>Show ${auto.length - 4} more</summary>${auto.slice(4).map(it => row(it, doneIco, `<span class="bk-q-score">${it.s.score}</span>`)).join("")}</details>` : ""}</div>`;
}

// ---------- Station card (kid view) ----------
function bkStationCardHTML(key, idx) {
  const kid = currentChild, th = bkTheme();
  const status = stationStatus(key);
  const active = activeTasks(key);
  const unlocked = active.filter(t => !isTaskLocked(t));
  const locked = active.filter(t => isTaskLocked(t));
  const doneN = unlocked.filter(t => state[key].tasks[t.id].done).length;
  const sentBack = bkSentBackTasks(key).length;
  let pill;
  if (status === "served") pill = bkPill("served", "Served", "check");
  else if (status === "burning") pill = bkPill("note", "Redo this station", "timer");
  else if (sentBack > 0) pill = bkPill("note", `${sentBack} chef’s ${bkPlural(sentBack, "note", "notes")}`, "timer");
  else if (status === "empty") pill = bkPill("neutral", "Nothing yet");
  else if (doneN > 0) pill = bkPill("neutral", "In progress");
  else pill = bkPill("neutral", "Not started");
  const segs = unlocked.map(t => {
    const s = state[key].tasks[t.id];
    const cls = s.done ? "done" : (s.sentBack ? "note" : "");
    return `<span class="bk-seg ${cls}"></span>`;
  }).join("");
  const tile = th.tiles[idx % th.tiles.length];
  return `<span class="bk-st-top"><span class="bk-st-tile" style="background:${tile}">${bkPastry(bkPastryFor(key), kid, 56)}</span>${pill}</span>
    <span class="bk-st-title">${bkTitle(DATA[key].name)}</span>
    <span class="bk-st-tag">${subjectTag(key)}</span>
    ${unlocked.length ? `<span class="bk-bar">${segs}</span><span class="bk-st-count">${doneN} of ${unlocked.length} done</span>` : `<span class="bk-st-count">Nothing loaded for this week yet</span>`}
    ${bkLockChips(locked)}`;
}

// ---------- Main render (replaces app.js render) ----------
window.render = function render() {
  regradeLegacyDictation();
  renderProjectsPanel();
  const kid = currentChild, isParent = currentView === "parent", th = bkTheme();
  const board = document.getElementById("board");
  board.dataset.child = kid;
  document.body.dataset.child = kid;
  document.body.dataset.view = currentView;
  document.getElementById("bkBrandIco").innerHTML = bkPastry("cupcake", kid, 36);

  document.getElementById("childSwitcher").innerHTML = Object.keys(CHILD_META).map(id =>
    `<button class="bk-kid kid-${id} ${kid === id ? "active" : ""}" aria-pressed="${kid === id}" onclick="switchChild('${id}')"><span class="bk-kid-badge">${bkPastry(bkTheme(id).emblem, id, 26)}</span>${CHILD_META[id].name}</button>`
  ).join("");
  const vt = document.getElementById("viewToggle");
  vt.classList.toggle("on", isParent);
  vt.setAttribute("aria-pressed", String(isParent));
  document.getElementById("toggleKnob").textContent = isParent ? "Back to student view" : "Parent view";

  if (isParent && parentNavWeek === null) parentNavWeek = currentWeek();
  if (isParent && !exportControlsReady) {
    document.getElementById("exportFromWeek").value = 1;
    document.getElementById("exportToWeek").value = currentWeek();
    exportControlsReady = true;
  }

  // Station bookkeeping (same rules as app.js)
  const keys = Object.keys(DATA);
  let doneCount = 0;
  keys.forEach(key => { if (stationDone(key)) doneCount++; });
  const sentBackAll = [];
  keys.forEach(key => bkSentBackTasks(key).forEach(t => sentBackAll.push({ key, t })));

  // Hero / parent header
  document.getElementById("bkHero").innerHTML = isParent ? bkParentHeroHTML(keys, doneCount) : bkKidHeroHTML(keys, doneCount, sentBackAll.length);

  // Parent tools, key, queue
  document.getElementById("bkTools").style.display = isParent ? "grid" : "none";
  document.getElementById("bkLegend").style.display = isParent ? "flex" : "none";
  document.getElementById("bkParentMain").style.display = isParent ? "grid" : "none";
  document.getElementById("reviewQueue").style.display = isParent ? "block" : "none";
  document.getElementById("waitingOnYou").style.display = "none";
  if (isParent) {
    document.querySelectorAll("#bkTools .bk-ico-slot").forEach(el => { el.innerHTML = bkIcon(el.dataset.ico, 18); });
    document.getElementById("gradeSheetLabel").textContent = `All of ${CHILD_META[kid].name}’s grades for the term`;
    const modeText = document.getElementById("monthlyTestModeText"), modeBtn = document.getElementById("monthlyTestBtn");
    if (settings.monthlyTestOverride === null) { modeText.textContent = `Monthly tests unlock automatically every 4th week (next: Week ${nextMonthlyTestWeek()}).`; modeBtn.textContent = "Override schedule"; }
    else if (settings.monthlyTestOverride === false) { modeText.textContent = "Monthly tests: forced open."; modeBtn.textContent = "Force locked instead"; }
    else { modeText.textContent = "Monthly tests: forced locked."; modeBtn.textContent = "Back to automatic"; }
    document.getElementById("termFinalText").textContent = settings.termFinalsUnlocked ? "Term finals are unlocked." : "Term finals are locked.";
    document.getElementById("termFinalBtn").textContent = settings.termFinalsUnlocked ? "Re-lock finals" : "Unlock finals";
    document.getElementById("bkLegend").innerHTML = bkLegendHTML();
    document.getElementById("bkQueue").innerHTML = bkQueueHTML();
  }

  // Advance-to-next-week card (parent) / all-done banner (kid)
  const advanceBanner = document.getElementById("advanceBanner");
  const served = allSubjectsServed();
  const name = CHILD_META[kid].name, nextWk = currentWeek() + 1;
  if (isParent) {
    const pending = pendingReviewCount();
    let text, btn;
    if (served && pending > 0) { text = `${name} finished every section of Week ${currentWeek()}. ${pending} written ${bkPlural(pending, "answer is", "answers are")} still waiting on your review.`; btn = `<button class="bk-btn plain" disabled>Advance to Week ${nextWk}</button>`; }
    else if (served) { text = `${name} finished every section of Week ${currentWeek()}.`; btn = `<button class="bk-btn primary" onclick="advanceWeek()">Advance to Week ${nextWk}</button>`; }
    else { text = `Week ${currentWeek()} isn’t fully finished yet${sentBackAll.length ? `: ${sentBackAll.length} ${bkPlural(sentBackAll.length, "item is", "items are")} still sent back` : ""}.`; btn = `<button class="bk-btn plain" onclick="advanceAnyway()">Advance to Week ${nextWk} anyway</button>`; }
    advanceBanner.className = "bk-tool";
    advanceBanner.innerHTML = `<div class="bk-tool-title">${bkIcon("arrow", 18)}Next Week</div><div class="bk-tool-text">${text}</div>${btn}`;
  } else {
    advanceBanner.className = "";
    advanceBanner.innerHTML = "";
  }

  // Kid: all-done banner + Chef's notes + stations + open station
  const notes = document.getElementById("sentBackBanner");
  const stationsWrap = document.getElementById("bkStationsWrap");
  const grid = document.getElementById("stationsGrid");
  const panel = document.getElementById("detailPanel");
  document.getElementById("weekReportPanel").style.display = isParent ? "block" : "none";
  if (isParent) {
    notes.innerHTML = ""; stationsWrap.style.display = "none"; panel.className = "detail"; panel.innerHTML = "";
    renderWeekReportPanel();
  } else {
    stationsWrap.style.display = "block";
    let notesHTML = "";
    if (served) {
      notesHTML += `<section class="bk-alldone">${bkPastry(th.emblem, kid, 48)}<div><h2>All Done With Week ${currentWeek()}!</h2><p>Nice work! Waiting for Mom to check everything, then you’ll move on to Week ${nextWk}.</p></div></section>`;
    }
    if (sentBackAll.length) {
      notesHTML += `<section class="bk-notes">
        <div class="bk-notes-head"><span class="bk-notes-ico">${bkIcon("timer", 22)}</span><div><h2>Chef’s Notes</h2><p>${sentBackAll.length === 1 ? "This plate is" : `These ${bkNumWord(sentBackAll.length).toLowerCase()} plates are`} back in the oven for a quick touch-up.</p></div></div>
        <div class="bk-notes-grid">${sentBackAll.map(({ key, t }) => `<button class="bk-note-row" onclick="bkOpenTask('${key}','${t.id}')">
            <span class="bk-note-pastry">${bkPastry(bkPastryFor(key), kid, 34)}</span>
            <span class="bk-note-text"><span class="bk-note-subj">${DATA[key].name}</span><span class="bk-note-task">${bkTitle(t.label)}</span></span>
            ${bkIcon("arrow", 20)}</button>`).join("")}</div>
      </section>`;
    }
    notes.innerHTML = notesHTML;
    document.getElementById("bkWeekMeta").textContent = `Week ${currentWeek()}${isMonthlyTestWeek() ? " · test week" : ""}`;

    grid.innerHTML = "";
    grid.className = `stations bk-stations count-${keys.length}`;
    keys.forEach((key, idx) => {
      const status = stationStatus(key);
      const card = document.createElement("button");
      card.type = "button";
      card.className = "station bk-station" + (status === "served" ? " done" : "") + (status === "burning" ? " burning" : "") + (openStation === key ? " active" : "");
      if (status === "served") {
        const stampId = `${currentChild}:${key}:${currentWeek()}`;
        if (!stampedStations.has(stampId)) {
          stampedStations.add(stampId);
          if (stampSeeded[currentChild]) card.classList.add("stamp-new");
        }
      }
      card.onclick = () => openStationFn(key);
      card.innerHTML = bkStationCardHTML(key, idx);
      grid.appendChild(card);
    });

    if (openStation && DATA[openStation]) {
      const d = DATA[openStation];
      const status = stationStatus(openStation);
      const active = activeTasks(openStation);
      const rows = active.length === 0
        ? `<div class="empty-note">Nothing loaded for Week ${currentWeek()} yet — check back once new content is added.</div>`
        : active.map(t => `<div class="task-row">${taskHeadHTML(openStation, t)}${isTaskLocked(t) ? "" : taskBodyHTML(openStation, t)}</div>`).join("");
      const note = status === "burning" ? `<div class="burn-note">This station scored below 70%. Redoing it resets the section and keeps the first scores for Mom to see.</div>` : "";
      panel.className = "detail open bk-detail";
      panel.innerHTML = `<div class="detail-head">
          <div class="bk-detail-title"><span class="bk-st-tile small">${bkPastry(bkPastryFor(openStation), kid, 40)}</span><div><div class="detail-tag">${subjectTag(openStation)}</div><div class="detail-title">${bkTitle(d.name)}</div>${note}</div></div>
          <div class="bk-detail-actions">
            ${status === "burning" ? `<button class="bk-btn primary" onclick="redoStation('${openStation}')">Redo this station</button>` : ""}
            <button class="detail-close bk-btn plain" onclick="openStationFn('${openStation}')">${bkIcon("close", 16)}Close</button>
          </div>
        </div>${rows}`;
    } else { panel.className = "detail"; panel.innerHTML = ""; }
  }
  stampSeeded[currentChild] = true;

  // Parent: review bank + redo log (same data as before)
  if (isParent) {
    const pool = loadPool();
    const activeW = pool.filter(p => p.status === "active");
    const mastered = pool.filter(p => p.status === "mastered");
    document.getElementById("reviewBank").innerHTML = pool.length === 0
      ? `<div class="empty-note">No missed words logged yet.</div>`
      : `<div class="bk-bank-group"><div class="bk-bank-label">Still practicing · ${activeW.length}</div><div class="bk-bank-words">${activeW.map(p => `<span class="bk-word" title="Missed ${p.timesMissed}× · from ${escHtml(p.lastSeen)}">${escHtml(p.word)}${p.timesMissed > 1 ? ` <b>×${p.timesMissed}</b>` : ""}</span>`).join("") || "none"}</div></div>
         <div class="bk-bank-group"><div class="bk-bank-label">Mastered · ${mastered.length}</div><div class="bk-bank-words">${mastered.map(p => `<span class="bk-word mastered">${escHtml(p.word)}</span>`).join("") || "none yet"}</div></div>`;
    const burnLog = burnLogCache[currentChild] || [];
    document.getElementById("burnLogList").innerHTML = burnLog.length === 0
      ? `<div class="empty-note">No sections have needed a redo yet.</div>`
      : burnLog.map(rec => `<div class="review-item"><strong>${rec.station} — ${rec.tag}</strong><div class="meta">${rec.date} · ${rec.reason}</div><div class="submitted-text">${rec.items.join("\n")}</div></div>`).join("");
  }
};

// ---------- Task header (kid + parent station list) ----------
window.taskHeadHTML = function taskHeadHTML(key, t) {
  const s = state[key].tasks[t.id];
  if (isTaskLocked(t)) {
    const msg = t.termFinal ? "Unlocks at end of term" : `Unlocks in Week ${nextMonthlyTestWeek()}`;
    return `<div class="task-head locked-task"><span class="bk-th-ico lock">${bkIcon("lock", 16)}</span><span class="label">${bkTitle(t.label)}</span><span class="status-text">${msg}</span></div>`;
  }
  let cls = "", pill;
  const isParent = currentView === "parent";
  if (s.sentBack && !s.done) { cls = "review"; pill = bkPill("note", "Chef’s note", "timer"); }
  else if (s.done && s.needsReview && !s.reviewed) { cls = "review"; pill = bkPill("neutral", isParent ? "Awaiting your review" : "Waiting for Mom"); }
  else if (s.done) { cls = "done"; pill = bkPill("served", s.score ? `Scored ${s.score}` : (s.reviewed ? "Approved" : "Done"), "check"); }
  else pill = bkPill("neutral", "Not started");
  const ico = s.done ? `<span class="bk-th-ico done">${bkIcon("check", 14)}</span>` : (s.sentBack ? `<span class="bk-th-ico note">${bkIcon("timer", 14)}</span>` : `<span class="bk-th-ico"></span>`);
  return `<div class="task-head ${cls}" role="button" tabindex="0" onclick="toggleTask('${key}','${t.id}')" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();toggleTask('${key}','${t.id}')}">
    ${ico}<span class="label">${bkTitle(t.label)}</span><span class="status-text">${pill}</span></div>`;
};

// Everything app.js draws inside a task keeps its behavior; only emoji become icons.
const bkOrigTaskBodyHTML = window.taskBodyHTML;
window.taskBodyHTML = function taskBodyHTML(key, t) { return bkDeEmoji(bkOrigTaskBodyHTML(key, t)); };
const bkOrigReadAloudButton = window.readAloudButton;
window.readAloudButton = function readAloudButton(elId, label) { return bkDeEmoji(bkOrigReadAloudButton(elId, label)); };
const bkOrigCaseFiles = window.caseFilesVocabReportHTML;
window.caseFilesVocabReportHTML = function caseFilesVocabReportHTML() { return bkDeEmoji(bkOrigCaseFiles()); };
const bkOrigProjectsPanel = window.renderProjectsPanel;
window.renderProjectsPanel = function renderProjectsPanel() {
  bkOrigProjectsPanel();
  const p = document.getElementById("projectsPanel");
  if (p) p.innerHTML = bkDeEmoji(p.innerHTML);
};
const bkOrigGradeSheetHTML = window.gradeSheetHTML;
window.gradeSheetHTML = function gradeSheetHTML() { return bkDeEmoji(bkOrigGradeSheetHTML()); };
const bkOrigOffline = window.updateOfflineBanner;
window.updateOfflineBanner = function updateOfflineBanner() {
  bkOrigOffline();
  const el = document.getElementById("syncBanner");
  if (el && el.style.display !== "none") el.textContent = bkStripEmojiText(el.textContent);
};
const bkOrigSyncError = window.showSyncError;
window.showSyncError = function showSyncError(msg) {
  bkOrigSyncError(msg);
  const el = document.getElementById("syncBanner");
  if (el) el.textContent = bkStripEmojiText(el.textContent);
};

// ---------- Earlier attempts (kid + parent) ----------
window.attemptHistoryHTML = function attemptHistoryHTML(t, s, opts) {
  const hist = s.history || [];
  if (!hist.length) return "";
  const forStudent = opts && opts.forStudent;
  const rows = hist.map((h, i) => {
    const when = new Date(h.at).toLocaleDateString(undefined, { month: "short", day: "numeric" });
    const who = h.by === "parent" ? "sent back" : "redone by her";
    return `<div class="attempt bk-attempt">
      <div class="attempt-head"><b>Attempt ${i + 1}</b> <span class="attempt-dim">· ${when}${forStudent ? "" : " · " + who}${h.score && !forStudent ? " · scored " + escHtml(h.score) : ""}</span></div>
      ${attemptSummaryHTML(t, h)}
      ${h.feedback ? `<div class="bk-bubble small"><div class="bk-block-label">${bkIcon("note", 14)}${forStudent ? "Feedback you got" : "Your note"}</div>${escHtml(h.feedback)}</div>` : ""}
    </div>`;
  }).join("");
  const label = forStudent ? `My Earlier Drafts (${hist.length})` : `Earlier Attempts (${hist.length})`;
  return `<details class="attempt-history bk-history"${opts && opts.open ? " open" : ""}><summary>${bkIcon("history", 16)}${label}</summary>${rows}</details>`;
};

// ---------- Parent: teaching note vs. sample answer ----------
// The Sheet currently stores both in one "sampleAnswer" field. Until a separate
// `teachingNote` field exists (Stage 3), split on the phrases the lesson plans use
// when they switch from a model answer to guidance for the parent.
const BK_GUIDANCE_MARKERS = [/Coaching note for you[^:]*:?/i, /A strong answer/i, /This is a personal-topic prompt/i, /Ask a concrete, curious question/i];
function bkSplitSample(t) {
  let sample = t.sampleAnswer || "", note = t.teachingNote || "", illustrative = false;
  if (!note && sample) {
    let cut = -1, marker = null;
    BK_GUIDANCE_MARKERS.forEach(re => { const m = sample.match(re); if (m && (cut === -1 || m.index < cut)) { cut = m.index; marker = m[0]; } });
    if (cut === 0) { note = sample; sample = ""; }
    else if (cut > 0) { note = sample.slice(cut); sample = sample.slice(0, cut); }
    if (/^Coaching note for you/i.test(note)) note = note.replace(/^Coaching note for you[^:]*:\s*/i, "");
  }
  const ill = sample.match(/^\s*An illustrative example only\s*(\([^)]*\))?\s*:?\s*/i);
  if (ill) { illustrative = true; sample = sample.slice(ill[0].length); }
  return { sample: sample.trim(), note: note.trim(), illustrative };
}
function bkTeachBlock(html) {
  return `<div class="bk-block bk-teach"><div class="bk-block-top"><span class="bk-block-label">${bkIcon("bulb", 16)}Teaching note</span><span class="bk-only">${bkIcon("eyeOff", 14)}Only you see this</span></div><div class="bk-block-body">${html}</div></div>`;
}
function bkSampleBlock(html, illustrative) {
  return `<div class="bk-block bk-sample"><div class="bk-block-top"><span class="bk-block-label">${bkIcon("key", 16)}Sample answer${illustrative ? " · one example, not a target" : ""}</span><span class="bk-only">${bkIcon("eyeOff", 14)}Only you see this</span></div><div class="bk-block-body">${html}</div></div>`;
}
function bkQuestionBlock(label, html) {
  return `<div class="bk-question"><div class="bk-q-label2">${label}</div><div class="bk-question-body">${html}</div></div>`;
}
function bkAnswerBlock(html, meta) {
  const kid = currentChild, name = CHILD_META[kid].name;
  return `<div class="bk-block bk-answer"><div class="bk-block-top"><span class="bk-block-label"><span class="bk-initial">${name[0]}</span>${name}’s answer</span>${meta ? `<span class="bk-answer-meta">${meta}</span>` : ""}</div><div class="bk-block-body bk-answer-text">${html}</div></div>`;
}
function bkNoteBubble(text) {
  const name = CHILD_META[currentChild].name;
  return `<div class="bk-bubble"><div class="bk-block-top"><span class="bk-block-label">${bkIcon("note", 16)}Your note to ${name}</span><span class="bk-only">${bkIcon("eye", 14)}She sees this</span></div><div class="bk-block-body">${escHtml(text)}</div></div>`;
}
function bkGradeEntryHTML(key, t, s) {
  const pct = gradePctOf(s);
  return `<div class="bk-grade">
    <div class="bk-grade-field"><label for="grade-${key}-${t.id}">Grade (0–100)</label>
      <input type="number" min="0" max="100" step="1" id="grade-${key}-${t.id}" value="${pct == null ? "" : pct}" placeholder="0–100"></div>
    <span class="bk-grade-letter">${pct == null ? "Not graded yet" : "= " + letterFor(pct)}</span>
    <button class="bk-btn primary" onclick="saveReflectionGrade('${key}','${t.id}')">${pct == null ? "Save grade" : "Update grade"}</button>
    <span class="bk-grade-hint">Counts toward her lessons grade. You can change it after she revises.</span>
  </div>`;
}
function bkStatusChips(t, s, awaiting) {
  const name = CHILD_META[currentChild].name;
  const chips = [];
  if (s.sentBack && !s.done) chips.push(bkPill("note", `Sent back · waiting on ${name}`, "timer"));
  else if (awaiting) chips.push(bkPill("review", "Waiting for your review", "eye"));
  else if (s.reviewed) chips.push(bkPill("served", "Approved", "check"));
  else if (s.done) chips.push(bkPill("neutral", "Complete"));
  else chips.push(bkPill("neutral", "Not started"));
  if (t.type === "reflection" && (s.done || s.sentBack || (s.answers && s.answers.text))) {
    const pct = gradePctOf(s);
    chips.push(pct == null ? bkPill("ungraded", "Not graded yet", "circle") : bkPill("graded", `Graded · ${letterFor(pct)} (${s.score})`, "check"));
  } else if (s.done && s.score && AUTO_GRADED_TYPES.includes(t.type)) {
    chips.push(bkPill("graded", `${t.type === "fluency-read" ? "Marked by you" : "Auto-scored"} · ${s.score}`, t.type === "fluency-read" ? "check" : "bolt"));
  }
  return chips.join("");
}

window.renderPastTaskReport = function renderPastTaskReport(key, t, s) {
  const name = CHILD_META[currentChild].name;
  const awaiting = t.type === "reflection" && s.needsReview && !s.reviewed;
  let body = "";
  if (t.type === "reflection") {
    body += bkQuestionBlock("Question · what she saw", `<p>${t.prompt}</p>`);
    const split = bkSplitSample(t);
    if (split.note || split.sample) body += `<div class="bk-pair${split.note && split.sample ? "" : " single"}">${split.note ? bkTeachBlock(split.note) : ""}${split.sample ? bkSampleBlock(split.sample, split.illustrative) : ""}</div>`;
    const hist = s.history || [];
    if (hist.length) body += attemptHistoryHTML(t, s, { open: false });
    if (s.answers && s.answers.text) {
      const attemptNo = hist.length + 1;
      const meta = `Attempt ${attemptNo}${hist.length ? " · latest" : ""}${s.sentBack && !s.done ? " · she hasn’t resubmitted yet" : ""}`;
      body += bkAnswerBlock(escHtml(s.answers.text), meta);
    } else if (!s.done) {
      body += `<div class="bk-empty-answer">${name} hasn’t answered yet.</div>`;
    }
    const lastFb = hist.length ? (hist[hist.length - 1].feedback || "Please take another look and resubmit.") : null;
    const showComment = s.parentComment && !(lastFb !== null && lastFb === s.parentComment && !s.reviewed);
    if (showComment) body += bkNoteBubble(s.parentComment);
    if (awaiting) {
      body += `<div class="bk-note-entry"><label for="comment-${key}-${t.id}">${bkIcon("note", 16)}Your note to ${name} <span>· she sees this</span></label>
        <textarea id="comment-${key}-${t.id}" placeholder="Optional. If you send it back, she’ll read this first."></textarea></div>`;
    }
    if (s.answers && s.answers.text) body += bkGradeEntryHTML(key, t, s);
    if (awaiting) {
      body += `<div class="bk-actions"><button class="bk-btn soft" onclick="sendBackReflection('${key}','${t.id}')">${bkIcon("sendBack", 16)}Send back</button><button class="bk-btn primary" onclick="approveReflection('${key}','${t.id}')">${bkIcon("check", 16)}Approve</button></div>`;
    }
  } else if (t.type === "read") {
    if (t.parentNotes) body += bkTeachBlock(t.parentNotes);
    body += `<details class="bk-lesson"><summary>${bkIcon("book", 16)}Lesson She Read</summary><div class="bk-lesson-body">${bkDeEmoji(t.content || "")}</div></details>`;
  } else {
    if (s.done && AUTO_GRADED_TYPES.includes(t.type)) {
      const work = taskBodyHTML(key, t).replace('class="task-body ', 'class="task-body open ');
      body += `<div class="bk-block bk-answer bk-work"><div class="bk-block-top"><span class="bk-block-label"><span class="bk-initial">${name[0]}</span>${name}’s answers</span></div><div class="bk-block-body">${work}</div></div>`;
    } else {
      body += bkQuestionBlock("What she saw", bkDeEmoji(renderTaskContent(t)));
    }
    if (isCaseFilesTask(t) && currentChild === "kenley") body += caseFilesVocabReportHTML();
    const histHtml = attemptHistoryHTML(t, s, { open: false });
    if (histHtml) body += histHtml;
  }
  if (s.done && !awaiting) {
    body += `<details class="sendback-box bk-sendback"><summary>${bkIcon("sendBack", 16)}Send Back Just This Section</summary>
      <label class="bk-sendback-label" for="sendback-${key}-${t.id}">Note for ${name} (she sees this)</label>
      <textarea id="sendback-${key}-${t.id}" placeholder="What should she fix?"></textarea>
      <div class="bk-actions"><button class="bk-btn soft" onclick="sendBackSection('${key}','${t.id}')">Send back &amp; reset this section</button></div>
    </details>`;
  }
  return `<article class="bk-item review-item${awaiting ? " needs-attention" : ""}" id="review-${key}-${t.id}">
    <header class="bk-item-head"><div><div class="bk-item-subj">${DATA[key].name}</div><h3>${bkTitle(t.label)}</h3></div><div class="bk-chips">${bkStatusChips(t, s, awaiting)}</div></header>
    ${body}
  </article>`;
};

window.renderUpcomingTaskPreview = function renderUpcomingTaskPreview(t) {
  return `<article class="bk-item review-item"><header class="bk-item-head"><div><div class="bk-item-subj">Coming up</div><h3>${bkTitle(t.label)}</h3></div><div class="bk-chips">${bkPill("neutral", "Preview")}</div></header>
    <details class="bk-lesson"><summary>${bkIcon("book", 16)}What She’ll See</summary><div class="bk-lesson-body">${bkDeEmoji(renderTaskContent(t))}</div></details></article>`;
};

window.renderWeekReportPanel = function renderWeekReportPanel() {
  const panel = document.getElementById("weekReportPanel");
  const week = parentNavWeek;
  const isPast = week <= currentWeek();
  const kid = currentChild;
  const sections = [];
  SUBJECT_ORDER.forEach(key => {
    const tasks = weekScopedTasks(key, week);
    if (tasks.length === 0) return;
    sections.push(`<div class="bk-report-subject">${bkPastry(bkPastryFor(key), kid, 34)}<span>${bkTitle(DATA[key].name)}</span><span class="bk-report-tag">${escHtml(DATA[key].tagsByWeek[week] || DATA[key].tag || "")}</span></div>`);
    tasks.forEach(t => {
      const s = state[key].tasks[t.id];
      sections.push(isPast ? renderPastTaskReport(key, t, s) : renderUpcomingTaskPreview(t));
    });
  });
  const titleSuffix = week < currentWeek() ? "Completed Record" : week === currentWeek() ? "In Progress" : "Preview";
  panel.innerHTML = `<div class="bk-report-title">Week ${week} · ${titleSuffix}</div>
    ${sections.length ? sections.join("") : `<div class="empty-note">Nothing planned yet for Week ${week}.</div>`}`;
};
