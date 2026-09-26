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
      ${typeof bkRankHTML === "function" ? bkRankHTML() : ""}
    </div>
    <div class="bk-plates">
      <div class="bk-plate-row">${keys.map(bkPlateHTML).join("")}</div>
      <div class="bk-plate-label">${doneCount} of ${keys.length} plates served</div>
      ${typeof bkTreatTeaserHTML === "function" ? bkTreatTeaserHTML() : ""}
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
      <div class="bk-stat"><b>${bkRankInfo().name}</b><span>chef rank · ${bkRankInfo().plates} ${bkPlural(bkRankInfo().plates, "plate", "plates")}</span></div>
      <div class="bk-stat"><b>${bkTreatsCollected()} / ${bkTreatsTotal()}</b><span>passport treats</span></div>
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
  if (status === "served" && typeof bkPerfected === "function" && bkPerfected(key)) pill = bkPill("perfected", "Perfected", "star");
  else if (status === "served") pill = bkPill("served", "Served", "check");
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
      card.dataset.key = key;
      if (typeof bkPerfected === "function" && bkPerfected(key)) card.classList.add("perfected");
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

// =====================================================================
// STAGE 2 — chef ranks, Perfected badge, served celebration, pastry passport.
// Everything here is worked out from data the Sheet already has.
// =====================================================================

let bkPassport = null;          // null = kitchen, "case" = passport case, number = that week's treat card
const bkWeekSeen = {};          // "kenley:1" -> was the week complete at the last render?
let bkPendingCelebration = null;

const BK_RANKS = [[0, "Commis"], [10, "Line Cook"], [25, "Pastry Cook"], [45, "Sous Chef"], [70, "Head Pastry Chef"]];

function bkIsTestTask(t) { return !!(t.monthlyTest || t.termFinal || t.dynamic); }
// true / false for a week's plate in one subject, or null if that subject had nothing that week
function bkPlateServed(key, w) {
  if (w === currentWeek()) return unlockedActiveTasks(key).length ? stationDone(key) : null;
  const tasks = DATA[key].tasks.filter(t => t.week_number === w && !bkIsTestTask(t));
  if (!tasks.length) return null;
  return tasks.every(t => state[key].tasks[t.id].done);
}
function bkWeekComplete(w) {
  const res = Object.keys(DATA).map(k => bkPlateServed(k, w)).filter(v => v !== null);
  return res.length > 0 && res.every(Boolean);
}
function bkPlatesServedTotal() {
  let n = 0;
  for (let w = 1; w <= currentWeek(); w++) Object.keys(DATA).forEach(k => { if (bkPlateServed(k, w) === true) n++; });
  return n;
}
function bkRankInfo() {
  const plates = bkPlatesServedTotal();
  let idx = 0;
  BK_RANKS.forEach((r, i) => { if (plates >= r[0]) idx = i; });
  const next = BK_RANKS[idx + 1];
  const from = BK_RANKS[idx][0];
  return { plates, name: BK_RANKS[idx][1], next: next ? next[1] : null, nextAt: next ? next[0] : null,
    pct: next ? Math.max(4, Math.round(((plates - from) / (next[0] - from)) * 100)) : 100 };
}
function bkTreatForWeek(w) { return (typeof BK_TREATS !== "undefined" ? BK_TREATS : []).find(t => t.week === w) || null; }
function bkTreatUnlocked(w) { return w <= currentWeek() && !!bkTreatForWeek(w) && bkWeekComplete(w); }
function bkTreatsCollected() { return (typeof BK_TREATS !== "undefined" ? BK_TREATS : []).filter(t => bkTreatUnlocked(t.week)).length; }
function bkTreatsTotal() { return (typeof BK_TREATS !== "undefined" ? BK_TREATS : []).length; }

// A station is Perfected when it's served and at least one of its tasks was sent back and then fixed.
function bkPerfected(key) {
  if (stationStatus(key) !== "served") return false;
  return unlockedActiveTasks(key).some(t => (state[key].tasks[t.id].history || []).some(h => h.by === "parent"));
}

function bkRankHTML() {
  const r = bkRankInfo();
  const sub = r.next ? `${r.plates} of ${r.nextAt} plates to ${r.next}` : `${r.plates} ${bkPlural(r.plates, "plate", "plates")} served. Top rank!`;
  return `<div class="bk-rank"><span class="bk-rank-chip"><span class="bk-rank-hat">${bkIcon("hat", 18)}</span>${r.name}</span>
    <span class="bk-rank-prog"><span class="bk-rank-bar"><span style="width:${r.pct}%"></span></span><span class="bk-rank-sub">${sub}</span></span></div>`;
}
function bkTreatTeaserHTML() {
  const w = currentWeek(), treat = bkTreatForWeek(w);
  if (!treat) return "";
  if (bkTreatUnlocked(w)) return `<button class="bk-teaser unlocked" onclick="bkOpenPassport(${w})">${bkIcon("check", 16)}Treat unlocked: ${treat.name}!</button>`;
  return `<button class="bk-teaser" onclick="bkOpenPassport('case')">${bkPassportIcon(16)}Serve all 5 to unlock a treat from ${treat.city}</button>`;
}
function bkPassportIcon(size) {
  return `<svg class="bk-ico" width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="5" y="3" width="14" height="18" rx="2"/><circle cx="12" cy="10" r="3.5"/><path d="M8.5 10h7M9 17h6"/></svg>`;
}
function bkClocheSVG(size, fill, stroke, mark) {
  return `<svg width="${size}" height="${size}" viewBox="0 0 64 64" aria-hidden="true"><path d="M11 45C11 31 20.5 21 32 21S53 31 53 45Z" fill="${fill}" stroke="${stroke}" stroke-width="2.5" stroke-linejoin="round"/><circle cx="32" cy="17.5" r="3.5" fill="${fill}" stroke="${stroke}" stroke-width="2.5"/><path d="M6 46.5H58" stroke="${stroke}" stroke-width="3" stroke-linecap="round"/>${mark ? `<text x="33" y="41" text-anchor="middle" font-family="Fredoka, sans-serif" font-size="15" font-weight="600" fill="${stroke}">?</text>` : ""}</svg>`;
}
function bkTreatPlateSVG(size) {
  return `<svg width="${size}" height="${size}" viewBox="0 0 64 64" aria-hidden="true"><ellipse cx="32" cy="40" rx="26" ry="10" fill="#FFFFFF" stroke="#E0B84A" stroke-width="2.5"/><path d="M14 38C14 26 22 20 32 20S50 26 50 38Z" fill="#F6C98A" stroke="#C98B45" stroke-width="2.5" stroke-linejoin="round"/><path d="M24 26l-2 10M32 23v13M40 26l2 10" stroke="#C98B45" stroke-width="2" stroke-linecap="round"/><path d="M50 12l2 4 4 .6-3 2.8.8 4.2-3.8-2-3.8 2 .8-4.2-3-2.8 4-.6z" fill="#F4D77A" stroke="#C9A43A" stroke-width="1.2" stroke-linejoin="round"/></svg>`;
}
const BK_CREMA_ART = `<svg width="100%" height="100%" viewBox="0 0 548 368" preserveAspectRatio="xMidYMid slice" aria-label="Illustration of crema catalana in a terracotta dish" role="img"><rect width="548" height="368" fill="#F4E9DA"/><path d="M0 60H548M0 150H548M0 240H548M0 330H548M70 0V368M190 0V368M310 0V368M430 0V368" stroke="#EBDCC7" stroke-width="18"/><ellipse cx="274" cy="222" rx="206" ry="112" fill="#000" opacity="0.08"/><ellipse cx="274" cy="206" rx="192" ry="104" fill="#A9542E"/><ellipse cx="274" cy="186" rx="192" ry="104" fill="#C8693D"/><ellipse cx="274" cy="182" rx="170" ry="90" fill="#DA8350"/><ellipse cx="274" cy="184" rx="156" ry="80" fill="#E0A04A"/><ellipse cx="226" cy="170" rx="52" ry="22" fill="#EFBE62"/><ellipse cx="324" cy="198" rx="46" ry="18" fill="#EFBE62"/><ellipse cx="300" cy="150" rx="30" ry="11" fill="#C07A2E"/><ellipse cx="200" cy="208" rx="26" ry="9" fill="#C07A2E"/><path d="M190 160l30 12 18-8 34 20M280 196l26-10 24 14 30-6M232 214l22-14M312 150l-10 18" stroke="#8F5420" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" fill="none"/><rect x="48" y="300" width="150" height="22" rx="11" fill="#9A5B34" transform="rotate(-14 123 311)"/><path d="M392 318c26-22 64-28 96-14c-8 16-44 30-78 30c-8 0-14-6-18-16z" fill="#F4D35E" stroke="#D9B53C" stroke-width="3" stroke-linejoin="round"/></svg>`;

// ---------- Navigation ----------
function bkOpenPassport(where) {
  bkPassport = where == null ? "case" : where;
  render();
  window.scrollTo({ top: 0, behavior: "smooth" });
}
function bkClosePassport() { bkPassport = null; render(); window.scrollTo({ top: 0 }); }
const bkOrigSwitchChild = window.switchChild;
window.switchChild = function switchChild(id) { bkPassport = null; return bkOrigSwitchChild(id); };
const bkOrigToggleView = window.toggleView;
window.toggleView = function toggleView() { bkPassport = null; return bkOrigToggleView(); };

// ---------- Passport case ----------
function bkPassportCaseHTML() {
  const kid = currentChild, name = CHILD_META[kid].name, th = bkTheme();
  const treats = BK_TREATS, cw = currentWeek();
  const cur = bkTreatForWeek(cw);
  const keys = Object.keys(DATA);
  const doneNow = keys.filter(k => stationDone(k)).length;
  let feature = "";
  if (cur && bkTreatUnlocked(cw)) {
    feature = `<div class="bk-pp-feature"><div class="bk-pp-feature-art">${bkTreatPlateSVG(120)}</div><div class="bk-pp-feature-text">
      <div class="bk-eyebrow">THIS WEEK’S TREAT · WEEK ${cw} · ${cur.city.toUpperCase()}, ${cur.country.toUpperCase()}</div>
      <h2>You Unlocked ${cur.name}!</h2><p>It’s in your case now. Open the card for its story and a recipe to bake together.</p>
      <button class="bk-btn primary" onclick="bkOpenPassport(${cw})">Open the treat card</button></div></div>`;
  } else if (cur) {
    feature = `<div class="bk-pp-feature"><div class="bk-pp-feature-art">${bkClocheSVG(124, "#FFFFFF", th.bar, true)}</div><div class="bk-pp-feature-text">
      <div class="bk-eyebrow">THIS WEEK’S TREAT · WEEK ${cw} · ${cur.city.toUpperCase()}, ${cur.country.toUpperCase()}</div>
      <h2>Something Delicious Is Hiding Under The Lid</h2><p>Serve all ${keys.length} plates this week to lift the lid and add it to your case.</p>
      <div class="bk-pp-progress"><span class="bk-bar">${keys.map((k, i) => `<span class="bk-seg ${i < doneNow ? "done" : ""}"></span>`).join("")}</span><b>${doneNow} of ${keys.length} plates</b></div></div></div>`;
  }
  const slot = t => {
    const unlocked = bkTreatUnlocked(t.week), isNow = t.week === cw;
    if (unlocked) {
      return `<button class="bk-slot unlocked" onclick="bkOpenPassport(${t.week})" aria-label="Open ${t.name}">
        <span class="bk-slot-plate"><img src="images/treats/${t.slug}.jpg" alt="" onerror="this.remove()">${bkTreatPlateSVG(62)}</span>
        <span class="bk-slot-name">${t.name}</span><span class="bk-slot-where">Week ${t.week} · ${t.city}</span></button>`;
    }
    return `<div class="bk-slot${isNow ? " now" : ""}">
      <span class="bk-slot-plate">${bkClocheSVG(58, isNow ? "#FFFFFF" : "#F1EEF3", isNow ? th.bar : "#A79DB3", isNow)}</span>
      ${isNow ? `<span class="bk-slot-badge">Baking Now</span>` : ""}
      <span class="bk-slot-name">${t.city}</span><span class="bk-slot-where">Week ${t.week} · ${t.country}</span></div>`;
  };
  const shelves = [];
  for (let i = 0; i < treats.length; i += 4) shelves.push(`<div class="bk-shelf">${treats.slice(i, i + 4).map(slot).join("")}</div>`);
  return `<div class="bk-pp">
    <div class="bk-pp-head">
      <div><button class="bk-back" onclick="bkClosePassport()">${bkIcon("back", 18)}Back To ${name}’s Kitchen</button>
        <h1>${name}’s Pastry Passport</h1><p>Finish a week, unlock a treat from somewhere new in the world.</p></div>
      <div class="bk-pp-count">${bkPassportIcon(28)}<div><b>${bkTreatsCollected()} of ${bkTreatsTotal()}</b><span>treats collected this term</span></div></div>
    </div>
    ${feature}
    <section class="bk-case"><div class="bk-case-head"><h2>My Pastry Case</h2><span>A new stop every week</span></div>${shelves.join("")}</section>
  </div>`;
}

// ---------- Treat card ----------
const BK_WHO = { kid: ["Kid Job", "kid"], grown: ["Grown-Up Job", "grown"], together: ["Together", "together"] };
function bkTreatCardHTML(w) {
  const t = bkTreatForWeek(w);
  if (!t || !bkTreatUnlocked(w)) return bkPassportCaseHTML();
  const r = t.recipe;
  const art = t.slug === "crema-catalana" ? BK_CREMA_ART : `<div class="bk-photo-ph">${bkTreatPlateSVG(140)}<span>${t.name}</span></div>`;
  return `<div class="bk-treat">
    <div class="bk-treat-head">
      <div class="bk-stamp"><span>WEEK ${t.week}</span><b>${t.city}</b><span>${t.country.toUpperCase()}</span></div>
      <div class="bk-treat-title"><button class="bk-back" onclick="bkOpenPassport('case')">${bkIcon("back", 18)}Back To My Passport</button>
        <h1>${t.name}</h1><p><i>${t.altName}</i> · ${t.tagline}</p></div>
      <span class="bk-pill bk-pill-served big">${bkIcon("check", 16)}In Your Case</span>
    </div>
    <div class="bk-treat-grid">
      <div class="bk-treat-left">
        <div class="bk-photo"><img src="images/treats/${t.slug}.jpg" alt="${t.name}" onerror="this.remove()">${art}</div>
        ${t.photoCredit ? `<div class="bk-photo-credit">Photo: ${t.photoCredit}</div>` : ""}
        <div class="bk-where"><span class="bk-where-ico">${bkIcon("pin", 26)}</span><div><div class="bk-where-label">Where It’s From</div><div class="bk-where-name">${t.where}</div><div class="bk-where-sub">${t.whereSub}</div></div></div>
      </div>
      <div class="bk-story"><h2>The Story</h2><p>${t.story}</p>
        <div class="bk-facts">${t.facts.map((f, i) => `<div class="bk-fact"><span class="bk-fact-n n${i}">${i + 1}</span><span>${f}</span></div>`).join("")}</div></div>
    </div>
    <section class="bk-recipe">
      <div class="bk-recipe-head"><h2>Let’s Make It Together</h2><div class="bk-recipe-meta"><span>${r.serves}</span><span>${r.time}</span><span>${r.tools}</span></div></div>
      ${r.note ? `<div class="bk-recipe-note">${r.note}</div>` : ""}
      <div class="bk-recipe-body">
        <div class="bk-ingredients"><div class="bk-where-label">Ingredients</div><ul>${r.ingredients.map(i => `<li>${i}</li>`).join("")}</ul></div>
        <ol class="bk-steps">${r.steps.map(([txt, who]) => `<li><span class="bk-step-text">${txt}</span><span class="bk-job ${BK_WHO[who][1]}">${BK_WHO[who][0]}</span></li>`).join("")}</ol>
      </div>
    </section>
  </div>`;
}
BK_ICON_PATHS.pin = `<path d="M12 21s-7-6.2-7-11.5a7 7 0 0 1 14 0C19 14.8 12 21 12 21z"/><circle cx="12" cy="9.5" r="2.5"/>`;
BK_ICON_PATHS.star = `<path d="M12 3l2.6 5.6 6.1.7-4.5 4.2 1.2 6L12 16.6 6.6 19.5l1.2-6L3.3 9.3l6.1-.7z"/>`;

// ---------- Served celebration ----------
function bkCelebrate(key, unlockedWeek) {
  const existing = document.querySelector(".bk-celebrate");
  if (existing) { if (!unlockedWeek) return; existing.remove(); }
  const kid = currentChild, keys = Object.keys(DATA);
  const served = keys.filter(k => stationDone(k)).length, left = keys.length - served;
  const treat = bkTreatForWeek(currentWeek());
  const reduce = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  let line2 = "";
  if (unlockedWeek && treat) line2 = `<div class="bk-cel-unlock">${bkPassportIcon(18)}You unlocked a treat from ${treat.city}!</div>
    <div class="bk-cel-btns"><button class="bk-btn primary" onclick="bkDismissCelebration();bkOpenPassport(${unlockedWeek})">See my treat</button><button class="bk-btn plain" onclick="bkDismissCelebration()">Later</button></div>`;
  else if (treat && left > 0) line2 = `<div class="bk-cel-sub2">${left} more to unlock this week’s treat from ${treat.city}</div>`;
  const colors = ["var(--bk-aw1)", "var(--bk-aw2)", "var(--bk-aw3)", "#F6D77C", "var(--bk-accent)"];
  const celStyle = (typeof bkKitchenFor === "function" ? bkKitchenFor().celebration : "sprinkles") || "sprinkles";
  const sprinkles = Array.from({ length: 14 }, (_, i) => {
    const c = colors[i % colors.length], angle = `--a:${i * (360 / 14)}deg;`;
    if (celStyle === "stars") return `<span class="bk-cel-spr shape" style="${angle}"><svg width="16" height="16" viewBox="0 0 24 24"><path d="M12 3l2.6 5.6 6.1.7-4.5 4.2 1.2 6L12 16.6 6.6 19.5l1.2-6L3.3 9.3l6.1-.7z" fill="${c}"/></svg></span>`;
    if (celStyle === "hearts") return `<span class="bk-cel-spr shape" style="${angle}"><svg width="16" height="16" viewBox="0 0 24 24"><path d="M12 21s-7-4.4-9.3-8.8C1.2 8.8 3 5.5 6.3 5c2-.3 3.7.6 4.7 2.3C12 5.6 13.7 4.7 15.7 5c3.3.5 5.1 3.8 3.6 7.2C17 16.6 12 21 12 21z" fill="${c}"/></svg></span>`;
    if (celStyle === "confetti") return `<span class="bk-cel-spr confetti" style="${angle}background:${c}"></span>`;
    return `<span class="bk-cel-spr" style="${angle}background:${c}"></span>`;
  }).join("");
  const el = document.createElement("div");
  el.className = "bk-celebrate" + (reduce ? " calm" : "") + (unlockedWeek ? " stay" : "");
  el.setAttribute("role", "status");
  el.setAttribute("aria-live", "polite");
  el.innerHTML = `<div class="bk-cel-card" onclick="event.stopPropagation()">
      <div class="bk-cel-stage">${sprinkles}<div class="bk-cel-plate">${bkPastry(bkPastryFor(key), kid, 84)}</div></div>
      <div class="bk-cel-title">Served!</div>
      <div class="bk-cel-sub">${bkTitle(DATA[key].name)} · ${served} of ${keys.length} plates served</div>
      ${line2}
    </div>`;
  el.addEventListener("click", bkDismissCelebration);
  document.body.appendChild(el);
  if (!unlockedWeek) setTimeout(bkDismissCelebration, reduce ? 2600 : 2400);
}
function bkDismissCelebration() {
  const el = document.querySelector(".bk-celebrate");
  if (!el) return;
  el.classList.add("out");
  setTimeout(() => el.remove(), 250);
}

// ---------- Hook Stage 2 into the Stage 1 renderer ----------
const bkStage1Render = window.render;
window.render = function render() {
  const kid = currentChild, isParent = currentView === "parent";
  if (isParent) bkPassport = null;
  const wkKey = `${kid}:${currentWeek()}`;
  const seeded = !!stampSeeded[kid];
  bkStage1Render();

  // Passport screens replace the kitchen (kid view only)
  const pv = document.getElementById("bkPassportView");
  const kitchenParts = ["bkHero", "sentBackBanner", "bkStationsWrap", "detailPanel"];
  if (!isParent && bkPassport !== null) {
    pv.innerHTML = bkPassport === "case" ? bkPassportCaseHTML() : bkTreatCardHTML(bkPassport);
    pv.style.display = "block";
    kitchenParts.forEach(id => { document.getElementById(id).style.display = "none"; });
  } else {
    pv.innerHTML = ""; pv.style.display = "none";
    kitchenParts.forEach(id => { document.getElementById(id).style.display = ""; });
    if (!isParent && !openStation) document.getElementById("detailPanel").style.display = "";
  }
  const pbtn = document.getElementById("bkPassportBtn");
  if (pbtn) pbtn.style.display = isParent ? "none" : "";

  // Newly served plates → celebrate (only after the first render for this child)
  const nowComplete = bkWeekComplete(currentWeek());
  const unlockedNow = seeded && nowComplete && bkWeekSeen[wkKey] === false && bkTreatForWeek(currentWeek()) ? currentWeek() : null;
  bkWeekSeen[wkKey] = nowComplete;
  if (!isParent && seeded) {
    const fresh = document.querySelector("#stationsGrid .bk-station.stamp-new");
    if (fresh && fresh.dataset.key && !fresh.dataset.celebrated) {
      fresh.dataset.celebrated = "1";
      setTimeout(() => bkCelebrate(fresh.dataset.key, unlockedNow), 150);
    } else if (unlockedNow) {
      setTimeout(() => bkCelebrate(Object.keys(DATA)[0], unlockedNow), 150);
    }
  }
};

// ======================================================================
// STAGE 3: Make It Your Kitchen (customization) + Word Recipe Cards
// ======================================================================

// ---------- Customization: options & persistence ----------
const BK_KITCHEN_OPTIONS = {
  pastries: ["croissant", "macaron", "cupcake", "donut", "cookie"],
  awnings: {
    kenley: [["mint", "Mint & Sky", ["#9ED6C6", "#A9D3EE", "#C4E3B5"]], ["ocean", "Ocean", ["#7FB8E0", "#B9DDF3", "#5E9BC9"]], ["sage", "Sage Garden", ["#C4E3B5", "#9ED6C6", "#E3F0D8"]], ["lagoon", "Lagoon", ["#8CD3D0", "#D7EFEA", "#A9D3EE"]]],
    adelyn: [["cotton", "Cotton Candy", ["#F6B8D0", "#CDB8EC", "#B9DDF3"]], ["berry", "Berry", ["#E893B8", "#F6B8D0", "#B99BE0"]], ["lilac", "Lilac Sky", ["#CDB8EC", "#B9DDF3", "#E6DCF7"]], ["bubblegum", "Bubblegum", ["#F6B8D0", "#FCE6EF", "#B9DDF3"]]]
  },
  styles: [["sweet", "Sweet Shop", "Sprinkles, stickers and extra treats"], ["pro", "Pro Kitchen", "Clean and simple, like a real pastry kitchen"]],
  lettering: [["classic", "Classic", "'Young Serif', Georgia, serif"], ["bubbly", "Bubbly", "'Fredoka', sans-serif"], ["fancy", "Fancy", "'Pacifico', cursive"]],
  patterns: ["plain", "polka", "gingham", "stripes", "sprinkles"],
  celebrations: ["sprinkles", "stars", "hearts", "confetti"],
  buddies: ["cat", "puppy", "unicorn", "none"]
};
const BK_KITCHEN_DEFAULTS = {
  kenley: { name: "Kenley’s Pâtisserie", pastry: "croissant", awning: "mint", style: "pro", lettering: "classic", pattern: "gingham", celebration: "stars", buddy: "cat", motto: "Bake it till you make it" },
  adelyn: { name: "Adelyn’s Pâtisserie", pastry: "macaron", awning: "cotton", style: "sweet", lettering: "fancy", pattern: "polka", celebration: "confetti", buddy: "unicorn", motto: "A sprinkle of sweetness every day" }
};
function bkKitchenFor(kid) {
  kid = kid || currentChild;
  let saved = {};
  try { saved = JSON.parse((childrenCache._settings || {})[`kitchen_${kid}`] || "{}"); } catch (e) {}
  return Object.assign({}, BK_KITCHEN_DEFAULTS[kid], saved);
}
function bkSaveKitchen(kid, patch) {
  const merged = Object.assign({}, bkKitchenFor(kid), patch);
  childrenCache._settings = childrenCache._settings || {};
  childrenCache._settings[`kitchen_${kid}`] = JSON.stringify(merged);
  apiPost("saveSetting", { key: `kitchen_${kid}`, value: JSON.stringify(merged) }).catch(() => {});
  return merged;
}
// Applies the saved awning colors + page pattern as CSS custom properties/attrs.
function bkApplyKitchenTheme() {
  const k = bkKitchenFor();
  const awn = (BK_KITCHEN_OPTIONS.awnings[currentChild] || []).find(a => a[0] === k.awning) || BK_KITCHEN_OPTIONS.awnings[currentChild][0];
  document.body.style.setProperty("--bk-aw1", awn[2][0]);
  document.body.style.setProperty("--bk-aw2", awn[2][1]);
  document.body.style.setProperty("--bk-aw3", awn[2][2]);
  const lett = BK_KITCHEN_OPTIONS.lettering.find(l => l[0] === k.lettering) || BK_KITCHEN_OPTIONS.lettering[0];
  document.body.style.setProperty("--bk-name-font", lett[2]);
  document.body.setAttribute("data-bk-pattern", k.pattern || "plain");
  document.body.setAttribute("data-bk-style", k.style || "pro");
}

// ---------- Sous-chef buddy artwork ----------
function bkBuddySVG(kind, size) {
  const s = size || 52;
  if (kind === "cat") return `<svg width="${s}" height="${s}" viewBox="0 0 64 64" aria-hidden="true"><path d="M15 30L17 10L29 21Z" fill="#E4E1DC" stroke="#8C8279" stroke-width="2.5" stroke-linejoin="round"/><path d="M49 30L47 10L35 21Z" fill="#E4E1DC" stroke="#8C8279" stroke-width="2.5" stroke-linejoin="round"/><circle cx="32" cy="37" r="19" fill="#E4E1DC" stroke="#8C8279" stroke-width="2.5"/><circle cx="25" cy="35" r="2.4" fill="#2B3B39"/><circle cx="39" cy="35" r="2.4" fill="#2B3B39"/><path d="M30 41h4l-2 2.5z" fill="#E88FA8"/><path d="M14 40h8M14 45l8-2M50 40h-8M50 45l-8-2" stroke="#8C8279" stroke-width="1.6" stroke-linecap="round"/></svg>`;
  if (kind === "puppy") return `<svg width="${s}" height="${s}" viewBox="0 0 64 64" aria-hidden="true"><circle cx="32" cy="36" r="18" fill="#F3DEC3" stroke="#B98A5E" stroke-width="2.5"/><ellipse cx="39" cy="33" rx="5.5" ry="5" fill="#E2C29E"/><path d="M14 22C8 24 7 38 12 44C16 40 19 32 20 26Z" fill="#C99A6B" stroke="#8F6440" stroke-width="2.5" stroke-linejoin="round"/><path d="M50 22C56 24 57 38 52 44C48 40 45 32 44 26Z" fill="#C99A6B" stroke="#8F6440" stroke-width="2.5" stroke-linejoin="round"/><circle cx="26" cy="34" r="2.4" fill="#2B3B39"/><circle cx="38" cy="34" r="2.4" fill="#2B3B39"/><ellipse cx="32" cy="41" rx="3.2" ry="2.4" fill="#3A2A20"/><path d="M28.5 45q3.5 3 7 0" stroke="#8F6440" stroke-width="1.8" stroke-linecap="round" fill="none"/><path d="M30.5 46.5q1.5 4 3 0z" fill="#F28CAB"/></svg>`;
  if (kind === "unicorn") return `<svg width="${s}" height="${s}" viewBox="0 0 64 64" aria-hidden="true"><path d="M17 27L16 13L26 21Z" fill="#FFFFFF" stroke="#B9A7C9" stroke-width="2.5" stroke-linejoin="round"/><path d="M47 27L48 13L38 21Z" fill="#FFFFFF" stroke="#B9A7C9" stroke-width="2.5" stroke-linejoin="round"/><circle cx="32" cy="38" r="18" fill="#FFFFFF" stroke="#B9A7C9" stroke-width="2.5"/><circle cx="22" cy="24" r="5" fill="#F7BCD3"/><circle cx="28" cy="21" r="4.5" fill="#CDB8EC"/><circle cx="37" cy="21.5" r="4.5" fill="#B9DDF3"/><path d="M32 4L28 21H36Z" fill="#F6D77C" stroke="#C9A43A" stroke-width="2" stroke-linejoin="round"/><path d="M29.8 12h4.4M29 16.5h6" stroke="#C9A43A" stroke-width="1.5"/><path d="M22 38q3-3 6 0M36 38q3-3 6 0" stroke="#2B3B39" stroke-width="2.2" stroke-linecap="round" fill="none"/><circle cx="21" cy="44" r="3" fill="#FCE6EF"/><circle cx="43" cy="44" r="3" fill="#FCE6EF"/><path d="M29 46q3 2.5 6 0" stroke="#B9A7C9" stroke-width="1.8" stroke-linecap="round" fill="none"/></svg>`;
  return "";
}
function bkBuddyBadgeHTML() {
  const k = bkKitchenFor();
  if (!k.buddy || k.buddy === "none") return "";
  return `<span class="bk-buddy-badge" aria-hidden="true">${bkBuddySVG(k.buddy, 44)}</span>`;
}

// ---------- Customize screen ----------
let bkCustomizeOpen = false;
let bkKitchenDraft = null;
function bkOpenCustomize() {
  bkKitchenDraft = bkKitchenFor();
  bkCustomizeOpen = true;
  render();
  window.scrollTo({ top: 0, behavior: "smooth" });
}
function bkCloseCustomize() { bkCustomizeOpen = false; bkKitchenDraft = null; render(); window.scrollTo({ top: 0 }); }
function bkSetDraft(field, value) { if (bkKitchenDraft) { bkKitchenDraft[field] = value; render(); } }
function bkSubmitKitchenName(input) { bkSetDraft("name", input.value); }
function bkSubmitKitchenMotto(input) { bkSetDraft("motto", input.value); }
function bkSaveKitchenDraft() {
  const nameEl = document.getElementById("bkKitchenNameInput");
  const mottoEl = document.getElementById("bkKitchenMottoInput");
  if (nameEl) bkKitchenDraft.name = nameEl.value;
  if (mottoEl) bkKitchenDraft.motto = mottoEl.value;
  bkSaveKitchen(currentChild, bkKitchenDraft);
  bkApplyKitchenTheme();
  bkCloseCustomize();
}
function bkChoiceBtn(active, onclick, inner, extraCls) {
  return `<button type="button" class="bk-choice${active ? " active" : ""}${extraCls ? " " + extraCls : ""}" aria-pressed="${active}" onclick="${onclick}">${inner}${active ? `<span class="bk-choice-check">${bkIcon("check", 12)}</span>` : ""}</button>`;
}
function bkPatternSwatchHTML(id) {
  if (id === "polka") return `<span class="bk-pat-swatch bk-pat-polka"></span>`;
  if (id === "gingham") return `<span class="bk-pat-swatch bk-pat-gingham"></span>`;
  if (id === "stripes") return `<span class="bk-pat-swatch bk-pat-stripes"></span>`;
  if (id === "sprinkles") return `<span class="bk-pat-swatch bk-pat-sprinkles"></span>`;
  return `<span class="bk-pat-swatch"></span>`;
}
function bkCelebrationSwatchHTML(id) {
  const c = ["var(--bk-aw1)", "var(--bk-aw2)", "var(--bk-aw3)"];
  let marks = "";
  if (id === "stars") marks = c.map((col, i) => `<svg width="16" height="16" viewBox="0 0 24 24" style="position:absolute;left:${18 + i * 24}px;top:${i % 2 ? 6 : 2}px"><path d="M12 3l2.6 5.6 6.1.7-4.5 4.2 1.2 6L12 16.6 6.6 19.5l1.2-6L3.3 9.3l6.1-.7z" fill="${col}"/></svg>`).join("");
  else if (id === "hearts") marks = c.map((col, i) => `<svg width="16" height="16" viewBox="0 0 24 24" style="position:absolute;left:${18 + i * 24}px;top:${i % 2 ? 6 : 2}px"><path d="M12 21s-7-4.4-9.3-8.8C1.2 8.8 3 5.5 6.3 5c2-.3 3.7.6 4.7 2.3C12 5.6 13.7 4.7 15.7 5c3.3.5 5.1 3.8 3.6 7.2C17 16.6 12 21 12 21z" fill="${col}"/></svg>`).join("");
  else if (id === "confetti") marks = c.map((col, i) => `<span style="position:absolute;left:${16 + i * 22}px;top:${i % 2 ? 8 : 4}px;width:9px;height:7px;border-radius:2px;background:${col};transform:rotate(${(i - 1) * 25}deg)"></span>`).join("");
  else marks = c.map((col, i) => `<span style="position:absolute;left:${16 + i * 22}px;top:${i % 2 ? 8 : 4}px;width:12px;height:5px;border-radius:3px;background:${col};transform:rotate(${(i - 1) * 20}deg)"></span>`).join("");
  return `<span class="bk-cel-swatch">${marks}</span>`;
}
function bkCustomizeHTML() {
  const kid = currentChild, name = CHILD_META[kid].name, d = bkKitchenDraft || bkKitchenFor();
  const awnList = BK_KITCHEN_OPTIONS.awnings[kid];
  const lett = BK_KITCHEN_OPTIONS.lettering.find(l => l[0] === d.lettering) || BK_KITCHEN_OPTIONS.lettering[0];
  return `<div class="bk-custom">
    <button class="bk-back" onclick="bkCloseCustomize()">${bkIcon("back", 18)}Back To ${name}’s Kitchen</button>
    <h1>Make It Your Kitchen</h1>
    <p class="bk-custom-sub">Pick a name, a signature pastry and your colors.</p>
    <div class="bk-custom-grid">
      <div class="bk-custom-main">
        <div class="bk-custom-field">
          <label class="bk-field-label" for="bkKitchenNameInput">Kitchen Name</label>
          <input id="bkKitchenNameInput" type="text" value="${bkAttr(d.name)}" oninput="bkSubmitKitchenName(this)">
        </div>
        <div class="bk-custom-field">
          <div class="bk-field-label">Signature Pastry</div>
          <div class="bk-choice-grid five">${BK_KITCHEN_OPTIONS.pastries.map(p => bkChoiceBtn(d.pastry === p, `bkSetDraft('pastry','${p}')`, `${bkPastry(p, kid, 44)}${bkTitle(p)}`)).join("")}</div>
        </div>
        <div class="bk-custom-field">
          <div class="bk-field-label">Awning Colors</div>
          <div class="bk-choice-grid four">${awnList.map(a => bkChoiceBtn(d.awning === a[0], `bkSetDraft('awning','${a[0]}')`, `<span class="bk-awn-swatch" style="background:linear-gradient(90deg,${a[2][0]} 33%,${a[2][1]} 33% 66%,${a[2][2]} 66%)"></span>${a[1]}`)).join("")}</div>
        </div>
        <div class="bk-custom-field">
          <div class="bk-field-label">Kitchen Style</div>
          <div class="bk-choice-grid two">${BK_KITCHEN_OPTIONS.styles.map(st => bkChoiceBtn(d.style === st[0], `bkSetDraft('style','${st[0]}')`, `<span class="bk-style-name">${st[1]}</span><span class="bk-style-sub">${st[2]}</span>`, "bk-choice-wide")).join("")}</div>
        </div>
      </div>
      <div class="bk-custom-preview">
        <div class="bk-field-label">Preview</div>
        <div class="bk-preview-card">
          <svg class="bk-awning" width="100%" height="30" aria-hidden="true" preserveAspectRatio="none"><defs><pattern id="bkPrevAwn" width="84" height="30" patternUnits="userSpaceOnUse"><rect width="28" height="16" fill="${awnList.find(a=>a[0]===d.awning)[2][0]}"/><rect x="28" width="28" height="16" fill="${awnList.find(a=>a[0]===d.awning)[2][1]}"/><rect x="56" width="28" height="16" fill="${awnList.find(a=>a[0]===d.awning)[2][2]}"/></pattern></defs><rect width="100%" height="30" fill="url(#bkPrevAwn)"/></svg>
          <div class="bk-preview-body bk-pat-${d.pattern}">
            <div class="bk-preview-emblem">${bkPastry(d.pastry, kid, 72)}</div>
            <div class="bk-preview-name" style="font-family:${lett[2]}">${bkAttr(d.name) || "Your Kitchen"}</div>
            <div class="bk-preview-motto">“${bkAttr(d.motto) || "Add a motto below"}”</div>
            <span class="bk-rank-chip small"><span class="bk-rank-hat">${bkIcon("hat", 14)}</span>${bkRankInfo().name}</span>
            ${d.buddy !== "none" ? `<span class="bk-preview-buddy">${bkBuddySVG(d.buddy, 52)}</span>` : ""}
          </div>
        </div>
      </div>
    </div>
    <div class="bk-custom-more">
      <h2>More Ways To Make It Yours</h2>
      <div class="bk-more-grid">
        <div class="bk-custom-field">
          <div class="bk-field-label">Name Lettering</div>
          <div class="bk-choice-grid three">${BK_KITCHEN_OPTIONS.lettering.map(l => bkChoiceBtn(d.lettering === l[0], `bkSetDraft('lettering','${l[0]}')`, `<span class="bk-lett-sample" style="font-family:${l[2]}">${name}</span>${l[1]}`)).join("")}</div>
        </div>
        <div class="bk-custom-field">
          <div class="bk-field-label">Page Pattern</div>
          <div class="bk-choice-grid five">${BK_KITCHEN_OPTIONS.patterns.map(p => bkChoiceBtn(d.pattern === p, `bkSetDraft('pattern','${p}')`, `${bkPatternSwatchHTML(p)}${bkTitle(p)}`)).join("")}</div>
        </div>
        <div class="bk-custom-field">
          <div class="bk-field-label">Celebration Style</div>
          <div class="bk-choice-grid four">${BK_KITCHEN_OPTIONS.celebrations.map(c => bkChoiceBtn(d.celebration === c, `bkSetDraft('celebration','${c}')`, `${bkCelebrationSwatchHTML(c)}${bkTitle(c)}`)).join("")}</div>
        </div>
        <div class="bk-custom-field">
          <div class="bk-field-label">Sous-Chef Buddy</div>
          <div class="bk-choice-grid four">${BK_KITCHEN_OPTIONS.buddies.map(b => bkChoiceBtn(d.buddy === b, `bkSetDraft('buddy','${b}')`, b === "none" ? `<span class="bk-no-buddy"></span>No Buddy` : `${bkBuddySVG(b, 40)}${bkTitle(b)}`)).join("")}</div>
        </div>
      </div>
    </div>
    <div class="bk-custom-footer">
      <div class="bk-custom-field grow">
        <label class="bk-field-label" for="bkKitchenMottoInput">Kitchen Motto</label>
        <input id="bkKitchenMottoInput" type="text" value="${bkAttr(d.motto)}" oninput="bkSubmitKitchenMotto(this)">
      </div>
      <button class="bk-btn primary big" onclick="bkSaveKitchenDraft()">Save My Kitchen</button>
    </div>
  </div>`;
}

// ---------- Word Recipe Cards ----------
// A morpheme-decomposition engine so a card can be built for ANY vocab word,
// plus a small curated set for words we know are being taught right now.
const BK_MORPHEMES = {
  prefixes: [
    ["incred", "in-", "not / without", "Latin"], // handled by curated entry, kept out of generic table
    ["un", "not / opposite of", "Old English"], ["re", "again / back", "Latin"], ["dis", "not / apart", "Latin"],
    ["mis", "wrongly", "Old English"], ["pre", "before", "Latin"], ["sub", "under", "Latin"], ["non", "not", "Latin"],
    ["inter", "between", "Latin"], ["trans", "across", "Latin"], ["anti", "against", "Greek"], ["semi", "half", "Latin"],
    ["super", "above / beyond", "Latin"], ["extra", "beyond", "Latin"], ["ex", "out of / former", "Latin"],
    ["im", "not", "Latin"], ["il", "not", "Latin"], ["ir", "not", "Latin"], ["in", "not / into", "Latin"],
    ["de", "down / away / reverse", "Latin"], ["over", "too much / above", "Old English"], ["under", "too little / below", "Old English"],
    ["mono", "one", "Greek"], ["uni", "one", "Latin"], ["bi", "two", "Latin"], ["tri", "three", "Latin"], ["poly", "many", "Greek"],
    ["micro", "small", "Greek"], ["auto", "self", "Greek"], ["tele", "far", "Greek"], ["mid", "middle", "Old English"]
  ],
  roots: [
    ["cred", "believe", "Latin"], ["duc", "lead", "Latin"], ["duct", "lead", "Latin"], ["spect", "look", "Latin"],
    ["scrib", "write", "Latin"], ["script", "write", "Latin"], ["port", "carry", "Latin"], ["dict", "say / speak", "Latin"],
    ["struct", "build", "Latin"], ["vis", "see", "Latin"], ["vid", "see", "Latin"], ["tract", "pull / drag", "Latin"],
    ["rupt", "break", "Latin"], ["ject", "throw", "Latin"], ["fer", "carry", "Latin"], ["mit", "send", "Latin"],
    ["miss", "send", "Latin"], ["pos", "place", "Latin"], ["pon", "place", "Latin"], ["grad", "step", "Latin"],
    ["gress", "step / go", "Latin"], ["cede", "go / yield", "Latin"], ["ceed", "go / yield", "Latin"], ["cess", "go / yield", "Latin"],
    ["cap", "take", "Latin"], ["cept", "take", "Latin"], ["voc", "call / voice", "Latin"], ["vok", "call", "Latin"],
    ["aud", "hear", "Latin"], ["photo", "light", "Greek"], ["graph", "write", "Greek"], ["bio", "life", "Greek"],
    ["geo", "earth", "Greek"], ["therm", "heat", "Greek"], ["meter", "measure", "Greek"], ["scope", "see / watch", "Greek"],
    ["phone", "sound", "Greek"], ["chron", "time", "Greek"], ["path", "feeling / suffering", "Greek"], ["log", "word / study", "Greek"],
    ["morph", "form / shape", "Greek"], ["sili", "jump / leap", "Latin (salire)"], ["ten", "hold", "Latin"], ["tain", "hold", "Latin"],
    ["sist", "stand", "Latin"], ["val", "worth / strength", "Latin"], ["fid", "faith / trust", "Latin"], ["metus", "fear", "Latin"],
    ["sens", "feel", "Latin"], ["sent", "feel", "Latin"], ["gen", "birth / kind", "Latin"], ["form", "shape", "Latin"],
    ["ven", "come", "Latin"], ["vert", "turn", "Latin"], ["vers", "turn", "Latin"], ["ann", "year", "Latin"], ["circ", "ring / around", "Latin"]
  ],
  suffixes: [
    ["ion", "act / state of", "Latin"], ["tion", "act / state of", "Latin"], ["ulous", "full of / tending to", "Latin"],
    ["able", "able to be", "Latin"], ["ible", "able to be", "Latin"], ["ent", "one who / state of", "Latin"],
    ["ant", "one who / state of", "Latin"], ["ive", "having the nature of", "Latin"], ["ous", "full of", "Latin"],
    ["ity", "state or quality of", "Latin"], ["ment", "result / act of", "Latin"], ["ful", "full of", "Old English"],
    ["less", "without", "Old English"], ["ly", "in a certain way", "Old English"], ["er", "one who / more", "Old English"],
    ["est", "most", "Old English"], ["ize", "to make", "Greek"], ["ist", "one who practices", "Greek"], ["ology", "study of", "Greek"]
  ]
};
// Curated cards for words currently in the girls' lesson content, hand-checked
// so the etymology is right even where the generic table would be too rough.
const BK_WORD_CARDS = {
  incredulous: { prefix: ["in-", "not"], root: ["cred", "believe"], suffix: ["-ulous", "full of / tending to"], origin: "Latin", meaning: "not willing or able to believe something; showing disbelief", example: "She gave an incredulous look when her brother said he'd cleaned his whole room in five minutes.", tip: "Same root as \"credit\" and \"credible\" — all about believing." },
  deduction: { prefix: ["de-", "down / from"], root: ["duct", "lead"], suffix: ["-ion", "act of"], origin: "Latin", meaning: "a conclusion reached by reasoning from general facts to a specific one", example: "From the muddy footprints, the detective made a deduction about which door the thief had used.", tip: "\"Lead down\" from facts to a conclusion — like Sherlock Holmes." },
  meticulous: { prefix: null, root: ["metus", "fear"], suffix: ["-ulous", "full of / tending to"], origin: "Latin", meaning: "showing great attention to detail; very careful and precise", example: "Adelyn was meticulous about lining up every sticker perfectly in her passport.", tip: "Comes from a word for \"fearful\" — being so careful you're almost afraid to make a mistake." },
  resilient: { prefix: ["re-", "back"], root: ["sili", "jump / leap"], suffix: ["-ent", "state of"], origin: "Latin", meaning: "able to recover quickly from difficulties; springing back into shape", example: "After a rough first attempt, Kenley was resilient and tried the recipe again the next day.", tip: "Think of a rubber band \"jumping back\" into shape." }
};
// Falls back to a generic morpheme breakdown for any word not in the curated list.
function bkAnalyzeWord(word) {
  const w = String(word || "").toLowerCase().trim();
  if (!w) return null;
  if (BK_WORD_CARDS[w]) return Object.assign({ word: w }, BK_WORD_CARDS[w]);
  const findSeg = (list, fromStart) => {
    let best = null;
    list.forEach(([seg, meaning, origin]) => {
      if (seg.length < 2) return;
      const hit = fromStart ? w.startsWith(seg) : w.endsWith(seg);
      if (hit && (!best || seg.length > best[0].length) && seg.length < w.length - 1) best = [seg, meaning, origin];
    });
    return best;
  };
  const pre = findSeg(BK_MORPHEMES.prefixes, true);
  const suf = findSeg(BK_MORPHEMES.suffixes, false);
  let mid = w;
  if (pre) mid = mid.slice(pre[0].length);
  if (suf && mid.endsWith(suf[0])) mid = mid.slice(0, mid.length - suf[0].length);
  const rootMatch = BK_MORPHEMES.roots.find(([seg]) => mid.includes(seg) || w.includes(seg));
  if (!pre && !suf && !rootMatch) return { word: w, prefix: null, root: null, suffix: null, origin: null, meaning: null, example: null, tip: null, unknown: true };
  return {
    word: w,
    prefix: pre ? [pre[0] + "-", pre[1]] : null,
    root: rootMatch ? [rootMatch[0], rootMatch[1]] : null,
    suffix: suf ? ["-" + suf[0], suf[1]] : null,
    origin: (rootMatch && rootMatch[2]) || (pre && pre[2]) || (suf && suf[2]) || null,
    meaning: null, example: null, tip: null
  };
}
function bkIngredientChip(label, part) {
  if (!part) return "";
  return `<div class="bk-ingredient"><span class="bk-ingredient-part">${label}</span><span class="bk-ingredient-seg">${part[0]}</span><span class="bk-ingredient-mean">${part[1]}</span></div>`;
}
function bkWordCardHTML(word) {
  const a = bkAnalyzeWord(word);
  if (!a) return `<div class="bk-recipecard empty">Type a word to see its recipe card.</div>`;
  if (a.unknown) {
    return `<div class="bk-recipecard">
      <div class="bk-rc-head"><span class="bk-rc-word">${bkTitle(a.word)}</span></div>
      <div class="bk-rc-empty">We don’t have a breakdown for this word yet — but you can still look up what it means and add it to the review bank.</div>
    </div>`;
  }
  return `<div class="bk-recipecard">
    <div class="bk-rc-head"><span class="bk-rc-word">${bkTitle(a.word)}</span>${a.origin ? `<span class="bk-rc-origin">${a.origin}</span>` : ""}</div>
    <div class="bk-rc-ingredients">
      ${bkIngredientChip("Prefix", a.prefix)}
      ${bkIngredientChip("Root", a.root)}
      ${bkIngredientChip("Suffix", a.suffix)}
    </div>
    ${a.meaning ? `<div class="bk-rc-block"><div class="bk-field-label">Meaning</div><p>${a.meaning}</p></div>` : ""}
    ${a.example ? `<div class="bk-rc-block"><div class="bk-field-label">In A Sentence</div><p>${a.example}</p></div>` : ""}
    ${a.tip ? `<div class="bk-rc-block bk-rc-tip"><div class="bk-field-label">${bkIcon("bulb", 14)}Memory Tip</div><p>${a.tip}</p></div>` : ""}
  </div>`;
}

// ---------- Word Recipe Box (browse words she's practiced + curated set) ----------
let bkRecipeBoxOpen = false;
let bkRecipeBoxWord = null;
function bkOpenRecipeBox() { bkRecipeBoxOpen = true; bkRecipeBoxWord = Object.keys(BK_WORD_CARDS)[0] || null; render(); window.scrollTo({ top: 0, behavior: "smooth" }); }
function bkCloseRecipeBox() { bkRecipeBoxOpen = false; render(); window.scrollTo({ top: 0 }); }
function bkPickRecipeWord(w) { bkRecipeBoxWord = w; render(); }
function bkRecipeBoxWordList() {
  const practiced = (reviewPoolCache[currentChild] || []).map(r => r.word).filter(Boolean);
  const curated = Object.keys(BK_WORD_CARDS);
  const seen = {}, out = [];
  curated.concat(practiced).forEach(w => { const key = String(w).toLowerCase(); if (!seen[key]) { seen[key] = true; out.push(w); } });
  return out;
}
function bkRecipeBoxHTML() {
  const name = CHILD_META[currentChild].name;
  const words = bkRecipeBoxWordList();
  const active = bkRecipeBoxWord || words[0];
  return `<div class="bk-custom">
    <button class="bk-back" onclick="bkCloseRecipeBox()">${bkIcon("back", 18)}Back To ${name}’s Kitchen</button>
    <h1>Word Recipe Box</h1>
    <p class="bk-custom-sub">Every word breaks down into ingredients — a prefix, a root and a suffix. Pick a word to see its card.</p>
    <div class="bk-recipebox-grid">
      <div class="bk-recipebox-list">
        ${words.length ? words.map(w => `<button class="bk-word-chip${String(w).toLowerCase() === String(active).toLowerCase() ? " active" : ""}" onclick="bkPickRecipeWord('${bkAttr(w)}')">${bkTitle(w)}</button>`).join("") : `<div class="bk-rc-empty">No practiced words yet — check back after a few vocabulary rounds.</div>`}
      </div>
      <div class="bk-recipebox-card">${active ? bkWordCardHTML(active) : ""}</div>
    </div>
  </div>`;
}

// ---------- Hook Stage 3 views into the renderer ----------
const bkStage2Render = window.render;
window.render = function render() {
  const isParent = currentView === "parent";
  if (isParent) { bkCustomizeOpen = false; bkRecipeBoxOpen = false; }
  bkApplyKitchenTheme();
  bkStage2Render();

  const pv = document.getElementById("bkPassportView");
  const kitchenParts = ["bkHero", "sentBackBanner", "bkStationsWrap", "detailPanel"];
  if (!isParent && (bkCustomizeOpen || bkRecipeBoxOpen)) {
    pv.innerHTML = bkCustomizeOpen ? bkCustomizeHTML() : bkRecipeBoxHTML();
    pv.style.display = "block";
    kitchenParts.forEach(id => { const el = document.getElementById(id); if (el) el.style.display = "none"; });
    const pbtn = document.getElementById("bkPassportBtn");
    if (pbtn) pbtn.style.display = "none";
    const cbtn = document.getElementById("bkCustomizeBtn"), rbtn = document.getElementById("bkRecipeBoxBtn");
    if (cbtn) cbtn.style.display = "none";
    if (rbtn) rbtn.style.display = "none";
  }

  // Buddy mascot rides along on the kid-view hero, next to the passport teaser.
  if (!isParent && !bkCustomizeOpen && !bkRecipeBoxOpen && bkPassport === null) {
    const hero = document.querySelector("#bkHero .bk-hero");
    if (hero && !hero.querySelector(".bk-buddy-badge")) hero.insertAdjacentHTML("beforeend", bkBuddyBadgeHTML());
  }
};
