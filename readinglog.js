// The ELA Pastry Kitchen — Reading Log station (beta only).
// Loaded AFTER beta.js. Adds a 6th station card that never counts toward the
// week's plates, grades, or advancing — it's just a place to write down reading.
//
// Kenley: a simple list of the books she's reading / has finished this year.
// Adelyn: a daily log (book defaults to that week's assigned Reading book,
// pages from–to, optional minutes), per-book progress, and "other reading".
//
// Storage follows the Projects pattern: every entry is its own Submissions row,
// so no Apps Script change is needed and the offline queue just works.
//   rbook_<id>  answers.readingBook = { id, title, author, totalPages, finished, finishedOn, addedOn, deleted }
//   rlog_<id>   answers.readingLog  = { id, kind: "assigned"|"extra", book, date, fromPage, toPage, minutes, note, week, deleted }

const readingCache = {}; // { kenley: { books: [...], logs: [...] }, adelyn: {...} }
let rlOpen = false;
let rlShowAll = false;
let rlExtraOpen = false;
let rlPagesEdit = null; // book title whose total-pages field is open
let rlProgressTitles = []; // titles in the order drawn, so buttons pass an index instead of a quoted title
let rlDraft = {};

function extractReading(submissions) {
  const books = [], logs = [];
  (submissions || []).forEach(s => {
    const id = String(s.task_id || "");
    if (id.indexOf("rbook_") === 0 && s.answers && s.answers.readingBook) books.push(Object.assign({}, s.answers.readingBook, { id }));
    if (id.indexOf("rlog_") === 0 && s.answers && s.answers.readingLog) logs.push(Object.assign({}, s.answers.readingLog, { id }));
  });
  return { books, logs };
}

const rlOrigBootstrap = window.apiGetBootstrap;
window.apiGetBootstrap = async function apiGetBootstrap(student) {
  const resp = await rlOrigBootstrap(student);
  readingCache[student] = extractReading(resp && resp.submissions);
  return resp;
};

function rlData() { return readingCache[currentChild] || (readingCache[currentChild] = { books: [], logs: [] }); }
function rlBooks() { return rlData().books.filter(b => !b.deleted); }
function rlLogs() { return rlData().logs.filter(l => !l.deleted); }
function rlNewId(prefix) { return prefix + Date.now().toString(36) + Math.random().toString(36).slice(2, 5); }
function rlSame(a, b) { return String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase(); }
function rlToday() { const d = new Date(); return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10); }
function rlFmtDate(iso) {
  if (!iso) return "";
  const d = new Date(iso + "T12:00:00");
  return isNaN(d) ? escHtml(iso) : d.toLocaleDateString(undefined, { weekday: "short", month: "short", day: "numeric" });
}
function rlFmtMin(m) {
  m = Number(m) || 0;
  if (!m) return "";
  const h = Math.floor(m / 60), r = m % 60;
  return h ? `${h} hr${r ? ` ${r} min` : ""}` : `${r} min`;
}
function rlNum(v) { const n = parseInt(v, 10); return Number.isFinite(n) && n >= 0 ? n : null; }

function rlPersist(kind, item) {
  const key = kind === "book" ? "readingBook" : "readingLog";
  apiPost("saveSubmission", {
    student: currentChild, task_id: item.id, status: "log", score: "", parent_comment: "",
    answers: { [key]: item }
  }).catch(() => {});
}

// This week's assigned Reading book, from the Reading subject's tag ("Harriet the Invincible · RL.4.3").
function rlAssignedBook() {
  const r = DATA && DATA.reading;
  if (!r) return "";
  const tag = r.tagsByWeek[currentWeek()] || r.tag || "";
  return String(tag).split(" · ")[0].trim();
}
function rlLastPage(book) {
  let max = 0;
  rlLogs().forEach(l => { if (rlSame(l.book, book) && Number(l.toPage) > max) max = Number(l.toPage); });
  return max;
}
function rlBookMeta(title) { return rlBooks().find(b => rlSame(b.title, title)); }
function rlKnownTitles() {
  const seen = new Map();
  const add = t => { const k = String(t || "").trim(); if (k && !seen.has(k.toLowerCase())) seen.set(k.toLowerCase(), k); };
  add(rlAssignedBook());
  rlLogs().slice().sort((a, b) => String(b.date).localeCompare(String(a.date))).forEach(l => add(l.book));
  rlBooks().forEach(b => add(b.title));
  return Array.from(seen.values());
}

// ---------- Actions ----------
function rlToggle() {
  rlOpen = !rlOpen;
  if (rlOpen) { openStation = null; rlDraft = {}; }
  render();
  if (rlOpen) { const p = document.getElementById("detailPanel"); if (p) p.scrollIntoView({ behavior: "smooth", block: "start" }); }
}
function rlVal(id) { const el = document.getElementById(id); return el ? el.value.trim() : ""; }
function rlSay(msg) { const el = document.getElementById("rlMsg"); if (el) { el.textContent = msg; el.style.display = msg ? "block" : "none"; } }

function rlBookInput(el) {
  rlDraft.book = el.value;
  const from = document.getElementById("rlFrom");
  if (from && !rlDraft.fromTyped) { const last = rlLastPage(el.value); from.value = last ? last + 1 : ""; }
}

function rlAddDaily() {
  const book = rlVal("rlBook"), date = rlVal("rlDate") || rlToday();
  const fromPage = rlNum(rlVal("rlFrom")), toPage = rlNum(rlVal("rlTo")), minutes = rlNum(rlVal("rlMin"));
  if (!book) return rlSay("Which book? Type the title first.");
  if (toPage === null) return rlSay("Add the page you stopped on.");
  if (fromPage !== null && toPage < fromPage) return rlSay("The last page has to be the same as or after the first page.");
  const entry = { id: rlNewId("rlog_"), kind: "assigned", book, date, fromPage, toPage, minutes, note: "", week: currentWeek(), addedOn: new Date().toISOString() };
  rlData().logs.push(entry);
  rlPersist("log", entry);
  rlDraft = {};
  render();
}
function rlAddExtra() {
  const book = rlVal("rlxBook"), date = rlVal("rlxDate") || rlToday();
  const fromPage = rlNum(rlVal("rlxFrom")), toPage = rlNum(rlVal("rlxTo")), minutes = rlNum(rlVal("rlxMin")), note = rlVal("rlxNote");
  if (!book) { const el = document.getElementById("rlxMsg"); if (el) { el.textContent = "What did she read? Add a title."; el.style.display = "block"; } return; }
  const entry = { id: rlNewId("rlog_"), kind: "extra", book, date, fromPage, toPage, minutes, note, week: currentWeek(), addedOn: new Date().toISOString() };
  rlData().logs.push(entry);
  rlPersist("log", entry);
  rlDraft = {};
  rlExtraOpen = false;
  render();
}
function rlRemoveLog(id) {
  const l = rlData().logs.find(x => x.id === id);
  if (!l || !confirm(`Remove this entry (${l.book}, ${rlFmtDate(l.date)})?`)) return;
  l.deleted = true;
  rlPersist("log", l);
  render();
}
// Book records: Kenley's whole list, and Adelyn's optional total-pages / finished flags.
function rlUpsertBook(title, changes) {
  let b = rlBookMeta(title);
  if (!b) { b = { id: rlNewId("rbook_"), title, author: "", totalPages: null, finished: false, finishedOn: "", addedOn: rlToday() }; rlData().books.push(b); }
  Object.assign(b, changes);
  rlPersist("book", b);
  return b;
}
function rlSavePages(i) {
  const title = rlProgressTitles[i];
  const n = rlNum(rlVal("rlPages"));
  rlUpsertBook(title, { totalPages: n || null });
  rlPagesEdit = null;
  render();
}
function rlEditPages(i) { rlPagesEdit = i === null ? null : rlProgressTitles[i]; render(); }
function rlToggleFinished(id) {
  const b = rlData().books.find(x => x.id === id);
  if (!b) return;
  b.finished = !b.finished;
  b.finishedOn = b.finished ? rlToday() : "";
  rlPersist("book", b);
  render();
}
function rlToggleFinishedTitle(i) {
  const title = rlProgressTitles[i];
  const b = rlBookMeta(title);
  if (b) rlToggleFinished(b.id);
  else { rlUpsertBook(title, { finished: true, finishedOn: rlToday() }); render(); }
}
function rlAddKenleyBook() {
  const title = rlVal("rlkTitle"), author = rlVal("rlkAuthor");
  const done = document.getElementById("rlkDone") && document.getElementById("rlkDone").checked;
  if (!title) return rlSay("Type the book’s title first.");
  if (rlBookMeta(title)) return rlSay("That book is already on your list.");
  const b = { id: rlNewId("rbook_"), title, author, totalPages: null, finished: !!done, finishedOn: done ? rlToday() : "", addedOn: rlToday() };
  rlData().books.push(b);
  rlPersist("book", b);
  rlDraft = {};
  render();
}
function rlRemoveBook(id) {
  const b = rlData().books.find(x => x.id === id);
  if (!b || !confirm(`Remove “${b.title}” from the list?`)) return;
  b.deleted = true;
  rlPersist("book", b);
  render();
}
function rlToggleAll() { rlShowAll = !rlShowAll; render(); }
function rlToggleExtra() { rlExtraOpen = !rlExtraOpen; render(); }

// ---------- Station card ----------
function rlLoggedThisWeek() {
  const days = new Set();
  rlLogs().forEach(l => { if (Number(l.week) === currentWeek()) days.add(l.date); });
  return days.size;
}
function rlStationCardHTML(idx) {
  const th = bkTheme();
  const tile = th.tiles[idx % th.tiles.length];
  let pill, tag, count;
  if (currentChild === "kenley") {
    const books = rlBooks(), fin = books.filter(b => b.finished).length, reading = books.length - fin;
    pill = fin ? bkPill("served", `${fin} finished`, "check") : bkPill("neutral", "Your book list");
    tag = reading ? `Reading ${reading === 1 ? "1 book" : reading + " books"} right now` : "Books you read this year";
    count = `${books.length} ${bkPlural(books.length, "book", "books")} this year`;
  } else {
    const loggedToday = rlLogs().some(l => l.date === rlToday());
    const days = rlLoggedThisWeek();
    pill = loggedToday ? bkPill("served", "Logged today", "check") : bkPill("neutral", "Log when you read");
    tag = rlAssignedBook() || "What we read each day";
    count = `${days} ${bkPlural(days, "day", "days")} logged this week`;
  }
  return `<span class="bk-st-top"><span class="bk-st-tile rl-tile" style="background:${tile}">${bkIcon("book", 44)}</span>${pill}</span>
    <span class="bk-st-title">Reading Log</span>
    <span class="bk-st-tag">${escHtml(tag)}</span>
    <span class="bk-st-count">${count}</span>`;
}

// ---------- Panel ----------
function rlDatalist() { return `<datalist id="rlTitles">${rlKnownTitles().map(t => `<option value="${escHtml(t)}"></option>`).join("")}</datalist>`; }
function rlD(k, fallback) { return escHtml(rlDraft[k] !== undefined ? rlDraft[k] : (fallback == null ? "" : fallback)); }

function rlKenleyHTML() {
  const books = rlBooks().slice().sort((a, b) => String(b.finishedOn || b.addedOn).localeCompare(String(a.finishedOn || a.addedOn)));
  const reading = books.filter(b => !b.finished), finished = books.filter(b => b.finished);
  const row = b => `<div class="rl-book${b.finished ? " finished" : ""}">
      <span class="rl-book-ico">${bkIcon(b.finished ? "check" : "book", 18)}</span>
      <span class="rl-book-text"><span class="rl-book-title">${escHtml(b.title)}</span>${b.author ? `<span class="rl-book-sub">by ${escHtml(b.author)}</span>` : ""}<span class="rl-book-sub">${b.finished ? `Finished ${rlFmtDate(b.finishedOn)}` : `Added ${rlFmtDate(b.addedOn)}`}</span></span>
      <span class="rl-book-actions">
        <button class="bk-btn ${b.finished ? "plain" : "outline"}" onclick="rlToggleFinished('${b.id}')">${b.finished ? "Still reading" : "I finished it!"}</button>
        <button class="bk-btn plain rl-x" onclick="rlRemoveBook('${b.id}')" aria-label="Remove ${escHtml(b.title)}">${bkIcon("close", 16)}</button>
      </span></div>`;
  return `<div class="rl-form">
      <div class="rl-form-title">Add a book</div>
      <div class="rl-grid">
        <label class="rl-field wide">Title<input type="text" id="rlkTitle" value="${rlD("kTitle")}" oninput="rlDraft.kTitle=this.value" placeholder="Book title"></label>
        <label class="rl-field"><span>Author <span class="rl-opt">(optional)</span></span><input type="text" id="rlkAuthor" value="${rlD("kAuthor")}" oninput="rlDraft.kAuthor=this.value"></label>
      </div>
      <label class="rl-check"><input type="checkbox" id="rlkDone"> I already finished it</label>
      <div class="rl-msg" id="rlMsg" style="display:none;"></div>
      <button class="bk-btn primary" onclick="rlAddKenleyBook()">${bkIcon("plus", 16)}Add to my list</button>
    </div>
    <div class="rl-group"><div class="rl-group-head">Reading now · ${reading.length}</div>${reading.length ? reading.map(row).join("") : `<div class="empty-note">Nothing in progress. Add the book you’re reading above.</div>`}</div>
    <div class="rl-group"><div class="rl-group-head">Finished this year · ${finished.length}</div>${finished.length ? finished.map(row).join("") : `<div class="empty-note">No finished books yet.</div>`}</div>`;
}

function rlAdelynHTML() {
  const assigned = rlAssignedBook();
  const book = rlDraft.book !== undefined ? rlDraft.book : assigned;
  const last = rlLastPage(book);
  const logs = rlLogs().slice().sort((a, b) => String(b.date).localeCompare(String(a.date)) || String(b.addedOn).localeCompare(String(a.addedOn)));

  // Progress per book, from every entry with pages.
  const progress = new Map();
  logs.forEach(l => {
    const k = String(l.book).trim().toLowerCase();
    const p = progress.get(k) || { title: l.book, maxPage: 0, sessions: 0, minutes: 0, lastDate: l.date };
    p.sessions++; p.minutes += Number(l.minutes) || 0;
    if (Number(l.toPage) > p.maxPage) p.maxPage = Number(l.toPage);
    progress.set(k, p);
  });
  const progressList = Array.from(progress.values()).filter(p => p.maxPage > 0);
  rlProgressTitles = progressList.map(p => p.title);
  const progressRows = progressList.map((p, i) => {
    const meta = rlBookMeta(p.title) || {};
    const total = Number(meta.totalPages) || 0;
    const pct = total ? Math.min(100, Math.round(p.maxPage / total * 100)) : null;
    const editing = rlSame(rlPagesEdit, p.title);
    return `<div class="rl-prog${meta.finished ? " finished" : ""}">
        <div class="rl-prog-top"><span class="rl-book-title">${escHtml(p.title)}</span>${meta.finished ? bkPill("served", "Finished", "check") : ""}</div>
        <div class="rl-prog-line">Up to page ${p.maxPage}${total ? ` of ${total}` : ""} · ${p.sessions} ${bkPlural(p.sessions, "session", "sessions")}${p.minutes ? ` · ${rlFmtMin(p.minutes)}` : ""}</div>
        ${pct !== null ? `<div class="rl-prog-bar"><span style="width:${pct}%"></span></div>` : ""}
        <div class="rl-prog-actions">
          ${editing
            ? `<label class="rl-inline">Pages in book <input type="number" id="rlPages" min="1" value="${total || ""}"></label><button class="bk-btn primary" onclick="rlSavePages(${i})">Save</button><button class="bk-btn plain" onclick="rlEditPages(null)">Cancel</button>`
            : `<button class="bk-btn plain" onclick="rlEditPages(${i})">${total ? "Change page count" : "Add page count"}</button><button class="bk-btn ${meta.finished ? "plain" : "outline"}" onclick="rlToggleFinishedTitle(${i})">${meta.finished ? "Not finished yet" : "Mark finished"}</button>`}
        </div></div>`;
  }).join("");

  const shown = rlShowAll ? logs : logs.slice(0, 10);
  const entryRow = l => `<div class="rl-entry">
      <span class="rl-entry-date">${rlFmtDate(l.date)}</span>
      <span class="rl-entry-main"><span class="rl-entry-book">${escHtml(l.book)}${l.kind === "extra" ? ` <span class="rl-tag">Other reading</span>` : ""}</span>
        <span class="rl-entry-sub">${[l.toPage != null ? (l.fromPage != null ? `pp. ${l.fromPage}–${l.toPage}` : `to p. ${l.toPage}`) : "", rlFmtMin(l.minutes), l.note ? escHtml(l.note) : ""].filter(Boolean).join(" · ") || "&nbsp;"}</span></span>
      <button class="bk-btn plain rl-x" onclick="rlRemoveLog('${l.id}')" aria-label="Remove entry">${bkIcon("close", 16)}</button></div>`;

  const extraForm = rlExtraOpen ? `<div class="rl-form rl-extra">
      <div class="rl-form-title">Other reading she did</div>
      <div class="rl-grid">
        <label class="rl-field wide">What she read<input type="text" id="rlxBook" list="rlTitles" value="${rlD("xBook")}" oninput="rlDraft.xBook=this.value" placeholder="Book, magazine, comic…"></label>
        <label class="rl-field">Date<input type="date" id="rlxDate" value="${rlD("xDate", rlToday())}" oninput="rlDraft.xDate=this.value"></label>
        <label class="rl-field"><span>From page <span class="rl-opt">(optional)</span></span><input type="number" id="rlxFrom" min="0" inputmode="numeric" value="${rlD("xFrom")}" oninput="rlDraft.xFrom=this.value"></label>
        <label class="rl-field"><span>To page <span class="rl-opt">(optional)</span></span><input type="number" id="rlxTo" min="0" inputmode="numeric" value="${rlD("xTo")}" oninput="rlDraft.xTo=this.value"></label>
        <label class="rl-field"><span>Minutes <span class="rl-opt">(optional)</span></span><input type="number" id="rlxMin" min="0" inputmode="numeric" value="${rlD("xMin")}" oninput="rlDraft.xMin=this.value"></label>
        <label class="rl-field full"><span>Note <span class="rl-opt">(optional)</span></span><input type="text" id="rlxNote" value="${rlD("xNote")}" oninput="rlDraft.xNote=this.value" placeholder="e.g. read it to her sister, finished the series"></label>
      </div>
      <div class="rl-msg" id="rlxMsg" style="display:none;"></div>
      <div class="rl-btns"><button class="bk-btn primary" onclick="rlAddExtra()">Save</button><button class="bk-btn plain" onclick="rlToggleExtra()">Cancel</button></div>
    </div>` : `<button class="bk-btn outline rl-extra-btn" onclick="rlToggleExtra()">${bkIcon("plus", 16)}Add other reading she did</button>`;

  return `${rlDatalist()}
    <div class="rl-form">
      <div class="rl-form-title">Log today’s reading</div>
      <div class="rl-grid">
        <label class="rl-field wide">Book<input type="text" id="rlBook" list="rlTitles" value="${escHtml(book)}" oninput="rlBookInput(this)"></label>
        <label class="rl-field">Date<input type="date" id="rlDate" value="${rlD("date", rlToday())}" oninput="rlDraft.date=this.value"></label>
        <label class="rl-field">From page<input type="number" id="rlFrom" min="0" inputmode="numeric" value="${rlD("from", last ? last + 1 : "")}" oninput="rlDraft.from=this.value;rlDraft.fromTyped=true"></label>
        <label class="rl-field">To page<input type="number" id="rlTo" min="0" inputmode="numeric" value="${rlD("to")}" oninput="rlDraft.to=this.value"></label>
        <label class="rl-field"><span>Minutes <span class="rl-opt">(optional)</span></span><input type="number" id="rlMin" min="0" inputmode="numeric" value="${rlD("min")}" oninput="rlDraft.min=this.value"></label>
      </div>
      ${assigned ? `<div class="rl-hint">This week’s book is <b>${escHtml(assigned)}</b>${last ? `. Last time you stopped on page ${last}.` : "."}</div>` : ""}
      <div class="rl-msg" id="rlMsg" style="display:none;"></div>
      <button class="bk-btn primary" onclick="rlAddDaily()">${bkIcon("check", 16)}Log it</button>
    </div>
    ${extraForm}
    ${progressRows ? `<div class="rl-group"><div class="rl-group-head">Book progress</div><div class="rl-prog-grid">${progressRows}</div></div>` : ""}
    <div class="rl-group"><div class="rl-group-head">Reading days · ${logs.length}</div>
      ${logs.length ? shown.map(entryRow).join("") : `<div class="empty-note">Nothing logged yet. Reading doesn’t have to happen every day; log it when it does.</div>`}
      ${logs.length > 10 ? `<button class="bk-btn plain" onclick="rlToggleAll()">${rlShowAll ? "Show fewer" : `Show all ${logs.length}`}</button>` : ""}
    </div>`;
}

function rlPanelHTML(withClose) {
  const kid = currentChild;
  return `<div class="detail-head">
      <div class="bk-detail-title"><span class="bk-st-tile small rl-tile">${bkIcon("book", 30)}</span><div><div class="detail-tag">${kid === "kenley" ? "Books I read this year" : "Doesn’t count toward the week. Log it when you read."}</div><div class="detail-title">${escHtml(CHILD_META[kid].name)}’s Reading Log</div></div></div>
      ${withClose ? `<div class="bk-detail-actions"><button class="detail-close bk-btn plain" onclick="rlToggle()">${bkIcon("close", 16)}Close</button></div>` : ""}
    </div>
    <div class="rl-body">${kid === "kenley" ? rlKenleyHTML() : rlAdelynHTML()}</div>`;
}

// ---------- Hook into render ----------
const rlOrigRender = window.render;
window.render = function render() {
  rlOrigRender();
  const isParent = currentView === "parent";
  const grid = document.getElementById("stationsGrid");
  const panel = document.getElementById("detailPanel");
  const parentBox = document.getElementById("rlParentBox");
  if (!isParent && grid) {
    const n = grid.children.length;
    const card = document.createElement("button");
    card.type = "button";
    card.className = "station bk-station rl-station" + (rlOpen ? " active" : "");
    card.onclick = rlToggle;
    card.innerHTML = rlStationCardHTML(n);
    grid.appendChild(card);
    grid.className = `stations bk-stations count-${n + 1}`;
    if (rlOpen) { panel.className = "detail open bk-detail rl-detail"; panel.innerHTML = rlPanelHTML(true); }
  }
  if (parentBox) {
    parentBox.style.display = isParent ? "block" : "none";
    parentBox.innerHTML = isParent ? rlPanelHTML(false) : "";
  }
};

// Opening a regular station closes the log, and vice versa; switching kids closes it.
const rlOrigOpenStation = window.openStationFn;
window.openStationFn = function openStationFn(key) { rlOpen = false; rlOrigOpenStation(key); };
const rlOrigSwitchChild = window.switchChild;
window.switchChild = function switchChild(id) { if (id !== currentChild) { rlOpen = false; rlDraft = {}; rlShowAll = false; rlExtraOpen = false; rlPagesEdit = null; } return rlOrigSwitchChild(id); };
