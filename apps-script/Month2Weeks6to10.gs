/**
 * The ELA Pastry Kitchen — Month 2, Weeks 6–10 content (Kenley).
 *
 * Generated Oct 13, 2026 (Kristi + Claude chat). Same pattern as
 * Month1Weeks2to5.gs. Purely additive: does not touch setupSheets() or any
 * existing rows, and never touches Adelyn's data.
 *
 * Run these from the Apps Script function dropdown, in this order. (Apps Script hides functions whose names END IN AN UNDERSCORE from the
 * dropdown, so use the run... wrapper names below; each just calls the underscore function next to it.)
 *
 *   1. runSeedMonth2Weeks6to10()           — adds 89 Schedule rows (idempotent: skips if task "ri0" or "g6a" exists)
 *   2. runExtendKenleyBanksMonth1()        — APPENDS Week 2–5 vocab / grammar / reading items to Kenley's Banks.
 *                                          Run BEFORE Kristi takes the Month 1 test. Idempotent.
 *   3. runPatchWeeks3to5Lookback()         — OPTIONAL. (a) Rewords the "As you read, watch for" lines in r3a / r4a / r5a
 *                                          now that Kenley has finished Mysterious Benedict Society. (b) Adds a "name your source"
 *                                          step to the Block 1 argument tasks w3a and w4a. Idempotent.
 *   4. runAppendKenleyBanksMonth2()        — RUN LATER, only after the Month 1 test has been taken (it checks
 *                                          MonthTestMarkers and skips any subject whose test hasn't been taken yet).
 *
 * Banks note: the app samples month tests from bank items AFTER the saved
 * MonthTestMarkers.bank_position (see getNewSinceLastTest in app.js), so Month 2
 * items must not be added until the Month 1 marker has been written.
 *
 * Content notes:
 *   - Reading Week 6 uses the Smithsonian interview with Mark Kurlansky on salt; questions were
 *     verified against the full text of the article (PDF supplied by Kristi).
 *   - "The Plate Check" (task ri0) is the one-page full-sentence answer guide that later reflections point back to.
 *     (Named "Plate Check", NOT "Chef's Notes" — Chef's Notes is already the app's parent-comment feature.)
 *   - Level 1 taggers tag participles as "Adjective" (MCT: a participle is an adjective made from a verb).
 *   - Level 3 phrase taggers from Week 7 include a third option, "Participial".
 *   - Spelling rows for Weeks 6–10 ARE included (15 rows, s6a–s10c) from AAS Level 7 Steps 12–16, plus a Homophone Watch-Out in each focus lesson.
 *   - Tasks ri3, ri4, ri5 carry a new optional content field refLink (salt article link) that the app renders once Claude Code adds support.
 */

// ---------- Run-from-dropdown wrappers (no trailing underscore, so they show in the Apps Script Run list) ----------
function runSeedMonth2Weeks6to10() { seedMonth2Weeks6to10_(); }
function runExtendKenleyBanksMonth1() { extendKenleyBanksMonth1_(); }
function runPatchWeeks3to5Lookback() { patchWeeks3to5Lookback_(); }
function runAppendKenleyBanksMonth2() { appendKenleyBanksMonth2_(); }

function m2Notify_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function seedMonth2Weeks6to10_() {
  var sh = getSheet_('Schedule');
  var existing = sheetToObjects_(sh);
  if (existing.some(function (r) { return r.student === 'kenley' && (r.task_id === 'ri0' || r.task_id === 'g6a'); })) {
    m2Notify_('Weeks 6–10 already appear to be seeded (found task ri0 or g6a) — skipping to avoid duplicates.');
    return;
  }
  var taken = {};
  existing.forEach(function (r) { if (r.student === 'kenley') taken[r.task_id] = true; });
  var clash = MONTH2_WEEKS_6_TO_10.filter(function (t) { return taken[t.id]; }).map(function (t) { return t.id; });
  if (clash.length) {
    m2Notify_('Aborted: these task ids already exist for Kenley: ' + clash.join(', '));
    return;
  }
  var rows = MONTH2_WEEKS_6_TO_10.map(function (t) {
    return ['kenley', t.key, t.name, t.tag, t.week, t.id, t.type, t.label, JSON.stringify(t.content), '', false, false];
  });
  var headers = SHEET_HEADERS.Schedule;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  m2Notify_('Added ' + rows.length + ' new Schedule rows for Kenley, Weeks 6–10.');
}

// ---------- Banks ----------

// Identity of a bank item, to keep appends idempotent.
function bankItemKey_(subjectKey, item) { return subjectKey === 'vocab' ? item.word : item.q; }

function appendToKenleyBank_(subjectKey, newItems) {
  var sh = getSheet_('Banks');
  var headers = getHeaders_(sh);
  var rows = sheetToObjects_(sh);
  var match = rows.find(function (r) { return r.student === 'kenley' && r.subject_key === subjectKey; });
  if (!match) return 'no kenley/' + subjectKey + ' Banks row found — skipped';
  var items = safeParse_(match.items_json, []);
  var have = {};
  items.forEach(function (it) { have[bankItemKey_(subjectKey, it)] = true; });
  var toAdd = newItems.filter(function (it) { return !have[bankItemKey_(subjectKey, it)]; });
  if (!toAdd.length) return subjectKey + ': already up to date (' + items.length + ' items)';
  var merged = items.concat(toAdd);
  var col = headers.indexOf('items_json') + 1;
  sh.getRange(match._row, col).setValue(JSON.stringify(merged));
  return subjectKey + ': ' + items.length + ' → ' + merged.length + ' items';
}

function extendKenleyBanksMonth1_() {
  var out = ['vocab', 'grammar', 'reading'].map(function (k) { return appendToKenleyBank_(k, BANKS_MONTH1[k]); });
  m2Notify_('Month 1 banks extended:\n' + out.join('\n'));
}

function appendKenleyBanksMonth2_() {
  var markers = {};
  sheetToObjects_(getSheet_('MonthTestMarkers'))
    .filter(function (r) { return r.student === 'kenley'; })
    .forEach(function (r) { markers[r.subject_key] = Number(r.bank_position) || 0; });
  var banks = {};
  sheetToObjects_(getSheet_('Banks'))
    .filter(function (r) { return r.student === 'kenley'; })
    .forEach(function (r) { banks[r.subject_key] = safeParse_(r.items_json, []); });
  var out = ['vocab', 'grammar', 'reading'].map(function (k) {
    var len = (banks[k] || []).length;
    // Only safe once the Month 1 test for this subject has been taken (marker caught up to the whole bank).
    if (!markers[k] || markers[k] < len) return k + ': SKIPPED — Month 1 test not taken yet (marker ' + (markers[k] || 0) + ' < bank ' + len + ')';
    return appendToKenleyBank_(k, BANKS_MONTH2[k]);
  });
  m2Notify_('Month 2 banks:\n' + out.join('\n'));
}

// ---------- Optional look-back reword for Weeks 3–5 ----------

function patchWeeks3to5Lookback_() {
  var sh = getSheet_('Schedule');
  var headers = getHeaders_(sh);
  var labelCol = headers.indexOf('label') + 1, contentCol = headers.indexOf('content_json') + 1;
  var rows = sheetToObjects_(sh);
  var done = [];
  Object.keys(LOOKBACK_PATCH).forEach(function (id) {
    var m = rows.find(function (r) { return r.student === 'kenley' && r.task_id === id; });
    if (!m) return;
    sh.getRange(m._row, labelCol).setValue(LOOKBACK_PATCH[id].label);
    sh.getRange(m._row, contentCol).setValue(JSON.stringify(LOOKBACK_PATCH[id].contentObj || { content: LOOKBACK_PATCH[id].content }));
    done.push(id);
  });
  m2Notify_('Look-back wording applied to: ' + (done.join(', ') || '(none found)'));
}

// ---------- Content ----------

var MONTH2_WEEKS_6_TO_10 = [
 {
  "week": 6.0,
  "key": "reading",
  "name": "Reading",
  "tag": "RI Week 1 · Central Idea & Objective Summary",
  "id": "ri0",
  "type": "read",
  "label": "The Plate Check: How To Write A Full-Sentence Answer",
  "content": {
   "content": "<div class=\"lesson-text\"><p>Every answer you type in this app is a real piece of writing. A thoughtful answer in 7th grade is <b>2–4 complete sentences</b> that sound finished when you read them out loud. Before you hit Submit, run the <b>Plate Check</b> — three quick checks, like a chef looking over a plate before it leaves the kitchen:</p><p><b>1. Capitals &amp; periods.</b> Every sentence starts with a capital letter and ends with punctuation. The word <i>I</i> is always a capital. Character names get capitals too — copy them from the book cover or the chapter page so they are spelled right.</p><p><b>2. Complete sentences.</b> Every sentence needs a <b>subject</b> (who or what) and a <b>verb</b> (what they did or are) and has to sound finished. <i>\"Because it is more interesting.\"</i> is a fragment — it leaves you hanging. Try not to start a sentence with <i>and</i>.</p><p><b>3. Prove it.</b> Put the question inside your answer, then give a reason with a detail from the book. That is the difference between a short answer and a thoughtful one.</p></div><table class=\"vocab-table\"><tr><td>Question</td><td>Why might the author hold back what the test is really testing?</td></tr><tr><td>Plate Check fail</td><td>\"it made me guess. and i think they delayed it to make it more mysterius.\"</td></tr><tr><td>Plate Check pass</td><td>\"The author held back what the test was really testing so that readers would guess along with the characters. Not knowing made me pay closer attention to small clues, like how the kids acted after the test.\"</td></tr></table>"
  }
 },
 {
  "week": 6.0,
  "key": "reading",
  "name": "Reading",
  "tag": "RI Week 1 · Central Idea & Objective Summary",
  "id": "ri1",
  "type": "read",
  "label": "This Week's Lens: Central Idea And Objective Summary",
  "content": {
   "content": "<div class=\"lesson-text\"><p>For two weeks you've been reading <i>stories</i>. This week is <b>informational reading</b> — real nonfiction. The skill: find the <b>central idea</b> and write an <b>objective summary</b>.</p><p><b>Central idea</b> = what the whole piece is mostly about <i>plus</i> the main point the author is making about it. <b>Details</b> are the facts, examples, and stories that support that point. A quick test: if you cover up any single detail, does the point still stand? If yes, it was a detail. The central idea needs <i>many</i> details to hold it up.</p><p><b>Objective summary</b> = a short retelling in your own words with the <b>facts only</b>. No <i>I think</i>, no <i>I liked</i>, no opinions. Use 2–3 sentences: one for the central idea, then one or two key details.</p><p><i>Try it on a short paragraph:</i> \"Vanilla is the second most expensive spice in the world. Each orchid flower opens for only one day, so farmers pollinate them by hand. The pods then take months to cure before they can be sold.\" The <b>central idea</b> is <i>why vanilla is so expensive because it takes so much careful hand work</i>. \"Orchid flowers open for one day\" is a <b>detail</b> that supports it. \"I love vanilla ice cream\" would be an <b>opinion</b> — it doesn't belong in an objective summary.</p><p><b>Summary frame:</b> In the article \"____,\" [author] explains that ____. He supports this by ____ and ____.</p><p><b>3-point rubric:</b> (1) states the real main idea, not just one detail · (2) includes one supporting detail · (3) stays objective, no opinions.</p></div>"
  }
 },
 {
  "week": 6.0,
  "key": "reading",
  "name": "Reading",
  "tag": "RI Week 1 · Central Idea & Objective Summary",
  "id": "ri2",
  "type": "external",
  "label": "Read The Article: Mark Kurlansky On Salt",
  "content": {
   "linkText": "Open The Smithsonian Interview",
   "note": "About 900 words — a Smithsonian interview with author Mark Kurlansky about his book on salt. Study this week's five Vocabulary words first (several appear in the article). Read it once straight through, then again with a pencil, marking the sentence in each answer that feels like the big point. The short list of specialty salts at the end can be skimmed.",
   "url": "https://www.smithsonianmag.com/arts-culture/mark-kurlansky-on-the-cultural-importance-of-salt-98044263/"
  }
 },
 {
  "week": 6.0,
  "key": "reading",
  "name": "Reading",
  "tag": "RI Week 1 · Central Idea & Objective Summary",
  "id": "ri3",
  "type": "reflection",
  "label": "Name The Central Idea",
  "content": {
   "prompt": "In one complete sentence, state what this whole interview is mostly about and the main point it makes. Then, in one more sentence, name one detail that is NOT the central idea but helps support it, and tell why it is only a detail.",
   "sampleAnswer": "Parent guide (checked against the article). A strong central-idea sentence names salt as an ordinary food that became economically, politically, and culturally important, and that lost its value once oil drilling and freezing changed things. A fine supporting detail: the salt tax was one of the great grievances behind the French Revolution, or Gandhi's salt march, or salt being the only way to ship food internationally before refrigeration, or salt being sealed into religious covenants. What matters is that she picks something narrower than the whole interview and says why it is narrower. Also look for the Plate Check (Week 6 lesson): capitals, complete sentences, 'I' capitalized.",
   "refLink": {
    "text": "Open The Smithsonian Interview",
    "url": "https://www.smithsonianmag.com/arts-culture/mark-kurlansky-on-the-cultural-importance-of-salt-98044263/"
   }
  }
 },
 {
  "week": 6.0,
  "key": "reading",
  "name": "Reading",
  "tag": "RI Week 1 · Central Idea & Objective Summary",
  "id": "ri4",
  "type": "graded-mc",
  "label": "Central Idea Or Detail?",
  "content": {
   "questions": [
    {
     "q": "Which sentence best states the central idea of the interview?",
     "options": [
      "Gandhi led a famous march to protest the salt tax.",
      "Salt, an ordinary food, has shaped trade, politics, and culture, and its importance changed as technology changed.",
      "Fleur de sel is a specialty salt.",
      "Salt is the best seasoning in the kitchen."
     ],
     "correct": 1,
     "explanation": "It covers the whole piece — salt's importance and how it changed. The others are single details or an opinion."
    },
    {
     "q": "Which of these is a supporting DETAIL, not the central idea?",
     "options": [
      "Salt has mattered economically, politically, and culturally.",
      "Before refrigeration, salt was essential for preserving food that was shipped long distances."
     ],
     "correct": 1,
     "explanation": "The shipping detail is one fact that helps prove the bigger point about salt's importance."
    },
    {
     "q": "According to the interview, why did salt lose its value?",
     "options": [
      "People stopped liking the taste of salt.",
      "Large salt deposits were found during oil exploration, and freezing replaced salt as the main way to preserve food.",
      "The French government banned the sale of salt."
     ],
     "correct": 1,
     "explanation": "Two changes together made salt cheap: a much bigger supply and a new way to preserve food."
    },
    {
     "q": "Which sentence would NOT belong in an objective summary?",
     "options": [
      "Kurlansky explains that salt became more important as societies began farming.",
      "I think salt is the most boring food in the world."
     ],
     "correct": 1,
     "explanation": "\"I think\" and opinion words don't belong in an objective summary. The first sentence is a fact reported from the article."
    },
    {
     "q": "A student writes: \"This article is about the salt tax.\" What is the problem with this central idea?",
     "options": [
      "It names only one detail, not what the whole piece is about.",
      "It is too long.",
      "It includes an opinion.",
      "There is no problem."
     ],
     "correct": 0,
     "explanation": "The salt tax is one detail. A central idea has to cover the whole interview."
    }
   ],
   "refLink": {
    "text": "Open The Smithsonian Interview",
    "url": "https://www.smithsonianmag.com/arts-culture/mark-kurlansky-on-the-cultural-importance-of-salt-98044263/"
   }
  }
 },
 {
  "week": 6.0,
  "key": "reading",
  "name": "Reading",
  "tag": "RI Week 1 · Central Idea & Objective Summary",
  "id": "ri5",
  "type": "reflection",
  "label": "Write The Objective Summary",
  "content": {
   "prompt": "Write a 2–3 sentence objective summary of the interview. Sentence 1: the central idea. Sentences 2–3: one or two key supporting details. No \"I think,\" no opinions, complete sentences. (Use the summary frame from This Week's Lens if it helps.)",
   "sampleAnswer": "3-point rubric: states the actual main idea, not just a detail (1) / includes one supporting detail (1) / stays objective, no opinion words (1). Model: \"In an interview with Smithsonian Magazine, author Mark Kurlansky explains that salt, a common food, became economically, politically, and culturally important because societies needed it for farming and diets, and to preserve food for trade. He notes that the salt tax was one of the great grievances behind the French Revolution, and that salt lost its value once the search for oil uncovered huge salt beds and freezing replaced salt for preserving food.\" Also check the Plate Check (Week 6 lesson): capitals, complete sentences, 'I' capitalized.",
   "refLink": {
    "text": "Open The Smithsonian Interview",
    "url": "https://www.smithsonianmag.com/arts-culture/mark-kurlansky-on-the-cultural-importance-of-salt-98044263/"
   }
  }
 },
 {
  "week": 6.0,
  "key": "vocab",
  "name": "Vocabulary",
  "tag": "Context Clues · Nonfiction Words",
  "id": "v6a",
  "type": "read",
  "label": "Study This Week's Words",
  "content": {
   "content": "<div class=\"lesson-text\"><p>Five words from the world of food and history. Several are in this week's article. Same job as Month 1: find the clue in the sentence that gives away the meaning — a comma-definition, a restating phrase, or a contrast.</p></div><table class=\"vocab-table\"><tr><td>commodity <button class=\"btn word-audio-btn\" onclick=\"speakWord('commodity')\">🔊</button></td><td>\"Salt was once a valuable commodity, a product people bought, sold, and traded like gold.\" (clue: comma-definition)</td></tr><tr><td>preserve <button class=\"btn word-audio-btn\" onclick=\"speakWord('preserve')\">🔊</button></td><td>\"Before refrigerators, cooks used salt to preserve fish, keeping it safe to eat for months instead of days.\" (clue: restating phrase)</td></tr><tr><td>grievance <button class=\"btn word-audio-btn\" onclick=\"speakWord('grievance')\">🔊</button></td><td>\"The salt tax was a major grievance, a complaint so serious that people were willing to protest in the streets.\" (clue: comma-definition)</td></tr><tr><td>ritual <button class=\"btn word-audio-btn\" onclick=\"speakWord('ritual')\">🔊</button></td><td>\"Many cultures made salt part of a ritual, a ceremony repeated in the same careful way every time.\" (clue: comma-definition)</td></tr><tr><td>abundant <button class=\"btn word-audio-btn\" onclick=\"speakWord('abundant')\">🔊</button></td><td>\"Once drilling uncovered huge deposits, salt became abundant — so plentiful that it was no longer rare or costly.\" (clue: restating phrase)</td></tr></table>"
  }
 },
 {
  "week": 6.0,
  "key": "vocab",
  "name": "Vocabulary",
  "tag": "Context Clues · Nonfiction Words",
  "id": "v6b",
  "type": "reflection",
  "label": "Which Clue Told You The Meaning?",
  "content": {
   "prompt": "Pick at least two of this week's words (commodity, preserve, grievance, ritual, abundant). For each one, copy the exact words in the sentence that gave away the meaning, and tell which clue type it was (comma-definition, restating phrase, or contrast). Then write your OWN sentence using the word in a way that shows you know what it means.",
   "sampleAnswer": "Look for: the exact clue words quoted (e.g., 'a product people bought, sold, and traded like gold' for commodity), the right clue type, and an original sentence where the word is used correctly (not just a copy of the example). Plate Check: complete sentences, capitals."
  }
 },
 {
  "week": 6.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · RI Week, Level 4 Booster",
  "id": "g6prep",
  "type": "read",
  "label": "Level 4 Booster: How To Spot A Dependent Clause",
  "content": {
   "content": "<div class=\"lesson-text\"><p>Level 4 is where we sort a sentence into <b>clauses</b>. A clause is a group of words with its own <b>subject and verb</b>. Some can stand alone (<b>independent</b>), and some can't (<b>dependent</b>). Use the <b>4-step Clause Check</b>:</p><p><b>1. Find every verb.</b> Each verb comes with its own subject — each subject + verb pair is its own clause. Two pairs means two clauses.</p><p><b>2. Look for the signal word at the front of a clause.</b> A dependent clause usually starts with a <b>subordinating conjunction</b>: <i>because, although, even though, while, when, since, if, unless, until, before, after, as, whereas</i>.</p><p><b>3. Say the clause alone, out loud.</b> \"Because salt preserved fish for months.\" Does it feel finished? No — you're left waiting. That means it's <b>dependent</b>. The clause that does sound finished is <b>independent</b>.</p><p><b>4. Check the comma.</b> When a dependent clause comes FIRST, a comma follows it. When it comes last, usually no comma is needed.</p></div><table class=\"vocab-table\"><tr><td>Trap to avoid</td><td>Marking a whole sentence \"independent\" because it has a noun and a verb somewhere. A sentence can be a mix — look for the signal word first!</td></tr><tr><td>This week's sentence</td><td>It's about salt, the same food as this week's article. Count the subject + verb pairs before you start.</td></tr></table>"
  }
 },
 {
  "week": 6.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · RI Week, Level 4 Booster",
  "id": "g6a",
  "type": "pos-tagger",
  "label": "Level 1 — Tag The Parts Of Speech",
  "content": {
   "sentence": [
    "Because",
    "salt",
    "preserved",
    "fish",
    "for",
    "months,",
    "traders,",
    "the",
    "earliest",
    "merchants,",
    "sold",
    "it",
    "across",
    "great",
    "distances."
   ],
   "answers": [
    "Conjunction",
    "Noun",
    "Verb",
    "Noun",
    "Preposition",
    "Noun",
    "Noun",
    "Adjective",
    "Adjective",
    "Noun",
    "Verb",
    "Pronoun",
    "Preposition",
    "Adjective",
    "Noun"
   ],
   "options": [
    "Noun",
    "Pronoun",
    "Verb",
    "Adjective",
    "Adverb",
    "Preposition",
    "Conjunction",
    "Interjection"
   ],
   "explanations": [
    "\"Because\" introduces the dependent clause and connects it to the main clause — that's a conjunction's job.",
    "It names a thing — \"salt\" is the subject of the first clause.",
    "It shows the action the subject performed: salt kept the fish from spoiling.",
    "It names a thing — what the salt preserved.",
    "It introduces a phrase telling how long, connecting to its object, \"months.\"",
    "It names a length of time — the object of the preposition \"for.\"",
    "It names people — \"traders\" is the subject of the main clause.",
    "Articles (a, an, the) are treated as adjectives — they modify the noun that follows.",
    "It describes \"merchants\" — \"earliest\" is a superlative adjective (the most early).",
    "It names people — a renamed identity for \"traders.\"",
    "It shows the action the traders performed.",
    "It stands in for \"salt\" so the sentence doesn't repeat the word — that's a pronoun's job.",
    "It introduces a phrase telling where the salt was sold, connecting to its object, \"distances.\"",
    "It describes \"distances\" — adjectives describe nouns.",
    "It names a thing — the object of the preposition \"across.\""
   ]
  }
 },
 {
  "week": 6.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · RI Week, Level 4 Booster",
  "id": "g6b",
  "type": "phrase-tagger",
  "label": "Level 2 — Tag The Parts Of The Sentence",
  "content": {
   "sentence": [
    "Because",
    "salt",
    "preserved",
    "fish",
    "for",
    "months,",
    "traders,",
    "the",
    "earliest",
    "merchants,",
    "sold",
    "it",
    "across",
    "great",
    "distances."
   ],
   "options": [
    "Subject",
    "Action Verb Predicate",
    "Linking Verb Predicate",
    "Direct Object",
    "Indirect Object",
    "Subject Complement"
   ],
   "phrases": [
    {
     "start": 1,
     "end": 1,
     "type": "Subject",
     "explanation": "\"Salt\" is who/what the dependent clause is about."
    },
    {
     "start": 2,
     "end": 2,
     "type": "Action Verb Predicate",
     "explanation": "\"Preserved\" shows a real action, and it takes a direct object — only action verbs can do that."
    },
    {
     "start": 3,
     "end": 3,
     "type": "Direct Object",
     "explanation": "\"Fish\" is what got preserved — the direct object answers \"what.\""
    },
    {
     "start": 6,
     "end": 6,
     "type": "Subject",
     "explanation": "\"Traders\" is who the main clause is about — the ones doing the action."
    },
    {
     "start": 10,
     "end": 10,
     "type": "Action Verb Predicate",
     "explanation": "\"Sold\" shows a real action the traders performed, and it takes a direct object."
    },
    {
     "start": 11,
     "end": 11,
     "type": "Direct Object",
     "explanation": "\"It\" (meaning the salt) is what got sold — the direct object answers \"what.\""
    }
   ]
  }
 },
 {
  "week": 6.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · RI Week, Level 4 Booster",
  "id": "g6c",
  "type": "phrase-tagger",
  "label": "Level 3 — Mark The Phrases",
  "content": {
   "sentence": [
    "Because",
    "salt",
    "preserved",
    "fish",
    "for",
    "months,",
    "traders,",
    "the",
    "earliest",
    "merchants,",
    "sold",
    "it",
    "across",
    "great",
    "distances."
   ],
   "options": [
    "Prepositional",
    "Appositive"
   ],
   "phrases": [
    {
     "start": 4,
     "end": 5,
     "type": "Prepositional",
     "explanation": "It starts with the preposition \"for\" and ends with its object \"months\" — that's the exact shape of a prepositional phrase."
    },
    {
     "start": 7,
     "end": 9,
     "type": "Appositive",
     "explanation": "It renames \"traders\" — telling you who they were, not where or when anything happened, which is what makes it an appositive."
    },
    {
     "start": 12,
     "end": 14,
     "type": "Prepositional",
     "explanation": "It starts with the preposition \"across\" and ends with its object \"distances.\""
    }
   ]
  }
 },
 {
  "week": 6.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · RI Week, Level 4 Booster",
  "id": "g6d",
  "type": "phrase-tagger",
  "label": "Level 4 — Mark The Clauses",
  "content": {
   "sentence": [
    "Because",
    "salt",
    "preserved",
    "fish",
    "for",
    "months,",
    "traders,",
    "the",
    "earliest",
    "merchants,",
    "sold",
    "it",
    "across",
    "great",
    "distances."
   ],
   "options": [
    "Independent Clause",
    "Dependent Clause"
   ],
   "phrases": [
    {
     "start": 0,
     "end": 5,
     "type": "Dependent Clause",
     "explanation": "It starts with the subordinating conjunction \"because\" and can't stand alone as a complete sentence — that makes it dependent."
    },
    {
     "start": 6,
     "end": 14,
     "type": "Independent Clause",
     "explanation": "It has its own subject and predicate (\"traders... sold...\") and could stand alone as a complete sentence — that makes it independent."
    }
   ]
  }
 },
 {
  "week": 6.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · RI Week, Level 4 Booster",
  "id": "g6e",
  "type": "graded-mc",
  "label": "Level 4 — Classify The Sentence Structure",
  "content": {
   "questions": [
    {
     "q": "\"Because salt preserved fish for months, traders, the earliest merchants, sold it across great distances.\" — one independent clause plus one dependent clause makes this sentence:",
     "options": [
      "Simple",
      "Compound",
      "Complex",
      "Compound-Complex"
     ],
     "correct": 2,
     "explanation": "One independent clause + one dependent clause = complex. (Simple = one independent clause alone. Compound = two or more independent clauses joined. Compound-complex = two+ independent clauses plus at least one dependent clause.)"
    },
    {
     "q": "In this sentence, \"across great distances\" tells where the traders sold the salt. What does it modify?",
     "options": [
      "The verb \"sold\" (adverbial)",
      "The noun \"salt\" (adjectival)"
     ],
     "correct": 0,
     "explanation": "It answers \"where did they sell it?\" — that's a job for an adverb, so the phrase is adverbial and modifies the verb \"sold.\""
    }
   ]
  }
 },
 {
  "week": 6.0,
  "key": "writing",
  "name": "Writing",
  "tag": "Book-Tied Analysis · Block 2",
  "id": "w6a",
  "type": "read",
  "label": "Block 2, Lesson 1: What Is An Author's Choice?",
  "content": {
   "content": "<div class=\"lesson-text\"><p>In Block 1 you argued for <i>your own</i> opinion. In Block 2 you flip it: you explain a choice <b>someone else</b> made. An <b>author's choice</b> is a decision about <i>how</i> to tell the story — not what happens, but how it is told. Examples: whose eyes we see through, what the author holds back, what we're shown instead of told.</p><p>Your analysis has three jobs (a <b>4-point rubric</b>): <b>(1)</b> name the choice clearly · <b>(2)</b> tell what we believed at first, what changed, and where · <b>(3)</b> explain <b>WHY</b> the author did it — what it did to you as a reader (<b>2 points</b>: this is the part that matters most).</p><p><b>The evidence rule for this block:</b> your evidence comes from the book, so it has to be something you can <i>point to</i>: an exact quote, a scene, or a chapter. Phrases like \"according to experts\" or \"surveys show\" have no place here. And after every piece of evidence, add a sentence that starts <i>This shows…</i></p><p><b>Model paragraph (a different book):</b> \"In the Harry Potter books, the author chooses to tell the story almost entirely through Harry's point of view, so we mostly know only what Harry knows. For six books, Harry and the reader believe Snape is an enemy because of the cruel things he says and does. In the last book we see Snape's memories and learn he was protecting Harry all along. By keeping us locked inside Harry's limited view, the author lets us feel the same suspicion Harry feels, so the reveal flips our opinion the same way it flips his. If the narrator had simply told us in book one, the surprise would have been lost.\"</p></div><table class=\"vocab-table\"><tr><td>Sentence 1–2</td><td>The choice (POV)</td></tr><tr><td>Sentence 2–3</td><td>What we believed, then what changed, and where</td></tr><tr><td>Sentence 4–5</td><td>WHY: the effect on the reader (suspicion, surprise) and what would be lost if told another way</td></tr></table>"
  }
 },
 {
  "week": 6.0,
  "key": "writing",
  "name": "Writing",
  "tag": "Book-Tied Analysis · Block 2",
  "id": "w6b",
  "type": "reflection",
  "label": "Choose Your Author's Choice And Plan",
  "content": {
   "prompt": "Pick ONE choice from The Mysterious Benedict Society: (A) The author never shows what happens inside the Waiting Room. (B) We stay inside Reynie's point of view the whole book, so Constance's secret reaches us only at the very end. (C) The author doesn't explain what the test is really testing until late. In 4–6 complete sentences, tell me: Which choice did you pick? What did you believe before? What changed, and where in the book? Why do you think the author did it that way? (Rough thinking is fine — this becomes your draft.)",
   "sampleAnswer": "Parent guide: grade this one for planning, not polish. Look for a clear choice, one specific moment in the book, and at least a first guess at 'why.' Best picks are usually (B) or (A) because the effect on the reader is easy to feel. If she is stuck on 'why,' ask: 'How would you have felt if the narrator had just told you?' Then have her write that sentence down. Her Week 2–5 discussion answers are raw material for this. Plate Check (Week 6 lesson) applies. Watch for: a vague 'the author shows…' with no specific moment. That's the same weak spot as 'according to surveys' in her argument paragraphs. If you see it, ask 'Which scene? Which page?'"
  }
 },
 {
  "week": 6.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 12",
  "id": "s6a",
  "type": "read",
  "label": "This Week's Focus",
  "content": {
   "content": "<div class=\"lesson-text\"><p><b>Step 12 — Words Ending in ANCE (Part 1).</b> Words ending in the sound /ents/ can be spelled <u>ance</u> or <u>ence</u>. There are a few clues to help you choose. This week’s clue: <b>if you can start with a whole word and want to add /ents/ to the end, you usually use <u>ance</u>.</b> allow → allow<u>ance</u>, clear → clear<u>ance</u>, resist → resist<u>ance</u>, avoid → avoid<u>ance</u>.</p><p><b>The exceptions.</b> The whole-word clue works most of the time, but not for these eight words, which are spelled <u>ence</u> even though they start with a whole word: <b>conference, correspondence, dependence, difference, existence, insistence, preference, reference</b>. Practice them so they look right to you. When in doubt, look it up.</p><p>Read through the Word Bank for ANCE this week. There are several ways to spell /ents/, and the more of these words you see, the easier it gets to pick the right ending.</p><p><b>Homophone Watch-Out.</b> These pairs sound the same but are spelled differently, so in this week’s dictated sentences the meaning decides the spelling (the same idea as <i>hymn / him</i> in Step 7). Listen for these:</p><p><b>their / there / they’re</b> — <i>their</i> = belonging to them (“brought the guests to <u>their</u> feet”); <i>there</i> = a place, or “there were” (“<u>There</u> were at least a dozen clowns”).</p><p><b>whole / hole</b> — <i>whole</i> = all of it (“I spent my <u>whole</u> allowance”). <b>bear / bare</b> — the animal (“the <u>bear’s</u> resistance”). <b>feet / feat</b> — body parts (“to their <u>feet</u>”). <b>rose / rows</b> — the flower (“the <u>rose</u> garden”).</p></div>"
  }
 },
 {
  "week": 6.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 12",
  "id": "s6b",
  "type": "graded-dictation",
  "label": "Word Cards 101–110 dictation",
  "content": {
   "prompt": "Tap ▶ to hear each word — no text shown, same as Mom reading from the card.",
   "words": [
    {
     "answer": "importance",
     "kind": "word"
    },
    {
     "answer": "performance",
     "kind": "word"
    },
    {
     "answer": "attendance",
     "kind": "word"
    },
    {
     "answer": "allowance",
     "kind": "word"
    },
    {
     "answer": "annoyance",
     "kind": "word"
    },
    {
     "answer": "inheritance",
     "kind": "word"
    },
    {
     "answer": "acceptance",
     "kind": "word"
    },
    {
     "answer": "appearance",
     "kind": "word"
    },
    {
     "answer": "resistance",
     "kind": "word"
    },
    {
     "answer": "avoidance",
     "kind": "word"
    }
   ]
  }
 },
 {
  "week": 6.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 12",
  "id": "s6c",
  "type": "graded-dictation",
  "label": "Dictated sentences",
  "content": {
   "prompt": "All 10 sentences from this week's Reinforcement section.",
   "words": [
    {
     "answer": "The graduates understand the importance of studying.",
     "kind": "sentence"
    },
    {
     "answer": "The musician's performance brought the guests to their feet.",
     "kind": "sentence"
    },
    {
     "answer": "He used his inheritance to construct an animal hospital.",
     "kind": "sentence"
    },
    {
     "answer": "The appearance of the monkeys caused great confusion.",
     "kind": "sentence"
    },
    {
     "answer": "I spent my whole allowance at the rodeo.",
     "kind": "sentence"
    },
    {
     "answer": "The invasion of snails in the rose garden is a real annoyance!",
     "kind": "sentence"
    },
    {
     "answer": "Keith's acceptance address was met with an explosion of giggles.",
     "kind": "sentence"
    },
    {
     "answer": "Her avoidance of the problem is driving me mad!",
     "kind": "sentence"
    },
    {
     "answer": "We were surprised by the bear's resistance to the bathtub.",
     "kind": "sentence"
    },
    {
     "answer": "There were at least a dozen clowns in attendance.",
     "kind": "sentence"
    }
   ]
  }
 },
 {
  "week": 7.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Imagery · Ch. 1–9",
  "id": "r7a",
  "type": "read",
  "label": "This Month's Lens: Figurative Language And Imagery",
  "content": {
   "content": "<div class=\"lesson-text\"><p>Authors paint pictures with words. Two tools do most of the work:</p><p><b>Imagery</b> = words that appeal to the five senses — what a place looks, sounds, smells, feels, or tastes like. <b>Figurative language</b> = words that mean something other than their plain, literal meaning, to make a picture or a feeling. This month's three tools:</p><p>• <b>Simile</b> — compares two unlike things using <i>like</i> or <i>as</i>. \"The sauce was as smooth as velvet.\"<br>• <b>Metaphor</b> — says one thing <i>is</i> another. \"The kitchen was a battlefield during the dinner rush.\"<br>• <b>Personification</b> — gives human qualities to something that isn't human. \"The oven grumbled as it warmed up.\"</p><p><b>Literal:</b> \"My hands were cold.\" <b>Figurative:</b> \"My hands were blocks of ice.\" Same fact — but the second one makes you <i>feel</i> it.</p><p>Naming the device is only half the job. Always ask: <b>(1)</b> What picture do I see? <b>(2)</b> Which exact words made it? <b>(3)</b> What does the comparison make me feel that a plain description wouldn't?</p></div>"
  }
 },
 {
  "week": 7.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Imagery · Ch. 1–9",
  "id": "r7b",
  "type": "read",
  "label": "This Week's Reading: Start Amari And The Night Brothers",
  "content": {
   "content": "<div class=\"lesson-text\"><p><b>Start your Book 3: <i>Amari and the Night Brothers</i> by B.B. Alston.</b> This week, read <b>Chapters 1–9</b> and stop at the end of Chapter 9 (about a quarter of the book).</p><p><b>As you read, watch for:</b></p><p>1. The difference between Amari's ordinary world and the secret world she is invited into. Which exact words make each place feel different — colors, sounds, light, smells, how people talk?</p><p>2. At least one sentence that makes a clear picture appear in your head. Notice the exact words that did it.</p><p>3. Any comparison the author makes — using <i>like</i>, <i>as</i>, or a sentence that says something <i>was</i> something else.</p><p><b>A magic-story warning:</b> in a fantasy book, some things that <i>sound</i> figurative are actually <i>true in the story</i>. If the author says the object really does it, that's the book's magic, so it's literal. Figurative language is when the author is comparing.</p></div>"
  }
 },
 {
  "week": 7.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Imagery · Ch. 1–9",
  "id": "r7c",
  "type": "reflection",
  "label": "Pick A Picture: Imagery",
  "content": {
   "prompt": "Choose ONE sentence from this week's reading that made a strong picture in your mind. Copy it exactly, with quotation marks. Then write 2–3 complete sentences: Which sense does it appeal to (sight, sound, smell, touch, taste)? Which exact words did the work? What feeling does the picture give you?",
   "sampleAnswer": "Parent guide: she should (1) quote exactly, (2) name a sense, (3) point to 2–3 specific words, (4) say a feeling. A weak answer says 'it was descriptive.' A strong answer says something like: 'The words \"flickering\" and \"damp\" make me see and feel a cold, shadowy room, which makes the place feel unsafe.' If she picks a food description, great. Plate Check (Week 6 lesson) applies. Amari-specific: strong picks often come from the contrast between Amari's everyday world (her school and neighborhood) and the Bureau of Supernatural Affairs. Don't choose the sentence for her; if she's stuck, ask 'where did you first feel like you were somewhere totally different?'"
  }
 },
 {
  "week": 7.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Imagery · Ch. 1–9",
  "id": "r7d",
  "type": "reflection",
  "label": "Literal Or Figurative?",
  "content": {
   "prompt": "Find ONE comparison in this week's reading — a simile, a metaphor, or personification. Copy it. Tell which kind it is, what two things are being compared, and what the comparison makes you see or feel that a plain description wouldn't. (If you can't find one, write your own for a place or character in your book, then explain it the same way.)",
   "sampleAnswer": "Parent guide: check that she names the right device (like/as = simile; 'is/was' = metaphor; human action by a non-human = personification) and that she explains the EFFECT, not just the label. An answer that only says 'it is a simile' is a 2-out-of-4 answer; push her to finish the sentence 'This comparison makes me feel ___ because ___.' Amari-specific trap: magic that is real in the story (talking or thinking objects, magical technology) is LITERAL, not personification. If she picks one of those, ask 'Is the author comparing it to something, or is that just what the Bureau's magic does?' A real simile or metaphor compares two unlike things; if she can't find one, accept her own invented one, per the prompt."
  }
 },
 {
  "week": 7.0,
  "key": "vocab",
  "name": "Vocabulary",
  "tag": "Greek & Latin Roots · Month 2",
  "id": "v7a",
  "type": "read",
  "label": "Study This Week's Words",
  "content": {
   "content": "<div class=\"lesson-text\"><p>Month 2 skill: <b>Greek and Latin roots</b>. One root unlocks a whole family of words. This week's root is <b>VERT / VERS</b> — Latin <i>vertere</i>, \"to turn.\" Watch for the root hiding inside each word.</p></div><table class=\"vocab-table\"><tr><td>versatile <button class=\"btn word-audio-btn\" onclick=\"speakWord('versatile')\">🔊</button></td><td>VERS- (turn). Something that can \"turn\" to many uses. Use it: name a kitchen tool or an ingredient that is versatile, and say two ways you could use it.</td></tr><tr><td>convert <button class=\"btn word-audio-btn\" onclick=\"speakWord('convert')\">🔊</button></td><td>CON- (together/completely) + VERT- (turn). To turn something into something else. Use it: explain how you would convert a recipe from cups to grams.</td></tr><tr><td>invert <button class=\"btn word-audio-btn\" onclick=\"speakWord('invert')\">🔊</button></td><td>IN- (in/upside) + VERT- (turn). To turn upside down or inside out. Use it: describe a time in the kitchen when you would invert a pan onto a plate.</td></tr><tr><td>controversy <button class=\"btn word-audio-btn\" onclick=\"speakWord('controversy')\">🔊</button></td><td>CONTRA- (against) + VERS- (turned). Opinions \"turned against\" each other. Use it: name a food controversy you've heard of, or invent a silly one.</td></tr><tr><td>glycemic <button class=\"btn word-audio-btn\" onclick=\"speakWord('glycemic')\">🔊</button></td><td>Greek GLYK- (sweet) + HAIMA (blood). Having to do with blood sugar — \"glycemic index\" shows how fast a food raises it. Use it: say whether you think a white-bread food or a lentil food has a higher glycemic effect, and why.</td></tr></table>"
  }
 },
 {
  "week": 7.0,
  "key": "vocab",
  "name": "Vocabulary",
  "tag": "Greek & Latin Roots · Month 2",
  "id": "v7b",
  "type": "reflection",
  "label": "Root Detective",
  "content": {
   "prompt": "Pick two of this week's words. For each one: name the root (VERT/VERS, or GLYK- + HAIMA), say what it means, and give ONE more word you know (or find in a dictionary) that has the same root. Then write your own sentence that uses the word correctly.",
   "sampleAnswer": "Parent guide: good extra VERT/VERS words — reverse, diverse, vertical, adverse, subvert, universe, introvert, extrovert, advertise, version, traverse. For glycemic, related words: glucose, glycogen, anemia (a-haima), hemoglobin. Look for a correct root meaning, a real family word, and an original sentence showing the meaning (not a copy of the example)."
  }
 },
 {
  "week": 7.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 1 (Participial Phrases)",
  "id": "g7prep",
  "type": "read",
  "label": "Participial Phrases: The Describing Phrase That Starts With A Verb Form",
  "content": {
   "content": "<div class=\"lesson-text\"><p>In MCT's words, a <b>participle</b> is an <b>adjective made from a verb</b>. A <b>participial phrase</b> is a participle plus its helpers, and the whole phrase <b>describes a noun</b>. Today we start with <b>present participles</b> — the <i>-ing</i> kind.</p><p><i>Humming softly, Waffles stirred the soup.</i> \"Humming softly\" is the participial phrase. It describes <b>Waffles</b> — he's the one humming.</p><p><b>The 3-step Participle Test:</b> (1) Find the participle — an <i>-ing</i> word that isn't a main verb. (2) Ask: <i>who or what does it describe?</i> (3) Check the placement — the noun it describes should come <b>right after the comma</b>.</p><p><b>Sentence opener:</b> A participial phrase makes a great way to <i>start</i> a sentence. A comma follows the phrase. Compare: <i>Waffles stirred the soup.</i> vs. <i>Humming softly, Waffles stirred the soup.</i> Same facts, but the second one has more movement and music.</p><p><b>Level 1 tip:</b> tag the participle as an <b>Adjective</b> — it is built from a verb, but it works as an adjective in the sentence. (Compare: <i>The dog is barking</i> — \"is barking\" is the verb; <i>The barking dog woke us</i> — \"barking\" is a participle describing dog; <i>Barking is annoying</i> — a gerund, which acts as a noun. Gerunds come next month.) <b>Level 3 now has a third button: Participial.</b></p></div><table class=\"vocab-table\"><tr><td>Clause check reminder</td><td>Count subject + verb pairs and look for the signal word. This week the dependent clause comes at the END, so there's no comma before it.</td></tr></table>"
  }
 },
 {
  "week": 7.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 1 (Participial Phrases)",
  "id": "g7a",
  "type": "pos-tagger",
  "label": "Level 1 — Tag The Parts Of Speech",
  "content": {
   "sentence": [
    "Humming",
    "softly,",
    "Waffles",
    "stirred",
    "the",
    "soup",
    "while",
    "the",
    "bread",
    "baked",
    "in",
    "the",
    "oven."
   ],
   "answers": [
    "Adjective",
    "Adverb",
    "Noun",
    "Verb",
    "Adjective",
    "Noun",
    "Conjunction",
    "Adjective",
    "Noun",
    "Verb",
    "Preposition",
    "Adjective",
    "Noun"
   ],
   "options": [
    "Noun",
    "Pronoun",
    "Verb",
    "Adjective",
    "Adverb",
    "Preposition",
    "Conjunction",
    "Interjection"
   ],
   "explanations": [
    "\"Humming\" is a participle — an adjective made from a verb — and it describes Waffles.",
    "It tells how Waffles was humming — adverbs describe verbs (and verb forms).",
    "It names a specific character — a proper noun.",
    "It shows the action Waffles performed.",
    "Articles (a, an, the) are treated as adjectives — they modify the noun that follows.",
    "It names a thing — the object of the action.",
    "\"While\" introduces the dependent clause and connects it to the main clause — that's a conjunction's job.",
    "Articles (a, an, the) are treated as adjectives — they modify the noun that follows.",
    "It names a thing — \"bread\" is the subject of the dependent clause.",
    "It shows what the bread did.",
    "It introduces a phrase telling where, connecting to its object, \"oven.\"",
    "Articles (a, an, the) are treated as adjectives — they modify the noun that follows.",
    "It names a place — the object of the preposition \"in.\""
   ]
  }
 },
 {
  "week": 7.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 1 (Participial Phrases)",
  "id": "g7b",
  "type": "phrase-tagger",
  "label": "Level 2 — Tag The Parts Of The Sentence",
  "content": {
   "sentence": [
    "Humming",
    "softly,",
    "Waffles",
    "stirred",
    "the",
    "soup",
    "while",
    "the",
    "bread",
    "baked",
    "in",
    "the",
    "oven."
   ],
   "options": [
    "Subject",
    "Action Verb Predicate",
    "Linking Verb Predicate",
    "Direct Object",
    "Indirect Object",
    "Subject Complement"
   ],
   "phrases": [
    {
     "start": 2,
     "end": 2,
     "type": "Subject",
     "explanation": "\"Waffles\" is who the main clause is about — the one doing the stirring."
    },
    {
     "start": 3,
     "end": 3,
     "type": "Action Verb Predicate",
     "explanation": "\"Stirred\" shows a real action and takes a direct object."
    },
    {
     "start": 5,
     "end": 5,
     "type": "Direct Object",
     "explanation": "\"Soup\" is what got stirred — the direct object answers \"what.\" Just the noun, same as how \"Waffles\" above is the Subject without \"the\" — the article is tagged separately at Level 1."
    },
    {
     "start": 8,
     "end": 8,
     "type": "Subject",
     "explanation": "\"Bread\" is who/what the dependent clause is about."
    },
    {
     "start": 9,
     "end": 9,
     "type": "Action Verb Predicate",
     "explanation": "\"Baked\" shows what the bread did. It's an action verb that doesn't need a direct object."
    }
   ]
  }
 },
 {
  "week": 7.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 1 (Participial Phrases)",
  "id": "g7c",
  "type": "phrase-tagger",
  "label": "Level 3 — Mark The Phrases",
  "content": {
   "sentence": [
    "Humming",
    "softly,",
    "Waffles",
    "stirred",
    "the",
    "soup",
    "while",
    "the",
    "bread",
    "baked",
    "in",
    "the",
    "oven."
   ],
   "options": [
    "Prepositional",
    "Appositive",
    "Participial"
   ],
   "phrases": [
    {
     "start": 0,
     "end": 1,
     "type": "Participial",
     "explanation": "\"Humming\" is a participle (an adjective made from a verb), and the whole phrase describes Waffles — that's a participial phrase."
    },
    {
     "start": 10,
     "end": 12,
     "type": "Prepositional",
     "explanation": "It starts with the preposition \"in\" and ends with its object \"oven.\""
    }
   ]
  }
 },
 {
  "week": 7.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 1 (Participial Phrases)",
  "id": "g7d",
  "type": "phrase-tagger",
  "label": "Level 4 — Mark The Clauses",
  "content": {
   "sentence": [
    "Humming",
    "softly,",
    "Waffles",
    "stirred",
    "the",
    "soup",
    "while",
    "the",
    "bread",
    "baked",
    "in",
    "the",
    "oven."
   ],
   "options": [
    "Independent Clause",
    "Dependent Clause"
   ],
   "phrases": [
    {
     "start": 0,
     "end": 5,
     "type": "Independent Clause",
     "explanation": "It has its own subject and predicate (\"Waffles stirred\") and could stand alone — that makes it independent. Notice the participial phrase rides along inside it."
    },
    {
     "start": 6,
     "end": 12,
     "type": "Dependent Clause",
     "explanation": "It starts with the subordinating conjunction \"while\" and can't stand alone — that makes it dependent. (At the END of a sentence, no comma is needed.)"
    }
   ]
  }
 },
 {
  "week": 7.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 1 (Participial Phrases)",
  "id": "g7e",
  "type": "graded-mc",
  "label": "Level 4 — Classify The Sentence Structure",
  "content": {
   "questions": [
    {
     "q": "\"Humming softly, Waffles stirred the soup while the bread baked in the oven.\" — what is this sentence's structure?",
     "options": [
      "Simple",
      "Compound",
      "Complex",
      "Compound-Complex"
     ],
     "correct": 2,
     "explanation": "One independent clause + one dependent clause = complex. The participial phrase is only a phrase, so it doesn't change the count."
    },
    {
     "q": "Who or what does the participial phrase \"Humming softly\" describe?",
     "options": [
      "Waffles",
      "the soup",
      "the bread"
     ],
     "correct": 0,
     "explanation": "Waffles is the one humming. A participial phrase should sit right next to the noun it describes — here it opens the sentence and Waffles comes right after the comma."
    }
   ]
  }
 },
 {
  "week": 7.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 1 (Participial Phrases)",
  "id": "g7f",
  "type": "reflection",
  "label": "Opener Challenge",
  "content": {
   "prompt": "Write TWO new sentences about something that happens in a kitchen (or while you are cooking). Each one must START with a present participial phrase (an -ing phrase) followed by a comma, and the noun right after the comma must be the one doing the -ing action. Underline your participial phrase in each.",
   "sampleAnswer": "Parent guide: check (1) the phrase starts with an -ing word, (2) a comma follows it, (3) the noun right after the comma is the one doing the action (no dangling). Example good answers: 'Stirring the pot slowly, Kenley watched the soup thicken.' / 'Chopping onions carefully, Dad kept his eyes dry by breathing through his mouth.' Watch for dangling: 'Chopping onions, the knife slipped.'"
  }
 },
 {
  "week": 7.0,
  "key": "writing",
  "name": "Writing",
  "tag": "Book-Tied Analysis · Block 2",
  "id": "w7a",
  "type": "read",
  "label": "Craft Lesson 2: The Hook, Not The Throat-Clearing",
  "content": {
   "content": "<div class=\"lesson-text\"><p>The first sentence of anything you write is a promise to the reader: <i>this is worth your time.</i> Most weak openers are <b>throat-clearing</b> — warming up instead of starting.</p></div><table class=\"vocab-table\"><tr><td>Throat-clearing</td><td>\"In today's world, many people enjoy reading fantasy books.\"</td></tr><tr><td>Hook</td><td>\"Reynie spends almost the whole book not knowing what his test was really testing — and the author wants it that way.\"</td></tr><tr><td>Throat-clearing</td><td>\"Have you ever wondered how bread rises?\"</td></tr><tr><td>Hook</td><td>\"A single teaspoon of yeast can turn a flat lump of dough into a loaf twice its size, and no one ever sees it happen.\"</td></tr><tr><td>Throat-clearing</td><td>\"This essay is about the market.\"</td></tr><tr><td>Hook</td><td>\"The first thing I noticed in the market was the smell of oranges — warm and sharp — long before I saw a single stall.\"</td></tr><tr><td>Three ways to start strong</td><td>1. Start in the middle of the action or the strangest detail. 2. Start with a surprising fact. 3. Start with your sharpest claim.</td></tr></table>"
  }
 },
 {
  "week": 7.0,
  "key": "writing",
  "name": "Writing",
  "tag": "Book-Tied Analysis · Block 2",
  "id": "w7b",
  "type": "reflection",
  "label": "Draft Paragraph 1: The Choice",
  "content": {
   "prompt": "Write Paragraph 1 of your analysis (4–6 complete sentences). Start with a hook, not throat-clearing. Then name the author's choice you picked last week in one clear sentence, and tell what a reader believes at the start. Do NOT write the 'why' yet — that comes in Week 9. Below your paragraph, copy the two other openers you tried and rejected. Then do a proofreading pass: read your paragraph out loud once, circle every word you aren't sure how to spell, look each one up, and fix it. Spell the characters' names exactly as the book does.",
   "sampleAnswer": "Parent guide: rubric points here are mostly for the 'name the choice clearly' job. The hook should be specific to the book (not 'In today's world…'). Having her write 3 openers and pick the best is the real lesson. Proofreading: her recent argument paragraphs had everyday words misspelled (e.g. 'according,' 'beautiful,' 'seasonal,' 'universal'), so this is where the read-aloud-and-look-it-up habit starts. Count spelling lightly (see the grading guide). If her draft jumps to explaining 'why,' praise it and tell her to save that for Paragraph 3. Plate Check applies."
  }
 },
 {
  "week": 7.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 13",
  "id": "s7a",
  "type": "read",
  "label": "This Week's Focus",
  "content": {
   "content": "<div class=\"lesson-text\"><p><b>Step 13 — Words Ending in ANCE (Part 2).</b> You already know clue 1: a whole word plus /ents/ usually takes <u>ance</u> (perform → performance). Here is the <b>second clue</b>: <b>if the last syllable of the word has an <u>a</u> in it, choose <u>ance</u>, which also has an a.</b> radiate → radi<u>ance</u>, instant → inst<u>ance</u>.</p><p>Some words don’t follow either clue. For those you either look them up or build a visual memory of how they should look, so read the Word Bank for ANCE again this week.</p><p><b>Words that change inside first.</b> Some words go through an internal change before the <u>ance</u> is added: abound → abundance, admit → admittance, endure → endurance, enter → entrance, guide → guidance, hinder → hindrance, ignore → ignorance, insure → insurance, maintain → maintenance, observe → observance, persevere → perseverance, remember → remembrance, resemble → resemblance, rid → riddance, sustain → sustenance.</p><p><b>Homophone Watch-Out.</b> These pairs sound the same but are spelled differently, so in this week’s dictated sentences the meaning decides the spelling (the same idea as <i>hymn / him</i> in Step 7). Listen for these:</p><p><b>led / lead</b> — <i>led</i> is the past tense of <i>to lead</i> (“led us right to the kitchen”). <b>right / write / rite</b> — “right to the kitchen” means straight to it.</p><p><b>sea / see</b> (“reach the <u>sea</u>”) · <b>great / grate</b> (“a <u>great</u> distance”) · <b>would / wood</b> (“I <u>would</u> not ask”) · <b>your / you’re</b> (“the color of <u>your</u> eyes”) · <b>by / buy / bye</b> (“damage <u>by</u> bears”) · <b>horse / hoarse</b> (“fell off his <u>horse</u>”).</p></div>"
  }
 },
 {
  "week": 7.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 13",
  "id": "s7b",
  "type": "graded-dictation",
  "label": "Word Cards 111–120 dictation",
  "content": {
   "prompt": "Tap ▶ to hear each word — no text shown, same as Mom reading from the card.",
   "words": [
    {
     "answer": "distance",
     "kind": "word"
    },
    {
     "answer": "fragrance",
     "kind": "word"
    },
    {
     "answer": "reluctance",
     "kind": "word"
    },
    {
     "answer": "tolerance",
     "kind": "word"
    },
    {
     "answer": "significance",
     "kind": "word"
    },
    {
     "answer": "insurance",
     "kind": "word"
    },
    {
     "answer": "resemblance",
     "kind": "word"
    },
    {
     "answer": "guidance",
     "kind": "word"
    },
    {
     "answer": "ambulance",
     "kind": "word"
    },
    {
     "answer": "entrance",
     "kind": "word"
    }
   ]
  }
 },
 {
  "week": 7.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 13",
  "id": "s7c",
  "type": "graded-dictation",
  "label": "Dictated sentences",
  "content": {
   "prompt": "All 10 sentences from this week's Reinforcement section.",
   "words": [
    {
     "answer": "I have a very low tolerance for bad grammar.",
     "kind": "sentence"
    },
    {
     "answer": "You have an incredible resemblance to my cousin Chris.",
     "kind": "sentence"
    },
    {
     "answer": "The entrance to the thief's cave was covered by thorns.",
     "kind": "sentence"
    },
    {
     "answer": "We called the ambulance when Mr. Pile fell off his horse.",
     "kind": "sentence"
    },
    {
     "answer": "The fragrance of cherry pies led us right to the kitchen.",
     "kind": "sentence"
    },
    {
     "answer": "I would not ask for guidance from that guy!",
     "kind": "sentence"
    },
    {
     "answer": "Those tourists had to travel a great distance to reach the sea.",
     "kind": "sentence"
    },
    {
     "answer": "Does the color of your eyes have any real significance?",
     "kind": "sentence"
    },
    {
     "answer": "Our insurance doesn't cover damage by bears.",
     "kind": "sentence"
    },
    {
     "answer": "The spy's reluctance to speak made us fear the worst.",
     "kind": "sentence"
    }
   ]
  }
 },
 {
  "week": 8.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Figurative Language · Ch. 10–16",
  "id": "r8a",
  "type": "read",
  "label": "This Week's Reading: Chapters 10–16",
  "content": {
   "content": "<div class=\"lesson-text\"><p><b>Keep reading <i>Amari and the Night Brothers</i>: Chapters 10–16.</b> Stop at the end of Chapter 16 (about halfway through the book). <b>As you read, watch for:</b></p><p>1. A simile or metaphor used to describe a <i>character</i> (not a place). What does it tell you about that person that a plain description wouldn't?</p><p>2. A moment where a place or object seems to <i>act</i> like a person. Ask yourself: is this <b>personification</b> (the author comparing it to a person), or is it <b>literal magic</b> in the story? How can you tell?</p><p>3. A word choice that makes a scene feel tense, unfair, magical, or exciting. Notice the exact words.</p></div>"
  }
 },
 {
  "week": 8.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Figurative Language · Ch. 10–16",
  "id": "r8b",
  "type": "reflection",
  "label": "Comparing A Character",
  "content": {
   "prompt": "Find a simile or metaphor the author uses to describe a character. Copy it exactly. In 3–4 complete sentences, tell what it suggests about that character that a plain description wouldn't, and whether you think it's a good comparison for that character — why or why not?",
   "sampleAnswer": "Parent guide: look for an exact quote, a correct label (simile vs. metaphor), and an effect — what trait or feeling the comparison adds. Push for the opinion piece ('good comparison because…') — that is the analysis. Amari-specific: good candidates are how the author describes Amari, her roommate Elsie, the Van Helsing twins (Lara and Dylan), or a Bureau official. Spoiler for you only: one of these characters is not what he seems, so a description that sounds friendly now may sound different on a re-read in Week 10. Don't hint at it."
  }
 },
 {
  "week": 8.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Figurative Language · Ch. 10–16",
  "id": "r8c",
  "type": "reflection",
  "label": "Structure: Titles And Repeated Phrases",
  "content": {
   "prompt": "Look at the chapter titles in this part of your book (or, if your book has no titles, a phrase or image the author repeats on purpose). Pick one title or one repeated phrase. What does it promise or point to, and does the chapter deliver on it? Why do you think the author chose it? (3–4 complete sentences.)",
   "sampleAnswer": "Parent guide: this is a first, light step toward RL.7.5 (how structure adds meaning), which the poetry book later in the year will deepen. Accept any thoughtful claim backed by a specific example. If she says 'it tells what's coming,' ask: 'Does it give it away or hide it? Which did the author choose, and why?'"
  }
 },
 {
  "week": 8.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Figurative Language · Ch. 10–16",
  "id": "r8d",
  "type": "reflection",
  "label": "Rouxbe Check-In: Domain Words And Text Structure",
  "content": {
   "prompt": "Open any Rouxbe lesson you've done in the last two weeks. (1) Pick ONE technical or cooking term from it and explain what it means using the clues in the lesson — not a dictionary. (2) Name the text structure of the lesson (sequence/steps, cause-effect, problem-solution, or compare-contrast) and give one example from the lesson that proves it. (3) Why does that structure fit teaching a cooking technique? (4–6 complete sentences total.)",
   "sampleAnswer": "Parent guide: this is the 'light-touch confirm' for RI.7.4 (domain vocabulary in context) and RI.7.5 (text structure) from the Reading doc, so it only needs to be done once in a while. Most Rouxbe lessons are sequence (steps in order) or cause-effect (what changes when you change heat/time). If she names a structure and a real example, mark it confirmed — no grade needed."
  }
 },
 {
  "week": 8.0,
  "key": "vocab",
  "name": "Vocabulary",
  "tag": "Greek & Latin Roots · Month 2",
  "id": "v8a",
  "type": "read",
  "label": "Study This Week's Words",
  "content": {
   "content": "<div class=\"lesson-text\"><p>This week's root is <b>DUC / DUCT</b> — Latin <i>ducere</i>, \"to lead.\" Last month you met it in <i>deduction</i> (\"leading away\" a conclusion from clues).</p></div><table class=\"vocab-table\"><tr><td>conduct <button class=\"btn word-audio-btn\" onclick=\"speakWord('conduct')\">🔊</button></td><td>CON- (together) + DUCT- (lead). To lead or carry out something; also to carry heat or electricity. Use it: explain how a metal pan conducts heat from the stove to the food.</td></tr><tr><td>conducive <button class=\"btn word-audio-btn\" onclick=\"speakWord('conducive')\">🔊</button></td><td>CON- + DUC- (lead). Leading toward a helpful result. Use it: describe one thing that makes a kitchen conducive to careful knife work.</td></tr><tr><td>induce <button class=\"btn word-audio-btn\" onclick=\"speakWord('induce')\">🔊</button></td><td>IN- (in/into) + DUC- (lead). To lead someone or something into doing or being. Use it: write what might induce you to try a new food.</td></tr><tr><td>produce <button class=\"btn word-audio-btn\" onclick=\"speakWord('produce')\">🔊</button></td><td>PRO- (forth) + DUC- (lead). To lead forth, to bring something into being. Use it: name something a farm can produce, and something a kitchen can produce from it.</td></tr><tr><td>contribution <button class=\"btn word-audio-btn\" onclick=\"speakWord('contribution')\">🔊</button></td><td>CON- (together) + TRIBUERE (to give). Something you give to a shared effort or pot. Use it: in Money Lab, say what a monthly contribution to your investment account does over time.</td></tr></table>"
  }
 },
 {
  "week": 8.0,
  "key": "vocab",
  "name": "Vocabulary",
  "tag": "Greek & Latin Roots · Month 2",
  "id": "v8b",
  "type": "reflection",
  "label": "Root Detective",
  "content": {
   "prompt": "Pick two of this week's words. For each one: name the root and what it means, give ONE more word you know with the same root, and write your own sentence that uses the word correctly.",
   "sampleAnswer": "Parent guide: more DUC words — introduce, reduce, educate (e- 'out' + duc, 'lead out'), aqueduct, seduce, deduce, abduct, ductile, viaduct. Look for a correct root meaning, a real family word, and an original sentence with correct use. For 'contribution', the root is TRIBUERE (to give), related to distribute, attribute, tribute."
  }
 },
 {
  "week": 8.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 2 (Past Participles)",
  "id": "g8prep",
  "type": "read",
  "label": "Past Participles: Toasted, Frozen, Burnt, Written",
  "content": {
   "content": "<div class=\"lesson-text\"><p>Last week's participles ended in <i>-ing</i>. This week's are <b>past participles</b> — the form you use after <i>has/have/had</i>: <i>toasted, frozen, burnt, written, hidden</i>. Regular ones end in <i>-ed</i>; irregular ones (<i>frozen, written, hidden</i>) have to be learned.</p><p>A past participle phrase works the same way: <i>Perfectly toasted, the bread vanished.</i> \"Perfectly toasted\" describes the <b>bread</b>.</p><p><b>Tell it apart from a main verb:</b> If the word is <i>the</i> verb of the clause, it's a predicate (<i>The bread was toasted</i>). If it's squeezed in as a describing word set off by a comma, it's a participle (<i>Toasted, the bread…</i>).</p><p><b>Reminder (Level 4):</b> look for the signal word. This week's dependent clause starts with <b>before</b> — a word that is sometimes a preposition and sometimes a subordinating conjunction. If it's followed by a subject and a verb, it's a conjunction and starts a clause.</p><p><b>Placement check:</b> the noun it describes should come right after the comma. Next week we find out what happens when it doesn't!</p></div>"
  }
 },
 {
  "week": 8.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 2 (Past Participles)",
  "id": "g8a",
  "type": "pos-tagger",
  "label": "Level 1 — Tag The Parts Of Speech",
  "content": {
   "sentence": [
    "Perfectly",
    "toasted,",
    "the",
    "bread",
    "vanished",
    "before",
    "the",
    "guests",
    "reached",
    "the",
    "table",
    "in",
    "the",
    "garden."
   ],
   "answers": [
    "Adverb",
    "Adjective",
    "Adjective",
    "Noun",
    "Verb",
    "Conjunction",
    "Adjective",
    "Noun",
    "Verb",
    "Adjective",
    "Noun",
    "Preposition",
    "Adjective",
    "Noun"
   ],
   "options": [
    "Noun",
    "Pronoun",
    "Verb",
    "Adjective",
    "Adverb",
    "Preposition",
    "Conjunction",
    "Interjection"
   ],
   "explanations": [
    "It tells how the bread was toasted — adverbs describe verbs and verb forms.",
    "\"Toasted\" is a past participle — an adjective made from a verb. It describes the bread.",
    "Articles (a, an, the) are treated as adjectives — they modify the noun that follows.",
    "It names a thing — \"bread\" is the subject of the main clause.",
    "It shows what the bread did.",
    "\"Before\" starts a clause here (it has its own subject and verb), so it works as a conjunction, not a preposition.",
    "Articles (a, an, the) are treated as adjectives — they modify the noun that follows.",
    "It names people — \"guests\" is the subject of the dependent clause.",
    "It shows the action the guests performed.",
    "Articles (a, an, the) are treated as adjectives — they modify the noun that follows.",
    "It names a thing — what the guests reached.",
    "It introduces a phrase telling where the table was, connecting to its object, \"garden.\"",
    "Articles (a, an, the) are treated as adjectives — they modify the noun that follows.",
    "It names a place — the object of the preposition \"in.\""
   ]
  }
 },
 {
  "week": 8.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 2 (Past Participles)",
  "id": "g8b",
  "type": "phrase-tagger",
  "label": "Level 2 — Tag The Parts Of The Sentence",
  "content": {
   "sentence": [
    "Perfectly",
    "toasted,",
    "the",
    "bread",
    "vanished",
    "before",
    "the",
    "guests",
    "reached",
    "the",
    "table",
    "in",
    "the",
    "garden."
   ],
   "options": [
    "Subject",
    "Action Verb Predicate",
    "Linking Verb Predicate",
    "Direct Object",
    "Indirect Object",
    "Subject Complement"
   ],
   "phrases": [
    {
     "start": 3,
     "end": 3,
     "type": "Subject",
     "explanation": "\"Bread\" is who/what the main clause is about."
    },
    {
     "start": 4,
     "end": 4,
     "type": "Action Verb Predicate",
     "explanation": "\"Vanished\" tells what the bread did. It's an action verb that doesn't take a direct object."
    },
    {
     "start": 7,
     "end": 7,
     "type": "Subject",
     "explanation": "\"Guests\" is who the dependent clause is about."
    },
    {
     "start": 8,
     "end": 8,
     "type": "Action Verb Predicate",
     "explanation": "\"Reached\" shows a real action and takes a direct object."
    },
    {
     "start": 10,
     "end": 10,
     "type": "Direct Object",
     "explanation": "\"Table\" is what got reached — the direct object answers \"what.\" Just the noun, same as how \"guests\" above is the Subject without \"the\" — the article is tagged separately at Level 1."
    }
   ]
  }
 },
 {
  "week": 8.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 2 (Past Participles)",
  "id": "g8c",
  "type": "phrase-tagger",
  "label": "Level 3 — Mark The Phrases",
  "content": {
   "sentence": [
    "Perfectly",
    "toasted,",
    "the",
    "bread",
    "vanished",
    "before",
    "the",
    "guests",
    "reached",
    "the",
    "table",
    "in",
    "the",
    "garden."
   ],
   "options": [
    "Prepositional",
    "Appositive",
    "Participial"
   ],
   "phrases": [
    {
     "start": 0,
     "end": 1,
     "type": "Participial",
     "explanation": "\"Toasted\" is a past participle acting like an adjective, and the phrase describes the bread — that's a participial phrase."
    },
    {
     "start": 11,
     "end": 13,
     "type": "Prepositional",
     "explanation": "It starts with the preposition \"in\" and ends with its object \"garden.\""
    }
   ]
  }
 },
 {
  "week": 8.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 2 (Past Participles)",
  "id": "g8d",
  "type": "phrase-tagger",
  "label": "Level 4 — Mark The Clauses",
  "content": {
   "sentence": [
    "Perfectly",
    "toasted,",
    "the",
    "bread",
    "vanished",
    "before",
    "the",
    "guests",
    "reached",
    "the",
    "table",
    "in",
    "the",
    "garden."
   ],
   "options": [
    "Independent Clause",
    "Dependent Clause"
   ],
   "phrases": [
    {
     "start": 0,
     "end": 4,
     "type": "Independent Clause",
     "explanation": "It has its own subject and predicate (\"bread vanished\") and could stand alone — that makes it independent."
    },
    {
     "start": 5,
     "end": 13,
     "type": "Dependent Clause",
     "explanation": "It starts with the subordinating conjunction \"before\" and can't stand alone — that makes it dependent."
    }
   ]
  }
 },
 {
  "week": 8.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 2 (Past Participles)",
  "id": "g8e",
  "type": "graded-mc",
  "label": "Level 4 — Classify The Sentence Structure",
  "content": {
   "questions": [
    {
     "q": "\"Perfectly toasted, the bread vanished before the guests reached the table in the garden.\" — what is this sentence's structure?",
     "options": [
      "Simple",
      "Compound",
      "Complex",
      "Compound-Complex"
     ],
     "correct": 2,
     "explanation": "One independent clause + one dependent clause = complex."
    },
    {
     "q": "Which word does the participial phrase \"Perfectly toasted\" describe?",
     "options": [
      "bread",
      "guests",
      "table"
     ],
     "correct": 0,
     "explanation": "The bread is what got toasted, and \"bread\" comes right after the comma — the phrase is placed correctly."
    }
   ]
  }
 },
 {
  "week": 8.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 2 (Past Participles)",
  "id": "g8f",
  "type": "reflection",
  "label": "Opener Challenge",
  "content": {
   "prompt": "Write TWO new sentences that START with a past participial phrase (like \"Perfectly toasted,\" or \"Frozen solid,\") followed by a comma. The noun right after the comma must be the thing being described. Underline the participial phrase in each.",
   "sampleAnswer": "Parent guide: check for a past participle (-ed, -en, or irregular), a comma, and the right noun right after it. Good examples: 'Dusted with flour, the loaf waited on the counter.' / 'Frozen solid, the berries rattled in the bowl.' Beware of 'Covered in sauce, the pasta…' (fine) vs. 'Covered in sauce, the spoon…' (could be fine too — talk about it)."
  }
 },
 {
  "week": 8.0,
  "key": "writing",
  "name": "Writing",
  "tag": "Book-Tied Analysis · Block 2",
  "id": "w8a",
  "type": "read",
  "label": "Quoting Evidence Smoothly",
  "content": {
   "content": "<div class=\"lesson-text\"><p>Strong analysis <b>shows</b> the reader the exact words from the book. A good quote sandwich has three layers:</p><p><b>1. Lead in:</b> tell where you are in the story. (<i>Early in the book…</i> / <i>In Chapter 12…</i>)<br><b>2. The quote:</b> copy the exact words in quotation marks. Keep it short — just the part that proves your point.<br><b>3. Explain:</b> say what the quote shows and how it connects to your idea. Never leave a quote standing by itself.</p></div><table class=\"vocab-table\"><tr><td>Example (made-up book)</td><td>Early in the novel, the author describes the kitchen as a place where \"the air smelled like a thunderstorm made of cinnamon.\" This shows that the author wants the room to feel exciting and a little dangerous, even though it is full of baking.</td></tr><tr><td>Not enough</td><td>The author describes the kitchen as exciting and dangerous.</td></tr><tr><td>Not evidence</td><td>Many readers agree the kitchen feels exciting.</td></tr><tr><td>The rule</td><td>If you can't point to it in the book, it isn't evidence.</td></tr></table>"
  }
 },
 {
  "week": 8.0,
  "key": "writing",
  "name": "Writing",
  "tag": "Book-Tied Analysis · Block 2",
  "id": "w8b",
  "type": "reflection",
  "label": "Draft Paragraph 2: What Changed",
  "content": {
   "prompt": "Write Paragraph 2 (4–6 complete sentences). What did a reader (and Reynie) believe at first about the choice you picked? What changed, and WHERE in the book? Use at least one exact quote or specific detail from the book, with a lead-in and an explanation (the quote sandwich). Put the chapter or scene in parentheses after each piece of evidence, and end each with a sentence that starts 'This shows…'. Still no 'why' yet.",
   "sampleAnswer": "Parent guide: the rubric point for this paragraph is 'accurately states what changed and when.' Specifics beat generalities — a chapter number or a recognizable scene. If she can't find a quote, it's OK to use a specific detail she remembers; encourage her to flip back to check. If she writes 'the author shows…' with no quote or scene, ask 'Which page?' If she can't point to it, it isn't evidence yet. Plate Check applies."
  }
 },
 {
  "week": 8.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 14",
  "id": "s8a",
  "type": "read",
  "label": "This Week's Focus",
  "content": {
   "content": "<div class=\"lesson-text\"><p><b>Step 14 — Words Ending in ENCE.</b> This week we switch from <u>ance</u> to <u>ence</u>, and there are two clues.</p><p><b>Clue 1: a soft g or soft c before the /ents/ sound means <u>ence</u>.</b> Say <i>diligence</i>: the g is soft because it is followed by an e. If we used <u>ance</u> the g would turn hard (“dili-gance”). The same goes for <i>innocence</i>: the c is soft, so we spell it <u>ence</u>.</p><p><b>Clue 2: if the last syllable has an <u>e</u> in it, choose <u>ence</u>, which also has an e.</b> compete → compet<u>ence</u>, silent → sil<u>ence</u>.</p><p>These clues work most of the time, not every time. Read through the Word Bank for ENCE so these words look familiar, and remember the Step 12 exceptions (<i>reference, conference, difference</i>) show up in this week’s cards too.</p><p><b>Homophone Watch-Out.</b> These pairs sound the same but are spelled differently, so in this week’s dictated sentences the meaning decides the spelling (the same idea as <i>hymn / him</i> in Step 7). Listen for these:</p><p><b>no / know</b> — “You have <u>no</u> evidence” means none; <i>know</i> is to be aware. <b>one / won</b> — “at least <u>one</u> reference” means a single one; <i>won</i> is the past tense of <i>win</i>.</p><p><b>your / you’re</b> (“<u>Your</u> absence”) · <b>there / their / they’re</b> (“<u>There</u> was nothing she could do”) · <b>night / knight</b> (“the silence of the <u>night</u>”) · <b>plans / planes</b> (“our <u>plans</u>”).</p></div>"
  }
 },
 {
  "week": 8.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 14",
  "id": "s8b",
  "type": "graded-dictation",
  "label": "Word Cards 121–130 dictation",
  "content": {
   "prompt": "Tap ▶ to hear each word — no text shown, same as Mom reading from the card.",
   "words": [
    {
     "answer": "innocence",
     "kind": "word"
    },
    {
     "answer": "confidence",
     "kind": "word"
    },
    {
     "answer": "absence",
     "kind": "word"
    },
    {
     "answer": "reference",
     "kind": "word"
    },
    {
     "answer": "silence",
     "kind": "word"
    },
    {
     "answer": "conference",
     "kind": "word"
    },
    {
     "answer": "intelligence",
     "kind": "word"
    },
    {
     "answer": "experience",
     "kind": "word"
    },
    {
     "answer": "difference",
     "kind": "word"
    },
    {
     "answer": "evidence",
     "kind": "word"
    }
   ]
  }
 },
 {
  "week": 8.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 14",
  "id": "s8c",
  "type": "graded-dictation",
  "label": "Dictated sentences",
  "content": {
   "prompt": "All 10 sentences from this week's Reinforcement section.",
   "words": [
    {
     "answer": "You have no evidence that I painted those gnomes pink!",
     "kind": "sentence"
    },
    {
     "answer": "The president held a private conference yesterday.",
     "kind": "sentence"
    },
    {
     "answer": "Dave can't tell the difference between sugar and salt.",
     "kind": "sentence"
    },
    {
     "answer": "That lamb has the intelligence of a fully grown sheep.",
     "kind": "sentence"
    },
    {
     "answer": "Your absence at the table interfered with our plans.",
     "kind": "sentence"
    },
    {
     "answer": "In my experience, books are better than television shows.",
     "kind": "sentence"
    },
    {
     "answer": "Tom speaks three languages with confidence.",
     "kind": "sentence"
    },
    {
     "answer": "There was nothing she could do to prove her innocence.",
     "kind": "sentence"
    },
    {
     "answer": "Include at least one reference in your report.",
     "kind": "sentence"
    },
    {
     "answer": "Snow fell in the silence of the night.",
     "kind": "sentence"
    }
   ]
  }
 },
 {
  "week": 9.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Mood & Word Choice · Ch. 17–24",
  "id": "r9a",
  "type": "read",
  "label": "This Week's Reading: Chapters 17–24",
  "content": {
   "content": "<div class=\"lesson-text\"><p><b>Keep reading <i>Amari and the Night Brothers</i>: Chapters 17–24.</b> Stop at the end of Chapter 24 (about three-quarters of the book). <b>As you read, watch for:</b></p><p>1. <b>Mood.</b> Pick a scene that feels tense, spooky, magical, or unfair. Which exact words make it feel that way?</p><p>2. <b>Connotation.</b> Two words can mean almost the same thing but feel different — <i>walked</i> vs. <i>crept</i>, <i>house</i> vs. <i>mansion</i>. Notice when the author chose the more powerful one.</p><p>3. <b>Repeats.</b> Has the author described the same object, color, or place in a similar way more than once? Notice what it seems to stand for.</p></div>"
  }
 },
 {
  "week": 9.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Mood & Word Choice · Ch. 17–24",
  "id": "r9b",
  "type": "reflection",
  "label": "Word Choice And Mood",
  "content": {
   "prompt": "Pick a passage (2–4 sentences) that creates a clear mood. Copy the best sentence. Name the mood, list THREE exact words from the passage that create it, and then swap one of those words for a plainer word. How does the feeling change? (4–5 complete sentences.)",
   "sampleAnswer": "Parent guide: this is connotation in action (Month 5 will formalize it). A strong answer names a precise mood (not 'good' or 'bad'), lists three actual words from the page, and makes the swap sound different: 'crept' vs. 'walked' feels sneaky vs. neutral. If she struggles, offer a pair like 'whispered/said' to practice the swap out loud first."
  }
 },
 {
  "week": 9.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Mood & Word Choice · Ch. 17–24",
  "id": "r9c",
  "type": "reflection",
  "label": "A Picture That Keeps Coming Back",
  "content": {
   "prompt": "Has the author described the same kind of image, object, color, or place more than once? Describe what keeps coming back, find at least TWO places it shows up, and tell what feeling or idea it seems to carry each time. If you can't find one, pick the image you remember most from the book so far and tell why it stuck with you.",
   "sampleAnswer": "Parent guide: this is an early seed for the Symbolism & Theme lens (Book 5). Don't introduce the word 'symbol' unless she uses it; if she notices a repeated image, just ask what it makes her think the book is 'really about.' Two specific examples beat one."
  }
 },
 {
  "week": 9.0,
  "key": "vocab",
  "name": "Vocabulary",
  "tag": "Greek & Latin Roots · Month 2",
  "id": "v9a",
  "type": "read",
  "label": "Study This Week's Words",
  "content": {
   "content": "<div class=\"lesson-text\"><p>This week's root is <b>FUS / FUND</b> — Latin <i>fundere</i>, \"to pour.\" Many of these words are about things that flow, spread, or mix — just like liquids in a kitchen.</p></div><table class=\"vocab-table\"><tr><td>infuse <button class=\"btn word-audio-btn\" onclick=\"speakWord('infuse')\">🔊</button></td><td>IN- (in) + FUS- (pour). To pour a quality into something, or to soak flavor into liquid. Use it: describe how you would infuse oil with garlic or tea with mint.</td></tr><tr><td>diffuse <button class=\"btn word-audio-btn\" onclick=\"speakWord('diffuse')\">🔊</button></td><td>DIS- (apart) + FUS- (pour). To pour out and spread in all directions. Use it: describe how a smell diffuses from the oven through a house.</td></tr><tr><td>profuse <button class=\"btn word-audio-btn\" onclick=\"speakWord('profuse')\">🔊</button></td><td>PRO- (forth) + FUS- (pour). Poured forth in great amounts. Use it: write about a time you gave or got profuse thanks, or profuse apologies.</td></tr><tr><td>effusive <button class=\"btn word-audio-btn\" onclick=\"speakWord('effusive')\">🔊</button></td><td>EX- (out) + FUS- (pour). Gushing — feelings that pour out freely. Use it: describe someone who is effusive when they taste something delicious.</td></tr><tr><td>confuse <button class=\"btn word-audio-btn\" onclick=\"speakWord('confuse')\">🔊</button></td><td>CON- (together) + FUND- (pour). To pour things together until they're mixed up. Use it: say two kitchen terms (or two spices) that people often confuse.</td></tr></table>"
  }
 },
 {
  "week": 9.0,
  "key": "vocab",
  "name": "Vocabulary",
  "tag": "Greek & Latin Roots · Month 2",
  "id": "v9b",
  "type": "reflection",
  "label": "Root Detective",
  "content": {
   "prompt": "Pick two of this week's words. For each one: name the root and what it means, give ONE more word you know with the same root (or find one), and write your own sentence that uses the word correctly. Bonus: explain how the word 'fondue' fits this family.",
   "sampleAnswer": "Parent guide: more FUS/FUND words — fuse, fusion, transfusion, refund (re- 'back' + fund 'pour back'), foundry, funnel (from the same pouring root, via Latin fundibulum), fondue (French, 'melted' — from fondre, 'to melt/pour'). Look for a correct root meaning, a real family word, and an original sentence with correct use. The fondue bonus is optional fun."
  }
 },
 {
  "week": 9.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 3 (Dangling Modifiers)",
  "id": "g9prep",
  "type": "read",
  "label": "Dangling Modifiers: Who Is Doing The Describing?",
  "content": {
   "content": "<div class=\"lesson-text\"><p>A participial phrase must <b>sit right next to the noun it describes</b> — and that noun has to be the one actually doing the action. When it's not, you get a <b>dangling modifier</b>, and the sentence says something silly.</p><p><b>Dangling:</b> <i>Running to catch the train, my coffee spilled.</i> The coffee can't run!<br><b>Fixed:</b> <i>Running to catch the train, I spilled my coffee.</i></p><p><b>Another one:</b> <i>Whisking the batter, the timer rang loudly.</i> A timer can't whisk. <b>Fix:</b> put the real whisker right after the comma: <i>Whisking the batter, Waffles heard the timer ring loudly.</i></p><p><b>The quick check:</b> read the participial phrase, put \"who is doing that?\" next to it, then look at the word right after the comma. If they don't match, you've found a dangler.</p><p><b>Why this matters:</b> dangling modifiers sneak into real writing all the time — including the Block 2 paragraph you're drafting. Read your openers out loud and ask who's doing the -ing.</p></div>"
  }
 },
 {
  "week": 9.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 3 (Dangling Modifiers)",
  "id": "g9a",
  "type": "pos-tagger",
  "label": "Level 1 — Tag The Parts Of Speech",
  "content": {
   "sentence": [
    "Tiptoeing",
    "quietly,",
    "Waffles",
    "checked",
    "the",
    "stove",
    "because",
    "the",
    "caramel",
    "in",
    "the",
    "pan",
    "smelled",
    "burnt."
   ],
   "answers": [
    "Adjective",
    "Adverb",
    "Noun",
    "Verb",
    "Adjective",
    "Noun",
    "Conjunction",
    "Adjective",
    "Noun",
    "Preposition",
    "Adjective",
    "Noun",
    "Verb",
    "Adjective"
   ],
   "options": [
    "Noun",
    "Pronoun",
    "Verb",
    "Adjective",
    "Adverb",
    "Preposition",
    "Conjunction",
    "Interjection"
   ],
   "explanations": [
    "\"Tiptoeing\" is a participle — an adjective made from a verb — and it describes Waffles.",
    "It tells how Waffles was tiptoeing.",
    "It names a specific character — a proper noun.",
    "It shows the action Waffles performed.",
    "Articles (a, an, the) are treated as adjectives — they modify the noun that follows.",
    "It names a thing — the object of the action.",
    "\"Because\" introduces the dependent clause and connects it to the main clause.",
    "Articles (a, an, the) are treated as adjectives — they modify the noun that follows.",
    "It names a thing — \"caramel\" is the subject of the dependent clause.",
    "It introduces a phrase telling where the caramel was, connecting to its object, \"pan.\"",
    "Articles (a, an, the) are treated as adjectives — they modify the noun that follows.",
    "It names a thing — the object of the preposition \"in.\"",
    "It links \"caramel\" to a description. We'll classify it as action or linking in Level 2.",
    "It describes the caramel — adjectives describe nouns."
   ]
  }
 },
 {
  "week": 9.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 3 (Dangling Modifiers)",
  "id": "g9b",
  "type": "phrase-tagger",
  "label": "Level 2 — Tag The Parts Of The Sentence",
  "content": {
   "sentence": [
    "Tiptoeing",
    "quietly,",
    "Waffles",
    "checked",
    "the",
    "stove",
    "because",
    "the",
    "caramel",
    "in",
    "the",
    "pan",
    "smelled",
    "burnt."
   ],
   "options": [
    "Subject",
    "Action Verb Predicate",
    "Linking Verb Predicate",
    "Direct Object",
    "Indirect Object",
    "Subject Complement"
   ],
   "phrases": [
    {
     "start": 2,
     "end": 2,
     "type": "Subject",
     "explanation": "\"Waffles\" is who the main clause is about."
    },
    {
     "start": 3,
     "end": 3,
     "type": "Action Verb Predicate",
     "explanation": "\"Checked\" shows a real action and takes a direct object."
    },
    {
     "start": 5,
     "end": 5,
     "type": "Direct Object",
     "explanation": "\"Stove\" is what got checked — the direct object answers \"what.\" Just the noun, same as how \"Waffles\" above is the Subject without \"the\" — the article is tagged separately at Level 1."
    },
    {
     "start": 8,
     "end": 8,
     "type": "Subject",
     "explanation": "\"Caramel\" is who/what the dependent clause is about."
    },
    {
     "start": 12,
     "end": 12,
     "type": "Linking Verb Predicate",
     "explanation": "Here \"smelled\" doesn't show an action — it links \"caramel\" to a description. (Compare: \"Waffles smelled the soup\" is an action.)"
    },
    {
     "start": 13,
     "end": 13,
     "type": "Subject Complement",
     "explanation": "\"Burnt\" describes the subject \"caramel\" after the linking verb — that's exactly what a subject complement does."
    }
   ]
  }
 },
 {
  "week": 9.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 3 (Dangling Modifiers)",
  "id": "g9c",
  "type": "phrase-tagger",
  "label": "Level 3 — Mark The Phrases",
  "content": {
   "sentence": [
    "Tiptoeing",
    "quietly,",
    "Waffles",
    "checked",
    "the",
    "stove",
    "because",
    "the",
    "caramel",
    "in",
    "the",
    "pan",
    "smelled",
    "burnt."
   ],
   "options": [
    "Prepositional",
    "Appositive",
    "Participial"
   ],
   "phrases": [
    {
     "start": 0,
     "end": 1,
     "type": "Participial",
     "explanation": "\"Tiptoeing\" is a participle (an adjective made from a verb), and the phrase describes Waffles — a participial phrase."
    },
    {
     "start": 9,
     "end": 11,
     "type": "Prepositional",
     "explanation": "It starts with the preposition \"in\" and ends with its object \"pan.\""
    }
   ]
  }
 },
 {
  "week": 9.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 3 (Dangling Modifiers)",
  "id": "g9d",
  "type": "phrase-tagger",
  "label": "Level 4 — Mark The Clauses",
  "content": {
   "sentence": [
    "Tiptoeing",
    "quietly,",
    "Waffles",
    "checked",
    "the",
    "stove",
    "because",
    "the",
    "caramel",
    "in",
    "the",
    "pan",
    "smelled",
    "burnt."
   ],
   "options": [
    "Independent Clause",
    "Dependent Clause"
   ],
   "phrases": [
    {
     "start": 0,
     "end": 5,
     "type": "Independent Clause",
     "explanation": "It has its own subject and predicate (\"Waffles checked\") and could stand alone — that makes it independent."
    },
    {
     "start": 6,
     "end": 13,
     "type": "Dependent Clause",
     "explanation": "It starts with the subordinating conjunction \"because\" and can't stand alone — that makes it dependent."
    }
   ]
  }
 },
 {
  "week": 9.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 3 (Dangling Modifiers)",
  "id": "g9e",
  "type": "graded-mc",
  "label": "Level 4 — Classify The Sentence Structure And Fix The Dangler",
  "content": {
   "questions": [
    {
     "q": "\"Tiptoeing quietly, Waffles checked the stove because the caramel in the pan smelled burnt.\" — what is this sentence's structure?",
     "options": [
      "Simple",
      "Compound",
      "Complex",
      "Compound-Complex"
     ],
     "correct": 2,
     "explanation": "One independent clause + one dependent clause = complex."
    },
    {
     "q": "Which sentence has a dangling modifier?",
     "options": [
      "\"Running to catch the train, my coffee spilled.\"",
      "\"Running to catch the train, I spilled my coffee.\""
     ],
     "correct": 0,
     "explanation": "Coffee can't run. The describing phrase doesn't sit next to the noun that is actually doing the running. In the second sentence, \"I\" is the runner, so the phrase is placed correctly."
    },
    {
     "q": "Which revision fixes this sentence? \"Whisking the batter, the timer rang loudly.\"",
     "options": [
      "\"Whisking the batter, Waffles heard the timer ring loudly.\"",
      "\"Whisking the batter, the timer rang loudly.\" (no change needed)",
      "\"The timer, whisking the batter, rang loudly.\""
     ],
     "correct": 0,
     "explanation": "A timer can't whisk. The fix puts the real whisker (Waffles) right after the comma. Moving the phrase next to \"timer\" (the third option) just makes the timer do the whisking in a different spot."
    }
   ]
  }
 },
 {
  "week": 9.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 3 (Dangling Modifiers)",
  "id": "g9f",
  "type": "reflection",
  "label": "Opener Challenge",
  "content": {
   "prompt": "Write ONE sentence with a dangling modifier on purpose (make it silly!). Then rewrite it so the participial phrase is correct. Show both versions, and underline the participial phrase in each.",
   "sampleAnswer": "Parent guide: check that the first version really dangles (the noun after the comma can't do the action) and that the fix puts the real doer right after the comma. Example: 'Sizzling in the pan, I smelled the garlic.' (the speaker isn't sizzling) → 'Sizzling in the pan, the garlic smelled wonderful.' Silly examples are the best way to remember this."
  }
 },
 {
  "week": 9.0,
  "key": "writing",
  "name": "Writing",
  "tag": "Book-Tied Analysis · Block 2",
  "id": "w9a",
  "type": "reflection",
  "label": "Draft Paragraph 3: Why Did The Author Do It?",
  "content": {
   "prompt": "Write Paragraph 3 (5–7 complete sentences). This is the most important paragraph. WHY did the author make this choice? Think about what it did to you as a reader (curious? tricked? tense? surprised?) and what the story would be like if the author had just told us instead. Use at least two 'because' sentences. End with one last sentence that adds something — not just a repeat of your first paragraph (a last sentence that only says your claim again doesn't add anything). If you get stuck, use these frames: 'Because the author ___, the reader ___.' and 'If the author had ___ instead, we would have ___.'",
   "sampleAnswer": "Parent guide: this paragraph carries 2 of the 4 rubric points. A strong answer connects the author's choice to a specific effect ('it made me guess along with the characters') and tests it ('if the narrator had told us, I would have stopped paying attention to small clues'). Weak: 'The author did this to make it more interesting' with no how/why. If she stalls, ask: 'What would you have missed?' Plate Check applies."
  }
 },
 {
  "week": 9.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 15",
  "id": "s9a",
  "type": "read",
  "label": "This Week's Focus",
  "content": {
   "content": "<div class=\"lesson-text\"><p><b>Step 15 — Words Ending in IZE.</b> <u>Ize</u> is a suffix meaning <i>to cause to be</i>, and we can add it to hundreds of words. In many cases you simply spell the base word and add the suffix: victim → victim<u>ize</u>, special → special<u>ize</u>, general → general<u>ize</u>.</p><p><b>Silent E.</b> If the base word ends in a silent e, drop the e before adding <u>ize</u>, just like with every vowel suffix: fertile → fertil<u>ize</u>.</p><p><b>Connectors.</b> Some words need connector letters before the <u>ize</u> to make the word easier to pronounce. In <i>systematize</i>, the connector <u>a-t</u> is added to the base word <i>system</i>, and then comes <u>ize</u>.</p><p><b>Organize.</b> <i>Organ</i> comes from a Greek word meaning <i>tool</i>, so if something is <i>organized</i>, it is like a tool. For each word on the cards, name the base word first, and then spell it.</p><p><b>Homophone Watch-Out.</b> These pairs sound the same but are spelled differently, so in this week’s dictated sentences the meaning decides the spelling (the same idea as <i>hymn / him</i> in Step 7). Listen for these:</p><p><b>knew / new</b> — “The doctor <u>knew</u> he had to…”: <i>knew</i> is the past tense of <i>know</i>; <i>new</i> means not old. <b>for / four / fore</b> — “apologize <u>for</u> eating” is the little word, not the number.</p><p><b>to / too / two</b> (“make music materialize”; “had <u>to</u> hospitalize”) · <b>rain / reign / rein</b> (“<u>rain</u> coats” means coats for wet weather) · <b>seems / seams</b> (“He <u>seems</u> to make music…”).</p></div>"
  }
 },
 {
  "week": 9.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 15",
  "id": "s9b",
  "type": "graded-dictation",
  "label": "Word Cards 131–140 dictation",
  "content": {
   "prompt": "Tap ▶ to hear each word — no text shown, same as Mom reading from the card.",
   "words": [
    {
     "answer": "victimize",
     "kind": "word"
    },
    {
     "answer": "generalize",
     "kind": "word"
    },
    {
     "answer": "departmentalize",
     "kind": "word"
    },
    {
     "answer": "specialize",
     "kind": "word"
    },
    {
     "answer": "materialize",
     "kind": "word"
    },
    {
     "answer": "hospitalize",
     "kind": "word"
    },
    {
     "answer": "systematize",
     "kind": "word"
    },
    {
     "answer": "fertilize",
     "kind": "word"
    },
    {
     "answer": "apologize",
     "kind": "word"
    },
    {
     "answer": "organize",
     "kind": "word"
    }
   ]
  }
 },
 {
  "week": 9.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 15",
  "id": "s9c",
  "type": "graded-dictation",
  "label": "Dictated sentences",
  "content": {
   "prompt": "All 10 sentences from this week's Reinforcement section.",
   "words": [
    {
     "answer": "We need to systematize the way we count the penguins.",
     "kind": "sentence"
    },
    {
     "answer": "Tomorrow we are helping Farmer Jim fertilize his fields.",
     "kind": "sentence"
    },
    {
     "answer": "They specialize in umbrellas and rain coats.",
     "kind": "sentence"
    },
    {
     "answer": "Let's organize a performance of famous scenes from plays.",
     "kind": "sentence"
    },
    {
     "answer": "You can't generalize about the habits of men and women.",
     "kind": "sentence"
    },
    {
     "answer": "The office will be more efficient if we departmentalize it.",
     "kind": "sentence"
    },
    {
     "answer": "He seems to make music materialize wherever he goes.",
     "kind": "sentence"
    },
    {
     "answer": "The doctor knew he had to hospitalize the athlete.",
     "kind": "sentence"
    },
    {
     "answer": "The criminals tried to victimize those helpless people.",
     "kind": "sentence"
    },
    {
     "answer": "Did Beth ever apologize for eating all the potatoes?",
     "kind": "sentence"
    }
   ]
  }
 },
 {
  "week": 10.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Finish & Capstone · Ch. 25–33",
  "id": "r10a",
  "type": "read",
  "label": "This Week's Reading: Finish Amari And The Night Brothers",
  "content": {
   "content": "<div class=\"lesson-text\"><p><b>Finish <i>Amari and the Night Brothers</i>: Chapters 25–33.</b> <b>As you read the last section, watch for:</b></p><p>1. The most vivid image in the climax. Which exact words make it vivid?</p><p>2. Amari's magic is all about making <b>pictures people believe</b>. How do the author's words help <i>you</i> see what she creates? (Think: the author is doing with words what Amari does with magic.)</p><p>3. Go back in your mind to the very first chapters. How does the world feel now compared with how it felt at the start — what changed, and what stayed the same?</p></div>"
  }
 },
 {
  "week": 10.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Finish & Capstone · Ch. 25–33",
  "id": "r10b",
  "type": "reflection",
  "label": "Capstone: The Best Image In The Book",
  "content": {
   "prompt": "Choose the single best image or comparison in the whole book — the one that would be on the poster if the book were a movie. Copy it exactly. In 4–5 complete sentences: what is it, which exact words create it, which sense or device is doing the work (imagery, simile, metaphor, personification), and why is it the best one?",
   "sampleAnswer": "Parent guide: this is the month's mastery check for RL.7.4 (figurative and connotative meaning). Look for exact quoting, accurate naming of the device, and — most of all — an effect explained with specific words. A 'mastered' answer says why the choice of words mattered, not just that it was pretty. Amari-specific: her magic is illusion, which means making pictures people believe, so her illusion scenes are natural picks. Bonus if she notices that the author is doing with words what Amari does with magic."
  }
 },
 {
  "week": 10.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Finish & Capstone · Ch. 25–33",
  "id": "r10c",
  "type": "reflection",
  "label": "Capstone: Three Hidden Worlds",
  "content": {
   "prompt": "You have now read three stories where a kid steps into a hidden or secret place — Nevermoor, The Mysterious Benedict Society, and <i>Amari and the Night Brothers</i>. Choose TWO of them. What is ONE thing the authors do the same way when they introduce the hidden world, and ONE thing they do differently? Use a specific detail from each book. (5–6 complete sentences.)",
   "sampleAnswer": "Parent guide: this is a light, early taste of the synthesis lens she'll use at the end of the year. Accept any real comparison with specifics from both books. Good threads: how each author shows the world's rules, whether the world is welcoming or threatening, and whether the main character is chosen or lucky. Don't expect polish. Amari threads: how each kid gets in (Amari is invited by a briefcase her missing brother left; compare Nevermoor and the Benedict Society's test), whether the hidden world welcomes her (some Bureau officials and trainees look down on her, unlike the more welcoming Benedict Society), and what each kid has to prove. A strong answer notices that Amari's hidden world has prejudice built in."
  }
 },
 {
  "week": 10.0,
  "key": "reading",
  "name": "Reading",
  "tag": "Amari And The Night Brothers · Finish & Capstone · Ch. 25–33",
  "id": "r10d",
  "type": "reflection",
  "label": "Capstone: Wrap-Up",
  "content": {
   "prompt": "Was 'Figurative Language & Imagery' a good lens for this particular book? Tell me one thing you noticed this month that you would NOT have noticed if you had just read for plot. Then tell me one thing you will look for in the next book you read.",
   "sampleAnswer": "Parent guide: this one is reflection, not graded for content. Look for honesty and one specific noticing. This also gives you a read on whether she is starting to see the author's word choice on her own."
  }
 },
 {
  "week": 10.0,
  "key": "vocab",
  "name": "Vocabulary",
  "tag": "Greek & Latin Roots · Month 2",
  "id": "v10a",
  "type": "read",
  "label": "Study This Week's Words",
  "content": {
   "content": "<div class=\"lesson-text\"><p>This week's root is <b>STA / STIT</b> — Latin <i>stare</i> and <i>statuere</i>, \"to stand\" or \"to set in place.\"</p></div><table class=\"vocab-table\"><tr><td>substitute <button class=\"btn word-audio-btn\" onclick=\"speakWord('substitute')\">🔊</button></td><td>SUB- (under/in place of) + STIT- (set). To put one thing in place of another. Use it: name a plant-based substitute for an egg or for milk, and tell what it does in the recipe.</td></tr><tr><td>obstacle <button class=\"btn word-audio-btn\" onclick=\"speakWord('obstacle')\">🔊</button></td><td>OB- (against) + STA- (stand). Something that stands in your way. Use it: describe an obstacle you faced in a recent recipe, and how you got past it.</td></tr><tr><td>constitute <button class=\"btn word-audio-btn\" onclick=\"speakWord('constitute')\">🔊</button></td><td>CON- (together) + STIT- (set). To make up or form. Use it: say how many tasks constitute one Rouxbe lesson, or what ingredients constitute a basic vinaigrette.</td></tr><tr><td>circumstance <button class=\"btn word-audio-btn\" onclick=\"speakWord('circumstance')\">🔊</button></td><td>CIRCUM- (around) + STA- (stand). The conditions standing around an event. Use it: describe a circumstance in which you would change your dinner plans.</td></tr><tr><td>principal <button class=\"btn word-audio-btn\" onclick=\"speakWord('principal')\">🔊</button></td><td>From Latin <i>princeps</i> (first, chief). The main or most important one; in Money Lab, the original amount of money before interest. Use it: say what 'principal' means in a savings account, and then explain how it is different from the spelling 'principle' (a rule or belief).</td></tr></table>"
  }
 },
 {
  "week": 10.0,
  "key": "vocab",
  "name": "Vocabulary",
  "tag": "Greek & Latin Roots · Month 2",
  "id": "v10b",
  "type": "reflection",
  "label": "Root Detective",
  "content": {
   "prompt": "Pick two of this week's words. For each one: name the root and what it means, give ONE more word you know with the same root (or find one), and write your own sentence that uses the word correctly. Then, in one sentence, explain the difference between 'principal' and 'principle' so you'll never mix them up.",
   "sampleAnswer": "Parent guide: more STA/STIT words — stable, status, station, statue, stamina, establish, institute, constitution, understand (stand + under), stationary/stationery (a homophone pair worth a look-up: stationary = not moving, stationery = paper and envelopes). 'Principal' is not from the STA root; it's from the 'first/chief' root, so it's a good chance to point out that not every word in a unit fits the pattern. Memory anchor for the spelling: 'the principal is your pal.'"
  }
 },
 {
  "week": 10.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 4 (Participial Openers In Your Writing)",
  "id": "g10prep",
  "type": "read",
  "label": "Using Participial Openers In Your Own Writing",
  "content": {
   "content": "<div class=\"lesson-text\"><p>Three weeks ago, participial phrases were new. Now you're going to use them in <b>your own writing</b>. Here is a recipe for a strong opener:</p><p><b>1.</b> Pick the person or thing doing the main action. <b>2.</b> Think of a second thing they were doing at the same time — an <i>-ing</i> or <i>-ed</i> description. <b>3.</b> Put that description first, add a comma, then name the person or thing.</p><p><i>Plain:</i> Kenley watched the cream thicken. <i>With an opener:</i> <b>Whisking furiously,</b> Kenley watched the cream thicken.</p><p><b>Your job this week:</b> add <b>at least one participial opener</b> to your Block 2 analysis, check it with the \"who is doing it?\" test, and copy it into the Opener Challenge below. Variety matters — don't start every sentence this way!</p></div><table class=\"vocab-table\"><tr><td>Reminder</td><td>Participial phrase = describes a noun. Dangling = the noun after the comma can't do the action. One comma after the opener.</td></tr></table>"
  }
 },
 {
  "week": 10.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 4 (Participial Openers In Your Writing)",
  "id": "g10a",
  "type": "pos-tagger",
  "label": "Level 1 — Tag The Parts Of Speech",
  "content": {
   "sentence": [
    "Whisking",
    "furiously,",
    "Kenley",
    "beat",
    "the",
    "cream",
    "until",
    "stiff",
    "peaks",
    "formed",
    "in",
    "the",
    "bowl."
   ],
   "answers": [
    "Adjective",
    "Adverb",
    "Noun",
    "Verb",
    "Adjective",
    "Noun",
    "Conjunction",
    "Adjective",
    "Noun",
    "Verb",
    "Preposition",
    "Adjective",
    "Noun"
   ],
   "options": [
    "Noun",
    "Pronoun",
    "Verb",
    "Adjective",
    "Adverb",
    "Preposition",
    "Conjunction",
    "Interjection"
   ],
   "explanations": [
    "\"Whisking\" is a participle — an adjective made from a verb — and it describes Kenley.",
    "It tells how Kenley was whisking.",
    "It names a specific person — a proper noun.",
    "It shows the action Kenley performed.",
    "Articles (a, an, the) are treated as adjectives — they modify the noun that follows.",
    "It names a thing — the object of the action.",
    "\"Until\" introduces the dependent clause and connects it to the main clause.",
    "It describes \"peaks\" — adjectives describe nouns.",
    "It names things — \"peaks\" is the subject of the dependent clause.",
    "It shows what the peaks did.",
    "It introduces a phrase telling where, connecting to its object, \"bowl.\"",
    "Articles (a, an, the) are treated as adjectives — they modify the noun that follows.",
    "It names a thing — the object of the preposition \"in.\""
   ]
  }
 },
 {
  "week": 10.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 4 (Participial Openers In Your Writing)",
  "id": "g10b",
  "type": "phrase-tagger",
  "label": "Level 2 — Tag The Parts Of The Sentence",
  "content": {
   "sentence": [
    "Whisking",
    "furiously,",
    "Kenley",
    "beat",
    "the",
    "cream",
    "until",
    "stiff",
    "peaks",
    "formed",
    "in",
    "the",
    "bowl."
   ],
   "options": [
    "Subject",
    "Action Verb Predicate",
    "Linking Verb Predicate",
    "Direct Object",
    "Indirect Object",
    "Subject Complement"
   ],
   "phrases": [
    {
     "start": 2,
     "end": 2,
     "type": "Subject",
     "explanation": "\"Kenley\" is who the main clause is about — the one doing the beating."
    },
    {
     "start": 3,
     "end": 3,
     "type": "Action Verb Predicate",
     "explanation": "\"Beat\" shows a real action and takes a direct object."
    },
    {
     "start": 5,
     "end": 5,
     "type": "Direct Object",
     "explanation": "\"Cream\" is what got beaten — the direct object answers \"what.\" Just the noun, same as how \"Kenley\" above is the Subject without \"the\" — the article is tagged separately at Level 1."
    },
    {
     "start": 8,
     "end": 8,
     "type": "Subject",
     "explanation": "\"Peaks\" is who/what the dependent clause is about."
    },
    {
     "start": 9,
     "end": 9,
     "type": "Action Verb Predicate",
     "explanation": "\"Formed\" shows what the peaks did. It doesn't need a direct object."
    }
   ]
  }
 },
 {
  "week": 10.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 4 (Participial Openers In Your Writing)",
  "id": "g10c",
  "type": "phrase-tagger",
  "label": "Level 3 — Mark The Phrases",
  "content": {
   "sentence": [
    "Whisking",
    "furiously,",
    "Kenley",
    "beat",
    "the",
    "cream",
    "until",
    "stiff",
    "peaks",
    "formed",
    "in",
    "the",
    "bowl."
   ],
   "options": [
    "Prepositional",
    "Appositive",
    "Participial"
   ],
   "phrases": [
    {
     "start": 0,
     "end": 1,
     "type": "Participial",
     "explanation": "\"Whisking\" is a participle (an adjective made from a verb), and the phrase describes Kenley — a participial phrase."
    },
    {
     "start": 10,
     "end": 12,
     "type": "Prepositional",
     "explanation": "It starts with the preposition \"in\" and ends with its object \"bowl.\""
    }
   ]
  }
 },
 {
  "week": 10.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 4 (Participial Openers In Your Writing)",
  "id": "g10d",
  "type": "phrase-tagger",
  "label": "Level 4 — Mark The Clauses",
  "content": {
   "sentence": [
    "Whisking",
    "furiously,",
    "Kenley",
    "beat",
    "the",
    "cream",
    "until",
    "stiff",
    "peaks",
    "formed",
    "in",
    "the",
    "bowl."
   ],
   "options": [
    "Independent Clause",
    "Dependent Clause"
   ],
   "phrases": [
    {
     "start": 0,
     "end": 5,
     "type": "Independent Clause",
     "explanation": "It has its own subject and predicate (\"Kenley beat\") and could stand alone — that makes it independent."
    },
    {
     "start": 6,
     "end": 12,
     "type": "Dependent Clause",
     "explanation": "It starts with the subordinating conjunction \"until\" and can't stand alone — that makes it dependent."
    }
   ]
  }
 },
 {
  "week": 10.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 4 (Participial Openers In Your Writing)",
  "id": "g10e",
  "type": "graded-mc",
  "label": "Level 4 — Classify The Sentence Structure",
  "content": {
   "questions": [
    {
     "q": "\"Whisking furiously, Kenley beat the cream until stiff peaks formed in the bowl.\" — what is this sentence's structure?",
     "options": [
      "Simple",
      "Compound",
      "Complex",
      "Compound-Complex"
     ],
     "correct": 2,
     "explanation": "One independent clause + one dependent clause = complex."
    },
    {
     "q": "Who does the participial phrase \"Whisking furiously\" describe?",
     "options": [
      "Kenley",
      "the cream",
      "the peaks"
     ],
     "correct": 0,
     "explanation": "Kenley is the one whisking, and \"Kenley\" comes right after the comma — the phrase is placed correctly."
    },
    {
     "q": "\"Kenley whisked the cream, and Waffles poured the syrup.\" — what is this sentence's structure?",
     "options": [
      "Simple",
      "Compound",
      "Complex",
      "Compound-Complex"
     ],
     "correct": 1,
     "explanation": "Two independent clauses joined by a comma + \"and\" = compound. Neither clause is dependent."
    }
   ]
  }
 },
 {
  "week": 10.0,
  "key": "grammar",
  "name": "Grammar",
  "tag": "MCT-Based · Month 2, Week 4 (Participial Openers In Your Writing)",
  "id": "g10f",
  "type": "reflection",
  "label": "Opener Challenge",
  "content": {
   "prompt": "Copy ONE sentence from your Block 2 analysis (the draft you are finishing this week) that starts with a participial phrase. Underline the participial phrase and tell which noun it describes. If your draft doesn't have one yet, add one now and copy it here.",
   "sampleAnswer": "Parent guide: check the phrase is correctly placed (no dangler), followed by a comma, and that she can name the noun it describes. One good opener is the goal; don't ask for more."
  }
 },
 {
  "week": 10.0,
  "key": "writing",
  "name": "Writing",
  "tag": "Book-Tied Analysis · Block 2",
  "id": "w10a",
  "type": "reflection",
  "label": "Final Draft: Put It Together And Revise",
  "content": {
   "prompt": "Paste your final, clean analysis here — Paragraphs 1, 2, and 3 in order. Before you paste, check each item: (1) hook opening, not throat-clearing · (2) the author's choice named in one clear sentence · (3) what we believed → what changed → where, with one exact quote or detail you could point to on the page, followed by a 'This shows…' sentence · (4) WHY the author did it, with the effect on you as a reader · (5) at least one sentence that opens with a participial phrase · (6) the Plate Check: capitals, complete sentences, 'I' capitalized, names spelled right, and every word you weren't sure about looked up.",
   "sampleAnswer": "4-point rubric from the Framework (Cycle 2 / Book-Tied Analysis): names a specific structural choice (1) / accurately states what changed and when (1) / explains the WHY of the craft choice, not just what happened (2). Add 1 bonus point for the craft moves visible: hook (Craft Lesson 2), quote sandwich, participial opener. A paragraph that only retells the plot caps at 2 out of 4."
  }
 },
 {
  "week": 10.0,
  "key": "writing",
  "name": "Writing",
  "tag": "Book-Tied Analysis · Block 2",
  "id": "w10b",
  "type": "reflection",
  "label": "Capstone Reflection",
  "content": {
   "prompt": "Look at your first draft of Paragraph 1 and your final version. What is ONE change you made that made it better, and how do you know it's better? Then name the Block 2 skill that was hardest for you (naming the choice, using a quote, or explaining WHY), and tell what you'd do differently next time.",
   "sampleAnswer": "Parent guide: not graded for correctness. This is a metacognition check — look for a real change named specifically and an honest 'hardest' skill. If 'explaining WHY' is hardest, that is typical; it's the skill the next Book-Tied piece (Cycle 4, POV retell) will build on."
  }
 },
 {
  "week": 10.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 16",
  "id": "s10a",
  "type": "read",
  "label": "This Week's Focus",
  "content": {
   "content": "<div class=\"lesson-text\"><p><b>Step 16 — /sh/ Spelled SS or SSI.</b> Two patterns for the /sh/ sound this week.</p><p><b>Spelled <u>ss</u>:</b> tissue, pressure, issue. Listen to <i>tissue</i>, and the /sh/ sound is made by the two s’s.</p><p><b>Spelled <u>ssi</u>, at the end as /shun/ = <u>ssion</u>:</b> intermission, confession, admission, session, impression, permission, transmission. Every one of these words ends in <u>ssion</u>.</p><p>Bonus words that follow the same patterns: assure, commission, fissure, mission, omission, submission.</p><p><b>Homophone Watch-Out.</b> These pairs sound the same but are spelled differently, so in this week’s dictated sentences the meaning decides the spelling (the same idea as <i>hymn / him</i> in Step 7). Listen for these:</p><p><b>needs / kneads</b> — “our truck <u>needs</u> a new transmission” means requires; <i>kneads</i> is what a baker does to dough. <b>new / knew</b> — “a <u>new</u> transmission” means not old.</p><p><b>our / hour</b> (“<u>our</u> science experiment”) · <b>no / know</b> (“<u>no</u> one liked”) · <b>one / won</b> (“no <u>one</u>”) · <b>for / four</b> (“pay <u>for</u> admission”) · <b>made / maid</b> (“made my head ache”) · <b>so / sew / sow</b> (“so awful”).</p></div>"
  }
 },
 {
  "week": 10.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 16",
  "id": "s10b",
  "type": "graded-dictation",
  "label": "Word Cards 141–150 dictation",
  "content": {
   "prompt": "Tap ▶ to hear each word — no text shown, same as Mom reading from the card.",
   "words": [
    {
     "answer": "tissue",
     "kind": "word"
    },
    {
     "answer": "pressure",
     "kind": "word"
    },
    {
     "answer": "issue",
     "kind": "word"
    },
    {
     "answer": "intermission",
     "kind": "word"
    },
    {
     "answer": "confession",
     "kind": "word"
    },
    {
     "answer": "admission",
     "kind": "word"
    },
    {
     "answer": "session",
     "kind": "word"
    },
    {
     "answer": "impression",
     "kind": "word"
    },
    {
     "answer": "permission",
     "kind": "word"
    },
    {
     "answer": "transmission",
     "kind": "word"
    }
   ]
  }
 },
 {
  "week": 10.0,
  "key": "spelling",
  "name": "Spelling",
  "tag": "AAS Level 7 · Step 16",
  "id": "s10c",
  "type": "graded-dictation",
  "label": "Dictated sentences",
  "content": {
   "prompt": "All 10 sentences from this week's Reinforcement section.",
   "words": [
    {
     "answer": "How much did you pay for admission to the circus?",
     "kind": "sentence"
    },
    {
     "answer": "The mechanic said that our truck needs a new transmission.",
     "kind": "sentence"
    },
    {
     "answer": "My article will appear in the next issue of the medical journal.",
     "kind": "sentence"
    },
    {
     "answer": "The cats ripped the tissue box to pieces again!",
     "kind": "sentence"
    },
    {
     "answer": "That session with my drum teacher made my head ache!",
     "kind": "sentence"
    },
    {
     "answer": "We got the impression that no one liked our science experiment.",
     "kind": "sentence"
    },
    {
     "answer": "The chief expects the thieves to make a full confession.",
     "kind": "sentence"
    },
    {
     "answer": "Did you already ask permission to dig for ancient fossils?",
     "kind": "sentence"
    },
    {
     "answer": "The actors were so awful that we left at intermission.",
     "kind": "sentence"
    },
    {
     "answer": "Pam always loses her composure when she's under pressure.",
     "kind": "sentence"
    }
   ]
  }
 }
];

var BANKS_MONTH1 = {
 "vocab": [
  {
   "word": "rival",
   "correct": "The two bakers were rivals, each hoping to win the same blue ribbon.",
   "wrong": "She felt rival after eating three bowls of soup."
  },
  {
   "word": "ominous",
   "correct": "Dark clouds gathered overhead, an ominous sign that the picnic was doomed.",
   "wrong": "The ominous puppy wagged its tail and begged to play."
  },
  {
   "word": "irrational",
   "correct": "Being afraid of a harmless spoon is irrational.",
   "wrong": "Her irrational plan was carefully tested and made perfect sense."
  },
  {
   "word": "perceptive",
   "correct": "A perceptive taster can notice a pinch of salt that others miss.",
   "wrong": "He was so perceptive that he never noticed anything going on around him."
  },
  {
   "word": "dismiss",
   "correct": "Don't dismiss her suggestion before you've tried it.",
   "wrong": "She dismissed the bowl of cereal and poured it with milk."
  },
  {
   "word": "deception",
   "correct": "Pretending to be sick to skip school was a deception.",
   "wrong": "Her deception was completely honest and truthful."
  },
  {
   "word": "reluctant",
   "correct": "I was reluctant to taste the beet smoothie, so I took a very small sip.",
   "wrong": "She was reluctant, so she raced to be first in line."
  },
  {
   "word": "camouflage",
   "correct": "The moth's brown wings were perfect camouflage against the tree bark.",
   "wrong": "She wore bright orange camouflage so everyone in the crowd would spot her."
  },
  {
   "word": "complicit",
   "correct": "Anyone who watched the prank and said nothing was complicit.",
   "wrong": "She was complicit because she tried to stop the prank and told a teacher."
  },
  {
   "word": "injustice",
   "correct": "Giving the punishment to the wrong student was an injustice.",
   "wrong": "The judge was proud of the injustice because everyone got a fair trial."
  },
  {
   "word": "escalate",
   "correct": "Arguments can escalate quickly if no one stays calm.",
   "wrong": "To escalate the fire, Mom poured water on the pan and put it out."
  },
  {
   "word": "vulnerable",
   "correct": "Baby birds are vulnerable because they can't protect themselves.",
   "wrong": "The castle was vulnerable to every attack because nothing could break through its walls."
  },
  {
   "word": "discreet",
   "correct": "A discreet whisper let her pass the note without anyone noticing.",
   "wrong": "He was discreet and announced his secret over the loudspeaker."
  },
  {
   "word": "underestimate",
   "correct": "Don't underestimate how hungry a hiker can get.",
   "wrong": "They underestimated the test, so they studied for weeks because they thought it would be extremely hard."
  },
  {
   "word": "resourceful",
   "correct": "A resourceful camper can build a shelter from sticks and a tarp.",
   "wrong": "He was so resourceful that he gave up the moment a problem appeared."
  },
  {
   "word": "resourcefulness",
   "correct": "Her resourcefulness turned a few leftover vegetables into a full soup.",
   "wrong": "His resourcefulness showed when he sat helplessly and waited for rescue."
  },
  {
   "word": "futile",
   "correct": "Shouting at the thunderstorm to stop was futile.",
   "wrong": "It was futile because her plan worked perfectly on the first try."
  },
  {
   "word": "vindicated",
   "correct": "When the test results proved her theory, she felt vindicated.",
   "wrong": "He felt vindicated when everyone proved he was completely wrong."
  },
  {
   "word": "resolution",
   "correct": "The story's resolution tied up every loose end.",
   "wrong": "The plot's resolution is the moment the villain's plan first begins."
  },
  {
   "word": "credible",
   "correct": "A credible witness tells the truth and has good evidence.",
   "wrong": "Her credible story was full of lies that nobody believed."
  }
 ],
 "grammar": [
  {
   "q": "\"Although the wizard seemed exhausted, Pickles devoured the pizza.\" — which part is the dependent clause?",
   "options": [
    "\"Although the wizard seemed exhausted\"",
    "\"Pickles devoured the pizza\""
   ],
   "correct": 0
  },
  {
   "q": "Which of these can stand alone as a complete sentence?",
   "options": [
    "The parrot ate the pizza.",
    "Because the parrot ate the pizza."
   ],
   "correct": 0
  },
  {
   "q": "\"While the professor grew nervous, Nugget wrecked the display.\" — what is this sentence's structure?",
   "options": [
    "Complex",
    "Compound"
   ],
   "correct": 0
  },
  {
   "q": "\"Bruno kicked the ball, but the window broke.\" — what is this sentence's structure?",
   "options": [
    "Compound",
    "Complex"
   ],
   "correct": 0
  },
  {
   "q": "\"During the movie, Pickles devoured an entire pizza.\" — what kind of phrase is \"During the movie\"?",
   "options": [
    "Prepositional",
    "Appositive"
   ],
   "correct": 0
  },
  {
   "q": "\"Nugget, the tiniest hamster, wrecked the display.\" — what does \"the tiniest hamster\" do?",
   "options": [
    "Renames Nugget",
    "Shows where Nugget was"
   ],
   "correct": 0
  },
  {
   "q": "\"The soup tasted disgusting.\" — what is \"disgusting\"?",
   "options": [
    "Subject Complement",
    "Direct Object"
   ],
   "correct": 0
  },
  {
   "q": "\"Since the referee looked furious, Bruno froze.\" — which word signals the dependent clause?",
   "options": [
    "Since",
    "Bruno"
   ],
   "correct": 0
  }
 ],
 "reading": [
  {
   "q": "Why does the author never show what happens inside the Waiting Room?",
   "options": [
    "It builds suspense — we only learn about it from how the kids act afterward",
    "The author ran out of time to write it"
   ],
   "correct": 0
  },
  {
   "q": "Through whose eyes do we see almost the entire story?",
   "options": [
    "Reynie's",
    "Mr. Curtain's"
   ],
   "correct": 0
  },
  {
   "q": "When the four kids suddenly feel a strong, confusing emotion, how do we find out what's going on?",
   "options": [
    "Reynie notices his own reaction and guesses about his friends",
    "A narrator explains everything right away"
   ],
   "correct": 0
  },
  {
   "q": "Which kid notices what others miss, partly because no one expects her to be paying attention?",
   "options": [
    "Constance",
    "Kate"
   ],
   "correct": 0
  },
  {
   "q": "Why does the Whisperer work so well on the kids?",
   "options": [
    "It targets each child's deepest fear or wish",
    "It forces them with loud noises"
   ],
   "correct": 0
  },
  {
   "q": "At the climax, whose stubbornness turns out to matter the most?",
   "options": [
    "Constance's",
    "Mr. Benedict's"
   ],
   "correct": 0
  },
  {
   "q": "Because the story stays inside Reynie's point of view, how do we learn the big surprise about Constance?",
   "options": [
    "At the very end, at the same time Reynie does",
    "In Chapter 1, before anyone else"
   ],
   "correct": 0
  },
  {
   "q": "Why might the author have built a team of four kids instead of one hero?",
   "options": [
    "Each child's different strength is needed to solve the problem",
    "A single hero would have been too boring"
   ],
   "correct": 0
  },
  {
   "q": "What is the 'setup' in a plot?",
   "options": [
    "The part where the characters, setting, and problem are introduced",
    "The final scene where all the problems are solved"
   ],
   "correct": 0
  },
  {
   "q": "Which clue shows a story is told in third person limited?",
   "options": [
    "It says 'he/she' but only shows one character's thoughts",
    "It says 'I' and 'we' throughout"
   ],
   "correct": 0
  },
  {
   "q": "How does the book show that the Institute is not what it seems?",
   "options": [
    "The Helpers have had their memories erased and the Messengers are being controlled",
    "Everyone is happy and nothing is wrong"
   ],
   "correct": 0
  }
 ]
};

var BANKS_MONTH2 = {
 "vocab": [
  {
   "word": "commodity",
   "correct": "Wheat is a commodity that farmers sell by the ton.",
   "wrong": "She lived in a commodity with her three roommates."
  },
  {
   "word": "preserve",
   "correct": "Pickling is one way to preserve cucumbers for the winter.",
   "wrong": "To preserve the bread, she ate it all in one sitting."
  },
  {
   "word": "grievance",
   "correct": "His main grievance was that the restaurant never answered complaints.",
   "wrong": "The grievance was a cheerful party with balloons and cake."
  },
  {
   "word": "ritual",
   "correct": "Every Sunday, the family followed the same ritual of baking bread together.",
   "wrong": "A ritual is a sudden surprise that happens only once."
  },
  {
   "word": "abundant",
   "correct": "Apples were abundant in the orchard, so we picked basket after basket.",
   "wrong": "Water was abundant in the desert, so we rationed every single drop."
  },
  {
   "word": "versatile",
   "correct": "Chickpeas are versatile; you can roast, mash, or blend them.",
   "wrong": "The versatile tool could do only one thing and nothing else."
  },
  {
   "word": "convert",
   "correct": "Use the chart to convert ounces to grams.",
   "wrong": "She wanted to convert the cake, so she left it exactly the same."
  },
  {
   "word": "invert",
   "correct": "Invert the pan so the cake slides out onto the plate.",
   "wrong": "To invert the recipe, he kept every step in the same order exactly as written."
  },
  {
   "word": "controversy",
   "correct": "There is a controversy over whether pineapple belongs on pizza.",
   "wrong": "Everyone agreed completely, which caused a huge controversy."
  },
  {
   "word": "glycemic",
   "correct": "A food with a high glycemic index raises blood sugar quickly.",
   "wrong": "The glycemic of the cake was seven inches wide."
  },
  {
   "word": "conduct",
   "correct": "Copper pans conduct heat quickly and evenly.",
   "wrong": "The conduct was delicious on warm toast."
  },
  {
   "word": "conducive",
   "correct": "A calm, tidy kitchen is conducive to careful cooking.",
   "wrong": "The weather was conducive to an outdoor picnic, so we got soaked by a thunderstorm."
  },
  {
   "word": "induce",
   "correct": "Warm milk may induce sleep.",
   "wrong": "She tried to induce the soup by throwing it out."
  },
  {
   "word": "produce",
   "correct": "A single tomato plant can produce dozens of tomatoes.",
   "wrong": "The farm will produce the apples by eating all of them."
  },
  {
   "word": "contribution",
   "correct": "Each family member's contribution made the potluck better.",
   "wrong": "She took the contribution out of the bowl and gave nothing."
  },
  {
   "word": "infuse",
   "correct": "Infuse the oil with garlic by warming it gently.",
   "wrong": "To infuse the tea, remove the leaves before the water ever touches them."
  },
  {
   "word": "diffuse",
   "correct": "The scent of baking bread began to diffuse through the house.",
   "wrong": "He tried to diffuse the smell by sealing it in an airtight box."
  },
  {
   "word": "profuse",
   "correct": "After the surprise party, she offered profuse thanks.",
   "wrong": "Her profuse apology was just two quiet words that no one heard."
  },
  {
   "word": "effusive",
   "correct": "Grandma was effusive in her praise of the pie.",
   "wrong": "He gave a short, flat, effusive grunt."
  },
  {
   "word": "confuse",
   "correct": "Don't confuse baking soda with baking powder.",
   "wrong": "To confuse the recipe, she followed every step clearly in order."
  },
  {
   "word": "substitute",
   "correct": "You can substitute applesauce for oil in some muffins.",
   "wrong": "The substitute was exactly the original ingredient, never changed."
  },
  {
   "word": "obstacle",
   "correct": "A fallen tree was an obstacle on the trail.",
   "wrong": "The open road with nothing blocking it was a huge obstacle."
  },
  {
   "word": "constitute",
   "correct": "Flour, water, salt, and yeast constitute a basic bread dough.",
   "wrong": "A single raisin constitutes an entire fruit salad of forty fruits."
  },
  {
   "word": "circumstance",
   "correct": "Under the circumstances, we decided to eat inside because of the rain.",
   "wrong": "The circumstance was a spoon made out of silver."
  },
  {
   "word": "principal",
   "correct": "The principal amount of a loan is the original sum borrowed.",
   "wrong": "The principal of good cooking is tasting as you go."
  }
 ],
 "grammar": [
  {
   "q": "\"Humming softly, Waffles stirred the soup.\" — what does \"Humming softly\" describe?",
   "options": [
    "Waffles",
    "the soup"
   ],
   "correct": 0
  },
  {
   "q": "\"Perfectly toasted, the bread vanished.\" — what kind of phrase is \"Perfectly toasted\"?",
   "options": [
    "Participial",
    "Prepositional"
   ],
   "correct": 0
  },
  {
   "q": "Which sentence has a dangling modifier?",
   "options": [
    "\"Walking to the market, the rain soaked my jacket.\"",
    "\"Walking to the market, I got soaked by the rain.\""
   ],
   "correct": 0
  },
  {
   "q": "What is the comma rule for a participial phrase at the start of a sentence?",
   "options": [
    "Put a comma after it",
    "No comma is needed"
   ],
   "correct": 0
  },
  {
   "q": "A participle is a verb form that acts like a(n):",
   "options": [
    "adjective",
    "noun"
   ],
   "correct": 0
  },
  {
   "q": "Which word is a past participle?",
   "options": [
    "frozen",
    "freezing"
   ],
   "correct": 0
  }
 ],
 "reading": [
  {
   "q": "Which is an example of a simile?",
   "options": [
    "\"Her voice was as sharp as a knife.\"",
    "\"Her voice was a knife.\""
   ],
   "correct": 0
  },
  {
   "q": "Which is an example of a metaphor?",
   "options": [
    "\"The kitchen was a battlefield.\"",
    "\"The kitchen was like a battlefield.\""
   ],
   "correct": 0
  },
  {
   "q": "\"The oven grumbled as it warmed up.\" is an example of:",
   "options": [
    "Personification",
    "Simile"
   ],
   "correct": 0
  },
  {
   "q": "What is imagery?",
   "options": [
    "Words that appeal to the senses",
    "A picture drawn in the book"
   ],
   "correct": 0
  },
  {
   "q": "Naming a device is only half the job. What is the other half?",
   "options": [
    "Explaining its effect — what it makes you see or feel",
    "Counting how many words it has"
   ],
   "correct": 0
  },
  {
   "q": "Two words can mean nearly the same thing but feel different. That feeling is called:",
   "options": [
    "connotation",
    "punctuation"
   ],
   "correct": 0
  }
 ]
};

var LOOKBACK_PATCH = {
 "r3a": {
  "label": "This Week's Look-Back: Chapters 17–24",
  "content": "<div class=\"lesson-text\"><p><b>You've already finished the whole book, so this week you're going back to Chapters 17–24 with fresh eyes.</b> Flip back through them (skim, don't reread everything) and look for:</p><p>1. Reynie’s growing discomfort with lying and cheating in order to move up in rank. Notice this is an internal conflict we only know about because we’re inside his head.</p><p>2. A scene where all four kids use a different personal strength to solve the same small problem. Notice how the author uses this to remind you who each character is, without stating it outright.</p><p>3. A moment where the mission’s cost becomes real — where someone innocent ends up taking the blame for something the team caused.</p></div>",
  "contentObj": null
 },
 "r4a": {
  "label": "This Week's Look-Back: Chapters 25–32",
  "content": "<div class=\"lesson-text\"><p><b>You've already finished the book, so this week you're looking back at Chapters 25–32.</b> Skim these chapters and look for:</p><p>1. As the hidden messages get stronger, notice that each character reacts a little differently — especially Constance.</p><p>2. A near-disaster scene where the team is almost caught. Notice the tension is built through small physical details (who’s watching, what’s being passed between people) rather than the author just telling you \"they almost got caught.\"</p><p>3. Notice how much the team ends up depending on the character who seemed the least capable at the start of the book.</p></div>",
  "contentObj": null
 },
 "r5a": {
  "label": "This Week's Look-Back: Chapters 33–39 (The Finale)",
  "content": "<div class=\"lesson-text\"><p><b>You've already read the ending, so this week you're looking back at Chapters 33–39 — now that you know how it all turns out.</b> Skim them and look for:</p><p>1. Kate and Constance’s escape from the tower. Notice a choice Kate makes about teamwork instead of going it alone — it echoes something Mr. Benedict told her earlier in the book.</p><p>2. The moment Reynie tries to trick the Whisperer machine itself. Notice <i>why</i> it doesn’t work — that failure tells you something important about honesty as a theme in this book.</p><p>3. The climax, where the team’s whole plan comes down to Constance specifically. Think back over everything you’ve noticed about her since Chapter 1.</p><p>4. The final chapter ties up a lot of loose ends — families reunited, secrets explained. Look for the one surprising reveal about Constance herself.</p></div>",
  "contentObj": null
 },
 "w3a": {
  "label": "Expand to Two Connected Paragraphs",
  "content": null,
  "contentObj": {
   "prompt": "Pull up the paragraph you wrote last week. Add a second claim-evidence-explain paragraph that supports the SAME overall argument — don’t start a new topic. When you’re done, check: does your new evidence pass the specificity test and the fact-vs-opinion test? And do the two paragraphs actually work together as one argument, or do they feel like two separate ones? One more rule: every fact needs a source you can name (a study, a website, a book, or something from your own life). “According to surveys” only counts if you can say which survey. Finish each piece of evidence with a sentence that starts “This shows…”."
  }
 },
 "w4a": {
  "label": "Revision Pass",
  "content": null,
  "contentObj": {
   "prompt": "Pull up your Week 2 or Week 3 draft. Go through every piece of evidence you used and re-run both tests: the specificity test (could this evidence be swapped into a totally different essay and still fit?) and the fact-vs-opinion test (is it actually checkable, not just another opinion in disguise?). Add a third test, the source test: for each fact, write down where it came from. If you can’t name the source, either look it up with Mom and name it, or replace it with a specific detail from your own life. Revise anything that fails, check your spelling (read it out loud and look up any word you’re unsure of), then paste your revised version here."
  }
 }
};
