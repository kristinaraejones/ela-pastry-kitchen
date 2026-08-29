/**
 * The ELA Pastry Kitchen — one-time setup.
 *
 * Run setupSheets() ONCE from the Apps Script editor (select it in the
 * function dropdown, click Run) after pasting Code.gs and this file into a
 * new Apps Script project bound to your Google Sheet. It creates every tab
 * with headers and seeds Kenley's real Week 1 content plus Adelyn's
 * placeholder shell, exactly as described in the build spec. Re-running it
 * later is safe — it only fills in tabs/rows that don't already exist, it
 * never overwrites work that's already been done in the sheet.
 */

var SHEET_HEADERS = {
  Schedule: ['student', 'subject_key', 'subject_name', 'subject_tag', 'week_number', 'task_id', 'task_type', 'label', 'content_json', 'dynamic_bank_key', 'term_final', 'monthly_test'],
  Submissions: ['student', 'task_id', 'timestamp', 'status', 'score', 'answers_json', 'parent_comment'],
  ReviewPool: ['student', 'word', 'times_missed', 'times_correct', 'last_seen', 'status', 'context_sentence'],
  MonthTestMarkers: ['student', 'subject_key', 'bank_position'],
  BurnLog: ['student', 'station', 'tag', 'date', 'reason', 'items_json'],
  Banks: ['student', 'subject_key', 'items_json'],
  Settings: ['key', 'value']
};

function setupSheets() {
  var ss = getSS_();
  Object.keys(SHEET_HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
    }
    if (sh.getLastRow() === 0) {
      sh.appendRow(SHEET_HEADERS[name]);
      sh.setFrozenRows(1);
    }
  });
  // Remove the default "Sheet1" if it's still there and empty.
  var def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0) ss.deleteSheet(def);

  seedScheduleIfEmpty_();
  seedBanksIfEmpty_();
  seedReviewPoolIfEmpty_();
  seedSettingsIfEmpty_();

  SpreadsheetApp.getUi().alert('Setup complete. Schedule, Banks, ReviewPool and Settings are seeded. Deploy this project as a Web App next (Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone).');
}

function seedScheduleIfEmpty_() {
  var sh = getSheet_('Schedule');
  if (sh.getLastRow() > 1) return; // already seeded
  var headers = SHEET_HEADERS.Schedule;
  var rows = [];
  ['kenley', 'adelyn'].forEach(function (student) {
    var data = DATA_BY_CHILD_SEED[student];
    Object.keys(data).forEach(function (subjectKey) {
      var subject = data[subjectKey];
      subject.tasks.forEach(function (t) {
        var content = {};
        Object.keys(t).forEach(function (k) {
          if (['id', 'label', 'type', 'dynamic', 'termFinal', 'monthlyTest'].indexOf(k) === -1) content[k] = t[k];
        });
        rows.push([
          student, subjectKey, subject.name, subject.tag, 1,
          t.id, t.type, t.label, JSON.stringify(content),
          t.dynamic || '', !!t.termFinal, !!t.monthlyTest
        ]);
      });
    });
  });
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function seedBanksIfEmpty_() {
  var sh = getSheet_('Banks');
  if (sh.getLastRow() > 1) return;
  var rows = [
    ['kenley', 'vocab', JSON.stringify(VOCAB_MONTH_WORDS_SEED)],
    ['kenley', 'spelling', JSON.stringify(SPELLING_MONTH_WORDS_SEED)],
    ['kenley', 'grammar', JSON.stringify(GRAMMAR_MONTH_BANK_SEED)],
    ['kenley', 'reading', JSON.stringify(READING_MONTH_BANK_SEED)],
    ['adelyn', 'vocab', JSON.stringify([])],
    ['adelyn', 'spelling', JSON.stringify([])],
    ['adelyn', 'grammar', JSON.stringify([])],
    ['adelyn', 'reading', JSON.stringify([])]
  ];
  sh.getRange(2, 1, rows.length, SHEET_HEADERS.Banks.length).setValues(rows);
}

function seedReviewPoolIfEmpty_() {
  var sh = getSheet_('ReviewPool');
  if (sh.getLastRow() > 1) return;
  var seed = [
    ['kenley', 'separate', 2, 0, 'Step 3', 'active', ''],
    ['kenley', 'believe', 1, 0, 'Step 5', 'active', ''],
    ['kenley', 'though', 1, 0, 'Step 1', 'active', ''],
    ['kenley', 'disagree', 2, 0, 'Step 4', 'active', '']
  ];
  sh.getRange(2, 1, seed.length, SHEET_HEADERS.ReviewPool.length).setValues(seed);
}

function seedSettingsIfEmpty_() {
  var sh = getSheet_('Settings');
  if (sh.getLastRow() > 1) return;
  var dueDate = new Date();
  dueDate.setDate(dueDate.getDate() + 4);
  var rows = [
    ['currentWeekNumber', 1],
    ['dueDate', dueDate.toISOString().slice(0, 10)],
    ['termFinalsUnlocked', 'false'],
    ['monthlyTestOverride', ''] // '' = automatic schedule, 'true' = forced locked, 'false' = forced open
  ];
  sh.getRange(2, 1, rows.length, SHEET_HEADERS.Settings.length).setValues(rows);
}

// ---------- Seed content, ported 1:1 from the validated mockup ----------

var GRAMMAR_LEVEL_OPTIONS_SEED = {
  pos: ['Noun', 'Pronoun', 'Verb', 'Adjective', 'Adverb', 'Preposition', 'Conjunction', 'Interjection'],
  sentence: ['Subject', 'Action Verb Predicate', 'Linking Verb Predicate', 'Direct Object', 'Indirect Object', 'Subject Complement'],
  phrase: ['Prepositional', 'Appositive'],
  clause: ['Independent Clause', 'Dependent Clause'],
  structure: ['Simple', 'Compound', 'Complex', 'Compound-Complex']
};

var VOCAB_MONTH_WORDS_SEED = [
  { word: 'incredulous', correct: "She stared at the letter, incredulous that she'd actually won.", wrong: 'She was incredulous about doing her chores every single day without complaint.' },
  { word: 'deduction', correct: 'From the muddy footprints, he made a quick deduction about where she\'d gone.', wrong: 'He made a loud deduction when the door slammed.' },
  { word: 'meticulous', correct: 'The chef was meticulous about leveling every measuring cup exactly.', wrong: 'The chef was meticulous and forgot the recipe entirely.' },
  { word: 'resilient', correct: 'Even after the frost, the little plant proved resilient and grew back.', wrong: 'The plant was resilient and died within a day.' },
  { word: 'dictate', correct: "The coach's packed schedule seemed to dictate everything about their week.", wrong: 'She tried to dictate a quiet nap in the corner.' },
  { word: 'corroborate', correct: 'Two separate witnesses corroborated her account of the accident.', wrong: 'He corroborated his sandwich with extra mustard.' },
  { word: 'candid', correct: 'Her candid opinion about the soup surprised the whole table.', wrong: 'He gave a candid punch to the wall.' },
  { word: 'ambiguous', correct: 'The instructions were so ambiguous that two people read them differently.', wrong: 'The instructions were ambiguous and perfectly clear to everyone.' },
  { word: 'caramelize', correct: 'The onions took twenty minutes to caramelize into a deep golden brown.', wrong: 'The onions began to caramelize into a bright, crisp white.' }
];

var SPELLING_MONTH_WORDS_SEED = [
  { word: 'onion' }, { word: 'opinion' }, { word: 'convenient' }, { word: 'companion' }, { word: 'union' },
  { word: 'autumn' }, { word: 'hymn', context: 'She sang a hymn in church.' }, { word: 'column' }, { word: 'condemn' }, { word: 'solemn' }
];

var GRAMMAR_MONTH_BANK_SEED = [
  { q: '"Reynie sat quietly among the other children in the room." — what kind of phrase is "among the other children"?', options: ['Prepositional', 'Appositive'], correct: 0 },
  { q: '"Sticky, the boy who never forgot anything, adjusted his glasses." — what kind of phrase is "the boy who never forgot anything"?', options: ['Prepositional', 'Appositive'], correct: 1 },
  { q: 'In "Reynie sat quietly among the other children in the room," what does "in the room" modify?', options: ['The verb "sat" (adverbial)', 'The noun "children" (adjectival)'], correct: 1 },
  { q: '"The letter, a mysterious invitation, arrived without warning." — what kind of phrase is "a mysterious invitation"?', options: ['Appositive', 'Prepositional'], correct: 0 },
  { q: '"She waited outside the classroom door." — what kind of phrase is "outside the classroom door"?', options: ['Prepositional', 'Appositive'], correct: 0 },
  { q: 'What does an appositive phrase do?', options: ['Renames a nearby noun', 'Shows location or direction'], correct: 0 },
  { q: '"The cake tasted delicious." — what kind of predicate is "tasted"?', options: ['Action Verb Predicate', 'Linking Verb Predicate'], correct: 1 },
  { q: '"Sticky memorized the entire manual." — what is "the entire manual"?', options: ['Direct Object', 'Subject Complement'], correct: 0 },
  { q: '"Kate tossed Constance the map." — what is "Constance" in this sentence?', options: ['Indirect Object', 'Direct Object'], correct: 0 },
  { q: '"Because the door creaked, everyone jumped." — which part is the dependent clause?', options: ['"Because the door creaked"', '"everyone jumped"'], correct: 0 },
  { q: '"The kids ran, and then they hid." — what is this sentence\'s structure?', options: ['Compound', 'Complex'], correct: 0 },
  { q: 'A sentence with one independent clause and no dependent clauses is:', options: ['Simple', 'Compound-Complex'], correct: 0 }
];

var READING_MONTH_BANK_SEED = [
  { q: 'Which analytical lens is Kenley using for this book?', options: ['Characterization', 'Plot Structure & POV'], correct: 1 },
  { q: 'This book is told in third person limited. What does that mean?', options: ['We only see inside one character\'s head — everyone else stays a mystery', 'We see inside every character\'s head at once'], correct: 0 },
  { q: 'Why might an author deliberately delay revealing key information (a plot structure choice)?', options: ['To build suspense and let readers form their own guesses', "It's usually just a mistake in early drafts"], correct: 0 },
  { q: 'Why do many kids panic during the strange test at the start of the book?', options: ['There are no right answers written down anywhere', 'The test room is completely empty'], correct: 0 },
  { q: 'Which of the four kids solves problems using whatever is in her bucket?', options: ['Kate', 'Sticky'], correct: 0 },
  { q: "What is Sticky's defining trait, based on the setup so far?", options: ['He never forgets anything he reads', 'He is the appointed leader of the group'], correct: 0 },
  { q: "The four kids are sent undercover into a school. What's the setup telling us about it?", options: ['Something is very wrong there', "It's an ordinary cooking school"], correct: 0 }
];

var SHARED_SENTENCE = ['Because', 'the', 'mission', 'seemed', 'dangerous,', 'Reynie,', 'the', 'newest', 'member,', 'carried', 'the', 'flashlight', 'through', 'the', 'tunnel.'];

var DATA_BY_CHILD_SEED = {
  kenley: {
    vocab: {
      name: 'Vocabulary', tag: 'Case Files · Vocab Set A',
      tasks: [
        { id: 'v1', label: 'Study Set A Words', type: 'read', content: '<div class="lesson-text"><p>Four words for this week, each tied to a root family you\'ll keep seeing:</p></div><table class="vocab-table"><tr><td>incredulous</td><td>CRED- (believe) + IN-. Use it: describe a moment you were incredulous about something a sibling or friend told you.</td></tr><tr><td>deduction</td><td>DUCT- (lead) + -TION. Use it: make a small deduction about something in the room right now, using only what you can observe.</td></tr><tr><td>meticulous</td><td>From Latin <i>metus</i> (fear) — "fearful" became "extremely careful." Use it: describe a task that requires being meticulous.</td></tr><tr><td>resilient</td><td>Latin <i>resilire</i>, "to rebound." Use it: name someone or something resilient.</td></tr></table>' },
        { id: 'v2', label: 'Play Mystery Round 3 — Case Files', type: 'external', linkText: 'Open Vocab Case Files', note: 'Score syncs back automatically once the game module is wired to the shared sheet.' },
        { id: 'v3', label: 'Monthly Comprehension Test', type: 'graded-mc', dynamic: 'vocabMonthBank', monthlyTest: true, questions: [] },
        { id: 'v4', label: 'Term Final — Vocabulary', type: 'graded-mc', dynamic: 'examVocabBank', termFinal: true, questions: [] }
      ]
    },
    spelling: {
      name: 'Spelling', tag: 'AAS Level 7 · Step 7',
      tasks: [
        { id: 's0', label: "This Week's Focus", type: 'read', content: '<div class="lesson-text"><p><b>Step 7 — /y/ Spelled I and /m/ Spelled MN.</b> Two patterns this week: the /y/ sound spelled with <u>i</u> (onion, opinion, convenient, companion, union), and the /m/ sound spelled <u>mn</u> at the end of a word (autumn, hymn, column, condemn, solemn).</p><p>Bonus pattern: when a suffix is added to an <u>mn</u> word, the m and n split into different syllables — <i>condemn → con-dem-na-tion</i>.</p><p><b>New homophone pair: hymn / him.</b> These sound identical but mean completely different things, so the only way to spell the right one is context, not sound.</p><p><b>hymn</b> — a religious song, usually sung in church. <i>"We stood to sing the closing hymn."</i> <i>"That hymn has four verses."</i></p><p><b>him</b> — the pronoun, referring to a boy or man (like "her" but for a male). <i>"I gave him the recipe."</i> <i>"She waited for him outside."</i></p><p>Quick check: if you could replace the word with "her" and the sentence still makes sense, it\'s <b>him</b>. If it\'s about singing or church music, it\'s <b>hymn</b>.</p></div>' },
        { id: 's1', label: 'Word Cards 51–60 dictation', type: 'graded-dictation', prompt: 'Tap ▶ to hear each word — no text shown, same as Mom reading from the card.', words: [
          { answer: 'onion', kind: 'word' }, { answer: 'opinion', kind: 'word' }, { answer: 'convenient', kind: 'word' },
          { answer: 'companion', kind: 'word' }, { answer: 'union', kind: 'word' }, { answer: 'autumn', kind: 'word' },
          { answer: 'hymn', kind: 'word', context: 'She sang a hymn in church.' },
          { answer: 'column', kind: 'word' }, { answer: 'condemn', kind: 'word' }, { answer: 'solemn', kind: 'word' }
        ] },
        { id: 's2', label: 'Dictated sentences', type: 'graded-dictation', prompt: "All 10 sentences from this week's Reinforcement section.", words: [
          { answer: 'No one ever asks for my opinion on the subject!', kind: 'sentence' },
          { answer: 'This stew is the perfect union of beef and vegetables.', kind: 'sentence' },
          { answer: 'She sang the hymn so loudly that we could not hear the orchestra.', kind: 'sentence' },
          { answer: 'Chris writes a column on travel for the newspaper.', kind: 'sentence' },
          { answer: "Mom said she will condemn my room if I don't clean it up!", kind: 'sentence' },
          { answer: 'The solemn actress walked slowly off the stage.', kind: 'sentence' },
          { answer: 'I eat two grapefruits and a raw onion every morning.', kind: 'sentence' },
          { answer: 'My horse Puddle is my best friend and companion.', kind: 'sentence' },
          { answer: 'Nothing is better than the gold leaves of autumn.', kind: 'sentence' },
          { answer: "I find it more convenient to let the cat lick the stamps.", kind: 'sentence' }
        ] },
        { id: 's3', label: 'Review Words (Auto-Pulled From Misses)', type: 'graded-dictation', dynamic: 'reviewPool', prompt: "These aren't hand-picked — they're pulled automatically from words Kenley has missed in past dictation, oldest/most-missed first. Get one right twice and it graduates out of rotation." },
        { id: 's4', label: 'Monthly Spelling Test', type: 'graded-dictation', dynamic: 'spellingMonthBank', monthlyTest: true, prompt: "Up to 15 words drawn from every step taught since the last monthly test — right now that's only Step 7's 10 words, so it'll use all 10 until more steps are loaded." },
        { id: 's5', label: 'Term Final — Spelling', type: 'graded-dictation', dynamic: 'examSpellingBank', termFinal: true, prompt: 'Up to 30 words drawn from every step taught this term.' }
      ]
    },
    grammar: {
      name: 'Grammar', tag: 'MCT-Based · Month 1, Level 3 Review',
      tasks: [
        { id: 'g1', label: "Prepositional vs. Appositive — What's the Difference", type: 'read', content: '<div class="lesson-text"><p><b>Prepositional phrase</b> — starts with a preposition (in, on, under, among, outside, before, after...) and ends with its object, a noun or pronoun. It tells you where, when, or how.</p><p><i>Example:</i> "Reynie sat quietly among the other children." The phrase "among the other children" starts with the preposition "among" and ends with the noun "children."</p><p><b>Appositive phrase</b> — renames or explains a nearby noun, usually set off by commas. It doesn\'t tell where or when — it just tells you more about who or what something is.</p><p><i>Example:</i> "Sticky, the boy who never forgot anything, adjusted his glasses." The phrase "the boy who never forgot anything" renames "Sticky" — it\'s not location or time information, it\'s an identity.</p><p><b>Quick test:</b> Does it start with a preposition and end with its object? That\'s prepositional. Does it rename the noun right next to it, and could you lift it out (commas and all) without losing the sentence\'s basic meaning? That\'s an appositive.</p><table class="vocab-table"><tr><td>Prepositional</td><td>"She waited outside the classroom door." — starts with "outside," ends with "door."</td></tr><tr><td>Appositive</td><td>"The letter, a mysterious invitation, arrived without warning." — "a mysterious invitation" renames "letter."</td></tr></table><p style="opacity:.8;font-size:0.8rem;">This week, all four levels of analysis work through the same sentence: <i>"Because the mission seemed dangerous, Reynie, the newest member, carried the flashlight through the tunnel."</i> Each level looks at it a different way.</p></div>' },
        { id: 'g2', label: 'Level 1 — Tag the Parts of Speech', type: 'pos-tagger', sentence: SHARED_SENTENCE,
          answers: ['Conjunction', 'Adjective', 'Noun', 'Verb', 'Adjective', 'Noun', 'Adjective', 'Adjective', 'Noun', 'Verb', 'Adjective', 'Noun', 'Preposition', 'Adjective', 'Noun'],
          options: GRAMMAR_LEVEL_OPTIONS_SEED.pos,
          explanations: [
            '"Because" introduces the dependent clause and connects it to the main clause — that\'s a conjunction\'s job.',
            'Articles (a, an, the) are treated as adjectives — they modify the noun that follows.',
            'It names a thing — nouns name people, places, things, or ideas.',
            'It shows what the subject "mission" is described as. We\'ll classify it as action or linking specifically in Level 2.',
            'It describes "mission" — adjectives describe nouns.',
            'It names a specific person — a proper noun.',
            'Article again.',
            'It describes "member," comparing degree.',
            'It names a person — a renamed identity for Reynie.',
            'It shows the action the subject performed.',
            'Article.',
            'It names a thing — the object of the action.',
            'It introduces a phrase showing where something happened, connecting to its object, "tunnel."',
            'Article.',
            'It names a place — the object of the preposition "through."'
          ] },
        { id: 'g3', label: 'Level 2 — Tag the Parts of the Sentence', type: 'phrase-tagger', sentence: SHARED_SENTENCE, options: GRAMMAR_LEVEL_OPTIONS_SEED.sentence,
          phrases: [
            { start: 2, end: 2, type: 'Subject', explanation: '"Mission" is who/what the dependent clause is about.' },
            { start: 3, end: 3, type: 'Linking Verb Predicate', explanation: '"Seemed" doesn\'t show an action — it links "mission" to a description. That makes it a linking verb predicate, not an action verb predicate.' },
            { start: 4, end: 4, type: 'Subject Complement', explanation: '"Dangerous" describes the subject "mission" after the linking verb "seemed" — that\'s exactly what a subject complement does.' },
            { start: 5, end: 5, type: 'Subject', explanation: '"Reynie" is who the main clause is about — the one doing the action.' },
            { start: 9, end: 9, type: 'Action Verb Predicate', explanation: '"Carried" shows a real action Reynie performed, and it takes a direct object — only action verbs can do that.' },
            { start: 10, end: 11, type: 'Direct Object', explanation: '"The flashlight" is what got carried — the direct object answers "what."' }
          ] },
        { id: 'g4', label: 'Level 3 — Mark the Phrases', type: 'phrase-tagger', sentence: SHARED_SENTENCE, options: GRAMMAR_LEVEL_OPTIONS_SEED.phrase,
          phrases: [
            { start: 6, end: 8, type: 'Appositive', explanation: 'It renames "Reynie" — telling you who Reynie is, not where or when anything happened, which is what makes it an appositive instead of prepositional.' },
            { start: 12, end: 14, type: 'Prepositional', explanation: 'It starts with the preposition "through" and ends with its object "tunnel" — that\'s the exact shape of a prepositional phrase.' }
          ] },
        { id: 'g5', label: 'Level 4 — Mark the Clauses', type: 'phrase-tagger', sentence: SHARED_SENTENCE, options: GRAMMAR_LEVEL_OPTIONS_SEED.clause,
          phrases: [
            { start: 0, end: 4, type: 'Dependent Clause', explanation: 'It starts with the subordinating conjunction "because" and can\'t stand alone as a complete sentence — that makes it dependent.' },
            { start: 5, end: 14, type: 'Independent Clause', explanation: 'It has its own subject and predicate ("Reynie... carried...") and could stand alone as a complete sentence — that makes it independent.' }
          ] },
        { id: 'g6', label: 'Level 4 — Classify the Sentence Structure', type: 'graded-mc', questions: [
          { q: '"Because the mission seemed dangerous, Reynie, the newest member, carried the flashlight through the tunnel." — one independent clause plus one dependent clause makes this sentence:', options: GRAMMAR_LEVEL_OPTIONS_SEED.structure, correct: 2,
            explanation: 'One independent clause + one dependent clause = complex. (Simple = one independent clause alone. Compound = two or more independent clauses joined. Compound-complex = two+ independent clauses plus at least one dependent clause.)' }
        ] },
        { id: 'g7', label: 'Identify the Phrase — Quick Check', type: 'graded-mc', questions: [
          { q: '"Reynie sat quietly among the other children in the room." — what kind of phrase is "among the other children"?', options: ['Prepositional', 'Appositive'], correct: 0 },
          { q: '"Sticky, the boy who never forgot anything, adjusted his glasses." — what kind of phrase is "the boy who never forgot anything"?', options: ['Prepositional', 'Appositive'], correct: 1 },
          { q: 'In sentence 1, what does "in the room" modify?', options: ['The verb "sat" (adverbial)', 'The noun "children" (adjectival)'], correct: 1 }
        ] },
        { id: 'g8', label: 'Monthly Comprehension Test', type: 'graded-mc', dynamic: 'grammarMonthBank', monthlyTest: true, questions: [] },
        { id: 'g9', label: 'Term Final — Grammar', type: 'graded-mc', dynamic: 'examGrammarBank', termFinal: true, questions: [] }
      ]
    },
    reading: {
      name: 'Reading', tag: 'The Mysterious Benedict Society · Pause Pt. 1',
      tasks: [
        { id: 'r0', label: "This Month's Lens: Plot Structure & POV", type: 'read', content: '<div class="lesson-text"><p><b>Plot structure</b> is the shape of how a story unfolds — what the author tells us, in what order, and why. Authors sometimes hold something back on purpose, not by accident, to build suspense or make a later reveal land harder. Think about how a good mystery never tells you who did it on page one — that\'s structure at work.</p><p><b>Point of view (POV)</b> is whose eyes we\'re seeing the story through. There are three main kinds:</p><p>• <b>First person</b> — the narrator is a character in the story, speaking as "I."<br>• <b>Third person omniscient</b> — the narrator says "he/she" but can see inside <i>every</i> character\'s head.<br>• <b>Third person limited</b> — the narrator says "he/she" but only shows us <i>one</i> character\'s thoughts. Everyone else stays a bit of a mystery.</p><p>This book is <b>third person limited</b>, staying close to Reynie the whole time — so we only ever know what Reynie personally sees, hears, or guesses.</p></div>' },
        { id: 'r1', label: 'Read chapters 1–8', type: 'read', content: '<div class="lesson-text"><p class="quote">There\'s a strange newspaper ad going around: "Are you a gifted child looking for special opportunities?" Hundreds of kids show up to take a test — except it isn\'t like any test you\'ve ever heard of...</p><p><b>As you read, watch for:</b></p><p>1. Moments where Reynie can only <i>guess</i> what another character is thinking or feeling — since we\'re locked to his POV, we never get confirmation, only his best guess. Notice when that guess turns out right or wrong.</p><p>2. Where you personally start suspecting what the "real test" is actually measuring, versus where the book confirms it. That gap is plot structure at work — the author controlling what you\'re allowed to know, and when.</p><p>3. Small details that seem like throwaway texture now — the kind of thing that might matter later. You don\'t need to solve anything yet, just start noticing.</p></div>' },
        { id: 'r2', label: 'Discussion — Plot Structure', type: 'reflection', prompt: 'The author waits a long time before revealing what the test was actually testing. Why do you think the author chose to delay that reveal instead of just telling us up front? What did the delay do for you as a reader — did it make you guess, get frustrated, pay closer attention?' },
        { id: 'r3', label: 'Discussion — Point of View', type: 'reflection', prompt: "Since we only see this story through Reynie's eyes, name one thing happening around him that we don't get to know for sure yet — something we might only find out later if the book ever showed us a different character's perspective." },
        { id: 'r4', label: 'Monthly Comprehension Test', type: 'graded-mc', dynamic: 'readingMonthBank', monthlyTest: true, questions: [] },
        { id: 'r5', label: 'Term Final — Reading & Literary Concepts', type: 'graded-mc', dynamic: 'examReadingBank', termFinal: true, questions: [] }
      ]
    },
    writing: {
      name: 'Writing', tag: 'MCT Essay Voyage (adapted) · Block 1',
      tasks: [
        { id: 'w1', label: 'Why This Matters', type: 'read', content: '<div class="lesson-text"><p>When you argue for something, it\'s easy to just repeat your opinion in different words and call that "support." Real arguments need reasons someone else could actually check — not just a louder version of the same belief. This block is about learning to tell the difference, so your writing convinces people instead of just restating what you already think.</p></div>' },
        { id: 'w2', label: 'Fact vs. Opinion — The Foundation', type: 'read', content: '<div class="lesson-text"><p><b>Fact:</b> a statement that can be checked and proven true or false — a date, a measurement, a documented result.</p><p><b>Opinion:</b> a personal judgment or preference that can\'t be objectively verified, even if lots of people agree with it.</p><p><b>Quick test:</b> Could a reasonable person disagree with this and not be "wrong"? If yes, it\'s an opinion. If there\'s a way to look it up or measure it, it\'s a fact.</p><table class="vocab-table"><tr><td>Opinion</td><td>"Chocolate is the best ice cream flavor." — no way to prove this; people can disagree forever.</td></tr><tr><td>Fact</td><td>"Vanilla outsells chocolate in the U.S. by a wide margin, per National Ice Cream Retailers data." — specific, checkable, cited.</td></tr><tr><td>Opinion</td><td>"Homework should be banned." — this is a position, not a checkable fact.</td></tr><tr><td>Fact-based evidence</td><td>"Students averaging under 7 hours of sleep scored lower on next-day tests in a 2019 Stanford study." — this could actually support the position above.</td></tr></table><p><b>Important:</b> Your <i>claim</i> is usually an opinion — that\'s normal, that\'s what an argument is. The mistake is supporting an opinion with <i>another</i> opinion instead of a fact. "Dogs are the best pets because they\'re the most lovable" doesn\'t work — "lovable" is just another opinion wearing a costume.</p></div>' },
        { id: 'w3', label: 'Claim, Evidence, Explain — The Specificity Test', type: 'read', content: '<div class="lesson-text"><p><b>Claim</b> — the specific point this paragraph is making (your opinion, stated clearly).</p><p><b>Evidence</b> — a fact, specific enough to check.</p><p><b>Explain</b> — one sentence connecting the evidence back to the claim in your own words. Never drop a fact and walk away — say why it matters.</p><p><b>The specificity test:</b> Could this evidence sentence be true of literally any argument on this topic, or is it locked to your exact point? If you could swap it into a totally different essay and it would still fit, it\'s too vague.</p></div>' },
        { id: 'w4', label: 'See It Done Well — A Model Paragraph', type: 'read', content: '<div class="lesson-text"><p><b>Claim:</b> Learning to cook at home teaches better food habits than eating out regularly.</p><p><b>Evidence:</b> A 2019 Cornell study found that people who cook dinner at home most nights eat significantly more vegetables than people who rarely cook.</p><p><b>Explain:</b> This shows the habit of cooking itself — not just having healthy food nearby — is what actually changes what people choose to eat, which is exactly why home cooking builds better habits than relying on restaurants.</p><p style="opacity:.8;font-size:0.8rem;">Notice: The evidence names a specific study and a specific finding — you couldn\'t paste it into an essay about a different topic and have it still make sense. That\'s what passing the specificity test looks like.</p></div>' },
        { id: 'w5', label: 'Practice Set A — Vague or Specific?', type: 'graded-mc', questions: [
          { q: 'Which is real evidence? A) "Dogs make the best pets because they\'re really loyal." B) "My dog waited by the door every day at 3:15 when I got home from school."', options: ['A', 'B'], correct: 1 },
          { q: 'Which is real evidence? A) "Homemade bread has no preservatives, so it goes stale in two days instead of two weeks." B) "Homemade bread tastes better because it\'s fresher and more natural."', options: ['A', 'B'], correct: 0 },
          { q: 'Which of these is a fact rather than an opinion?', options: ['"Recipe videos are the best way to learn to cook."', '"A 2021 survey found 60% of home cooks under 30 learned primarily from online video."'], correct: 1 }
        ] },
        { id: 'w6', label: 'Draft Your Free-Topic Argument', type: 'reflection', prompt: 'Pick a real opinion you hold. Write your draft paragraph — claim, evidence, explain — then apply the specificity test and the fact-vs-opinion check to your evidence before you\'re done.' }
      ]
    }
  },
  adelyn: {
    vocab: {
      name: 'Vocabulary', tag: 'Grade 4 · Custom Game — In Design',
      tasks: [
        { id: 'av1', label: 'Vocab Game — Coming Soon', type: 'read', content: '<div class="lesson-text"><p>Adelyn\'s vocab game is being designed separately (a memory/matching game — flip word cards and definition cards, match pairs, no root/antonym reasoning required). Once the game and Grade 4 word lists are ready, this station works exactly like Kenley\'s: weekly words, a game link, and a monthly comprehension test.</p></div>' }
      ]
    },
    spelling: {
      name: 'Spelling', tag: 'AAS Level 4 · Step 11',
      tasks: [
        { id: 'as1', label: 'Waiting on the AAS PDF', type: 'read', content: '<div class="lesson-text"><p>Once you drop in the AAS Level 4 teacher\'s material, this station will mirror Kenley\'s exactly: a "This Week\'s Focus" summary, real word/sentence dictation with audio playback, an auto-updating review pool of missed words, and a monthly test — all pulled straight from the actual Step 11 content, the same way Kenley\'s Step 7 was built.</p></div>' }
      ]
    },
    grammar: {
      name: 'Grammar', tag: 'Grade 4 · Pending Curriculum',
      tasks: [
        { id: 'ag1', label: 'Waiting on Grammar Curriculum', type: 'read', content: '<div class="lesson-text"><p>Drop in Adelyn\'s grammar scope and sequence whenever it\'s ready, and this becomes a full four-level analysis sequence like Kenley\'s: real lesson, then Levels 1–4 as interactive tap-to-tag exercises on one shared sentence, plus monthly and term tests.</p></div>' }
      ]
    },
    reading: {
      name: 'Reading', tag: 'Grade 4 · Pending Book & Lens',
      tasks: [
        { id: 'ar1', label: 'Waiting on Reading Curriculum', type: 'read', content: '<div class="lesson-text"><p>Once you share Adelyn\'s book and analytical lens, this becomes a full sequence like Kenley\'s: a lesson that actually teaches the lens, reading with concrete things to watch for, discussion questions tied to the lesson, and a monthly test.</p></div>' }
      ]
    },
    writing: {
      name: 'Writing', tag: 'Grade 4 · Pending Curriculum',
      tasks: [
        { id: 'aw1', label: 'Waiting on Writing Curriculum', type: 'read', content: '<div class="lesson-text"><p>Drop in Adelyn\'s writing block whenever it\'s ready, and this becomes a full self-sufficient sequence like Kenley\'s — the concept taught from the ground up, a worked model example, guided practice, and a draft assignment with a parent-review loop.</p></div>' }
      ]
    }
  }
};
