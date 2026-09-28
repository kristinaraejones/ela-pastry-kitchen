/**
 * One-time live-data patch: inserts a "Prepositions & Prepositional Phrases
 * — Quick Review" lesson into Kenley's Grammar for Weeks 2-5, each one
 * right before that week's Level 1 — Tag the Parts of Speech task
 * (g2prep before g2a, g3prep before g3a, etc). Month 1 Weeks 2-5 were
 * already seeded into the live Schedule sheet in an earlier session, so
 * the matching edit to Month1Weeks2to5.gs's source doesn't reach the live
 * data on its own — this patch does.
 *
 * Each lesson references that week's own sentence and its actual
 * prepositional phrase (during the movie / into the garden / through the
 * window / before lunch), so it's a real preview, not generic filler.
 *
 * readSchedule_ returns rows in raw sheet order with no sort, so a plain
 * append would land these AFTER that week's tagging tasks instead of
 * before — this uses insertRowBefore to physically place each new row
 * ahead of its week's g*a row in the sheet.
 *
 * Run once from the Apps Script editor (addGrammarPrepLessons > Run).
 * Safe to re-run: any week whose *prep row already exists is skipped.
 */
function addGrammarPrepLessons() {
  var sh = getSheet_('Schedule');
  var headers = getHeaders_(sh);

  var LESSONS = {
    g2prep: { before: 'g2a', body:
      '<div class="lesson-text"><p><b>Preposition</b> — a word that shows how a noun or pronoun relates to another word in the sentence: where, when, how, or which one. Common ones: <i>in, on, at, by, with, about, above, below, before, after, during, through, between, among, under, over, without, into, onto, from, near, until.</i></p>' +
      '<p><b>Prepositional phrase</b> — a preposition plus its object (a noun or pronoun), sometimes with a describing word or two in between. It always starts with the preposition and ends with its object.</p>' +
      '<p><i>Formula:</i> preposition + (describing words) + noun/pronoun</p>' +
      '<table class="vocab-table"><tr><td>"under the old bridge"</td><td>preposition "under," object "bridge," with "the old" describing the object in between.</td></tr>' +
      '<tr><td>"with her"</td><td>preposition "with," object "her" — pronouns can be the object of a preposition too.</td></tr>' +
      '<tr><td>"during the movie"</td><td>preposition "during," object "movie" — no describing words needed for a phrase to count.</td></tr></table>' +
      '<p><b>Quick test:</b> Find a word from the list above. Does it lead straight into a noun or pronoun (maybe with a describing word or two along the way)? That whole chunk — preposition through object — is the prepositional phrase.</p>' +
      '<p style="opacity:.8;font-size:0.8rem;">This week\'s sentence has one: "...Pickles, the ridiculous parrot, devoured an entire pizza <b>during the movie</b>." "During" is the preposition and "movie" is its object — that\'s the prepositional phrase to watch for when you tag Level 1 and beyond.</p></div>' },
    g3prep: { before: 'g3a', body:
      '<div class="lesson-text"><p><b>Preposition</b> — a word that shows how a noun or pronoun relates to another word in the sentence: where, when, how, or which one. Common ones: <i>in, on, at, by, with, about, above, below, before, after, during, through, between, among, under, over, without, into, onto, from, near, until.</i></p>' +
      '<p><b>Prepositional phrase</b> — a preposition plus its object (a noun or pronoun), sometimes with a describing word or two in between. It always starts with the preposition and ends with its object.</p>' +
      '<p><i>Formula:</i> preposition + (describing words) + noun/pronoun</p>' +
      '<table class="vocab-table"><tr><td>"during the movie"</td><td>preposition "during," object "movie," with "the" describing the object in between.</td></tr>' +
      '<tr><td>"with her"</td><td>preposition "with," object "her" — pronouns can be the object of a preposition too.</td></tr></table>' +
      '<p><b>Quick test:</b> Find a word from the list above. Does it lead straight into a noun or pronoun (maybe with a describing word or two along the way)? That whole chunk — preposition through object — is the prepositional phrase.</p>' +
      '<p style="opacity:.8;font-size:0.8rem;">This week\'s sentence has one: "...dumped the entire pot <b>into the garden</b>." "Into" is the preposition and "garden" is its object — that\'s the prepositional phrase to watch for when you tag Level 1 and beyond.</p></div>' },
    g4prep: { before: 'g4a', body:
      '<div class="lesson-text"><p><b>Preposition</b> — a word that shows how a noun or pronoun relates to another word in the sentence: where, when, how, or which one. Common ones: <i>in, on, at, by, with, about, above, below, before, after, during, through, between, among, under, over, without, into, onto, from, near, until.</i></p>' +
      '<p><b>Prepositional phrase</b> — a preposition plus its object (a noun or pronoun), sometimes with a describing word or two in between. It always starts with the preposition and ends with its object.</p>' +
      '<p><i>Formula:</i> preposition + (describing words) + noun/pronoun</p>' +
      '<table class="vocab-table"><tr><td>"into the garden"</td><td>preposition "into," object "garden," with "the" describing the object in between.</td></tr>' +
      '<tr><td>"with her"</td><td>preposition "with," object "her" — pronouns can be the object of a preposition too.</td></tr></table>' +
      '<p><b>Quick test:</b> Find a word from the list above. Does it lead straight into a noun or pronoun (maybe with a describing word or two along the way)? That whole chunk — preposition through object — is the prepositional phrase.</p>' +
      '<p style="opacity:.8;font-size:0.8rem;">This week\'s sentence has one: "...launched the ball straight <b>through the window</b>." "Through" is the preposition and "window" is its object. Watch out — "straight" right before it is an adverb describing "launched," not part of the phrase.</p></div>' },
    g5prep: { before: 'g5a', body:
      '<div class="lesson-text"><p><b>Preposition</b> — a word that shows how a noun or pronoun relates to another word in the sentence: where, when, how, or which one. Common ones: <i>in, on, at, by, with, about, above, below, before, after, during, through, between, among, under, over, without, into, onto, from, near, until.</i></p>' +
      '<p><b>Prepositional phrase</b> — a preposition plus its object (a noun or pronoun), sometimes with a describing word or two in between. It always starts with the preposition and ends with its object.</p>' +
      '<p><i>Formula:</i> preposition + (describing words) + noun/pronoun</p>' +
      '<table class="vocab-table"><tr><td>"through the window"</td><td>preposition "through," object "window," with "the" describing the object in between.</td></tr>' +
      '<tr><td>"before lunch"</td><td>preposition "before," object "lunch" — no describing word needed for a phrase to count.</td></tr></table>' +
      '<p><b>Quick test:</b> Find a word from the list above. Does it lead straight into a noun or pronoun (maybe with a describing word or two along the way)? That whole chunk — preposition through object — is the prepositional phrase.</p>' +
      '<p style="opacity:.8;font-size:0.8rem;">This week\'s sentence has one: "...wrecked the entire display <b>before lunch</b>." "Before" is the preposition and "lunch" is its object — no article or describing word this time, and that\'s still a complete prepositional phrase.</p></div>' }
  };

  var added = 0;
  Object.keys(LESSONS).forEach(function (lessonId) {
    var lesson = LESSONS[lessonId];
    // Re-read fresh each iteration since earlier inserts shift row numbers.
    var rows = sheetToObjects_(sh);
    var already = rows.some(function (r) {
      return r.student === 'kenley' && r.subject_key === 'grammar' && r.task_id === lessonId;
    });
    if (already) { Logger.log(lessonId + ' already present — skipping.'); return; }

    var target = rows.find(function (r) {
      return r.student === 'kenley' && r.subject_key === 'grammar' && r.task_id === lesson.before;
    });
    if (!target) { Logger.log('Could not find kenley/grammar/' + lesson.before + ' — skipping ' + lessonId + '.'); return; }

    var newRow = {
      student: 'kenley',
      subject_key: 'grammar',
      subject_name: 'Grammar',
      subject_tag: target.subject_tag,
      week_number: target.week_number,
      task_id: lessonId,
      task_type: 'read',
      label: 'Prepositions & Prepositional Phrases — Quick Review',
      content_json: JSON.stringify({ content: lesson.body }),
      dynamic_bank_key: '',
      term_final: '',
      monthly_test: ''
    };
    var rowValues = headers.map(function (h) { return newRow[h] !== undefined ? newRow[h] : ''; });

    sh.insertRowBefore(target._row);
    sh.getRange(target._row, 1, 1, headers.length).setValues([rowValues]);
    added++;
    Logger.log('Inserted ' + lessonId + ' before ' + lesson.before + ' at row ' + target._row + '.');
  });
  Logger.log('Added ' + added + ' prep lesson(s).');
}
