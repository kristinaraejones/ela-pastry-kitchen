/**
 * The ELA Pastry Kitchen — Adelyn's Vocabulary & Reading, Weeks 1-36.
 *
 * Adds net-new Schedule rows for Adelyn's "vocab" and "reading" subjects,
 * covering the full 36-week year, converted from
 * Adelyns_Vocabulary_Course_Grade4.docx and Adelyns_Reading_Program.docx
 * into the same content_json shape used for Kenley's content. Does NOT
 * touch setupSheets(), Month1Weeks2to5.gs, or any Kenley rows — this is
 * purely additive, and only touches Adelyn's placeholder vocab/reading rows.
 *
 * Run seedAdelynVocabReading_() ONCE from the function dropdown, after
 * setupSheets() has already been run at least once. It's idempotent: if it
 * detects task_id "av1" already in the Schedule sheet, it assumes this has
 * already run and does nothing. It also removes Adelyn's original single
 * placeholder rows for "vocab" and "reading" ("av1" waiting-shell and "ar1"
 * waiting-shell — note the real content below reuses "ar1" for Week 1's
 * actual reading lesson, so the OLD placeholder "ar1" is deleted by task
 * label match before the new rows are added, to avoid an id collision).
 *
 * Content structure per week:
 *   Vocabulary — one "read" task with that week's 3 root words + 3
 *     frequency words (table format, matching Kenley's vocab-table style).
 *     Every 4th week (4, 8, ... 36) adds a "Review Quiz" graded-mc task
 *     (monthlyTest: true) built directly from the course's own Word Bank +
 *     Parent Answer Key — each definition becomes a question, each week's
 *     12-word bank becomes that question's multiple-choice options, so the
 *     parent answer key is baked directly into the "correct" index. Week 36
 *     additionally adds the End-of-Year Vocabulary Test (termFinal: true).
 *   Reading — one "read" task (book, standard focus, pacing goal, and the
 *     week's Concept Spotlight teaching text), one "graded-dictation" task
 *     built from that week's Appendix A phonics warm-up word list plus its
 *     practice sentence (reuses the exact same listen-and-type mechanic as
 *     Kenley's spelling dictation — she taps to hear each word/sentence via
 *     text-to-speech and types what she hears), and one "reflection" task
 *     per Talk It Over discussion question, each with the course's own
 *     Sample Answer carried over as that task's sampleAnswer field (shown
 *     only in Parent view, to help grade her response).
 *
 * A "🔊 Read this to me" button (text-to-speech via speakElementText in
 * app.js) is available on every read/reflection/dictation/graded-mc task
 * for both children, not just Adelyn — added alongside this content so she
 * can have any lesson, question, or instruction read aloud on tap.
 *
 * NOTE ON FILE SPLIT: this seed function lives in its own small file,
 * separate from AdelynVocabReadingData.gs (which just holds the big
 * ADELYN_VOCAB_READING_WEEKS array, no functions of its own — Apps Script
 * shares one global scope across all files in a project, so this is fine
 * at runtime). This split exists purely because the Apps Script web
 * editor's "select function" dropdown silently fails to index functions
 * in very large files (this content alone is 300KB+) — keeping the
 * runnable function in a small file keeps it discoverable in the editor.
 */

function seedAdelynVocabReading_() {
  var sh = getSheet_('Schedule');
  var existing = sheetToObjects_(sh);
  if (existing.some(function (r) { return r.task_id === 'av1' && r.label === "Study This Week's Words"; })) {
    var alreadyMsg = 'Adelyn\'s Vocabulary/Reading content already appears to be seeded (found task av1) — skipping to avoid duplicates.';
    try { SpreadsheetApp.getUi().alert(alreadyMsg); } catch (e) { Logger.log(alreadyMsg); }
    return;
  }

  // Remove Adelyn's original single placeholder rows for vocab/reading
  // ("av1"/"ar1" waiting-shells) so the new "ar1" (Week 1's real reading
  // lesson) doesn't collide with the old placeholder id.
  var toDelete = existing.filter(function (r) {
    return r.student === 'adelyn' && (r.subject_key === 'vocab' || r.subject_key === 'reading') &&
      (r.task_id === 'av1' || r.task_id === 'ar1') && r.label.indexOf('Waiting') !== -1;
  });
  toDelete.sort(function (a, b) { return b._row - a._row; }).forEach(function (r) { sh.deleteRow(r._row); });

  var rows = [];
  function addTasks(subjectKey, subjectName, subjectTag, weekNumber, tasks) {
    tasks.forEach(function (t) {
      var content = {};
      Object.keys(t).forEach(function (k) {
        if (['id', 'label', 'type', 'dynamic', 'termFinal', 'monthlyTest'].indexOf(k) === -1) content[k] = t[k];
      });
      rows.push([
        'adelyn', subjectKey, subjectName, subjectTag, weekNumber,
        t.id, t.type, t.label, JSON.stringify(content),
        t.dynamic || '', !!t.termFinal, !!t.monthlyTest
      ]);
    });
  }

  ADELYN_VOCAB_READING_WEEKS.forEach(function (week) {
    Object.keys(week.subjects).forEach(function (subjectKey) {
      var subj = week.subjects[subjectKey];
      addTasks(subjectKey, subj.name, subj.tag, week.week_number, subj.tasks);
    });
  });

  var headers = SHEET_HEADERS.Schedule;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  var doneMsg = 'Added ' + rows.length + ' new Schedule rows for Adelyn — Vocabulary & Reading, all 36 weeks.';
  try { SpreadsheetApp.getUi().alert(doneMsg); } catch (e) { Logger.log(doneMsg); }
}
