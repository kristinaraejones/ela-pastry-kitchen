/**
 * One-time live-data patch: adds the "Play in the Word Bakery" external-link
 * task (task_id avwb) to every week of Adelyn's Vocabulary in the live
 * Schedule sheet. The seed data already has it, but seedScheduleIfEmpty_ only
 * seeds an empty sheet, so already-live rows need this instead.
 *
 * Run once from the Apps Script editor (addWordBakeryTaskToLiveVocab > Run).
 * Safe to re-run: weeks that already have an avwb row are skipped.
 * Each new row is a copy of that week's study-words row (av1, av2, ...) with the task fields swapped.
 */
function addWordBakeryTaskToLiveVocab() {
  var sh = getSheet_('Schedule');
  var values = sh.getDataRange().getValues();
  var h = values[0];
  var studentCol = h.indexOf('student'), subjectCol = h.indexOf('subject_key');
  var weekCol = h.indexOf('week_number'), taskCol = h.indexOf('task_id');
  var typeCol = h.indexOf('task_type'), labelCol = h.indexOf('label');
  var contentCol = h.indexOf('content_json');
  var content = JSON.stringify({
    url: 'https://kristinaraejones.github.io/ela-pastry-kitchen/word-bakery/',
    linkText: 'Open Word Bakery',
    note: "Opens in a new tab. Works offline once you've loaded it there at least once. Covers all her vocab words so far, not just this week's — great for review any week."
  });

  var hasTask = {};
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (r[studentCol] === 'adelyn' && r[subjectCol] === 'vocab' && r[taskCol] === 'avwb') hasTask[r[weekCol]] = true;
  }
  var newRows = [];
  for (var k = 1; k < values.length; k++) {
    var t = values[k];
    if (t[studentCol] !== 'adelyn' || t[subjectCol] !== 'vocab' || !/^av\d+$/.test(t[taskCol])) continue;
    if (hasTask[t[weekCol]]) continue;
    var copy = t.slice();
    copy[taskCol] = 'avwb';
    copy[typeCol] = 'external';
    copy[labelCol] = 'Play in the Word Bakery';
    copy[contentCol] = content;
    newRows.push(copy);
  }
  if (newRows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
  Logger.log('Added the Word Bakery task to ' + newRows.length + ' week(s).');
}
