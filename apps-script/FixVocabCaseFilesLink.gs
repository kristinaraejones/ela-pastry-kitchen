/**
 * One-time live-data patch: the "Play Mystery Round 3 — Case Files" task
 * (kenley / vocab / v2) was seeded before the Case Files game had a real
 * URL, so its content_json never got a `url` field and the button in the
 * app was a dead link (fixed in app.js/Setup.gs, but seedScheduleIfEmpty_
 * only seeds an empty sheet, so the already-live row needs this instead).
 *
 * Run once from the Apps Script editor (Fix Vocab Case Files Link ▸ Run).
 * Safe to re-run — it just overwrites this one row's content_json with the
 * corrected object, same as it should have been seeded in the first place.
 */
function fixVocabCaseFilesLink() {
  var sh = getSheet_('Schedule');
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var studentCol = headers.indexOf('student');
  var subjectCol = headers.indexOf('subject_key');
  var taskCol = headers.indexOf('task_id');
  var contentCol = headers.indexOf('content_json');

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row[studentCol] === 'kenley' && row[subjectCol] === 'vocab' && row[taskCol] === 'v2') {
      var content = JSON.parse(row[contentCol] || '{}');
      content.url = 'https://root-and-bloom-case-files.vercel.app';
      content.linkText = 'Open Vocab Case Files';
      content.note = "Opens in a new tab. Works offline once you've loaded it there at least once.";
      sh.getRange(i + 1, contentCol + 1).setValue(JSON.stringify(content));
      Logger.log('Fixed: Kenley’s Vocab Case Files task now links to ' + content.url);
      return;
    }
  }
  Logger.log('Row not found (student=kenley, subject_key=vocab, task_id=v2) — nothing changed.');
}
