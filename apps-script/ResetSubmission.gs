/**
 * One-time cleanup: deletes one student's submission row for one task, so
 * that task shows as "Not started" again in the app. Use this to clear out
 * a task you were just testing/trying out yourself.
 *
 * Paste this file into the live Apps Script project, then in the function
 * dropdown select resetKenleyGrammarG2 and click Run — that clears
 * Kenley's Week 1 Grammar "Level 1 — Tag the Parts of Speech" (task g2),
 * the only grammar submission on record for her Week 1.
 *
 * To reset a different task later, call resetSubmission('kenley'|'adelyn',
 * 'taskId') directly from the editor's "Execute function" box, or add
 * another small wrapper like resetKenleyGrammarG2 below.
 */
function resetKenleyGrammarG2() {
  resetSubmission('kenley', 'g2');
  // g2 is the only grammar submission on record for Kenley's Week 1, so
  // every Grammar-subject AnswerLog row for her belongs to this same test.
  deleteAnswerLogRows('kenley', 'Grammar');
}

function deleteAnswerLogRows(student, subject) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AnswerLog');
  if (!sh) return;
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var studentCol = headers.indexOf('student');
  var subjectCol = headers.indexOf('subject');
  var removed = 0;
  for (var r = values.length - 1; r >= 1; r--) {
    if (values[r][studentCol] === student && values[r][subjectCol] === subject) {
      sh.deleteRow(r + 1);
      removed++;
    }
  }
  Logger.log('Removed ' + removed + ' AnswerLog row(s) for ' + student + '/' + subject + '.');
}

function resetSubmission(student, taskId) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Submissions');
  if (!sh) return;
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var studentCol = headers.indexOf('student');
  var taskCol = headers.indexOf('task_id');
  var removed = 0;
  for (var r = values.length - 1; r >= 1; r--) {
    if (values[r][studentCol] === student && values[r][taskCol] === taskId) {
      sh.deleteRow(r + 1);
      removed++;
    }
  }
  Logger.log('Removed ' + removed + ' submission row(s) for ' + student + '/' + taskId + '.');
}
