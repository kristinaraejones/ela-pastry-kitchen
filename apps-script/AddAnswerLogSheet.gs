/**
 * One-time patch: adds the AnswerLog sheet used by the new `logAnswer`
 * action (Code.gs) so Kenley's Case Files and Adelyn's Word Bakery can
 * report real answer-by-answer history into the parent view.
 *
 * Paste this file into the live Apps Script project and run
 * addAnswerLogSheet() once from the editor (select it in the function
 * dropdown, click Run). Safe to re-run — it only creates the sheet/header
 * row if missing, never touches existing data.
 */
function addAnswerLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var headers = ['student', 'timestamp', 'game', 'subject', 'word', 'question', 'given_answer', 'correct_answer', 'correct'];
  var sh = ss.getSheetByName('AnswerLog');
  if (!sh) sh = ss.insertSheet('AnswerLog');
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  try {
    SpreadsheetApp.getUi().alert('AnswerLog sheet ready. Re-deploy the Web App (Deploy > Manage deployments > edit > New version) so the new logAnswer action goes live.');
  } catch (e) {
    // getUi() only works when triggered from the Sheet's own UI, not when
    // run from the Apps Script editor's Run button — the sheet is already
    // created by this point regardless, so there's nothing to fix here.
    Logger.log('AnswerLog sheet ready. Re-deploy the Web App (Deploy > Manage deployments > edit > New version) so the new logAnswer action goes live.');
  }
}
