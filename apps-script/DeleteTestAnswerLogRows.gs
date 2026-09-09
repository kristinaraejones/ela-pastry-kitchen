/**
 * One-time cleanup: removes the "TEST-connection-check" rows that were
 * pushed into AnswerLog while verifying the logAnswer endpoint was live.
 *
 * Paste this file into the live Apps Script project and run
 * deleteTestAnswerLogRows() once from the editor (select it in the
 * function dropdown, click Run). Only removes rows where subject is
 * exactly "TEST" — never touches real answer data.
 */
function deleteTestAnswerLogRows() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AnswerLog');
  if (!sh) return;
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var subjectCol = headers.indexOf('subject');
  var removed = 0;
  // Walk bottom-up so deleting a row doesn't shift the indices of rows
  // still to be checked.
  for (var r = values.length - 1; r >= 1; r--) {
    if (values[r][subjectCol] === 'TEST') {
      sh.deleteRow(r + 1); // +1: values[] is 0-indexed, sheet rows are 1-indexed
      removed++;
    }
  }
  Logger.log('Removed ' + removed + ' TEST row(s) from AnswerLog.');
}
