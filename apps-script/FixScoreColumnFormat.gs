/**
 * One-time fix for a Sheets gotcha: a score like "7/10" looks exactly like
 * a date (July 10) to Sheets' auto-detection, so writing it as a plain
 * string still silently got stored as a real Date cell — that's why some
 * scores were showing up as timestamps (e.g. "2026-07-10T00:00:00.000Z")
 * instead of "7/10". Code.gs's readSubmissions_ now reconstructs the
 * fraction from a Date cell on read (normalizeScore_), so this fixes
 * display immediately without needing this script — but this script also
 * rewrites the affected cells back to real plain-text "M/D" values and
 * locks the column to Plain Text formatting, so the raw sheet itself is
 * clean and this can't recur for future scores.
 *
 * Paste into the live Apps Script project and run fixScoreColumnFormat()
 * once. Safe to re-run — rows already stored as text are left alone.
 *
 * Updated Oct 2026: (1) converts each Date back to "M/d" using the
 * SPREADSHEET's time zone (the script's zone is Africa/Abidjan = UTC, which
 * can be a day off from where Sheets stored the date), and (2) locks the
 * WHOLE score column (every row the sheet has, not just rows that exist
 * today) to Plain Text so future rows can't be converted either.
 */
function fixScoreColumnFormat() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Submissions');
  if (!sh) return;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var scoreCol = headers.indexOf('score') + 1;
  if (scoreCol === 0) return;

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var dataRange = sh.getRange(2, scoreCol, lastRow - 1, 1);
  var values = dataRange.getValues();
  var fixed = 0;
  for (var i = 0; i < values.length; i++) {
    var v = values[i][0];
    if (v instanceof Date) {
      values[i][0] = Utilities.formatDate(v, tz, 'M/d');
      fixed++;
    }
  }
  // Lock the entire column (to the sheet's max rows) BEFORE writing values back,
  // so already-existing AND not-yet-created rows are plain text.
  sh.getRange(2, scoreCol, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  dataRange.setValues(values);
  Logger.log('Fixed ' + fixed + ' corrupted score cell(s); score column locked to plain text.');
}
