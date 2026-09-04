/**
 * Temporary helper menu for running Adelyn's seed functions from the
 * actual Google Sheet, instead of the Apps Script editor's "Run" button.
 *
 * WHY THIS EXISTS: the Apps Script editor's function dropdown has been
 * unreliable in this project (not indexing newly-pushed functions,
 * independent of file size). A custom Sheet menu references functions by
 * name directly and doesn't depend on that dropdown at all, so it's a
 * reliable way to invoke them regardless.
 *
 * HOW TO USE: after this file is pushed, reload the actual Google Sheet
 * (not the Apps Script editor) in your browser — reloading re-runs
 * onOpen() and a new "Adelyn Setup" menu appears next to Extensions/Help.
 * Click each item once, in any order. Each one is idempotent (safe to
 * click more than once — it'll just tell you it's already done, or for
 * the full-replace seeders, just re-applies the current content).
 *
 * DELETE THIS FILE once everything on it has been run — it's scaffolding,
 * not part of the app's ongoing behavior.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Adelyn Setup')
    .addItem('1. Seed Vocabulary & Reading', 'seedAdelynVocabReading_')
    .addItem('2. Seed Spelling', 'seedAdelynSpelling_')
    .addItem('3. Seed Grammar (Unit 1)', 'seedAdelynGrammarUnit1_')
    .addItem('4. Seed Writing', 'seedAdelynWriting_')
    .addItem('5. Seed Missed Words From Trackers (Kenley + Adelyn)', 'seedMissedWordsFromTrackers')
    .addToUi();
}
