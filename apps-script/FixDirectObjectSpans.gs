/**
 * One-time live-data patch: narrows the "Direct Object" span in Kenley's
 * Level 2 grammar (Weeks 1-5) down to just the head noun, matching how
 * Subject is already tagged (bare noun, no article/adjectives — those are
 * Level 1's job). Previously Direct Object spanned the whole noun phrase
 * ("the flashlight," "an entire pizza," "the entire pot," "the ball,"
 * "the entire display"), which was inconsistent with Subject and with
 * real MCT-style four-level analysis. See the convention note above
 * GRAMMAR_LEVEL_OPTIONS_SEED in Setup.gs for the full reasoning.
 *
 * Because Level 2's phrase-tagger score is re-graded against the current
 * answer key every time the app loads (regradeLegacyDictation in app.js),
 * correcting this key is enough on its own — any of Kenley's past Level 2
 * submissions will automatically re-score against the corrected key the
 * next time she (or a parent) opens the app. No submission data needs
 * touching here, only the Schedule row holding the answer key itself.
 *
 * Run once from the Apps Script editor (fixDirectObjectSpans > Run).
 * Safe to re-run: rows already matching the corrected span are skipped.
 */
function fixDirectObjectSpans() {
  var sh = getSheet_('Schedule');
  var rows = sheetToObjects_(sh);
  var headers = getHeaders_(sh);
  var contentCol = headers.indexOf('content_json');
  if (contentCol === -1) throw new Error('content_json column not found.');

  var FIXES = {
    g3: { start: 11, end: 11, explanation: '"Flashlight" is what got carried — the direct object answers "what." Just the head noun, same as how "Reynie" above is the Subject without "the" — the article is already tagged separately as an adjective at Level 1.' },
    g2b: { start: 12, end: 12, explanation: '"Pizza" is what got devoured — the direct object answers "what." Just the head noun, same as how "Pickles" above is the Subject without "the" — "an" and "entire" are already tagged separately at Level 1.' },
    g3b: { start: 12, end: 12, explanation: '"Pot" is what got dumped — the direct object answers "what." Just the head noun, same as how "Waffles" above is the Subject without "the" — "the" and "entire" are already tagged separately at Level 1.' },
    g4b: { start: 11, end: 11, explanation: '"Ball" is what got launched — the direct object answers "what." Just the head noun, same as how "Bruno" above is the Subject without "the" — "the" is already tagged separately at Level 1.' },
    g5b: { start: 12, end: 12, explanation: '"Display" is what got wrecked — the direct object answers "what." Just the head noun, same as how "Nugget" above is the Subject without "the" — "the" and "entire" are already tagged separately at Level 1.' }
  };

  var fixed = 0, skipped = 0;
  rows.forEach(function (r) {
    if (r.student !== 'kenley' || r.subject_key !== 'grammar' || !FIXES[r.task_id]) return;
    var fix = FIXES[r.task_id];
    var content = safeParse_(r.content_json, null);
    if (!content || !content.phrases) { Logger.log('No phrases[] on ' + r.task_id + ' — skipping.'); return; }

    var target = content.phrases.find(function (p) { return p.type === 'Direct Object'; });
    if (!target) { Logger.log('No Direct Object entry on ' + r.task_id + ' — skipping.'); return; }

    if (target.start === fix.start && target.end === fix.end) { skipped++; return; }

    target.start = fix.start;
    target.end = fix.end;
    target.explanation = fix.explanation;
    sh.getRange(r._row, contentCol + 1).setValue(JSON.stringify(content));
    fixed++;
    Logger.log('Fixed Direct Object span on ' + r.task_id + ' (row ' + r._row + ').');
  });
  Logger.log('Fixed ' + fixed + ' row(s), ' + skipped + ' already correct.');
}
