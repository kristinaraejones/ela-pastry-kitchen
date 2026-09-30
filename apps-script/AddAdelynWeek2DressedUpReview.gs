/**
 * One-time live-data patch: adds a short "Quick Review: Dressed-Up Sentences"
 * task (task_id awr2) to Adelyn's Writing Week 2, placed BEFORE the
 * "Action Words With Personality" lesson (awl2).
 *
 * Adelyn has been adding extra simple sentences instead of adding adjectives/
 * adverbs inside one sentence, so this reteaches it and has her write 2 more.
 *
 * Run once from the Apps Script editor (addAdelynWeek2DressedUpReview > Run).
 * Safe to re-run: it does nothing if an awr2 row already exists.
 * The same task was also added to AdelynWriting.gs so a future re-seed keeps it.
 */
function addAdelynWeek2DressedUpReview() {
  var sh = getSheet_('Schedule');
  var values = sh.getDataRange().getValues();
  var h = values[0];
  var studentCol = h.indexOf('student'), subjectCol = h.indexOf('subject_key');
  var taskCol = h.indexOf('task_id'), typeCol = h.indexOf('task_type');
  var labelCol = h.indexOf('label'), contentCol = h.indexOf('content_json');
  var task = {
    "id": "awr2",
    "label": "Quick Review: Dressed-Up Sentences",
    "type": "reflection",
    "prompt": "<p><b>Before we try action words, let's make sure our dressed-up sentences are really dressed up!</b></p><p>Last week, lots of dressed-up sentences turned into <i>extra simple sentences</i>. That's a good start, but dressing up a sentence means adding describing words <b>inside the same sentence</b>, not writing more sentences after it.</p><table class=\"vocab-table\"><tr><td><b>Simple sentence</b></td><td>The dog ran.</td></tr><tr><td><b>Not dressed up</b> (just more simple sentences)</td><td>The dog ran. It was big. It was fast.</td></tr><tr><td><b>Dressed up</b> (ONE sentence, 2 describing words)</td><td>The <b>enormous</b> dog ran <b>swiftly</b>.</td></tr></table><p><b>The two kinds of describing words to pick from:</b></p><ul><li><b>Adjective</b>: describes a <i>noun</i> (a person, place or thing). It answers \"what kind?\" (<i>enormous</i> dog, <i>sparkly</i> leotard).</li><li><b>Adverb</b>: describes a <i>verb</i> (the action). It answers \"how?\" (ran <i>swiftly</i>, stirred <i>carefully</i>). Many end in <b>-ly</b>.</li></ul><p><b>How to do it:</b></p><ol><li>Write a <b>simple sentence</b>: who or what, and what they did.</li><li>Pick <b>2 describing words</b> (adjectives or adverbs, your choice).</li><li>Squeeze them into the <b>same sentence</b>. Don't start a new one!</li><li>Check: is it still <b>one sentence</b> with <b>one period</b>?</li></ol><p><b>Your turn:</b> do this two times, about cooking or gymnastics. Use this shape for each one:</p><p>Simple sentence: ______<br>Dressed-up sentence: ______ (circle or name your 2 describing words)</p>",
    "sampleAnswer": "<i>An illustrative example only, not a target:</i><br>Simple: Adelyn stirred the batter. → Dressed up: Adelyn stirred the <b>sticky</b> batter <b>carefully</b>.<br>Simple: The gymnast landed. → Dressed up: The <b>tiny</b> gymnast landed <b>proudly</b>.<br><br><b>Coaching note for you:</b> What to look for: each dressed-up sentence should be ONE sentence with exactly 2 new describing words added inside it. If she adds extra simple sentences again (\"The batter was sticky. It was gooey.\"), don't correct it; ask, \"Can you squeeze those describing words into the first sentence?\" A nudge like \"What kind of batter? How did she stir it?\" helps her find the two words. If she does it well, approve it and celebrate before the Week 2 lesson."
  };

  var anchorRow = -1;
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (r[studentCol] !== 'adelyn' || r[subjectCol] !== 'writing') continue;
    if (r[taskCol] === 'awr2') { Logger.log('awr2 already exists - nothing to do.'); return; }
    if (r[taskCol] === 'awl2') anchorRow = i + 1; // 1-based sheet row
  }
  if (anchorRow < 0) throw new Error("Couldn't find Adelyn's Week 2 lesson row (awl2).");

  sh.insertRowBefore(anchorRow);
  var copy = values[anchorRow - 1].slice();
  copy[taskCol] = task.id;
  copy[typeCol] = task.type;
  copy[labelCol] = task.label;
  copy[contentCol] = JSON.stringify({ prompt: task.prompt, sampleAnswer: task.sampleAnswer });
  sh.getRange(anchorRow, 1, 1, copy.length).setValues([copy]);
  Logger.log('Added awr2 before awl2 at sheet row ' + anchorRow + '.');
}
