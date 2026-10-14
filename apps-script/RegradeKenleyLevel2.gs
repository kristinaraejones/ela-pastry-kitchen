/**
 * One-time cleanup after the Level 2 convention change (Direct Object and
 * Indirect Object are only the head noun/pronoun — "pizza", not "an entire
 * pizza"). Run AFTER fixDirectObjectSpans() has narrowed the keys in the
 * Schedule sheet, and AFTER fixScoreColumnFormat().
 *
 * What it does (all idempotent — safe to re-run):
 *   1. Re-scores every Kenley Submission for a Level 2 phrase-tagger
 *      (options include Direct Object) from its stored selections against the
 *      CURRENT answer key, using the same rules as phraseScore() in app.js
 *      (exact span match; a right chunk with the wrong label loses one item;
 *      chunks that aren't any answer add to the total). Writes the score back
 *      as plain text ("3/6") and logs old -> new for each task.
 *   2. Corrects the matching AnswerLog rows: a Direct/Indirect Object row that
 *      was logged against the old whole-phrase key ("the flashlight",
 *      "(missed)", FALSE) becomes the head noun ("flashlight") with the given
 *      answer filled in and correct = TRUE, when her stored tap was right.
 *      Rows are found by student + question + correct_answer + word text +
 *      timestamp (never by row number). "extra chunk" rows are left alone.
 *   3. Safeguard: a submission whose Direct/Indirect Object tap is a
 *      multi-token span (i.e. she tapped a whole old-style phrase) is
 *      grandfathered — its old score is kept and it is listed in the log.
 *      (Note: the app's on-load re-score in app.js would still re-score such
 *      a task; none exist in the Oct 13 export. If one ever shows up, tell
 *      Claude so the app can exempt it.)
 *
 * Run from the Apps Script editor: regradeKenleyLevel2 > Run.
 * Needs Code.gs (getSheet_, sheetToObjects_, getHeaders_, safeParse_,
 * normalizeScore_).
 */
function regradeKenleyLevel2() {
  var subSheet = getSheet_('Submissions');
  var subHeaders = getHeaders_(subSheet);
  var scoreCol = subHeaders.indexOf('score') + 1;
  if (scoreCol === 0) throw new Error('score column not found in Submissions.');

  // Level 2 tasks from the Schedule (current keys).
  var tasks = {};
  sheetToObjects_(getSheet_('Schedule')).forEach(function (r) {
    if (r.student !== 'kenley' || r.task_type !== 'phrase-tagger') return;
    var c = safeParse_(r.content_json, null);
    if (!c || !c.phrases || !c.sentence || !c.options) return;
    if (c.options.indexOf('Direct Object') === -1) return; // Level 2 option set only
    tasks[r.task_id] = c;
  });

  var report = [];
  var scoreWrites = [];
  var logTargets = []; // {task, phrase, submissionTime}
  var subRows = sheetToObjects_(subSheet).filter(function (r) { return r.student === 'kenley' && tasks[r.task_id]; });

  subRows.forEach(function (r) {
    var c = tasks[r.task_id];
    var a = safeParse_(r.answers_json, {});
    var sels = (a && a.selections) || [];
    var oldScore = normalizeScore_(r.score);
    if (!sels.length) { report.push([r.task_id, oldScore, '(no selections stored)', 'skipped']); return; }

    // Grandfather any whole-phrase Direct/Indirect Object tap.
    var multi = sels.filter(function (sel) {
      return (sel.type === 'Direct Object' || sel.type === 'Indirect Object') && sel.end > sel.start;
    });
    if (multi.length) {
      report.push([r.task_id, oldScore, oldScore, 'GRANDFATHERED (multi-word ' + multi[0].type + ' tap) — score kept']);
      return;
    }

    var sc = level2Score_(c, sels);
    var fresh = sc.ok + '/' + sc.total;
    if (fresh !== oldScore || r.score instanceof Date) {
      scoreWrites.push({ row: r._row, value: fresh });
      report.push([r.task_id, oldScore, fresh, fresh === oldScore ? 'rewritten as text (was a Date cell)' : 'CHANGED']);
    } else {
      report.push([r.task_id, oldScore, fresh, 'unchanged']);
    }
    c.phrases.forEach(function (p) {
      if ((p.type === 'Direct Object' || p.type === 'Indirect Object') && sels.some(function (sel) { return level2Hit_(p, sel); })) {
        logTargets.push({ task: r.task_id, phrase: p, c: c, time: r.timestamp });
      }
    });
  });

  scoreWrites.forEach(function (w) {
    var cell = subSheet.getRange(w.row, scoreCol);
    cell.setNumberFormat('@');
    cell.setValue(w.value);
  });

  // ---- AnswerLog ----
  var logSheet = getSheet_('AnswerLog');
  var logHeaders = getHeaders_(logSheet);
  var col = function (h) { return logHeaders.indexOf(h) + 1; };
  var logRows = sheetToObjects_(logSheet).filter(function (r) {
    return r.student === 'kenley' && r.question === 'Phrase/clause tagging';
  });
  var logFixed = 0;
  var flippedMisses = {}; // old word text (lowercase) -> number of FALSE rows flipped to TRUE
  logTargets.forEach(function (tg) {
    var head = tg.c.sentence.slice(tg.phrase.start, tg.phrase.end + 1).join(' ').replace(/[.,;:!?]+$/, '');
    var cutoff = tg.time ? new Date(tg.time).getTime() + 24 * 3600 * 1000 : null;
    logRows.forEach(function (lr) {
      if (String(lr.correct_answer) !== tg.phrase.type) return;
      var w = String(lr.word);
      if (w.toLowerCase() !== head.toLowerCase() && !new RegExp('\\s' + head.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '$', 'i').test(w)) return;
      if (cutoff && lr.timestamp && new Date(lr.timestamp).getTime() > cutoff) return; // a later attempt, not this submission
      var isTrue = lr.correct === true || String(lr.correct).toLowerCase() === 'true';
      if (w === head && isTrue) return; // already corrected
      if (!isTrue) flippedMisses[w.toLowerCase()] = (flippedMisses[w.toLowerCase()] || 0) + 1;
      logSheet.getRange(lr._row, col('word')).setValue(head);
      logSheet.getRange(lr._row, col('given_answer')).setValue(tg.phrase.type);
      logSheet.getRange(lr._row, col('correct')).setValue(true);
      logFixed++;
      Logger.log('AnswerLog row ' + lr._row + ' (' + tg.task + '): "' + w + '" -> "' + head + '", correct -> TRUE');
    });
  });

  // ---- ReviewPool ----
  // Every logged miss also lands in ReviewPool (Code.gs logAnswer), so a wrong
  // "miss" of "the flashlight" left a bogus active review word. If the pool row
  // is nothing but those flipped misses (never answered correctly, not missed more
  // often than we flipped), retire it as 'mastered' instead of deleting anything.
  var poolSheet = getSheet_('ReviewPool');
  var poolHeaders = getHeaders_(poolSheet);
  var statusCol = poolHeaders.indexOf('status') + 1;
  var poolRetired = 0;
  sheetToObjects_(poolSheet).forEach(function (pr) {
    if (pr.student !== 'kenley') return;
    var n = flippedMisses[String(pr.word).toLowerCase()];
    if (!n || String(pr.status) !== 'active') return;
    if ((Number(pr.times_correct) || 0) > 0 || (Number(pr.times_missed) || 0) > n) return;
    if (statusCol) poolSheet.getRange(pr._row, statusCol).setValue('mastered');
    poolRetired++;
    Logger.log('ReviewPool: retired "' + pr.word + '" (it was only a wrongly-logged Level 2 miss).');
  });

  Logger.log('--- Level 2 regrade: task | old score | new score | note ---');
  report.forEach(function (x) { Logger.log(x.join(' | ')); });
  var changed = report.filter(function (x) { return x[3] === 'CHANGED'; }).length;
  Logger.log('Submissions re-scored: ' + changed + ' changed, ' + report.length + ' Level 2 submission(s) examined. AnswerLog rows corrected: ' + logFixed + '. ReviewPool entries retired: ' + poolRetired + '.');
  if (!changed) Logger.log('No Submissions score changed — only AnswerLog rows (and score cell formats) were corrected.');
}

// Same rules as spanHit / phraseHit / phraseScore in app.js.
function level2Hit_(p, sel) { return p.start === sel.start && p.end === sel.end; }
function level2Score_(c, selections) {
  var seen = {};
  var sels = selections.filter(function (sel) {
    var k = sel.start + '-' + sel.end + '-' + sel.type;
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  });
  var ok = sels.filter(function (sel) { return c.phrases.some(function (p) { return p.type === sel.type && level2Hit_(p, sel); }); }).length;
  var extra = sels.filter(function (sel) { return !c.phrases.some(function (p) { return level2Hit_(p, sel); }); }).length;
  return { ok: ok, total: c.phrases.length + extra };
}
