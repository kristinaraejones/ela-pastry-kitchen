/**
 * One-time backfill: adds the missed spelling words Kristina tracked by hand
 * in Kenley_Spelling_Words_Missed.docx and Adelyn_Spelling_Words_Missed.docx
 * into the live ReviewPool sheet, so they show up in each kid's "missed
 * words" review rotation alongside anything the app has already tracked
 * from real dictation misses.
 *
 * Run seedMissedWordsFromTrackers() ONCE from the Apps Script editor. It is
 * additive and skip-safe: for any (student, word) pair that already exists
 * in ReviewPool — whether from the app's own tracking or a prior run of
 * this function — it leaves that row untouched rather than overwriting or
 * double-counting times_missed. Only genuinely new words get appended.
 */

function seedMissedWordsFromTrackers() {
  var sh = getSheet_('ReviewPool');
  var existing = sheetToObjects_(sh);
  var existingKeys = {};
  existing.forEach(function (r) {
    existingKeys[r.student + '|' + String(r.word).toLowerCase()] = true;
  });

  var headers = getHeaders_(sh);
  var toAppend = [];
  var skipped = 0;

  MISSED_WORD_TRACKER_DATA_.forEach(function (entry) {
    var key = entry.student + '|' + entry.word.toLowerCase();
    if (existingKeys[key]) { skipped++; return; }
    toAppend.push(headers.map(function (h) {
      switch (h) {
        case 'student': return entry.student;
        case 'word': return entry.word;
        case 'times_missed': return entry.timesMissed;
        case 'times_correct': return 0;
        case 'last_seen': return entry.lastSeen;
        case 'status': return 'active';
        case 'context_sentence': return entry.context || '';
        default: return '';
      }
    }));
    existingKeys[key] = true; // guard against dupes within this same data set
  });

  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, headers.length).setValues(toAppend);
  }

  var msg = 'Added ' + toAppend.length + ' missed word(s) to ReviewPool, skipped ' + skipped + ' already tracked.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* no UI when run outside the Sheet */ }
}

// word: string | lastSeen: the tracker's lesson/section label | context: optional clarifier
// pulled from parenthetical notes in the source doc (e.g. "sight (as in vision)").
var MISSED_WORD_TRACKER_DATA_ = [
  { student: 'kenley', word: 'bathtub', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'beautiful', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'cabbage', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'circus', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'comfortable', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'kenley', word: 'comparative', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'kenley', word: 'compatible', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'kenley', word: 'complacent', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'kenley', word: 'complaint', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'kenley', word: 'curious', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'cycling', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'decision', timesMissed: 1, lastSeen: 'Lesson 3', context: '' },
  { student: 'kenley', word: 'descend', timesMissed: 1, lastSeen: 'Lesson 2', context: '' },
  { student: 'kenley', word: 'descendant', timesMissed: 1, lastSeen: 'Lesson 2', context: '' },
  { student: 'kenley', word: 'detective', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'detectives', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'kenley', word: "didn't", timesMissed: 1, lastSeen: 'Lesson 5', context: 'review contractions' },
  { student: 'kenley', word: 'division', timesMissed: 1, lastSeen: 'Lesson 3', context: '' },
  { student: 'kenley', word: 'educate', timesMissed: 1, lastSeen: 'Lesson 6', context: '' },
  { student: 'kenley', word: 'excelled', timesMissed: 1, lastSeen: 'Lesson 5', context: '' },
  { student: 'kenley', word: 'fascinate', timesMissed: 1, lastSeen: 'Lesson 2', context: '' },
  { student: 'kenley', word: 'foreclosure', timesMissed: 1, lastSeen: 'Lesson 5', context: '' },
  { student: 'kenley', word: 'fossil', timesMissed: 1, lastSeen: 'Lesson 5', context: '' },
  { student: 'kenley', word: 'garbage', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'giggling', timesMissed: 1, lastSeen: 'Lesson 5', context: '' },
  { student: 'kenley', word: 'gradual', timesMissed: 1, lastSeen: 'Lesson 6', context: '' },
  { student: 'kenley', word: 'guard', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'icicle', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'individual', timesMissed: 1, lastSeen: 'Lesson 6', context: '' },
  { student: 'kenley', word: 'invasion', timesMissed: 1, lastSeen: 'Lesson 3', context: '' },
  { student: 'kenley', word: 'jealous', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'Keith', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'kitchen', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'knight', timesMissed: 1, lastSeen: 'Lesson 5', context: 'vs. night' },
  { student: 'kenley', word: 'lazy', timesMissed: 1, lastSeen: 'Lesson 2', context: '' },
  { student: 'kenley', word: 'miracle', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'miscellaneous', timesMissed: 1, lastSeen: 'Lesson 2', context: '' },
  { student: 'kenley', word: 'neighbors', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'nervous', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'kenley', word: 'nickel', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'pickers', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'kenley', word: 'plumber', timesMissed: 2, lastSeen: 'Lesson 4', context: '' },
  { student: 'kenley', word: 'raccoon', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'recommend', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'kenley', word: 'residual', timesMissed: 1, lastSeen: 'Lesson 6', context: '' },
  { student: 'kenley', word: 'schedule', timesMissed: 1, lastSeen: 'Lesson 6', context: '' },
  { student: 'kenley', word: 'scissors', timesMissed: 1, lastSeen: 'Lesson 2', context: '' },
  { student: 'kenley', word: 'Sheila', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'sight', timesMissed: 1, lastSeen: 'Lesson 1', context: 'as in vision' },
  { student: 'kenley', word: 'soap', timesMissed: 1, lastSeen: 'Lesson 2', context: '' },
  { student: 'kenley', word: 'spectacle', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'story', timesMissed: 1, lastSeen: 'Lesson 3', context: '' },
  { student: 'kenley', word: 'tipped', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'utensils', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'vegetables', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'kenley', word: 'wear', timesMissed: 1, lastSeen: 'Lesson 5', context: 'vs. where' },
  { student: 'kenley', word: 'weird', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'kenley', word: 'whispered', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'kenley', word: 'won', timesMissed: 1, lastSeen: 'Lesson 2', context: 'as in "I won a prize"' },

  { student: 'adelyn', word: 'among', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: 'attention', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: 'babies', timesMissed: 1, lastSeen: 'Lesson 7', context: '' },
  { student: 'adelyn', word: 'been', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: 'birds', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'boat', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'brown', timesMissed: 1, lastSeen: 'Lesson 8', context: '' },
  { student: 'adelyn', word: 'brownish', timesMissed: 1, lastSeen: 'Lesson 2', context: '' },
  { student: 'adelyn', word: 'bus station', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: 'busy', timesMissed: 1, lastSeen: 'Lesson 7', context: '' },
  { student: 'adelyn', word: 'checking', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'adelyn', word: 'cheerful', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'circle', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: 'city', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'cleaned', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'coats', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'cries', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'curve', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: 'December', timesMissed: 1, lastSeen: 'Lesson 8', context: '' },
  { student: 'adelyn', word: 'destroy', timesMissed: 1, lastSeen: 'Lesson 6', context: '' },
  { student: 'adelyn', word: 'different', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: 'direction', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: 'disturb', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: "don't", timesMissed: 2, lastSeen: 'Lesson 8', context: '' },
  { student: 'adelyn', word: 'down', timesMissed: 1, lastSeen: 'Lesson 2', context: '' },
  { student: 'adelyn', word: 'everything', timesMissed: 1, lastSeen: 'Lesson 6', context: '' },
  { student: 'adelyn', word: 'feed', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'flowers', timesMissed: 2, lastSeen: 'Lesson 2', context: '' },
  { student: 'adelyn', word: 'foggy', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'front', timesMissed: 1, lastSeen: 'Lesson 3', context: '' },
  { student: 'adelyn', word: 'gracefully', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: 'gym', timesMissed: 1, lastSeen: 'Lesson 6', context: 'as in gymnasium' },
  { student: 'adelyn', word: 'heard', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: 'hill', timesMissed: 1, lastSeen: 'Lesson 3', context: '' },
  { student: 'adelyn', word: 'history', timesMissed: 1, lastSeen: 'Lesson 6', context: '' },
  { student: 'adelyn', word: 'house', timesMissed: 2, lastSeen: 'Lesson 3', context: '' },
  { student: 'adelyn', word: 'how', timesMissed: 1, lastSeen: 'Lesson 7', context: '' },
  { student: 'adelyn', word: 'July', timesMissed: 1, lastSeen: 'Lesson 8', context: '' },
  { student: 'adelyn', word: 'junk', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'lawn', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'lazy', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'loud', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'misjudge', timesMissed: 1, lastSeen: 'Lesson 5', context: '' },
  { student: 'adelyn', word: 'muddy', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'needle', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'adelyn', word: 'neighbor', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: 'our', timesMissed: 1, lastSeen: 'Lesson 2', context: 'as in belonging to us' },
  { student: 'adelyn', word: 'overdue', timesMissed: 1, lastSeen: 'Lesson 5', context: '' },
  { student: 'adelyn', word: 'paint', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'paths', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'adelyn', word: 'photo', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: 'places', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'adelyn', word: 'plays', timesMissed: 1, lastSeen: 'Lesson 2', context: '' },
  { student: 'adelyn', word: 'popcorn', timesMissed: 1, lastSeen: 'Lesson 2', context: '' },
  { student: 'adelyn', word: 'presoak', timesMissed: 1, lastSeen: 'Lesson 5', context: '' },
  { student: 'adelyn', word: 'purple', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'pushed', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'put', timesMissed: 1, lastSeen: 'Lesson 2', context: '' },
  { student: 'adelyn', word: 'puzzle', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'quick', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'really', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'return', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: 'road', timesMissed: 2, lastSeen: 'Lesson 8', context: '' },
  { student: 'adelyn', word: 'runner', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'semicircle', timesMissed: 1, lastSeen: 'Lesson 5', context: '' },
  { student: 'adelyn', word: 'skirt', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'speak', timesMissed: 1, lastSeen: 'Lesson 3', context: '' },
  { student: 'adelyn', word: 'speech', timesMissed: 1, lastSeen: 'Lesson 2', context: '' },
  { student: 'adelyn', word: 'thirty', timesMissed: 2, lastSeen: 'Level 4 Book Review Words', context: '' },
  { student: 'adelyn', word: 'told', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'Tom', timesMissed: 1, lastSeen: 'Lesson 6', context: '' },
  { student: 'adelyn', word: 'too', timesMissed: 1, lastSeen: 'Lesson 7', context: 'as in "me too"' },
  { student: 'adelyn', word: 'toothless', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'trumpets', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'uncle', timesMissed: 2, lastSeen: 'Lesson 7', context: '' },
  { student: 'adelyn', word: 'very', timesMissed: 1, lastSeen: 'Lesson 8', context: '' },
  { student: 'adelyn', word: 'want', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'warm', timesMissed: 1, lastSeen: 'Lesson 1', context: '' },
  { student: 'adelyn', word: 'won', timesMissed: 1, lastSeen: 'Lesson 1', context: 'as in "the students won"' },
  { student: 'adelyn', word: 'wood', timesMissed: 1, lastSeen: 'Lesson 4', context: '' },
  { student: 'adelyn', word: 'worry', timesMissed: 1, lastSeen: 'Lesson 7', context: '' },
  { student: 'adelyn', word: 'writing', timesMissed: 1, lastSeen: 'Level 4 Book Review Words', context: '' }
];
