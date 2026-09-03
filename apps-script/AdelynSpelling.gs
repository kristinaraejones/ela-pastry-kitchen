/**
 * The ELA Pastry Kitchen — Adelyn's Spelling, All About Spelling Level 5,
 * Steps 2-28 (Weeks 1-27).
 *
 * Adds net-new Schedule rows for Adelyn's "spelling" subject, converted
 * from Adelyn_AAS_Level5_Spelling_Weeks_2-28.pdf (weekly lesson, word
 * list, homophone pairs, dictation sentences) and
 * Adelyn_AAS_Level5_Quizzes_and_Final_Exam.docx (the six 4-week checkpoint
 * quizzes plus final exam, each with its own parent answer key already
 * baked in) into the same content_json shape used for Kenley's spelling.
 * Does NOT touch setupSheets(), Month1Weeks2to5.gs, AdelynVocabReading.gs,
 * or any Kenley rows — purely additive, and only touches Adelyn's
 * placeholder "spelling" row.
 *
 * Run seedAdelynSpelling_() ONCE from the function dropdown, after
 * setupSheets() has already been run at least once. It's idempotent: if it
 * detects task_id "asf1" already in the Schedule sheet, it assumes this has
 * already run and does nothing. It also removes Adelyn's original single
 * placeholder row for "spelling" ("as1" waiting-shell), so she won't see
 * both the placeholder and the real content at once.
 *
 * Content structure per week:
 *   "This Week's Focus" (read) — that week's lesson: the sound/spelling
 *     rule, any RULE/RULE BREAKER/MEMORY TRICKS callouts, and homophone
 *     pair explanations with example sentences, all carried over from the
 *     source PDF nearly verbatim (only the paragraph breaks are
 *     re-flowed, since the PDF only preserves wrapped display lines).
 *   "Word List Dictation" (graded-dictation) — that week's new spelling
 *     words (skipped on the 5 pure-review weeks that introduce no new
 *     words: weeks 5, 9, 12, 15, 27). A word that's part of this week's
 *     homophone pair carries its example sentence as dictation context
 *     (spoken word-in-sentence-word, exactly like Kenley's hymn/him week),
 *     so the two spellings are actually distinguishable by ear.
 *   "Dictated Sentences" (graded-dictation) — the 12 provided sentences
 *     for the week, every week (review weeks included).
 *   "Review Words (Auto-Pulled From Misses)" — added ONCE, at week 1,
 *     using the same dynamic:'reviewPool' mechanic as Kenley's spelling
 *     review task. Because dynamic tasks show up regardless of the
 *     student's current week (see app.js's activeTasks), this alone
 *     satisfies "missed-word review every week" without needing to repeat
 *     the task 27 times — it pulls from the shared ReviewPool sheet,
 *     which any dictation task (in any subject) feeds automatically
 *     whenever she misses a word.
 *   Checkpoint quizzes (weeks 4, 8, 12, 16, 20, 24) and the cumulative
 *     Final Exam (week 27) — graded-dictation tasks (monthlyTest /
 *     termFinal respectively) using the FIXED word lists and answer keys
 *     from the Quizzes/Final Exam document, not a randomly-sampled bank
 *     like Kenley's spellingMonthBank/examSpellingBank. This is
 *     deliberate: Kenley's dynamic bank sampling only works because her
 *     Bank sheet is populated incrementally, in step with what's actually
 *     been taught — since Adelyn's whole year loads at once here, a
 *     dynamic bank would risk quizzing her on words from weeks she hasn't
 *     reached yet. The source document's own curated, already-scoped word
 *     lists sidestep that problem entirely and also come with a ready-made
 *     bonus dictation sentence per quiz.
 */

function seedAdelynSpelling_() {
  var sh = getSheet_('Schedule');
  var existing = sheetToObjects_(sh);
  if (existing.some(function (r) { return r.task_id === 'asf1' && r.subject_key === 'spelling'; })) {
    SpreadsheetApp.getUi().alert('Adelyn\'s Spelling content already appears to be seeded (found task asf1) — skipping to avoid duplicates.');
    return;
  }

  // Remove Adelyn's original single placeholder row for spelling
  // ("as1" waiting-shell) so it doesn't sit alongside the real content.
  var toDelete = existing.filter(function (r) {
    return r.student === 'adelyn' && r.subject_key === 'spelling' && r.task_id === 'as1' && r.label.indexOf('Waiting') !== -1;
  });
  toDelete.sort(function (a, b) { return b._row - a._row; }).forEach(function (r) { sh.deleteRow(r._row); });

  var rows = [];
  function addTasks(subjectKey, subjectName, subjectTag, weekNumber, tasks) {
    tasks.forEach(function (t) {
      var content = {};
      Object.keys(t).forEach(function (k) {
        if (['id', 'label', 'type', 'dynamic', 'termFinal', 'monthlyTest'].indexOf(k) === -1) content[k] = t[k];
      });
      rows.push([
        'adelyn', subjectKey, subjectName, subjectTag, weekNumber,
        t.id, t.type, t.label, JSON.stringify(content),
        t.dynamic || '', !!t.termFinal, !!t.monthlyTest
      ]);
    });
  }

  ADELYN_SPELLING_WEEKS.forEach(function (week) {
    Object.keys(week.subjects).forEach(function (subjectKey) {
      var subj = week.subjects[subjectKey];
      addTasks(subjectKey, subj.name, subj.tag, week.week_number, subj.tasks);
    });
  });

  var headers = SHEET_HEADERS.Schedule;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  SpreadsheetApp.getUi().alert('Added ' + rows.length + ' new Schedule rows for Adelyn — Spelling, all 27 weeks (AAS Level 5, Steps 2-28).');
}

// ---------- Content ----------

var ADELYN_SPELLING_WEEKS = [
  {
    "week_number": 1,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 2",
        "tasks": [
          {
            "id": "asf1",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>/kt/ Spelled CT</b> — Step 2.</p><p style=\"opacity:.8;\">The /kt/ sound at the end of a word, plus a look at accented syllables.</p><p>The sound /kt/ at the end of a word is spelled ct — you can hear it in select, collect, subject, and object. This week also introduces the idea of accenting a syllable. In a multisyllable word, one syllable is usually said a little louder than the others — that's the accented syllable. Some words change meaning depending on which syllable gets the accent: OB-ject (a thing) versus ob-JECT (to disagree). Words spelled the same but pronounced differently like this are called homographs — homo means 'same' and graph means 'writing.' A few words on this week's list are homographs: perfect, subject, object, and contract. Try saying each one both ways — PER-fect vs. per-FECT — and notice how the meaning shifts with the accent.</p></div>"
          },
          {
            "id": "asw1",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "select",
                "kind": "word"
              },
              {
                "answer": "perfect",
                "kind": "word"
              },
              {
                "answer": "collect",
                "kind": "word"
              },
              {
                "answer": "subject",
                "kind": "word"
              },
              {
                "answer": "respect",
                "kind": "word"
              },
              {
                "answer": "inspect",
                "kind": "word"
              },
              {
                "answer": "object",
                "kind": "word"
              },
              {
                "answer": "contract",
                "kind": "word"
              },
              {
                "answer": "direct",
                "kind": "word"
              },
              {
                "answer": "effect",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass1",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "Sam and Amy collect yellow bugs every summer.",
                "kind": "sentence"
              },
              {
                "answer": "Which prize did they select for their brother?",
                "kind": "sentence"
              },
              {
                "answer": "I knew the judge would convict the robber!",
                "kind": "sentence"
              },
              {
                "answer": "The milkman broke his contract with the farmer.",
                "kind": "sentence"
              },
              {
                "answer": "Dan chose horses as the subject of his speech.",
                "kind": "sentence"
              },
              {
                "answer": "This stuffed bear is the object of my love!",
                "kind": "sentence"
              },
              {
                "answer": "I think the teams in our district are the best!",
                "kind": "sentence"
              },
              {
                "answer": "Sue played such perfect music at the party.",
                "kind": "sentence"
              },
              {
                "answer": "The flood had a big effect on all of us.",
                "kind": "sentence"
              },
              {
                "answer": "Who will direct the play if Bob goes away?",
                "kind": "sentence"
              },
              {
                "answer": "Please inspect those flowers for ants.",
                "kind": "sentence"
              },
              {
                "answer": "These goats never show us any respect!",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "asr",
            "label": "Review Words (Auto-Pulled From Misses)",
            "type": "graded-dictation",
            "dynamic": "reviewPool",
            "prompt": "These aren't hand-picked — they're pulled automatically from words Adelyn has missed in past dictation, oldest/most-missed first. Get one right twice and it graduates out of rotation."
          }
        ]
      }
    }
  },
  {
    "week_number": 2,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 3",
        "tasks": [
          {
            "id": "asf2",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Clues for /shŭn/ Words, Part 1</b> — Step 3.</p><p style=\"opacity:.8;\">Two more clues for choosing tion vs. sion — and a closer look at tion words.</p><p>You already know that /shŭn/ at the end of a word is usually spelled tion. This week gives you two clues that help predict when to use tion.</p><p><b>CLUE #1</b></p><p>Find the base word inside the /shŭn/ word. If the base word ends in the sound /t/, the ending is spelled tion — object → objection, connect → connection, educate → education.</p><p><b>CLUE #2</b></p><p>Over 700 words end in the sound /a-shŭn/, and every one of them is spelled a-tion — information, vacation, nation, location, combination. Not every /shŭn/ word has a helpful base word — you'll learn a third clue in an upcoming lesson for words like motion, which don't end in /t/ and don't fit the a-tion pattern either.</p></div>"
          },
          {
            "id": "asw2",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "objection",
                "kind": "word"
              },
              {
                "answer": "connection",
                "kind": "word"
              },
              {
                "answer": "education",
                "kind": "word"
              },
              {
                "answer": "election",
                "kind": "word"
              },
              {
                "answer": "celebration",
                "kind": "word"
              },
              {
                "answer": "inspection",
                "kind": "word"
              },
              {
                "answer": "infection",
                "kind": "word"
              },
              {
                "answer": "protection",
                "kind": "word"
              },
              {
                "answer": "invention",
                "kind": "word"
              },
              {
                "answer": "intention",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass2",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "I have no objection to painting the barn pink.",
                "kind": "sentence"
              },
              {
                "answer": "The cat had an operation on her front paw yesterday.",
                "kind": "sentence"
              },
              {
                "answer": "We hardly need protection from the toothless old lions!",
                "kind": "sentence"
              },
              {
                "answer": "If you don't clean that cut, you'll get an infection.",
                "kind": "sentence"
              },
              {
                "answer": "The whole family voted in the election this year.",
                "kind": "sentence"
              },
              {
                "answer": "The construction of our tree house is nearing completion.",
                "kind": "sentence"
              },
              {
                "answer": "Beth showed us her gold medal at the celebration.",
                "kind": "sentence"
              },
              {
                "answer": "I must finish the inspection of the monkey cages.",
                "kind": "sentence"
              },
              {
                "answer": "A good education is so important!",
                "kind": "sentence"
              },
              {
                "answer": "Check the telephone connection before you call.",
                "kind": "sentence"
              },
              {
                "answer": "Sliced bread was a great invention.",
                "kind": "sentence"
              },
              {
                "answer": "Is it your intention to wear those itchy pants all day?",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 3,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 4",
        "tasks": [
          {
            "id": "asf3",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Words Ending in SS</b> — Step 4.</p><p style=\"opacity:.8;\">Multisyllable words ending in the sound of /s/ spelled ss.</p><p>At the end of a base word, the sound /s/ is often spelled with double s-s — you can hear it in princess, impress, confess, and discuss. This week's list also includes the contraction o'clock, which is short for 'of the clock' — the apostrophe replaces the letters that were removed (of and the f in the).</p><p><b>RULE BREAKER: BUSINESS</b></p><p>The base word is busy, with the y changed to an i — but that i is not pronounced. One trick: remember there's a 'bus' hiding in busi-ness.</p><p><b>RULE BREAKER: ANSWER</b></p><p>The w is completely silent.</p></div>"
          },
          {
            "id": "asw3",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "confess",
                "kind": "word"
              },
              {
                "answer": "discuss",
                "kind": "word"
              },
              {
                "answer": "impress",
                "kind": "word"
              },
              {
                "answer": "recess",
                "kind": "word"
              },
              {
                "answer": "address",
                "kind": "word"
              },
              {
                "answer": "actress",
                "kind": "word"
              },
              {
                "answer": "princess",
                "kind": "word"
              },
              {
                "answer": "o'clock",
                "kind": "word"
              },
              {
                "answer": "business",
                "kind": "word"
              },
              {
                "answer": "answer",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass3",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "His tall tale about the talking bunny did not impress me.",
                "kind": "sentence"
              },
              {
                "answer": "Why is your brother wearing a winter hat in July?",
                "kind": "sentence"
              },
              {
                "answer": "You have five minutes to answer our question!",
                "kind": "sentence"
              },
              {
                "answer": "I can't remember which actress played the princess.",
                "kind": "sentence"
              },
              {
                "answer": "Ted needs your address so he can reply to your kind letter.",
                "kind": "sentence"
              },
              {
                "answer": "I must confess that I don't like to eat sticky things.",
                "kind": "sentence"
              },
              {
                "answer": "My sister enjoys photographing family events.",
                "kind": "sentence"
              },
              {
                "answer": "We made a tray of fudge and cupcakes for the bake sale.",
                "kind": "sentence"
              },
              {
                "answer": "Tonight we will discuss the plans for our camping vacation.",
                "kind": "sentence"
              },
              {
                "answer": "The knights have been napping since nine o'clock.",
                "kind": "sentence"
              },
              {
                "answer": "Does Pam prefer to play ball or swing at recess?",
                "kind": "sentence"
              },
              {
                "answer": "Those bears have no business walking backward like that!",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 4,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 5",
        "tasks": [
          {
            "id": "asf4",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Clues for /shŭn/ Words, Part 2</b> — Step 5.</p><p style=\"opacity:.8;\">A third clue for /shŭn/ — plus how to stack two suffixes on one word.</p><p>Here's your third clue for spelling /shŭn/: if the base word ends in the sound /s/, the ending is spelled sion — confess → confession, discuss → discussion, impress → impression. Not every /shŭn/ word has a base word at all — motion doesn't end in /t/, doesn't end in /a-shŭn/, and has no base word, so it just has to be memorized as tion. This week also covers stacking two suffixes onto one word — like turning hope into hopeful, then into hopefully. When two suffixes are added in a row, each one keeps its own spelling, even if it creates a double letter, like the two l's in hopefully.</p><p><b>THREE CLUES FOR /SHŬN/ SO FAR</b></p><p>Base word ends in /t/ → tion. Word ends in the sound /a-shŭn/ → a-tion. Base word ends in /s/ → sion.</p><p><b>🔀 Homophone Pairs</b></p><p><i>whether vs. weather</i> — Whether introduces a choice between options. Weather refers to what's happening outside (rain, sun, etc.) — a word you learned back in Level Four.</p><p style=\"opacity:.8;margin-left:12px;\">whether: \"She goes running whether there's rain or sun. I don't know whether or not to believe her!\"</p><p style=\"opacity:.8;margin-left:12px;\">weather: \"Wendy never listens to the weather report. She told me that she likes all types of weather.\"</p></div>"
          },
          {
            "id": "asw4",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "confession",
                "kind": "word"
              },
              {
                "answer": "discussion",
                "kind": "word"
              },
              {
                "answer": "impression",
                "kind": "word"
              },
              {
                "answer": "recession",
                "kind": "word"
              },
              {
                "answer": "clothing",
                "kind": "word"
              },
              {
                "answer": "farther",
                "kind": "word"
              },
              {
                "answer": "therefore",
                "kind": "word"
              },
              {
                "answer": "rather",
                "kind": "word"
              },
              {
                "answer": "whether",
                "kind": "word",
                "context": "She goes running whether there's rain or sun. I don't know whether or not to believe her!"
              },
              {
                "answer": "further",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass4",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "We hoped to invent a robot that could fold shirts.",
                "kind": "sentence"
              },
              {
                "answer": "Would you rather go to parties or to school?",
                "kind": "sentence"
              },
              {
                "answer": "I get the impression that Amy prefers blue clothing.",
                "kind": "sentence"
              },
              {
                "answer": "Dave can ride his bike farther up the hill than I can.",
                "kind": "sentence"
              },
              {
                "answer": "Her uncles do not know whether to go fishing or sailing.",
                "kind": "sentence"
              },
              {
                "answer": "What is that furry thing under the front porch?",
                "kind": "sentence"
              },
              {
                "answer": "We giggled when the dogs got their tails stuck in the fence.",
                "kind": "sentence"
              },
              {
                "answer": "The queen would not write a confession and thus was set free.",
                "kind": "sentence"
              },
              {
                "answer": "You will have to study further if you want to pass the test.",
                "kind": "sentence"
              },
              {
                "answer": "It is hard to find jobs when the nation is in a recession.",
                "kind": "sentence"
              },
              {
                "answer": "The bug was intent on swinging from one leaf to another.",
                "kind": "sentence"
              },
              {
                "answer": "We had a long discussion about those five funny skunks.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "asq1",
            "label": "Spelling Quiz 1 — Weeks 1–4",
            "type": "graded-dictation",
            "monthlyTest": true,
            "prompt": "A quick, low-stakes checkpoint on words already taught in this block — not a re-test of every word.",
            "words": [
              {
                "answer": "select",
                "kind": "word"
              },
              {
                "answer": "object",
                "kind": "word"
              },
              {
                "answer": "contract",
                "kind": "word"
              },
              {
                "answer": "connection",
                "kind": "word"
              },
              {
                "answer": "education",
                "kind": "word"
              },
              {
                "answer": "celebration",
                "kind": "word"
              },
              {
                "answer": "princess",
                "kind": "word"
              },
              {
                "answer": "business",
                "kind": "word"
              },
              {
                "answer": "answer",
                "kind": "word"
              },
              {
                "answer": "confession",
                "kind": "word"
              },
              {
                "answer": "The inspection of the princess's mansion revealed no objection to the celebration.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 5,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 6",
        "tasks": [
          {
            "id": "asf5",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Ways to Spell /shŭn/</b> — Step 6.</p><p style=\"opacity:.8;\">A review of the two most common ways to spell /shŭn/.</p><p>This week is a review, pulling together what you've learned about spelling /shŭn/. Two spellings cover the vast majority of words: tion and sion. Try sorting these words by ending: completion, protection, impression, recession, confession, connection, operation, election, infection, discussion, education, invention. Say each word aloud, listen for /shŭn/, and think about which of your three clues from the last two weeks applies.</p><p><b>REMEMBER</b></p><p>tion is used far more often than sion — when you're truly unsure and none of the clues apply, tion is the better guess.</p></div>"
          },
          {
            "id": "ass5",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "Ben munched on a plateful of fried clams.",
                "kind": "sentence"
              },
              {
                "answer": "If you connect all the dots you will see a short knight.",
                "kind": "sentence"
              },
              {
                "answer": "How many families will be at the lunch on Saturday?",
                "kind": "sentence"
              },
              {
                "answer": "The wind blew wildly as the stormy night raged on.",
                "kind": "sentence"
              },
              {
                "answer": "We were very pleased to meet your pet muskrat yesterday.",
                "kind": "sentence"
              },
              {
                "answer": "This is the crunchiest apple pie I have ever had.",
                "kind": "sentence"
              },
              {
                "answer": "The sly horse tugged on the rotting rope until he broke free.",
                "kind": "sentence"
              },
              {
                "answer": "Our neighbors gave us a selection of fish and shrimp.",
                "kind": "sentence"
              },
              {
                "answer": "Have you ever cooked anything over an open fire?",
                "kind": "sentence"
              },
              {
                "answer": "Beth tested the water with her toe before jumping in.",
                "kind": "sentence"
              },
              {
                "answer": "Those sharks completely destroyed our raft!",
                "kind": "sentence"
              },
              {
                "answer": "You'll have to rewash the cat if she naps in the mud again.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 6,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 7",
        "tasks": [
          {
            "id": "asf6",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>/eks/ and /egz/ Spelled EX</b> — Step 7.</p><p style=\"opacity:.8;\">At the start of a word, ex can say /eks/ or /egz/.</p><p>When you hear the sound /eks/ at the beginning of a word, it's spelled ex — listen for it in expense, extreme, and express. Sometimes that same ex spelling sounds more like /egz/ instead — you can hear this softer sound in example, exist, and exam.</p><p><b>RULE</b></p><p>Same spelling, two sounds: ex at the start of a word can say either /eks/ or /egz/ depending on the word — there's no way to tell just by looking, so it helps to get familiar with both groups of words.</p></div>"
          },
          {
            "id": "asw6",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "except",
                "kind": "word"
              },
              {
                "answer": "expense",
                "kind": "word"
              },
              {
                "answer": "extreme",
                "kind": "word"
              },
              {
                "answer": "extra",
                "kind": "word"
              },
              {
                "answer": "express",
                "kind": "word"
              },
              {
                "answer": "expect",
                "kind": "word"
              },
              {
                "answer": "exit",
                "kind": "word"
              },
              {
                "answer": "example",
                "kind": "word"
              },
              {
                "answer": "exam",
                "kind": "word"
              },
              {
                "answer": "exist",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass6",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "Ed had eighty extra electric ice makers.",
                "kind": "sentence"
              },
              {
                "answer": "This extreme heat makes my hair stand up in all directions.",
                "kind": "sentence"
              },
              {
                "answer": "How many bubbles do you expect me to blow in one hour?",
                "kind": "sentence"
              },
              {
                "answer": "An unending ball of yellow yarn is exactly what I wanted!",
                "kind": "sentence"
              },
              {
                "answer": "Ken is very helpful and sets a good example for his little brother.",
                "kind": "sentence"
              },
              {
                "answer": "A mother bear will do anything to protect her cubs.",
                "kind": "sentence"
              },
              {
                "answer": "Did you pass your exam about the different types of worms?",
                "kind": "sentence"
              },
              {
                "answer": "Please exit the room before you infect us all with your germs!",
                "kind": "sentence"
              },
              {
                "answer": "Sam sobbed silently every day except Thursday.",
                "kind": "sentence"
              },
              {
                "answer": "I've learned that bugs will exist whether I like it or not!",
                "kind": "sentence"
              },
              {
                "answer": "It was quite an expense to take all of our sheep on vacation.",
                "kind": "sentence"
              },
              {
                "answer": "Kim prefers to express herself with paintings of lovely gardens.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 7,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 8",
        "tasks": [
          {
            "id": "asf7",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>The Sound of /ŭff/ Spelled OUGH</b> — Step 8.</p><p style=\"opacity:.8;\">One new phonogram (ough) and how to spell the names of holidays.</p><p>The letters ough are unusual — this one spelling can represent six different sounds, but it only shows up in about twenty-five words total, so it's easiest to just get familiar with them by sound group. This week focuses on one of those sounds: /ŭff/, as in rough, tough, and enough. A helpful sentence to remember them: 'Is it rough and tough enough?' This week also covers spelling the names of holidays.</p><p><b>RULE</b></p><p>The names of holidays are always capitalized. New Year's Day needs an apostrophe-s, because the Day belongs to the New Year.</p><p><b>RULE BREAKER: CHRISTMAS</b></p><p>The t is silent — even though the holiday is named after 'Christ,' the t in Christmas is never pronounced.</p></div>"
          },
          {
            "id": "asw7",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "rough",
                "kind": "word"
              },
              {
                "answer": "tough",
                "kind": "word"
              },
              {
                "answer": "enough",
                "kind": "word"
              },
              {
                "answer": "holiday",
                "kind": "word"
              },
              {
                "answer": "New Year's Day",
                "kind": "word"
              },
              {
                "answer": "Memorial Day",
                "kind": "word"
              },
              {
                "answer": "Independence Day",
                "kind": "word"
              },
              {
                "answer": "Columbus Day",
                "kind": "word"
              },
              {
                "answer": "Thanksgiving",
                "kind": "word"
              },
              {
                "answer": "Christmas",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass7",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "Our holiday events are mostly loud and joyful!",
                "kind": "sentence"
              },
              {
                "answer": "Let's string popcorn to hang on the Christmas tree.",
                "kind": "sentence"
              },
              {
                "answer": "We don't have enough candles for her birthday cake.",
                "kind": "sentence"
              },
              {
                "answer": "Does your mother have to work at the bank on Columbus Day?",
                "kind": "sentence"
              },
              {
                "answer": "Who did they elect to head the Independence Day celebrations?",
                "kind": "sentence"
              },
              {
                "answer": "A hundred friendly ducks dashed out of the woods.",
                "kind": "sentence"
              },
              {
                "answer": "Our Thanksgiving turkey was tough, but the plum pie was good.",
                "kind": "sentence"
              },
              {
                "answer": "We weighed the roundest pigs on New Year's Day.",
                "kind": "sentence"
              },
              {
                "answer": "The cheerful bride wed her husband on Memorial Day.",
                "kind": "sentence"
              },
              {
                "answer": "A dozen people crept across the bridge in the semidarkness.",
                "kind": "sentence"
              },
              {
                "answer": "She grabbed an armful of scratchy fabric and threw it all away.",
                "kind": "sentence"
              },
              {
                "answer": "The clawless hawk had a rough time hanging onto the branch.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 8,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 9",
        "tasks": [
          {
            "id": "asf8",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>OR in Unaccented Syllables</b> — Step 9.</p><p style=\"opacity:.8;\">or can spell /er/ in an unaccented syllable.</p><p>You already know or can say /or/, as in corn. But in an unaccented syllable — one that's said a little softer — or can also spell the muffled /er/ sound, like in flavor and works. This pattern shows up often in job words: inventor, author, doctor, mayor, and actor all end this way. Other examples you'll see later include director, editor, and navigator.</p><p><b>RULE</b></p><p>or says /or/ in an accented syllable, but softens to /er/ in an unaccented syllable.</p><p><b>🔀 Homophone Pairs</b></p><p><i>mayor vs. mare</i> — A mayor is the elected leader of a town or city. A mare is a female horse.</p><p style=\"opacity:.8;margin-left:12px;\">mayor: \"Our town mayor is shaped like a pear. Yes, the mayor of our town is a funny guy!\"</p><p style=\"opacity:.8;margin-left:12px;\">mare: \"He wanders the streets on a sweet old mare. We throw the mare apples as she trots by.\"</p></div>"
          },
          {
            "id": "asw8",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "inventor",
                "kind": "word"
              },
              {
                "answer": "odor",
                "kind": "word"
              },
              {
                "answer": "flavor",
                "kind": "word"
              },
              {
                "answer": "author",
                "kind": "word"
              },
              {
                "answer": "actor",
                "kind": "word"
              },
              {
                "answer": "doctor",
                "kind": "word"
              },
              {
                "answer": "comfort",
                "kind": "word"
              },
              {
                "answer": "mayor",
                "kind": "word",
                "context": "Our town mayor is shaped like a pear. Yes, the mayor of our town is a funny guy!"
              },
              {
                "answer": "effort",
                "kind": "word"
              },
              {
                "answer": "color",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass8",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "The odor of garlic in the kitchen just knocked me down!",
                "kind": "sentence"
              },
              {
                "answer": "The people threw roses at the actor as he left the stage.",
                "kind": "sentence"
              },
              {
                "answer": "Uncle Dan is the author of seventeen books about puzzles.",
                "kind": "sentence"
              },
              {
                "answer": "Who was the inventor of the first telephone?",
                "kind": "sentence"
              },
              {
                "answer": "My doctor said I had the largest ears of all the kids in town.",
                "kind": "sentence"
              },
              {
                "answer": "The smallest crow crunched loudly on a crispy leaf.",
                "kind": "sentence"
              },
              {
                "answer": "What flavor is the cake you're baking for us tonight?",
                "kind": "sentence"
              },
              {
                "answer": "There is nothing like the comfort of an old pair of socks!",
                "kind": "sentence"
              },
              {
                "answer": "Who did the mayor put in charge of the city gardens?",
                "kind": "sentence"
              },
              {
                "answer": "Frank put little effort into fixing that plate of fudge!",
                "kind": "sentence"
              },
              {
                "answer": "What color are the candies in your backpack?",
                "kind": "sentence"
              },
              {
                "answer": "The owl blinked and shook the morning dew from his wings.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "asq2",
            "label": "Spelling Quiz 2 — Weeks 6–8",
            "type": "graded-dictation",
            "monthlyTest": true,
            "prompt": "A quick, low-stakes checkpoint on words already taught in this block — not a re-test of every word.",
            "words": [
              {
                "answer": "expect",
                "kind": "word"
              },
              {
                "answer": "example",
                "kind": "word"
              },
              {
                "answer": "exist",
                "kind": "word"
              },
              {
                "answer": "enough",
                "kind": "word"
              },
              {
                "answer": "tough",
                "kind": "word"
              },
              {
                "answer": "Christmas",
                "kind": "word"
              },
              {
                "answer": "inventor",
                "kind": "word"
              },
              {
                "answer": "author",
                "kind": "word"
              },
              {
                "answer": "mayor",
                "kind": "word"
              },
              {
                "answer": "flavor",
                "kind": "word"
              },
              {
                "answer": "The author didn't expect the inventor's tough new example to exist by Christmas.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 9,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 10",
        "tasks": [
          {
            "id": "asf9",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Ways to Spell /er/</b> — Step 10.</p><p style=\"opacity:.8;\">A review of five ways to spell the /er/ sound.</p><p>This week reviews everything you know so far about spelling the /er/ sound. There are five spellings to keep straight: er, ur, ir, or, and ear. Try sorting these words by spelling: flavor, letter, heard, curve, shirt, prefer, thirty, work, turn, serve, summer, early, skirt, world, mayor, church, learn, dirt, offer, inventor, hurt, worst, search, birthday, father. Say each word aloud, listen for /er/, and think about which spelling it uses.</p></div>"
          },
          {
            "id": "ass9",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "I wish the pretty princess would extend an invitation to us.",
                "kind": "sentence"
              },
              {
                "answer": "Why are half the hens hopping around on one foot?",
                "kind": "sentence"
              },
              {
                "answer": "We are extremely upset about the letters we found in the attic.",
                "kind": "sentence"
              },
              {
                "answer": "The water was dirty and therefore not very healthy.",
                "kind": "sentence"
              },
              {
                "answer": "What caused all that giggling I heard in class today?",
                "kind": "sentence"
              },
              {
                "answer": "The puppy planted a big wet kiss on my cheek!",
                "kind": "sentence"
              },
              {
                "answer": "She passed the afternoon joyfully among good friends.",
                "kind": "sentence"
              },
              {
                "answer": "Who sent me a box of thorny stems with no flowers on the ends?",
                "kind": "sentence"
              },
              {
                "answer": "Mr. March is deciding whether or not to give us the examination.",
                "kind": "sentence"
              },
              {
                "answer": "I knew it was you who crashed into my cow!",
                "kind": "sentence"
              },
              {
                "answer": "I think his grassy lawn is a little patch of perfection!",
                "kind": "sentence"
              },
              {
                "answer": "The sly spy slipped into the room unseen by the lady in white.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 10,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 11",
        "tasks": [
          {
            "id": "asf10",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>I-Before-E Generalization, Part 1</b> — Step 11.</p><p style=\"opacity:.8;\">A classic spelling poem: i before e, when the sound is /ē/.</p><p>Here's a poem that will help you remember one of the most famous spelling patterns in English: When the sound is /ē/, It's i before e Except after c. This week covers the first part of that poem — words where the sound of /ē/ is spelled ie, like in field, chief, believe, relief, and piece.</p><p><b>RULE</b></p><p>When you hear the /ē/ sound and need to choose between ie and ei, the letter i comes before e — unless the letter right before it is a c (that part of the poem comes into play in the next lesson). Keep in mind there are a few exceptions to this poem that you'll learn later, after getting plenty of practice with words that do follow the pattern.</p><p><b>🔀 Homophone Pairs</b></p><p><i>piece vs. peace</i> — A piece is a part or portion of something. Peace means calm, or the absence of conflict.</p><p style=\"opacity:.8;margin-left:12px;\">piece: \"She used a piece of chocolate for the design. Then she gave one piece of cake to everyone in our family.\"</p><p style=\"opacity:.8;margin-left:12px;\">peace: \"Alice decorated the cake with doves, the symbol of peace. And now we can all have a bit of peace!\"</p></div>"
          },
          {
            "id": "asw10",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "field",
                "kind": "word"
              },
              {
                "answer": "chief",
                "kind": "word"
              },
              {
                "answer": "believe",
                "kind": "word"
              },
              {
                "answer": "relief",
                "kind": "word"
              },
              {
                "answer": "piece",
                "kind": "word",
                "context": "She used a piece of chocolate for the design. Then she gave one piece of cake to everyone in our family."
              },
              {
                "answer": "niece",
                "kind": "word"
              },
              {
                "answer": "achieve",
                "kind": "word"
              },
              {
                "answer": "yield",
                "kind": "word"
              },
              {
                "answer": "thief",
                "kind": "word"
              },
              {
                "answer": "shield",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass10",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "We spent hours grinding grain in the field.",
                "kind": "sentence"
              },
              {
                "answer": "I believe the judges rated my jokes unfairly.",
                "kind": "sentence"
              },
              {
                "answer": "Just how many drumsticks her niece ate is unknown.",
                "kind": "sentence"
              },
              {
                "answer": "Do you remember the exact words the chief whispered to you?",
                "kind": "sentence"
              },
              {
                "answer": "This mix will yield about a dozen moist cupcakes.",
                "kind": "sentence"
              },
              {
                "answer": "Use this sheet to shield yourself from the muddy ground.",
                "kind": "sentence"
              },
              {
                "answer": "We felt such relief when we found our lost stork in the hedge!",
                "kind": "sentence"
              },
              {
                "answer": "The little thief took twenty tires and rolled them all away.",
                "kind": "sentence"
              },
              {
                "answer": "How did that piece of purple cloth wind up in the ditch?",
                "kind": "sentence"
              },
              {
                "answer": "Beth works hard because she wants to achieve a lot in her life.",
                "kind": "sentence"
              },
              {
                "answer": "Black birds sat on the benches and boldly begged for bread.",
                "kind": "sentence"
              },
              {
                "answer": "How did he light that flame with those wet matches?",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 11,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 12",
        "tasks": [
          {
            "id": "asf11",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>I-Before-E Generalization, Part 2</b> — Step 12.</p><p style=\"opacity:.8;\">The second half of the poem: ei right after the letter c.</p><p>Remember the poem from last week? This week covers its last line: except after c. Right after the letter c, the /ē/ sound flips to ei instead of ie — you can hear this in receive and ceiling. This week also reviews the jobs of Silent E, since several of this week's words end in it: relative, prove, engine, promise, move, service, and debate each use Silent E for a different reason — to keep the word from ending in v, to make a vowel long, as a 'Handyman E' that doesn't fit another category, or to make a c or g soft.</p><p><b>RULE</b></p><p>i before e, except after c — right after a c, flip the order to ei.</p><p><b>MEMORY TRICKS</b></p><p>Sometimes a memory trick (also called a mnemonic) is the easiest way to remember a tricky spelling. For receipt, remember: when you pay, you get a receipt — both start with p. For business, remember there's a bus hiding inside: busi-ness.</p><p><b>RULE BREAKER: RECEIPT</b></p><p>The p is silent.</p></div>"
          },
          {
            "id": "asw11",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "receive",
                "kind": "word"
              },
              {
                "answer": "ceiling",
                "kind": "word"
              },
              {
                "answer": "receipt",
                "kind": "word"
              },
              {
                "answer": "relative",
                "kind": "word"
              },
              {
                "answer": "prove",
                "kind": "word"
              },
              {
                "answer": "engine",
                "kind": "word"
              },
              {
                "answer": "promise",
                "kind": "word"
              },
              {
                "answer": "move",
                "kind": "word"
              },
              {
                "answer": "service",
                "kind": "word"
              },
              {
                "answer": "debate",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass11",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "How did your footprints get on the ceiling?",
                "kind": "sentence"
              },
              {
                "answer": "It does not matter much if we win or lose the debate.",
                "kind": "sentence"
              },
              {
                "answer": "We were served a flavorful lunch of fried oysters.",
                "kind": "sentence"
              },
              {
                "answer": "Can you prove that this stack of quarters belongs to you?",
                "kind": "sentence"
              },
              {
                "answer": "I did not expect to receive such great service and attention!",
                "kind": "sentence"
              },
              {
                "answer": "We paid a large amount for that pony but didn't get a receipt.",
                "kind": "sentence"
              },
              {
                "answer": "Is dear Mrs. Maple a relative of yours?",
                "kind": "sentence"
              },
              {
                "answer": "Be careful when you move that priceless piece of art!",
                "kind": "sentence"
              },
              {
                "answer": "Eighty itchy inventors scratched their heads all afternoon.",
                "kind": "sentence"
              },
              {
                "answer": "Why is Amy brushing her teeth with a paintbrush?",
                "kind": "sentence"
              },
              {
                "answer": "I ride an elk, so I never worry that my engine will fail.",
                "kind": "sentence"
              },
              {
                "answer": "Did Pam keep her promise to feed the pigs pancakes?",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 12,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 13",
        "tasks": [
          {
            "id": "asf12",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Ways to Spell /ē/</b> — Step 13.</p><p style=\"opacity:.8;\">A review of nine ways to spell the /ē/ sound.</p><p>This week reviews the many ways to spell /ē/ — quite a few, since it's one of the most common vowel sounds in English! The nine spellings are: e, e-e, ee, ea, y, i, ey, ie, and ei. Try sorting these words by spelling: monkey, street, cheek, real, Steve, these, niece, money, seed, copier, receipt, lucky, east, field, beef, penny, speak, honey, here, valley, even, year, she, puppy, happier, because, shield. Say each word aloud, listen for /ē/, and think about which spelling it uses.</p></div>"
          },
          {
            "id": "ass12",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "I find it very funny that the ox is oinking.",
                "kind": "sentence"
              },
              {
                "answer": "We watched the ducks dashing madly across the yard.",
                "kind": "sentence"
              },
              {
                "answer": "The child had a cheek full of chewy candy.",
                "kind": "sentence"
              },
              {
                "answer": "I think we should throw away all that junky stuff in our attic.",
                "kind": "sentence"
              },
              {
                "answer": "Did you know that eating too much fast food is unhealthy?",
                "kind": "sentence"
              },
              {
                "answer": "My brother unlocked the barn and pushed the colt inside.",
                "kind": "sentence"
              },
              {
                "answer": "April looked quite ill after spinning around like a top.",
                "kind": "sentence"
              },
              {
                "answer": "Rick held the smelly fish right under my nose.",
                "kind": "sentence"
              },
              {
                "answer": "Your knees seem to get dirtier by the second!",
                "kind": "sentence"
              },
              {
                "answer": "Her perfect words worked like magic to make us all feel better!",
                "kind": "sentence"
              },
              {
                "answer": "Can you believe how many snails came out after the storm?",
                "kind": "sentence"
              },
              {
                "answer": "We met an odorless skunk while walking along the path.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "asq3",
            "label": "Spelling Quiz 3 — Weeks 10–12",
            "type": "graded-dictation",
            "monthlyTest": true,
            "prompt": "A quick, low-stakes checkpoint on words already taught in this block — not a re-test of every word.",
            "words": [
              {
                "answer": "believe",
                "kind": "word"
              },
              {
                "answer": "field",
                "kind": "word"
              },
              {
                "answer": "thief",
                "kind": "word"
              },
              {
                "answer": "receive",
                "kind": "word"
              },
              {
                "answer": "ceiling",
                "kind": "word"
              },
              {
                "answer": "receipt",
                "kind": "word"
              },
              {
                "answer": "promise",
                "kind": "word"
              },
              {
                "answer": "engine",
                "kind": "word"
              },
              {
                "answer": "service",
                "kind": "word"
              },
              {
                "answer": "debate",
                "kind": "word"
              },
              {
                "answer": "I believe the thief in the field will receive a receipt for his promise.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 13,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 14",
        "tasks": [
          {
            "id": "asf13",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>More Words with Silent E</b> — Step 14.</p><p style=\"opacity:.8;\">Reviewing the jobs Silent E does, with a set of trickier words.</p><p>This week is all about practicing what you already know about Silent E, using some longer, trickier words: sentence, possible, square, change, article, escape, divide, arrange, entire, and estate. For each word, try to name the job Silent E is doing — is it making a vowel long, making a c or g soft, adding a vowel to a syllable, or acting as a 'Handyman E' that doesn't fit any other category? In a word like arrange, the a in the middle is in an unaccented syllable, so it gets muffled toward an /ŭ/ sound. Saying the word slowly and clearly — ah-RANGE — can help you hear which letter it should be.</p></div>"
          },
          {
            "id": "asw13",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "sentence",
                "kind": "word"
              },
              {
                "answer": "possible",
                "kind": "word"
              },
              {
                "answer": "square",
                "kind": "word"
              },
              {
                "answer": "change",
                "kind": "word"
              },
              {
                "answer": "article",
                "kind": "word"
              },
              {
                "answer": "escape",
                "kind": "word"
              },
              {
                "answer": "divide",
                "kind": "word"
              },
              {
                "answer": "arrange",
                "kind": "word"
              },
              {
                "answer": "entire",
                "kind": "word"
              },
              {
                "answer": "estate",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass13",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "The spoiled queen lined up her invitations and smiled happily.",
                "kind": "sentence"
              },
              {
                "answer": "We have been unable to arrange a meeting for Tuesday morning.",
                "kind": "sentence"
              },
              {
                "answer": "Did you reread the article about that acting school?",
                "kind": "sentence"
              },
              {
                "answer": "Would it be possible to dig a huge hole with a small spoon?",
                "kind": "sentence"
              },
              {
                "answer": "Shall I declare my love for you now or at seven o'clock?",
                "kind": "sentence"
              },
              {
                "answer": "The second sentence of the story stopped me in my tracks.",
                "kind": "sentence"
              },
              {
                "answer": "The look in her eyes seems to change with every passing minute.",
                "kind": "sentence"
              },
              {
                "answer": "Mr. Swift sold me square screws instead of round ones!",
                "kind": "sentence"
              },
              {
                "answer": "The old man left his entire estate to his mule.",
                "kind": "sentence"
              },
              {
                "answer": "The convict planned his great escape in the dead of night.",
                "kind": "sentence"
              },
              {
                "answer": "Hopefully this tale will end with a sweet surprise.",
                "kind": "sentence"
              },
              {
                "answer": "Will Mike divide this line of mice into five nice piles?",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 14,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 15",
        "tasks": [
          {
            "id": "asf14",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>More Words with /z/ Spelled S</b> — Step 15.</p><p style=\"opacity:.8;\">No rule tells you s vs. z for /z/ — so you turn to other strategies.</p><p>The /z/ sound in the middle of a word can be spelled with either s or z, and there's no rule that tells you which one to use. When you're not sure, two strategies can help: try writing the word both ways and see which one 'looks right' (Scratch Paper Spelling), or look the word up in a dictionary.</p><p><b>HELPFUL PATTERN</b></p><p>When s sits between two vowels, or right before an m, it often says /z/ instead of /s/. Since s is the more common spelling for /z/ overall, it's usually worth trying an s first.</p><p><b>🔀 Homophone Pairs</b></p><p><i>raise vs. rays</i> — Raise means to lift up or to bring up (like raising animals). Rays are beams of light.</p><p style=\"opacity:.8;margin-left:12px;\">raise: \"Denny and Pam have decided to raise a family of bats. \"Those bats raise the hair on my head,\" said their mother.\"</p><p style=\"opacity:.8;margin-left:12px;\">rays: \"They must keep their bats away from the sun's rays. \"Why? They don't shoot poison rays from their eyes!\" they replied.\"</p></div>"
          },
          {
            "id": "asw14",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "present",
                "kind": "word"
              },
              {
                "answer": "desire",
                "kind": "word"
              },
              {
                "answer": "lose",
                "kind": "word"
              },
              {
                "answer": "season",
                "kind": "word"
              },
              {
                "answer": "visitor",
                "kind": "word"
              },
              {
                "answer": "raise",
                "kind": "word",
                "context": "Denny and Pam have decided to raise a family of bats. \"Those bats raise the hair on my head,\" said their mother."
              },
              {
                "answer": "reason",
                "kind": "word"
              },
              {
                "answer": "president",
                "kind": "word"
              },
              {
                "answer": "closet",
                "kind": "word"
              },
              {
                "answer": "thousand",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass14",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "Deb has no desire to dig for worms after dark.",
                "kind": "sentence"
              },
              {
                "answer": "The visitor was lost in the museum for over six hours.",
                "kind": "sentence"
              },
              {
                "answer": "I know you have a good reason to be dressed like a crab.",
                "kind": "sentence"
              },
              {
                "answer": "How many matches did your team lose last season?",
                "kind": "sentence"
              },
              {
                "answer": "Rick just can't resist those greenish sticks of gum.",
                "kind": "sentence"
              },
              {
                "answer": "The president decided to close the biggest prison in the state.",
                "kind": "sentence"
              },
              {
                "answer": "Frank heard a noise in the closet and dove under the covers.",
                "kind": "sentence"
              },
              {
                "answer": "Did she deserve to win the starring role in the play?",
                "kind": "sentence"
              },
              {
                "answer": "We were present when the thousand knights bravely rode to war.",
                "kind": "sentence"
              },
              {
                "answer": "Jumping in a muddy puddle may result in dirty feet.",
                "kind": "sentence"
              },
              {
                "answer": "The milkman and his wife raise mostly cows and monkeys.",
                "kind": "sentence"
              },
              {
                "answer": "I suppose I could enclose those forty fields of yellow roses.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 15,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 16",
        "tasks": [
          {
            "id": "asf15",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Make It Plural Book</b> — Step 16.</p><p style=\"opacity:.8;\">Reviewing four patterns for making words plural.</p><p>This week reviews four patterns you already know for making a word plural: adding s (farmers, wings, mayors), adding es after certain endings (houses, faces, wires), changing y to i and adding es (armies, replies, cities), and a few other familiar patterns (candies, pennies, flies). There are no new spelling words this week — instead, it's a good chance to notice these patterns in words you already know how to spell.</p></div>"
          },
          {
            "id": "ass15",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "The children continued to imprison the bugs in glass jars.",
                "kind": "sentence"
              },
              {
                "answer": "Please don't misunderstand what I'm trying to tell you!",
                "kind": "sentence"
              },
              {
                "answer": "The wild waves of the sea came crashing down on the rocks.",
                "kind": "sentence"
              },
              {
                "answer": "We must fix the broken chimney bricks before Christmas.",
                "kind": "sentence"
              },
              {
                "answer": "They finished testing the telephones on Wednesday.",
                "kind": "sentence"
              },
              {
                "answer": "Mrs. White photographed the mayor for the newspaper article.",
                "kind": "sentence"
              },
              {
                "answer": "The inventor pulled on the brake to stop his flying bike.",
                "kind": "sentence"
              },
              {
                "answer": "It was funny to see the cats dart about when I honked my horn.",
                "kind": "sentence"
              },
              {
                "answer": "We could hear the noise of the engines from ten miles away.",
                "kind": "sentence"
              },
              {
                "answer": "I'm not sure why she used her teeth to trim the fabric edges.",
                "kind": "sentence"
              },
              {
                "answer": "That square object in the closet is exactly the color of mold.",
                "kind": "sentence"
              },
              {
                "answer": "The doctor drifted into a ditch one clear dawn in December.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 16,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 17",
        "tasks": [
          {
            "id": "asf16",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Plurals of Words Ending in F and FE</b> — Step 17.</p><p style=\"opacity:.8;\">Many words ending in f or fe change to v before adding the plural.</p><p>For some words ending in f or f-e, making the word plural just means adding s — roof becomes roofs. But for about half of these words, something unusual happens: the f changes to a v, and then you add es — knife becomes knives, not 'knifes.' You can hear the new /v/ sound clearly when you say the plural aloud.</p><p><b>RULE</b></p><p>For roughly half of all words ending in f or fe, change the f to v and add es: calf → calves, self → selves, knife → knives, wife → wives, leaf → leaves, thief → thieves, loaf → loaves. For the other half, just add s, as in roofs. A couple of words can actually go either way — both hoofs/hooves and scarfs/scarves are considered correct.</p></div>"
          },
          {
            "id": "asw16",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "calves",
                "kind": "word"
              },
              {
                "answer": "selves",
                "kind": "word"
              },
              {
                "answer": "shelves",
                "kind": "word"
              },
              {
                "answer": "lives",
                "kind": "word"
              },
              {
                "answer": "knives",
                "kind": "word"
              },
              {
                "answer": "wives",
                "kind": "word"
              },
              {
                "answer": "wolves",
                "kind": "word"
              },
              {
                "answer": "leaves",
                "kind": "word"
              },
              {
                "answer": "thieves",
                "kind": "word"
              },
              {
                "answer": "loaves",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass16",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "Her shelves are full of torn clothes and unmatched socks.",
                "kind": "sentence"
              },
              {
                "answer": "Let's shed our old selves and try to become better people.",
                "kind": "sentence"
              },
              {
                "answer": "A hundred husbands and wives circled the parking lot.",
                "kind": "sentence"
              },
              {
                "answer": "They spent their whole lives caring for sick and lost puppies.",
                "kind": "sentence"
              },
              {
                "answer": "The cook used his sharpest knives to slice the tough beef.",
                "kind": "sentence"
              },
              {
                "answer": "What lovely white wolves we saw on the snowy hills!",
                "kind": "sentence"
              },
              {
                "answer": "Their intention is to keep their inventions to themselves.",
                "kind": "sentence"
              },
              {
                "answer": "We silently raked the leaves under a sinking sun.",
                "kind": "sentence"
              },
              {
                "answer": "Aunt April easily ate eleven loaves of nut bread.",
                "kind": "sentence"
              },
              {
                "answer": "Why is the garden hose all tied up in knots?",
                "kind": "sentence"
              },
              {
                "answer": "Calves pass their time charging around and chewing grass.",
                "kind": "sentence"
              },
              {
                "answer": "Each of the thieves had weak knees and tiny feet.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "asq4",
            "label": "Spelling Quiz 4 — Weeks 13–16",
            "type": "graded-dictation",
            "monthlyTest": true,
            "prompt": "A quick, low-stakes checkpoint on words already taught in this block — not a re-test of every word.",
            "words": [
              {
                "answer": "present",
                "kind": "word"
              },
              {
                "answer": "thousand",
                "kind": "word"
              },
              {
                "answer": "president",
                "kind": "word"
              },
              {
                "answer": "sentence",
                "kind": "word"
              },
              {
                "answer": "possible",
                "kind": "word"
              },
              {
                "answer": "escape",
                "kind": "word"
              },
              {
                "answer": "knives",
                "kind": "word"
              },
              {
                "answer": "wolves",
                "kind": "word"
              },
              {
                "answer": "leaves",
                "kind": "word"
              },
              {
                "answer": "thieves",
                "kind": "word"
              },
              {
                "answer": "Is it possible that a thousand wolves could escape into the falling leaves?",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 17,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 18",
        "tasks": [
          {
            "id": "asf17",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Plurals of Words Ending in O</b> — Step 18.</p><p style=\"opacity:.8;\">Four new letter teams, plus how to pluralize words ending in o.</p><p>This week introduces four new letter teams: ui (says /o͞o/, as in fruit), gn (says /n/, used at the beginning or end of a word), our (says /er/, as in journey), and ci (says /sh/ — since c is the shorter letter, it's nicknamed 'short-letter / sh/,' compared to ti, 'tall-letter /sh/'). For pluralizing words ending in o: if the word ends in a vowel plus o, just add s — radio becomes radios. If the word ends in a consonant plus o, you'll usually need to check a dictionary, since some take s (pianos) and others take es (heroes, potatoes, tomatoes).</p><p><b>RULE</b></p><p>Vowel + o → add s (radios, videos, patios). Consonant + o → check the dictionary; many take es (heroes, potatoes), though some take just s (pianos, autos).</p><p><b>RULE BREAKER: LISTEN</b></p><p>The t is completely silent — this happens in a few other words too, where a base word is pronounced differently once it becomes part of a longer word (child → children, cloth → clothing, know → knowledge).</p></div>"
          },
          {
            "id": "asw17",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "radios",
                "kind": "word"
              },
              {
                "answer": "potatoes",
                "kind": "word"
              },
              {
                "answer": "tomatoes",
                "kind": "word"
              },
              {
                "answer": "heroes",
                "kind": "word"
              },
              {
                "answer": "pianos",
                "kind": "word"
              },
              {
                "answer": "tornadoes",
                "kind": "word"
              },
              {
                "answer": "volcanoes",
                "kind": "word"
              },
              {
                "answer": "solos",
                "kind": "word"
              },
              {
                "answer": "rodeos",
                "kind": "word"
              },
              {
                "answer": "listen",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass17",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "Twenty-two tornadoes swept swiftly across the nation.",
                "kind": "sentence"
              },
              {
                "answer": "Ron gave plastic radios to all the kids in the neighborhood.",
                "kind": "sentence"
              },
              {
                "answer": "Every year we plant and pick our own potatoes.",
                "kind": "sentence"
              },
              {
                "answer": "We made working volcanoes out of paper and glue.",
                "kind": "sentence"
              },
              {
                "answer": "Those tigers have long claws and noses as red as tomatoes.",
                "kind": "sentence"
              },
              {
                "answer": "The president pinned badges on the shirts of the heroes.",
                "kind": "sentence"
              },
              {
                "answer": "Tuning pianos and teaching music is our family business.",
                "kind": "sentence"
              },
              {
                "answer": "Sam played three trumpet solos at the event last night.",
                "kind": "sentence"
              },
              {
                "answer": "We take our horses to ten different rodeos each summer.",
                "kind": "sentence"
              },
              {
                "answer": "Some of the quickest birds can run fifteen miles an hour.",
                "kind": "sentence"
              },
              {
                "answer": "Our goats escaped from the barn and fled to the cliffs.",
                "kind": "sentence"
              },
              {
                "answer": "If you listen to the river you'll hear music in the water.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 18,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 19",
        "tasks": [
          {
            "id": "asf18",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Words Ending in /ij/</b> — Step 19.</p><p style=\"opacity:.8;\">The ending /ij/ is spelled age, plus a look at irregular plurals.</p><p>At the end of a word, the sound /ij/ is spelled a-g-e — listen for it in postage, garbage, message, and package. In everyday speech, part of this ending often gets swallowed — average is really pronounced 'av-er-age,' and pronouncing it that way (Pronounce for Spelling) can help you remember every syllable. This week also covers irregular plurals — a small set of words that don't follow any of the regular patterns you've learned. You already know most of them by ear: the plural of mouse is mice, not 'mouses.' The plural of child is children, and the plural of goose is geese.</p><p><b>RULE</b></p><p>The ending sound /ij/ is spelled age.</p><p><b>IRREGULAR PLURALS</b></p><p>Most words form the plural using regular patterns, but a handful are irregular and simply need to be memorized: child → children, goose → geese, man → men, woman → women, tooth → teeth, foot → feet.</p></div>"
          },
          {
            "id": "asw18",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "postage",
                "kind": "word"
              },
              {
                "answer": "average",
                "kind": "word"
              },
              {
                "answer": "garbage",
                "kind": "word"
              },
              {
                "answer": "message",
                "kind": "word"
              },
              {
                "answer": "package",
                "kind": "word"
              },
              {
                "answer": "mileage",
                "kind": "word"
              },
              {
                "answer": "damage",
                "kind": "word"
              },
              {
                "answer": "manage",
                "kind": "word"
              },
              {
                "answer": "children",
                "kind": "word"
              },
              {
                "answer": "geese",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass18",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "Did he put enough postage on the package before he mailed it?",
                "kind": "sentence"
              },
              {
                "answer": "The students are writing an average of eight pages a day.",
                "kind": "sentence"
              },
              {
                "answer": "Those ungraceful geese have been in the garbage again!",
                "kind": "sentence"
              },
              {
                "answer": "A man dressed in black inspected the message for clues.",
                "kind": "sentence"
              },
              {
                "answer": "I don't see the advantage of saying the alphabet backward.",
                "kind": "sentence"
              },
              {
                "answer": "Do your trucks get better mileage in the city?",
                "kind": "sentence"
              },
              {
                "answer": "If you ask me, this kitchen has a real shortage of cupcakes!",
                "kind": "sentence"
              },
              {
                "answer": "We heard that the whales caused a lot of damage to the ship.",
                "kind": "sentence"
              },
              {
                "answer": "The cheerful nurses manage to keep all the children smiling.",
                "kind": "sentence"
              },
              {
                "answer": "The hostage was taken from his hotel and put in prison.",
                "kind": "sentence"
              },
              {
                "answer": "We live in the largest stone cottage in the village.",
                "kind": "sentence"
              },
              {
                "answer": "I baked a clam and cabbage pie to celebrate your achievement!",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 19,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 20",
        "tasks": [
          {
            "id": "asf19",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>The Sound of /o͞o/ Spelled UI</b> — Step 20.</p><p style=\"opacity:.8;\">ui spells /o͞o/, plus words with matching singular and plural forms.</p><p>The letters ui spell the sound /o͞o/ — listen for it in juice, suit, fruit, and bruise. A few words in English stay exactly the same whether you're talking about one or many — the plural of sheep is still sheep, and the same goes for deer, moose, pants, shrimp, and elk.</p><p><b>RULE</b></p><p>ui spells the sound /o͞o/.</p><p><b>RULE BREAKER: OFTEN</b></p><p>The t is completely silent.</p><p><b>🔀 Homophone Pairs</b></p><p><i>bruise vs. brews</i> — A bruise is a dark mark left on skin after an injury. Brews means makes tea or coffee (or steers/handles, in some uses).</p><p style=\"opacity:.8;margin-left:12px;\">bruise: \"Bertha once tripped on a rope and got a bruise on her knee.\"</p><p style=\"opacity:.8;margin-left:12px;\">brews: \"She loves the funny captain, who brews tea in his cabin.\"</p><p><i>cruise vs. crews</i> — A cruise is a trip on a ship. Crews are teams of people working together (like a ship's crew).</p><p style=\"opacity:.8;margin-left:12px;\">cruise: \"Bertha goes on a cruise to Greece on the same ship every year.\"</p><p style=\"opacity:.8;margin-left:12px;\">crews: \"Sometimes, however, the deck crews are a little sloppy.\"</p></div>"
          },
          {
            "id": "asw19",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "fruit",
                "kind": "word"
              },
              {
                "answer": "bruise",
                "kind": "word",
                "context": "Bertha once tripped on a rope and got a bruise on her knee."
              },
              {
                "answer": "cruise",
                "kind": "word",
                "context": "Bertha goes on a cruise to Greece on the same ship every year."
              },
              {
                "answer": "juice",
                "kind": "word"
              },
              {
                "answer": "suit",
                "kind": "word"
              },
              {
                "answer": "grapefruit",
                "kind": "word"
              },
              {
                "answer": "lawsuit",
                "kind": "word"
              },
              {
                "answer": "suitcase",
                "kind": "word"
              },
              {
                "answer": "recruit",
                "kind": "word"
              },
              {
                "answer": "often",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass19",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "My uncle often wears his spacesuit to dinner.",
                "kind": "sentence"
              },
              {
                "answer": "Fruit fell from the trees and landed with a thud.",
                "kind": "sentence"
              },
              {
                "answer": "Mark put on his new blue suit and walked along the avenue.",
                "kind": "sentence"
              },
              {
                "answer": "The baby made a face when she ate that piece of grapefruit.",
                "kind": "sentence"
              },
              {
                "answer": "We are thinking about taking a long cruise next winter.",
                "kind": "sentence"
              },
              {
                "answer": "How did you get that bruise on your knee?",
                "kind": "sentence"
              },
              {
                "answer": "Ted is trying to recruit someone to complete the hardest tasks.",
                "kind": "sentence"
              },
              {
                "answer": "The judge threw out our lawsuit and told us to go home.",
                "kind": "sentence"
              },
              {
                "answer": "My mother is undecided about the color for the bedrooms.",
                "kind": "sentence"
              },
              {
                "answer": "His suitcase sank in the sea like a sack of potatoes.",
                "kind": "sentence"
              },
              {
                "answer": "I'd offer you a glass of juice but all I have is water.",
                "kind": "sentence"
              },
              {
                "answer": "Where did you hide the fruitcake I gave you last Thanksgiving?",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 20,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 21",
        "tasks": [
          {
            "id": "asf20",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>The Sound of /n/ Spelled GN</b> — Step 21.</p><p style=\"opacity:.8;\">Another way to spell /n/: gn, usually at the start of a word.</p><p>You know /n/ can be spelled n or two-letter kn. This week adds a third spelling: gn, as in sign, gnat, gnarled, and campaign.</p><p><b>RULE BREAKER: SANDWICH</b></p><p>After a short vowel, /ch/ is usually spelled tch, but sandwich breaks that pattern — the /ch/ sound here is just spelled ch.</p></div>"
          },
          {
            "id": "asw20",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "sign",
                "kind": "word"
              },
              {
                "answer": "gnaw",
                "kind": "word"
              },
              {
                "answer": "gnat",
                "kind": "word"
              },
              {
                "answer": "design",
                "kind": "word"
              },
              {
                "answer": "assign",
                "kind": "word"
              },
              {
                "answer": "resign",
                "kind": "word"
              },
              {
                "answer": "assignment",
                "kind": "word"
              },
              {
                "answer": "campaign",
                "kind": "word"
              },
              {
                "answer": "gnarled",
                "kind": "word"
              },
              {
                "answer": "sandwich",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass20",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "Would you design a fine sign that I could call mine?",
                "kind": "sentence"
              },
              {
                "answer": "The hawks flew down and grabbed the sandwich from my hands.",
                "kind": "sentence"
              },
              {
                "answer": "Our assignment is to rearrange the items in the closet.",
                "kind": "sentence"
              },
              {
                "answer": "Did you listen to his entire campaign speech?",
                "kind": "sentence"
              },
              {
                "answer": "I'm only half mad that the gnat is nesting in my hat.",
                "kind": "sentence"
              },
              {
                "answer": "I knew a gnu who grew too tall and died too soon.",
                "kind": "sentence"
              },
              {
                "answer": "The articles said the mayor will resign at the end of the week.",
                "kind": "sentence"
              },
              {
                "answer": "We saw the ox gnaw on a block of rotting wood.",
                "kind": "sentence"
              },
              {
                "answer": "This study proves that raising wolves in the house is a bad plan.",
                "kind": "sentence"
              },
              {
                "answer": "Your job is to assign rooms to the visitors at all of our hotels.",
                "kind": "sentence"
              },
              {
                "answer": "The unhelpful gnome threw grapefruit at my relatives.",
                "kind": "sentence"
              },
              {
                "answer": "The bat hung from a gnarled branch and looked at me all night.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "asq5",
            "label": "Spelling Quiz 5 — Weeks 17–20",
            "type": "graded-dictation",
            "monthlyTest": true,
            "prompt": "A quick, low-stakes checkpoint on words already taught in this block — not a re-test of every word.",
            "words": [
              {
                "answer": "tornadoes",
                "kind": "word"
              },
              {
                "answer": "heroes",
                "kind": "word"
              },
              {
                "answer": "potatoes",
                "kind": "word"
              },
              {
                "answer": "listen",
                "kind": "word"
              },
              {
                "answer": "average",
                "kind": "word"
              },
              {
                "answer": "package",
                "kind": "word"
              },
              {
                "answer": "message",
                "kind": "word"
              },
              {
                "answer": "fruit",
                "kind": "word"
              },
              {
                "answer": "suitcase",
                "kind": "word"
              },
              {
                "answer": "sign",
                "kind": "word"
              },
              {
                "answer": "Please listen for the sign that the package with the message has arrived.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 21,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 22",
        "tasks": [
          {
            "id": "asf21",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>The Doubling Rule</b> — Step 22.</p><p style=\"opacity:.8;\">A new rule for doubling the final consonant on multisyllable words.</p><p>You already know the 1-1-1 Rule for one-syllable words: if a word has one syllable, one vowel, and one consonant at the end, you double that consonant before adding a vowel suffix (as in win → winning). This week extends that idea to two-syllable words.</p><p><b>THE DOUBLING RULE</b></p><p>For a multisyllable word, ask two questions about the base word: Does the last syllable end in one vowel followed by one consonant? Is the accent on the last syllable? If the answer to both questions is yes, double the last letter before adding a vowel suffix — begin → beginning. If either answer is no, don't double it.</p><p><b>WORKED EXAMPLES</b></p><p>offer + ed → offered (no double — the accent isn't on the last syllable). admit + ing → admitting (double the t — both answers are yes). disturb + ing → disturbing (no double — doesn't end in one vowel + one consonant). prefer + ed → preferred (double the r). return + ed → returned (no double — doesn't end in one vowel + one consonant). excel + ing → excelling (double the l).</p></div>"
          },
          {
            "id": "asw21",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "answered",
                "kind": "word"
              },
              {
                "answer": "referring",
                "kind": "word"
              },
              {
                "answer": "disturbing",
                "kind": "word"
              },
              {
                "answer": "beginning",
                "kind": "word"
              },
              {
                "answer": "preferred",
                "kind": "word"
              },
              {
                "answer": "forgetting",
                "kind": "word"
              },
              {
                "answer": "returned",
                "kind": "word"
              },
              {
                "answer": "admitting",
                "kind": "word"
              },
              {
                "answer": "excelled",
                "kind": "word"
              },
              {
                "answer": "offered",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass21",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "Are we disturbing you by playing our trumpets and pianos?",
                "kind": "sentence"
              },
              {
                "answer": "He answered our questions in a whisper we could hardly hear.",
                "kind": "sentence"
              },
              {
                "answer": "Amy preferred the flavor of honey to that of unripe plums.",
                "kind": "sentence"
              },
              {
                "answer": "Dan keeps forgetting to divide the geese from the hens.",
                "kind": "sentence"
              },
              {
                "answer": "The color of his swimsuit changes when he goes in the water.",
                "kind": "sentence"
              },
              {
                "answer": "Ron selected thicker covers to keep the cows warm.",
                "kind": "sentence"
              },
              {
                "answer": "Are you referring to those calves who like fruit juice and grapes?",
                "kind": "sentence"
              },
              {
                "answer": "The doll that was once loved now lay forgotten in the attic.",
                "kind": "sentence"
              },
              {
                "answer": "Pam excelled at gardening and often offered us fresh tomatoes.",
                "kind": "sentence"
              },
              {
                "answer": "The farmer returned nightly to his cold cabin across the river.",
                "kind": "sentence"
              },
              {
                "answer": "Is Frank admitting that he failed to fix fifty fences in the field?",
                "kind": "sentence"
              },
              {
                "answer": "I'm beginning to understand why my relatives never visit.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 22,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 23",
        "tasks": [
          {
            "id": "asf22",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>AR in Unaccented Syllables</b> — Step 23.</p><p style=\"opacity:.8;\">ar can spell the muffled /er/ sound in an unaccented syllable.</p><p>You know ar says /ar/, as in far. But in an unaccented syllable, that sound gets muffled toward /er/ instead — you can hear this in dollar, calendar, grammar, and collar.</p><p><b>RULE</b></p><p>ar says /ar/ in an accented syllable, but softens toward /er/ in an unaccented syllable — similar to what you learned about or a few lessons ago.</p><p><b>RULE BREAKER: SUGAR</b></p><p>The s says /sh/ instead of its usual /s/ sound. Only two words in English do this: sugar and sure (which you learned back in Level Four).</p><p><b>🔀 Homophone Pairs</b></p><p><i>mustard vs. mustered</i> — Mustard is the yellow condiment. Mustered means gathered or summoned up (courage, troops, etc.).</p><p style=\"opacity:.8;margin-left:12px;\">mustard: \"Sergeant Sam sadly caressed the mustard stain on his uniform. Oh, how to tell them there's no more mustard in the mess hall?\"</p><p style=\"opacity:.8;margin-left:12px;\">mustered: \"Tooting his trumpet, he mustered his troops in the misty dawn. From deep in his soul, Sam mustered the courage he needed.\"</p></div>"
          },
          {
            "id": "asw22",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "dollar",
                "kind": "word"
              },
              {
                "answer": "calendar",
                "kind": "word"
              },
              {
                "answer": "grammar",
                "kind": "word"
              },
              {
                "answer": "collar",
                "kind": "word"
              },
              {
                "answer": "lizard",
                "kind": "word"
              },
              {
                "answer": "mustard",
                "kind": "word",
                "context": "Sergeant Sam sadly caressed the mustard stain on his uniform. Oh, how to tell them there's no more mustard in the mess hall?"
              },
              {
                "answer": "popular",
                "kind": "word"
              },
              {
                "answer": "circular",
                "kind": "word"
              },
              {
                "answer": "particular",
                "kind": "word"
              },
              {
                "answer": "sugar",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass22",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "We earned one dollar for every pearl we pried from an oyster.",
                "kind": "sentence"
              },
              {
                "answer": "Hopefully this circular path will lead us to the right address.",
                "kind": "sentence"
              },
              {
                "answer": "Ed unglued the stamps from all the damaged packaging.",
                "kind": "sentence"
              },
              {
                "answer": "The numbering of the dates on this calendar is all wrong!",
                "kind": "sentence"
              },
              {
                "answer": "Grammar is such an interesting subject to study.",
                "kind": "sentence"
              },
              {
                "answer": "He grabbed my collar and demanded I hand him the mustard.",
                "kind": "sentence"
              },
              {
                "answer": "The beginner made an effortless and graceful turn on her toes.",
                "kind": "sentence"
              },
              {
                "answer": "Who is preventing that popular song from being on the radio?",
                "kind": "sentence"
              },
              {
                "answer": "Are you sure she wants sugar in her shellfish stew?",
                "kind": "sentence"
              },
              {
                "answer": "These particular prizes are worth their weight in gold.",
                "kind": "sentence"
              },
              {
                "answer": "Why is Aunt Amy inspecting that handful of acorns?",
                "kind": "sentence"
              },
              {
                "answer": "Mr. Zip is a lazy lizard who lives in a log by the lake.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 23,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 24",
        "tasks": [
          {
            "id": "asf23",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>/awt/ Spelled OUGHT and A Followed by L</b> — Step 24.</p><p style=\"opacity:.8;\">Another ough sound, plus how a often changes before l.</p><p>Here's another sound for the tricky ough letter team: /awt/, spelled ough-t, as in fought, bought, ought, and thought. This week also looks at what happens when a is followed by an l. Sometimes a keeps a sound close to its third sound (as in always), but often the l shifts it into the 'all' sound you hear in also, almost, already, bald, and false.</p><p><b>RULE</b></p><p>Listen for the /aw/-type sound whenever a is followed by l — it's a signal that the vowel sound is shifting away from its usual short or long sound.</p><p><b>🔀 Homophone Pairs</b></p><p><i>already vs. all ready</i> — Already (one word) means 'by this time.' All ready (two words) means 'completely prepared.'</p><p style=\"opacity:.8;margin-left:12px;\">already: \"\"I'm bringing the baby over already,\" my sister said cheerfully.\"</p><p style=\"opacity:.8;margin-left:12px;\">all ready: \"I was dead tired and all ready for a nap when the telephone rang.\"</p><p><i>bald vs. balled / bawled</i> — Bald means having no hair. Balled means formed into a ball; bawled means cried loudly.</p><p style=\"opacity:.8;margin-left:12px;\">bald: \"Ten minutes later I had a small bald bundle to baby-sit.\"</p><p style=\"opacity:.8;margin-left:12px;\">balled / bawled: \"I balled up my fists in frustration as I stomped toward the phone. She's cute, but she bawled all day, and I never got my nap!\"</p></div>"
          },
          {
            "id": "asw23",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "fought",
                "kind": "word"
              },
              {
                "answer": "bought",
                "kind": "word"
              },
              {
                "answer": "ought",
                "kind": "word"
              },
              {
                "answer": "thought",
                "kind": "word"
              },
              {
                "answer": "brought",
                "kind": "word"
              },
              {
                "answer": "also",
                "kind": "word"
              },
              {
                "answer": "almost",
                "kind": "word"
              },
              {
                "answer": "already",
                "kind": "word",
                "context": "\"I'm bringing the baby over already,\" my sister said cheerfully."
              },
              {
                "answer": "bald",
                "kind": "word",
                "context": "Ten minutes later I had a small bald bundle to baby-sit."
              },
              {
                "answer": "false",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass23",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "I thought you might like this ham and pickle sandwich.",
                "kind": "sentence"
              },
              {
                "answer": "Do you have the receipts for all the suits you bought?",
                "kind": "sentence"
              },
              {
                "answer": "She spent many moonless nights crying into her pillow.",
                "kind": "sentence"
              },
              {
                "answer": "Our neighbors ought to trim their hedges once in a while.",
                "kind": "sentence"
              },
              {
                "answer": "Fifty men and women bravely fought the flames.",
                "kind": "sentence"
              },
              {
                "answer": "A bunch of bald babies blew bubbles on the bridge.",
                "kind": "sentence"
              },
              {
                "answer": "That particular author also wrote a few works of fiction.",
                "kind": "sentence"
              },
              {
                "answer": "We almost forgot to hang the sign above the candy store.",
                "kind": "sentence"
              },
              {
                "answer": "Are you already repacking the suitcases for your flight home?",
                "kind": "sentence"
              },
              {
                "answer": "The ship sailed silently under a starless sky.",
                "kind": "sentence"
              },
              {
                "answer": "The princess did not expect to have so many false friends.",
                "kind": "sentence"
              },
              {
                "answer": "I confess it was Sue who brought the gnu on the cruise.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 24,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 25",
        "tasks": [
          {
            "id": "asf24",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>More Words with Long E Spelled EA</b> — Step 25.</p><p style=\"opacity:.8;\">More practice with ea spelling the /ē/ sound.</p><p>This week adds more words to your /ē/-spelled-ea list: least, dream, appear, peach, increase, eagle, steam, meat, and beauty. For beauty, try over-emphasizing each syllable as you say it — 'bee-YOU-tee' — to help you remember all the letters.</p><p><b>RULE BREAKER: WHOLE</b></p><p>The wh says /h/ here instead of its usual /hw/ sound.</p><p><b>🔀 Homophone Pairs</b></p><p><i>least vs. leased</i> — Least means the smallest amount. Leased means rented.</p><p style=\"opacity:.8;margin-left:12px;\">least: \"His home may be small, but at least it's full of good friends!\"</p><p style=\"opacity:.8;margin-left:12px;\">leased: \"After hunting around, he leased one from his cousin Harvey.\"</p><p><i>meat vs. meet</i> — Meat is food from an animal. Meet means to come together with someone.</p><p style=\"opacity:.8;margin-left:12px;\">meat: \"Rufus served carrots and weeds, but certainly no meat!\"</p><p style=\"opacity:.8;margin-left:12px;\">meet: \"Then he invited all of his friends to meet him there for a party.\"</p><p><i>whole vs. hole</i> — Whole means complete or entire. A hole is an opening or gap.</p><p style=\"opacity:.8;margin-left:12px;\">whole: \"He decorated his whole home with garlands of grass and daisies.\"</p><p style=\"opacity:.8;margin-left:12px;\">hole: \"Rufus Rabbit decided to move into a brand new rabbit hole.\"</p></div>"
          },
          {
            "id": "asw24",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "least",
                "kind": "word",
                "context": "His home may be small, but at least it's full of good friends!"
              },
              {
                "answer": "dream",
                "kind": "word"
              },
              {
                "answer": "appear",
                "kind": "word"
              },
              {
                "answer": "peach",
                "kind": "word"
              },
              {
                "answer": "increase",
                "kind": "word"
              },
              {
                "answer": "eagle",
                "kind": "word"
              },
              {
                "answer": "steam",
                "kind": "word"
              },
              {
                "answer": "meat",
                "kind": "word",
                "context": "Rufus served carrots and weeds, but certainly no meat!"
              },
              {
                "answer": "beauty",
                "kind": "word"
              },
              {
                "answer": "whole",
                "kind": "word",
                "context": "He decorated his whole home with garlands of grass and daisies."
              }
            ]
          },
          {
            "id": "ass24",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "The gnat fought the fly for that juicy piece of meat pie.",
                "kind": "sentence"
              },
              {
                "answer": "I promise to send you a message at least once a week.",
                "kind": "sentence"
              },
              {
                "answer": "Why do I always dream about a gnome in a spacesuit?",
                "kind": "sentence"
              },
              {
                "answer": "These fruits appear to be turning blacker by the hour.",
                "kind": "sentence"
              },
              {
                "answer": "Are you cooking the whole turkey or just a quarter of it?",
                "kind": "sentence"
              },
              {
                "answer": "June spent nearly an hour selecting that pair of jeans.",
                "kind": "sentence"
              },
              {
                "answer": "You need to increase the amount of salt in the rice.",
                "kind": "sentence"
              },
              {
                "answer": "We saw the bald eagle nesting high on the cliffs.",
                "kind": "sentence"
              },
              {
                "answer": "Do you believe that beauty is only skin deep?",
                "kind": "sentence"
              },
              {
                "answer": "Wavy lines of steam rose from the burning pavement.",
                "kind": "sentence"
              },
              {
                "answer": "I fear the thief has stolen our only ear of corn.",
                "kind": "sentence"
              },
              {
                "answer": "Shall we eat this perfect peach or save it for the pie?",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "asq6",
            "label": "Spelling Quiz 6 — Weeks 21–24",
            "type": "graded-dictation",
            "monthlyTest": true,
            "prompt": "A quick, low-stakes checkpoint on words already taught in this block — not a re-test of every word.",
            "words": [
              {
                "answer": "beginning",
                "kind": "word"
              },
              {
                "answer": "preferred",
                "kind": "word"
              },
              {
                "answer": "referring",
                "kind": "word"
              },
              {
                "answer": "grammar",
                "kind": "word"
              },
              {
                "answer": "calendar",
                "kind": "word"
              },
              {
                "answer": "sugar",
                "kind": "word"
              },
              {
                "answer": "thought",
                "kind": "word"
              },
              {
                "answer": "already",
                "kind": "word"
              },
              {
                "answer": "beauty",
                "kind": "word"
              },
              {
                "answer": "increase",
                "kind": "word"
              },
              {
                "answer": "I thought grammar was already the beginning of a beauty of a subject, not a chore.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 25,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 26",
        "tasks": [
          {
            "id": "asf25",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>The Sound of /ŭ/ Spelled OU</b> — Step 26.</p><p style=\"opacity:.8;\">ou can spell the short-u sound /ŭ/.</p><p>You know ou can spell /ow/, as in out. This week shows another sound it can make: /ŭ/, as in touch, double, famous, and country.</p><p><b>RULE BREAKER: BUY</b></p><p>The u is completely silent.</p><p><b>RULE BREAKER: BUILD</b></p><p>The u is completely silent here too.</p><p><b>🔀 Homophone Pairs</b></p><p><i>buy vs. by / bye</i> — Buy means to purchase something. By means near or through; bye is a casual goodbye.</p><p style=\"opacity:.8;margin-left:12px;\">buy: \"Harold and I went to town to buy a big lizard. Tonight I'd like to buy ninety-nine entirely white lights.\"</p><p style=\"opacity:.8;margin-left:12px;\">by / bye: \"One day as our neighbor ran by, the lizard jumped onto her head. Now when she sees us she doesn't even say \"hi\" or \"bye\"!\"</p><p><i>build vs. billed</i> — Build means to construct something. Billed means charged or invoiced.</p><p style=\"opacity:.8;margin-left:12px;\">build: \"Then we got enough wood to build a tree house for our pet.\"</p><p style=\"opacity:.8;margin-left:12px;\">billed: \"She billed us for the damage to her fashionable hat.\"</p></div>"
          },
          {
            "id": "asw25",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "young",
                "kind": "word"
              },
              {
                "answer": "trouble",
                "kind": "word"
              },
              {
                "answer": "touch",
                "kind": "word"
              },
              {
                "answer": "country",
                "kind": "word"
              },
              {
                "answer": "couple",
                "kind": "word"
              },
              {
                "answer": "double",
                "kind": "word"
              },
              {
                "answer": "cousin",
                "kind": "word"
              },
              {
                "answer": "famous",
                "kind": "word"
              },
              {
                "answer": "buy",
                "kind": "word",
                "context": "Harold and I went to town to buy a big lizard. Tonight I'd like to buy ninety-nine entirely white lights."
              },
              {
                "answer": "build",
                "kind": "word",
                "context": "Then we got enough wood to build a tree house for our pet."
              }
            ]
          },
          {
            "id": "ass25",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "I was stung by a young bee while returning from the pond.",
                "kind": "sentence"
              },
              {
                "answer": "The children are demanding new equipment to build their robots.",
                "kind": "sentence"
              },
              {
                "answer": "Tonight I'd like to buy ninety-nine entirely white lights.",
                "kind": "sentence"
              },
              {
                "answer": "Did those sheep cause such trouble all by themselves?",
                "kind": "sentence"
              },
              {
                "answer": "Please use your magic touch to arrange these pretty flowers.",
                "kind": "sentence"
              },
              {
                "answer": "We took some photos of that row of neat country houses.",
                "kind": "sentence"
              },
              {
                "answer": "My cousin made a double batch of sugary frosted cupcakes.",
                "kind": "sentence"
              },
              {
                "answer": "Any mention of the French elections was unintended.",
                "kind": "sentence"
              },
              {
                "answer": "Fred is famous for visiting Frank every Friday at five.",
                "kind": "sentence"
              },
              {
                "answer": "Who knows why Kim talked like a fox and walked like a lizard?",
                "kind": "sentence"
              },
              {
                "answer": "I'm tired of answering questions about things that don't matter!",
                "kind": "sentence"
              },
              {
                "answer": "A couple of cute cows crept carefully across the carpet.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 26,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 27",
        "tasks": [
          {
            "id": "asf26",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>The /er/ of Journey</b> — Step 27.</p><p style=\"opacity:.8;\">Another spelling for /er/: our, plus the rare oo sound in door.</p><p>This week adds one more spelling to your collection for the /er/ sound: our, as in journey, courage, journal, and nourish. It also introduces a rare sound for oo: /ō/, as in door and floor. This is one of the least common patterns you'll learn — only door, floor, and brooch use oo this way.</p></div>"
          },
          {
            "id": "asw26",
            "label": "Word List Dictation",
            "type": "graded-dictation",
            "prompt": "Tap 🔊 to hear each word — no text shown, same as reading it aloud from a card.",
            "words": [
              {
                "answer": "journey",
                "kind": "word"
              },
              {
                "answer": "courage",
                "kind": "word"
              },
              {
                "answer": "journal",
                "kind": "word"
              },
              {
                "answer": "nourish",
                "kind": "word"
              },
              {
                "answer": "tourist",
                "kind": "word"
              },
              {
                "answer": "glamour",
                "kind": "word"
              },
              {
                "answer": "encourage",
                "kind": "word"
              },
              {
                "answer": "discourage",
                "kind": "word"
              },
              {
                "answer": "door",
                "kind": "word"
              },
              {
                "answer": "floor",
                "kind": "word"
              }
            ]
          },
          {
            "id": "ass26",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "Our journey led us over rocky hilltops and into green valleys.",
                "kind": "sentence"
              },
              {
                "answer": "The glamour of the busy city did not impress us.",
                "kind": "sentence"
              },
              {
                "answer": "Why won't Jim sketch owls or crows in his journal?",
                "kind": "sentence"
              },
              {
                "answer": "These gentle rains will nourish the beautiful ferns and flowers.",
                "kind": "sentence"
              },
              {
                "answer": "The thankful tourist trusted Tom to throw her trunk on the train.",
                "kind": "sentence"
              },
              {
                "answer": "We would never discourage you from starring in rodeos!",
                "kind": "sentence"
              },
              {
                "answer": "My dream is to design a door made of gnarled wood.",
                "kind": "sentence"
              },
              {
                "answer": "I thought they fought at dawn one day but I was really wrong.",
                "kind": "sentence"
              },
              {
                "answer": "The number of bald eagles appears to be increasing.",
                "kind": "sentence"
              },
              {
                "answer": "Why must she encourage Aunt Sue to sleep in the extra bathtub?",
                "kind": "sentence"
              },
              {
                "answer": "How did you manage to spill that sticky fruit juice on the floor?",
                "kind": "sentence"
              },
              {
                "answer": "I just didn't have the courage to eat twelve pieces of toast.",
                "kind": "sentence"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 27,
    "subjects": {
      "spelling": {
        "name": "Spelling",
        "tag": "AAS Level 5 · Step 28",
        "tasks": [
          {
            "id": "asf27",
            "label": "This Week's Focus",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Ways to Spell /er/ — Final Lesson of Level 5</b> — Step 28.</p><p style=\"opacity:.8;\">A full review of all six ways to spell the /er/ sound.</p><p>This final lesson pulls together everything you've learned about spelling /er/ across all of Level 5. There are six spellings in total: er, ur, ir, or, ear, and our. Try sorting these words by spelling: together, stir, prefer, turn, third, turkey, dirt, word, ruler, purple, power, inventor, labor, cover, direction, circle, worst, mayor, slower, firm, curve, brother, busier, thirteen, Thursday.</p><p><b>A SENTENCE TO REMEMBER THEM BY</b></p><p>Here's a sentence that uses all six common spellings of /er/: Her nurse first works early on the journey.</p></div>"
          },
          {
            "id": "ass27",
            "label": "Dictated Sentences",
            "type": "graded-dictation",
            "prompt": "All 12 sentences from this week's dictation set.",
            "words": [
              {
                "answer": "We closed the door to escape the odor of that smelly fish!",
                "kind": "sentence"
              },
              {
                "answer": "You must resign yourself to the fact that gnomes do not exist.",
                "kind": "sentence"
              },
              {
                "answer": "The builder bought enough oak and maple to finish the ceiling.",
                "kind": "sentence"
              },
              {
                "answer": "Can we discuss the problem of your mismatched clothes?",
                "kind": "sentence"
              },
              {
                "answer": "Let's move to a country cottage far away from traffic and noise.",
                "kind": "sentence"
              },
              {
                "answer": "They are offering their visitors a beautiful room in a windmill.",
                "kind": "sentence"
              },
              {
                "answer": "We thanked him for defending his friends from the wild wolves.",
                "kind": "sentence"
              },
              {
                "answer": "The ship tipped and flipped the dripping tourists into the sea.",
                "kind": "sentence"
              },
              {
                "answer": "My doctor might have done that operation differently.",
                "kind": "sentence"
              },
              {
                "answer": "You ought to buy a couple of geese to honk at the neighbors.",
                "kind": "sentence"
              },
              {
                "answer": "Amy always has exact change to buy food at the zoo.",
                "kind": "sentence"
              },
              {
                "answer": "Is it too much trouble to keep the muskrats off the carpets?",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "asfinal",
            "label": "Final Exam — All About Spelling Level 5",
            "type": "graded-dictation",
            "termFinal": true,
            "prompt": "Part A checks the last new words of the level (Steps 26–28). Part B is a cumulative review of the 20 words from across all of Level 5 that trip up students most often.",
            "words": [
              {
                "answer": "young",
                "kind": "word"
              },
              {
                "answer": "trouble",
                "kind": "word"
              },
              {
                "answer": "touch",
                "kind": "word"
              },
              {
                "answer": "country",
                "kind": "word"
              },
              {
                "answer": "cousin",
                "kind": "word"
              },
              {
                "answer": "famous",
                "kind": "word"
              },
              {
                "answer": "buy",
                "kind": "word"
              },
              {
                "answer": "build",
                "kind": "word"
              },
              {
                "answer": "journey",
                "kind": "word"
              },
              {
                "answer": "courage",
                "kind": "word"
              },
              {
                "answer": "journal",
                "kind": "word"
              },
              {
                "answer": "discourage",
                "kind": "word"
              },
              {
                "answer": "door",
                "kind": "word"
              },
              {
                "answer": "floor",
                "kind": "word"
              },
              {
                "answer": "business",
                "kind": "word"
              },
              {
                "answer": "answer",
                "kind": "word"
              },
              {
                "answer": "receipt",
                "kind": "word"
              },
              {
                "answer": "ceiling",
                "kind": "word"
              },
              {
                "answer": "thousand",
                "kind": "word"
              },
              {
                "answer": "calendar",
                "kind": "word"
              },
              {
                "answer": "grammar",
                "kind": "word"
              },
              {
                "answer": "mustard",
                "kind": "word"
              },
              {
                "answer": "sugar",
                "kind": "word"
              },
              {
                "answer": "build",
                "kind": "word"
              },
              {
                "answer": "buy",
                "kind": "word"
              },
              {
                "answer": "campaign",
                "kind": "word"
              },
              {
                "answer": "sandwich",
                "kind": "word"
              },
              {
                "answer": "Christmas",
                "kind": "word"
              },
              {
                "answer": "listen",
                "kind": "word"
              },
              {
                "answer": "often",
                "kind": "word"
              },
              {
                "answer": "already",
                "kind": "word"
              },
              {
                "answer": "beauty",
                "kind": "word"
              },
              {
                "answer": "bruise",
                "kind": "word"
              },
              {
                "answer": "courage",
                "kind": "word"
              }
            ]
          }
        ]
      }
    }
  }
];
