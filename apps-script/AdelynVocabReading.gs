/**
 * The ELA Pastry Kitchen — Adelyn's Vocabulary & Reading, Weeks 1-36.
 *
 * Adds net-new Schedule rows for Adelyn's "vocab" and "reading" subjects,
 * covering the full 36-week year, converted from
 * Adelyns_Vocabulary_Course_Grade4.docx and Adelyns_Reading_Program.docx
 * into the same content_json shape used for Kenley's content. Does NOT
 * touch setupSheets(), Month1Weeks2to5.gs, or any Kenley rows — this is
 * purely additive, and only touches Adelyn's placeholder vocab/reading rows.
 *
 * Run seedAdelynVocabReading_() ONCE from the function dropdown, after
 * setupSheets() has already been run at least once. It's idempotent: if it
 * detects task_id "av1" already in the Schedule sheet, it assumes this has
 * already run and does nothing. It also removes Adelyn's original single
 * placeholder rows for "vocab" and "reading" ("av1" waiting-shell and "ar1"
 * waiting-shell — note the real content below reuses "ar1" for Week 1's
 * actual reading lesson, so the OLD placeholder "ar1" is deleted by task
 * label match before the new rows are added, to avoid an id collision).
 *
 * Content structure per week:
 *   Vocabulary — one "read" task with that week's 3 root words + 3
 *     frequency words (table format, matching Kenley's vocab-table style).
 *     Every 4th week (4, 8, ... 36) adds a "Review Quiz" graded-mc task
 *     (monthlyTest: true) built directly from the course's own Word Bank +
 *     Parent Answer Key — each definition becomes a question, each week's
 *     12-word bank becomes that question's multiple-choice options, so the
 *     parent answer key is baked directly into the "correct" index. Week 36
 *     additionally adds the End-of-Year Vocabulary Test (termFinal: true).
 *   Reading — one "read" task (book, standard focus, pacing goal, and the
 *     week's Concept Spotlight teaching text), one "graded-dictation" task
 *     built from that week's Appendix A phonics warm-up word list plus its
 *     practice sentence (reuses the exact same listen-and-type mechanic as
 *     Kenley's spelling dictation — she taps to hear each word/sentence via
 *     text-to-speech and types what she hears), and one "reflection" task
 *     per Talk It Over discussion question, each with the course's own
 *     Sample Answer carried over as that task's sampleAnswer field (shown
 *     only in Parent view, to help grade her response).
 *
 * A "🔊 Read this to me" button (text-to-speech via speakElementText in
 * app.js) is available on every read/reflection/dictation/graded-mc task
 * for both children, not just Adelyn — added alongside this content so she
 * can have any lesson, question, or instruction read aloud on tap.
 */

function seedAdelynVocabReading_() {
  var sh = getSheet_('Schedule');
  var existing = sheetToObjects_(sh);
  if (existing.some(function (r) { return r.task_id === 'av1' && r.label === "Study This Week's Words"; })) {
    SpreadsheetApp.getUi().alert('Adelyn\'s Vocabulary/Reading content already appears to be seeded (found task av1) — skipping to avoid duplicates.');
    return;
  }

  // Remove Adelyn's original single placeholder rows for vocab/reading
  // ("av1"/"ar1" waiting-shells) so the new "ar1" (Week 1's real reading
  // lesson) doesn't collide with the old placeholder id.
  var toDelete = existing.filter(function (r) {
    return r.student === 'adelyn' && (r.subject_key === 'vocab' || r.subject_key === 'reading') &&
      (r.task_id === 'av1' || r.task_id === 'ar1') && r.label.indexOf('Waiting') !== -1;
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

  ADELYN_VOCAB_READING_WEEKS.forEach(function (week) {
    Object.keys(week.subjects).forEach(function (subjectKey) {
      var subj = week.subjects[subjectKey];
      addTasks(subjectKey, subj.name, subj.tag, week.week_number, subj.tasks);
    });
  });

  var headers = SHEET_HEADERS.Schedule;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  SpreadsheetApp.getUi().alert('Added ' + rows.length + ' new Schedule rows for Adelyn — Vocabulary & Reading, all 36 weeks.');
}

// ---------- Content ----------

var ADELYN_VOCAB_READING_WEEKS = [
  {
    "week_number": 1,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "TELE- — far, far off (Greek prefix)",
        "tasks": [
          {
            "id": "av1",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>TELE-</b> — far, far off (Greek prefix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>telephone</td><td><i>tele- (far) + phone (sound)</i> — a device used to talk to someone who is far away<br><span style=\"opacity:.75;\">\"Adelyn called her grandmother on the telephone to tell her about the trip.\"</span></td></tr><tr><td>television</td><td><i>tele- (far) + vision (to see)</i> — a device that shows moving pictures and sound sent from far away<br><span style=\"opacity:.75;\">\"The family watched a nature show on television after dinner.\"</span></td></tr><tr><td>telescope</td><td><i>tele- (far) + scope (to look at)</i> — a tool that makes faraway objects, like stars, look closer<br><span style=\"opacity:.75;\">\"Through the telescope, Adelyn could see the craters on the moon.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>however</td><td>used to introduce a statement that contrasts with something already said<br><span style=\"opacity:.75;\">\"Adelyn wanted to go to the beach; however, it started raining.\"</span></td></tr><tr><td>although</td><td>even though; despite the fact that<br><span style=\"opacity:.75;\">\"Although she was tired, Adelyn finished her stretching routine.\"</span></td></tr><tr><td>particular</td><td>specific; one certain thing and not others<br><span style=\"opacity:.75;\">\"Kenley had one particular market stall she visited every week.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Harriet the Invincible · RL.4.3 (intro, brief)",
        "tasks": [
          {
            "id": "ar1",
            "label": "Week 1 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Harriet the Invincible (quick retrospective — already read)</p><p><b>Standard Focus:</b> RL.4.3 (intro, brief)</p><p><b>Pacing Goal:</b> No new reading this week — flip back to 2-3 favorite pages together.</p><p>Since this book is already behind you, this is a single condensed look-back rather than a fresh read. Introduce characterization for the first time now: characters show us who they are through their thoughts, words, and actions — not just what the narrator tells us. Flip to a favorite scene or two (maybe one with Wilbur for contrast, and a picture that adds something the words don't) and talk it through with that lens before moving to a brand-new book next week.</p></div>"
          },
          {
            "id": "ard1",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Short Vowels &amp; Closed Syllables</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "cat",
                "kind": "word"
              },
              {
                "answer": "plan",
                "kind": "word"
              },
              {
                "answer": "trip",
                "kind": "word"
              },
              {
                "answer": "black",
                "kind": "word"
              },
              {
                "answer": "stomp",
                "kind": "word"
              },
              {
                "answer": "disgust",
                "kind": "word"
              },
              {
                "answer": "contest",
                "kind": "word"
              },
              {
                "answer": "blanket",
                "kind": "word"
              },
              {
                "answer": "Kenley packed a snack for the trip.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar1q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "Thinking back on the whole book, what's one thing Harriet did or said that showed who she is as a character?",
            "sampleAnswer": "Harriet goes looking for dangerous quests instead of sitting quietly and waiting for her curse to happen — that shows she's stubborn and brave, and doesn't accept being told what a princess ‘should’ do."
          },
          {
            "id": "ar1q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "How was Harriet different from Wilbur, and how do we know?",
            "sampleAnswer": "Wilbur tends to be cautious and practical, often trying to talk Harriet out of her riskiest plans, while Harriet charges ahead anyway — we know this from how differently they react to the same dangerous situations."
          }
        ]
      }
    }
  },
  {
    "week_number": 2,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "GRAPH / GRAM — write (Greek root)",
        "tasks": [
          {
            "id": "av2",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>GRAPH / GRAM</b> — write (Greek root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>paragraph</td><td><i>para- (beside) + graph (write)</i> — a group of sentences about one main idea<br><span style=\"opacity:.75;\">\"Adelyn wrote a paragraph describing her favorite market in Lisbon.\"</span></td></tr><tr><td>biography</td><td><i>bio- (life) + graph (write)</i> — a true story of a person's life written by someone else<br><span style=\"opacity:.75;\">\"Kenley read a biography about a famous inventor.\"</span></td></tr><tr><td>photograph</td><td><i>photo- (light) + graph (write)</i> — a picture made using light, usually with a camera<br><span style=\"opacity:.75;\">\"Dad took a photograph of the sunset over the mountains.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>determine</td><td>to decide something after careful thought; to figure out<br><span style=\"opacity:.75;\">\"They used a map to determine the fastest route to the museum.\"</span></td></tr><tr><td>additional</td><td>extra; more than what is already there<br><span style=\"opacity:.75;\">\"The teacher gave Adelyn an additional worksheet for practice.\"</span></td></tr><tr><td>recognize</td><td>to know or identify someone or something from before<br><span style=\"opacity:.75;\">\"Adelyn didn't recognize the city at night with all the lights.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "The Case of the Stolen Sixpence · RL.4.3",
        "tasks": [
          {
            "id": "ar2",
            "label": "Week 2 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> The Case of the Stolen Sixpence (Maisie Hitchins, Book 1)</p><p><b>Standard Focus:</b> RL.4.3</p><p><b>Pacing Goal:</b> Read to roughly the 35% mark (through the puppy rescue and George's firing).</p><p>Fresh book, fresh start: this time we catch characterization as it happens, in real time. Maisie is a “noticing sort of person” who wants to be a detective — watch for a moment where what she does or says (not the narrator) shows that about her.</p></div>"
          },
          {
            "id": "ard2",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Silent-E (Magic E)</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "cake",
                "kind": "word"
              },
              {
                "answer": "smile",
                "kind": "word"
              },
              {
                "answer": "stripe",
                "kind": "word"
              },
              {
                "answer": "invite",
                "kind": "word"
              },
              {
                "answer": "compete",
                "kind": "word"
              },
              {
                "answer": "decide",
                "kind": "word"
              },
              {
                "answer": "costume",
                "kind": "word"
              },
              {
                "answer": "parade",
                "kind": "word"
              },
              {
                "answer": "Adelyn will compete in the costume parade.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar2q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "Find one thing Maisie does or says that shows she's observant or determined. Read it out loud.",
            "sampleAnswer": "Maisie notices small details other people miss — like realizing the puppy in the sack was deliberately abandoned, not just lost — and decides right then to look into it herself instead of just feeling bad about it."
          },
          {
            "id": "ar2q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "What kind of detective do you think Maisie will be, based on what we've seen so far?",
            "sampleAnswer": "A hands-on one — she doesn't wait for permission or worry about getting in trouble; she just starts asking questions and following leads, the way her hero Gilbert Carrington would."
          }
        ]
      }
    }
  },
  {
    "week_number": 3,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "PORT — carry (Latin root)",
        "tasks": [
          {
            "id": "av3",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>PORT</b> — carry (Latin root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>transport</td><td><i>trans- (across) + port (carry)</i> — to carry people or things from one place to another<br><span style=\"opacity:.75;\">\"A big ship can transport hundreds of shipping containers across the ocean.\"</span></td></tr><tr><td>portable</td><td><i>port (carry) + -able (able to be)</i> — able to be carried or moved easily<br><span style=\"opacity:.75;\">\"Adelyn packed a portable speaker for the long train ride.\"</span></td></tr><tr><td>report</td><td><i>re- (back) + port (carry)</i> — to carry information back and share it with others<br><span style=\"opacity:.75;\">\"Adelyn will report on what she learned about Portuguese history.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>purpose</td><td>the reason something is done; a goal<br><span style=\"opacity:.75;\">\"The purpose of the trip was to visit Kenley's old friend.\"</span></td></tr><tr><td>environment</td><td>the surroundings or conditions in which someone lives<br><span style=\"opacity:.75;\">\"The rainforest environment was hot, humid, and full of sound.\"</span></td></tr><tr><td>achieve</td><td>to succeed in reaching a goal<br><span style=\"opacity:.75;\">\"Adelyn worked hard all year to achieve a perfect cartwheel.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "The Case of the Stolen Sixpence · RL.4.3",
        "tasks": [
          {
            "id": "ar3",
            "label": "Week 3 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> The Case of the Stolen Sixpence</p><p><b>Standard Focus:</b> RL.4.3</p><p><b>Pacing Goal:</b> Read to roughly the 70% mark.</p><p>Characters are also defined by contrast with the people around them — and by their sidekicks. Look at Maisie alongside Eddie (her dog) or George. What does the way she treats them show about her?</p></div>"
          },
          {
            "id": "ard3",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Vowel Teams (ai/ay, ee/ea, oa/ow, ie)</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "rain",
                "kind": "word"
              },
              {
                "answer": "play",
                "kind": "word"
              },
              {
                "answer": "tea",
                "kind": "word"
              },
              {
                "answer": "feet",
                "kind": "word"
              },
              {
                "answer": "coat",
                "kind": "word"
              },
              {
                "answer": "snow",
                "kind": "word"
              },
              {
                "answer": "pie",
                "kind": "word"
              },
              {
                "answer": "field",
                "kind": "word"
              },
              {
                "answer": "We saw rain clouds over the green field.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar3q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "How does Maisie treat Eddie or George, and what does that reveal about her character?",
            "sampleAnswer": "She's loyal and protective — she keeps Eddie even though she knows her grandmother won't approve, and she takes George's side without hesitation when he's blamed for something she doesn't believe he did."
          },
          {
            "id": "ar3q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "What's a moment where Maisie's personality causes a problem or a breakthrough in the case?",
            "sampleAnswer": "Her stubborn curiosity gets her into scrapes, like poking around where she's not supposed to be, but that same trait is exactly what leads her to clues the grown-ups overlook."
          }
        ]
      }
    }
  },
  {
    "week_number": 4,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "DICT — say, speak (Latin root)",
        "tasks": [
          {
            "id": "av4",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>DICT</b> — say, speak (Latin root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>predict</td><td><i>pre- (before) + dict (say)</i> — to say what will happen before it happens<br><span style=\"opacity:.75;\">\"Kenley tried to predict how the soup would taste before adding the spices.\"</span></td></tr><tr><td>dictionary</td><td><i>dict (say/word) + -ary (place/collection)</i> — a book that lists words and their meanings<br><span style=\"opacity:.75;\">\"When Adelyn didn't know a word, she looked it up in the dictionary.\"</span></td></tr><tr><td>dictate</td><td><i>dict (say) + -ate (to do)</i> — to say words aloud for someone else to write down<br><span style=\"opacity:.75;\">\"Mom dictated the grocery list while Adelyn wrote it down.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>ordinary</td><td>normal; not special or unusual<br><span style=\"opacity:.75;\">\"It was just an ordinary Tuesday until the surprise phone call.\"</span></td></tr><tr><td>opportunity</td><td>a chance to do something<br><span style=\"opacity:.75;\">\"Living abroad gave Adelyn the opportunity to learn Spanish.\"</span></td></tr><tr><td>various</td><td>several different<br><span style=\"opacity:.75;\">\"The market sold various spices from all over the region.\"</span></td></tr></table>"
          },
          {
            "id": "avq1",
            "label": "Review Quiz 1 — Weeks 1–4",
            "type": "graded-mc",
            "monthlyTest": true,
            "questions": [
              {
                "q": "A picture made using light, usually with a camera",
                "options": [
                  "biography",
                  "dictate",
                  "dictionary",
                  "paragraph",
                  "photograph",
                  "portable",
                  "predict",
                  "report",
                  "telephone",
                  "telescope",
                  "television",
                  "transport"
                ],
                "correct": 4
              },
              {
                "q": "A true story of a person's life written by someone else",
                "options": [
                  "biography",
                  "dictate",
                  "dictionary",
                  "paragraph",
                  "photograph",
                  "portable",
                  "predict",
                  "report",
                  "telephone",
                  "telescope",
                  "television",
                  "transport"
                ],
                "correct": 0
              },
              {
                "q": "A group of sentences about one main idea",
                "options": [
                  "biography",
                  "dictate",
                  "dictionary",
                  "paragraph",
                  "photograph",
                  "portable",
                  "predict",
                  "report",
                  "telephone",
                  "telescope",
                  "television",
                  "transport"
                ],
                "correct": 3
              },
              {
                "q": "A tool that makes faraway objects, like stars, look closer",
                "options": [
                  "biography",
                  "dictate",
                  "dictionary",
                  "paragraph",
                  "photograph",
                  "portable",
                  "predict",
                  "report",
                  "telephone",
                  "telescope",
                  "television",
                  "transport"
                ],
                "correct": 9
              },
              {
                "q": "A device that shows moving pictures and sound sent from far away",
                "options": [
                  "biography",
                  "dictate",
                  "dictionary",
                  "paragraph",
                  "photograph",
                  "portable",
                  "predict",
                  "report",
                  "telephone",
                  "telescope",
                  "television",
                  "transport"
                ],
                "correct": 10
              },
              {
                "q": "A device used to talk to someone who is far away",
                "options": [
                  "biography",
                  "dictate",
                  "dictionary",
                  "paragraph",
                  "photograph",
                  "portable",
                  "predict",
                  "report",
                  "telephone",
                  "telescope",
                  "television",
                  "transport"
                ],
                "correct": 8
              },
              {
                "q": "To say words aloud for someone else to write down",
                "options": [
                  "biography",
                  "dictate",
                  "dictionary",
                  "paragraph",
                  "photograph",
                  "portable",
                  "predict",
                  "report",
                  "telephone",
                  "telescope",
                  "television",
                  "transport"
                ],
                "correct": 1
              },
              {
                "q": "A book that lists words and their meanings",
                "options": [
                  "biography",
                  "dictate",
                  "dictionary",
                  "paragraph",
                  "photograph",
                  "portable",
                  "predict",
                  "report",
                  "telephone",
                  "telescope",
                  "television",
                  "transport"
                ],
                "correct": 2
              },
              {
                "q": "To say what will happen before it happens",
                "options": [
                  "biography",
                  "dictate",
                  "dictionary",
                  "paragraph",
                  "photograph",
                  "portable",
                  "predict",
                  "report",
                  "telephone",
                  "telescope",
                  "television",
                  "transport"
                ],
                "correct": 6
              },
              {
                "q": "To carry information back and share it with others",
                "options": [
                  "biography",
                  "dictate",
                  "dictionary",
                  "paragraph",
                  "photograph",
                  "portable",
                  "predict",
                  "report",
                  "telephone",
                  "telescope",
                  "television",
                  "transport"
                ],
                "correct": 7
              },
              {
                "q": "Able to be carried or moved easily",
                "options": [
                  "biography",
                  "dictate",
                  "dictionary",
                  "paragraph",
                  "photograph",
                  "portable",
                  "predict",
                  "report",
                  "telephone",
                  "telescope",
                  "television",
                  "transport"
                ],
                "correct": 5
              },
              {
                "q": "To carry people or things from one place to another",
                "options": [
                  "biography",
                  "dictate",
                  "dictionary",
                  "paragraph",
                  "photograph",
                  "portable",
                  "predict",
                  "report",
                  "telephone",
                  "telescope",
                  "television",
                  "transport"
                ],
                "correct": 11
              },
              {
                "q": "To know or identify someone or something from before",
                "options": [
                  "achieve",
                  "additional",
                  "although",
                  "determine",
                  "environment",
                  "however",
                  "opportunity",
                  "ordinary",
                  "particular",
                  "purpose",
                  "recognize",
                  "various"
                ],
                "correct": 10
              },
              {
                "q": "Extra; more than what is already there",
                "options": [
                  "achieve",
                  "additional",
                  "although",
                  "determine",
                  "environment",
                  "however",
                  "opportunity",
                  "ordinary",
                  "particular",
                  "purpose",
                  "recognize",
                  "various"
                ],
                "correct": 1
              },
              {
                "q": "To decide something after careful thought; to figure out",
                "options": [
                  "achieve",
                  "additional",
                  "although",
                  "determine",
                  "environment",
                  "however",
                  "opportunity",
                  "ordinary",
                  "particular",
                  "purpose",
                  "recognize",
                  "various"
                ],
                "correct": 3
              },
              {
                "q": "Specific; one certain thing and not others",
                "options": [
                  "achieve",
                  "additional",
                  "although",
                  "determine",
                  "environment",
                  "however",
                  "opportunity",
                  "ordinary",
                  "particular",
                  "purpose",
                  "recognize",
                  "various"
                ],
                "correct": 8
              },
              {
                "q": "Even though; despite the fact that",
                "options": [
                  "achieve",
                  "additional",
                  "although",
                  "determine",
                  "environment",
                  "however",
                  "opportunity",
                  "ordinary",
                  "particular",
                  "purpose",
                  "recognize",
                  "various"
                ],
                "correct": 2
              },
              {
                "q": "Used to introduce a statement that contrasts with something already said",
                "options": [
                  "achieve",
                  "additional",
                  "although",
                  "determine",
                  "environment",
                  "however",
                  "opportunity",
                  "ordinary",
                  "particular",
                  "purpose",
                  "recognize",
                  "various"
                ],
                "correct": 5
              },
              {
                "q": "Several different",
                "options": [
                  "achieve",
                  "additional",
                  "although",
                  "determine",
                  "environment",
                  "however",
                  "opportunity",
                  "ordinary",
                  "particular",
                  "purpose",
                  "recognize",
                  "various"
                ],
                "correct": 11
              },
              {
                "q": "A chance to do something",
                "options": [
                  "achieve",
                  "additional",
                  "although",
                  "determine",
                  "environment",
                  "however",
                  "opportunity",
                  "ordinary",
                  "particular",
                  "purpose",
                  "recognize",
                  "various"
                ],
                "correct": 6
              },
              {
                "q": "Normal; not special or unusual",
                "options": [
                  "achieve",
                  "additional",
                  "although",
                  "determine",
                  "environment",
                  "however",
                  "opportunity",
                  "ordinary",
                  "particular",
                  "purpose",
                  "recognize",
                  "various"
                ],
                "correct": 7
              },
              {
                "q": "To succeed in reaching a goal",
                "options": [
                  "achieve",
                  "additional",
                  "although",
                  "determine",
                  "environment",
                  "however",
                  "opportunity",
                  "ordinary",
                  "particular",
                  "purpose",
                  "recognize",
                  "various"
                ],
                "correct": 0
              },
              {
                "q": "The surroundings or conditions in which someone lives",
                "options": [
                  "achieve",
                  "additional",
                  "although",
                  "determine",
                  "environment",
                  "however",
                  "opportunity",
                  "ordinary",
                  "particular",
                  "purpose",
                  "recognize",
                  "various"
                ],
                "correct": 4
              },
              {
                "q": "The reason something is done; a goal",
                "options": [
                  "achieve",
                  "additional",
                  "although",
                  "determine",
                  "environment",
                  "however",
                  "opportunity",
                  "ordinary",
                  "particular",
                  "purpose",
                  "recognize",
                  "various"
                ],
                "correct": 9
              }
            ]
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "The Case of the Stolen Sixpence · RL.4.7",
        "tasks": [
          {
            "id": "ar4",
            "label": "Week 4 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> The Case of the Stolen Sixpence</p><p><b>Standard Focus:</b> RL.4.7</p><p><b>Pacing Goal:</b> Finish the book.</p><p>This book uses black-and-white line drawings throughout. Pick an illustration and talk about what it shows that the words alone don't — then wrap up the mystery.</p></div>"
          },
          {
            "id": "ard4",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>R-Controlled Vowels (ar, er, ir, or, ur)</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "star",
                "kind": "word"
              },
              {
                "answer": "market",
                "kind": "word"
              },
              {
                "answer": "bird",
                "kind": "word"
              },
              {
                "answer": "storm",
                "kind": "word"
              },
              {
                "answer": "turn",
                "kind": "word"
              },
              {
                "answer": "sister",
                "kind": "word"
              },
              {
                "answer": "corner",
                "kind": "word"
              },
              {
                "answer": "garden",
                "kind": "word"
              },
              {
                "answer": "The market is just around the corner.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar4q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "Find a picture that adds something the words didn't say. What is it?",
            "sampleAnswer": "A good one to look for is an illustration where Eddie's expression or body language reacts to something before the text even mentions he noticed it — the art doing some of the storytelling on its own."
          },
          {
            "id": "ar4q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "How was the mystery solved, and did Maisie turn out to be a good detective?",
            "sampleAnswer": "Maisie pieces together clues the adults ignored and proves who actually took the sixpence — showing that paying close attention and refusing to give up made her a real detective, just like she hoped."
          }
        ]
      }
    }
  },
  {
    "week_number": 5,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "SPECT — look, see (Latin root)",
        "tasks": [
          {
            "id": "av5",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>SPECT</b> — look, see (Latin root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>inspect</td><td><i>in- (into) + spect (look)</i> — to look at something closely and carefully<br><span style=\"opacity:.75;\">\"The mechanic will inspect the van before our long road trip.\"</span></td></tr><tr><td>respect</td><td><i>re- (again) + spect (look)</i> — a feeling of admiration for someone; to treat someone with consideration<br><span style=\"opacity:.75;\">\"Adelyn showed respect for her coach by listening carefully.\"</span></td></tr><tr><td>spectator</td><td><i>spect (look) + -ator (one who)</i> — a person who watches an event without taking part<br><span style=\"opacity:.75;\">\"Hundreds of spectators cheered from the stands during the gymnastics meet.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>especially</td><td>more than usual; particularly<br><span style=\"opacity:.75;\">\"Adelyn loves fruit, especially mangoes.\"</span></td></tr><tr><td>considerable</td><td>fairly large in amount or size<br><span style=\"opacity:.75;\">\"It took a considerable amount of time to pack for the trip.\"</span></td></tr><tr><td>eventually</td><td>after some time; in the end<br><span style=\"opacity:.75;\">\"The rain stopped eventually, and they went outside to play.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "The Tea Dragon Society · RL.4.5",
        "tasks": [
          {
            "id": "ar5",
            "label": "Week 5 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> The Tea Dragon Society (Kay O'Neill)</p><p><b>Standard Focus:</b> RL.4.5</p><p><b>Pacing Goal:</b> Read Chapter 1 (Summer).</p><p>This graphic novel is built as four chapters, one per season — a clear, visible shape. Greta discovers a lost tea dragon and meets the tea shop owners who teach her about a fading art. Notice how starting a new season signals a new stage in the story.</p></div>"
          },
          {
            "id": "ard5",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Blends &amp; Digraphs in Multisyllable Words</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "splash",
                "kind": "word"
              },
              {
                "answer": "thunder",
                "kind": "word"
              },
              {
                "answer": "chapter",
                "kind": "word"
              },
              {
                "answer": "shrimp",
                "kind": "word"
              },
              {
                "answer": "stretch",
                "kind": "word"
              },
              {
                "answer": "whisper",
                "kind": "word"
              },
              {
                "answer": "scratch",
                "kind": "word"
              },
              {
                "answer": "sunscreen",
                "kind": "word"
              },
              {
                "answer": "Adelyn read one more chapter before the thunder started.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar5q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What's the season, and what happens in it that moves the story forward?",
            "sampleAnswer": "It's Summer — Greta finds an abandoned tea dragon and brings it to the Tea Dragon Society shop, which is what introduces her to Hesekiel, Erik, and this whole hidden world."
          },
          {
            "id": "ar5q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "What's Greta's first impression of the tea dragons and the people who care for them?",
            "sampleAnswer": "She's curious but a little skeptical at first — tea dragon care seems slow and old-fashioned compared to the quick, visible results she's used to from blacksmithing."
          }
        ]
      }
    }
  },
  {
    "week_number": 6,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "STRUCT — build (Latin root)",
        "tasks": [
          {
            "id": "av6",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>STRUCT</b> — build (Latin root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>construct</td><td><i>con- (together) + struct (build)</i> — to build or put something together<br><span style=\"opacity:.75;\">\"The workers used bamboo scaffolding to construct the new building.\"</span></td></tr><tr><td>instruct</td><td><i>in- (in) + struct (build)</i> — to teach or give directions<br><span style=\"opacity:.75;\">\"The coach will instruct the class on a new stretching routine.\"</span></td></tr><tr><td>structure</td><td><i>struct (build) + -ure (result of)</i> — something that has been built; the way parts are arranged<br><span style=\"opacity:.75;\">\"The old stone structure had stood in the village for centuries.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>appropriate</td><td>suitable or correct for the situation<br><span style=\"opacity:.75;\">\"Kenley picked an appropriate outfit for the fancy dinner.\"</span></td></tr><tr><td>obvious</td><td>easy to see or understand<br><span style=\"opacity:.75;\">\"It was obvious that Adelyn had practiced her routine every day.\"</span></td></tr><tr><td>generally</td><td>usually; in most cases<br><span style=\"opacity:.75;\">\"Adelyn generally wakes up early, even on travel days.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "The Tea Dragon Society · RL.4.5",
        "tasks": [
          {
            "id": "ar6",
            "label": "Week 6 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> The Tea Dragon Society</p><p><b>Standard Focus:</b> RL.4.5</p><p><b>Pacing Goal:</b> Read Chapter 2 (Autumn).</p><p>Keep tracking the four-part shape: are the seasonal chapters more like one continuous story, or separate connected vignettes held together by Greta's growth? Either answer is fine — the point is noticing how the author chose to organize time.</p></div>"
          },
          {
            "id": "ard6",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Common Prefixes (re-, un-, dis-, pre-, mis-, non-)</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "redo",
                "kind": "word"
              },
              {
                "answer": "unpack",
                "kind": "word"
              },
              {
                "answer": "distrust",
                "kind": "word"
              },
              {
                "answer": "preview",
                "kind": "word"
              },
              {
                "answer": "mislead",
                "kind": "word"
              },
              {
                "answer": "nonstop",
                "kind": "word"
              },
              {
                "answer": "rewind",
                "kind": "word"
              },
              {
                "answer": "unfold",
                "kind": "word"
              },
              {
                "answer": "Kenley had to unpack and redo her bag.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar6q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "Does this chapter feel like a direct continuation of Chapter 1, or more like its own mini-story? What makes you say that?",
            "sampleAnswer": "It feels like a continuation with its own small arc — Greta keeps coming back to visit, so there's an ongoing relationship building, but each visit also has its own small moment or lesson."
          },
          {
            "id": "ar6q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "What is Greta learning to be patient about?",
            "sampleAnswer": "She's learning that some things, like caring for a tea dragon or mastering a craft, take a long time and can't be rushed just because she's eager for a result."
          }
        ]
      }
    }
  },
  {
    "week_number": 7,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "RUPT — break (Latin root)",
        "tasks": [
          {
            "id": "av7",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>RUPT</b> — break (Latin root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>interrupt</td><td><i>inter- (between) + rupt (break)</i> — to break into a conversation or activity<br><span style=\"opacity:.75;\">\"It's polite to wait your turn instead of interrupting someone.\"</span></td></tr><tr><td>erupt</td><td><i>e- (out) + rupt (break)</i> — to burst out suddenly, like a volcano<br><span style=\"opacity:.75;\">\"The tour guide explained how the volcano could erupt without much warning.\"</span></td></tr><tr><td>disrupt</td><td><i>dis- (apart) + rupt (break)</i> — to break apart or interrupt the normal order of something<br><span style=\"opacity:.75;\">\"A power outage disrupted the cooking class for almost an hour.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>entire</td><td>whole; complete<br><span style=\"opacity:.75;\">\"They explored the entire old town in one afternoon.\"</span></td></tr><tr><td>immediately</td><td>right away; at once<br><span style=\"opacity:.75;\">\"Adelyn immediately recognized her cousin at the airport.\"</span></td></tr><tr><td>gradually</td><td>slowly, little by little<br><span style=\"opacity:.75;\">\"Her Spanish improved gradually over the year of travel.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "The Tea Dragon Society · RL.4.7",
        "tasks": [
          {
            "id": "ar7",
            "label": "Week 7 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> The Tea Dragon Society</p><p><b>Standard Focus:</b> RL.4.7</p><p><b>Pacing Goal:</b> Read Chapter 3 (Winter).</p><p>This is a graphic novel, so illustrations aren't just decoration — whole beats of the story happen in panels with little or no text, especially the dreamy tea-memory sequences. Find a quiet, mostly-wordless page and talk through what it communicates.</p></div>"
          },
          {
            "id": "ard7",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Common Suffixes (-tion, -sion, -ing, -ed, -ly, -ful, -less)</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "action",
                "kind": "word"
              },
              {
                "answer": "mission",
                "kind": "word"
              },
              {
                "answer": "jumping",
                "kind": "word"
              },
              {
                "answer": "packed",
                "kind": "word"
              },
              {
                "answer": "quickly",
                "kind": "word"
              },
              {
                "answer": "careful",
                "kind": "word"
              },
              {
                "answer": "fearless",
                "kind": "word"
              },
              {
                "answer": "direction",
                "kind": "word"
              },
              {
                "answer": "Adelyn quickly packed her bag for the mission.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar7q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "Find a page with little or no text. What is it telling us just through the pictures?",
            "sampleAnswer": "Look for one of the dreamy ‘tea memory’ sequences — often nearly wordless, using soft colors and imagery to show a character's past feelings rather than explaining them outright."
          },
          {
            "id": "ar7q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "How does the art's color or mood change between a happy scene and a quieter one?",
            "sampleAnswer": "Happier scenes tend to use warmer, brighter colors, while quieter or more reflective moments shift to cooler, softer tones — the color palette is doing emotional work the dialogue doesn't spell out."
          }
        ]
      }
    }
  },
  {
    "week_number": 8,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "SCRIB / SCRIPT — write (Latin root)",
        "tasks": [
          {
            "id": "av8",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>SCRIB / SCRIPT</b> — write (Latin root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>describe</td><td><i>de- (down) + scrib (write)</i> — to tell or write details about something<br><span style=\"opacity:.75;\">\"Adelyn tried to describe the taste of the strange new fruit.\"</span></td></tr><tr><td>subscribe</td><td><i>sub- (under) + scrib (write)</i> — to sign up to regularly receive something<br><span style=\"opacity:.75;\">\"Kenley asked to subscribe to a cooking magazine.\"</span></td></tr><tr><td>script</td><td><i>script (write)</i> — the written words of a play, movie, or speech<br><span style=\"opacity:.75;\">\"The actors memorized their script before the show.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>numerous</td><td>many<br><span style=\"opacity:.75;\">\"There were numerous stalls selling handmade jewelry.\"</span></td></tr><tr><td>tremendous</td><td>very great in size or amount<br><span style=\"opacity:.75;\">\"The gymnast made a tremendous effort during her final routine.\"</span></td></tr><tr><td>enormous</td><td>extremely large<br><span style=\"opacity:.75;\">\"The old castle had an enormous courtyard.\"</span></td></tr></table>"
          },
          {
            "id": "avq2",
            "label": "Review Quiz 2 — Weeks 5–8",
            "type": "graded-mc",
            "monthlyTest": true,
            "questions": [
              {
                "q": "Something that has been built; the way parts are arranged",
                "options": [
                  "construct",
                  "describe",
                  "disrupt",
                  "erupt",
                  "inspect",
                  "instruct",
                  "interrupt",
                  "respect",
                  "script",
                  "spectator",
                  "structure",
                  "subscribe"
                ],
                "correct": 10
              },
              {
                "q": "To teach or give directions",
                "options": [
                  "construct",
                  "describe",
                  "disrupt",
                  "erupt",
                  "inspect",
                  "instruct",
                  "interrupt",
                  "respect",
                  "script",
                  "spectator",
                  "structure",
                  "subscribe"
                ],
                "correct": 5
              },
              {
                "q": "To build or put something together",
                "options": [
                  "construct",
                  "describe",
                  "disrupt",
                  "erupt",
                  "inspect",
                  "instruct",
                  "interrupt",
                  "respect",
                  "script",
                  "spectator",
                  "structure",
                  "subscribe"
                ],
                "correct": 0
              },
              {
                "q": "A person who watches an event without taking part",
                "options": [
                  "construct",
                  "describe",
                  "disrupt",
                  "erupt",
                  "inspect",
                  "instruct",
                  "interrupt",
                  "respect",
                  "script",
                  "spectator",
                  "structure",
                  "subscribe"
                ],
                "correct": 9
              },
              {
                "q": "A feeling of admiration for someone; to treat someone with consideration",
                "options": [
                  "construct",
                  "describe",
                  "disrupt",
                  "erupt",
                  "inspect",
                  "instruct",
                  "interrupt",
                  "respect",
                  "script",
                  "spectator",
                  "structure",
                  "subscribe"
                ],
                "correct": 7
              },
              {
                "q": "To look at something closely and carefully",
                "options": [
                  "construct",
                  "describe",
                  "disrupt",
                  "erupt",
                  "inspect",
                  "instruct",
                  "interrupt",
                  "respect",
                  "script",
                  "spectator",
                  "structure",
                  "subscribe"
                ],
                "correct": 4
              },
              {
                "q": "The written words of a play, movie, or speech",
                "options": [
                  "construct",
                  "describe",
                  "disrupt",
                  "erupt",
                  "inspect",
                  "instruct",
                  "interrupt",
                  "respect",
                  "script",
                  "spectator",
                  "structure",
                  "subscribe"
                ],
                "correct": 8
              },
              {
                "q": "To sign up to regularly receive something",
                "options": [
                  "construct",
                  "describe",
                  "disrupt",
                  "erupt",
                  "inspect",
                  "instruct",
                  "interrupt",
                  "respect",
                  "script",
                  "spectator",
                  "structure",
                  "subscribe"
                ],
                "correct": 11
              },
              {
                "q": "To tell or write details about something",
                "options": [
                  "construct",
                  "describe",
                  "disrupt",
                  "erupt",
                  "inspect",
                  "instruct",
                  "interrupt",
                  "respect",
                  "script",
                  "spectator",
                  "structure",
                  "subscribe"
                ],
                "correct": 1
              },
              {
                "q": "To break apart or interrupt the normal order of something",
                "options": [
                  "construct",
                  "describe",
                  "disrupt",
                  "erupt",
                  "inspect",
                  "instruct",
                  "interrupt",
                  "respect",
                  "script",
                  "spectator",
                  "structure",
                  "subscribe"
                ],
                "correct": 2
              },
              {
                "q": "To burst out suddenly, like a volcano",
                "options": [
                  "construct",
                  "describe",
                  "disrupt",
                  "erupt",
                  "inspect",
                  "instruct",
                  "interrupt",
                  "respect",
                  "script",
                  "spectator",
                  "structure",
                  "subscribe"
                ],
                "correct": 3
              },
              {
                "q": "To break into a conversation or activity",
                "options": [
                  "construct",
                  "describe",
                  "disrupt",
                  "erupt",
                  "inspect",
                  "instruct",
                  "interrupt",
                  "respect",
                  "script",
                  "spectator",
                  "structure",
                  "subscribe"
                ],
                "correct": 6
              },
              {
                "q": "Usually; in most cases",
                "options": [
                  "appropriate",
                  "considerable",
                  "enormous",
                  "entire",
                  "especially",
                  "eventually",
                  "generally",
                  "gradually",
                  "immediately",
                  "numerous",
                  "obvious",
                  "tremendous"
                ],
                "correct": 6
              },
              {
                "q": "Easy to see or understand",
                "options": [
                  "appropriate",
                  "considerable",
                  "enormous",
                  "entire",
                  "especially",
                  "eventually",
                  "generally",
                  "gradually",
                  "immediately",
                  "numerous",
                  "obvious",
                  "tremendous"
                ],
                "correct": 10
              },
              {
                "q": "Suitable or correct for the situation",
                "options": [
                  "appropriate",
                  "considerable",
                  "enormous",
                  "entire",
                  "especially",
                  "eventually",
                  "generally",
                  "gradually",
                  "immediately",
                  "numerous",
                  "obvious",
                  "tremendous"
                ],
                "correct": 0
              },
              {
                "q": "After some time; in the end",
                "options": [
                  "appropriate",
                  "considerable",
                  "enormous",
                  "entire",
                  "especially",
                  "eventually",
                  "generally",
                  "gradually",
                  "immediately",
                  "numerous",
                  "obvious",
                  "tremendous"
                ],
                "correct": 5
              },
              {
                "q": "Fairly large in amount or size",
                "options": [
                  "appropriate",
                  "considerable",
                  "enormous",
                  "entire",
                  "especially",
                  "eventually",
                  "generally",
                  "gradually",
                  "immediately",
                  "numerous",
                  "obvious",
                  "tremendous"
                ],
                "correct": 1
              },
              {
                "q": "More than usual; particularly",
                "options": [
                  "appropriate",
                  "considerable",
                  "enormous",
                  "entire",
                  "especially",
                  "eventually",
                  "generally",
                  "gradually",
                  "immediately",
                  "numerous",
                  "obvious",
                  "tremendous"
                ],
                "correct": 4
              },
              {
                "q": "Extremely large",
                "options": [
                  "appropriate",
                  "considerable",
                  "enormous",
                  "entire",
                  "especially",
                  "eventually",
                  "generally",
                  "gradually",
                  "immediately",
                  "numerous",
                  "obvious",
                  "tremendous"
                ],
                "correct": 2
              },
              {
                "q": "Very great in size or amount",
                "options": [
                  "appropriate",
                  "considerable",
                  "enormous",
                  "entire",
                  "especially",
                  "eventually",
                  "generally",
                  "gradually",
                  "immediately",
                  "numerous",
                  "obvious",
                  "tremendous"
                ],
                "correct": 11
              },
              {
                "q": "Many",
                "options": [
                  "appropriate",
                  "considerable",
                  "enormous",
                  "entire",
                  "especially",
                  "eventually",
                  "generally",
                  "gradually",
                  "immediately",
                  "numerous",
                  "obvious",
                  "tremendous"
                ],
                "correct": 9
              },
              {
                "q": "Slowly, little by little",
                "options": [
                  "appropriate",
                  "considerable",
                  "enormous",
                  "entire",
                  "especially",
                  "eventually",
                  "generally",
                  "gradually",
                  "immediately",
                  "numerous",
                  "obvious",
                  "tremendous"
                ],
                "correct": 7
              },
              {
                "q": "Right away; at once",
                "options": [
                  "appropriate",
                  "considerable",
                  "enormous",
                  "entire",
                  "especially",
                  "eventually",
                  "generally",
                  "gradually",
                  "immediately",
                  "numerous",
                  "obvious",
                  "tremendous"
                ],
                "correct": 8
              },
              {
                "q": "Whole; complete",
                "options": [
                  "appropriate",
                  "considerable",
                  "enormous",
                  "entire",
                  "especially",
                  "eventually",
                  "generally",
                  "gradually",
                  "immediately",
                  "numerous",
                  "obvious",
                  "tremendous"
                ],
                "correct": 3
              }
            ]
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "The Tea Dragon Society · RL.4.3 (Characterization Checkpoint #2)",
        "tasks": [
          {
            "id": "ar8",
            "label": "Week 8 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> The Tea Dragon Society</p><p><b>Standard Focus:</b> RL.4.3 (Characterization Checkpoint #2)</p><p><b>Pacing Goal:</b> Read Chapter 4 (Spring) and finish the book.</p><p>Time for the first real characterization check-in since Month 1 — see if the skill transfers to a quiet, gentle character without prompting. Greta shows growth through actions (returning, practicing, staying curious) rather than being told she's changed.</p></div>"
          },
          {
            "id": "ard8",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Syllable Division Patterns</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "napkin",
                "kind": "word"
              },
              {
                "answer": "robot",
                "kind": "word"
              },
              {
                "answer": "cabin",
                "kind": "word"
              },
              {
                "answer": "apple",
                "kind": "word"
              },
              {
                "answer": "magnet",
                "kind": "word"
              },
              {
                "answer": "tiger",
                "kind": "word"
              },
              {
                "answer": "camel",
                "kind": "word"
              },
              {
                "answer": "puzzle",
                "kind": "word"
              },
              {
                "answer": "The tiny robot solved the puzzle.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar8q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What does Greta do by the end that shows she's changed since Chapter 1, without the story just saying so?",
            "sampleAnswer": "She keeps returning and sticking with something even when it's slow or difficult, instead of giving up the way she might have at the start — that persistence itself is the evidence of her growth."
          },
          {
            "id": "ar8q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "How is Greta's kind of bravery different from Harriet's or Maisie's?",
            "sampleAnswer": "Harriet and Maisie are bold and active, jumping into danger or mysteries; Greta's bravery is quieter — it's about being patient and open to something unfamiliar, which is its own kind of courage."
          }
        ]
      }
    }
  },
  {
    "week_number": 9,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "BIO — life (Greek root)",
        "tasks": [
          {
            "id": "av9",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>BIO</b> — life (Greek root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>biology</td><td><i>bio- (life) + -logy (study of)</i> — the scientific study of living things<br><span style=\"opacity:.75;\">\"In biology, Adelyn learned how plants make their own food.\"</span></td></tr><tr><td>antibiotic</td><td><i>anti- (against) + bio (life)</i> — a medicine that fights harmful bacteria<br><span style=\"opacity:.75;\">\"The doctor gave Kenley an antibiotic to help fight the infection.\"</span></td></tr><tr><td>biodegradable</td><td><i>bio- (life) + degrade (break down) + -able (able to be)</i> — able to break down naturally over time<br><span style=\"opacity:.75;\">\"They packed biodegradable utensils for the picnic.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>rarely</td><td>not often<br><span style=\"opacity:.75;\">\"They rarely stayed in one city for more than a month.\"</span></td></tr><tr><td>frequently</td><td>often; many times<br><span style=\"opacity:.75;\">\"Adelyn frequently practiced her Spanish with the shopkeepers.\"</span></td></tr><tr><td>constantly</td><td>happening all the time, without stopping<br><span style=\"opacity:.75;\">\"The busy market was constantly full of noise and color.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Nonfiction Detour #1 · RI.4.1/4.2",
        "tasks": [
          {
            "id": "ar9",
            "label": "Week 9 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Nonfiction Detour #1 (see Appendix D for article options)</p><p><b>Standard Focus:</b> RI.4.1/4.2</p><p><b>Pacing Goal:</b> One short article this week (a single sitting).</p><p>Switch gears this week to a short nonfiction article — anything that interests her (an animal, a place you're visiting, or a science topic she's already learning). The “main idea” is what the piece is mostly about, and the details are the evidence that support it.</p></div>"
          },
          {
            "id": "ard9",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Homophones &amp; Frequently Confused Words</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "their",
                "kind": "word"
              },
              {
                "answer": "there",
                "kind": "word"
              },
              {
                "answer": "they're",
                "kind": "word"
              },
              {
                "answer": "your",
                "kind": "word"
              },
              {
                "answer": "you're",
                "kind": "word"
              },
              {
                "answer": "its",
                "kind": "word"
              },
              {
                "answer": "it's",
                "kind": "word"
              },
              {
                "answer": "whose",
                "kind": "word"
              },
              {
                "answer": "Their bags are over there, and they're ready to go.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar9q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "In one sentence, what is this article mostly about?",
            "sampleAnswer": "‘This article is about how red pandas use their long, striped tails to balance and stay warm.’ A strong answer names the specific topic and the main point, not just the general subject."
          },
          {
            "id": "ar9q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "What's one detail from the text that supports that main idea?",
            "sampleAnswer": "A good answer points to a specific fact, not a vague impression — e.g., ‘It says their tails are almost as long as their bodies, which helps them balance in trees.’"
          }
        ]
      }
    }
  },
  {
    "week_number": 10,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "PHOTO — light (Greek root)",
        "tasks": [
          {
            "id": "av10",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>PHOTO</b> — light (Greek root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>photocopy</td><td><i>photo- (light) + copy</i> — a copy of a document made using a machine that uses light<br><span style=\"opacity:.75;\">\"The teacher made a photocopy of the map for every student.\"</span></td></tr><tr><td>photosynthesis</td><td><i>photo- (light) + synthesis (putting together)</i> — the process plants use to turn sunlight into food<br><span style=\"opacity:.75;\">\"During photosynthesis, a leaf uses sunlight, water, and air to make sugar.\"</span></td></tr><tr><td>photogenic</td><td><i>photo- (light) + gen (produce) + -ic (having qualities of)</i> — looking good in photographs<br><span style=\"opacity:.75;\">\"Adelyn's cartwheel photo turned out very photogenic.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>suddenly</td><td>happening quickly and unexpectedly<br><span style=\"opacity:.75;\">\"Suddenly, the sky turned dark and it began to pour.\"</span></td></tr><tr><td>previously</td><td>before now; earlier<br><span style=\"opacity:.75;\">\"They had previously visited that same city two years ago.\"</span></td></tr><tr><td>currently</td><td>happening right now<br><span style=\"opacity:.75;\">\"The family is currently living in a small apartment overseas.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Nonfiction Detour #1 · RI.4.2",
        "tasks": [
          {
            "id": "ar10",
            "label": "Week 10 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Nonfiction Detour #1 (new short article — Appendix D)</p><p><b>Standard Focus:</b> RI.4.2</p><p><b>Pacing Goal:</b> One short article this week.</p><p>Practice summarizing this week — telling the important parts of the article in just two or three sentences, in her own words, without going through it line by line.</p></div>"
          },
          {
            "id": "ard10",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Short Vowels &amp; Closed Syllables</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "sand",
                "kind": "word"
              },
              {
                "answer": "drift",
                "kind": "word"
              },
              {
                "answer": "skip",
                "kind": "word"
              },
              {
                "answer": "splash",
                "kind": "word"
              },
              {
                "answer": "insect",
                "kind": "word"
              },
              {
                "answer": "hamster",
                "kind": "word"
              },
              {
                "answer": "dentist",
                "kind": "word"
              },
              {
                "answer": "publish",
                "kind": "word"
              },
              {
                "answer": "The hamster sniffed at the sandy path.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar10q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "Can you tell me what this article said, in just three sentences?",
            "sampleAnswer": "‘Red pandas live in mountain forests in Asia. They eat mostly bamboo. Their reddish fur and long tails help them hide and balance in trees.’ A strong summary hits the key facts without extra detail."
          },
          {
            "id": "ar10q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "What's the most surprising fact you learned?",
            "sampleAnswer": "Any genuine, specific reaction counts — the goal is that she can point to something particular rather than say ‘I don't know.’"
          }
        ]
      }
    }
  },
  {
    "week_number": 11,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "MIT / MIS — send (Latin root)",
        "tasks": [
          {
            "id": "av11",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>MIT / MIS</b> — send (Latin root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>permit</td><td><i>per- (through) + mit (send)</i> — to allow something; an official document that gives permission<br><span style=\"opacity:.75;\">\"You need a permit to camp in this national park.\"</span></td></tr><tr><td>submit</td><td><i>sub- (under) + mit (send)</i> — to hand in or turn something in<br><span style=\"opacity:.75;\">\"Adelyn had to submit her essay before Friday.\"</span></td></tr><tr><td>dismiss</td><td><i>dis- (away) + miss (send)</i> — to send away or allow to leave<br><span style=\"opacity:.75;\">\"The teacher will dismiss the class when the bell rings.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>approximately</td><td>close to an exact amount; about<br><span style=\"opacity:.75;\">\"The hike took approximately three hours to finish.\"</span></td></tr><tr><td>examine</td><td>to look at something closely and carefully<br><span style=\"opacity:.75;\">\"The doctor examined Adelyn's ankle after her fall.\"</span></td></tr><tr><td>observe</td><td>to watch something carefully<br><span style=\"opacity:.75;\">\"They stopped to observe the birds building a nest.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Nonfiction Detour #1 · RI.4.4",
        "tasks": [
          {
            "id": "ar11",
            "label": "Week 11 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Nonfiction Detour #1 (new short article — Appendix D)</p><p><b>Standard Focus:</b> RI.4.4</p><p><b>Pacing Goal:</b> One short article this week.</p><p>Nonfiction often introduces special vocabulary specific to the topic. Hunt for one or two words that were new or tricky, and use context clues (the surrounding sentences) to guess the meaning before checking a dictionary.</p></div>"
          },
          {
            "id": "ard11",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Silent-E (Magic E)</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "brave",
                "kind": "word"
              },
              {
                "answer": "shine",
                "kind": "word"
              },
              {
                "answer": "close",
                "kind": "word"
              },
              {
                "answer": "escape",
                "kind": "word"
              },
              {
                "answer": "athlete",
                "kind": "word"
              },
              {
                "answer": "complete",
                "kind": "word"
              },
              {
                "answer": "mistake",
                "kind": "word"
              },
              {
                "answer": "suitcase",
                "kind": "word"
              },
              {
                "answer": "The brave athlete packed her suitcase.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar11q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "Find one word you didn't know. What clues in the sentence helped you guess its meaning?",
            "sampleAnswer": "If the word is ‘nocturnal,’ a sentence like ‘Because it is nocturnal, the animal sleeps all day and hunts at night’ gives strong context clues pointing to ‘active at night.’"
          },
          {
            "id": "ar11q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "Use that new word in a sentence of your own.",
            "sampleAnswer": "‘My cat is nocturnal because she sleeps all day and wants to play at 2 a.m.’"
          }
        ]
      }
    }
  },
  {
    "week_number": 12,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "VIS / VID — see (Latin root)",
        "tasks": [
          {
            "id": "av12",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>VIS / VID</b> — see (Latin root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>visible</td><td><i>vis (see) + -ible (able to be)</i> — able to be seen<br><span style=\"opacity:.75;\">\"From the rooftop, the entire old city was visible.\"</span></td></tr><tr><td>invisible</td><td><i>in- (not) + vis (see) + -ible (able to be)</i> — not able to be seen<br><span style=\"opacity:.75;\">\"The tiny insect was almost invisible against the green leaf.\"</span></td></tr><tr><td>vision</td><td><i>vis (see) + -ion (act of)</i> — the ability to see; a picture in the mind of a future goal<br><span style=\"opacity:.75;\">\"Adelyn's eye doctor checked her vision before the trip.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>compare</td><td>to look at two or more things to see how they are alike or different<br><span style=\"opacity:.75;\">\"Adelyn compared the price of mangoes at two different stalls.\"</span></td></tr><tr><td>contrast</td><td>a difference between two things<br><span style=\"opacity:.75;\">\"There was a clear contrast between the busy city and the quiet countryside.\"</span></td></tr><tr><td>illustrate</td><td>to explain or show something clearly, often with pictures or examples<br><span style=\"opacity:.75;\">\"The teacher used a diagram to illustrate how volcanoes erupt.\"</span></td></tr></table>"
          },
          {
            "id": "avq3",
            "label": "Review Quiz 3 — Weeks 9–12",
            "type": "graded-mc",
            "monthlyTest": true,
            "questions": [
              {
                "q": "Looking good in photographs",
                "options": [
                  "antibiotic",
                  "biodegradable",
                  "biology",
                  "dismiss",
                  "invisible",
                  "permit",
                  "photocopy",
                  "photogenic",
                  "photosynthesis",
                  "submit",
                  "visible",
                  "vision"
                ],
                "correct": 7
              },
              {
                "q": "The process plants use to turn sunlight into food",
                "options": [
                  "antibiotic",
                  "biodegradable",
                  "biology",
                  "dismiss",
                  "invisible",
                  "permit",
                  "photocopy",
                  "photogenic",
                  "photosynthesis",
                  "submit",
                  "visible",
                  "vision"
                ],
                "correct": 8
              },
              {
                "q": "A copy of a document made using a machine that uses light",
                "options": [
                  "antibiotic",
                  "biodegradable",
                  "biology",
                  "dismiss",
                  "invisible",
                  "permit",
                  "photocopy",
                  "photogenic",
                  "photosynthesis",
                  "submit",
                  "visible",
                  "vision"
                ],
                "correct": 6
              },
              {
                "q": "Able to break down naturally over time",
                "options": [
                  "antibiotic",
                  "biodegradable",
                  "biology",
                  "dismiss",
                  "invisible",
                  "permit",
                  "photocopy",
                  "photogenic",
                  "photosynthesis",
                  "submit",
                  "visible",
                  "vision"
                ],
                "correct": 1
              },
              {
                "q": "A medicine that fights harmful bacteria",
                "options": [
                  "antibiotic",
                  "biodegradable",
                  "biology",
                  "dismiss",
                  "invisible",
                  "permit",
                  "photocopy",
                  "photogenic",
                  "photosynthesis",
                  "submit",
                  "visible",
                  "vision"
                ],
                "correct": 0
              },
              {
                "q": "The scientific study of living things",
                "options": [
                  "antibiotic",
                  "biodegradable",
                  "biology",
                  "dismiss",
                  "invisible",
                  "permit",
                  "photocopy",
                  "photogenic",
                  "photosynthesis",
                  "submit",
                  "visible",
                  "vision"
                ],
                "correct": 2
              },
              {
                "q": "The ability to see; a picture in the mind of a future goal",
                "options": [
                  "antibiotic",
                  "biodegradable",
                  "biology",
                  "dismiss",
                  "invisible",
                  "permit",
                  "photocopy",
                  "photogenic",
                  "photosynthesis",
                  "submit",
                  "visible",
                  "vision"
                ],
                "correct": 11
              },
              {
                "q": "Not able to be seen",
                "options": [
                  "antibiotic",
                  "biodegradable",
                  "biology",
                  "dismiss",
                  "invisible",
                  "permit",
                  "photocopy",
                  "photogenic",
                  "photosynthesis",
                  "submit",
                  "visible",
                  "vision"
                ],
                "correct": 4
              },
              {
                "q": "Able to be seen",
                "options": [
                  "antibiotic",
                  "biodegradable",
                  "biology",
                  "dismiss",
                  "invisible",
                  "permit",
                  "photocopy",
                  "photogenic",
                  "photosynthesis",
                  "submit",
                  "visible",
                  "vision"
                ],
                "correct": 10
              },
              {
                "q": "To send away or allow to leave",
                "options": [
                  "antibiotic",
                  "biodegradable",
                  "biology",
                  "dismiss",
                  "invisible",
                  "permit",
                  "photocopy",
                  "photogenic",
                  "photosynthesis",
                  "submit",
                  "visible",
                  "vision"
                ],
                "correct": 3
              },
              {
                "q": "To hand in or turn something in",
                "options": [
                  "antibiotic",
                  "biodegradable",
                  "biology",
                  "dismiss",
                  "invisible",
                  "permit",
                  "photocopy",
                  "photogenic",
                  "photosynthesis",
                  "submit",
                  "visible",
                  "vision"
                ],
                "correct": 9
              },
              {
                "q": "To allow something; an official document that gives permission",
                "options": [
                  "antibiotic",
                  "biodegradable",
                  "biology",
                  "dismiss",
                  "invisible",
                  "permit",
                  "photocopy",
                  "photogenic",
                  "photosynthesis",
                  "submit",
                  "visible",
                  "vision"
                ],
                "correct": 5
              },
              {
                "q": "Happening right now",
                "options": [
                  "approximately",
                  "compare",
                  "constantly",
                  "contrast",
                  "currently",
                  "examine",
                  "frequently",
                  "illustrate",
                  "observe",
                  "previously",
                  "rarely",
                  "suddenly"
                ],
                "correct": 4
              },
              {
                "q": "Before now; earlier",
                "options": [
                  "approximately",
                  "compare",
                  "constantly",
                  "contrast",
                  "currently",
                  "examine",
                  "frequently",
                  "illustrate",
                  "observe",
                  "previously",
                  "rarely",
                  "suddenly"
                ],
                "correct": 9
              },
              {
                "q": "Happening quickly and unexpectedly",
                "options": [
                  "approximately",
                  "compare",
                  "constantly",
                  "contrast",
                  "currently",
                  "examine",
                  "frequently",
                  "illustrate",
                  "observe",
                  "previously",
                  "rarely",
                  "suddenly"
                ],
                "correct": 11
              },
              {
                "q": "Happening all the time, without stopping",
                "options": [
                  "approximately",
                  "compare",
                  "constantly",
                  "contrast",
                  "currently",
                  "examine",
                  "frequently",
                  "illustrate",
                  "observe",
                  "previously",
                  "rarely",
                  "suddenly"
                ],
                "correct": 2
              },
              {
                "q": "Often; many times",
                "options": [
                  "approximately",
                  "compare",
                  "constantly",
                  "contrast",
                  "currently",
                  "examine",
                  "frequently",
                  "illustrate",
                  "observe",
                  "previously",
                  "rarely",
                  "suddenly"
                ],
                "correct": 6
              },
              {
                "q": "Not often",
                "options": [
                  "approximately",
                  "compare",
                  "constantly",
                  "contrast",
                  "currently",
                  "examine",
                  "frequently",
                  "illustrate",
                  "observe",
                  "previously",
                  "rarely",
                  "suddenly"
                ],
                "correct": 10
              },
              {
                "q": "To explain or show something clearly, often with pictures or examples",
                "options": [
                  "approximately",
                  "compare",
                  "constantly",
                  "contrast",
                  "currently",
                  "examine",
                  "frequently",
                  "illustrate",
                  "observe",
                  "previously",
                  "rarely",
                  "suddenly"
                ],
                "correct": 7
              },
              {
                "q": "A difference between two things",
                "options": [
                  "approximately",
                  "compare",
                  "constantly",
                  "contrast",
                  "currently",
                  "examine",
                  "frequently",
                  "illustrate",
                  "observe",
                  "previously",
                  "rarely",
                  "suddenly"
                ],
                "correct": 3
              },
              {
                "q": "To look at two or more things to see how they are alike or different",
                "options": [
                  "approximately",
                  "compare",
                  "constantly",
                  "contrast",
                  "currently",
                  "examine",
                  "frequently",
                  "illustrate",
                  "observe",
                  "previously",
                  "rarely",
                  "suddenly"
                ],
                "correct": 1
              },
              {
                "q": "To watch something carefully",
                "options": [
                  "approximately",
                  "compare",
                  "constantly",
                  "contrast",
                  "currently",
                  "examine",
                  "frequently",
                  "illustrate",
                  "observe",
                  "previously",
                  "rarely",
                  "suddenly"
                ],
                "correct": 8
              },
              {
                "q": "To look at something closely and carefully",
                "options": [
                  "approximately",
                  "compare",
                  "constantly",
                  "contrast",
                  "currently",
                  "examine",
                  "frequently",
                  "illustrate",
                  "observe",
                  "previously",
                  "rarely",
                  "suddenly"
                ],
                "correct": 5
              },
              {
                "q": "Close to an exact amount; about",
                "options": [
                  "approximately",
                  "compare",
                  "constantly",
                  "contrast",
                  "currently",
                  "examine",
                  "frequently",
                  "illustrate",
                  "observe",
                  "previously",
                  "rarely",
                  "suddenly"
                ],
                "correct": 0
              }
            ]
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Isadora Moon Goes to School · RL.4.6 (preview)",
        "tasks": [
          {
            "id": "ar12",
            "label": "Week 12 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Isadora Moon Goes to School (begin)</p><p><b>Standard Focus:</b> RL.4.6 (preview)</p><p><b>Pacing Goal:</b> Read to roughly the 20% mark.</p><p>Start the new book this week — Isadora Moon Goes to School, book one of the series. Before we focus on the standard next month, just get to know Isadora and notice: is the story told by Isadora herself (“I”), or by an outside narrator (“she”)?</p></div>"
          },
          {
            "id": "ard12",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Vowel Teams</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "train",
                "kind": "word"
              },
              {
                "answer": "stay",
                "kind": "word"
              },
              {
                "answer": "bead",
                "kind": "word"
              },
              {
                "answer": "road",
                "kind": "word"
              },
              {
                "answer": "grow",
                "kind": "word"
              },
              {
                "answer": "chief",
                "kind": "word"
              },
              {
                "answer": "wait",
                "kind": "word"
              },
              {
                "answer": "sea",
                "kind": "word"
              },
              {
                "answer": "They will stay by the sea and wait for the train.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar12q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "Does the story use ‘I’ or ‘she’ to talk about Isadora? How can you tell?",
            "sampleAnswer": "This series is told in first person — ‘I’ shows up right from the first page, since Isadora is narrating her own story."
          },
          {
            "id": "ar12q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "What makes Isadora different from Harriet, Maisie, or Kitty?",
            "sampleAnswer": "She's half-fairy, half-vampire, which makes her feel like she doesn't quite fit into either world — a different kind of ‘different’ than the other heroines, who mostly stand out because of what they do rather than what they are."
          }
        ]
      }
    }
  },
  {
    "week_number": 13,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "AUD — hear (Latin root)",
        "tasks": [
          {
            "id": "av13",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>AUD</b> — hear (Latin root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>audience</td><td><i>aud (hear) + -ience (state of)</i> — a group of people watching or listening to a performance<br><span style=\"opacity:.75;\">\"The audience clapped loudly after the gymnastics routine.\"</span></td></tr><tr><td>audible</td><td><i>aud (hear) + -ible (able to be)</i> — loud enough to be heard<br><span style=\"opacity:.75;\">\"Her whisper was barely audible over the noisy market.\"</span></td></tr><tr><td>audio</td><td><i>aud (hear) + -io</i> — sound, especially recorded or broadcast sound<br><span style=\"opacity:.75;\">\"The museum offered an audio guide in five languages.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>summarize</td><td>to give the main points of something in a shorter form<br><span style=\"opacity:.75;\">\"Adelyn summarized the chapter in just three sentences.\"</span></td></tr><tr><td>conclude</td><td>to decide something after considering the facts; to end<br><span style=\"opacity:.75;\">\"After weighing both trails, they concluded the shorter one was better.\"</span></td></tr><tr><td>suggest</td><td>to offer an idea or plan for someone to consider<br><span style=\"opacity:.75;\">\"Kenley suggested they try the small bakery on the corner.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Isadora Moon Goes to School · RL.4.6",
        "tasks": [
          {
            "id": "ar13",
            "label": "Week 13 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Isadora Moon Goes to School</p><p><b>Standard Focus:</b> RL.4.6</p><p><b>Pacing Goal:</b> Read to roughly the 40% mark.</p><p>A story told in first person (“I saw…”) feels like we're inside the character's head. Third person (“She saw…”) feels like we're watching from outside. Confirm which one this book uses and talk about how it changes the feel of the story.</p></div>"
          },
          {
            "id": "ard13",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>R-Controlled Vowels</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "chart",
                "kind": "word"
              },
              {
                "answer": "farmer",
                "kind": "word"
              },
              {
                "answer": "third",
                "kind": "word"
              },
              {
                "answer": "fork",
                "kind": "word"
              },
              {
                "answer": "purse",
                "kind": "word"
              },
              {
                "answer": "winter",
                "kind": "word"
              },
              {
                "answer": "border",
                "kind": "word"
              },
              {
                "answer": "artist",
                "kind": "word"
              },
              {
                "answer": "The artist packed her chart before winter.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar13q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "Is this book written in first or third person? What word clues tell you?",
            "sampleAnswer": "First person — words like ‘I’ and ‘my’ show up constantly, since Isadora is telling us her own thoughts and feelings directly."
          },
          {
            "id": "ar13q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "How might the story feel different if it were told the other way?",
            "sampleAnswer": "In third person, we'd only see Isadora from the outside — we'd lose the sense of hearing her private worries about fitting in, which is a big part of what makes this book feel personal."
          }
        ]
      }
    }
  },
  {
    "week_number": 14,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "THERM — heat (Greek root)",
        "tasks": [
          {
            "id": "av14",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>THERM</b> — heat (Greek root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>thermometer</td><td><i>therm (heat) + meter (measure)</i> — a tool used to measure temperature<br><span style=\"opacity:.75;\">\"The thermometer showed it was over 90 degrees outside.\"</span></td></tr><tr><td>thermal</td><td><i>therm (heat) + -al (relating to)</i> — relating to heat<br><span style=\"opacity:.75;\">\"They wore thermal socks to stay warm on the mountain hike.\"</span></td></tr><tr><td>thermos</td><td><i>therm (heat)</i> — a container that keeps drinks hot or cold for a long time<br><span style=\"opacity:.75;\">\"Kenley packed hot soup in a thermos for the train ride.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>indicate</td><td>to show or point out<br><span style=\"opacity:.75;\">\"Dark clouds indicated that a storm was coming.\"</span></td></tr><tr><td>reveal</td><td>to make something known that was hidden<br><span style=\"opacity:.75;\">\"The old letter revealed a secret about the house's history.\"</span></td></tr><tr><td>demonstrate</td><td>to show clearly how something works or is done<br><span style=\"opacity:.75;\">\"The coach demonstrated the correct way to land a cartwheel.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Isadora Moon Goes to School · RL.4.6",
        "tasks": [
          {
            "id": "ar14",
            "label": "Week 14 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Isadora Moon Goes to School</p><p><b>Standard Focus:</b> RL.4.6</p><p><b>Pacing Goal:</b> Read to roughly the 60% mark.</p><p>Compare this book's narration style to Harriet, Maisie, or Kitty — were those first or third person too?</p></div>"
          },
          {
            "id": "ard14",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Blends &amp; Digraphs</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "shrink",
                "kind": "word"
              },
              {
                "answer": "thankful",
                "kind": "word"
              },
              {
                "answer": "splendid",
                "kind": "word"
              },
              {
                "answer": "chopstick",
                "kind": "word"
              },
              {
                "answer": "shepherd",
                "kind": "word"
              },
              {
                "answer": "throttle",
                "kind": "word"
              },
              {
                "answer": "sandwich",
                "kind": "word"
              },
              {
                "answer": "spring",
                "kind": "word"
              },
              {
                "answer": "Kenley made a splendid sandwich in spring.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar14q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "Was Maisie's book told in first or third person? What about Kitty's?",
            "sampleAnswer": "Maisie's book is told in third person (‘Maisie thought…’), and Kitty's books use third person too — so Isadora Moon actually stands out this year as one of the few first-person books."
          },
          {
            "id": "ar14q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "Which style do you like better as a reader — and why?",
            "sampleAnswer": "There's no wrong answer here — some readers like feeling ‘inside’ a character's head with first person, others prefer the wider view third person gives. The useful part is being able to explain why."
          }
        ]
      }
    }
  },
  {
    "week_number": 15,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "AUTO — self (Greek root)",
        "tasks": [
          {
            "id": "av15",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>AUTO</b> — self (Greek root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>automatic</td><td><i>auto- (self) + -matic (acting)</i> — happening by itself, without a person controlling it<br><span style=\"opacity:.75;\">\"The automatic doors opened as soon as they walked close.\"</span></td></tr><tr><td>automobile</td><td><i>auto- (self) + mobile (moving)</i> — a car; a self-moving vehicle<br><span style=\"opacity:.75;\">\"The family rented an automobile to explore the countryside.\"</span></td></tr><tr><td>autopilot</td><td><i>auto- (self) + pilot</i> — a system that flies a plane or steers a ship without a person controlling it directly<br><span style=\"opacity:.75;\">\"The pilot switched to autopilot once the plane reached cruising height.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>establish</td><td>to set up or create something that will last<br><span style=\"opacity:.75;\">\"The family worked to establish a daily homeschool routine while traveling.\"</span></td></tr><tr><td>maintain</td><td>to keep something going or in good condition<br><span style=\"opacity:.75;\">\"It takes practice to maintain flexibility in gymnastics.\"</span></td></tr><tr><td>require</td><td>to need something<br><span style=\"opacity:.75;\">\"The recipe requires fresh basil and ripe tomatoes.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Isadora Moon Goes to School · RL.4.5",
        "tasks": [
          {
            "id": "ar15",
            "label": "Week 15 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Isadora Moon Goes to School</p><p><b>Standard Focus:</b> RL.4.5</p><p><b>Pacing Goal:</b> Read to roughly the 80% mark.</p><p>Every story has a shape: a beginning that sets things up, a middle where a problem grows, and an end where it gets solved. Map out where we are in that shape.</p></div>"
          },
          {
            "id": "ard15",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Common Prefixes</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "replay",
                "kind": "word"
              },
              {
                "answer": "unwrap",
                "kind": "word"
              },
              {
                "answer": "disagree",
                "kind": "word"
              },
              {
                "answer": "pretest",
                "kind": "word"
              },
              {
                "answer": "misplace",
                "kind": "word"
              },
              {
                "answer": "nonsense",
                "kind": "word"
              },
              {
                "answer": "refill",
                "kind": "word"
              },
              {
                "answer": "unlock",
                "kind": "word"
              },
              {
                "answer": "Please unwrap and refill the snack bag.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar15q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What's the main problem Isadora is facing so far?",
            "sampleAnswer": "Likely something about starting school and figuring out where she fits in — whether to lean into being different or try to blend in with either the fairy or vampire side of her family."
          },
          {
            "id": "ar15q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "Do you think we're in the beginning, middle, or end of the story? What makes you think that?",
            "sampleAnswer": "If the problem is still unresolved and new complications keep showing up, that signals the middle — the beginning usually just sets up who's who and what the problem is."
          }
        ]
      }
    }
  },
  {
    "week_number": 16,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "GEO — earth (Greek root)",
        "tasks": [
          {
            "id": "av16",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>GEO</b> — earth (Greek root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>geography</td><td><i>geo- (earth) + graph (write)</i> — the study of Earth's land, features, and people<br><span style=\"opacity:.75;\">\"In geography class, Adelyn learned about the mountains of South America.\"</span></td></tr><tr><td>geology</td><td><i>geo- (earth) + -logy (study of)</i> — the study of rocks, minerals, and the structure of the Earth<br><span style=\"opacity:.75;\">\"The geology teacher brought in samples of volcanic rock.\"</span></td></tr><tr><td>geometry</td><td><i>geo- (earth) + -metry (measurement)</i> — the branch of math that studies shapes, sizes, and space<br><span style=\"opacity:.75;\">\"Adelyn used geometry to figure out the angles of the kite she built.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>accomplish</td><td>to successfully complete something<br><span style=\"opacity:.75;\">\"Adelyn accomplished her goal of learning ten new words a week.\"</span></td></tr><tr><td>attempt</td><td>to try to do something<br><span style=\"opacity:.75;\">\"She made a brave attempt at the new gymnastics trick.\"</span></td></tr><tr><td>succeed</td><td>to achieve a desired result<br><span style=\"opacity:.75;\">\"After many tries, Adelyn finally succeeded at the backflip.\"</span></td></tr></table>"
          },
          {
            "id": "avq4",
            "label": "Review Quiz 4 — Weeks 13–16",
            "type": "graded-mc",
            "monthlyTest": true,
            "questions": [
              {
                "q": "A container that keeps drinks hot or cold for a long time",
                "options": [
                  "audible",
                  "audience",
                  "audio",
                  "automatic",
                  "automobile",
                  "autopilot",
                  "geography",
                  "geology",
                  "geometry",
                  "thermal",
                  "thermometer",
                  "thermos"
                ],
                "correct": 11
              },
              {
                "q": "Relating to heat",
                "options": [
                  "audible",
                  "audience",
                  "audio",
                  "automatic",
                  "automobile",
                  "autopilot",
                  "geography",
                  "geology",
                  "geometry",
                  "thermal",
                  "thermometer",
                  "thermos"
                ],
                "correct": 9
              },
              {
                "q": "A tool used to measure temperature",
                "options": [
                  "audible",
                  "audience",
                  "audio",
                  "automatic",
                  "automobile",
                  "autopilot",
                  "geography",
                  "geology",
                  "geometry",
                  "thermal",
                  "thermometer",
                  "thermos"
                ],
                "correct": 10
              },
              {
                "q": "Sound, especially recorded or broadcast sound",
                "options": [
                  "audible",
                  "audience",
                  "audio",
                  "automatic",
                  "automobile",
                  "autopilot",
                  "geography",
                  "geology",
                  "geometry",
                  "thermal",
                  "thermometer",
                  "thermos"
                ],
                "correct": 2
              },
              {
                "q": "Loud enough to be heard",
                "options": [
                  "audible",
                  "audience",
                  "audio",
                  "automatic",
                  "automobile",
                  "autopilot",
                  "geography",
                  "geology",
                  "geometry",
                  "thermal",
                  "thermometer",
                  "thermos"
                ],
                "correct": 0
              },
              {
                "q": "A group of people watching or listening to a performance",
                "options": [
                  "audible",
                  "audience",
                  "audio",
                  "automatic",
                  "automobile",
                  "autopilot",
                  "geography",
                  "geology",
                  "geometry",
                  "thermal",
                  "thermometer",
                  "thermos"
                ],
                "correct": 1
              },
              {
                "q": "The branch of math that studies shapes, sizes, and space",
                "options": [
                  "audible",
                  "audience",
                  "audio",
                  "automatic",
                  "automobile",
                  "autopilot",
                  "geography",
                  "geology",
                  "geometry",
                  "thermal",
                  "thermometer",
                  "thermos"
                ],
                "correct": 8
              },
              {
                "q": "The study of rocks, minerals, and the structure of the Earth",
                "options": [
                  "audible",
                  "audience",
                  "audio",
                  "automatic",
                  "automobile",
                  "autopilot",
                  "geography",
                  "geology",
                  "geometry",
                  "thermal",
                  "thermometer",
                  "thermos"
                ],
                "correct": 7
              },
              {
                "q": "The study of Earth's land, features, and people",
                "options": [
                  "audible",
                  "audience",
                  "audio",
                  "automatic",
                  "automobile",
                  "autopilot",
                  "geography",
                  "geology",
                  "geometry",
                  "thermal",
                  "thermometer",
                  "thermos"
                ],
                "correct": 6
              },
              {
                "q": "A system that flies a plane or steers a ship without a person controlling it directly",
                "options": [
                  "audible",
                  "audience",
                  "audio",
                  "automatic",
                  "automobile",
                  "autopilot",
                  "geography",
                  "geology",
                  "geometry",
                  "thermal",
                  "thermometer",
                  "thermos"
                ],
                "correct": 5
              },
              {
                "q": "A car; a self-moving vehicle",
                "options": [
                  "audible",
                  "audience",
                  "audio",
                  "automatic",
                  "automobile",
                  "autopilot",
                  "geography",
                  "geology",
                  "geometry",
                  "thermal",
                  "thermometer",
                  "thermos"
                ],
                "correct": 4
              },
              {
                "q": "Happening by itself, without a person controlling it",
                "options": [
                  "audible",
                  "audience",
                  "audio",
                  "automatic",
                  "automobile",
                  "autopilot",
                  "geography",
                  "geology",
                  "geometry",
                  "thermal",
                  "thermometer",
                  "thermos"
                ],
                "correct": 3
              },
              {
                "q": "To show clearly how something works or is done",
                "options": [
                  "accomplish",
                  "attempt",
                  "conclude",
                  "demonstrate",
                  "establish",
                  "indicate",
                  "maintain",
                  "require",
                  "reveal",
                  "succeed",
                  "suggest",
                  "summarize"
                ],
                "correct": 3
              },
              {
                "q": "To make something known that was hidden",
                "options": [
                  "accomplish",
                  "attempt",
                  "conclude",
                  "demonstrate",
                  "establish",
                  "indicate",
                  "maintain",
                  "require",
                  "reveal",
                  "succeed",
                  "suggest",
                  "summarize"
                ],
                "correct": 8
              },
              {
                "q": "To show or point out",
                "options": [
                  "accomplish",
                  "attempt",
                  "conclude",
                  "demonstrate",
                  "establish",
                  "indicate",
                  "maintain",
                  "require",
                  "reveal",
                  "succeed",
                  "suggest",
                  "summarize"
                ],
                "correct": 5
              },
              {
                "q": "To offer an idea or plan for someone to consider",
                "options": [
                  "accomplish",
                  "attempt",
                  "conclude",
                  "demonstrate",
                  "establish",
                  "indicate",
                  "maintain",
                  "require",
                  "reveal",
                  "succeed",
                  "suggest",
                  "summarize"
                ],
                "correct": 10
              },
              {
                "q": "To decide something after considering the facts; to end",
                "options": [
                  "accomplish",
                  "attempt",
                  "conclude",
                  "demonstrate",
                  "establish",
                  "indicate",
                  "maintain",
                  "require",
                  "reveal",
                  "succeed",
                  "suggest",
                  "summarize"
                ],
                "correct": 2
              },
              {
                "q": "To give the main points of something in a shorter form",
                "options": [
                  "accomplish",
                  "attempt",
                  "conclude",
                  "demonstrate",
                  "establish",
                  "indicate",
                  "maintain",
                  "require",
                  "reveal",
                  "succeed",
                  "suggest",
                  "summarize"
                ],
                "correct": 11
              },
              {
                "q": "To achieve a desired result",
                "options": [
                  "accomplish",
                  "attempt",
                  "conclude",
                  "demonstrate",
                  "establish",
                  "indicate",
                  "maintain",
                  "require",
                  "reveal",
                  "succeed",
                  "suggest",
                  "summarize"
                ],
                "correct": 9
              },
              {
                "q": "To try to do something",
                "options": [
                  "accomplish",
                  "attempt",
                  "conclude",
                  "demonstrate",
                  "establish",
                  "indicate",
                  "maintain",
                  "require",
                  "reveal",
                  "succeed",
                  "suggest",
                  "summarize"
                ],
                "correct": 1
              },
              {
                "q": "To successfully complete something",
                "options": [
                  "accomplish",
                  "attempt",
                  "conclude",
                  "demonstrate",
                  "establish",
                  "indicate",
                  "maintain",
                  "require",
                  "reveal",
                  "succeed",
                  "suggest",
                  "summarize"
                ],
                "correct": 0
              },
              {
                "q": "To need something",
                "options": [
                  "accomplish",
                  "attempt",
                  "conclude",
                  "demonstrate",
                  "establish",
                  "indicate",
                  "maintain",
                  "require",
                  "reveal",
                  "succeed",
                  "suggest",
                  "summarize"
                ],
                "correct": 7
              },
              {
                "q": "To keep something going or in good condition",
                "options": [
                  "accomplish",
                  "attempt",
                  "conclude",
                  "demonstrate",
                  "establish",
                  "indicate",
                  "maintain",
                  "require",
                  "reveal",
                  "succeed",
                  "suggest",
                  "summarize"
                ],
                "correct": 6
              },
              {
                "q": "To set up or create something that will last",
                "options": [
                  "accomplish",
                  "attempt",
                  "conclude",
                  "demonstrate",
                  "establish",
                  "indicate",
                  "maintain",
                  "require",
                  "reveal",
                  "succeed",
                  "suggest",
                  "summarize"
                ],
                "correct": 4
              }
            ]
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Isadora Moon Goes to School · RL.4.5",
        "tasks": [
          {
            "id": "ar16",
            "label": "Week 16 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Isadora Moon Goes to School</p><p><b>Standard Focus:</b> RL.4.5</p><p><b>Pacing Goal:</b> Finish the book.</p><p>Finish the book this week and look at the whole shape of the story from start to finish.</p></div>"
          },
          {
            "id": "ard16",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Common Suffixes</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "vacation",
                "kind": "word"
              },
              {
                "answer": "tension",
                "kind": "word"
              },
              {
                "answer": "stretching",
                "kind": "word"
              },
              {
                "answer": "traveled",
                "kind": "word"
              },
              {
                "answer": "gracefully",
                "kind": "word"
              },
              {
                "answer": "playful",
                "kind": "word"
              },
              {
                "answer": "endless",
                "kind": "word"
              },
              {
                "answer": "celebration",
                "kind": "word"
              },
              {
                "answer": "They traveled gracefully through the vacation.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar16q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "How did the problem get solved by the end?",
            "sampleAnswer": "Likely something about Isadora learning to accept and even celebrate being different rather than needing to pick one side of her identity."
          },
          {
            "id": "ar16q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "If you were telling a friend the story in order, what are the three most important things that happened?",
            "sampleAnswer": "1) Isadora starts a new school and feels like she doesn't fit in. 2) She faces a challenge related to being half-fairy, half-vampire. 3) She finds a way to be herself, and it works out."
          }
        ]
      }
    }
  },
  {
    "week_number": 17,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "TRI- — three (Latin/Greek prefix)",
        "tasks": [
          {
            "id": "av17",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>TRI-</b> — three (Latin/Greek prefix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>triangle</td><td><i>tri- (three) + angle</i> — a shape with three sides and three angles<br><span style=\"opacity:.75;\">\"The roof of the little cabin was shaped like a triangle.\"</span></td></tr><tr><td>tricycle</td><td><i>tri- (three) + cycle (wheel)</i> — a vehicle with three wheels<br><span style=\"opacity:.75;\">\"Adelyn's little cousin loves riding his tricycle around the courtyard.\"</span></td></tr><tr><td>tripod</td><td><i>tri- (three) + pod (foot)</i> — a three-legged stand used to hold a camera steady<br><span style=\"opacity:.75;\">\"Dad set up the tripod to film Adelyn's cartwheel.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>struggle</td><td>to try hard to do something difficult<br><span style=\"opacity:.75;\">\"Kenley struggled at first to read the foreign menu.\"</span></td></tr><tr><td>hesitate</td><td>to pause because of uncertainty<br><span style=\"opacity:.75;\">\"Adelyn didn't hesitate before jumping into the pool.\"</span></td></tr><tr><td>wonder</td><td>to think about something with curiosity<br><span style=\"opacity:.75;\">\"Adelyn began to wonder how far away the stars really were.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Nonfiction Detour #2 · RI.4.5",
        "tasks": [
          {
            "id": "ar17",
            "label": "Week 17 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Nonfiction Detour #2 (short article)</p><p><b>Standard Focus:</b> RI.4.5</p><p><b>Pacing Goal:</b> One short article this week.</p><p>Nonfiction writers organize information in different ways — sometimes step-by-step, sometimes cause-and-effect, sometimes comparing two things. Figure out how this week's article is organized.</p></div>"
          },
          {
            "id": "ard17",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Syllable Division Patterns</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "basket",
                "kind": "word"
              },
              {
                "answer": "pilot",
                "kind": "word"
              },
              {
                "answer": "rapid",
                "kind": "word"
              },
              {
                "answer": "table",
                "kind": "word"
              },
              {
                "answer": "picnic",
                "kind": "word"
              },
              {
                "answer": "spider",
                "kind": "word"
              },
              {
                "answer": "model",
                "kind": "word"
              },
              {
                "answer": "candle",
                "kind": "word"
              },
              {
                "answer": "They set a picnic basket on the table.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar17q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "Is this article organized like a list of steps, a cause-and-effect chain, or a comparison? How can you tell?",
            "sampleAnswer": "Look for clue words — ‘first, next, then’ signal steps; ‘because, as a result’ signal cause-and-effect; ‘unlike, similarly’ signal a comparison."
          },
          {
            "id": "ar17q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "Why do you think the author chose to organize it that way?",
            "sampleAnswer": "A process, like how volcanoes form, makes the most sense told in steps, since order actually matters to understanding it."
          }
        ]
      }
    }
  },
  {
    "week_number": 18,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "UN- — not (prefix)",
        "tasks": [
          {
            "id": "av18",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>UN-</b> — not (prefix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>unhappy</td><td><i>un- (not) + happy</i> — not happy; sad<br><span style=\"opacity:.75;\">\"Adelyn felt unhappy when the rain canceled the picnic.\"</span></td></tr><tr><td>unable</td><td><i>un- (not) + able</i> — not able to do something<br><span style=\"opacity:.75;\">\"The team was unable to finish the routine because of the storm.\"</span></td></tr><tr><td>unfair</td><td><i>un- (not) + fair</i> — not fair or just<br><span style=\"opacity:.75;\">\"It felt unfair that the match was postponed twice.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>imagine</td><td>to form a picture or idea in your mind<br><span style=\"opacity:.75;\">\"Try to imagine what life was like a hundred years ago.\"</span></td></tr><tr><td>realize</td><td>to become aware of something<br><span style=\"opacity:.75;\">\"Adelyn suddenly realized she had left her water bottle at the market.\"</span></td></tr><tr><td>recall</td><td>to remember something<br><span style=\"opacity:.75;\">\"Kenley could still recall the taste of her grandmother's soup.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Nonfiction Detour #2 · RI.4.3",
        "tasks": [
          {
            "id": "ar18",
            "label": "Week 18 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Nonfiction Detour #2 (a how-to or process article — a recipe write-up works great)</p><p><b>Standard Focus:</b> RI.4.3</p><p><b>Pacing Goal:</b> One short article this week.</p><p>Try a “how” or “why” nonfiction piece this week — a simple recipe write-up is a fun tie to cooking with Kenley, or use a short science-process article. Practice explaining not just what happened, but why.</p></div>"
          },
          {
            "id": "ard18",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Homophones &amp; Confused Words</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "to",
                "kind": "word"
              },
              {
                "answer": "too",
                "kind": "word"
              },
              {
                "answer": "two",
                "kind": "word"
              },
              {
                "answer": "know",
                "kind": "word"
              },
              {
                "answer": "no",
                "kind": "word"
              },
              {
                "answer": "hear",
                "kind": "word"
              },
              {
                "answer": "here",
                "kind": "word"
              },
              {
                "answer": "I know there are two seats here.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar18q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What happens first, and what happens because of that?",
            "sampleAnswer": "‘First the butter is creamed with sugar, which is what lets air into the batter so the cookies turn out fluffy instead of dense.’"
          },
          {
            "id": "ar18q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "Can you explain the process in your own words, step by step?",
            "sampleAnswer": "A strong answer walks through the process in order without skipping steps, using her own phrasing rather than repeating the article word for word."
          }
        ]
      }
    }
  },
  {
    "week_number": 19,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "RE- — again, back (prefix)",
        "tasks": [
          {
            "id": "av19",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>RE-</b> — again, back (prefix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>rewrite</td><td><i>re- (again) + write</i> — to write something again, usually to improve it<br><span style=\"opacity:.75;\">\"Adelyn had to rewrite her paragraph to add more details.\"</span></td></tr><tr><td>replay</td><td><i>re- (again) + play</i> — to play something again, like a video<br><span style=\"opacity:.75;\">\"They watched a replay of the winning routine three times.\"</span></td></tr><tr><td>rebuild</td><td><i>re- (again) + build</i> — to build something again after it was broken or damaged<br><span style=\"opacity:.75;\">\"The village worked together to rebuild the bridge after the flood.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>remind</td><td>to help someone remember something<br><span style=\"opacity:.75;\">\"Mom set an alarm to remind them about the train.\"</span></td></tr><tr><td>remain</td><td>to stay in the same place or condition<br><span style=\"opacity:.75;\">\"Despite the chaos, Adelyn remained calm and focused.\"</span></td></tr><tr><td>continue</td><td>to keep doing something without stopping<br><span style=\"opacity:.75;\">\"They continued walking even after the sun went down.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Nonfiction Detour #2 · RI.4.6",
        "tasks": [
          {
            "id": "ar19",
            "label": "Week 19 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Nonfiction Detour #2 (firsthand vs. secondhand pair — Appendix D)</p><p><b>Standard Focus:</b> RI.4.6</p><p><b>Pacing Goal:</b> Two short pieces this week (a firsthand account and a secondhand entry on the same place).</p><p>A fun one for a traveling family: compare a firsthand account (a postcard, journal entry, or blog post from someone who was actually there) to a secondhand account (an encyclopedia entry or guidebook) about the same place. Notice how the focus and information differ.</p></div>"
          },
          {
            "id": "ard19",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Short Vowels &amp; Closed Syllables</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "crunch",
                "kind": "word"
              },
              {
                "answer": "stretch",
                "kind": "word"
              },
              {
                "answer": "twist",
                "kind": "word"
              },
              {
                "answer": "gymnast",
                "kind": "word"
              },
              {
                "answer": "contract",
                "kind": "word"
              },
              {
                "answer": "admit",
                "kind": "word"
              },
              {
                "answer": "absent",
                "kind": "word"
              },
              {
                "answer": "magnet",
                "kind": "word"
              },
              {
                "answer": "The gymnast did a quick twist and stretch.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar19q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What does the firsthand account tell us that the guidebook doesn't?",
            "sampleAnswer": "Personal details and feelings — e.g., a journal entry might describe how loud and crowded a market felt, or something funny that happened, which a guidebook's factual entry wouldn't include."
          },
          {
            "id": "ar19q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "Which one made you feel more like you were really there? Why?",
            "sampleAnswer": "Usually the firsthand account, because of the specific sensory details and personal voice — though a guidebook might give a clearer overall picture of the place."
          }
        ]
      }
    }
  },
  {
    "week_number": 20,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "PRE- — before (prefix)",
        "tasks": [
          {
            "id": "av20",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>PRE-</b> — before (prefix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>preview</td><td><i>pre- (before) + view</i> — a look at something before it fully happens<br><span style=\"opacity:.75;\">\"The class got a preview of next week's spelling words.\"</span></td></tr><tr><td>prepare</td><td><i>pre- (before) + pare (make ready)</i> — to get ready for something ahead of time<br><span style=\"opacity:.75;\">\"Kenley began to prepare dinner while Adelyn set the table.\"</span></td></tr><tr><td>preheat</td><td><i>pre- (before) + heat</i> — to heat something, like an oven, before using it<br><span style=\"opacity:.75;\">\"Remember to preheat the oven before baking the bread.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>reduce</td><td>to make something smaller or less<br><span style=\"opacity:.75;\">\"Packing light helped reduce the weight of their suitcases.\"</span></td></tr><tr><td>increase</td><td>to make or become greater<br><span style=\"opacity:.75;\">\"Daily stretching helped increase Adelyn's flexibility.\"</span></td></tr><tr><td>decrease</td><td>to make or become smaller<br><span style=\"opacity:.75;\">\"The noise began to decrease as the market closed for the night.\"</span></td></tr></table>"
          },
          {
            "id": "avq5",
            "label": "Review Quiz 5 — Weeks 17–20",
            "type": "graded-mc",
            "monthlyTest": true,
            "questions": [
              {
                "q": "Not fair or just",
                "options": [
                  "preheat",
                  "prepare",
                  "preview",
                  "rebuild",
                  "replay",
                  "rewrite",
                  "triangle",
                  "tricycle",
                  "tripod",
                  "unable",
                  "unfair",
                  "unhappy"
                ],
                "correct": 10
              },
              {
                "q": "Not able to do something",
                "options": [
                  "preheat",
                  "prepare",
                  "preview",
                  "rebuild",
                  "replay",
                  "rewrite",
                  "triangle",
                  "tricycle",
                  "tripod",
                  "unable",
                  "unfair",
                  "unhappy"
                ],
                "correct": 9
              },
              {
                "q": "Not happy; sad",
                "options": [
                  "preheat",
                  "prepare",
                  "preview",
                  "rebuild",
                  "replay",
                  "rewrite",
                  "triangle",
                  "tricycle",
                  "tripod",
                  "unable",
                  "unfair",
                  "unhappy"
                ],
                "correct": 11
              },
              {
                "q": "A three-legged stand used to hold a camera steady",
                "options": [
                  "preheat",
                  "prepare",
                  "preview",
                  "rebuild",
                  "replay",
                  "rewrite",
                  "triangle",
                  "tricycle",
                  "tripod",
                  "unable",
                  "unfair",
                  "unhappy"
                ],
                "correct": 8
              },
              {
                "q": "A vehicle with three wheels",
                "options": [
                  "preheat",
                  "prepare",
                  "preview",
                  "rebuild",
                  "replay",
                  "rewrite",
                  "triangle",
                  "tricycle",
                  "tripod",
                  "unable",
                  "unfair",
                  "unhappy"
                ],
                "correct": 7
              },
              {
                "q": "A shape with three sides and three angles",
                "options": [
                  "preheat",
                  "prepare",
                  "preview",
                  "rebuild",
                  "replay",
                  "rewrite",
                  "triangle",
                  "tricycle",
                  "tripod",
                  "unable",
                  "unfair",
                  "unhappy"
                ],
                "correct": 6
              },
              {
                "q": "To heat something, like an oven, before using it",
                "options": [
                  "preheat",
                  "prepare",
                  "preview",
                  "rebuild",
                  "replay",
                  "rewrite",
                  "triangle",
                  "tricycle",
                  "tripod",
                  "unable",
                  "unfair",
                  "unhappy"
                ],
                "correct": 0
              },
              {
                "q": "To get ready for something ahead of time",
                "options": [
                  "preheat",
                  "prepare",
                  "preview",
                  "rebuild",
                  "replay",
                  "rewrite",
                  "triangle",
                  "tricycle",
                  "tripod",
                  "unable",
                  "unfair",
                  "unhappy"
                ],
                "correct": 1
              },
              {
                "q": "A look at something before it fully happens",
                "options": [
                  "preheat",
                  "prepare",
                  "preview",
                  "rebuild",
                  "replay",
                  "rewrite",
                  "triangle",
                  "tricycle",
                  "tripod",
                  "unable",
                  "unfair",
                  "unhappy"
                ],
                "correct": 2
              },
              {
                "q": "To build something again after it was broken or damaged",
                "options": [
                  "preheat",
                  "prepare",
                  "preview",
                  "rebuild",
                  "replay",
                  "rewrite",
                  "triangle",
                  "tricycle",
                  "tripod",
                  "unable",
                  "unfair",
                  "unhappy"
                ],
                "correct": 3
              },
              {
                "q": "To play something again, like a video",
                "options": [
                  "preheat",
                  "prepare",
                  "preview",
                  "rebuild",
                  "replay",
                  "rewrite",
                  "triangle",
                  "tricycle",
                  "tripod",
                  "unable",
                  "unfair",
                  "unhappy"
                ],
                "correct": 4
              },
              {
                "q": "To write something again, usually to improve it",
                "options": [
                  "preheat",
                  "prepare",
                  "preview",
                  "rebuild",
                  "replay",
                  "rewrite",
                  "triangle",
                  "tricycle",
                  "tripod",
                  "unable",
                  "unfair",
                  "unhappy"
                ],
                "correct": 5
              },
              {
                "q": "To remember something",
                "options": [
                  "continue",
                  "decrease",
                  "hesitate",
                  "imagine",
                  "increase",
                  "realize",
                  "recall",
                  "reduce",
                  "remain",
                  "remind",
                  "struggle",
                  "wonder"
                ],
                "correct": 6
              },
              {
                "q": "To become aware of something",
                "options": [
                  "continue",
                  "decrease",
                  "hesitate",
                  "imagine",
                  "increase",
                  "realize",
                  "recall",
                  "reduce",
                  "remain",
                  "remind",
                  "struggle",
                  "wonder"
                ],
                "correct": 5
              },
              {
                "q": "To form a picture or idea in your mind",
                "options": [
                  "continue",
                  "decrease",
                  "hesitate",
                  "imagine",
                  "increase",
                  "realize",
                  "recall",
                  "reduce",
                  "remain",
                  "remind",
                  "struggle",
                  "wonder"
                ],
                "correct": 3
              },
              {
                "q": "To think about something with curiosity",
                "options": [
                  "continue",
                  "decrease",
                  "hesitate",
                  "imagine",
                  "increase",
                  "realize",
                  "recall",
                  "reduce",
                  "remain",
                  "remind",
                  "struggle",
                  "wonder"
                ],
                "correct": 11
              },
              {
                "q": "To pause because of uncertainty",
                "options": [
                  "continue",
                  "decrease",
                  "hesitate",
                  "imagine",
                  "increase",
                  "realize",
                  "recall",
                  "reduce",
                  "remain",
                  "remind",
                  "struggle",
                  "wonder"
                ],
                "correct": 2
              },
              {
                "q": "To try hard to do something difficult",
                "options": [
                  "continue",
                  "decrease",
                  "hesitate",
                  "imagine",
                  "increase",
                  "realize",
                  "recall",
                  "reduce",
                  "remain",
                  "remind",
                  "struggle",
                  "wonder"
                ],
                "correct": 10
              },
              {
                "q": "To make or become smaller",
                "options": [
                  "continue",
                  "decrease",
                  "hesitate",
                  "imagine",
                  "increase",
                  "realize",
                  "recall",
                  "reduce",
                  "remain",
                  "remind",
                  "struggle",
                  "wonder"
                ],
                "correct": 1
              },
              {
                "q": "To make or become greater",
                "options": [
                  "continue",
                  "decrease",
                  "hesitate",
                  "imagine",
                  "increase",
                  "realize",
                  "recall",
                  "reduce",
                  "remain",
                  "remind",
                  "struggle",
                  "wonder"
                ],
                "correct": 4
              },
              {
                "q": "To make something smaller or less",
                "options": [
                  "continue",
                  "decrease",
                  "hesitate",
                  "imagine",
                  "increase",
                  "realize",
                  "recall",
                  "reduce",
                  "remain",
                  "remind",
                  "struggle",
                  "wonder"
                ],
                "correct": 7
              },
              {
                "q": "To keep doing something without stopping",
                "options": [
                  "continue",
                  "decrease",
                  "hesitate",
                  "imagine",
                  "increase",
                  "realize",
                  "recall",
                  "reduce",
                  "remain",
                  "remind",
                  "struggle",
                  "wonder"
                ],
                "correct": 0
              },
              {
                "q": "To stay in the same place or condition",
                "options": [
                  "continue",
                  "decrease",
                  "hesitate",
                  "imagine",
                  "increase",
                  "realize",
                  "recall",
                  "reduce",
                  "remain",
                  "remind",
                  "struggle",
                  "wonder"
                ],
                "correct": 8
              },
              {
                "q": "To help someone remember something",
                "options": [
                  "continue",
                  "decrease",
                  "hesitate",
                  "imagine",
                  "increase",
                  "realize",
                  "recall",
                  "reduce",
                  "remain",
                  "remind",
                  "struggle",
                  "wonder"
                ],
                "correct": 9
              }
            ]
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Julieta and the Diamond Enigma · RL.4.4 (preview)",
        "tasks": [
          {
            "id": "ar20",
            "label": "Week 20 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Julieta and the Diamond Enigma (begin)</p><p><b>Standard Focus:</b> RL.4.4 (preview)</p><p><b>Pacing Goal:</b> Read to roughly the 20% mark.</p><p>Start Julieta and the Diamond Enigma this week. The book weaves in French and Spanish words, plus references to art history and Greek mythology. Keep a running list of any words or names you want to look up together — it's part of the fun of this book.</p></div>"
          },
          {
            "id": "ard20",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Silent-E (Magic E)</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "grape",
                "kind": "word"
              },
              {
                "answer": "slide",
                "kind": "word"
              },
              {
                "answer": "spice",
                "kind": "word"
              },
              {
                "answer": "arrive",
                "kind": "word"
              },
              {
                "answer": "delete",
                "kind": "word"
              },
              {
                "answer": "compare",
                "kind": "word"
              },
              {
                "answer": "cascade",
                "kind": "word"
              },
              {
                "answer": "mistake",
                "kind": "word"
              },
              {
                "answer": "They will arrive at the spice market soon.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar20q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What's one French or Spanish word from the book, and what does it mean?",
            "sampleAnswer": "This depends on which word she encounters — a strong answer names the actual word and gives its meaning based on context or a quick dictionary check, not just a guess."
          },
          {
            "id": "ar20q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "What do you think the ‘Diamond Enigma’ in the title means?",
            "sampleAnswer": "An ‘enigma’ is a mystery or puzzle, so the title is hinting that there's a mysterious diamond at the center of the plot that needs solving."
          }
        ]
      }
    }
  },
  {
    "week_number": 21,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "DIS- — not, opposite of (prefix)",
        "tasks": [
          {
            "id": "av21",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>DIS-</b> — not, opposite of (prefix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>disagree</td><td><i>dis- (not) + agree</i> — to have a different opinion<br><span style=\"opacity:.75;\">\"Adelyn and Kenley disagreed about which movie to watch.\"</span></td></tr><tr><td>dislike</td><td><i>dis- (not) + like</i> — to not like something<br><span style=\"opacity:.75;\">\"Adelyn used to dislike spicy food, but now she loves it.\"</span></td></tr><tr><td>disappear</td><td><i>dis- (opposite) + appear</i> — to go out of sight<br><span style=\"opacity:.75;\">\"The cat seemed to disappear behind the market stalls.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>improve</td><td>to make or become better<br><span style=\"opacity:.75;\">\"Adelyn's cartwheels improved with every week of practice.\"</span></td></tr><tr><td>prevent</td><td>to stop something from happening<br><span style=\"opacity:.75;\">\"Wearing a helmet can prevent serious injuries.\"</span></td></tr><tr><td>protect</td><td>to keep someone or something safe from harm<br><span style=\"opacity:.75;\">\"Sunscreen helps protect skin from sunburn.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Julieta and the Diamond Enigma · RL.4.1, RL.4.3 (Characterization Checkpoint #3)",
        "tasks": [
          {
            "id": "ar21",
            "label": "Week 21 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Julieta and the Diamond Enigma</p><p><b>Standard Focus:</b> RL.4.1, RL.4.3 (Characterization Checkpoint #3)</p><p><b>Pacing Goal:</b> Read to roughly the 40% mark.</p><p>Mysteries are the best genre for practicing inference — using clues plus what you already know to figure something out. This week also doubles as our third characterization check-in: notice how Julieta's own curiosity and persistence, not the narrator telling us so, drive her to investigate.</p></div>"
          },
          {
            "id": "ard21",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Vowel Teams</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "paint",
                "kind": "word"
              },
              {
                "answer": "tray",
                "kind": "word"
              },
              {
                "answer": "dream",
                "kind": "word"
              },
              {
                "answer": "roast",
                "kind": "word"
              },
              {
                "answer": "glow",
                "kind": "word"
              },
              {
                "answer": "thief",
                "kind": "word"
              },
              {
                "answer": "mail",
                "kind": "word"
              },
              {
                "answer": "cream",
                "kind": "word"
              },
              {
                "answer": "The mail cart carried a tray of cream.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar21q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What's one clue we've found so far, and what does it suggest?",
            "sampleAnswer": "This depends on where the plot is — a good answer names a specific clue, not just ‘something suspicious happened,’ and explains what it might point toward."
          },
          {
            "id": "ar21q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "What does Julieta do that shows she's persistent or brave, rather than the story just telling us so?",
            "sampleAnswer": "Likely something like continuing to investigate even after being told to stop or warned it's risky — action rather than just being labeled ‘brave’ by the narrator."
          }
        ]
      }
    }
  },
  {
    "week_number": 22,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "MIS- — wrongly (prefix)",
        "tasks": [
          {
            "id": "av22",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>MIS-</b> — wrongly (prefix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>mistake</td><td><i>mis- (wrongly) + take</i> — an error; something done incorrectly<br><span style=\"opacity:.75;\">\"It was an honest mistake to add sugar instead of salt.\"</span></td></tr><tr><td>misspell</td><td><i>mis- (wrongly) + spell</i> — to spell a word incorrectly<br><span style=\"opacity:.75;\">\"Adelyn used to misspell the word 'because,' but not anymore.\"</span></td></tr><tr><td>mislead</td><td><i>mis- (wrongly) + lead</i> — to give someone a false idea; to lead in the wrong direction<br><span style=\"opacity:.75;\">\"The confusing sign nearly misled the hikers down the wrong path.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>provide</td><td>to give someone something they need<br><span style=\"opacity:.75;\">\"The host family provided fresh bread every morning.\"</span></td></tr><tr><td>produce</td><td>to make or create something<br><span style=\"opacity:.75;\">\"The small farm produces enough vegetables for the whole village.\"</span></td></tr><tr><td>create</td><td>to make something new<br><span style=\"opacity:.75;\">\"Adelyn used old fabric scraps to create a colorful bag.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Julieta and the Diamond Enigma · RL.4.2",
        "tasks": [
          {
            "id": "ar22",
            "label": "Week 22 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Julieta and the Diamond Enigma</p><p><b>Standard Focus:</b> RL.4.2</p><p><b>Pacing Goal:</b> Read to roughly the 60% mark.</p><p>A theme is the bigger idea underneath the plot — not just ‘what happens’ but ‘what it's really about.’ Start thinking about what this story might be about deep down (trust, bravery, family, curiosity…).</p></div>"
          },
          {
            "id": "ard22",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>R-Controlled Vowels</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "park",
                "kind": "word"
              },
              {
                "answer": "harbor",
                "kind": "word"
              },
              {
                "answer": "stir",
                "kind": "word"
              },
              {
                "answer": "north",
                "kind": "word"
              },
              {
                "answer": "curl",
                "kind": "word"
              },
              {
                "answer": "letter",
                "kind": "word"
              },
              {
                "answer": "market",
                "kind": "word"
              },
              {
                "answer": "dirt",
                "kind": "word"
              },
              {
                "answer": "The harbor market was north of the park.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar22q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What do you think this story is really about, underneath the mystery plot?",
            "sampleAnswer": "Possible themes include trusting yourself, family loyalty, or the idea that curiosity and paying attention can solve problems adults overlook."
          },
          {
            "id": "ar22q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "What's a moment where Julieta shows bravery or determination?",
            "sampleAnswer": "Any moment where she keeps going despite obstacles or fear counts — the important part is pointing to a specific scene, not just asserting she's brave."
          }
        ]
      }
    }
  },
  {
    "week_number": 23,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "SUB- — under (prefix)",
        "tasks": [
          {
            "id": "av23",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>SUB-</b> — under (prefix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>submarine</td><td><i>sub- (under) + marine (sea)</i> — a ship that can travel underwater<br><span style=\"opacity:.75;\">\"They watched a documentary about a submarine exploring the deep ocean.\"</span></td></tr><tr><td>subway</td><td><i>sub- (under) + way</i> — an underground train system<br><span style=\"opacity:.75;\">\"They took the subway to get across the busy city quickly.\"</span></td></tr><tr><td>substitute</td><td><i>sub- (in place of) + statute (set up)</i> — a person or thing that takes the place of another<br><span style=\"opacity:.75;\">\"They used honey as a substitute for sugar in the recipe.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>develop</td><td>to grow or become more advanced over time<br><span style=\"opacity:.75;\">\"Reading every night helped Adelyn develop a bigger vocabulary.\"</span></td></tr><tr><td>discover</td><td>to find something for the first time<br><span style=\"opacity:.75;\">\"They discovered a hidden waterfall behind the trees.\"</span></td></tr><tr><td>explore</td><td>to travel through a place to learn about it<br><span style=\"opacity:.75;\">\"The family loves to explore new cities on foot.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Julieta and the Diamond Enigma · RL.4.2",
        "tasks": [
          {
            "id": "ar23",
            "label": "Week 23 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Julieta and the Diamond Enigma</p><p><b>Standard Focus:</b> RL.4.2</p><p><b>Pacing Goal:</b> Read to roughly the 80% mark.</p><p>Practice backing up a theme idea with evidence from the story — a good theme statement can always point to a specific moment that supports it.</p></div>"
          },
          {
            "id": "ard23",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Blends &amp; Digraphs</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "strong",
                "kind": "word"
              },
              {
                "answer": "shatter",
                "kind": "word"
              },
              {
                "answer": "chicken",
                "kind": "word"
              },
              {
                "answer": "thistle",
                "kind": "word"
              },
              {
                "answer": "spinach",
                "kind": "word"
              },
              {
                "answer": "scratchy",
                "kind": "word"
              },
              {
                "answer": "whistle",
                "kind": "word"
              },
              {
                "answer": "splash",
                "kind": "word"
              },
              {
                "answer": "The chicken and spinach recipe needs a splash of lemon.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar23q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "You said the theme might be about ___. What's a scene that shows that?",
            "sampleAnswer": "Whatever theme she named earlier, a strong answer points back to a specific event that demonstrates it, not just a restatement of the theme in different words."
          },
          {
            "id": "ar23q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "How does Julieta's relationship with her dad connect to that theme?",
            "sampleAnswer": "If the theme is about trust or family, her dad's involvement in the mystery might be exactly what tests or proves that trust."
          }
        ]
      }
    }
  },
  {
    "week_number": 24,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "INTER- — between (prefix)",
        "tasks": [
          {
            "id": "av24",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>INTER-</b> — between (prefix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>international</td><td><i>inter- (between) + national (nation)</i> — involving two or more countries<br><span style=\"opacity:.75;\">\"Adelyn attends an international school with kids from many countries.\"</span></td></tr><tr><td>interact</td><td><i>inter- (between) + act</i> — to communicate or do things together with others<br><span style=\"opacity:.75;\">\"Adelyn loves to interact with kids from different countries while traveling.\"</span></td></tr><tr><td>interview</td><td><i>inter- (between) + view</i> — a meeting where one person asks another questions<br><span style=\"opacity:.75;\">\"A local reporter asked to interview the young gymnastics champion.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>investigate</td><td>to carefully look into something to learn the facts<br><span style=\"opacity:.75;\">\"The kids investigated the strange noise coming from the attic.\"</span></td></tr><tr><td>research</td><td>to study a subject carefully to learn new facts<br><span style=\"opacity:.75;\">\"Adelyn did research on the country before they visited.\"</span></td></tr><tr><td>analyze</td><td>to study something carefully in order to understand it<br><span style=\"opacity:.75;\">\"Scientists analyze rock samples to learn about the Earth's history.\"</span></td></tr></table>"
          },
          {
            "id": "avq6",
            "label": "Review Quiz 6 — Weeks 21–24",
            "type": "graded-mc",
            "monthlyTest": true,
            "questions": [
              {
                "q": "To give someone a false idea; to lead in the wrong direction",
                "options": [
                  "disagree",
                  "disappear",
                  "dislike",
                  "interact",
                  "international",
                  "interview",
                  "mislead",
                  "misspell",
                  "mistake",
                  "submarine",
                  "substitute",
                  "subway"
                ],
                "correct": 6
              },
              {
                "q": "To spell a word incorrectly",
                "options": [
                  "disagree",
                  "disappear",
                  "dislike",
                  "interact",
                  "international",
                  "interview",
                  "mislead",
                  "misspell",
                  "mistake",
                  "submarine",
                  "substitute",
                  "subway"
                ],
                "correct": 7
              },
              {
                "q": "An error; something done incorrectly",
                "options": [
                  "disagree",
                  "disappear",
                  "dislike",
                  "interact",
                  "international",
                  "interview",
                  "mislead",
                  "misspell",
                  "mistake",
                  "submarine",
                  "substitute",
                  "subway"
                ],
                "correct": 8
              },
              {
                "q": "To go out of sight",
                "options": [
                  "disagree",
                  "disappear",
                  "dislike",
                  "interact",
                  "international",
                  "interview",
                  "mislead",
                  "misspell",
                  "mistake",
                  "submarine",
                  "substitute",
                  "subway"
                ],
                "correct": 1
              },
              {
                "q": "To not like something",
                "options": [
                  "disagree",
                  "disappear",
                  "dislike",
                  "interact",
                  "international",
                  "interview",
                  "mislead",
                  "misspell",
                  "mistake",
                  "submarine",
                  "substitute",
                  "subway"
                ],
                "correct": 2
              },
              {
                "q": "To have a different opinion",
                "options": [
                  "disagree",
                  "disappear",
                  "dislike",
                  "interact",
                  "international",
                  "interview",
                  "mislead",
                  "misspell",
                  "mistake",
                  "submarine",
                  "substitute",
                  "subway"
                ],
                "correct": 0
              },
              {
                "q": "A meeting where one person asks another questions",
                "options": [
                  "disagree",
                  "disappear",
                  "dislike",
                  "interact",
                  "international",
                  "interview",
                  "mislead",
                  "misspell",
                  "mistake",
                  "submarine",
                  "substitute",
                  "subway"
                ],
                "correct": 5
              },
              {
                "q": "To communicate or do things together with others",
                "options": [
                  "disagree",
                  "disappear",
                  "dislike",
                  "interact",
                  "international",
                  "interview",
                  "mislead",
                  "misspell",
                  "mistake",
                  "submarine",
                  "substitute",
                  "subway"
                ],
                "correct": 3
              },
              {
                "q": "Involving two or more countries",
                "options": [
                  "disagree",
                  "disappear",
                  "dislike",
                  "interact",
                  "international",
                  "interview",
                  "mislead",
                  "misspell",
                  "mistake",
                  "submarine",
                  "substitute",
                  "subway"
                ],
                "correct": 4
              },
              {
                "q": "A person or thing that takes the place of another",
                "options": [
                  "disagree",
                  "disappear",
                  "dislike",
                  "interact",
                  "international",
                  "interview",
                  "mislead",
                  "misspell",
                  "mistake",
                  "submarine",
                  "substitute",
                  "subway"
                ],
                "correct": 10
              },
              {
                "q": "An underground train system",
                "options": [
                  "disagree",
                  "disappear",
                  "dislike",
                  "interact",
                  "international",
                  "interview",
                  "mislead",
                  "misspell",
                  "mistake",
                  "submarine",
                  "substitute",
                  "subway"
                ],
                "correct": 11
              },
              {
                "q": "A ship that can travel underwater",
                "options": [
                  "disagree",
                  "disappear",
                  "dislike",
                  "interact",
                  "international",
                  "interview",
                  "mislead",
                  "misspell",
                  "mistake",
                  "submarine",
                  "substitute",
                  "subway"
                ],
                "correct": 9
              },
              {
                "q": "To make something new",
                "options": [
                  "analyze",
                  "create",
                  "develop",
                  "discover",
                  "explore",
                  "improve",
                  "investigate",
                  "prevent",
                  "produce",
                  "protect",
                  "provide",
                  "research"
                ],
                "correct": 1
              },
              {
                "q": "To make or create something",
                "options": [
                  "analyze",
                  "create",
                  "develop",
                  "discover",
                  "explore",
                  "improve",
                  "investigate",
                  "prevent",
                  "produce",
                  "protect",
                  "provide",
                  "research"
                ],
                "correct": 8
              },
              {
                "q": "To give someone something they need",
                "options": [
                  "analyze",
                  "create",
                  "develop",
                  "discover",
                  "explore",
                  "improve",
                  "investigate",
                  "prevent",
                  "produce",
                  "protect",
                  "provide",
                  "research"
                ],
                "correct": 10
              },
              {
                "q": "To keep someone or something safe from harm",
                "options": [
                  "analyze",
                  "create",
                  "develop",
                  "discover",
                  "explore",
                  "improve",
                  "investigate",
                  "prevent",
                  "produce",
                  "protect",
                  "provide",
                  "research"
                ],
                "correct": 9
              },
              {
                "q": "To stop something from happening",
                "options": [
                  "analyze",
                  "create",
                  "develop",
                  "discover",
                  "explore",
                  "improve",
                  "investigate",
                  "prevent",
                  "produce",
                  "protect",
                  "provide",
                  "research"
                ],
                "correct": 7
              },
              {
                "q": "To make or become better",
                "options": [
                  "analyze",
                  "create",
                  "develop",
                  "discover",
                  "explore",
                  "improve",
                  "investigate",
                  "prevent",
                  "produce",
                  "protect",
                  "provide",
                  "research"
                ],
                "correct": 5
              },
              {
                "q": "To study something carefully in order to understand it",
                "options": [
                  "analyze",
                  "create",
                  "develop",
                  "discover",
                  "explore",
                  "improve",
                  "investigate",
                  "prevent",
                  "produce",
                  "protect",
                  "provide",
                  "research"
                ],
                "correct": 0
              },
              {
                "q": "To study a subject carefully to learn new facts",
                "options": [
                  "analyze",
                  "create",
                  "develop",
                  "discover",
                  "explore",
                  "improve",
                  "investigate",
                  "prevent",
                  "produce",
                  "protect",
                  "provide",
                  "research"
                ],
                "correct": 11
              },
              {
                "q": "To carefully look into something to learn the facts",
                "options": [
                  "analyze",
                  "create",
                  "develop",
                  "discover",
                  "explore",
                  "improve",
                  "investigate",
                  "prevent",
                  "produce",
                  "protect",
                  "provide",
                  "research"
                ],
                "correct": 6
              },
              {
                "q": "To travel through a place to learn about it",
                "options": [
                  "analyze",
                  "create",
                  "develop",
                  "discover",
                  "explore",
                  "improve",
                  "investigate",
                  "prevent",
                  "produce",
                  "protect",
                  "provide",
                  "research"
                ],
                "correct": 4
              },
              {
                "q": "To find something for the first time",
                "options": [
                  "analyze",
                  "create",
                  "develop",
                  "discover",
                  "explore",
                  "improve",
                  "investigate",
                  "prevent",
                  "produce",
                  "protect",
                  "provide",
                  "research"
                ],
                "correct": 3
              },
              {
                "q": "To grow or become more advanced over time",
                "options": [
                  "analyze",
                  "create",
                  "develop",
                  "discover",
                  "explore",
                  "improve",
                  "investigate",
                  "prevent",
                  "produce",
                  "protect",
                  "provide",
                  "research"
                ],
                "correct": 2
              }
            ]
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Julieta and the Diamond Enigma · RL.4.4",
        "tasks": [
          {
            "id": "ar24",
            "label": "Week 24 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Julieta and the Diamond Enigma</p><p><b>Standard Focus:</b> RL.4.4</p><p><b>Pacing Goal:</b> Finish the book.</p><p>Finish the book this week and circle back to vocabulary — the art, mythology, and multilingual words woven throughout.</p></div>"
          },
          {
            "id": "ard24",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Common Prefixes</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "rebuild",
                "kind": "word"
              },
              {
                "answer": "unhappy",
                "kind": "word"
              },
              {
                "answer": "dislike",
                "kind": "word"
              },
              {
                "answer": "prepay",
                "kind": "word"
              },
              {
                "answer": "misread",
                "kind": "word"
              },
              {
                "answer": "nonfiction",
                "kind": "word"
              },
              {
                "answer": "restart",
                "kind": "word"
              },
              {
                "answer": "uneven",
                "kind": "word"
              },
              {
                "answer": "She had to restart and reread the nonfiction page.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar24q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What's your favorite new word or fact you learned from this book?",
            "sampleAnswer": "Any specific, genuine answer works — the goal is that she can name something particular rather than say ‘I don't know.’"
          },
          {
            "id": "ar24q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "How was the mystery solved? Did it surprise you?",
            "sampleAnswer": "A strong answer explains the actual resolution in her own words and gives a genuine reaction, not just ‘yes’ or ‘no.’"
          }
        ]
      }
    }
  },
  {
    "week_number": 25,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "-FUL — full of (suffix)",
        "tasks": [
          {
            "id": "av25",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>-FUL</b> — full of (suffix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>joyful</td><td><i>joy + -ful (full of)</i> — full of joy or happiness<br><span style=\"opacity:.75;\">\"The whole family felt joyful after the reunion.\"</span></td></tr><tr><td>careful</td><td><i>care + -ful (full of)</i> — paying close attention; full of care<br><span style=\"opacity:.75;\">\"Be careful on the slippery cobblestone streets.\"</span></td></tr><tr><td>helpful</td><td><i>help + -ful (full of)</i> — giving help; useful<br><span style=\"opacity:.75;\">\"The friendly shopkeeper was very helpful with directions.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>calculate</td><td>to figure out an answer using math<br><span style=\"opacity:.75;\">\"Adelyn calculated how many days were left until their next flight.\"</span></td></tr><tr><td>estimate</td><td>to make a careful guess about an amount<br><span style=\"opacity:.75;\">\"Kenley estimated how much flour the recipe would need.\"</span></td></tr><tr><td>measure</td><td>to find the size, length, or amount of something<br><span style=\"opacity:.75;\">\"They used a tape measure to check the height of the doorway.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Sleeping Beauty · RL.4.3 (Characterization Checkpoint #4) / RL.4.9 (preview)",
        "tasks": [
          {
            "id": "ar25",
            "label": "Week 25 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Sleeping Beauty (a gentle traditional retelling, e.g. Ladybird Tales) + Harriet the Invincible (revisited)</p><p><b>Standard Focus:</b> RL.4.3 (Characterization Checkpoint #4) / RL.4.9 (preview)</p><p><b>Pacing Goal:</b> Sleeping Beauty is a single-sitting picture book; Harriet needs no rereading — you already know it well.</p><p>Full circle: Harriet the Invincible's whole curse-and-hundred-year-sleep premise is a spunky twist on Sleeping Beauty. Read (or reread) the traditional tale — a princess cursed to prick her finger and sleep, saved when the thorns simply part for a prince — then talk about how differently Harriet handled the exact same curse back in Month 1. As a fourth characterization check-in, revisit what made Harriet Harriet.</p></div>"
          },
          {
            "id": "ard25",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Common Suffixes</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "attention",
                "kind": "word"
              },
              {
                "answer": "decision",
                "kind": "word"
              },
              {
                "answer": "cooking",
                "kind": "word"
              },
              {
                "answer": "practiced",
                "kind": "word"
              },
              {
                "answer": "happily",
                "kind": "word"
              },
              {
                "answer": "joyful",
                "kind": "word"
              },
              {
                "answer": "colorless",
                "kind": "word"
              },
              {
                "answer": "education",
                "kind": "word"
              },
              {
                "answer": "Kenley happily practiced her cooking for the celebration.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar25q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "How is Harriet's response to her curse different from the traditional sleeping princess's?",
            "sampleAnswer": "The traditional princess is passive — she waits, sleeps, and gets rescued. Harriet actively fights the curse, seeking out dangerous quests to prove she's tough enough to face anything, including her own fate."
          },
          {
            "id": "ar25q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "What does Harriet do about her curse that shows her personality, rather than just waiting for a prince?",
            "sampleAnswer": "She goes looking for danger on purpose instead of avoiding it — that stubborn, do-it-herself attitude is very different from quietly waiting to be saved."
          }
        ]
      }
    }
  },
  {
    "week_number": 26,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "-LESS — without (suffix)",
        "tasks": [
          {
            "id": "av26",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>-LESS</b> — without (suffix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>careless</td><td><i>care + -less (without)</i> — not paying attention; done without care<br><span style=\"opacity:.75;\">\"A careless mistake caused the tower of blocks to fall.\"</span></td></tr><tr><td>fearless</td><td><i>fear + -less (without)</i> — without fear; very brave<br><span style=\"opacity:.75;\">\"The fearless gymnast attempted the trick without hesitation.\"</span></td></tr><tr><td>harmless</td><td><i>harm + -less (without)</i> — not able or likely to cause harm<br><span style=\"opacity:.75;\">\"The garden snake was harmless and quickly slithered away.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>include</td><td>to have something as part of a group<br><span style=\"opacity:.75;\">\"The tour price includes lunch and a boat ride.\"</span></td></tr><tr><td>exclude</td><td>to leave something out; to not include<br><span style=\"opacity:.75;\">\"The recipe excludes nuts because of Kenley's allergy.\"</span></td></tr><tr><td>combine</td><td>to join two or more things together<br><span style=\"opacity:.75;\">\"Combine the flour and water to make the dough.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Sleeping Beauty + Harriet the Invincible · RL.4.9",
        "tasks": [
          {
            "id": "ar26",
            "label": "Week 26 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Sleeping Beauty + Harriet the Invincible</p><p><b>Standard Focus:</b> RL.4.9</p><p><b>Pacing Goal:</b> No new reading — comparison discussion.</p><p>Now compare the two directly: same starting curse (prick a finger, fall into a deep sleep), completely different heroine response. Talk through what's the same and what's changed.</p></div>"
          },
          {
            "id": "ard26",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Syllable Division Patterns</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "cactus",
                "kind": "word"
              },
              {
                "answer": "hotel",
                "kind": "word"
              },
              {
                "answer": "seven",
                "kind": "word"
              },
              {
                "answer": "bubble",
                "kind": "word"
              },
              {
                "answer": "gadget",
                "kind": "word"
              },
              {
                "answer": "tiny",
                "kind": "word"
              },
              {
                "answer": "planet",
                "kind": "word"
              },
              {
                "answer": "gentle",
                "kind": "word"
              },
              {
                "answer": "The tiny gadget looked like a little planet.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar26q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What's the same between the traditional Sleeping Beauty and Harriet's story?",
            "sampleAnswer": "Both start with a curse triggered by pricking a finger on a spinning wheel (or something similar), leading to a deep sleep."
          },
          {
            "id": "ar26q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "What's different — and why do you think the Hamster Princess author changed it?",
            "sampleAnswer": "Harriet takes action instead of waiting — the author likely changed it to give young readers a heroine who solves her own problems instead of needing to be rescued."
          }
        ]
      }
    }
  },
  {
    "week_number": 27,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "-ABLE / -IBLE — capable of (suffix)",
        "tasks": [
          {
            "id": "av27",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>-ABLE / -IBLE</b> — capable of (suffix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>comfortable</td><td><i>comfort + -able (capable of)</i> — providing comfort; at ease<br><span style=\"opacity:.75;\">\"The old hammock was surprisingly comfortable.\"</span></td></tr><tr><td>flexible</td><td><i>flex + -ible (capable of)</i> — able to bend or change easily<br><span style=\"opacity:.75;\">\"Years of gymnastics made Adelyn incredibly flexible.\"</span></td></tr><tr><td>reliable</td><td><i>rely + -able (capable of)</i> — able to be trusted or depended on<br><span style=\"opacity:.75;\">\"Their old van turned out to be a reliable way to travel.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>separate</td><td>to divide or keep apart<br><span style=\"opacity:.75;\">\"They separated the laundry into light and dark colors.\"</span></td></tr><tr><td>divide</td><td>to split something into parts<br><span style=\"opacity:.75;\">\"Adelyn divided the mangoes evenly between her and Kenley.\"</span></td></tr><tr><td>connect</td><td>to join or link together<br><span style=\"opacity:.75;\">\"A long bridge connects the two islands.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Sleeping Beauty + Harriet the Invincible · RL.4.9",
        "tasks": [
          {
            "id": "ar27",
            "label": "Week 27 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Sleeping Beauty + Harriet the Invincible (wrap-up)</p><p><b>Standard Focus:</b> RL.4.9</p><p><b>Pacing Goal:</b> No new reading — wrap-up discussion.</p><p>Fairy tales get retold in almost every culture, often with a twist that fits the storyteller's own values. Talk about what Harriet's version says about waiting to be rescued versus taking action.</p></div>"
          },
          {
            "id": "ard27",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Homophones &amp; Confused Words</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "which",
                "kind": "word"
              },
              {
                "answer": "witch",
                "kind": "word"
              },
              {
                "answer": "right",
                "kind": "word"
              },
              {
                "answer": "write",
                "kind": "word"
              },
              {
                "answer": "meat",
                "kind": "word"
              },
              {
                "answer": "meet",
                "kind": "word"
              },
              {
                "answer": "Write down which witch you will meet.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar27q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What lesson does the traditional Sleeping Beauty seem to teach? What lesson does Harriet's story teach instead?",
            "sampleAnswer": "The traditional tale suggests patience and waiting are rewarded; Harriet's version suggests that facing your fears head-on is more empowering than waiting passively."
          },
          {
            "id": "ar27q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "If you were cursed like Sleeping Beauty, would you rather wait for rescue or do what Harriet did?",
            "sampleAnswer": "There's no wrong answer — the value is in her reasoning, e.g., ‘I'd want to do what Harriet did because waiting sounds boring and scary.’"
          }
        ]
      }
    }
  },
  {
    "week_number": 28,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "-TION / -SION — act or state of (suffix)",
        "tasks": [
          {
            "id": "av28",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>-TION / -SION</b> — act or state of (suffix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>celebration</td><td><i>celebrate + -tion (act of)</i> — an event or activity to honor a special occasion<br><span style=\"opacity:.75;\">\"The whole village joined the celebration for the harvest festival.\"</span></td></tr><tr><td>decision</td><td><i>decide + -sion (act of)</i> — a choice made after thinking it over<br><span style=\"opacity:.75;\">\"Choosing which country to visit next was a tough decision.\"</span></td></tr><tr><td>invention</td><td><i>invent + -ion (act of)</i> — something newly created or designed<br><span style=\"opacity:.75;\">\"The lightbulb was one of the most useful inventions in history.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>relate</td><td>to show a connection between two things<br><span style=\"opacity:.75;\">\"The story relates closely to something that happened to Adelyn.\"</span></td></tr><tr><td>depend</td><td>to rely on someone or something<br><span style=\"opacity:.75;\">\"The whole trip depended on good weather.\"</span></td></tr><tr><td>rely</td><td>to trust and depend on someone or something<br><span style=\"opacity:.75;\">\"Adelyn learned to rely on herself when trying something new.\"</span></td></tr></table>"
          },
          {
            "id": "avq7",
            "label": "Review Quiz 7 — Weeks 25–28",
            "type": "graded-mc",
            "monthlyTest": true,
            "questions": [
              {
                "q": "Not able or likely to cause harm",
                "options": [
                  "careful",
                  "careless",
                  "celebration",
                  "comfortable",
                  "decision",
                  "fearless",
                  "flexible",
                  "harmless",
                  "helpful",
                  "invention",
                  "joyful",
                  "reliable"
                ],
                "correct": 7
              },
              {
                "q": "Without fear; very brave",
                "options": [
                  "careful",
                  "careless",
                  "celebration",
                  "comfortable",
                  "decision",
                  "fearless",
                  "flexible",
                  "harmless",
                  "helpful",
                  "invention",
                  "joyful",
                  "reliable"
                ],
                "correct": 5
              },
              {
                "q": "Not paying attention; done without care",
                "options": [
                  "careful",
                  "careless",
                  "celebration",
                  "comfortable",
                  "decision",
                  "fearless",
                  "flexible",
                  "harmless",
                  "helpful",
                  "invention",
                  "joyful",
                  "reliable"
                ],
                "correct": 1
              },
              {
                "q": "Giving help; useful",
                "options": [
                  "careful",
                  "careless",
                  "celebration",
                  "comfortable",
                  "decision",
                  "fearless",
                  "flexible",
                  "harmless",
                  "helpful",
                  "invention",
                  "joyful",
                  "reliable"
                ],
                "correct": 8
              },
              {
                "q": "Paying close attention; full of care",
                "options": [
                  "careful",
                  "careless",
                  "celebration",
                  "comfortable",
                  "decision",
                  "fearless",
                  "flexible",
                  "harmless",
                  "helpful",
                  "invention",
                  "joyful",
                  "reliable"
                ],
                "correct": 0
              },
              {
                "q": "Full of joy or happiness",
                "options": [
                  "careful",
                  "careless",
                  "celebration",
                  "comfortable",
                  "decision",
                  "fearless",
                  "flexible",
                  "harmless",
                  "helpful",
                  "invention",
                  "joyful",
                  "reliable"
                ],
                "correct": 10
              },
              {
                "q": "Something newly created or designed",
                "options": [
                  "careful",
                  "careless",
                  "celebration",
                  "comfortable",
                  "decision",
                  "fearless",
                  "flexible",
                  "harmless",
                  "helpful",
                  "invention",
                  "joyful",
                  "reliable"
                ],
                "correct": 9
              },
              {
                "q": "A choice made after thinking it over",
                "options": [
                  "careful",
                  "careless",
                  "celebration",
                  "comfortable",
                  "decision",
                  "fearless",
                  "flexible",
                  "harmless",
                  "helpful",
                  "invention",
                  "joyful",
                  "reliable"
                ],
                "correct": 4
              },
              {
                "q": "An event or activity to honor a special occasion",
                "options": [
                  "careful",
                  "careless",
                  "celebration",
                  "comfortable",
                  "decision",
                  "fearless",
                  "flexible",
                  "harmless",
                  "helpful",
                  "invention",
                  "joyful",
                  "reliable"
                ],
                "correct": 2
              },
              {
                "q": "Able to be trusted or depended on",
                "options": [
                  "careful",
                  "careless",
                  "celebration",
                  "comfortable",
                  "decision",
                  "fearless",
                  "flexible",
                  "harmless",
                  "helpful",
                  "invention",
                  "joyful",
                  "reliable"
                ],
                "correct": 11
              },
              {
                "q": "Able to bend or change easily",
                "options": [
                  "careful",
                  "careless",
                  "celebration",
                  "comfortable",
                  "decision",
                  "fearless",
                  "flexible",
                  "harmless",
                  "helpful",
                  "invention",
                  "joyful",
                  "reliable"
                ],
                "correct": 6
              },
              {
                "q": "Providing comfort; at ease",
                "options": [
                  "careful",
                  "careless",
                  "celebration",
                  "comfortable",
                  "decision",
                  "fearless",
                  "flexible",
                  "harmless",
                  "helpful",
                  "invention",
                  "joyful",
                  "reliable"
                ],
                "correct": 3
              },
              {
                "q": "To join two or more things together",
                "options": [
                  "calculate",
                  "combine",
                  "connect",
                  "depend",
                  "divide",
                  "estimate",
                  "exclude",
                  "include",
                  "measure",
                  "relate",
                  "rely",
                  "separate"
                ],
                "correct": 1
              },
              {
                "q": "To leave something out; to not include",
                "options": [
                  "calculate",
                  "combine",
                  "connect",
                  "depend",
                  "divide",
                  "estimate",
                  "exclude",
                  "include",
                  "measure",
                  "relate",
                  "rely",
                  "separate"
                ],
                "correct": 6
              },
              {
                "q": "To have something as part of a group",
                "options": [
                  "calculate",
                  "combine",
                  "connect",
                  "depend",
                  "divide",
                  "estimate",
                  "exclude",
                  "include",
                  "measure",
                  "relate",
                  "rely",
                  "separate"
                ],
                "correct": 7
              },
              {
                "q": "To find the size, length, or amount of something",
                "options": [
                  "calculate",
                  "combine",
                  "connect",
                  "depend",
                  "divide",
                  "estimate",
                  "exclude",
                  "include",
                  "measure",
                  "relate",
                  "rely",
                  "separate"
                ],
                "correct": 8
              },
              {
                "q": "To make a careful guess about an amount",
                "options": [
                  "calculate",
                  "combine",
                  "connect",
                  "depend",
                  "divide",
                  "estimate",
                  "exclude",
                  "include",
                  "measure",
                  "relate",
                  "rely",
                  "separate"
                ],
                "correct": 5
              },
              {
                "q": "To figure out an answer using math",
                "options": [
                  "calculate",
                  "combine",
                  "connect",
                  "depend",
                  "divide",
                  "estimate",
                  "exclude",
                  "include",
                  "measure",
                  "relate",
                  "rely",
                  "separate"
                ],
                "correct": 0
              },
              {
                "q": "To trust and depend on someone or something",
                "options": [
                  "calculate",
                  "combine",
                  "connect",
                  "depend",
                  "divide",
                  "estimate",
                  "exclude",
                  "include",
                  "measure",
                  "relate",
                  "rely",
                  "separate"
                ],
                "correct": 10
              },
              {
                "q": "To rely on someone or something",
                "options": [
                  "calculate",
                  "combine",
                  "connect",
                  "depend",
                  "divide",
                  "estimate",
                  "exclude",
                  "include",
                  "measure",
                  "relate",
                  "rely",
                  "separate"
                ],
                "correct": 3
              },
              {
                "q": "To show a connection between two things",
                "options": [
                  "calculate",
                  "combine",
                  "connect",
                  "depend",
                  "divide",
                  "estimate",
                  "exclude",
                  "include",
                  "measure",
                  "relate",
                  "rely",
                  "separate"
                ],
                "correct": 9
              },
              {
                "q": "To join or link together",
                "options": [
                  "calculate",
                  "combine",
                  "connect",
                  "depend",
                  "divide",
                  "estimate",
                  "exclude",
                  "include",
                  "measure",
                  "relate",
                  "rely",
                  "separate"
                ],
                "correct": 2
              },
              {
                "q": "To split something into parts",
                "options": [
                  "calculate",
                  "combine",
                  "connect",
                  "depend",
                  "divide",
                  "estimate",
                  "exclude",
                  "include",
                  "measure",
                  "relate",
                  "rely",
                  "separate"
                ],
                "correct": 4
              },
              {
                "q": "To divide or keep apart",
                "options": [
                  "calculate",
                  "combine",
                  "connect",
                  "depend",
                  "divide",
                  "estimate",
                  "exclude",
                  "include",
                  "measure",
                  "relate",
                  "rely",
                  "separate"
                ],
                "correct": 11
              }
            ]
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Nonfiction Detour #3 · RI.4.9",
        "tasks": [
          {
            "id": "ar28",
            "label": "Week 28 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Nonfiction Detour #3 (two articles, same topic — Appendix D)</p><p><b>Standard Focus:</b> RI.4.9</p><p><b>Pacing Goal:</b> Two short articles this week.</p><p>Back to nonfiction for two weeks. Pick a topic Adelyn's curious about and find two different short articles on it (a kids' science site and a simple encyclopedia entry work well). Compare what each one adds.</p></div>"
          },
          {
            "id": "ard28",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Short Vowels &amp; Closed Syllables</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "brand",
                "kind": "word"
              },
              {
                "answer": "drink",
                "kind": "word"
              },
              {
                "answer": "blast",
                "kind": "word"
              },
              {
                "answer": "plastic",
                "kind": "word"
              },
              {
                "answer": "expand",
                "kind": "word"
              },
              {
                "answer": "athletic",
                "kind": "word"
              },
              {
                "answer": "mishap",
                "kind": "word"
              },
              {
                "answer": "sunset",
                "kind": "word"
              },
              {
                "answer": "They watched the sunset after the athletic contest.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar28q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What did the first article tell you that the second one didn't?",
            "sampleAnswer": "Depends on the two sources chosen — a good answer names specific information unique to the first one, e.g., ‘It talked about their diet, which the other article didn't mention.’"
          },
          {
            "id": "ar28q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "If you combined both into one summary, what would you say?",
            "sampleAnswer": "A strong combined summary blends unique facts from both sources into a few sentences, rather than just repeating one source."
          }
        ]
      }
    }
  },
  {
    "week_number": 29,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "-OLOGY — study of (suffix)",
        "tasks": [
          {
            "id": "av29",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>-OLOGY</b> — study of (suffix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>zoology</td><td><i>zoo- (animal) + -logy (study of)</i> — the scientific study of animals<br><span style=\"opacity:.75;\">\"In zoology, Adelyn learned how octopuses change color to hide.\"</span></td></tr><tr><td>technology</td><td><i>techno- (skill/craft) + -logy (study of)</i> — the use of science to create tools and machines<br><span style=\"opacity:.75;\">\"New technology made it easier to video-call family back home.\"</span></td></tr><tr><td>mythology</td><td><i>myth + -logy (study of)</i> — the study of traditional stories, often about gods and heroes<br><span style=\"opacity:.75;\">\"Greek mythology is full of stories about brave heroes and powerful gods.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>trust</td><td>to believe that someone is honest or reliable<br><span style=\"opacity:.75;\">\"It's important to trust your gut when trying a new trick.\"</span></td></tr><tr><td>value</td><td>to think something is important or useful<br><span style=\"opacity:.75;\">\"The family values spending time together while traveling.\"</span></td></tr><tr><td>appreciate</td><td>to be thankful for something; to recognize its worth<br><span style=\"opacity:.75;\">\"Adelyn learned to appreciate the little things, like a warm meal.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Nonfiction Detour #3 · RI.4.7",
        "tasks": [
          {
            "id": "ar29",
            "label": "Week 29 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Nonfiction Detour #3 (an article with a map, chart, or infographic — Appendix D)</p><p><b>Standard Focus:</b> RI.4.7</p><p><b>Pacing Goal:</b> One article this week.</p><p>Choose an article that includes a map, chart, or infographic. Notice how the picture adds information beyond the words — and if the article makes a claim, see if you can spot the evidence behind it (a stretch skill — no pressure to nail it perfectly).</p></div>"
          },
          {
            "id": "ard29",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Silent-E (Magic E)</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "flame",
                "kind": "word"
              },
              {
                "answer": "ride",
                "kind": "word"
              },
              {
                "answer": "note",
                "kind": "word"
              },
              {
                "answer": "invite",
                "kind": "word"
              },
              {
                "answer": "explode",
                "kind": "word"
              },
              {
                "answer": "promote",
                "kind": "word"
              },
              {
                "answer": "sunshine",
                "kind": "word"
              },
              {
                "answer": "homemade",
                "kind": "word"
              },
              {
                "answer": "The homemade note sat in the sunshine.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar29q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What does the map or chart show that the words alone don't?",
            "sampleAnswer": "A map might show exactly where something is located relative to other places, which is hard to picture from a text description alone."
          },
          {
            "id": "ar29q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "Does the author give any reasons or evidence for a claim they make? What are they?",
            "sampleAnswer": "Look for specific facts or examples backing up a stated opinion — e.g., ‘The article says pollution is a problem, and backs it up with a statistic about water quality.’"
          }
        ]
      }
    }
  },
  {
    "week_number": 30,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "HYDRO — water (Greek root)",
        "tasks": [
          {
            "id": "av30",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>HYDRO</b> — water (Greek root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>dehydrate</td><td><i>de- (remove) + hydro (water) + -ate (to make)</i> — to remove water from something; to lose too much water<br><span style=\"opacity:.75;\">\"It's easy to dehydrate when hiking in hot weather, so drink plenty of water.\"</span></td></tr><tr><td>hydroelectric</td><td><i>hydro- (water) + electric</i> — producing electricity using the power of moving water<br><span style=\"opacity:.75;\">\"The dam produces hydroelectric power for the whole region.\"</span></td></tr><tr><td>hydrogen</td><td><i>hydro- (water) + -gen (produce)</i> — a gas that combines with oxygen to form water<br><span style=\"opacity:.75;\">\"Hydrogen is the lightest and most common element in the universe.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>admire</td><td>to respect or think highly of someone<br><span style=\"opacity:.75;\">\"Adelyn admired the gymnast's years of hard work.\"</span></td></tr><tr><td>encourage</td><td>to give someone support or confidence<br><span style=\"opacity:.75;\">\"Kenley encouraged her sister before the big competition.\"</span></td></tr><tr><td>inspire</td><td>to give someone the desire or idea to do something<br><span style=\"opacity:.75;\">\"The mountain view inspired Adelyn to try painting.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Escape from Mr. Lemoncello's Library · RL.4.10",
        "tasks": [
          {
            "id": "ar30",
            "label": "Week 30 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Escape from Mr. Lemoncello's Library</p><p><b>Standard Focus:</b> RL.4.10</p><p><b>Pacing Goal:</b> Read chapters 1-14. This book is more than double the length of anything else this year — the point of this month is exactly that stretch.</p><p>This month is a stamina month — a fun, engaging, noticeably longer book with no new standard to learn. The goal is just getting comfortable with a story read over many short sittings (the chapters are quick, just a lot of them).</p></div>"
          },
          {
            "id": "ard30",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Vowel Teams</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "braid",
                "kind": "word"
              },
              {
                "answer": "clay",
                "kind": "word"
              },
              {
                "answer": "cream",
                "kind": "word"
              },
              {
                "answer": "boast",
                "kind": "word"
              },
              {
                "answer": "throw",
                "kind": "word"
              },
              {
                "answer": "believe",
                "kind": "word"
              },
              {
                "answer": "trail",
                "kind": "word"
              },
              {
                "answer": "sweet",
                "kind": "word"
              },
              {
                "answer": "She can't believe how sweet the trail mix is.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar30q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What's hooking you into this story so far?",
            "sampleAnswer": "Likely the puzzle-solving, competitive game element — Grabenstein's books tend to hook readers with mystery and challenge rather than danger."
          },
          {
            "id": "ar30q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "Predict: what do you think will happen next?",
            "sampleAnswer": "Any reasonable prediction based on clues so far counts — the goal is practicing prediction, not being ‘right.’"
          }
        ]
      }
    }
  },
  {
    "week_number": 31,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "MAL- — bad (prefix)",
        "tasks": [
          {
            "id": "av31",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>MAL-</b> — bad (prefix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>malfunction</td><td><i>mal- (bad) + function</i> — to fail to work correctly<br><span style=\"opacity:.75;\">\"The old camera began to malfunction halfway through the trip.\"</span></td></tr><tr><td>malnutrition</td><td><i>mal- (bad) + nutrition (nourishment)</i> — poor health caused by not getting enough of the right food<br><span style=\"opacity:.75;\">\"Doctors work to prevent malnutrition in communities with little access to fresh food.\"</span></td></tr><tr><td>malady</td><td><i>mal- (bad)</i> — an illness or unhealthy condition<br><span style=\"opacity:.75;\">\"A mild malady kept Kenley resting for a day before she felt better.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>motivate</td><td>to give someone a reason to do something<br><span style=\"opacity:.75;\">\"A gold star chart helped motivate Adelyn during spelling practice.\"</span></td></tr><tr><td>persuade</td><td>to convince someone to do or believe something<br><span style=\"opacity:.75;\">\"Adelyn tried to persuade her mom to let her stay up late.\"</span></td></tr><tr><td>convince</td><td>to make someone believe something is true<br><span style=\"opacity:.75;\">\"It took some convincing before Kenley agreed to try the spicy dish.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Escape from Mr. Lemoncello's Library · RL.4.3/4.6 (review)",
        "tasks": [
          {
            "id": "ar31",
            "label": "Week 31 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Escape from Mr. Lemoncello's Library</p><p><b>Standard Focus:</b> RL.4.3/4.6 (review)</p><p><b>Pacing Goal:</b> Read chapters 15-28.</p><p>Revisit characterization and narration style with this new book, just as light review.</p></div>"
          },
          {
            "id": "ard31",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>R-Controlled Vowels</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "yard",
                "kind": "word"
              },
              {
                "answer": "doctor",
                "kind": "word"
              },
              {
                "answer": "first",
                "kind": "word"
              },
              {
                "answer": "thorn",
                "kind": "word"
              },
              {
                "answer": "curve",
                "kind": "word"
              },
              {
                "answer": "summer",
                "kind": "word"
              },
              {
                "answer": "order",
                "kind": "word"
              },
              {
                "answer": "birthday",
                "kind": "word"
              },
              {
                "answer": "Her birthday party is in the summer yard.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar31q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What have you learned about the main character so far, and how do you know?",
            "sampleAnswer": "Likely that the protagonist is clever and resourceful, especially good at games and puzzles — shown through how he approaches challenges rather than being told directly."
          },
          {
            "id": "ar31q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "Is this story told in first or third person?",
            "sampleAnswer": "Escape from Mr. Lemoncello's Library is told in third person."
          }
        ]
      }
    }
  },
  {
    "week_number": 32,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "NON- — not (prefix)",
        "tasks": [
          {
            "id": "av32",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>NON-</b> — not (prefix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>nonfiction</td><td><i>non- (not) + fiction</i> — writing based on real facts and events<br><span style=\"opacity:.75;\">\"Adelyn prefers nonfiction books about animals and geography.\"</span></td></tr><tr><td>nonstop</td><td><i>non- (not) + stop</i> — without stopping<br><span style=\"opacity:.75;\">\"They took a nonstop flight straight to their next destination.\"</span></td></tr><tr><td>nonsense</td><td><i>non- (not) + sense</i> — words or ideas that have no meaning or logic<br><span style=\"opacity:.75;\">\"The made-up song was silly nonsense, and everyone laughed.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>argue</td><td>to give reasons for or against something, sometimes with disagreement<br><span style=\"opacity:.75;\">\"The sisters argued about whose turn it was to choose the movie.\"</span></td></tr><tr><td>debate</td><td>a discussion between people with different opinions<br><span style=\"opacity:.75;\">\"The class held a debate about the best way to protect the ocean.\"</span></td></tr><tr><td>discuss</td><td>to talk about something with someone<br><span style=\"opacity:.75;\">\"They discussed their plans for the weekend over breakfast.\"</span></td></tr></table>"
          },
          {
            "id": "avq8",
            "label": "Review Quiz 8 — Weeks 29–32",
            "type": "graded-mc",
            "monthlyTest": true,
            "questions": [
              {
                "q": "A gas that combines with oxygen to form water",
                "options": [
                  "dehydrate",
                  "hydroelectric",
                  "hydrogen",
                  "malady",
                  "malfunction",
                  "malnutrition",
                  "mythology",
                  "nonfiction",
                  "nonsense",
                  "nonstop",
                  "technology",
                  "zoology"
                ],
                "correct": 2
              },
              {
                "q": "Producing electricity using the power of moving water",
                "options": [
                  "dehydrate",
                  "hydroelectric",
                  "hydrogen",
                  "malady",
                  "malfunction",
                  "malnutrition",
                  "mythology",
                  "nonfiction",
                  "nonsense",
                  "nonstop",
                  "technology",
                  "zoology"
                ],
                "correct": 1
              },
              {
                "q": "To remove water from something; to lose too much water",
                "options": [
                  "dehydrate",
                  "hydroelectric",
                  "hydrogen",
                  "malady",
                  "malfunction",
                  "malnutrition",
                  "mythology",
                  "nonfiction",
                  "nonsense",
                  "nonstop",
                  "technology",
                  "zoology"
                ],
                "correct": 0
              },
              {
                "q": "The study of traditional stories, often about gods and heroes",
                "options": [
                  "dehydrate",
                  "hydroelectric",
                  "hydrogen",
                  "malady",
                  "malfunction",
                  "malnutrition",
                  "mythology",
                  "nonfiction",
                  "nonsense",
                  "nonstop",
                  "technology",
                  "zoology"
                ],
                "correct": 6
              },
              {
                "q": "The use of science to create tools and machines",
                "options": [
                  "dehydrate",
                  "hydroelectric",
                  "hydrogen",
                  "malady",
                  "malfunction",
                  "malnutrition",
                  "mythology",
                  "nonfiction",
                  "nonsense",
                  "nonstop",
                  "technology",
                  "zoology"
                ],
                "correct": 10
              },
              {
                "q": "The scientific study of animals",
                "options": [
                  "dehydrate",
                  "hydroelectric",
                  "hydrogen",
                  "malady",
                  "malfunction",
                  "malnutrition",
                  "mythology",
                  "nonfiction",
                  "nonsense",
                  "nonstop",
                  "technology",
                  "zoology"
                ],
                "correct": 11
              },
              {
                "q": "Words or ideas that have no meaning or logic",
                "options": [
                  "dehydrate",
                  "hydroelectric",
                  "hydrogen",
                  "malady",
                  "malfunction",
                  "malnutrition",
                  "mythology",
                  "nonfiction",
                  "nonsense",
                  "nonstop",
                  "technology",
                  "zoology"
                ],
                "correct": 8
              },
              {
                "q": "Without stopping",
                "options": [
                  "dehydrate",
                  "hydroelectric",
                  "hydrogen",
                  "malady",
                  "malfunction",
                  "malnutrition",
                  "mythology",
                  "nonfiction",
                  "nonsense",
                  "nonstop",
                  "technology",
                  "zoology"
                ],
                "correct": 9
              },
              {
                "q": "Writing based on real facts and events",
                "options": [
                  "dehydrate",
                  "hydroelectric",
                  "hydrogen",
                  "malady",
                  "malfunction",
                  "malnutrition",
                  "mythology",
                  "nonfiction",
                  "nonsense",
                  "nonstop",
                  "technology",
                  "zoology"
                ],
                "correct": 7
              },
              {
                "q": "An illness or unhealthy condition",
                "options": [
                  "dehydrate",
                  "hydroelectric",
                  "hydrogen",
                  "malady",
                  "malfunction",
                  "malnutrition",
                  "mythology",
                  "nonfiction",
                  "nonsense",
                  "nonstop",
                  "technology",
                  "zoology"
                ],
                "correct": 3
              },
              {
                "q": "Poor health caused by not getting enough of the right food",
                "options": [
                  "dehydrate",
                  "hydroelectric",
                  "hydrogen",
                  "malady",
                  "malfunction",
                  "malnutrition",
                  "mythology",
                  "nonfiction",
                  "nonsense",
                  "nonstop",
                  "technology",
                  "zoology"
                ],
                "correct": 5
              },
              {
                "q": "To fail to work correctly",
                "options": [
                  "dehydrate",
                  "hydroelectric",
                  "hydrogen",
                  "malady",
                  "malfunction",
                  "malnutrition",
                  "mythology",
                  "nonfiction",
                  "nonsense",
                  "nonstop",
                  "technology",
                  "zoology"
                ],
                "correct": 4
              },
              {
                "q": "To give someone the desire or idea to do something",
                "options": [
                  "admire",
                  "appreciate",
                  "argue",
                  "convince",
                  "debate",
                  "discuss",
                  "encourage",
                  "inspire",
                  "motivate",
                  "persuade",
                  "trust",
                  "value"
                ],
                "correct": 7
              },
              {
                "q": "To give someone support or confidence",
                "options": [
                  "admire",
                  "appreciate",
                  "argue",
                  "convince",
                  "debate",
                  "discuss",
                  "encourage",
                  "inspire",
                  "motivate",
                  "persuade",
                  "trust",
                  "value"
                ],
                "correct": 6
              },
              {
                "q": "To respect or think highly of someone",
                "options": [
                  "admire",
                  "appreciate",
                  "argue",
                  "convince",
                  "debate",
                  "discuss",
                  "encourage",
                  "inspire",
                  "motivate",
                  "persuade",
                  "trust",
                  "value"
                ],
                "correct": 0
              },
              {
                "q": "To be thankful for something; to recognize its worth",
                "options": [
                  "admire",
                  "appreciate",
                  "argue",
                  "convince",
                  "debate",
                  "discuss",
                  "encourage",
                  "inspire",
                  "motivate",
                  "persuade",
                  "trust",
                  "value"
                ],
                "correct": 1
              },
              {
                "q": "To think something is important or useful",
                "options": [
                  "admire",
                  "appreciate",
                  "argue",
                  "convince",
                  "debate",
                  "discuss",
                  "encourage",
                  "inspire",
                  "motivate",
                  "persuade",
                  "trust",
                  "value"
                ],
                "correct": 11
              },
              {
                "q": "To believe that someone is honest or reliable",
                "options": [
                  "admire",
                  "appreciate",
                  "argue",
                  "convince",
                  "debate",
                  "discuss",
                  "encourage",
                  "inspire",
                  "motivate",
                  "persuade",
                  "trust",
                  "value"
                ],
                "correct": 10
              },
              {
                "q": "To talk about something with someone",
                "options": [
                  "admire",
                  "appreciate",
                  "argue",
                  "convince",
                  "debate",
                  "discuss",
                  "encourage",
                  "inspire",
                  "motivate",
                  "persuade",
                  "trust",
                  "value"
                ],
                "correct": 5
              },
              {
                "q": "A discussion between people with different opinions",
                "options": [
                  "admire",
                  "appreciate",
                  "argue",
                  "convince",
                  "debate",
                  "discuss",
                  "encourage",
                  "inspire",
                  "motivate",
                  "persuade",
                  "trust",
                  "value"
                ],
                "correct": 4
              },
              {
                "q": "To give reasons for or against something, sometimes with disagreement",
                "options": [
                  "admire",
                  "appreciate",
                  "argue",
                  "convince",
                  "debate",
                  "discuss",
                  "encourage",
                  "inspire",
                  "motivate",
                  "persuade",
                  "trust",
                  "value"
                ],
                "correct": 2
              },
              {
                "q": "To make someone believe something is true",
                "options": [
                  "admire",
                  "appreciate",
                  "argue",
                  "convince",
                  "debate",
                  "discuss",
                  "encourage",
                  "inspire",
                  "motivate",
                  "persuade",
                  "trust",
                  "value"
                ],
                "correct": 3
              },
              {
                "q": "To convince someone to do or believe something",
                "options": [
                  "admire",
                  "appreciate",
                  "argue",
                  "convince",
                  "debate",
                  "discuss",
                  "encourage",
                  "inspire",
                  "motivate",
                  "persuade",
                  "trust",
                  "value"
                ],
                "correct": 9
              },
              {
                "q": "To give someone a reason to do something",
                "options": [
                  "admire",
                  "appreciate",
                  "argue",
                  "convince",
                  "debate",
                  "discuss",
                  "encourage",
                  "inspire",
                  "motivate",
                  "persuade",
                  "trust",
                  "value"
                ],
                "correct": 8
              }
            ]
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Escape from Mr. Lemoncello's Library · RL.4.5 (review)",
        "tasks": [
          {
            "id": "ar32",
            "label": "Week 32 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Escape from Mr. Lemoncello's Library</p><p><b>Standard Focus:</b> RL.4.5 (review)</p><p><b>Pacing Goal:</b> Read chapters 29-42.</p><p>Check in on the story's shape — beginning, middle, and where the problem is heading.</p></div>"
          },
          {
            "id": "ard32",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Blends &amp; Digraphs</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "strap",
                "kind": "word"
              },
              {
                "answer": "thicken",
                "kind": "word"
              },
              {
                "answer": "chuckle",
                "kind": "word"
              },
              {
                "answer": "shiver",
                "kind": "word"
              },
              {
                "answer": "sprinkle",
                "kind": "word"
              },
              {
                "answer": "whittle",
                "kind": "word"
              },
              {
                "answer": "scratch",
                "kind": "word"
              },
              {
                "answer": "thumbprint",
                "kind": "word"
              },
              {
                "answer": "She had to chuckle at the scratchy sweater.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar32q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What's the main problem in this book, and how is it building toward a solution?",
            "sampleAnswer": "The central challenge is escaping/winning the library game by solving puzzles — tension builds as the stakes and puzzle difficulty increase."
          },
          {
            "id": "ar32q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "What's your favorite twist or surprise so far?",
            "sampleAnswer": "Any specific plot twist she names counts as a strong answer."
          }
        ]
      }
    }
  },
  {
    "week_number": 33,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "-IST — one who (suffix)",
        "tasks": [
          {
            "id": "av33",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>-IST</b> — one who (suffix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>artist</td><td><i>art + -ist (one who)</i> — a person who creates art<br><span style=\"opacity:.75;\">\"The street artist painted a colorful mural on the old wall.\"</span></td></tr><tr><td>scientist</td><td><i>science + -ist (one who)</i> — a person who studies and researches science<br><span style=\"opacity:.75;\">\"A marine scientist explained how coral reefs grow.\"</span></td></tr><tr><td>tourist</td><td><i>tour + -ist (one who)</i> — a person traveling for pleasure<br><span style=\"opacity:.75;\">\"The small town welcomed tourists from all over the world.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>explain</td><td>to make something clear or easy to understand<br><span style=\"opacity:.75;\">\"Kenley explained the recipe step by step.\"</span></td></tr><tr><td>clarify</td><td>to make something clearer or easier to understand<br><span style=\"opacity:.75;\">\"The teacher clarified the instructions before the quiz began.\"</span></td></tr><tr><td>confuse</td><td>to make someone unsure or unable to understand<br><span style=\"opacity:.75;\">\"The maze of narrow streets began to confuse them.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Escape from Mr. Lemoncello's Library · RL.4.1/4.10",
        "tasks": [
          {
            "id": "ar33",
            "label": "Week 33 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Escape from Mr. Lemoncello's Library</p><p><b>Standard Focus:</b> RL.4.1/4.10</p><p><b>Pacing Goal:</b> Finish the book — chapters 43-56.</p><p>Finish the book and reflect on the whole thing.</p></div>"
          },
          {
            "id": "ard33",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Common Prefixes</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "recheck",
                "kind": "word"
              },
              {
                "answer": "unroll",
                "kind": "word"
              },
              {
                "answer": "discolor",
                "kind": "word"
              },
              {
                "answer": "preheat",
                "kind": "word"
              },
              {
                "answer": "mistype",
                "kind": "word"
              },
              {
                "answer": "nonstick",
                "kind": "word"
              },
              {
                "answer": "renew",
                "kind": "word"
              },
              {
                "answer": "unwind",
                "kind": "word"
              },
              {
                "answer": "Please preheat the pan and unroll the dough.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar33q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "How did the story wrap up? Did it end how you expected?",
            "sampleAnswer": "A strong answer explains the actual ending and honestly compares it to what she predicted earlier in the month."
          },
          {
            "id": "ar33q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "Would you recommend this book to a friend? Why or why not?",
            "sampleAnswer": "Any genuine opinion with a reason counts — e.g., ‘Yes, because the puzzles were fun to guess along with.’"
          }
        ]
      }
    }
  },
  {
    "week_number": 34,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "-ER / -OR — one who (suffix)",
        "tasks": [
          {
            "id": "av34",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>-ER / -OR</b> — one who (suffix). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>inventor</td><td><i>invent + -or (one who)</i> — a person who creates something new<br><span style=\"opacity:.75;\">\"The young inventor built a machine that sorted seashells by size.\"</span></td></tr><tr><td>visitor</td><td><i>visit + -or (one who)</i> — a person who visits a place<br><span style=\"opacity:.75;\">\"The museum welcomed visitors from dozens of countries.\"</span></td></tr><tr><td>actor</td><td><i>act + -or (one who)</i> — a person who performs in plays or movies<br><span style=\"opacity:.75;\">\"The actor practiced her lines before the big performance.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>puzzle</td><td>to confuse someone; a problem that is hard to solve<br><span style=\"opacity:.75;\">\"The old riddle puzzled Adelyn for the whole afternoon.\"</span></td></tr><tr><td>curious</td><td>eager to learn or know something<br><span style=\"opacity:.75;\">\"Adelyn is curious about how other kids learn in different countries.\"</span></td></tr><tr><td>eager</td><td>very excited and ready to do something<br><span style=\"opacity:.75;\">\"She was eager to try the new gymnastics gym in town.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Dragonbreath · RL.4.3/4.5 (Characterization Checkpoint #5, cumulative)",
        "tasks": [
          {
            "id": "ar34",
            "label": "Week 34 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Dragonbreath (Book 1)</p><p><b>Standard Focus:</b> RL.4.3/4.5 (Characterization Checkpoint #5, cumulative)</p><p><b>Pacing Goal:</b> Read to roughly the 35% mark.</p><p>Our final book of the year is another comic-hybrid adventure — Book 1 specifically, the mildest entry in the series. Bring together everything: characterization, setting, and story structure, all in one conversation.</p></div>"
          },
          {
            "id": "ard34",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Common Suffixes</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "motion",
                "kind": "word"
              },
              {
                "answer": "expression",
                "kind": "word"
              },
              {
                "answer": "balancing",
                "kind": "word"
              },
              {
                "answer": "landed",
                "kind": "word"
              },
              {
                "answer": "silently",
                "kind": "word"
              },
              {
                "answer": "wonderful",
                "kind": "word"
              },
              {
                "answer": "hopeless",
                "kind": "word"
              },
              {
                "answer": "invitation",
                "kind": "word"
              },
              {
                "answer": "She landed silently after balancing on the beam.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar34q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "What do we learn about the main character through what he does, not just what the narrator says?",
            "sampleAnswer": "Danny Dragonbreath tends to show his impulsiveness and loyalty to his friend Wendell through the choices he makes, rather than the narrator directly describing his personality."
          },
          {
            "id": "ar34q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "Where are we in the story's shape — beginning, middle, or end?",
            "sampleAnswer": "If we're still meeting characters and learning the setup, that's the beginning; if complications are actively growing, that's the middle."
          }
        ]
      }
    }
  },
  {
    "week_number": 35,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "CRED — believe (Latin root)",
        "tasks": [
          {
            "id": "av35",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>CRED</b> — believe (Latin root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>credible</td><td><i>cred (believe) + -ible (capable of)</i> — believable; able to be trusted<br><span style=\"opacity:.75;\">\"The witness gave a credible account of what happened.\"</span></td></tr><tr><td>credit</td><td><i>cred (believe/trust)</i> — trust or belief in someone's honesty; recognition for an achievement<br><span style=\"opacity:.75;\">\"Adelyn deserves credit for practicing every single day.\"</span></td></tr><tr><td>incredible</td><td><i>in- (not) + cred (believe) + -ible (capable of)</i> — impossible or hard to believe; amazing<br><span style=\"opacity:.75;\">\"The view from the mountaintop was absolutely incredible.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>anxious</td><td>worried or nervous about something<br><span style=\"opacity:.75;\">\"Adelyn felt a little anxious before her first competition abroad.\"</span></td></tr><tr><td>nervous</td><td>feeling worried or uneasy<br><span style=\"opacity:.75;\">\"Kenley was nervous before tasting the unusual fruit.\"</span></td></tr><tr><td>confident</td><td>feeling sure of yourself and your abilities<br><span style=\"opacity:.75;\">\"After months of practice, Adelyn felt confident about her routine.\"</span></td></tr></table>"
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Dragonbreath · RL.4.9 (cumulative)",
        "tasks": [
          {
            "id": "ar35",
            "label": "Week 35 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Dragonbreath (Book 1)</p><p><b>Standard Focus:</b> RL.4.9 (cumulative)</p><p><b>Pacing Goal:</b> Read to roughly the 70% mark.</p><p>Look back at the whole year of reading. Compare this book to an earlier favorite.</p></div>"
          },
          {
            "id": "ard35",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Syllable Division Patterns</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "contest",
                "kind": "word"
              },
              {
                "answer": "pirate",
                "kind": "word"
              },
              {
                "answer": "cover",
                "kind": "word"
              },
              {
                "answer": "little",
                "kind": "word"
              },
              {
                "answer": "insect",
                "kind": "word"
              },
              {
                "answer": "siren",
                "kind": "word"
              },
              {
                "answer": "salad",
                "kind": "word"
              },
              {
                "answer": "uncle",
                "kind": "word"
              },
              {
                "answer": "Her uncle made a little salad for the contest.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar35q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "How is this book similar to or different from an earlier favorite in humor, adventure, or characters?",
            "sampleAnswer": "‘Dragonbreath uses the same comic-panel style as The Tea Dragon Society, but it's much more focused on jokes and slapstick adventure instead of quiet, cozy moments.’"
          },
          {
            "id": "ar35q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "Which book from this year was your favorite, and why?",
            "sampleAnswer": "Any genuine, specific answer with reasoning counts — the goal is she can articulate why, not just name a title."
          }
        ]
      }
    }
  },
  {
    "week_number": 36,
    "subjects": {
      "vocab": {
        "name": "Vocabulary",
        "tag": "PHON — sound (Greek root)",
        "tasks": [
          {
            "id": "av36",
            "label": "Study This Week's Words",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p>Root family this week: <b>PHON</b> — sound (Greek root). Three root words share this piece, plus three everyday \"frequency\" words that show up constantly in her reading.</p></div><table class=\"vocab-table\"><tr><td colspan=\"2\"><b>Strand 1: Root Words</b></td></tr><tr><td>microphone</td><td><i>micro- (small) + phon (sound)</i> — a device that picks up and makes sound louder<br><span style=\"opacity:.75;\">\"Adelyn spoke into the microphone before her speech.\"</span></td></tr><tr><td>symphony</td><td><i>sym- (together) + phon (sound)</i> — a long piece of music played by a full orchestra<br><span style=\"opacity:.75;\">\"The orchestra performed a beautiful symphony in the old concert hall.\"</span></td></tr><tr><td>phonics</td><td><i>phon (sound) + -ics (study/system)</i> — a method of learning to read by connecting sounds to letters<br><span style=\"opacity:.75;\">\"Phonics helps young readers sound out new words.\"</span></td></tr><tr><td colspan=\"2\"><b>Strand 2: Frequency Words</b></td></tr><tr><td>cautious</td><td>being careful to avoid danger or mistakes<br><span style=\"opacity:.75;\">\"They were cautious while hiking along the rocky cliff.\"</span></td></tr><tr><td>courageous</td><td>brave; willing to face fear or difficulty<br><span style=\"opacity:.75;\">\"It was courageous of Adelyn to try the trick in front of a crowd.\"</span></td></tr><tr><td>grateful</td><td>feeling or showing thanks<br><span style=\"opacity:.75;\">\"They felt grateful for the kind family who hosted them.\"</span></td></tr></table>"
          },
          {
            "id": "avq9",
            "label": "Review Quiz 9 — Weeks 33–36",
            "type": "graded-mc",
            "monthlyTest": true,
            "questions": [
              {
                "q": "A person who performs in plays or movies",
                "options": [
                  "actor",
                  "artist",
                  "credible",
                  "credit",
                  "incredible",
                  "inventor",
                  "microphone",
                  "phonics",
                  "scientist",
                  "symphony",
                  "tourist",
                  "visitor"
                ],
                "correct": 0
              },
              {
                "q": "A person who visits a place",
                "options": [
                  "actor",
                  "artist",
                  "credible",
                  "credit",
                  "incredible",
                  "inventor",
                  "microphone",
                  "phonics",
                  "scientist",
                  "symphony",
                  "tourist",
                  "visitor"
                ],
                "correct": 11
              },
              {
                "q": "A person who creates something new",
                "options": [
                  "actor",
                  "artist",
                  "credible",
                  "credit",
                  "incredible",
                  "inventor",
                  "microphone",
                  "phonics",
                  "scientist",
                  "symphony",
                  "tourist",
                  "visitor"
                ],
                "correct": 5
              },
              {
                "q": "A person traveling for pleasure",
                "options": [
                  "actor",
                  "artist",
                  "credible",
                  "credit",
                  "incredible",
                  "inventor",
                  "microphone",
                  "phonics",
                  "scientist",
                  "symphony",
                  "tourist",
                  "visitor"
                ],
                "correct": 10
              },
              {
                "q": "A person who studies and researches science",
                "options": [
                  "actor",
                  "artist",
                  "credible",
                  "credit",
                  "incredible",
                  "inventor",
                  "microphone",
                  "phonics",
                  "scientist",
                  "symphony",
                  "tourist",
                  "visitor"
                ],
                "correct": 8
              },
              {
                "q": "A person who creates art",
                "options": [
                  "actor",
                  "artist",
                  "credible",
                  "credit",
                  "incredible",
                  "inventor",
                  "microphone",
                  "phonics",
                  "scientist",
                  "symphony",
                  "tourist",
                  "visitor"
                ],
                "correct": 1
              },
              {
                "q": "A method of learning to read by connecting sounds to letters",
                "options": [
                  "actor",
                  "artist",
                  "credible",
                  "credit",
                  "incredible",
                  "inventor",
                  "microphone",
                  "phonics",
                  "scientist",
                  "symphony",
                  "tourist",
                  "visitor"
                ],
                "correct": 7
              },
              {
                "q": "A long piece of music played by a full orchestra",
                "options": [
                  "actor",
                  "artist",
                  "credible",
                  "credit",
                  "incredible",
                  "inventor",
                  "microphone",
                  "phonics",
                  "scientist",
                  "symphony",
                  "tourist",
                  "visitor"
                ],
                "correct": 9
              },
              {
                "q": "A device that picks up and makes sound louder",
                "options": [
                  "actor",
                  "artist",
                  "credible",
                  "credit",
                  "incredible",
                  "inventor",
                  "microphone",
                  "phonics",
                  "scientist",
                  "symphony",
                  "tourist",
                  "visitor"
                ],
                "correct": 6
              },
              {
                "q": "Impossible or hard to believe; amazing",
                "options": [
                  "actor",
                  "artist",
                  "credible",
                  "credit",
                  "incredible",
                  "inventor",
                  "microphone",
                  "phonics",
                  "scientist",
                  "symphony",
                  "tourist",
                  "visitor"
                ],
                "correct": 4
              },
              {
                "q": "Trust or belief in someone's honesty; recognition for an achievement",
                "options": [
                  "actor",
                  "artist",
                  "credible",
                  "credit",
                  "incredible",
                  "inventor",
                  "microphone",
                  "phonics",
                  "scientist",
                  "symphony",
                  "tourist",
                  "visitor"
                ],
                "correct": 3
              },
              {
                "q": "Believable; able to be trusted",
                "options": [
                  "actor",
                  "artist",
                  "credible",
                  "credit",
                  "incredible",
                  "inventor",
                  "microphone",
                  "phonics",
                  "scientist",
                  "symphony",
                  "tourist",
                  "visitor"
                ],
                "correct": 2
              },
              {
                "q": "Very excited and ready to do something",
                "options": [
                  "anxious",
                  "cautious",
                  "clarify",
                  "confident",
                  "confuse",
                  "courageous",
                  "curious",
                  "eager",
                  "explain",
                  "grateful",
                  "nervous",
                  "puzzle"
                ],
                "correct": 7
              },
              {
                "q": "Eager to learn or know something",
                "options": [
                  "anxious",
                  "cautious",
                  "clarify",
                  "confident",
                  "confuse",
                  "courageous",
                  "curious",
                  "eager",
                  "explain",
                  "grateful",
                  "nervous",
                  "puzzle"
                ],
                "correct": 6
              },
              {
                "q": "To confuse someone; a problem that is hard to solve",
                "options": [
                  "anxious",
                  "cautious",
                  "clarify",
                  "confident",
                  "confuse",
                  "courageous",
                  "curious",
                  "eager",
                  "explain",
                  "grateful",
                  "nervous",
                  "puzzle"
                ],
                "correct": 11
              },
              {
                "q": "To make someone unsure or unable to understand",
                "options": [
                  "anxious",
                  "cautious",
                  "clarify",
                  "confident",
                  "confuse",
                  "courageous",
                  "curious",
                  "eager",
                  "explain",
                  "grateful",
                  "nervous",
                  "puzzle"
                ],
                "correct": 4
              },
              {
                "q": "To make something clearer or easier to understand",
                "options": [
                  "anxious",
                  "cautious",
                  "clarify",
                  "confident",
                  "confuse",
                  "courageous",
                  "curious",
                  "eager",
                  "explain",
                  "grateful",
                  "nervous",
                  "puzzle"
                ],
                "correct": 2
              },
              {
                "q": "To make something clear or easy to understand",
                "options": [
                  "anxious",
                  "cautious",
                  "clarify",
                  "confident",
                  "confuse",
                  "courageous",
                  "curious",
                  "eager",
                  "explain",
                  "grateful",
                  "nervous",
                  "puzzle"
                ],
                "correct": 8
              },
              {
                "q": "Feeling or showing thanks",
                "options": [
                  "anxious",
                  "cautious",
                  "clarify",
                  "confident",
                  "confuse",
                  "courageous",
                  "curious",
                  "eager",
                  "explain",
                  "grateful",
                  "nervous",
                  "puzzle"
                ],
                "correct": 9
              },
              {
                "q": "Brave; willing to face fear or difficulty",
                "options": [
                  "anxious",
                  "cautious",
                  "clarify",
                  "confident",
                  "confuse",
                  "courageous",
                  "curious",
                  "eager",
                  "explain",
                  "grateful",
                  "nervous",
                  "puzzle"
                ],
                "correct": 5
              },
              {
                "q": "Being careful to avoid danger or mistakes",
                "options": [
                  "anxious",
                  "cautious",
                  "clarify",
                  "confident",
                  "confuse",
                  "courageous",
                  "curious",
                  "eager",
                  "explain",
                  "grateful",
                  "nervous",
                  "puzzle"
                ],
                "correct": 1
              },
              {
                "q": "Feeling sure of yourself and your abilities",
                "options": [
                  "anxious",
                  "cautious",
                  "clarify",
                  "confident",
                  "confuse",
                  "courageous",
                  "curious",
                  "eager",
                  "explain",
                  "grateful",
                  "nervous",
                  "puzzle"
                ],
                "correct": 3
              },
              {
                "q": "Feeling worried or uneasy",
                "options": [
                  "anxious",
                  "cautious",
                  "clarify",
                  "confident",
                  "confuse",
                  "courageous",
                  "curious",
                  "eager",
                  "explain",
                  "grateful",
                  "nervous",
                  "puzzle"
                ],
                "correct": 10
              },
              {
                "q": "Worried or nervous about something",
                "options": [
                  "anxious",
                  "cautious",
                  "clarify",
                  "confident",
                  "confuse",
                  "courageous",
                  "curious",
                  "eager",
                  "explain",
                  "grateful",
                  "nervous",
                  "puzzle"
                ],
                "correct": 0
              }
            ]
          },
          {
            "id": "avfinal",
            "label": "End-of-Year Vocabulary Test",
            "type": "graded-mc",
            "termFinal": true,
            "questions": [
              {
                "q": "To tell or write details about something",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 4
              },
              {
                "q": "To break into a conversation or activity",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 11
              },
              {
                "q": "To build or put something together",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 3
              },
              {
                "q": "A feeling of admiration for someone; to treat someone with consideration",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 16
              },
              {
                "q": "A book that lists words and their meanings",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 6
              },
              {
                "q": "To say what will happen before it happens",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 13
              },
              {
                "q": "To carry information back and share it with others",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 15
              },
              {
                "q": "To carry people or things from one place to another",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 19
              },
              {
                "q": "A device that shows moving pictures and sound sent from far away",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 18
              },
              {
                "q": "A device used to talk to someone who is far away",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 17
              },
              {
                "q": "To find something for the first time",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 7
              },
              {
                "q": "To make something clear or easy to understand",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 9
              },
              {
                "q": "To look at two or more things to see how they are alike or different",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 2
              },
              {
                "q": "A chance to do something",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 12
              },
              {
                "q": "To succeed in reaching a goal",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 0
              },
              {
                "q": "The surroundings or conditions in which someone lives",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 8
              },
              {
                "q": "The reason something is done; a goal",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 14
              },
              {
                "q": "To decide something after careful thought; to figure out",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 5
              },
              {
                "q": "Even though; despite the fact that",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 1
              },
              {
                "q": "Used to introduce a statement that contrasts with something already said",
                "options": [
                  "achieve",
                  "although",
                  "compare",
                  "construct",
                  "describe",
                  "determine",
                  "dictionary",
                  "discover",
                  "environment",
                  "explain",
                  "however",
                  "interrupt",
                  "opportunity",
                  "predict",
                  "purpose",
                  "report",
                  "respect",
                  "telephone",
                  "television",
                  "transport"
                ],
                "correct": 10
              },
              {
                "q": "Which word parts make up \"telephone\"?",
                "options": [
                  "tele- (far) + phone (sound)",
                  "tele- (far) + vision (to see)"
                ],
                "correct": 0
              },
              {
                "q": "Which word parts make up \"television\"?",
                "options": [
                  "trans- (across) + port (carry)",
                  "tele- (far) + vision (to see)"
                ],
                "correct": 1
              },
              {
                "q": "Which word parts make up \"transport\"?",
                "options": [
                  "trans- (across) + port (carry)",
                  "re- (back) + port (carry)"
                ],
                "correct": 0
              },
              {
                "q": "Which word parts make up \"report\"?",
                "options": [
                  "pre- (before) + dict (say)",
                  "re- (back) + port (carry)"
                ],
                "correct": 1
              },
              {
                "q": "Which word parts make up \"predict\"?",
                "options": [
                  "pre- (before) + dict (say)",
                  "dict (say/word) + -ary (place/collection)"
                ],
                "correct": 0
              },
              {
                "q": "Which word parts make up \"dictionary\"?",
                "options": [
                  "re- (again) + spect (look)",
                  "dict (say/word) + -ary (place/collection)"
                ],
                "correct": 1
              },
              {
                "q": "Which word parts make up \"respect\"?",
                "options": [
                  "re- (again) + spect (look)",
                  "con- (together) + struct (build)"
                ],
                "correct": 0
              },
              {
                "q": "Which word parts make up \"construct\"?",
                "options": [
                  "inter- (between) + rupt (break)",
                  "con- (together) + struct (build)"
                ],
                "correct": 1
              },
              {
                "q": "Which word parts make up \"interrupt\"?",
                "options": [
                  "inter- (between) + rupt (break)",
                  "de- (down) + scrib (write)"
                ],
                "correct": 0
              },
              {
                "q": "Which word parts make up \"describe\"?",
                "options": [
                  "tele- (far) + phone (sound)",
                  "de- (down) + scrib (write)"
                ],
                "correct": 1
              }
            ]
          }
        ]
      },
      "reading": {
        "name": "Reading",
        "tag": "Dragonbreath · RF.4.4 (fluency check-in)",
        "tasks": [
          {
            "id": "ar36",
            "label": "Week 36 Reading & Discussion Lesson",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Book:</b> Dragonbreath (Book 1, wrap-up) — Reading Year Celebration</p><p><b>Standard Focus:</b> RF.4.4 (fluency check-in)</p><p><b>Pacing Goal:</b> Finish the book.</p><p>Celebrate a full year of reading! As a light, no-pressure fluency check-in, ask her to read a favorite short passage aloud from any book this year — purely for the joy of hearing how far she's come, not a test.</p></div>"
          },
          {
            "id": "ard36",
            "label": "Phonics Fluency Warm-Up (Listen & Spell)",
            "type": "graded-dictation",
            "prompt": "This week's review category: <b>Homophones &amp; Confused Words</b>. Tap each 🔊 button to hear the word (or the practice sentence), then type what you hear. This is quick review of sounds she already knows — the goal is speed, not new learning.",
            "words": [
              {
                "answer": "whole",
                "kind": "word"
              },
              {
                "answer": "hole",
                "kind": "word"
              },
              {
                "answer": "break",
                "kind": "word"
              },
              {
                "answer": "brake",
                "kind": "word"
              },
              {
                "answer": "plain",
                "kind": "word"
              },
              {
                "answer": "plane",
                "kind": "word"
              },
              {
                "answer": "The whole plane ride went by without a single break.",
                "kind": "sentence"
              }
            ]
          },
          {
            "id": "ar36q1",
            "label": "Talk It Over — Question 1",
            "type": "reflection",
            "prompt": "Which character from this whole year would you want to be friends with, and why?",
            "sampleAnswer": "‘Maisie, because she's brave and I'd want to solve mysteries with her’ — any specific character with genuine reasoning works."
          },
          {
            "id": "ar36q2",
            "label": "Talk It Over — Question 2",
            "type": "reflection",
            "prompt": "What's one thing about reading that felt easier by the end of the year than at the beginning?",
            "sampleAnswer": "‘Reading longer books without getting tired,’ or ‘figuring out a story's theme faster’ — the goal is a specific, honest reflection."
          }
        ]
      }
    }
  }
];
