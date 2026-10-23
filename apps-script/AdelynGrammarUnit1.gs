/**
 * The ELA Pastry Kitchen — Adelyn's Grammar Town, Unit 1 (Weeks 1-10):
 * The Eight Parts of Speech, refreshed and taken to Town-level depth.
 *
 * Adds net-new Schedule rows for Adelyn's "grammar" subject, converted
 * from Adelyns_Grammar_Town_Full_Year.docx, Unit 1 only. Does NOT touch
 * setupSheets(), Month1Weeks2to5.gs, AdelynVocabReading.gs,
 * AdelynSpelling.gs, or any Kenley rows — purely additive, and only
 * touches Adelyn's placeholder "grammar" row.
 *
 * THIS IS A PARTIAL SEED, DELIBERATELY. Grammar Town runs 36 weeks across
 * 5 units, introducing an evolving four-level analysis system (parts of
 * speech, then parts of the sentence, then phrases, then clauses) rather
 * than applying one fixed set of categories all year the way Kenley's
 * grammar does. Given how much design judgment the word-by-word tagging
 * requires, Units 2-5 (Weeks 11-36) were deliberately left for a
 * follow-up pass after this one is reviewed — see the build notes below
 * before continuing them.
 *
 * Run seedAdelynGrammarUnit1_() ONCE from the function dropdown, after
 * setupSheets() has already been run at least once. It's idempotent: if
 * it detects task_id "agl1" already in the Schedule sheet, it assumes
 * this has already run and does nothing. It also removes Adelyn's
 * original single placeholder row for "grammar" ("ag1" waiting-shell —
 * note the real content below reuses "ag1_1" etc., not "ag1", so there's
 * no id collision to worry about).
 *
 * Content structure per week:
 *   One "read" task (the week's lesson, written directly to Adelyn) plus
 *   one "pos-tagger" task per practice sentence (2-4 per week, matching
 *   Kenley's Level-1 tap-to-tag mechanic exactly — tap a word, choose its
 *   part of speech from the week's options list).
 *
 * Two-step design (updated 2026-10-23, per MCT): Step 1 (pos-tagger) uses
 * ONLY the 8 parts of speech, by the word's job in the sentence. A gerund is
 * tagged Noun, a participle Adjective, and both "to" and its verb in an
 * infinitive are tagged Verb ("to" before a noun stays Preposition). Step 2
 * (concept-check, id = sentence id + "c") then asks which verb-built words
 * are gerunds, participles, infinitives, or regular verbs. Any later week
 * that adds a verbal must follow the same split. Sentences in Weeks 1-5
 * contain no infinitives, since infinitives are not taught until Week 6. The
 * source doc's fuller nuance is kept in each word's explanations entry,
 * shown after she submits.
 *
 * A few sentences also needed hand-built overrides rather than automatic
 * parsing, because the source docx lost some paragraph breaks on export
 * (two answer keys ran together with no separator, and one sentence's
 * answer key silently skipped three words) — flagged and fixed by hand,
 * not guessed at silently. If continuing Units 2-5 from the same source
 * document, re-check for this failure mode.
 */

function seedAdelynGrammarUnit1_() {
  var sh = getSheet_('Schedule');
  var existing = sheetToObjects_(sh);

  // Delete EVERY existing Adelyn row for grammar — old placeholder, an
  // older version of this same content, anything — and rebuild fresh from
  // ADELYN_GRAMMAR_UNIT1_WEEKS below. Safe to re-run any time this content
  // is updated; it always fully replaces rather than skipping.
  var toDelete = existing.filter(function (r) {
    return r.student === 'adelyn' && r.subject_key === 'grammar';
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

  ADELYN_GRAMMAR_UNIT1_WEEKS.forEach(function (week) {
    Object.keys(week.subjects).forEach(function (subjectKey) {
      var subj = week.subjects[subjectKey];
      addTasks(subjectKey, subj.name, subj.tag, week.week_number, subj.tasks);
    });
  });

  var headers = SHEET_HEADERS.Schedule;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  var doneMsg = 'Added ' + rows.length + ' new Schedule rows for Adelyn — Grammar, Unit 1 (Weeks 1-10).';
  try { SpreadsheetApp.getUi().alert(doneMsg); } catch (e) { Logger.log(doneMsg); }
}

// ---------- Content ----------

var ADELYN_GRAMMAR_UNIT1_WEEKS = [
  {
    "week_number": 1,
    "subjects": {
      "grammar": {
        "name": "Grammar",
        "tag": "Grammar Town · Unit 1, Week 1",
        "tasks": [
          {
            "id": "agl1",
            "label": "Week 1 Lesson: Nouns Refresher — Plus Collective & Compound Nouns",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Week 1: Nouns Refresher — Plus Collective &amp; Compound Nouns</b></p><p style=\"opacity:.7;font-size:0.78rem;\">REVIEW / PRACTICE — Standards: L.4.6 (academic vocabulary); foundation for L.4.1</p><p>Hi Adelyn!</p><p>You already know nouns from Grammar Island — a noun names a person, place, thing, or idea, and it's either common (any old gym) or proper (one specific place, like Portugal).</p><p>Quick refresher, then we add something new.</p><p><b>NEW</b>: a <b>COLLECTIVE</b> noun names a whole group acting as one unit — team, family, herd, class.</p><p>A <b>COMPOUND</b> noun is two words smooshed into one noun idea — handstand, backflip, breakfast.</p><p>This year we're using the same four-level analysis system from Island, but we're going to go faster and deeper — like moving from tricycle to bicycle.</p><p>Same balance skills, more speed.</p><p>Ready?</p><p>Let's find some nouns, including a few sneaky new kinds.</p><p style=\"opacity:.8;\">For each sentence, label the part of speech of every word.</p></div>"
          },
          {
            "id": "ag1_1",
            "label": "Sentence 1",
            "type": "pos-tagger",
            "sentence": [
              "Adelyn's",
              "gymnastics",
              "team",
              "practiced",
              "a",
              "new",
              "handstand",
              "in",
              "the",
              "courtyard",
              "of",
              "the",
              "castle."
            ],
            "answers": [
              "Noun",
              "Adjective",
              "Noun",
              "Verb",
              "Adjective",
              "Adjective",
              "Noun",
              "Preposition",
              "Adjective",
              "Noun",
              "Preposition",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "possessive proper noun",
              "noun used as an adjective (describes team)",
              "common/collective noun",
              "action verb",
              "article",
              "adjective",
              "common/compound noun",
              "preposition",
              "article",
              "common noun",
              "preposition",
              "article",
              "common noun"
            ]
          },
          {
            "id": "ag1_1c",
            "label": "Sentence 1 — Noun Check",
            "type": "concept-check",
            "sentence": [
              "Adelyn's",
              "gymnastics",
              "team",
              "practiced",
              "a",
              "new",
              "handstand",
              "in",
              "the",
              "courtyard",
              "of",
              "the",
              "castle."
            ],
            "prompt": "Now look at the nouns you found in this sentence. Is each one common, proper, collective, or compound?",
            "options": [
              "Common",
              "Proper",
              "Collective",
              "Compound"
            ],
            "targets": [
              {
                "index": 0,
                "answer": "Proper",
                "explanation": "names one specific person, place, or thing, so it's capitalized"
              },
              {
                "index": 2,
                "answer": [
                  "Common",
                  "Collective"
                ],
                "explanation": "it is BOTH: common (a general name, not one specific one) and collective (a group acting as one unit)"
              },
              {
                "index": 6,
                "answer": [
                  "Common",
                  "Compound"
                ],
                "explanation": "it is BOTH: common (a general name, not one specific one) and compound (two words joined into one noun idea)"
              },
              {
                "index": 9,
                "answer": [
                  "Common",
                  "Compound"
                ],
                "explanation": "it is BOTH: common (a general name, not one specific one) and compound (two words joined into one noun idea)"
              },
              {
                "index": 12,
                "answer": "Common",
                "explanation": "names any person, place, or thing — not one specific one"
              }
            ]
          },
          {
            "id": "ag1_2",
            "label": "Sentence 2",
            "type": "pos-tagger",
            "sentence": [
              "The",
              "whole",
              "family",
              "gathered",
              "for",
              "breakfast",
              "before",
              "Kenley's",
              "cooking",
              "class."
            ],
            "answers": [
              "Adjective",
              "Adjective",
              "Noun",
              "Verb",
              "Preposition",
              "Noun",
              "Preposition",
              "Noun",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "article",
              "adjective",
              "common/collective noun",
              "action verb",
              "preposition",
              "common/compound noun",
              "preposition",
              "possessive proper noun",
              "noun used as an adjective (describes class)",
              "common noun"
            ]
          },
          {
            "id": "ag1_2c",
            "label": "Sentence 2 — Noun Check",
            "type": "concept-check",
            "sentence": [
              "The",
              "whole",
              "family",
              "gathered",
              "for",
              "breakfast",
              "before",
              "Kenley's",
              "cooking",
              "class."
            ],
            "prompt": "Now look at the nouns you found in this sentence. Is each one common, proper, collective, or compound?",
            "options": [
              "Common",
              "Proper",
              "Collective",
              "Compound"
            ],
            "targets": [
              {
                "index": 2,
                "answer": [
                  "Common",
                  "Collective"
                ],
                "explanation": "it is BOTH: common (a general name, not one specific one) and collective (a group acting as one unit)"
              },
              {
                "index": 5,
                "answer": [
                  "Common",
                  "Compound"
                ],
                "explanation": "it is BOTH: common (a general name, not one specific one) and compound (two words joined into one noun idea)"
              },
              {
                "index": 7,
                "answer": "Proper",
                "explanation": "names one specific person, place, or thing, so it's capitalized"
              },
              {
                "index": 9,
                "answer": [
                  "Common",
                  "Collective"
                ],
                "explanation": "it is BOTH: common (a general name, not one specific one) and collective (a group acting as one unit)"
              }
            ]
          },
          {
            "id": "ag1_3",
            "label": "Sentence 3",
            "type": "pos-tagger",
            "sentence": [
              "Kenley's",
              "cooking",
              "class",
              "visited",
              "a",
              "busy",
              "marketplace",
              "full",
              "of",
              "colorful",
              "spices."
            ],
            "answers": [
              "Noun",
              "Adjective",
              "Noun",
              "Verb",
              "Adjective",
              "Adjective",
              "Noun",
              "Adjective",
              "Preposition",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "possessive proper noun",
              "noun used as an adjective (describes class)",
              "common noun",
              "action verb",
              "article",
              "adjective",
              "common/compound noun",
              "adjective",
              "preposition",
              "adjective",
              "common noun"
            ]
          },
          {
            "id": "ag1_3c",
            "label": "Sentence 3 — Noun Check",
            "type": "concept-check",
            "sentence": [
              "Kenley's",
              "cooking",
              "class",
              "visited",
              "a",
              "busy",
              "marketplace",
              "full",
              "of",
              "colorful",
              "spices."
            ],
            "prompt": "Now look at the nouns you found in this sentence. Is each one common, proper, collective, or compound?",
            "options": [
              "Common",
              "Proper",
              "Collective",
              "Compound"
            ],
            "targets": [
              {
                "index": 0,
                "answer": "Proper",
                "explanation": "names one specific person, place, or thing, so it's capitalized"
              },
              {
                "index": 2,
                "answer": [
                  "Common",
                  "Collective"
                ],
                "explanation": "it is BOTH: common (a general name, not one specific one) and collective (a group acting as one unit)"
              },
              {
                "index": 6,
                "answer": [
                  "Common",
                  "Compound"
                ],
                "explanation": "it is BOTH: common (a general name, not one specific one) and compound (two words joined into one noun idea)"
              },
              {
                "index": 10,
                "answer": "Common",
                "explanation": "names any person, place, or thing — not one specific one"
              }
            ]
          },
          {
            "id": "ag1_4",
            "label": "Sentence 4",
            "type": "pos-tagger",
            "sentence": [
              "The",
              "gymnastics",
              "squad",
              "and",
              "the",
              "yoga",
              "group",
              "shared",
              "the",
              "same",
              "sunny",
              "studio."
            ],
            "answers": [
              "Adjective",
              "Adjective",
              "Noun",
              "Conjunction",
              "Adjective",
              "Adjective",
              "Noun",
              "Verb",
              "Adjective",
              "Adjective",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "article",
              "noun used as an adjective (describes squad)",
              "common/collective noun",
              "conjunction",
              "article",
              "noun used as an adjective (describes group)",
              "common/collective noun",
              "action verb",
              "article",
              "adjective",
              "adjective",
              "common noun"
            ]
          },
          {
            "id": "ag1_4c",
            "label": "Sentence 4 — Noun Check",
            "type": "concept-check",
            "sentence": [
              "The",
              "gymnastics",
              "squad",
              "and",
              "the",
              "yoga",
              "group",
              "shared",
              "the",
              "same",
              "sunny",
              "studio."
            ],
            "prompt": "Now look at the nouns you found in this sentence. Is each one common, proper, collective, or compound?",
            "options": [
              "Common",
              "Proper",
              "Collective",
              "Compound"
            ],
            "targets": [
              {
                "index": 2,
                "answer": [
                  "Common",
                  "Collective"
                ],
                "explanation": "it is BOTH: common (a general name, not one specific one) and collective (a group acting as one unit)"
              },
              {
                "index": 6,
                "answer": [
                  "Common",
                  "Collective"
                ],
                "explanation": "it is BOTH: common (a general name, not one specific one) and collective (a group acting as one unit)"
              },
              {
                "index": 11,
                "answer": "Common",
                "explanation": "names any person, place, or thing — not one specific one"
              }
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 2,
    "subjects": {
      "grammar": {
        "name": "Grammar",
        "tag": "Grammar Town · Unit 1, Week 2",
        "tasks": [
          {
            "id": "agl2",
            "label": "Week 2 Lesson: Pronouns Refresher — Plus Indefinite Pronouns",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Week 2: Pronouns Refresher — Plus Indefinite Pronouns</b></p><p style=\"opacity:.7;font-size:0.78rem;\">REVIEW / PRACTICE — Standards: L.4.1a (foundation)</p><p>Hi Adelyn!</p><p>Quick refresher: pronouns replace nouns so we don't repeat names constantly.</p><p>You already know subject pronouns (I, she, they), object pronouns (me, her, them), and possessive pronouns (my, her, their).</p><p><b>NEW</b>: <b>INDEFINITE</b> pronouns refer to people or things without naming exactly who or what — everyone, somebody, nothing, anything, few, several, all, both.</p><p>They're 'indefinite' because they're intentionally vague!'<b>EVERYONE</b> clapped when Adelyn landed the trick.'</p><p>Everyone doesn't name a specific person — it's indefinite, and it's still the subject of the sentence.</p><p style=\"opacity:.8;\">For each sentence, label the part of speech of every word.</p></div>"
          },
          {
            "id": "ag2_1",
            "label": "Sentence 1",
            "type": "pos-tagger",
            "sentence": [
              "Everyone",
              "cheered",
              "for",
              "her",
              "after",
              "she",
              "landed",
              "the",
              "cartwheel",
              "perfectly."
            ],
            "answers": [
              "Pronoun",
              "Verb",
              "Preposition",
              "Pronoun",
              "Conjunction",
              "Pronoun",
              "Verb",
              "Adjective",
              "Noun",
              "Adverb"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "indefinite pronoun (subject)",
              "action verb",
              "preposition",
              "object pronoun",
              "starts a clause here, so it works like a conjunction (you'll learn the name 'subordinating conjunction' in Week 25)",
              "subject pronoun",
              "action verb",
              "article",
              "noun",
              "adverb"
            ]
          },
          {
            "id": "ag2_1c",
            "label": "Sentence 1 — Pronoun Check",
            "type": "concept-check",
            "sentence": [
              "Everyone",
              "cheered",
              "for",
              "her",
              "after",
              "she",
              "landed",
              "the",
              "cartwheel",
              "perfectly."
            ],
            "prompt": "Now look at the pronouns you found in this sentence. Is each one subject, object, possessive, or indefinite?",
            "options": [
              "Subject",
              "Object",
              "Possessive",
              "Indefinite"
            ],
            "targets": [
              {
                "index": 0,
                "answer": [
                  "Subject",
                  "Indefinite"
                ],
                "explanation": "it is BOTH: an indefinite pronoun (no specific person or thing) and the subject in its clause"
              },
              {
                "index": 3,
                "answer": "Object",
                "explanation": "it's receiving the action, or following a preposition"
              },
              {
                "index": 5,
                "answer": "Subject",
                "explanation": "it's doing the action in its clause"
              }
            ]
          },
          {
            "id": "ag2_2",
            "label": "Sentence 2",
            "type": "pos-tagger",
            "sentence": [
              "Somebody",
              "left",
              "their",
              "yoga",
              "mat",
              "outside,",
              "but",
              "nobody",
              "knew",
              "whose",
              "it",
              "was."
            ],
            "answers": [
              "Pronoun",
              "Verb",
              "Pronoun",
              "Adjective",
              "Noun",
              "Adverb",
              "Conjunction",
              "Pronoun",
              "Verb",
              "Pronoun",
              "Pronoun",
              "Verb"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "indefinite pronoun",
              "action verb",
              "possessive pronoun",
              "noun used as an adjective (describes mat)",
              "noun",
              "adverb",
              "conjunction",
              "indefinite pronoun",
              "action verb",
              "possessive pronoun",
              "subject pronoun",
              "linking verb"
            ]
          },
          {
            "id": "ag2_2c",
            "label": "Sentence 2 — Pronoun Check",
            "type": "concept-check",
            "sentence": [
              "Somebody",
              "left",
              "their",
              "yoga",
              "mat",
              "outside,",
              "but",
              "nobody",
              "knew",
              "whose",
              "it",
              "was."
            ],
            "prompt": "Now look at the pronouns you found in this sentence. Is each one subject, object, possessive, or indefinite?",
            "options": [
              "Subject",
              "Object",
              "Possessive",
              "Indefinite"
            ],
            "targets": [
              {
                "index": 0,
                "answer": [
                  "Subject",
                  "Indefinite"
                ],
                "explanation": "it is BOTH: an indefinite pronoun (no specific person or thing) and the subject in its clause"
              },
              {
                "index": 2,
                "answer": "Possessive",
                "explanation": "it shows ownership"
              },
              {
                "index": 7,
                "answer": [
                  "Subject",
                  "Indefinite"
                ],
                "explanation": "it is BOTH: an indefinite pronoun (no specific person or thing) and the subject in its clause"
              },
              {
                "index": 9,
                "answer": "Possessive",
                "explanation": "it shows ownership"
              },
              {
                "index": 10,
                "answer": "Subject",
                "explanation": "it's doing the action in its clause"
              }
            ]
          },
          {
            "id": "ag2_3",
            "label": "Sentence 3",
            "type": "pos-tagger",
            "sentence": [
              "Everybody",
              "tried",
              "her",
              "new",
              "yoga",
              "pose,",
              "but",
              "nobody",
              "could",
              "balance",
              "as",
              "well",
              "as",
              "she",
              "could."
            ],
            "answers": [
              "Pronoun",
              "Verb",
              "Pronoun",
              "Adjective",
              "Adjective",
              "Noun",
              "Conjunction",
              "Pronoun",
              "Verb",
              "Verb",
              "Conjunction",
              "Adverb",
              "Conjunction",
              "Pronoun",
              "Verb"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "indefinite pronoun",
              "action verb",
              "possessive pronoun",
              "adjective",
              "noun used as an adjective (describes pose)",
              "noun",
              "conjunction",
              "indefinite pronoun",
              "helping verb",
              "action verb",
              "comparison word (part of \"as well as\")",
              "adverb",
              "comparison word (part of \"as well as\")",
              "subject pronoun",
              "helping verb"
            ]
          },
          {
            "id": "ag2_3c",
            "label": "Sentence 3 — Pronoun Check",
            "type": "concept-check",
            "sentence": [
              "Everybody",
              "tried",
              "her",
              "new",
              "yoga",
              "pose,",
              "but",
              "nobody",
              "could",
              "balance",
              "as",
              "well",
              "as",
              "she",
              "could."
            ],
            "prompt": "Now look at the pronouns you found in this sentence. Is each one subject, object, possessive, or indefinite?",
            "options": [
              "Subject",
              "Object",
              "Possessive",
              "Indefinite"
            ],
            "targets": [
              {
                "index": 0,
                "answer": [
                  "Subject",
                  "Indefinite"
                ],
                "explanation": "it is BOTH: an indefinite pronoun (no specific person or thing) and the subject in its clause"
              },
              {
                "index": 2,
                "answer": "Possessive",
                "explanation": "it shows ownership"
              },
              {
                "index": 7,
                "answer": [
                  "Subject",
                  "Indefinite"
                ],
                "explanation": "it is BOTH: an indefinite pronoun (no specific person or thing) and the subject in its clause"
              },
              {
                "index": 13,
                "answer": "Subject",
                "explanation": "it's doing the action in its clause"
              }
            ]
          },
          {
            "id": "ag2_4",
            "label": "Sentence 4",
            "type": "pos-tagger",
            "sentence": [
              "Someone",
              "left",
              "something",
              "delicious",
              "on",
              "the",
              "counter,",
              "and",
              "everyone",
              "wanted",
              "a",
              "bite."
            ],
            "answers": [
              "Pronoun",
              "Verb",
              "Pronoun",
              "Adjective",
              "Preposition",
              "Adjective",
              "Noun",
              "Conjunction",
              "Pronoun",
              "Verb",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "indefinite pronoun",
              "action verb",
              "indefinite pronoun",
              "adjective",
              "preposition",
              "article",
              "noun",
              "conjunction",
              "indefinite pronoun",
              "action verb",
              "article",
              "noun"
            ]
          },
          {
            "id": "ag2_4c",
            "label": "Sentence 4 — Pronoun Check",
            "type": "concept-check",
            "sentence": [
              "Someone",
              "left",
              "something",
              "delicious",
              "on",
              "the",
              "counter,",
              "and",
              "everyone",
              "wanted",
              "a",
              "bite."
            ],
            "prompt": "Now look at the pronouns you found in this sentence. Is each one subject, object, possessive, or indefinite?",
            "options": [
              "Subject",
              "Object",
              "Possessive",
              "Indefinite"
            ],
            "targets": [
              {
                "index": 0,
                "answer": [
                  "Subject",
                  "Indefinite"
                ],
                "explanation": "it is BOTH: an indefinite pronoun (no specific person or thing) and the subject in its clause"
              },
              {
                "index": 2,
                "answer": [
                  "Object",
                  "Indefinite"
                ],
                "explanation": "it is BOTH: an indefinite pronoun (no specific person or thing) and the object in its clause"
              },
              {
                "index": 8,
                "answer": [
                  "Subject",
                  "Indefinite"
                ],
                "explanation": "it is BOTH: an indefinite pronoun (no specific person or thing) and the subject in its clause"
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
      "grammar": {
        "name": "Grammar",
        "tag": "Grammar Town · Unit 1, Week 3",
        "tasks": [
          {
            "id": "agl3",
            "label": "Week 3 Lesson: Verbs Refresher — Action, Linking & Helping, All at Once",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Week 3: Verbs Refresher — Action, Linking &amp; Helping, All at Once</b></p><p style=\"opacity:.7;font-size:0.78rem;\">REVIEW / PRACTICE — Standards: L.4.1 (foundation for L.4.1b)</p><p>Hi Adelyn!</p><p>Fast refresher, three verb types in one week since you already met all three in Island: <b>ACTION</b> verbs show something happening (leaped, whisked).</p><p><b>LINKING</b> verbs connect the subject to a description, like an equals sign (is, seems, smelled).</p><p><b>HELPING</b> verbs team up with a main verb (was, has, will, can).</p><p>Since you've got this, here's the test to nail it every time: try swapping the verb for an equals sign.</p><p>If the sentence still makes sense, it's linking.</p><p>If not, it's action (or working with a helper).</p><p>This week's sentences mix all three types — see if you can sort them without hesitating.</p><p style=\"opacity:.8;\">For each sentence, label the part of speech of every word.</p></div>"
          },
          {
            "id": "ag3_1",
            "label": "Sentence 1",
            "type": "pos-tagger",
            "sentence": [
              "The",
              "soup",
              "smelled",
              "delicious,",
              "and",
              "Kenley",
              "was",
              "stirring",
              "it",
              "carefully."
            ],
            "answers": [
              "Adjective",
              "Noun",
              "Verb",
              "Adjective",
              "Conjunction",
              "Noun",
              "Verb",
              "Verb",
              "Pronoun",
              "Adverb"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "article",
              "noun",
              "linking verb (soup = delicious)",
              "describes soup (a predicate adjective — more on this in Week 15)",
              "conjunction",
              "proper noun",
              "helping verb",
              "action verb",
              "object pronoun",
              "adverb"
            ]
          },
          {
            "id": "ag3_1c",
            "label": "Sentence 1 — Verb Check",
            "type": "concept-check",
            "sentence": [
              "The",
              "soup",
              "smelled",
              "delicious,",
              "and",
              "Kenley",
              "was",
              "stirring",
              "it",
              "carefully."
            ],
            "prompt": "Now look at the verbs you found in this sentence. Is each one action, linking, or helping?",
            "options": [
              "Action",
              "Linking",
              "Helping"
            ],
            "targets": [
              {
                "index": 2,
                "answer": "Linking",
                "explanation": "connects the subject to a word that renames or describes it, instead of showing action"
              },
              {
                "index": 6,
                "answer": "Helping",
                "explanation": "works with the main verb to show tense, ability, or mood"
              },
              {
                "index": 7,
                "answer": "Action",
                "explanation": "shows something the subject does"
              }
            ]
          },
          {
            "id": "ag3_2",
            "label": "Sentence 2",
            "type": "pos-tagger",
            "sentence": [
              "Adelyn",
              "has",
              "practiced",
              "every",
              "day,",
              "and",
              "she",
              "seems",
              "much",
              "stronger",
              "now."
            ],
            "answers": [
              "Noun",
              "Verb",
              "Verb",
              "Adjective",
              "Noun",
              "Conjunction",
              "Pronoun",
              "Verb",
              "Adverb",
              "Adjective",
              "Adverb"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "proper noun",
              "helping verb",
              "action verb",
              "adjective",
              "noun",
              "conjunction",
              "subject pronoun",
              "linking verb",
              "adverb (describes stronger)",
              "describes she (a predicate adjective — more in Week 15)",
              "adverb"
            ]
          },
          {
            "id": "ag3_2c",
            "label": "Sentence 2 — Verb Check",
            "type": "concept-check",
            "sentence": [
              "Adelyn",
              "has",
              "practiced",
              "every",
              "day,",
              "and",
              "she",
              "seems",
              "much",
              "stronger",
              "now."
            ],
            "prompt": "Now look at the verbs you found in this sentence. Is each one action, linking, or helping?",
            "options": [
              "Action",
              "Linking",
              "Helping"
            ],
            "targets": [
              {
                "index": 1,
                "answer": "Helping",
                "explanation": "works with the main verb to show tense, ability, or mood"
              },
              {
                "index": 2,
                "answer": "Action",
                "explanation": "shows something the subject does"
              },
              {
                "index": 7,
                "answer": "Linking",
                "explanation": "connects the subject to a word that renames or describes it, instead of showing action"
              }
            ]
          },
          {
            "id": "ag3_3",
            "label": "Sentence 3",
            "type": "pos-tagger",
            "sentence": [
              "Adelyn",
              "is",
              "stretching",
              "now,",
              "but",
              "she",
              "was",
              "resting",
              "earlier."
            ],
            "answers": [
              "Noun",
              "Verb",
              "Verb",
              "Adverb",
              "Conjunction",
              "Pronoun",
              "Verb",
              "Verb",
              "Adverb"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "proper noun",
              "helping verb",
              "action verb",
              "adverb",
              "conjunction",
              "subject pronoun",
              "helping verb",
              "action verb",
              "adverb"
            ]
          },
          {
            "id": "ag3_3c",
            "label": "Sentence 3 — Verb Check",
            "type": "concept-check",
            "sentence": [
              "Adelyn",
              "is",
              "stretching",
              "now,",
              "but",
              "she",
              "was",
              "resting",
              "earlier."
            ],
            "prompt": "Now look at the verbs you found in this sentence. Is each one action, linking, or helping?",
            "options": [
              "Action",
              "Linking",
              "Helping"
            ],
            "targets": [
              {
                "index": 1,
                "answer": "Helping",
                "explanation": "works with the main verb to show tense, ability, or mood"
              },
              {
                "index": 2,
                "answer": "Action",
                "explanation": "shows something the subject does"
              },
              {
                "index": 6,
                "answer": "Helping",
                "explanation": "works with the main verb to show tense, ability, or mood"
              },
              {
                "index": 7,
                "answer": "Action",
                "explanation": "shows something the subject does"
              }
            ]
          },
          {
            "id": "ag3_4",
            "label": "Sentence 4",
            "type": "pos-tagger",
            "sentence": [
              "The",
              "bread",
              "tasted",
              "amazing,",
              "and",
              "Kenley",
              "felt",
              "very",
              "proud",
              "of",
              "her",
              "baking."
            ],
            "answers": [
              "Adjective",
              "Noun",
              "Verb",
              "Adjective",
              "Conjunction",
              "Noun",
              "Verb",
              "Adverb",
              "Adjective",
              "Preposition",
              "Pronoun",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "article",
              "noun",
              "linking verb",
              "predicate adjective",
              "conjunction",
              "proper noun",
              "linking verb",
              "adverb",
              "predicate adjective",
              "preposition",
              "possessive pronoun",
              "gerund, acting as a noun (object of 'of')"
            ]
          },
          {
            "id": "ag3_4c",
            "label": "Sentence 4 — Verb Check",
            "type": "concept-check",
            "sentence": [
              "The",
              "bread",
              "tasted",
              "amazing,",
              "and",
              "Kenley",
              "felt",
              "very",
              "proud",
              "of",
              "her",
              "baking."
            ],
            "prompt": "Now look at the verbs you found in this sentence. Is each one action, linking, or helping?",
            "options": [
              "Action",
              "Linking",
              "Helping"
            ],
            "targets": [
              {
                "index": 2,
                "answer": "Linking",
                "explanation": "connects the subject to a word that renames or describes it, instead of showing action"
              },
              {
                "index": 6,
                "answer": "Linking",
                "explanation": "connects the subject to a word that renames or describes it, instead of showing action"
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
      "grammar": {
        "name": "Grammar",
        "tag": "Grammar Town · Unit 1, Week 4",
        "tasks": [
          {
            "id": "agl4",
            "label": "Week 4 Lesson: NEW: Verbals — Gerunds (Verbs Acting as Nouns)",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Week 4: NEW: Verbals — Gerunds (Verbs Acting as Nouns)</b></p><p style=\"opacity:.7;font-size:0.78rem;\">NEW TOPIC — Standards: L.4.1 (Town-level enrichment beyond CCSS minimum; supports W.4.3 sentence variety)</p><p>Hi Adelyn!</p><p>Here's something totally new, straight from Town-level grammar: a <b>VERBAL</b> is a verb form that stopped acting like a verb and started doing a different job.</p><p>The first kind is the <b>GERUND</b> — a verb + -ing that acts as a <b>NOUN</b>.'<b>CARTWHEELING</b> is Adelyn's favorite activity.'</p><p>Cartwheeling looks like a verb, but here it's the <b>SUBJECT</b> of the sentence — a noun job!</p><p>Compare that to 'Adelyn is cartwheeling,' where cartwheeling is just part of the verb.</p><p>The trick: ask what job the -ing word is doing.</p><p>If it's acting as a person/place/thing (subject, object, etc.), it's a gerund.</p><p>If it's paired with a helping verb showing action happening, it's just a regular verb.</p><p>In Step 1 you still label a gerund as a <b>NOUN</b>, because that is its job. Step 2 is where you name it.</p><p style=\"opacity:.8;\"><b>Step 1:</b> For each sentence, label every word using only the 8 parts of speech, based on the job the word is doing in that sentence.</p><p style=\"opacity:.8;\"><b>Step 2:</b> After a sentence, a Gerund Check asks you to spot the gerunds.</p></div>"
          },
          {
            "id": "ag4_1",
            "label": "Sentence 1",
            "type": "pos-tagger",
            "sentence": [
              "Cooking",
              "is",
              "Kenley's",
              "favorite",
              "activity",
              "after",
              "school."
            ],
            "answers": [
              "Noun",
              "Verb",
              "Noun",
              "Adjective",
              "Noun",
              "Preposition",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "noun: it names the activity, so it does a noun's job here",
              "linking verb",
              "possessive proper noun",
              "adjective",
              "noun",
              "preposition",
              "noun"
            ]
          },
          {
            "id": "ag4_1c",
            "label": "Sentence 1 — Gerund Check",
            "type": "concept-check",
            "sentence": [
              "Cooking",
              "is",
              "Kenley's",
              "favorite",
              "activity",
              "after",
              "school."
            ],
            "prompt": "Now look at the -ing words in this sentence. Each one is built from a verb. Is it working as a regular verb, or as a gerund (a verb form doing a noun's job)?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 0,
                "answer": "Gerund",
                "explanation": "an -ing verb form that names an activity, so it works as a noun"
              }
            ]
          },
          {
            "id": "ag4_2",
            "label": "Sentence 2",
            "type": "pos-tagger",
            "sentence": [
              "Adelyn",
              "is",
              "stretching",
              "now,",
              "but",
              "she",
              "loves",
              "stretching",
              "before",
              "practice."
            ],
            "answers": [
              "Noun",
              "Verb",
              "Verb",
              "Adverb",
              "Conjunction",
              "Pronoun",
              "Verb",
              "Noun",
              "Preposition",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "proper noun",
              "helping verb",
              "action verb (part of the verb phrase \"is stretching\")",
              "adverb",
              "conjunction",
              "subject pronoun",
              "action verb",
              "noun: it names the activity, so it does a noun's job here",
              "preposition",
              "noun"
            ]
          },
          {
            "id": "ag4_2c",
            "label": "Sentence 2 — Gerund Check",
            "type": "concept-check",
            "sentence": [
              "Adelyn",
              "is",
              "stretching",
              "now,",
              "but",
              "she",
              "loves",
              "stretching",
              "before",
              "practice."
            ],
            "prompt": "Now look at the -ing words in this sentence. Each one is built from a verb. Is it working as a regular verb, or as a gerund (a verb form doing a noun's job)?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 2,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 7,
                "answer": "Gerund",
                "explanation": "an -ing verb form that names an activity, so it works as a noun"
              }
            ]
          },
          {
            "id": "ag4_3",
            "label": "Sentence 3",
            "type": "pos-tagger",
            "sentence": [
              "Traveling",
              "is",
              "Adelyn's",
              "favorite",
              "part",
              "of",
              "homeschooling."
            ],
            "answers": [
              "Noun",
              "Verb",
              "Noun",
              "Adjective",
              "Noun",
              "Preposition",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "noun: it names the activity, so it does a noun's job here",
              "linking verb",
              "possessive proper noun",
              "adjective",
              "noun",
              "preposition",
              "noun: it names the activity, so it does a noun's job here"
            ]
          },
          {
            "id": "ag4_3c",
            "label": "Sentence 3 — Gerund Check",
            "type": "concept-check",
            "sentence": [
              "Traveling",
              "is",
              "Adelyn's",
              "favorite",
              "part",
              "of",
              "homeschooling."
            ],
            "prompt": "Now look at the -ing words in this sentence. Each one is built from a verb. Is it working as a regular verb, or as a gerund (a verb form doing a noun's job)?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 0,
                "answer": "Gerund",
                "explanation": "an -ing verb form that names an activity, so it works as a noun"
              },
              {
                "index": 6,
                "answer": "Gerund",
                "explanation": "an -ing verb form that names an activity, so it works as a noun"
              }
            ]
          },
          {
            "id": "ag4_4",
            "label": "Sentence 4",
            "type": "pos-tagger",
            "sentence": [
              "Kenley",
              "enjoys",
              "baking",
              "more",
              "than",
              "cooking",
              "pasta."
            ],
            "answers": [
              "Noun",
              "Verb",
              "Noun",
              "Adverb",
              "Conjunction",
              "Noun",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "proper noun",
              "action verb",
              "noun: it names the activity, so it does a noun's job here",
              "adverb",
              "comparison word",
              "noun: it names the activity, so it does a noun's job here",
              "noun (what is being cooked)"
            ]
          },
          {
            "id": "ag4_4c",
            "label": "Sentence 4 — Gerund Check",
            "type": "concept-check",
            "sentence": [
              "Kenley",
              "enjoys",
              "baking",
              "more",
              "than",
              "cooking",
              "pasta."
            ],
            "prompt": "Now look at the -ing words in this sentence. Each one is built from a verb. Is it working as a regular verb, or as a gerund (a verb form doing a noun's job)?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 2,
                "answer": "Gerund",
                "explanation": "an -ing verb form that names an activity, so it works as a noun"
              },
              {
                "index": 5,
                "answer": "Gerund",
                "explanation": "an -ing verb form that names an activity, so it works as a noun"
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
      "grammar": {
        "name": "Grammar",
        "tag": "Grammar Town · Unit 1, Week 5",
        "tasks": [
          {
            "id": "agl5",
            "label": "Week 5 Lesson: NEW: Verbals — Participles (Verbs Acting as Adjectives)",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Week 5: NEW: Verbals — Participles (Verbs Acting as Adjectives)</b></p><p style=\"opacity:.7;font-size:0.78rem;\">NEW TOPIC — Standards: L.4.1 (Town-level enrichment; supports W.4.3)</p><p>Hi Adelyn!</p><p>Second verbal: the <b>PARTICIPLE</b> — a verb form (usually ending in -ing or -ed) that acts as an <b>ADJECTIVE</b>, describing a noun.'The <b>GIGGLING</b> gymnasts stretched.'</p><p>Giggling looks like a verb, but here it describes gymnasts — an adjective job!</p><p>Same word family as a gerund, totally different job depending on what it's doing in the sentence.</p><p>Compare: 'Adelyn was <b>LAUGHING</b>' (verb, part of the action) vs.</p><p>'the <b>LAUGHING</b> chef' (participle, describing chef).</p><p>Same word, different job — grammar is sneaky like that!</p><p>In Step 1 you label a participle as an <b>ADJECTIVE</b>, because that is its job. Step 2 is where you name it.</p><p style=\"opacity:.8;\"><b>Step 1:</b> For each sentence, label every word using only the 8 parts of speech, based on the job the word is doing in that sentence.</p><p style=\"opacity:.8;\"><b>Step 2:</b> After a sentence, a Participle Check asks you to spot the participles.</p></div>"
          },
          {
            "id": "ag5_1",
            "label": "Sentence 1",
            "type": "pos-tagger",
            "sentence": [
              "The",
              "exhausted",
              "gymnasts",
              "collapsed",
              "onto",
              "the",
              "mats",
              "after",
              "the",
              "giggling",
              "coach",
              "called",
              "a",
              "break."
            ],
            "answers": [
              "Adjective",
              "Adjective",
              "Noun",
              "Verb",
              "Preposition",
              "Adjective",
              "Noun",
              "Conjunction",
              "Adjective",
              "Adjective",
              "Noun",
              "Verb",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "article",
              "adjective: it describes a noun here",
              "noun",
              "action verb",
              "preposition",
              "article",
              "noun",
              "starts a clause here, acts like a conjunction (subordinating conjunction, Week 25)",
              "article",
              "adjective: it describes a noun here",
              "noun",
              "action verb",
              "article",
              "noun"
            ]
          },
          {
            "id": "ag5_1c",
            "label": "Sentence 1 — Participle Check",
            "type": "concept-check",
            "sentence": [
              "The",
              "exhausted",
              "gymnasts",
              "collapsed",
              "onto",
              "the",
              "mats",
              "after",
              "the",
              "giggling",
              "coach",
              "called",
              "a",
              "break."
            ],
            "prompt": "Now look at the words built from verbs (-ing and -ed words) in this sentence. Is each one a regular verb, a gerund (doing a noun's job), or a participle (describing a noun)?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 1,
                "answer": "Participle",
                "explanation": "built from a verb, but here it describes a noun, so it works as an adjective"
              },
              {
                "index": 3,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 9,
                "answer": "Participle",
                "explanation": "built from a verb, but here it describes a noun, so it works as an adjective"
              },
              {
                "index": 11,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              }
            ]
          },
          {
            "id": "ag5_2",
            "label": "Sentence 2",
            "type": "pos-tagger",
            "sentence": [
              "Kenley",
              "served",
              "the",
              "steaming",
              "soup",
              "to",
              "her",
              "smiling",
              "little",
              "sister."
            ],
            "answers": [
              "Noun",
              "Verb",
              "Adjective",
              "Adjective",
              "Noun",
              "Preposition",
              "Pronoun",
              "Adjective",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "proper noun",
              "action verb",
              "article",
              "adjective: it describes a noun here",
              "noun",
              "preposition",
              "possessive pronoun",
              "adjective: it describes a noun here",
              "adjective",
              "noun"
            ]
          },
          {
            "id": "ag5_2c",
            "label": "Sentence 2 — Participle Check",
            "type": "concept-check",
            "sentence": [
              "Kenley",
              "served",
              "the",
              "steaming",
              "soup",
              "to",
              "her",
              "smiling",
              "little",
              "sister."
            ],
            "prompt": "Now look at the words built from verbs (-ing and -ed words) in this sentence. Is each one a regular verb, a gerund (doing a noun's job), or a participle (describing a noun)?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 1,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 3,
                "answer": "Participle",
                "explanation": "built from a verb, but here it describes a noun, so it works as an adjective"
              },
              {
                "index": 7,
                "answer": "Participle",
                "explanation": "built from a verb, but here it describes a noun, so it works as an adjective"
              }
            ]
          },
          {
            "id": "ag5_3",
            "label": "Sentence 3",
            "type": "pos-tagger",
            "sentence": [
              "The",
              "smiling",
              "coach",
              "clapped",
              "for",
              "the",
              "exhausted",
              "but",
              "excited",
              "gymnasts."
            ],
            "answers": [
              "Adjective",
              "Adjective",
              "Noun",
              "Verb",
              "Preposition",
              "Adjective",
              "Adjective",
              "Conjunction",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "article",
              "adjective: it describes a noun here",
              "noun",
              "action verb",
              "preposition",
              "article",
              "adjective: it describes a noun here",
              "conjunction",
              "adjective: it describes a noun here",
              "noun"
            ]
          },
          {
            "id": "ag5_3c",
            "label": "Sentence 3 — Participle Check",
            "type": "concept-check",
            "sentence": [
              "The",
              "smiling",
              "coach",
              "clapped",
              "for",
              "the",
              "exhausted",
              "but",
              "excited",
              "gymnasts."
            ],
            "prompt": "Now look at the words built from verbs (-ing and -ed words) in this sentence. Is each one a regular verb, a gerund (doing a noun's job), or a participle (describing a noun)?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 1,
                "answer": "Participle",
                "explanation": "built from a verb, but here it describes a noun, so it works as an adjective"
              },
              {
                "index": 3,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 6,
                "answer": "Participle",
                "explanation": "built from a verb, but here it describes a noun, so it works as an adjective"
              },
              {
                "index": 8,
                "answer": "Participle",
                "explanation": "built from a verb, but here it describes a noun, so it works as an adjective"
              }
            ]
          },
          {
            "id": "ag5_4",
            "label": "Sentence 4",
            "type": "pos-tagger",
            "sentence": [
              "Kenley",
              "poured",
              "the",
              "sizzling",
              "batter",
              "into",
              "a",
              "waiting",
              "pan."
            ],
            "answers": [
              "Noun",
              "Verb",
              "Adjective",
              "Adjective",
              "Noun",
              "Preposition",
              "Adjective",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "proper noun",
              "action verb",
              "article",
              "adjective: it describes a noun here",
              "noun",
              "preposition",
              "article",
              "adjective: it describes a noun here",
              "noun"
            ]
          },
          {
            "id": "ag5_4c",
            "label": "Sentence 4 — Participle Check",
            "type": "concept-check",
            "sentence": [
              "Kenley",
              "poured",
              "the",
              "sizzling",
              "batter",
              "into",
              "a",
              "waiting",
              "pan."
            ],
            "prompt": "Now look at the words built from verbs (-ing and -ed words) in this sentence. Is each one a regular verb, a gerund (doing a noun's job), or a participle (describing a noun)?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 1,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 3,
                "answer": "Participle",
                "explanation": "built from a verb, but here it describes a noun, so it works as an adjective"
              },
              {
                "index": 7,
                "answer": "Participle",
                "explanation": "built from a verb, but here it describes a noun, so it works as an adjective"
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
      "grammar": {
        "name": "Grammar",
        "tag": "Grammar Town · Unit 1, Week 6",
        "tasks": [
          {
            "id": "agl6",
            "label": "Week 6 Lesson: NEW: Verbals — Infinitives (To + Verb)",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Week 6: NEW: Verbals — Infinitives (To + Verb)</b></p><p style=\"opacity:.7;font-size:0.78rem;\">NEW TOPIC — Standards: L.4.1 (Town-level enrichment; supports W.4.3)</p><p>Hi Adelyn!</p><p>Third and final verbal: the <b>INFINITIVE</b> — 'to' + a verb (to jump, to cook, to travel).</p><p>Infinitives can act as a noun, an adjective, or an adverb, depending on the sentence.'Adelyn wants TO <b>PRACTICE</b> every morning.'</p><p>To practice acts as a noun (wants <b>WHAT</b>? to practice).</p><p>'Kenley has a recipe TO <b>TRY</b>.'</p><p>To try acts as an adjective (describes which recipe).</p><p>'They traveled TO <b>EXPLORE</b> new places.'</p><p>To explore acts as an adverb (traveled <b>WHY</b>?).</p><p>Don't worry about perfectly labeling noun/adjective/adverb infinitive jobs yet — for now, just get great at <b>SPOTTING</b> an infinitive: 'to' + a verb, working as a team.</p><p>In Step 1, label both \"to\" and its verb as <b>VERB</b> when they work as a team (\"to master\"). But \"to\" before a noun (\"to the market\") is a <b>PREPOSITION</b>. Step 2 is where you name the infinitive.</p><p style=\"opacity:.8;\"><b>Step 1:</b> For each sentence, label every word using only the 8 parts of speech, based on the job the word is doing in that sentence.</p><p style=\"opacity:.8;\"><b>Step 2:</b> After a sentence, a Infinitive Check asks you to spot the infinitives.</p></div>"
          },
          {
            "id": "ag6_1",
            "label": "Sentence 1",
            "type": "pos-tagger",
            "sentence": [
              "Adelyn",
              "is",
              "ready",
              "to",
              "master",
              "a",
              "backflip",
              "by",
              "the",
              "end",
              "of",
              "the",
              "year."
            ],
            "answers": [
              "Noun",
              "Verb",
              "Adjective",
              "Verb",
              "Verb",
              "Adjective",
              "Noun",
              "Preposition",
              "Adjective",
              "Noun",
              "Preposition",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "proper noun",
              "linking verb",
              "describes Adelyn (predicate adjective)",
              "\"to\" and the verb after it work together as one verb form",
              "verb: it follows \"to\"",
              "article",
              "compound noun",
              "preposition",
              "article",
              "noun",
              "preposition",
              "article",
              "noun"
            ]
          },
          {
            "id": "ag6_1c",
            "label": "Sentence 1 — Infinitive Check",
            "type": "concept-check",
            "sentence": [
              "Adelyn",
              "is",
              "ready",
              "to",
              "master",
              "a",
              "backflip",
              "by",
              "the",
              "end",
              "of",
              "the",
              "year."
            ],
            "prompt": "Now look at the words built from verbs in this sentence. For an infinitive, tap the verb that comes after \"to\". Is each one a regular verb, a gerund, a participle, or an infinitive?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Infinitive",
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 4,
                "answer": "Infinitive",
                "explanation": "the verb that follows \"to\"; \"to\" and the verb work as a team"
              }
            ]
          },
          {
            "id": "ag6_2",
            "label": "Sentence 2",
            "type": "pos-tagger",
            "sentence": [
              "Kenley",
              "traveled",
              "to",
              "the",
              "market",
              "to",
              "buy",
              "fresh",
              "mangoes",
              "for",
              "breakfast."
            ],
            "answers": [
              "Noun",
              "Verb",
              "Preposition",
              "Adjective",
              "Noun",
              "Verb",
              "Verb",
              "Adjective",
              "Noun",
              "Preposition",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "proper noun",
              "action verb",
              "preposition (with 'the market')",
              "article",
              "noun",
              "\"to\" and the verb after it work together as one verb form",
              "verb: it follows \"to\"",
              "adjective",
              "noun",
              "preposition",
              "noun"
            ]
          },
          {
            "id": "ag6_2c",
            "label": "Sentence 2 — Infinitive Check",
            "type": "concept-check",
            "sentence": [
              "Kenley",
              "traveled",
              "to",
              "the",
              "market",
              "to",
              "buy",
              "fresh",
              "mangoes",
              "for",
              "breakfast."
            ],
            "prompt": "Now look at the words built from verbs in this sentence. For an infinitive, tap the verb that comes after \"to\". Is each one a regular verb, a gerund, a participle, or an infinitive?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Infinitive",
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 1,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 6,
                "answer": "Infinitive",
                "explanation": "the verb that follows \"to\"; \"to\" and the verb work as a team"
              }
            ]
          },
          {
            "id": "ag6_3",
            "label": "Sentence 3",
            "type": "pos-tagger",
            "sentence": [
              "Adelyn",
              "stopped",
              "to",
              "stretch",
              "before",
              "she",
              "continued",
              "to",
              "practice."
            ],
            "answers": [
              "Noun",
              "Verb",
              "Verb",
              "Verb",
              "Conjunction",
              "Pronoun",
              "Verb",
              "Verb",
              "Verb"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "proper noun",
              "action verb",
              "\"to\" and the verb after it work together as one verb form",
              "verb: it follows \"to\"",
              "subordinating conjunction",
              "subject pronoun",
              "action verb",
              "\"to\" and the verb after it work together as one verb form",
              "verb: it follows \"to\""
            ]
          },
          {
            "id": "ag6_3c",
            "label": "Sentence 3 — Infinitive Check",
            "type": "concept-check",
            "sentence": [
              "Adelyn",
              "stopped",
              "to",
              "stretch",
              "before",
              "she",
              "continued",
              "to",
              "practice."
            ],
            "prompt": "Now look at the words built from verbs in this sentence. For an infinitive, tap the verb that comes after \"to\". Is each one a regular verb, a gerund, a participle, or an infinitive?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Infinitive",
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 1,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 3,
                "answer": "Infinitive",
                "explanation": "the verb that follows \"to\"; \"to\" and the verb work as a team"
              },
              {
                "index": 6,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 8,
                "answer": "Infinitive",
                "explanation": "the verb that follows \"to\"; \"to\" and the verb work as a team"
              }
            ]
          },
          {
            "id": "ag6_4",
            "label": "Sentence 4",
            "type": "pos-tagger",
            "sentence": [
              "Kenley",
              "needs",
              "more",
              "spices",
              "to",
              "finish",
              "the",
              "recipe."
            ],
            "answers": [
              "Noun",
              "Verb",
              "Adjective",
              "Noun",
              "Verb",
              "Verb",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "proper noun",
              "action verb",
              "adjective",
              "noun",
              "\"to\" and the verb after it work together as one verb form",
              "verb: it follows \"to\"",
              "article",
              "noun"
            ]
          },
          {
            "id": "ag6_4c",
            "label": "Sentence 4 — Infinitive Check",
            "type": "concept-check",
            "sentence": [
              "Kenley",
              "needs",
              "more",
              "spices",
              "to",
              "finish",
              "the",
              "recipe."
            ],
            "prompt": "Now look at the words built from verbs in this sentence. For an infinitive, tap the verb that comes after \"to\". Is each one a regular verb, a gerund, a participle, or an infinitive?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Infinitive",
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 5,
                "answer": "Infinitive",
                "explanation": "the verb that follows \"to\"; \"to\" and the verb work as a team"
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
      "grammar": {
        "name": "Grammar",
        "tag": "Grammar Town · Unit 1, Week 7",
        "tasks": [
          {
            "id": "agl7",
            "label": "Week 7 Lesson: Adjectives Refresher — Plus Demonstrative & Interrogative Pronouns",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Week 7: Adjectives Refresher — Plus Demonstrative &amp; Interrogative Pronouns</b></p><p style=\"opacity:.7;font-size:0.78rem;\">REVIEW / PRACTICE — Standards: L.4.1a, L.4.1d (foundation)</p><p>Hi Adelyn!</p><p>Quick refresher: adjectives describe nouns (color, size, number, feeling), and a/an/the are special article-adjectives.</p><p>You've had this since Island.</p><p><b>NEW</b>: <b>DEMONSTRATIVE</b> words (this, that, these, those) point to a specific noun.</p><p>They're pronouns when they stand alone ('<b>THAT</b> is my mat') and adjectives when they describe a noun right next to them ('<b>THAT</b> mat is mine').</p><p><b>INTERROGATIVE</b> pronouns (who, what, which, whose) ask a question — and you'll meet them again later this year in a much bigger role!</p><p>This week, spot the regular adjectives <b>AND</b> these two new pronoun types.</p><p style=\"opacity:.8;\">For each sentence, label the part of speech of every word.</p></div>"
          },
          {
            "id": "ag7_1",
            "label": "Sentence 1",
            "type": "pos-tagger",
            "sentence": [
              "Which",
              "spicy",
              "curry",
              "did",
              "Kenley",
              "make,",
              "and",
              "is",
              "that",
              "tiny",
              "bowl",
              "mine?"
            ],
            "answers": [
              "Pronoun",
              "Adjective",
              "Noun",
              "Verb",
              "Noun",
              "Verb",
              "Conjunction",
              "Verb",
              "Adjective",
              "Adjective",
              "Noun",
              "Pronoun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "interrogative word, used here as an adjective describing curry",
              "adjective",
              "noun",
              "helping verb",
              "proper noun",
              "action verb",
              "conjunction",
              "linking verb",
              "demonstrative adjective, describes bowl",
              "adjective",
              "noun",
              "possessive pronoun"
            ]
          },
          {
            "id": "ag7_2",
            "label": "Sentence 2",
            "type": "pos-tagger",
            "sentence": [
              "This",
              "new",
              "leotard",
              "is",
              "comfortable,",
              "but",
              "those",
              "old",
              "ones",
              "felt",
              "scratchy."
            ],
            "answers": [
              "Adjective",
              "Adjective",
              "Noun",
              "Verb",
              "Adjective",
              "Conjunction",
              "Adjective",
              "Adjective",
              "Pronoun",
              "Verb",
              "Adjective"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "demonstrative adjective, describes leotard",
              "adjective",
              "noun",
              "linking verb",
              "predicate adjective",
              "conjunction",
              "demonstrative adjective, describes ones",
              "adjective",
              "indefinite pronoun",
              "linking verb",
              "predicate adjective"
            ]
          },
          {
            "id": "ag7_3",
            "label": "Sentence 3",
            "type": "pos-tagger",
            "sentence": [
              "What",
              "flavor",
              "is",
              "this",
              "soup,",
              "and",
              "which",
              "spice",
              "did",
              "you",
              "add?"
            ],
            "answers": [
              "Adjective",
              "Noun",
              "Verb",
              "Adjective",
              "Noun",
              "Conjunction",
              "Adjective",
              "Noun",
              "Verb",
              "Pronoun",
              "Verb"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "interrogative adjective, describes flavor",
              "noun",
              "linking verb",
              "demonstrative adjective, describes soup",
              "noun",
              "conjunction",
              "interrogative adjective, describes spice",
              "noun",
              "helping verb",
              "subject pronoun",
              "action verb"
            ]
          },
          {
            "id": "ag7_4",
            "label": "Sentence 4",
            "type": "pos-tagger",
            "sentence": [
              "These",
              "purple",
              "leotards",
              "are",
              "softer",
              "than",
              "those",
              "old",
              "blue",
              "ones."
            ],
            "answers": [
              "Adjective",
              "Adjective",
              "Noun",
              "Verb",
              "Adjective",
              "Conjunction",
              "Adjective",
              "Adjective",
              "Adjective",
              "Pronoun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "demonstrative adjective, describes leotards",
              "adjective",
              "noun",
              "linking verb",
              "predicate adjective",
              "comparison word",
              "demonstrative adjective, describes ones",
              "adjective",
              "adjective",
              "indefinite pronoun"
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 8,
    "subjects": {
      "grammar": {
        "name": "Grammar",
        "tag": "Grammar Town · Unit 1, Week 8",
        "tasks": [
          {
            "id": "agl8",
            "label": "Week 8 Lesson: Adverbs & Prepositions Refresher",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Week 8: Adverbs &amp; Prepositions Refresher</b></p><p style=\"opacity:.7;font-size:0.78rem;\">REVIEW / PRACTICE — Standards: L.4.1, L.4.1e (foundation)</p><p>Hi Adelyn!</p><p>Two quick refreshers in one week since you know both from Island.</p><p><b>ADVERBS</b> describe verbs (and sometimes adjectives or other adverbs), telling how, when, where, or how much — gracefully, yesterday, extremely.</p><p><b>PREPOSITIONS</b> show a relationship (location, direction, time) and always travel with a noun friend, forming a prepositional phrase — in the courtyard, during breakfast, beside the fountain.</p><p>This week's sentences mix both, plus a verbal or two from the last few weeks — see if your eye catches everything.</p><p style=\"opacity:.8;\"><b>Step 1:</b> For each sentence, label every word using only the 8 parts of speech, based on the job the word is doing in that sentence.</p><p style=\"opacity:.8;\"><b>Step 2:</b> After a sentence, a Verbal Check asks you to spot any gerunds, participles, and infinitives.</p></div>"
          },
          {
            "id": "ag8_1",
            "label": "Sentence 1",
            "type": "pos-tagger",
            "sentence": [
              "Adelyn",
              "gracefully",
              "cartwheeled",
              "across",
              "the",
              "courtyard",
              "yesterday,",
              "still",
              "smiling."
            ],
            "answers": [
              "Noun",
              "Adverb",
              "Verb",
              "Preposition",
              "Adjective",
              "Noun",
              "Adverb",
              "Adverb",
              "Adjective"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "proper noun",
              "adverb",
              "action verb",
              "preposition",
              "article",
              "noun",
              "adverb",
              "adverb",
              "adjective: it describes a noun here"
            ]
          },
          {
            "id": "ag8_1c",
            "label": "Sentence 1 — Verbal Check",
            "type": "concept-check",
            "sentence": [
              "Adelyn",
              "gracefully",
              "cartwheeled",
              "across",
              "the",
              "courtyard",
              "yesterday,",
              "still",
              "smiling."
            ],
            "prompt": "Now look at the words built from verbs in this sentence. For an infinitive, tap the verb that comes after \"to\". Is each one a regular verb, a gerund, a participle, or an infinitive?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Infinitive",
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 2,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 8,
                "answer": "Participle",
                "explanation": "built from a verb, but here it describes a noun, so it works as an adjective"
              }
            ]
          },
          {
            "id": "ag8_2",
            "label": "Sentence 2",
            "type": "pos-tagger",
            "sentence": [
              "During",
              "breakfast,",
              "Kenley",
              "quickly",
              "told",
              "an",
              "extremely",
              "silly",
              "joke",
              "about",
              "her",
              "cooking."
            ],
            "answers": [
              "Preposition",
              "Noun",
              "Noun",
              "Adverb",
              "Verb",
              "Adjective",
              "Adverb",
              "Adjective",
              "Noun",
              "Preposition",
              "Pronoun",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "preposition",
              "noun",
              "proper noun",
              "adverb",
              "action verb",
              "article",
              "adverb, describes silly",
              "adjective",
              "noun",
              "preposition",
              "possessive pronoun",
              "noun: it names the activity, so it does a noun's job here"
            ]
          },
          {
            "id": "ag8_2c",
            "label": "Sentence 2 — Verbal Check",
            "type": "concept-check",
            "sentence": [
              "During",
              "breakfast,",
              "Kenley",
              "quickly",
              "told",
              "an",
              "extremely",
              "silly",
              "joke",
              "about",
              "her",
              "cooking."
            ],
            "prompt": "Now look at the words built from verbs in this sentence. For an infinitive, tap the verb that comes after \"to\". Is each one a regular verb, a gerund, a participle, or an infinitive?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Infinitive",
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 11,
                "answer": "Gerund",
                "explanation": "an -ing verb form that names an activity, so it works as a noun"
              }
            ]
          },
          {
            "id": "ag8_3",
            "label": "Sentence 3",
            "type": "pos-tagger",
            "sentence": [
              "Adelyn",
              "quietly",
              "tiptoed",
              "around",
              "the",
              "sleeping",
              "kitten",
              "near",
              "the",
              "market",
              "stall."
            ],
            "answers": [
              "Noun",
              "Adverb",
              "Verb",
              "Preposition",
              "Adjective",
              "Adjective",
              "Noun",
              "Preposition",
              "Adjective",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "proper noun",
              "adverb",
              "action verb",
              "preposition",
              "article",
              "adjective: it describes a noun here",
              "noun",
              "preposition",
              "article",
              "noun used as an adjective",
              "noun"
            ]
          },
          {
            "id": "ag8_3c",
            "label": "Sentence 3 — Verbal Check",
            "type": "concept-check",
            "sentence": [
              "Adelyn",
              "quietly",
              "tiptoed",
              "around",
              "the",
              "sleeping",
              "kitten",
              "near",
              "the",
              "market",
              "stall."
            ],
            "prompt": "Now look at the words built from verbs in this sentence. For an infinitive, tap the verb that comes after \"to\". Is each one a regular verb, a gerund, a participle, or an infinitive?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Infinitive",
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 2,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 5,
                "answer": "Participle",
                "explanation": "built from a verb, but here it describes a noun, so it works as an adjective"
              }
            ]
          },
          {
            "id": "ag8_4",
            "label": "Sentence 4",
            "type": "pos-tagger",
            "sentence": [
              "Kenley",
              "happily",
              "worked",
              "inside",
              "the",
              "kitchen",
              "until",
              "dinner",
              "was",
              "finally",
              "ready."
            ],
            "answers": [
              "Noun",
              "Adverb",
              "Verb",
              "Preposition",
              "Adjective",
              "Noun",
              "Conjunction",
              "Noun",
              "Verb",
              "Adverb",
              "Adjective"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "proper noun",
              "adverb",
              "action verb",
              "preposition",
              "article",
              "noun",
              "starts a clause here (subordinating conjunction, Week 25 preview)",
              "noun",
              "linking verb",
              "adverb",
              "predicate adjective"
            ]
          }
        ]
      }
    }
  },
  {
    "week_number": 9,
    "subjects": {
      "grammar": {
        "name": "Grammar",
        "tag": "Grammar Town · Unit 1, Week 9",
        "tasks": [
          {
            "id": "agl9",
            "label": "Week 9 Lesson: Conjunctions & Interjections Refresher — Full Cumulative Review",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Week 9: Conjunctions &amp; Interjections Refresher — Full Cumulative Review</b></p><p style=\"opacity:.7;font-size:0.78rem;\">REVIEW / PRACTICE — Standards: L.4.1 (cumulative)</p><p>Hi Adelyn!</p><p>Last quick refresher: <b>CONJUNCTIONS</b> join things together (<b>FANBOYS</b>: for, and, nor, but, or, yet, so).</p><p><b>INTERJECTIONS</b> burst out with sudden feeling (Wow!</p><p>Yikes!).</p><p>Both straight from Island — you've got these.</p><p>Now for the big picture: you know all 8 original parts of speech, <b>PLUS</b> this unit's new depth — collective/compound nouns, indefinite pronouns, gerunds, participles, infinitives, and demonstrative/interrogative words.</p><p>That's a serious upgrade from Island!</p><p>This week, label every word with one of the 8 parts of speech. After a sentence that has a verb form doing a new job, a Verbal Check asks you to name it.</p><p style=\"opacity:.8;\"><b>Step 1:</b> For each sentence, label every word using only the 8 parts of speech, based on the job the word is doing in that sentence.</p><p style=\"opacity:.8;\"><b>Step 2:</b> After a sentence, a Verbal Check asks you to spot any gerunds, participles, and infinitives.</p></div>"
          },
          {
            "id": "ag9_1",
            "label": "Sentence 1",
            "type": "pos-tagger",
            "sentence": [
              "Wow,",
              "everyone",
              "watched",
              "Adelyn's",
              "stunning",
              "cartwheel,",
              "but",
              "nobody",
              "expected",
              "her",
              "giggling!"
            ],
            "answers": [
              "Interjection",
              "Pronoun",
              "Verb",
              "Noun",
              "Adjective",
              "Noun",
              "Conjunction",
              "Pronoun",
              "Verb",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "interjection",
              "indefinite pronoun",
              "action verb",
              "possessive proper noun",
              "adjective: it describes a noun here",
              "noun",
              "conjunction",
              "indefinite pronoun",
              "action verb",
              "possessive pronoun",
              "noun: it names the activity, so it does a noun's job here"
            ]
          },
          {
            "id": "ag9_1c",
            "label": "Sentence 1 — Verbal Check",
            "type": "concept-check",
            "sentence": [
              "Wow,",
              "everyone",
              "watched",
              "Adelyn's",
              "stunning",
              "cartwheel,",
              "but",
              "nobody",
              "expected",
              "her",
              "giggling!"
            ],
            "prompt": "Now look at the words built from verbs in this sentence. For an infinitive, tap the verb that comes after \"to\". Is each one a regular verb, a gerund, a participle, or an infinitive?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Infinitive",
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 2,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 4,
                "answer": "Participle",
                "explanation": "built from a verb, but here it describes a noun, so it works as an adjective"
              },
              {
                "index": 8,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 10,
                "answer": "Gerund",
                "explanation": "an -ing verb form that names an activity, so it works as a noun"
              }
            ]
          },
          {
            "id": "ag9_2",
            "label": "Sentence 2",
            "type": "pos-tagger",
            "sentence": [
              "This",
              "delicious",
              "soup,",
              "which",
              "Kenley",
              "made",
              "this",
              "morning,",
              "is",
              "perfect",
              "for",
              "sharing",
              "with",
              "the",
              "whole",
              "family."
            ],
            "answers": [
              "Adjective",
              "Adjective",
              "Noun",
              "Pronoun",
              "Noun",
              "Verb",
              "Adjective",
              "Noun",
              "Verb",
              "Adjective",
              "Preposition",
              "Noun",
              "Preposition",
              "Adjective",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "demonstrative adjective, describes soup",
              "adjective",
              "noun",
              "relative pronoun (preview of Week 27)",
              "proper noun",
              "action verb",
              "demonstrative adjective, describes morning",
              "noun",
              "linking verb",
              "predicate adjective",
              "preposition",
              "noun: it names the activity, so it does a noun's job here",
              "preposition",
              "article",
              "adjective",
              "collective noun"
            ]
          },
          {
            "id": "ag9_2c",
            "label": "Sentence 2 — Verbal Check",
            "type": "concept-check",
            "sentence": [
              "This",
              "delicious",
              "soup,",
              "which",
              "Kenley",
              "made",
              "this",
              "morning,",
              "is",
              "perfect",
              "for",
              "sharing",
              "with",
              "the",
              "whole",
              "family."
            ],
            "prompt": "Now look at the words built from verbs in this sentence. For an infinitive, tap the verb that comes after \"to\". Is each one a regular verb, a gerund, a participle, or an infinitive?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Infinitive",
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 11,
                "answer": "Gerund",
                "explanation": "an -ing verb form that names an activity, so it works as a noun"
              }
            ]
          },
          {
            "id": "ag9_3",
            "label": "Sentence 3",
            "type": "pos-tagger",
            "sentence": [
              "Hooray,",
              "Adelyn",
              "and",
              "Kenley",
              "both",
              "finished",
              "their",
              "routines,",
              "and",
              "everyone",
              "celebrated",
              "with",
              "mango",
              "smoothies!"
            ],
            "answers": [
              "Interjection",
              "Noun",
              "Conjunction",
              "Noun",
              "Pronoun",
              "Verb",
              "Pronoun",
              "Noun",
              "Conjunction",
              "Pronoun",
              "Verb",
              "Preposition",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "interjection",
              "proper noun",
              "conjunction",
              "proper noun",
              "indefinite pronoun, describes Adelyn and Kenley",
              "action verb",
              "possessive pronoun",
              "noun",
              "conjunction",
              "indefinite pronoun",
              "action verb",
              "preposition",
              "noun used as an adjective",
              "noun"
            ]
          },
          {
            "id": "ag9_4",
            "label": "Sentence 4",
            "type": "pos-tagger",
            "sentence": [
              "Ouch!",
              "Something",
              "pinched",
              "her",
              "toe,",
              "but",
              "nothing",
              "could",
              "stop",
              "her",
              "from",
              "finishing",
              "the",
              "cartwheel."
            ],
            "answers": [
              "Interjection",
              "Pronoun",
              "Verb",
              "Pronoun",
              "Noun",
              "Conjunction",
              "Pronoun",
              "Verb",
              "Verb",
              "Pronoun",
              "Preposition",
              "Noun",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "interjection",
              "indefinite pronoun",
              "action verb",
              "possessive pronoun",
              "noun",
              "conjunction",
              "indefinite pronoun",
              "helping verb",
              "action verb",
              "object pronoun",
              "preposition",
              "noun: it names the activity, so it does a noun's job here",
              "article",
              "noun"
            ]
          },
          {
            "id": "ag9_4c",
            "label": "Sentence 4 — Verbal Check",
            "type": "concept-check",
            "sentence": [
              "Ouch!",
              "Something",
              "pinched",
              "her",
              "toe,",
              "but",
              "nothing",
              "could",
              "stop",
              "her",
              "from",
              "finishing",
              "the",
              "cartwheel."
            ],
            "prompt": "Now look at the words built from verbs in this sentence. For an infinitive, tap the verb that comes after \"to\". Is each one a regular verb, a gerund, a participle, or an infinitive?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Infinitive",
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 2,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 11,
                "answer": "Gerund",
                "explanation": "an -ing verb form that names an activity, so it works as a noun"
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
      "grammar": {
        "name": "Grammar",
        "tag": "Grammar Town · Unit 1, Week 10",
        "tasks": [
          {
            "id": "agl10",
            "label": "Week 10 Lesson: Unit 1 Assessment — Full Grammar Safari",
            "type": "read",
            "content": "<div class=\"lesson-text\"><p><b>Week 10: Unit 1 Assessment — Full Grammar Safari</b></p><p style=\"opacity:.7;font-size:0.78rem;\">REVIEW / PRACTICE — Standards: L.4.1, L.4.6 (cumulative assessment)</p><p>Hi Adelyn!</p><p>Time for a full Grammar Safari!</p><p>Label every single word with one of the 8 parts of speech in two whole sentences, then name any gerunds, participles, and infinitives in the checks that follow, all by yourself, pulling together your Island knowledge plus everything new from this unit.</p><p>Don't stress about perfection — even professional grammarians debate a few tricky words.</p><p>The goal is showing me you understand the <b>JOB</b> each word is doing, at a real Town-level depth now.</p><p>When you finish, do a victory cartwheel — you've officially leveled up past Grammar Island.</p><p>Level 1 of the Four-Level Analysis, Town edition, is yours!</p><p style=\"opacity:.8;\"><b>Step 1:</b> For each sentence, label every word using only the 8 parts of speech, based on the job the word is doing in that sentence.</p><p style=\"opacity:.8;\"><b>Step 2:</b> After a sentence, a Verbal Check asks you to spot any gerunds, participles, and infinitives.</p></div>"
          },
          {
            "id": "ag10_1",
            "label": "Sentence 1",
            "type": "pos-tagger",
            "sentence": [
              "The",
              "tired",
              "but",
              "happy",
              "gymnast,",
              "smiling",
              "proudly,",
              "loved",
              "stretching",
              "near",
              "the",
              "warm,",
              "cozy",
              "fire."
            ],
            "answers": [
              "Adjective",
              "Adjective",
              "Conjunction",
              "Adjective",
              "Noun",
              "Adjective",
              "Adverb",
              "Verb",
              "Noun",
              "Preposition",
              "Adjective",
              "Adjective",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "article",
              "adjective",
              "conjunction",
              "adjective",
              "common noun",
              "adjective: it describes a noun here",
              "adverb",
              "action verb",
              "noun: it names the activity, so it does a noun's job here",
              "preposition",
              "article",
              "adjective",
              "adjective",
              "common noun"
            ]
          },
          {
            "id": "ag10_1c",
            "label": "Sentence 1 — Verbal Check",
            "type": "concept-check",
            "sentence": [
              "The",
              "tired",
              "but",
              "happy",
              "gymnast,",
              "smiling",
              "proudly,",
              "loved",
              "stretching",
              "near",
              "the",
              "warm,",
              "cozy",
              "fire."
            ],
            "prompt": "Now look at the words built from verbs in this sentence. For an infinitive, tap the verb that comes after \"to\". Is each one a regular verb, a gerund, a participle, or an infinitive?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Infinitive",
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 5,
                "answer": "Participle",
                "explanation": "built from a verb, but here it describes a noun, so it works as an adjective"
              },
              {
                "index": 7,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 8,
                "answer": "Gerund",
                "explanation": "an -ing verb form that names an activity, so it works as a noun"
              }
            ]
          },
          {
            "id": "ag10_2",
            "label": "Sentence 2",
            "type": "pos-tagger",
            "sentence": [
              "Everyone",
              "wanted",
              "to",
              "visit",
              "the",
              "delicious",
              "little",
              "bakery",
              "that",
              "Kenley",
              "found",
              "near",
              "our",
              "new",
              "apartment."
            ],
            "answers": [
              "Pronoun",
              "Verb",
              "Verb",
              "Verb",
              "Adjective",
              "Adjective",
              "Adjective",
              "Noun",
              "Pronoun",
              "Noun",
              "Verb",
              "Preposition",
              "Pronoun",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "indefinite pronoun",
              "action verb",
              "\"to\" and the verb after it work together as one verb form",
              "verb: it follows \"to\"",
              "article",
              "adjective",
              "adjective",
              "common noun",
              "relative pronoun (preview of Week 27)",
              "proper noun",
              "action verb",
              "preposition",
              "possessive pronoun",
              "adjective",
              "common noun"
            ]
          },
          {
            "id": "ag10_2c",
            "label": "Sentence 2 — Verbal Check",
            "type": "concept-check",
            "sentence": [
              "Everyone",
              "wanted",
              "to",
              "visit",
              "the",
              "delicious",
              "little",
              "bakery",
              "that",
              "Kenley",
              "found",
              "near",
              "our",
              "new",
              "apartment."
            ],
            "prompt": "Now look at the words built from verbs in this sentence. For an infinitive, tap the verb that comes after \"to\". Is each one a regular verb, a gerund, a participle, or an infinitive?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Infinitive",
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 1,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 3,
                "answer": "Infinitive",
                "explanation": "the verb that follows \"to\"; \"to\" and the verb work as a team"
              }
            ]
          },
          {
            "id": "ag10_3",
            "label": "Sentence 3",
            "type": "pos-tagger",
            "sentence": [
              "Wow,",
              "everyone",
              "loved",
              "watching",
              "Adelyn's",
              "stunning",
              "routine",
              "and",
              "Kenley's",
              "delicious",
              "cooking!"
            ],
            "answers": [
              "Interjection",
              "Pronoun",
              "Verb",
              "Noun",
              "Noun",
              "Adjective",
              "Noun",
              "Conjunction",
              "Noun",
              "Adjective",
              "Noun"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "interjection",
              "indefinite pronoun",
              "action verb",
              "noun: it names the activity, so it does a noun's job here",
              "possessive proper noun",
              "adjective: it describes a noun here",
              "noun",
              "conjunction",
              "possessive proper noun",
              "adjective",
              "noun: it names the activity, so it does a noun's job here"
            ]
          },
          {
            "id": "ag10_3c",
            "label": "Sentence 3 — Verbal Check",
            "type": "concept-check",
            "sentence": [
              "Wow,",
              "everyone",
              "loved",
              "watching",
              "Adelyn's",
              "stunning",
              "routine",
              "and",
              "Kenley's",
              "delicious",
              "cooking!"
            ],
            "prompt": "Now look at the words built from verbs in this sentence. For an infinitive, tap the verb that comes after \"to\". Is each one a regular verb, a gerund, a participle, or an infinitive?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Infinitive",
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 2,
                "answer": "Verb",
                "explanation": "it shows the action (or helps the action) in its clause, so it is a regular verb"
              },
              {
                "index": 3,
                "answer": "Gerund",
                "explanation": "an -ing verb form that names an activity, so it works as a noun"
              },
              {
                "index": 5,
                "answer": "Participle",
                "explanation": "built from a verb, but here it describes a noun, so it works as an adjective"
              },
              {
                "index": 10,
                "answer": "Gerund",
                "explanation": "an -ing verb form that names an activity, so it works as a noun"
              }
            ]
          },
          {
            "id": "ag10_4",
            "label": "Sentence 4",
            "type": "pos-tagger",
            "sentence": [
              "This",
              "tiny",
              "market,",
              "which",
              "sells",
              "fresh",
              "mangoes,",
              "is",
              "our",
              "favorite",
              "place",
              "to",
              "visit."
            ],
            "answers": [
              "Adjective",
              "Adjective",
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Noun",
              "Verb",
              "Pronoun",
              "Adjective",
              "Noun",
              "Verb",
              "Verb"
            ],
            "options": [
              "Noun",
              "Pronoun",
              "Verb",
              "Adjective",
              "Adverb",
              "Preposition",
              "Conjunction",
              "Interjection"
            ],
            "explanations": [
              "demonstrative adjective, describes market",
              "adjective",
              "noun",
              "relative pronoun (preview of Week 27)",
              "action verb",
              "adjective",
              "noun",
              "linking verb",
              "possessive pronoun",
              "adjective",
              "noun",
              "\"to\" and the verb after it work together as one verb form",
              "verb: it follows \"to\""
            ]
          },
          {
            "id": "ag10_4c",
            "label": "Sentence 4 — Verbal Check",
            "type": "concept-check",
            "sentence": [
              "This",
              "tiny",
              "market,",
              "which",
              "sells",
              "fresh",
              "mangoes,",
              "is",
              "our",
              "favorite",
              "place",
              "to",
              "visit."
            ],
            "prompt": "Now look at the words built from verbs in this sentence. For an infinitive, tap the verb that comes after \"to\". Is each one a regular verb, a gerund, a participle, or an infinitive?",
            "question": "What is \"{w}\" doing in this sentence?",
            "options": [
              "Infinitive",
              "Participle",
              "Gerund",
              "Verb"
            ],
            "targets": [
              {
                "index": 12,
                "answer": "Infinitive",
                "explanation": "the verb that follows \"to\"; \"to\" and the verb work as a team"
              }
            ]
          }
        ]
      }
    }
  }
];
