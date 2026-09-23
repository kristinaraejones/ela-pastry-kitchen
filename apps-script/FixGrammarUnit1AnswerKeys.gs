/**
 * One-time patch: fixes the two Grammar Town Unit 1 answer-key issues found
 * 2026-09-16 without needing to re-paste the whole AdelynGrammarUnit1.gs file:
 *  1) Sentence 1 ("Adelyn's gymnastics team...") tagged "gymnastics" as plain
 *     Noun instead of Adjective (noun-adjunct), inconsistent with every other
 *     sentence in the unit that uses this pattern (cooking class, yoga group, etc).
 *  2) The Noun/Pronoun/Verb Check tasks (Weeks 1-3) showed which answer was
 *     right on a miss but never explained why -- added a short explanation
 *     to each of their 43 targets.
 *
 * Run patchGrammarUnit1AnswerKeys() ONCE from the function dropdown. It only
 * touches Adelyn's already-seeded Schedule rows for the 13 task ids below --
 * it does not touch Submissions/scores, and is safe to re-run (idempotent,
 * since it just overwrites the same fields each time).
 */

function patchGrammarUnit1AnswerKeys() {
  var PATCH = {
  "ag1_1": {
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
  "ag1_1c": {
    "targets": [
      {
        "index": 0,
        "answer": "Proper",
        "explanation": "names one specific person, place, or thing, so it's capitalized"
      },
      {
        "index": 2,
        "answer": ["Common", "Collective"],
        "explanation": "it is BOTH: common (a general name, not one specific one) and collective (a group acting as one unit)"
      },
      {
        "index": 6,
        "answer": ["Common", "Compound"],
        "explanation": "it is BOTH: common (a general name, not one specific one) and compound (two words joined into one noun idea)"
      },
      {
        "index": 9,
        "answer": ["Common", "Compound"],
        "explanation": "it is BOTH: common (a general name, not one specific one) and compound (two words joined into one noun idea)"
      },
      {
        "index": 12,
        "answer": "Common",
        "explanation": "names any person, place, or thing — not one specific one"
      }
    ]
  },
  "ag1_2c": {
    "targets": [
      {
        "index": 2,
        "answer": ["Common", "Collective"],
        "explanation": "it is BOTH: common (a general name, not one specific one) and collective (a group acting as one unit)"
      },
      {
        "index": 5,
        "answer": ["Common", "Compound"],
        "explanation": "it is BOTH: common (a general name, not one specific one) and compound (two words joined into one noun idea)"
      },
      {
        "index": 7,
        "answer": "Proper",
        "explanation": "names one specific person, place, or thing, so it's capitalized"
      },
      {
        "index": 9,
        "answer": ["Common","Collective"],
        "explanation": "it is BOTH: common (a general name, not one specific one) and collective (a group acting as one unit)"
      }
    ]
  },
  "ag1_3c": {
    "targets": [
      {
        "index": 0,
        "answer": "Proper",
        "explanation": "names one specific person, place, or thing, so it's capitalized"
      },
      {
        "index": 2,
        "answer": ["Common","Collective"],
        "explanation": "it is BOTH: common (a general name, not one specific one) and collective (a group acting as one unit)"
      },
      {
        "index": 6,
        "answer": ["Common", "Compound"],
        "explanation": "it is BOTH: common (a general name, not one specific one) and compound (two words joined into one noun idea)"
      },
      {
        "index": 10,
        "answer": "Common",
        "explanation": "names any person, place, or thing — not one specific one"
      }
    ]
  },
  "ag1_4c": {
    "targets": [
      {
        "index": 2,
        "answer": ["Common", "Collective"],
        "explanation": "it is BOTH: common (a general name, not one specific one) and collective (a group acting as one unit)"
      },
      {
        "index": 6,
        "answer": ["Common", "Collective"],
        "explanation": "it is BOTH: common (a general name, not one specific one) and collective (a group acting as one unit)"
      },
      {
        "index": 11,
        "answer": "Common",
        "explanation": "names any person, place, or thing — not one specific one"
      }
    ]
  },
  "ag2_1c": {
    "targets": [
      {
        "index": 0,
        "answer": ["Subject","Indefinite"],
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
  "ag2_2c": {
    "targets": [
      {
        "index": 0,
        "answer": ["Subject","Indefinite"],
        "explanation": "it is BOTH: an indefinite pronoun (no specific person or thing) and the subject in its clause"
      },
      {
        "index": 2,
        "answer": "Possessive",
        "explanation": "it shows ownership"
      },
      {
        "index": 7,
        "answer": ["Subject","Indefinite"],
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
  "ag2_3c": {
    "targets": [
      {
        "index": 0,
        "answer": ["Subject","Indefinite"],
        "explanation": "it is BOTH: an indefinite pronoun (no specific person or thing) and the subject in its clause"
      },
      {
        "index": 4,
        "answer": "Possessive",
        "explanation": "it shows ownership"
      },
      {
        "index": 9,
        "answer": ["Subject","Indefinite"],
        "explanation": "it is BOTH: an indefinite pronoun (no specific person or thing) and the subject in its clause"
      },
      {
        "index": 15,
        "answer": "Subject",
        "explanation": "it's doing the action in its clause"
      }
    ]
  },
  "ag2_4c": {
    "targets": [
      {
        "index": 0,
        "answer": ["Subject","Indefinite"],
        "explanation": "it is BOTH: an indefinite pronoun (no specific person or thing) and the subject in its clause"
      },
      {
        "index": 2,
        "answer": ["Object","Indefinite"],
        "explanation": "it is BOTH: an indefinite pronoun (no specific person or thing) and the object in its clause"
      },
      {
        "index": 8,
        "answer": ["Subject","Indefinite"],
        "explanation": "it is BOTH: an indefinite pronoun (no specific person or thing) and the subject in its clause"
      }
    ]
  },
  "ag3_1c": {
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
  "ag3_2c": {
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
  "ag3_3c": {
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
  "ag3_4c": {
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
};

  var sh = getSheet_("Schedule");
  var headers = SHEET_HEADERS.Schedule;
  var studentCol = headers.indexOf("student");
  var taskIdCol = headers.indexOf("task_id");
  var contentCol = headers.indexOf("content_json");
  var lastRow = sh.getLastRow();
  var range = sh.getRange(2, 1, lastRow - 1, headers.length);
  var values = range.getValues();
  var patched = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (row[studentCol] !== "adelyn") continue;
    var taskId = row[taskIdCol];
    if (!PATCH.hasOwnProperty(taskId)) continue;
    var content = JSON.parse(row[contentCol]);
    var fix = PATCH[taskId];
    Object.keys(fix).forEach(function (k) { content[k] = fix[k]; });
    values[i][contentCol] = JSON.stringify(content);
    patched++;
  }

  range.setValues(values);
  var msg = "Patched " + patched + " of " + Object.keys(PATCH).length + " Grammar Unit 1 rows.";
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}