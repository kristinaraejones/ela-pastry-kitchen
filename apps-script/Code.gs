/**
 * The ELA Pastry Kitchen — Apps Script API
 *
 * Fronts a Google Sheet acting as the database for the ELA Pastry Kitchen app.
 * Deploy this project as a Web App (Execute as: Me, Who has access: Anyone)
 * and paste the resulting URL into the frontend's setup screen.
 *
 * Sheets (tabs), created/seeded by Setup.gs -> setupSheets():
 *   Schedule          | student | subject_key | subject_name | subject_tag | week_number | task_id | task_type | label | content_json | dynamic_bank_key | term_final | monthly_test
 *   Submissions       | student | task_id | timestamp | status | score | answers_json | parent_comment
 *   ReviewPool        | student | word | times_missed | times_correct | last_seen | status | context_sentence
 *   MonthTestMarkers  | student | subject_key | bank_position
 *   BurnLog           | student | station | tag | date | reason | items_json
 *   Banks             | student | subject_key | items_json
 *   Settings          | key | value
 *
 * Every read/write below is scoped by a `student` field. Nothing here queries
 * across students — that guarantee lives in every function that takes a
 * `student` argument and filters by it before returning or writing anything.
 */

function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  var sh = getSS_().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name + '. Run setupSheets() first.');
  return sh;
}

/** Reads a sheet into an array of plain objects keyed by its header row. */
function sheetToObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (row.join('') === '') continue; // skip fully blank rows
    var obj = {};
    for (var c = 0; c < headers.length; c++) obj[headers[c]] = row[c];
    obj._row = r + 1; // 1-indexed sheet row, for in-place updates
    out.push(obj);
  }
  return out;
}

function appendRowObject_(sheet, headers, obj) {
  var row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sheet.appendRow(row);
}

function getHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/** Finds a row matching all keyVals; updates it in place, or appends a new row. */
function upsertRow_(sheet, keyCols, obj) {
  var headers = getHeaders_(sheet);
  var rows = sheetToObjects_(sheet);
  var match = rows.find(function (r) {
    return keyCols.every(function (k) { return String(r[k]) === String(obj[k]); });
  });
  var rowValues = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  if (match) {
    sheet.getRange(match._row, 1, 1, headers.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
}

function safeParse_(json, fallback) {
  if (json === '' || json === undefined || json === null) return fallback;
  try { return JSON.parse(json); } catch (e) { return fallback; }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Read helpers, each scoped by student ----------

function readSchedule_(student) {
  return sheetToObjects_(getSheet_('Schedule'))
    .filter(function (r) { return r.student === student; })
    .map(function (r) {
      return {
        subject_key: r.subject_key,
        subject_name: r.subject_name,
        subject_tag: r.subject_tag,
        week_number: r.week_number,
        id: r.task_id,
        type: r.task_type,
        label: r.label,
        dynamic: r.dynamic_bank_key || undefined,
        termFinal: !!r.term_final,
        monthlyTest: !!r.monthly_test,
        content: safeParse_(r.content_json, {})
      };
    });
}

function readSubmissions_(student) {
  return sheetToObjects_(getSheet_('Submissions'))
    .filter(function (r) { return r.student === student; })
    .map(function (r) {
      return {
        task_id: r.task_id,
        timestamp: r.timestamp,
        status: r.status,
        score: r.score,
        parent_comment: r.parent_comment,
        answers: safeParse_(r.answers_json, {})
      };
    });
}

function readReviewPool_(student) {
  return sheetToObjects_(getSheet_('ReviewPool'))
    .filter(function (r) { return r.student === student; })
    .map(function (r) {
      return {
        word: r.word,
        timesMissed: Number(r.times_missed) || 0,
        timesCorrect: Number(r.times_correct) || 0,
        lastSeen: r.last_seen,
        status: r.status,
        context: r.context_sentence || null
      };
    });
}

function readMarkers_(student) {
  var out = {};
  sheetToObjects_(getSheet_('MonthTestMarkers'))
    .filter(function (r) { return r.student === student; })
    .forEach(function (r) { out[r.subject_key] = Number(r.bank_position) || 0; });
  return out;
}

function readBurnLog_(student) {
  return sheetToObjects_(getSheet_('BurnLog'))
    .filter(function (r) { return r.student === student; })
    .map(function (r) {
      return {
        station: r.station,
        tag: r.tag,
        date: r.date,
        reason: r.reason,
        items: safeParse_(r.items_json, [])
      };
    })
    .reverse(); // most recent first, matching the mockup's unshift-based log
}

function readBanks_(student) {
  var out = {};
  sheetToObjects_(getSheet_('Banks'))
    .filter(function (r) { return r.student === student; })
    .forEach(function (r) { out[r.subject_key] = safeParse_(r.items_json, []); });
  return out;
}

function readSettings_() {
  var out = {};
  sheetToObjects_(getSheet_('Settings')).forEach(function (r) { out[r.key] = r.value; });
  return out;
}

// ---------- Entry points ----------

function doGet(e) {
  var action = e.parameter.action;
  var student = e.parameter.student;
  try {
    if (action === 'bootstrap') {
      if (!student) throw new Error('Missing student');
      return jsonOut_({
        ok: true,
        schedule: readSchedule_(student),
        submissions: readSubmissions_(student),
        reviewPool: readReviewPool_(student),
        markers: readMarkers_(student),
        burnLog: readBurnLog_(student),
        banks: readBanks_(student),
        settings: readSettings_()
      });
    }
    throw new Error('Unknown action: ' + action);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = safeParse_(e.postData.contents, {});
    var action = body.action;
    var student = body.student;

    if (action === 'saveSetting') {
      upsertRow_(getSheet_('Settings'), ['key'], { key: body.key, value: body.value });
      return jsonOut_({ ok: true });
    }

    if (!student) throw new Error('Missing student');

    if (action === 'saveSubmission') {
      upsertRow_(getSheet_('Submissions'), ['student', 'task_id'], {
        student: student,
        task_id: body.task_id,
        timestamp: new Date().toISOString(),
        status: body.status || '',
        score: body.score || '',
        answers_json: JSON.stringify(body.answers || {}),
        parent_comment: body.parent_comment || ''
      });
      return jsonOut_({ ok: true });
    }

    if (action === 'saveReviewWord') {
      upsertRow_(getSheet_('ReviewPool'), ['student', 'word'], {
        student: student,
        word: body.word,
        times_missed: body.timesMissed,
        times_correct: body.timesCorrect,
        last_seen: body.lastSeen,
        status: body.status,
        context_sentence: body.context || ''
      });
      return jsonOut_({ ok: true });
    }

    if (action === 'saveMarker') {
      upsertRow_(getSheet_('MonthTestMarkers'), ['student', 'subject_key'], {
        student: student,
        subject_key: body.subject_key,
        bank_position: body.bank_position
      });
      return jsonOut_({ ok: true });
    }

    if (action === 'addBurnLog') {
      appendRowObject_(getSheet_('BurnLog'), getHeaders_(getSheet_('BurnLog')), {
        student: student,
        station: body.station,
        tag: body.tag,
        date: body.date,
        reason: body.reason,
        items_json: JSON.stringify(body.items || [])
      });
      return jsonOut_({ ok: true });
    }

    throw new Error('Unknown action: ' + action);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
