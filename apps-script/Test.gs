function patchWeek1SampleAnswers_() {
  var patches = {
    'r2': "The author probably delayed the reveal so readers would form their own theories and pay closer attention, the same way the kids in the story are confused. A strong answer names a specific effect it had on her as a reader — for example, that not knowing made her pay closer attention to small details, or made her want to keep reading to see if her guess was right — rather than just restating that it 'built suspense' without saying how.",
    'r3': "A strong answer names something specific Reynie can only guess at because he's not inside another character's head — for example, what the adults running the test are actually planning, or what another kid is really thinking about him — and explains that we'd only know for sure if the book switched to that character's POV.",
    'w6': "This is a personal-topic prompt, so there's no single correct paragraph — grade the structure, not the opinion. Example of the shape to look for: Claim — 'Students should get a short walk break between long stretches of schoolwork.' Evidence — a specific, checkable fact (e.g. a real study on attention and breaks), not just a feeling. Explain — connects that evidence back to the claim (why the fact actually supports it). Specificity test — the evidence should be tied closely enough to the claim that it couldn't be dropped into a totally different essay. Fact-vs-opinion check — the evidence itself should be verifiable, not just a restated opinion."
  };

  var sh = getSheet_('Schedule');
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var taskIdCol = headers.indexOf('task_id');
  var contentCol = headers.indexOf('content_json');
  var studentCol = headers.indexOf('student');

  var patched = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[studentCol] !== 'kenley') continue;
    var taskId = row[taskIdCol];
    if (!patches.hasOwnProperty(taskId)) continue;
    var content = JSON.parse(row[contentCol]);
    content.sampleAnswer = patches[taskId];
    sh.getRange(i + 1, contentCol + 1).setValue(JSON.stringify(content));
    patched.push(taskId);
  }

  var msg = 'Patched sampleAnswer into rows: ' + patched.join(', ');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

function runPatch() {
  patchWeek1SampleAnswers_();
}
