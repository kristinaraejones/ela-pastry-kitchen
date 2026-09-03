# Deploying The ELA Pastry Kitchen

Two one-time steps: stand up the Google Sheet + Apps Script backend, then turn on GitHub Pages for the frontend. Both need to happen from your own Google/GitHub accounts — I can't do either of these for you.

## 1. Backend: Google Sheet + Apps Script

1. Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet. Name it whatever you like, e.g. "ELA Pastry Kitchen DB".
2. In the Sheet, go to **Extensions → Apps Script**. This opens a new Apps Script project already bound to your Sheet.
3. Delete the placeholder `Code.gs` content, then copy the full contents of this repo's [`apps-script/Code.gs`](apps-script/Code.gs) and paste it in.
4. Click the **+** next to "Files" and add a new script file named `Setup` (Apps Script will name the file `Setup.gs`). Copy the full contents of this repo's [`apps-script/Setup.gs`](apps-script/Setup.gs) and paste it in.
5. Save the project (Ctrl/Cmd+S).
6. In the function dropdown at the top (next to Debug/Run), select **setupSheets**, then click **Run**.
   - The first time you run it, Google will ask you to authorize the script (it needs permission to edit the Sheet it's bound to). Click through the consent screen — you'll likely see an "unverified app" warning since this is your own personal script; click **Advanced → Go to (project name) (unsafe)** to proceed. This is expected for any script you write yourself and haven't submitted to Google for review.
   - When it finishes, you'll get a popup confirming setup is complete. Check your Sheet — it should now have tabs: Schedule, Submissions, ReviewPool, MonthTestMarkers, BurnLog, Banks, Settings, with Schedule/Banks/ReviewPool/Settings already filled in.
7. Now deploy it as a web app: **Deploy → New deployment**.
   - Click the gear icon next to "Select type" and choose **Web app**.
   - Description: anything, e.g. "ELA Kitchen API v1".
   - **Execute as: Me** (your account).
   - **Who has access: Anyone**. (This does not mean anyone can find it — it's an unguessable, unlisted URL. It just means Google won't additionally require the caller to be logged into a Google account, which the kids' browsers aren't.)
   - Click **Deploy**, then authorize again if asked.
   - Copy the **Web app URL** it gives you (looks like `https://script.google.com/macros/s/AKfycb.../exec`). That's the one and only credential this whole backend has — treat it like you would any unlisted link.

If you ever change `Code.gs` or `Setup.gs` later, you'll need to create a **new deployment version** (Deploy → Manage deployments → edit → New version) for the changes to take effect — just saving the script isn't enough.

## 1a. Already deployed before? Run the one-time pacing migration

If your Sheet was set up before the due-date system was removed, its Settings tab still has old `dueDate` / `currentWeekNumber` rows. After pasting the updated `Code.gs` and `Setup.gs` and creating a new deployment version (see above), open the Apps Script editor, select **migrateToPerStudentWeeks** in the function dropdown, and click **Run** once. It removes the old rows and adds `kenley_current_week` / `adelyn_current_week`, both starting at week 1. Safe to run more than once.

## 1b. Loading Kenley's Month 1, Weeks 2–5 content

This is separate from `setupSheets()` and only needs to run once, after `setupSheets()` has already populated your Sheet:

1. In the Apps Script editor, click the **+** next to "Files" and add a new script file named `Month1Weeks2to5`. Copy the full contents of this repo's [`apps-script/Month1Weeks2to5.gs`](apps-script/Month1Weeks2to5.gs) and paste it in. Save.
2. Select **seedMonth1Weeks2to5_** in the function dropdown and click **Run**.
3. You'll get a popup confirming how many rows were added. It's idempotent — running it again does nothing if it detects the content is already there (no duplicate rows).

No new deployment version is needed for this one — it's a one-time data-loading function, not part of the live API surface (`Code.gs`), so it doesn't affect what the deployed Web App serves at all, only what's *in* the Sheet for it to read.

## 1c. Loading Adelyn's Vocabulary & Reading content (all 36 weeks)

Same idea as 1b, but for Adelyn — her full year of Vocabulary and Reading, converted from her own course documents. Only needs to run once, after `setupSheets()` has already populated your Sheet:

1. In the Apps Script editor, click the **+** next to "Files" and add a new script file named `AdelynVocabReading`. Copy the full contents of this repo's [`apps-script/AdelynVocabReading.gs`](apps-script/AdelynVocabReading.gs) and paste it in. Save.
2. Select **seedAdelynVocabReading_** in the function dropdown and click **Run**.
3. You'll get a popup confirming how many rows were added. It's idempotent — running it again does nothing if it detects the content is already there (no duplicate rows). It also automatically removes Adelyn's old "waiting on curriculum" placeholder rows for Vocabulary and Reading, so she won't see both the placeholder and the real content at once.

No new deployment version is needed for this one either — it's a one-time data-loading function, same as 1b.

## 1d. Loading Adelyn's Spelling content (all 27 weeks)

Same idea again, for Adelyn's Spelling — All About Spelling Level 5, Steps 2–28, converted from her own course PDF plus her checkpoint-quiz/final-exam document. Only needs to run once, after `setupSheets()` has already populated your Sheet:

1. In the Apps Script editor, click the **+** next to "Files" and add a new script file named `AdelynSpelling`. Copy the full contents of this repo's [`apps-script/AdelynSpelling.gs`](apps-script/AdelynSpelling.gs) and paste it in. Save.
2. Select **seedAdelynSpelling_** in the function dropdown and click **Run**.
3. You'll get a popup confirming how many rows were added. It's idempotent — running it again does nothing if it detects the content is already there (no duplicate rows). It also automatically removes Adelyn's old "waiting on curriculum" placeholder row for Spelling.

No new deployment version is needed for this one either — it's a one-time data-loading function, same as 1b/1c.

## 1e. Loading Adelyn's Grammar content (Unit 1 only — Weeks 1-10)

Adelyn's Grammar Town runs 36 weeks across 5 units with an evolving four-level sentence-analysis system (parts of speech, then sentence parts, then phrases, then clauses). Only Unit 1 (the eight parts of speech, Weeks 1-10) has been converted so far — see the build notes at the top of `AdelynGrammarUnit1.gs` for why the rest is a deliberate follow-up rather than an oversight. Only needs to run once, after `setupSheets()` has already populated your Sheet:

1. In the Apps Script editor, click the **+** next to "Files" and add a new script file named `AdelynGrammarUnit1`. Copy the full contents of this repo's [`apps-script/AdelynGrammarUnit1.gs`](apps-script/AdelynGrammarUnit1.gs) and paste it in. Save.
2. Select **seedAdelynGrammarUnit1_** in the function dropdown and click **Run**.
3. You'll get a popup confirming how many rows were added. It's idempotent — running it again does nothing if it detects the content is already there (no duplicate rows). It also automatically removes Adelyn's old "waiting on curriculum" placeholder row for Grammar.

No new deployment version is needed for this one either — it's a one-time data-loading function, same as 1b/1c/1d.

## 1f. Loading Adelyn's Writing content (Weeks 1-4 — "Sentence Play")

Adelyn's Writing Workshop is fundamentally a parent-led, mostly-oral program; the source document itself only has full weekly sessions written for Weeks 1-4 (Weeks 5-36 are previewed or not yet written, deliberately — see the build notes at the top of `AdelynWriting.gs`). Each week has a lesson (the parent-facing session script, run mostly away from the screen) and one independent-writing task Adelyn actually types into the app. There's no answer key by design — the program's philosophy is that there's no single correct sentence — so the parent-review flow instead shows that week's coaching notes (what to praise, what not to correct). Only needs to run once, after `setupSheets()` has already populated your Sheet:

1. In the Apps Script editor, click the **+** next to "Files" and add a new script file named `AdelynWriting`. Copy the full contents of this repo's [`apps-script/AdelynWriting.gs`](apps-script/AdelynWriting.gs) and paste it in. Save.
2. Select **seedAdelynWriting_** in the function dropdown and click **Run**.
3. You'll get a popup confirming how many rows were added. It's idempotent — running it again does nothing if it detects the content is already there (no duplicate rows). It also automatically removes Adelyn's old "waiting on curriculum" placeholder row for Writing.

No new deployment version is needed for this one either — it's a one-time data-loading function, same as 1b/1c/1d/1e.

## 2. Frontend: GitHub Pages

1. On GitHub, go to this repo's **Settings → Pages**.
2. Under "Build and deployment", set **Source: Deploy from a branch**, branch **main**, folder **/ (root)**.
3. Save. GitHub will give you a URL like `https://kristinaraejones.github.io/ela-pastry-kitchen/` within a minute or two.

## 3. Connect them

Open the Pages URL. The app will show a "Connect the Kitchen" screen asking for the Apps Script Web App URL from step 1. Paste it in and click Connect — it's saved in that browser's storage from then on, so you only do this once per device.

To skip that step entirely on every future device (recommended once you're happy with it), open `config.js` in the repo, set:

```js
const DEFAULT_API_URL = "https://script.google.com/macros/s/AKfycb.../exec";
```

and push that change. Every device that loads the page will then connect automatically with no setup screen.

## Verifying it's working

1. Open the app as Kenley, complete a "read" task (click "Mark as read").
2. Flip the toggle to Parent view — you should see the station's progress update.
3. Open the same Pages URL on a second device or browser tab — the completed task should show as done there too, confirming the Sheet is genuinely the shared source of truth and not per-device `localStorage` state.

## Adding Adelyn's real content later

Once her AAS Step 11 PDF and curriculum docs are ready, the cleanest path is: tell me (or whoever's driving Claude Code next) to convert them into the same `content_json` shape used for Kenley's tasks, and add rows to the `Schedule` tab (student = `adelyn`) the same way `Setup.gs` seeded Kenley's. No frontend code changes needed — the app renders whatever's in the Schedule sheet for that student.
