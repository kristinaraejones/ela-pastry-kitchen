# The ELA Pastry Kitchen

A shared, kitchen/bakery-themed daily-work app for Kenley (grade 7) and Adelyn (grade 4) — ELA only (Vocabulary, Spelling, Grammar, Reading, Writing). Homeschool Planet stays the system of record for everything else.

Built from `ELA_Pastry_Kitchen_Build_Spec.md`, using the validated `Kenley_Week1_Dashboard_Mockup.html` prototype as the reference implementation for every mechanic (status system, Refire loop, dictation grading, four-level grammar analysis, monthly/term test sampling, etc.).

## Architecture

- **Frontend** — static HTML/CSS/JS (`index.html`, `app.js`, `styles.css`), hosted on GitHub Pages. No build step. Kids use it via "Add to Home Screen" in iPad Safari.
- **Backend** — a Google Sheet as the database, fronted by a Google Apps Script Web App (`apps-script/Code.gs`, `apps-script/Setup.gs`) acting as a small REST-ish API. This is what replaced the mockup's `localStorage` — the Sheet is now the single source of truth, synced across every device.
- **No real login.** A profile picker (Kenley/Adelyn) plus a Kid/Parent view toggle, exactly like the mockup. This is a household tool for 2–4 people, not a multi-tenant product.

Every read and write is scoped by a `student` field. Nothing ever queries across Kenley's and Adelyn's data except the parent's own view of whichever child is currently selected.

## One-time setup (you'll need to do this — see [DEPLOY.md](DEPLOY.md))

1. Create the Google Sheet + Apps Script backend and deploy it as a Web App. Full steps in `DEPLOY.md`.
2. Turn on GitHub Pages for this repo (Settings → Pages → Deploy from branch → `main` → `/ (root)`).
3. Open the Pages URL on each device (or your laptop first) and paste the Apps Script Web App URL into the "Connect the Kitchen" screen. It's remembered per-device from then on. If you'd rather every device pick it up automatically with zero setup, put the URL in `config.js`'s `DEFAULT_API_URL` and push that instead.

I can't do steps 1–2 for you — they require being logged into your own Google and GitHub accounts.

## Data model (Google Sheet tabs)

- **Schedule** — one row per task, per child. `content_json` holds whatever that task type needs, mirroring the mockup's task objects.
- **Submissions** — the current state of each task per student (status, score, free-text/graded answers as JSON, parent comment). Upserted per task rather than kept as unbounded history — the audit trail that matters (past scores, missed words) lives in `BurnLog` and `ReviewPool` instead.
- **ReviewPool** — missed-word bank per student, feeds the auto-generated dictation review task.
- **MonthTestMarkers** — how far into each subject's month-wide bank a student has already been tested, so monthly/term tests sample what's new.
- **BurnLog** — history of redone ("burnt") sections.
- **Banks** *(addition beyond the literal spec text, needed to make "sample since last test" actually work)* — each subject's growing month/term word or question bank per student.
- **Settings** — a `key | value` table holding each student's `current_week` pointer (`kenley_current_week`, `adelyn_current_week`) plus the parent's global Term Final / Monthly Test overrides. There are no dates anywhere in this sheet.

## Task types

Ported 1:1 from the mockup: `read`, `external`, `reflection` (with the Approve/Refire review loop), `graded-mc`, `graded-dictation` (word/sentence, word-by-word spelling + separate grammar checklist), `pos-tagger` (Level 1), and `phrase-tagger` (Levels 2–4, driven by `GRAMMAR_LEVEL_OPTIONS`-shaped `options` arrays).

## Pacing: completion-based, not calendar-based

There are no due dates anywhere in this app — Homeschool Planet remains the system of record for actual calendar pacing. Instead, each student has their own persistent `current_week` pointer in the Settings sheet. Only that week's `Schedule` rows (plus the "unlocked whenever" monthly/term-test tasks, which aren't tied to any single week) count as her active work — past and future weeks' content simply isn't part of the active set, it's not "locked."

Advancing is her own deliberate click, not the parent's and not automatic: once every station for her current week reaches **Served**, an "Advance to Week N+1" button appears in her own view. Clicking it increments her `current_week` and nothing else — a week can otherwise sit half-finished indefinitely. A station with no content loaded yet for the active week shows as **empty** (not Served, not counted toward "plates served," and not advance-eligible), specifically so advancing can never accidentally chain through weeks that haven't been authored yet.

## Status system

**Served ✓** / **Burning 🔥** — computed live from completion + score, never a manually-set flag. There is no "Burnt" status — an earlier version tied a third status to due dates; that's been removed by deliberate decision. A station that isn't done yet just isn't done yet, with no calendar pressure or auto-flagging. See the build spec for the exact rules.

## Parent-view week navigator

A `‹ Week N ›` browser, visible only in Parent view, independent of whatever week the student herself is currently on. It covers three states:

- **Current** (`Week N === current_week`) — identical to the regular live Parent dashboard: interactive station cards, Approve/Refire, everything as usual.
- **Past** (`Week N < current_week`) — a read-only record: per subject, per task, its completion status, score if graded, and (for reflections) the full submitted text plus any parent comment from the Refire loop. No Refire/Approve/redo controls here — correcting a past week is a deliberate separate action, not something browsing back should make easy to do by accident.
- **Upcoming** (`Week N > current_week`, up to however far content has actually been authored) — a preview of the lesson content and task list with no submission data, since nothing's been done yet. A week with literally nothing authored shows "nothing planned yet" rather than an error.

Monthly/term-test tasks are deliberately left out of every week-scoped view (past, upcoming, and PDF export alike) — they're "unlocked whenever," not tied to one week, and `Submissions` only ever keeps a task's latest attempt (no per-attempt history). Attributing a since-retaken test's current score to whichever week it happened to be filed under would misrepresent that week's actual record. Their live status is unaffected and still shows normally on the current-week dashboard.

## PDF export

A parent-facing "Export weeks — to — → Export PDF" control, for the currently-selected child (the existing Kenley/Adelyn switcher is the student selector; there's no separate one). Generates a PDF client-side with [jsPDF](https://github.com/parallax/jsPDF) (loaded from a CDN — no server round-trip, no extra Google Drive/Docs permission scopes), organized by week then by subject: task label, type, completion status, score if graded, and the full text of any reflection submission plus its parent comment if it went through Refire. Dictation/tagging exercises show just the final score, not a blow-by-blow of every word or tag.

## What's real content vs. placeholder right now

- **Kenley** — real content for all of Month 1, Weeks 1–5, across all five subjects (Vocab context-clue words, AAS Steps 7–11, four weeks of four-level grammar analysis on original sentences, MBS Chapters 1–39 — the whole book — with Plot Structure & POV discussion questions, and the full Writing Block 1 arc from taught lesson through capstone reflection). Weeks 2–5 were converted from `Month1_Weeks2-5_Grammar_Spelling_Writing.md` / `Month1_Weeks2-5_Reading_Vocab.md` by `apps-script/Month1Weeks2to5.gs` (see that file's header comment for exactly which bits are Claude's own drafted content — the four spelling rule explanations and Writing Week 2's practice-set examples — versus content pulled directly from the source docs).
- **Adelyn** — placeholder shell only, same 5-subject structure, awaiting her curriculum docs and her separate vocab matching-game component.

## Known simplifications from the spec (intentional, documented here so they're not mistaken for bugs)

- In-progress answers (an unsubmitted MC selection, a half-typed grammar tagging pass) aren't synced to the Sheet — only submitted/graded results are. Syncing every keystroke would mean a network write per tap; this matches how often the original mockup's own state actually changed meaningfully.
- Submissions are upserted (latest state per task), not append-only history — see the Submissions bullet above for why that's still faithful to the audit-trail intent.
- Term Final / Monthly Test overrides are one shared, global setting — not per-student — since they're the parent's own toggles, not something each kid paces independently. `current_week` itself, by contrast, is per-student.
- A reflection sent back for revision in a week the student has since advanced past won't reappear in her own view (which only ever shows her *current* week) — it still shows up in the parent's "Needs Your Eyes" queue regardless of week, but there's currently no way for her to reopen and resubmit it herself once she's moved on. The parent's past-week report is read-only by design (no Refire there either) — this is a known gap, not something either built feature currently solves.

## Local development

Everything is static files; open `index.html` via GitHub Pages or any local static server. There's no build step and no dependencies.
