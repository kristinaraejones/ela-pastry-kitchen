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
- **Settings** — the shared weekly due date, current week number, and the parent's Term Final / Monthly Test overrides. The mockup's "demo controls" became real persisted settings here, per the spec's own note that they "should be a real per-family setting, not just a demo toggle, once built for real."

## Task types

Ported 1:1 from the mockup: `read`, `external`, `reflection` (with the Approve/Refire review loop), `graded-mc`, `graded-dictation` (word/sentence, word-by-word spelling + separate grammar checklist), `pos-tagger` (Level 1), and `phrase-tagger` (Levels 2–4, driven by `GRAMMAR_LEVEL_OPTIONS`-shaped `options` arrays).

## Status system

**Served ✓** / **Burning 🔥** / **Burnt ⚫** — computed live from completion + score + the shared due date, never a manually-set flag. See the build spec for the exact rules.

## What's real content vs. placeholder right now

- **Kenley** — Week 1 / Month 1 real content for all five subjects (Vocab Set A, AAS Step 7, Month 1 four-level grammar analysis, MBS Pause Point 1, Writing Block 1).
- **Adelyn** — placeholder shell only, same 5-subject structure, awaiting her curriculum docs and her separate vocab matching-game component.

## Known simplifications from the spec (intentional, documented here so they're not mistaken for bugs)

- In-progress answers (an unsubmitted MC selection, a half-typed grammar tagging pass) aren't synced to the Sheet — only submitted/graded results are. Syncing every keystroke would mean a network write per tap; this matches how often the original mockup's own state actually changed meaningfully.
- Submissions are upserted (latest state per task), not append-only history — see the Submissions bullet above for why that's still faithful to the audit-trail intent.
- The due date, current week, and test-lock overrides are one shared, global setting — not per-student — matching the spec's own "shared weekly due date" language.

## Local development

Everything is static files; open `index.html` via GitHub Pages or any local static server. There's no build step and no dependencies.
