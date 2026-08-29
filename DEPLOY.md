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
