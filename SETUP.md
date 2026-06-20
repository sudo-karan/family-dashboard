# Setup — one-time, about 15 minutes

You will end up with:

- a **Google Sheet** that holds all the FDs (the family can edit it directly, too),
- a tiny **Apps Script web app** bound to that sheet (the backend — free, never sleeps),
- **`index.html`** hosted on GitHub Pages (the app — free, never sleeps),
- one **shared family password**, checked by the backend on every request.

---

## 1. Create the Google Sheet

1. Go to [sheets.new](https://sheets.new) and create a blank spreadsheet. Name it e.g. **Fixed Deposits**.
2. Rename the first tab to exactly **`FDs`** and paste this header row into row 1:

   ```
   ID	Account	Status	FD Account Number	Start Date	End Date	Tenure (Days)	Interest %	Principal Amount	Maturity Amount	Interest On Maturity	Compounding	Show In Dashboard	Remarks
   ```

   (Copy the line above and paste into cell A1 — it is tab-separated, so each
   heading lands in its own column.)

3. Add a second tab named exactly **`Accounts`** with this header row:

   ```
   Name	Person	Bank
   ```

4. **Format three columns as plain text** so Sheets never mangles them:
   select the **FD Account Number**, **Start Date** and **End Date** columns
   (click the column letters), then **Format → Number → Plain text**.
   This keeps huge FD numbers out of scientific notation, preserves leading
   zeros, and keeps dates as literal `YYYY-MM-DD` strings.

> The column names mirror the old workbook on purpose (`Status`,
> `FD Account Number`, `Start Date`, … in the same order). The backend matches
> columns **by header name, not position**, so you can reorder columns or add
> your own extra columns later — nothing breaks. It also understands the
> camelCase names (`startDate`, `rate`, …) if you ever prefer those.

## 2. Add the backend script

1. In the sheet: **Extensions → Apps Script**. Delete the placeholder code and
   paste the entire contents of **`Code.gs`** from this repo. Save (💾).
2. Set the password: **Project Settings (⚙ in the left sidebar) → Script
   Properties → Add script property**:
   - Property: `FAMILY_PASSWORD`
   - Value: *the password your family will share*

   The password lives only here — never in the code, never in the web page.

## 3. Deploy it as a web app

1. **Deploy → New deployment**.
2. Click the ⚙ next to "Select type" → **Web app**.
3. Fill in:
   - Description: anything
   - **Execute as: Me**
   - **Who has access: Anyone** *(sometimes worded "Anyone with the link")*
4. Click **Deploy**, authorize the permissions prompt (it asks because the
   script reads/writes your own spreadsheet), and copy the **Web app URL** —
   it ends in `/exec`.

**Smoke test:** open that `/exec` URL in a browser tab. You should see a small
JSON reply like `{"ok":true,"app":"fd-tracker",...}`. If you see a Google
sign-in or permission page instead, redo step 3 (access must be "Anyone").

> Editing `Code.gs` later? Changes only go live after **Deploy → Manage
> deployments → ✏ edit → Version: New version → Deploy**. The URL stays the same.

## 4. Point the app at your backend

Open `index.html` in a text editor, find the line near the top of the first
`<script>` block:

```js
const APPS_SCRIPT_URL = "";
```

and paste your `/exec` URL between the quotes. While it is empty the app runs
in **Demo mode** (sample data, nothing saved) — handy for trying the UI first.

## 5. Host it on GitHub Pages

1. Create a GitHub repository (private is fine — Pages still works) and commit
   **all the files in this repo** (`index.html`, `manifest.webmanifest`,
   `sw.js`, and the `icons/` folder) to the `main` branch.
2. Repo **Settings → Pages → Build and deployment**: Source = **Deploy from a
   branch**, Branch = **main**, folder = **/ (root)**. Save.
3. After a minute your app is at `https://<username>.github.io/<repo>/`.
   Open it, sign in with the family password, and you're live. The sign-in is
   remembered on each device until that person taps **Sign out**.

`index.html` also works by simply **double-clicking the file** — useful as a
backup if GitHub is ever unreachable (the PWA extras are ignored there).

## 5b. Install it like an app (optional, recommended for parents)

- **Android (Chrome):** open the Pages URL → ⋮ menu → **Add to Home screen**
  (or **Install app**). It opens full-screen with its own ₹ icon.
- **iPhone / iPad (Safari):** open the URL → Share □↑ → **Add to Home Screen**.

Because the sign-in is remembered, the installed app opens straight into the
dashboard.

## 6. (Optional) Import the existing FDs

If you have the old multi-sheet workbook (`Fixed_Deposits.xlsx`):

```bash
pip install openpyxl
python3 migrate.py /path/to/Fixed_Deposits.xlsx --outdir migrated
```

This writes `migrated/FDs.csv` and `migrated/Accounts.csv` and prints totals —
**compare them with the old workbook's Summary Sheet** before importing.

To import each CSV into the Google Sheet:

1. Open the sheet, select cell **A1** of the matching tab (`FDs` or `Accounts`).
2. **File → Import → Upload** → choose the CSV.
3. Import location: **Replace data at selected cell**, and set
   **Convert text to numbers, dates and formulas → No**. ← this matters:
   it keeps FD account numbers and dates as plain text.
4. After importing, re-check that the three plain-text columns from step 1.4
   still show full FD numbers (no `5.03E+13`) and `YYYY-MM-DD` dates.

## The change log

A third tab named **`Log`** appears automatically the first time anyone adds,
edits or deletes something through the app. It powers the dashboard's
"Recent changes" panel and the **Undo** button — don't rename it or edit its
rows by hand (the Before/After columns hold the data Undo restores). The
optional "Device name" on the sign-in screen is what shows up as *who* made
each change.

> Updating from an earlier version? Re-paste `Code.gs` and publish a **new
> version** (Deploy → Manage deployments → ✏ → New version → Deploy),
> otherwise changes won't be logged and Undo won't work.

## Maturity reminders (push notifications) — optional

Get a phone notification **2 days before, 1 day before, and on the maturity
date**, around **10 AM IST**, on the installed app. It works by a daily Apps
Script trigger that sends a Web Push to each device that has opted in. The
notification is deliberately generic ("a fixed deposit is maturing soon — tap
to open") so no amounts ever land on a lock screen; tap it to open the app and
see which FDs in the password-protected dashboard.

One-time setup:

1. **Generate VAPID keys.** In the Apps Script editor, select the function
   `generateVapidKeys` and **Run**. Open **View → Logs**; it prints
   `VAPID_PUBLIC = …` and `VAPID_PRIVATE = …`.
2. **Project Settings → Script Properties**, add three properties:
   - `VAPID_PUBLIC` = the printed public value
   - `VAPID_PRIVATE` = the printed private value
   - `VAPID_SUBJECT` = `mailto:you@example.com` (any contact email)
3. **Set the timezone to IST.** Project Settings → **Time zone → (GMT+05:30)
   India Standard Time**. The 10 AM trigger fires in this timezone.
4. **Install the daily trigger.** Select the function `installMaturityReminders`
   and **Run** once (authorize if asked). It creates a single daily trigger for
   `sendMaturityPush` at ~10 AM. (To stop reminders entirely, run
   `removeMaturityReminders`.)
5. **Re-deploy** the web app as a **new version** (so the new `subscribe`
   action is live).
6. On each phone, open the installed app, go to the **Reminders** tab, tap
   **Enable on this device**, and allow notifications. That device is now
   subscribed. The **Reminders** tab also lists every subscribed device (by the
   "Device name" entered at sign-in) with an **Unsubscribe** button, and a
   `PushSubs` tab appears in the sheet with the same list.

Notes & limits:
- Works well on **Android** (Chrome / installed PWA). On **iPhone/iPad** it
  needs iOS **16.4+** and the app must be added to the Home Screen; Apple's
  web-push is a bit less reliable.
- Apps Script time triggers fire within ~15 minutes of the hour, so it's
  "around 10 AM IST", not to the exact minute.
- A device stays subscribed until you tap **Enable reminders** again to turn it
  off (it toggles), or the browser drops the subscription (the server prunes
  dead ones automatically).
- It only pushes on days when something is actually due, so no daily spam.

## Adding rows by hand later

The app is the comfortable way to add FDs, but editing the sheet directly is
fine — that's the point of keeping the sheet as the source of truth. If you
add a row by hand you may leave **ID** blank; the backend assigns the next
number automatically the next time anyone opens the app.

## The security model, honestly

- The password is checked **server-side** (in Apps Script) on every request;
  it never appears in `index.html` or in the browser's code.
- After a successful sign-in the password is remembered **on that device** in
  a cookie (~6 months) so the family doesn't retype it. **Sign out** forgets
  it; changing the `FAMILY_PASSWORD` script property signs every device out
  at once. Anyone who can unlock a family phone can open the app — that is
  the convenience trade-off, chosen deliberately.
- Anyone who has **both** the web-app URL **and** the password can read and
  edit the data. Anyone with only the URL gets `unauthorized`.
- Traffic is HTTPS end-to-end. The data itself sits in your Google account,
  protected by your Google login like any other sheet.
- There is no rate limiting, no audit log, and one shared password —
  **family-grade, not bank-grade**. Don't reuse a password you care about,
  and don't post the URL publicly.

## Troubleshooting

| Symptom | Fix |
|---|---|
| Login says *Wrong password* | `FAMILY_PASSWORD` script property missing or different — check step 2.2 |
| Login says *Could not open the ledger* | `APPS_SCRIPT_URL` empty/mistyped, or deployment access isn't "Anyone" |
| *sheet 'FDs' not found* | A tab is named differently — names must be exactly `FDs` and `Accounts` |
| FD numbers look like `5.03E+13` | Column wasn't plain text **before** the data arrived — reformat the column, then re-import or re-paste the numbers |
| Edits in the app don't appear in the sheet | You deployed an old version — see the note at the end of step 3 |
