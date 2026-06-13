# Family FD Tracker

A small, zero-maintenance web app for one family to track fixed deposits held
across many banks and four people. It replaces a multi-sheet Excel workbook
whose consolidated view was stitched together from hundreds of hardcoded
cross-sheet references (`='DAD HDFC'!B5`) that broke every time a row was
added.

- **Dashboard** — total principal, total at maturity, interest to earn,
  what's maturing in the next 120 days (or overdue), breakdowns by person and
  by account.
- **All FDs** — filter, sort, search, edit, delete, export CSV.
- **Accounts** — add, edit, delete, and **transfer** FDs between accounts from
  the app; renaming an account updates every FD under it, and deleting an
  account that still holds FDs walks you through moving them first.
- The dashboard summarises the **counted** deposits by default and can flip to
  show the **hidden & inactive** ones as their own set.
- **Add/edit form** — enter any two of *start / end / tenure* and any two of
  *rate / principal / maturity*; the third fills in automatically, and every
  computed value stays manually overridable (banks round differently — the
  sheet's number wins). One exception: a **periodic-payout** FD's maturity is
  locked to its principal and the field is disabled.
- **Readable amounts** — every big number carries its exact Indian-words
  equivalent (₹82,50,000.10 → "82 lakh 50 thousand and 10 paise") on the
  dashboard, in the table, and in the form, so a misplaced zero is obvious.
- **Installable (PWA)** — "Add to Home Screen" on Android/iPhone gives it an
  app icon and a standalone, browser-chrome-free window.
- **Remembered sign-in** — the password is kept in a cookie on the device
  (parents sign in once); a **Sign out** button forgets it. Changing
  `FAMILY_PASSWORD` in Apps Script signs every device out at once.
- **Light & dark mode** — follows the device by default, with a toggle.
- **Change log + Undo** — every add/edit/delete shows up under a **History**
  tab (with the device name that did it). Tap an entry to reopen the FD or
  account in a read-only form showing exactly what was recorded, and the
  latest change to any item can be undone — a deleted FD comes back with the
  same ID, an edit reverts to its previous values.
- **Google Sheets is the database** — the family can keep editing the sheet
  directly; the app and the sheet never fight.

## Architecture (in prose)

There are exactly three parts and no build step.

**The frontend is `index.html`** (plus a few PWA side-files: a manifest,
icons, and a tiny network-first service worker). React, ReactDOM and Babel
are loaded from cdnjs and the JSX is transpiled in the browser — no build
step; the page also runs by double-clicking it locally. With
`APPS_SCRIPT_URL` left empty it boots into a self-contained demo with sample
data; with the URL set it talks to the backend. The only things kept on the
device are two cookies: the remembered sign-in (cleared by Sign out) and the
theme choice — all data lives in memory and in the sheet.

**The backend is a Google Apps Script web app** (`Code.gs`) bound to the
spreadsheet. It exposes one `doPost` endpoint accepting
`{action: list | create | update | delete | accountCreate | accountUpdate |
accountDelete | transferFds | undo, password, device?, …}` and
always replies with the full fresh state `{ok, fds, accounts, log}`, so the client
just replaces what it has — no merging, no cache to go stale. The shared
family password is stored as a **script property**, never in code, and is
checked server-side on every request. Writes are serialized with
`LockService`. Requests are sent as `text/plain` so they stay "simple" CORS
requests that Apps Script can answer without a preflight.

**The database is the Google Sheet itself**, two tabs:

- `FDs` — one row per deposit. The columns intentionally mirror the old
  workbook's ledgers (`Status, FD Account Number, Start Date, End Date,
  Tenure (Days), Interest %, Principal Amount, Maturity Amount` in the same
  order), with `ID` and `Account` in front and three new fields after
  (`Interest On Maturity`, `Compounding`, `Show In Dashboard`), plus
  `Remarks`. The backend matches columns **by header name** (friendly names
  *and* camelCase are both recognised), so reordering or adding columns never
  breaks anything — the exact failure mode of the old workbook.
- `Accounts` — `Name, Person, Bank`. One row per account ledger (what used to
  be a tab, e.g. `DAD HDFC`). All grouping joins through this tab, so people
  and banks are data, not code.
- `Log` — auto-created on the first change. One row per add/edit/delete/undo
  with a JSON *before/after image* of the affected row — those images are
  what make Undo trustworthy. **Don't edit this tab by hand.** An entry is
  undoable only while it is the latest change to that FD/account (so an undo
  can never silently clobber a newer edit); undo entries themselves can't be
  undone.

## Repo map

| File | What it is |
|---|---|
| `index.html` | The frontend (one page, no build step) |
| `manifest.webmanifest`, `sw.js`, `icons/` | PWA bits: install metadata, network-first service worker, app icons |
| `Code.gs` | Apps Script backend — paste into the sheet's script editor |
| `SETUP.md` | One-time deployment guide (sheet → script → Pages) |
| `migrate.py` | One-time converter: old workbook → `FDs.csv` + `Accounts.csv` |

## Domain rules worth knowing

- Default interest model is **quarterly compounding on days/365**
  (`M = P·(1 + r/400)^(4·t)`), verified against the family's real bank
  figures; per-FD you can pick monthly / half-yearly / annual / simple, or
  switch the FD to **periodic payout**, in which case maturity = principal
  and the rate is informational.
- Dashboard totals count **Active FDs marked Show In Dashboard**. Inactive
  FDs never count — marking an FD Inactive automatically takes it off the
  dashboard (the form disables the checkbox). `Show In Dashboard` exists so
  an *Active* FD can still be kept out of the totals (e.g. something tracked
  separately).
- FD account numbers are **opaque strings** end to end — `50301262368467`,
  `000140451094222`, `NA`, `nsc 20159108247` all survive verbatim (the
  columns are plain-text formatted; the backend never parses them as
  numbers).
- Indian number formatting throughout: `₹1,01,59,956`.

## Try it

Open `index.html` (or the hosted page) with `APPS_SCRIPT_URL` empty, type any
password, and explore the demo. Deployment for real use: see [SETUP.md](SETUP.md).
