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
- **Add/edit form** — enter any two of *start / end / tenure* and any two of
  *rate / principal / maturity*; the third fills in automatically, and every
  computed value stays manually overridable (banks round differently — the
  sheet's number wins).
- **Google Sheets is the database** — the family can keep editing the sheet
  directly; the app and the sheet never fight.

## Architecture (in prose)

There are exactly three parts and no build step.

**The frontend is a single `index.html`.** React, ReactDOM and Babel are
loaded from cdnjs and the JSX is transpiled in the browser, so the identical
file runs by double-clicking it locally and when hosted on GitHub Pages. With
`APPS_SCRIPT_URL` left empty it boots into a self-contained demo with sample
data; with the URL set it talks to the backend. It deliberately uses no
browser storage of any kind — state lives in memory and in the sheet.

**The backend is a Google Apps Script web app** (`Code.gs`) bound to the
spreadsheet. It exposes one `doPost` endpoint accepting
`{action: list | create | update | delete, password, fd, id, newAccount}` and
always replies with the full fresh state `{ok, fds, accounts}`, so the client
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

## Repo map

| File | What it is |
|---|---|
| `index.html` | The entire frontend (single file, no build) |
| `Code.gs` | Apps Script backend — paste into the sheet's script editor |
| `SETUP.md` | One-time deployment guide (sheet → script → Pages) |
| `migrate.py` | One-time converter: old workbook → `FDs.csv` + `Accounts.csv` |

## Domain rules worth knowing

- Default interest model is **quarterly compounding on days/365**
  (`M = P·(1 + r/400)^(4·t)`), verified against the family's real bank
  figures; per-FD you can pick monthly / half-yearly / annual / simple, or
  switch the FD to **periodic payout**, in which case maturity = principal
  and the rate is informational.
- `Status` (Active/Inactive) and `Show In Dashboard` are independent flags:
  dashboard totals sum **everything marked Show In Dashboard, regardless of
  status**.
- FD account numbers are **opaque strings** end to end — `50301262368467`,
  `000140451094222`, `NA`, `nsc 20159108247` all survive verbatim (the
  columns are plain-text formatted; the backend never parses them as
  numbers).
- Indian number formatting throughout: `₹1,01,59,956`.

## Try it

Open `index.html` (or the hosted page) with `APPS_SCRIPT_URL` empty, type any
password, and explore the demo. Deployment for real use: see [SETUP.md](SETUP.md).
