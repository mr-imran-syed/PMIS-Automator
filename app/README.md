# PMIS Automation Tool

A small portable Windows utility that automates repetitive data entry in PMIS
(Heifer India) using Selenium.

Current features:

| Option | What it does |
|---|---|
| **Geographical Details** | Adds Block / GP / Village / Hamlet entries to PMIS from the `Geographical Details` sheet. Resumable. |
| **Participant - Geo Checker** | Read-only. Verifies every unique location cascade in the participant sheet exists, per FPO, and reports what's missing. |
| **Participant Master Data** | Placeholder — not implemented yet. |

---

## Running SHG tasks

Both **SHG Creation** and **SHG - Geo Checker** work on the rows whose
`PMIS-Update` cell says `Load`. Type `Load` against the rows you want, run the
task, and the result replaces it. A row that is not marked `Load` is left alone.

SHG Creation runs in two passes. First it checks every marked name against
PMIS and writes `SHG Name already exists` for the ones already there; then it
creates the rest and writes `SHG Created`. Anything else in the cell is the
error PMIS gave back. The Geo Checker writes `Geo OK`, a near-miss suggestion
(`Block not found: GHATAGAON -> GHATGAON`), or `No Location`.

## For colleagues: installing and running

1. **Install Python 3.10+** from <https://www.python.org/downloads/> and tick
   **"Add python.exe to PATH"** on the first screen.
2. Install **Firefox** or **Chrome** (Edge also works and is already on Windows).
3. Download this repo (green **Code** button → **Download ZIP**) and unzip it.
4. Double-click **`run.bat`**.
5. In the **System Readiness** panel, click **Install** next to anything marked
   missing. The app installs its own packages into a local `libs/` folder —
   your global Python is left untouched. Needs internet the first time.
6. Choose **Training** or **Live**, enter your own PMIS username and password,
   pick what to update, then **Start**.

Your password is never stored — you type it each run.

---

## The data file

The app reads and writes **`PMIS Data.xlsx`**, which sits next to `run.bat`.

Only **`PMIS Data.template.xlsx`** (headers, no rows) is committed here. On
first launch the app copies the template to `PMIS Data.xlsx` for you.

> **`PMIS Data.xlsx` is deliberately git-ignored.** It contains participant
> personal data — names, phone numbers, Aadhaar and bank account numbers.
> Never commit it or attach it to an issue.

Sheets:

- `Geographical Details` — `State, District, Block, GP, Village, Hamlet_Name, PMIS` (`Hamlet_Name` optional - blank cells are skipped)
- `Participant Master Data` — PMIS bulk-upload columns (`fpo_code`, `province`,
  `district`, `block`, `grampanchayat`, `village`, …)
- `code_book`

Keep the file **closed in Excel while a run is in progress** — the app writes
progress into it and will prompt you if it's locked.

### Resuming

Completed rows are stamped `Checked` in the status column and skipped next
time, so a run that stops (Stop, crash, network) picks up where it left off.

---

## Updating

On startup the app checks GitHub for a newer version and shows
**⚠ Update Available** in System Readiness with a **Get Update** button. An
available update never blocks a run.

The check reads the latest release tag, falling back to the `VERSION` file on
the default branch.

**Maintainer:** bump `VERSION` (single source of truth — `APP_VERSION` and the
footer both read it), commit, and publish a release with a matching tag:

```bash
echo 0.1.1 > VERSION
git commit -am "Release v0.1.1"
git tag v0.1.1 && git push && git push --tags
```

---

## Configuration

Everything environment-specific lives in [`config.py`](config.py):

| Setting | Purpose |
|---|---|
| `ENVIRONMENTS` | Training / Live base URLs |
| `GITHUB_REPO` | `owner/repo` used for the update check |
| `UPDATE_TYPES` | Dropdown options → sheet, status column, runner, validator |
| `PARTICIPANT_GEO_COLUMNS` | Cascade level → sheet column (update here if headers get renamed) |
| `PARTICIPANT_FPO_COLUMN` | Column holding the FPO id (a number, or `NA`) |
| `DEPENDENCIES` | The System Readiness checklist |

## Layout

A colleague opening the folder sees three things; the code is tucked away.

```
run.bat              Launcher - double-click to start
PMIS Data.xlsx       The workbook (gitignored: it holds participant PII)
Read me first.txt    One-page guide for colleagues
app/
  app.py             Tkinter UI (readiness, workflow, status, log)
  logo.py            The steaming-idli mark, drawn in pixels at runtime
  config.py          All configuration
  deps.py            Dependency checks + in-app pip install/update
  browsers.py        Firefox -> Chrome -> Edge detection + driver factory
  data_file.py       Workbook paths, row reader, open-in-Excel detection
  pmis_geo.py        Login + AJAX-safe dropdown helpers + GEO automation
  geo_check.py       Participant - Geo Checker
  participant.py     Member registration form
  shg.py             SHG creation + SHG geo checker
  credstore.py       Windows DPAPI credential storage
  settings.py        User settings
  build.bat          Package a shareable ZIP
  VERSION            Bumped automatically on every commit
  libs/              App-managed packages (selenium, openpyxl)
  test_shg.py        Self-checks - python test_shg.py
  test_geo_cascade.py
  test_geo_check.py
```

---

Made by Imran Syed
