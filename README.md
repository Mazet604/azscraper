# AgencyZoom Scraper

Scrapes AgencyZoom reports (Sales, Monthly, Annual) and lead/service pipelines
into Excel workbooks. It drives your **own installed browser** (Brave, Chrome,
or Edge) via Playwright, pausing so you can log in by hand, then reads the data.

## Requirements

- **Python 3.x** (developed on 3.14). Install from <https://www.python.org/downloads/>
  and tick **"Add python.exe to PATH"**. The Windows installer also includes
  `tkinter`, which the GUI front-end needs.
- **A Chromium-based browser installed**: Brave, Chrome, or Edge. The scraper
  uses whichever it finds first — it does **not** download its own browser.
- Python packages: `openpyxl`, `playwright` (see [requirements.txt](requirements.txt)).

## Setup (first time on any machine)

> Note: on this setup `pip` and `playwright` are not standalone commands — run
> them through Python with `python -m ...` as shown below.

```powershell
# 1. Clone and enter the repo
git clone https://github.com/Mazet604/azscraper.git
cd azscraper

# 2. Create and activate a virtual environment (recommended)
python -m venv .venv
.\.venv\Scripts\Activate.ps1
#   If activation is blocked once, run:
#   Set-ExecutionPolicy -Scope CurrentUser RemoteSigned

# 3. Install dependencies
python -m pip install -r requirements.txt
```

That's all. `playwright install chromium` is **not** required — this scraper
launches your system Brave/Chrome/Edge, not Playwright's bundled Chromium.

## Running

All scripts live in `agencyzoom_scraper/version 1/`.

```powershell
cd "agencyzoom_scraper\version 1"
```

**GUI (easiest):**
```powershell
python lil_scraper.py
```
Scrapes reports + all pipelines into `agencyzoom_all_data.xlsx`.

**Command line — all reports in one file:**
```powershell
python scrape_all_reports.py        # -> agencyzoom_reports.xlsx
```

**Individual scrapes:**
```powershell
python scrape_lead_pipelines.py
python scrape_service_pipelines.py
```

When a script runs, a browser window opens. **Log in to AgencyZoom**, and once
you see the dashboard, return to the terminal (or the GUI prompt) to continue.
The output `.xlsx` is written next to the script.

### Running from VS Code (the ▶ button)

The repo ships a [`.vscode/settings.json`](.vscode/settings.json) that points
both the Python extension and the Code Runner ▶ button at `.venv`. After you
create the venv (setup step 2), **reload the window once**
(Ctrl+Shift+P → "Developer: Reload Window") so the ▶ button uses the venv's
Python. Without that, the ▶ button uses the *global* Python and fails with
`ModuleNotFoundError: No module named 'openpyxl'`.

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `pip` / `playwright` "not recognized" | They aren't standalone commands on PATH | Run them via Python: `python -m pip ...`, `python -m playwright ...` |
| `ModuleNotFoundError: No module named 'openpyxl'` when using the ▶ button | Code Runner ran the **global** Python, not the venv | Reload the window so it picks up `.vscode/settings.json`; or run from an activated venv terminal |
| `Activate.ps1 ... cannot be loaded because running scripts is disabled` | PowerShell execution policy | `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` (once) |
| `Activate.ps1 is not recognized` | Wrong path to the venv | Activate the venv in **this** project: `& ".\.venv\Scripts\Activate.ps1"` |
| `Unexpected token '-u'` in PowerShell | Pasting the Code Runner command template by hand — a quoted path needs the `&` call operator | Don't type that template. Just run `python "<script>"`, or use `&`: `& ".\.venv\Scripts\python.exe" "<script>"` |
| `can't open file '...\version'` (path cut at a space) | The script path wasn't quoted as one argument | Quote the whole path: `python "azscraper\agencyzoom_scraper\version 1\lil_scraper.py"` |

**The short version of what tripped us up:** the packages were installed in
`.venv`, but the ▶ button was launching a *different* Python. Pointing Code
Runner at the venv (done in `.vscode/settings.json`) + reloading the window
fixes it. From the terminal, always activate the venv first and quote paths
that contain spaces.

## Notes

- Each new terminal session: re-activate the venv with `.\.venv\Scripts\Activate.ps1`
  before running scripts (or call `.\.venv\Scripts\python.exe <script>` directly).
- `.venv/` and the `*_profile/` browser folders are git-ignored — they're
  recreated per machine and the profile folders can hold a logged-in session.
- `agencyzoom_scraper/beta old scrape/` holds archived earlier versions.
