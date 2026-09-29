"""Central configuration for the PMIS Automation Tool.

Kept declarative and extensible: environments, update types (each mapped to a
sheet + status column + runner), and the dependency checklist all live here so
new ones can be added without touching the UI.
"""

import os
import sys

APP_NAME = "PMIS Automation Tool"
DATA_FILE_NAME = "PMIS Data.xlsx"
DATA_TEMPLATE_NAME = "PMIS Data.template.xlsx"

# GitHub repo used for the "app update available" check on startup.
# Format: "owner/repo". Leave as None to disable the check entirely.
GITHUB_REPO = "mr-imran-syed/PMIS-Automator"
GITHUB_BRANCH = "main"


def _read_version():
    """Version comes from the VERSION file so there is a single place to bump."""
    import os as _os
    try:
        with open(_os.path.join(_here(), "VERSION"), encoding="utf-8") as f:
            v = f.read().strip()
            if v:
                return v
    except OSError:
        pass
    return "0.1.0"


def _here():
    import os as _os
    if getattr(sys, "frozen", False):
        return _os.path.dirname(sys.executable)
    return _os.path.dirname(_os.path.abspath(__file__))


APP_VERSION = _read_version()
APP_FOOTER = f"Made by Imran Syed · v{APP_VERSION}"

# --- Environments (admin operation) --------------------------------------
ENVIRONMENTS = {
    "Training": "https://in-pmis-test.heifer.org/",
    "Live": "https://in-pmis.heifer.org/",
}


# --- Paths ----------------------------------------------------------------
def app_dir():
    """Folder the app lives in (works for a plain script or a frozen build)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def libs_dir():
    """App-managed package folder (pip --target installs here)."""
    return os.path.join(app_dir(), "libs")


def data_file_path():
    return os.path.join(app_dir(), DATA_FILE_NAME)


def data_template_path():
    return os.path.join(app_dir(), DATA_TEMPLATE_NAME)


def ensure_libs_on_path():
    """Put the app-managed libs folder first on sys.path so packages installed
    by the app are importable (and win over older global copies)."""
    d = libs_dir()
    os.makedirs(d, exist_ok=True)
    if d not in sys.path:
        sys.path.insert(0, d)


# --- Update types ---------------------------------------------------------
# Each maps a dropdown label to the sheet it reads, the column that records
# progress, the columns it requires, and the runner that performs it.
# `runner` is a dotted path "module:function" resolved lazily, or None (stub).
UPDATE_TYPES = {
    "Geographical Details": {
        "sheet": "Geographical Details",
        "status_col": "PMIS",
        "required": ["State", "District", "Block", "GP", "Village"],
        "runner": "pmis_geo:run_geo",
    },
    "Participant Master Data": {
        "sheet": "Participant Master Data",
        "status_col": "PMIS-Update",
        "required": ["name", "district", "block", "grampanchayat",
                     "village"],
        "runner": None,   # not implemented in v0.1
    },
    # Read-only check: verifies every unique location cascade in the
    # participant sheet actually exists on the Create Member page.
    "Participant - Geo Checker": {
        "sheet": "Participant Master Data",
        "status_col": "PMIS-Update",     # not written to; this check is read-only
        "required": ["fpo_code", "province", "district", "block",
                     "grampanchayat", "village"],
        "runner": "geo_check:run_geo_check",
        # Run on the UI thread before the browser starts; shows a popup and
        # aborts the run if it returns (False, message).
        "validator": "geo_check:validate_fpo_codes",
    },
}

# Member page used by the Participant - Geo Checker. The FPO code is added
# as ?coopId=<fpo_code>, so the cascade is checked in that coop's context.
#   {environment_url}/member/createdirect?coopId={fpo_code}
# Confirmed working when opened manually while logged in:
#   https://in-pmis.heifer.org/member/createdirect?coopId=697
# Tried in order; the first that renders the location cascade wins.
CREATE_DIRECT_PATHS = [
    "member/createdirect",
]
CREATE_DIRECT_PATH = CREATE_DIRECT_PATHS[0]   # kept for reference/logging

# Column holding the FPO / cooperative id in the participant sheet.
# Must be a whole number or the literal "NA".
PARTICIPANT_FPO_COLUMN = "fpo_code"

# Maps the cascade levels on the PMIS page to the participant sheet's
# column names. The workbook uses PMIS-style lowercase headers, so if those
# headers are renamed again, this mapping is the only place to update.
# NOTE: the sheet also has a "hamlet" column, but the page cascade has no
# hamlet control, so it is not part of the check.
PARTICIPANT_GEO_COLUMNS = {
    "State": "province",
    "District": "district",
    "Block": "block",
    "GP": "grampanchayat",
    "Village": "village",
}


# --- Dependency checklist (dynamic / extensible) --------------------------
# kind:
#   "python"  - the interpreter itself (always present while the app runs)
#   "package" - a pip package (import_name + pip_name); can Install/Update
#   "browser" - a web browser; Install opens the official download page
#   "driver"  - the Selenium driver, fetched by Selenium Manager
DEPENDENCIES = [
    {"key": "python", "name": "Python", "kind": "python", "mandatory": True},
    {"key": "selenium", "name": "Selenium", "kind": "package",
     "import_name": "selenium", "pip_name": "selenium", "mandatory": True},
    {"key": "pandas", "name": "pandas", "kind": "package",
     "import_name": "pandas", "pip_name": "pandas", "mandatory": True},
    {"key": "openpyxl", "name": "openpyxl", "kind": "package",
     "import_name": "openpyxl", "pip_name": "openpyxl", "mandatory": True},
    {"key": "browser", "name": "Web Browser", "kind": "browser",
     "mandatory": True},
    # Not mandatory: an available update must never block Start.
    {"key": "app_update", "name": "PMIS Automation Tool", "kind": "app",
     "mandatory": False},
    {"key": "driver", "name": "Browser Driver", "kind": "driver",
     "mandatory": True},
]

# Where to send users if no browser is installed.
BROWSER_DOWNLOAD_URL = "https://www.google.com/chrome/"
