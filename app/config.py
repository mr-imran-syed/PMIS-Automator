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
    """Where the code and its managed packages live (the 'app' folder)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def home_dir():
    """The folder the user actually sees: run.bat and the Excel workbook.

    Everything else is tucked inside app/ so a colleague opening the folder
    is not faced with twenty .py files.
    """
    return os.path.dirname(app_dir())


def libs_dir():
    """App-managed package folder (pip --target installs here)."""
    return os.path.join(app_dir(), "libs")


def data_file_path():
    return os.path.join(home_dir(), DATA_FILE_NAME)


def data_template_path():
    """Headers-only copy shipped with the code; never holds real data."""
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
# The participant geo checker writes here, NOT into PMIS-Update: that column
# is the entry task's done-marker, and the two must not overwrite each other.
PARTICIPANT_GEO_STATUS_COL = "Geo-Check"

UPDATE_TYPES = {
    "Geographical Details": {
        "sheet": "Geographical Details",
        "status_col": "PMIS",
        "required": ["State", "District", "Block", "GP", "Village"],
        "runner": "pmis_geo:run_geo",
    },
    # Case 1 only for now (FOAB member, not SHG). Fills the member
    # registration form and stops - it never submits.
    "Participant Master Data": {
        "sheet": "Participant Master Data",
        "status_col": "PMIS-Update",
        "required": ["fpo_code", "name", "gender", "age", "province",
                     "district", "block", "grampanchayat", "village"],
        "runner": "participant:run_participant_entry",
        "validator": "geo_check:validate_fpo_codes",
    },
    # Creates OG SHGs from the `shg` sheet. Outcomes (including the real
    # error text) are written back into the sheet's PMIS-Update column.
    "SHG Creation": {
        "sheet": "shg",
        "status_col": "PMIS-Update",
        "required": ["SHG_Name", "SHG_Type", "partner_name",
                     "work_order_number", "Date_Of_Formation", "State_Name",
                     "District_Name", "Block_Name", "Grampanchayat_Name",
                     "Village_Name", "Commodities_Name"],
        "runner": "shg:run_shg_creation",
    },
    # Read-only. Same idea as the participant checker, but it walks the SHG
    # sheet's cascades against the Create SHG page (and flags hamlets too).
    "SHG - Geo Checker": {
        "sheet": "shg",
        "status_col": "PMIS-Update",     # verdicts are written here
        "required": ["State_Name", "District_Name", "Block_Name",
                     "Grampanchayat_Name", "Village_Name"],
        "runner": "shg:run_shg_geo_check",
    },
    # Read-only check: verifies every unique location cascade in the
    # participant sheet actually exists on the Create Member page.
    "Participant - Geo Checker": {
        "sheet": "Participant Master Data",
        # Verdicts go in their OWN column: PMIS-Update is the entry task's
        # done-marker, and the two tasks must not overwrite each other.
        "status_col": PARTICIPANT_GEO_STATUS_COL,
        "required": ["fpo_code", "province", "district", "block",
                     "grampanchayat", "village"],
        "runner": "geo_check:run_geo_check",
        # Run on the UI thread before the browser starts; shows a popup and
        # aborts the run if it returns (False, message).
        "validator": "geo_check:validate_geo_rows",
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

# --- Member registration form (Participant Master Data) -----------------
# Logical field -> column in the Participant Master Data sheet. Rename-proof:
# if the workbook headers change, update here only.
PARTICIPANT_FORM_COLUMNS = {
    "name":          "name",
    "gender":        "gender",
    "age":           "age",
    "relation":      "contact_person_relation",
    "relative_name": "contact_person_name",
    # The form's Category dropdown is the caste. Read "caste" first and fall
    # back to "category" - the PMIS export fills one or the other.
    "caste":         "caste",
    "category":      "category",
    "ethnicity":     "ethnicity",      # -> form's EthnicityId select
    "contact_no":    "contact_no",
    "adhar":         "adhar_card",
    "household":     "status_in_house_hold",
    "member_type":   "member_type_name",
    "shg_code":      "shg_code",       # blank => Case 1 (not an SHG member)
}

# Valid values for the form's Category (caste) dropdown, read from the live
# member form. Anything else is flagged in the log.
PARTICIPANT_CASTE_OPTIONS = ["SC", "ST", "OBC", "GENERAL", "RELIGIOUS MINORITY"]

# Auto-calculated column the app maintains in the sheet, derived from `age`
# (the form's YearOfBirthAttr takes a full dd-mm-yyyy date).
PARTICIPANT_YOB_COLUMN = "year_of_birth"

# The sheet has no status column populated, so registrations default to this.
PARTICIPANT_DEFAULT_STATUS = "Active"

# The SHG dropdown falls back to this when the sheet gives no SHG name/code.
PARTICIPANT_SHG_NA = "NA"

# Written into the status column once a participant has been saved to PMIS,
# so an auto-save run can resume without re-entering anyone.
PARTICIPANT_DONE_MARKER = "Updated"

# Save & Continue on the member registration form.
PARTICIPANT_SAVE_BUTTON = '//input[@id="SubmitContinueButton"]'

# The thin loading bar that appears across the top of the page after saving.
# Checked in order; the first one that is actually visible is used as the
# "page is busy" indicator. If none match, the app logs the progress-bar-like
# elements it found so the right selector can be added here.
PARTICIPANT_PROGRESS_SELECTORS = [
    "#nprogress", "#nprogress .bar",          # NProgress
    ".pace.pace-active", ".pace-progress",    # Pace.js
    ".turbolinks-progress-bar",
    "#loading-bar", ".loading-bar",
    "#progressBar", "#progressbar", "#progress-bar",
    ".progress-bar-top", "#top-progress",
]

# After "Save & Continue" the page does a FULL refresh and the dropdowns are
# repopulated by the new document. That is the trigger for the next entry:
#   old DOM goes stale -> document.readyState complete -> dropdowns populated.
PARTICIPANT_RELOAD_TIMEOUT = 45   # ceiling for the whole refresh
PARTICIPANT_RELOAD_SETTLE = 0.6   # brief pause once the dropdowns are there

# Confirmed against the live form: the site uses Pace.js. <body> carries
# "pace-running" while a request is in flight and "pace-done" when idle -
# that class is the authoritative "page is busy" signal.
PARTICIPANT_BUSY_BODY_CLASS = "pace-running"

# How long to wait for an AJAX-loaded cascade option to turn up. There is no
# fixed settle: the option is taken the instant it appears.
PARTICIPANT_GEO_TIMEOUT = 12

# Speed tuning for the member form.
# The geo cascade is re-checked per participant, but "Save & Continue" keeps
# the previous location, so a level already showing the right value is left
# alone instead of being re-selected (which would cost a full AJAX settle).
PARTICIPANT_GEO_LEAD = 0.25      # AJAX head start before judging the options
PARTICIPANT_GEO_SETTLE = 0.35    # options must be stable this long
PARTICIPANT_GEO_RETRY = 3        # seconds to keep retrying a missing option
PARTICIPANT_ALERT_TIMEOUT = 1.2  # the GP/Village alert appears fast or not at all

# Stand-in for a blank/zero mobile number (matches the notebook's behaviour).
PARTICIPANT_PLACEHOLDER_MOBILE = "9000000000"

# Value-chain checkboxes, driven by columns in the participant sheet.
# A cell of True/Yes/1 ticks the box; anything else (NULL/blank/False) clears it.
#
# CAUTION on `value_id`: the source notebook had BYP_vc() and goat_vc() BOTH
# clicking //input[@value="8"] (a copy-paste bug), and clean_form() shows three
# boxes: 8, 17 and 20 (20 = agri). So BYP=8 is known, Goat=17 is only inferred
# by elimination. The app therefore matches the checkbox's visible LABEL first
# and uses value_id only as a fallback - and logs which control it actually
# used, so the mapping can be confirmed from the log.
PARTICIPANT_VALUE_CHAINS = {
    "BYP": {
        "column": "BYP Value Chain",
        "value_id": "8",
        "labels": ["backyard poultry", "byp", "poultry"],
    },
    "Goat": {
        "column": "Goat Value Chain",
        "value_id": "17",
        "labels": ["goat"],
    },
}


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


# --- SHG creation --------------------------------------------------------
SHG_SHEET = "shg"
SHG_LIST_PATH = "shgs/list"          # carries "Partner :" / "Project :"
SHG_CREATE_PATH = "shg/create?type=Shg"   # the OG "Create SHG" link

# Only this SHG_Type is handled (FOAB SHGs use ?type=Coop and are out of scope).
SHG_TYPE_ALLOWED = "OG"

SHG_STATUS_COL = "PMIS-Update"
SHG_DONE_MARKER = "SHG Created"
# Only rows carrying this marker are processed - every task that queues work.
LOAD_MARKER = "Load"
# Written when a location is missing and nothing in the dropdown is close.
NO_LOCATION = "No Location"
# difflib.get_close_matches settings for the "did you mean" suggestions.
FUZZY_MAX = 3
FUZZY_CUTOFF = 0.6
# Written by either geo checker when a cascade is fully valid.
GEO_OK_MARKER = "Geo OK"
SHG_DEFAULT_STATUS = "Active"        # the form's GroupStatus dropdown

# Hamlet values that mean "there isn't one" - the field is then left alone.
SHG_NO_HAMLET_VALUES = {"NO HAMLET", "NA", "NONE", "-"}

# Pause after saving before reading back any validation message.
SHG_SAVE_SETTLE = 1.0
# How long to wait for the "Are You Sure?" confirm, and for the save to land.
SHG_CONFIRM_TIMEOUT = 8
# The geography cascade must read correct for this long before we trust it.
# PMIS reloads a dropdown's options when its AJAX reply lands, wiping the
# selection, and those requests are invisible to Pace - so stability is the
# only signal that every reply has arrived.
# Rows marked done before the workbook is written back to disk. Saving an
# .xlsx costs far more than setting a cell; this is the resume granularity.
GEO_FLUSH_EVERY = 10
SHG_FLUSH_EVERY = 10

SHG_CASCADE_SETTLE = 1.5
SHG_CASCADE_TIMEOUT = 30

# How long to give PMIS's /SHG/IsSHGNameExists duplicate check.
SHG_NAME_CHECK_TIMEOUT = 8
SHG_SAVE_TIMEOUT = 45

# Sheet column for each piece of data.
SHG_COLUMNS = {
    "name":      "SHG_Name",
    "type":      "SHG_Type",
    "partner":   "partner_name",
    "workorder": "work_order_number",
    "date":      "Date_Of_Formation",
    "state":     "State_Name",
    "district":  "District_Name",
    "block":     "Block_Name",
    "gp":        "Grampanchayat_Name",
    "village":   "Village_Name",
    "hamlet":    "Hamlet_Name",
    "commodity": "Commodities_Name",
}

# Controls on the Create SHG form (confirmed against the live page).
SHG_FORM = {
    "name":   '//input[@id="ShgName"]',
    "date":   '//input[@id="DateOfFormationAttr"]',
    "status": '//select[@name="GroupStatus"]',
    "save":   '//input[@id="SubmitButton"]',
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
