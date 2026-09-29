"""Participant - Geo Checker.

Verifies that every location cascade used by the participant master data
actually exists in PMIS, by driving the cascading dropdowns on the member
page for that participant's FPO:

    {environment_url}/member/createdirect?coopId={fpo_code}

Because the page is scoped to a coop, the check is done per
(fpo_code + location cascade) pair - the same village may resolve under one
FPO and not another. Each coop's page is loaded once and all of its cascades
are checked before moving on.

Ported from the 'Geo Location Check' section of PMIS_autobot hhai bs.ipynb, with
one important change: selection uses GeoBot.find_and_select (settle + retry)
rather than a bare select_by_visible_text. The dropdowns load via AJAX, so a
plain lookup can miss an option that simply has not arrived yet and wrongly
report an existing location as missing.
"""

import pandas as pd

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import config
from pmis_geo import GeoBot, StopRequested

# The cascade, in order, with the control that selects each on the page.
# Which SHEET COLUMN feeds each level lives in config.PARTICIPANT_GEO_COLUMNS,
# so a header rename in the workbook is a one-line change there.
LEVEL_ORDER = ["State", "District", "Block", "GP", "Village"]
LEVEL_XPATHS = {
    "State":    '//select[@name="ProvinceId"]',
    "District": '//select[@name="DistrictId"]',
    "Block":    '//select[@name="MunicipalityId"]',
    "GP":       '//select[@name="GrampanchayatId"]',
    "Village":  '//select[@name="VillageId"]',
}


def _read_sheet(excel_path, sheet):
    """Read a sheet WITHOUT pandas' default NaN conversion.

    Critical: pandas treats the literal string "NA" as NaN by default, which
    would turn a valid fpo_code of "NA" into a blank and wrongly reject it.
    keep_default_na=False keeps it as the text "NA"; empty cells arrive as ''.
    """
    return pd.read_excel(excel_path, sheet_name=sheet,
                         keep_default_na=False, na_values=[])


# ------------------------------------------------------------- fpo_code
def normalize_fpo(value):
    """Return (normalized, is_valid) for one fpo_code cell.

    Valid values are a whole number (-> '697') or the literal 'NA'.
    Blank, NaN, decimals and anything else are invalid.
    """
    if value is None:
        return "", False
    try:
        if pd.isna(value):
            return "", False
    except (TypeError, ValueError):
        pass

    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return "", False
    if s.upper() == "NA":
        return "NA", True
    try:
        f = float(s)
    except ValueError:
        return s, False
    if f.is_integer():
        return str(int(f)), True      # handles pandas' 697.0 -> '697'
    return s, False


def validate_fpo_codes(excel_path, sheet):
    """Pre-flight check run on the UI thread before the browser starts.

    Returns (ok, message). When not ok, `message` is shown in a popup.
    """
    col = config.PARTICIPANT_FPO_COLUMN
    try:
        df = _read_sheet(excel_path, sheet)
    except Exception as e:
        return False, f"Could not read sheet '{sheet}':\n{e}"

    if col not in df.columns:
        return False, (f"The sheet '{sheet}' has no '{col}' column.\n\n"
                       f"The Geo Checker needs an FPO code for each row to "
                       f"open the right member page.")
    if df.empty:
        return False, f"The sheet '{sheet}' has no rows to check."

    bad = []
    for idx, value in df[col].items():
        _, ok = normalize_fpo(value)
        if not ok:
            shown = "(blank)" if str(value).strip() in ("", "nan", "None") \
                else repr(value)
            bad.append((idx + 2, shown))      # +2 -> Excel row incl. header

    if not bad:
        return True, ""

    preview = "\n".join(f"    row {r}:  {v}" for r, v in bad[:10])
    more = f"\n    ...and {len(bad) - 10} more" if len(bad) > 10 else ""
    return False, (
        f"No valid FPO code in {len(bad)} row(s) of '{sheet}'.\n\n"
        f"'{col}' must be a number (e.g. 697) or \"NA\".\n\n"
        f"{preview}{more}\n\n"
        f"Fix these rows in the data file, then Start again.")


class GeoChecker(GeoBot):
    """Reuses GeoBot's login + AJAX-safe dropdown helpers, but never creates
    anything - this check is strictly read-only."""

    def _page_snippet(self, limit=300):
        """Short, whitespace-collapsed sample of the page body, for diagnostics."""
        try:
            text = self.driver.find_element(By.TAG_NAME, "body").text
        except Exception:
            return "(could not read page text)"
        text = " ".join(text.split())
        return text[:limit] + ("..." if len(text) > limit else "")

    def _selects_on_page(self, limit=30):
        """name/id of every <select> on the page.

        If the member page uses different control names than the geolocation
        page, this is what reveals it - the page loads fine but our XPaths
        never match.
        """
        found = []
        try:
            for el in self.driver.find_elements(By.TAG_NAME, "select")[:limit]:
                try:
                    found.append(el.get_attribute("name")
                                 or el.get_attribute("id") or "(unnamed)")
                except Exception:
                    continue
        except Exception:
            pass
        return found

    def _try_member_url(self, url, timeout=20):
        """Navigate and report whether the location cascade appeared.

        PMIS redirects unauthenticated requests to
        /SignIn/Index?ReturnUrl=..., so if we land there we sign in again and
        re-request the page.
        """
        self.driver.get(url)
        if self.is_signin_page():
            self.reauth_if_signed_out()
            self.driver.get(url)        # request it again, now authenticated
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, LEVEL_XPATHS["State"])))
            return True
        except TimeoutException:
            return False

    def goto_fpo_member_page(self, base_url, coop_id):
        """Open the member page for one FPO. Returns True if it loaded.

        Tries the configured path first, then the '/cooperative/...' variant,
        since the plain create page lives under /cooperative/member/create.
        On failure it logs what the browser actually landed on, so a wrong
        path / expired session / bad coopId can be told apart.
        """
        base = base_url.rstrip("/")
        candidates = [f"{base}/{p.strip('/')}?coopId={coop_id}"
                      for p in config.CREATE_DIRECT_PATHS]

        for url in candidates:
            if self._try_member_url(url):
                self.log(f"FPO {coop_id}  ->  {url}")
                return True

        # Nothing worked - report exactly what happened on the last attempt.
        landed = self.driver.current_url
        self.log(f"  ! Member page did not load for FPO {coop_id}")
        for u in candidates:
            self.log(f"    tried    : {u}")
        self.log(f"    landed on: {landed}")
        try:
            self.log(f"    title    : {self.driver.title}")
        except Exception:
            pass
        self.log(f"    page says: {self._page_snippet()}")

        selects = self._selects_on_page()
        if selects:
            self.log(f"    <select> controls found: {', '.join(selects)}")
            self.log("    -> the page DID load, but none matched "
                     f"{LEVEL_XPATHS['State']}. The cascade control names on "
                     "this page differ; send me this list and I'll remap them.")
        else:
            self.log("    <select> controls found: none")
        if "signin" in landed.lower() or "login" in landed.lower():
            self.log("    -> redirected to sign-in: the session was not "
                     "carried over to this page.")
        return False

    def build_location_df(self, excel_path, sheet):
        """Unique (fpo_code + location cascade) pairs from the participant sheet.

        Sheet columns are mapped onto canonical level names via
        config.PARTICIPANT_GEO_COLUMNS, so the rest of the check is insulated
        from header renames in the workbook.
        """
        colmap = config.PARTICIPANT_GEO_COLUMNS      # level -> sheet column
        fpo_col = config.PARTICIPANT_FPO_COLUMN
        wanted = [colmap[lvl] for lvl in LEVEL_ORDER]

        df = _read_sheet(excel_path, sheet)
        missing = [c for c in [fpo_col] + wanted if c not in df.columns]
        if missing:
            raise ValueError(
                f"Sheet '{sheet}' is missing column(s): {', '.join(missing)}. "
                f"The Geo Checker needs: {', '.join([fpo_col] + wanted)}.")

        loc = df[[fpo_col] + wanted].copy()
        loc.columns = ["FPO"] + LEVEL_ORDER          # canonical names
        loc["FPO"] = loc["FPO"].map(lambda v: normalize_fpo(v)[0])
        for c in LEVEL_ORDER:
            loc[c] = loc[c].astype(str).str.strip()
        # Drop incomplete rows: no FPO, or any cascade level blank.
        keep = loc["FPO"] != ""
        for c in ["District", "Block", "GP", "Village"]:
            keep &= loc[c] != ""
        loc = loc[keep]
        # One row per distinct FPO + full cascade.
        loc = loc.drop_duplicates(
            subset=["FPO"] + LEVEL_ORDER).reset_index(drop=True)
        return loc

    def _dismiss_alert(self, timeout=3):
        """Selecting a village pops an alert on this page; accept it if shown."""
        try:
            WebDriverWait(self.driver, timeout).until(EC.alert_is_present())
            self.alert.accept()
        except TimeoutException:
            pass

    def _check_one(self, row, retry):
        """Walk one cascade; return the level that failed, or None if all ok."""
        for level in LEVEL_ORDER:
            value = row[level]
            if not value or value.lower() == "nan":
                return level
            if not self.find_and_select(LEVEL_XPATHS[level], value, retry=retry):
                return level
            if level == "Village":
                self._dismiss_alert()
        return None

    def check_locations(self, loc_df, base_url, retry=4):
        """Check every cascade, one FPO page at a time. Returns unavailable rows."""
        not_available = []
        total = len(loc_df)
        done = 0

        for coop_id, group in loc_df.groupby("FPO", sort=False):
            self._check_stop()
            self.log("")
            page_ok = self.goto_fpo_member_page(base_url, coop_id)

            for _, row in group.iterrows():
                self._check_stop()
                done += 1
                trail = (f"{row['District']} > {row['Block']} > "
                         f"{row['GP']} > {row['Village']}")
                self.log(f"#{done}/{total}  [FPO {coop_id}]  {trail}")

                if not page_ok:
                    rec = row.to_dict()
                    rec["Missing"] = "FPO page"
                    not_available.append(rec)
                    self.log("    SKIPPED - member page did not load")
                    continue

                missing_at = self._check_one(row, retry)
                if missing_at:
                    self.log(f"    NOT FOUND - {missing_at}: {row[missing_at]}")
                    rec = row.to_dict()
                    rec["Missing"] = missing_at
                    not_available.append(rec)
                else:
                    self.log("    ok")
        return not_available


def _report(log, loc_df, not_available):
    log("")
    log("=" * 60)
    log(f"Participant - Geo Checker: {len(loc_df)} location(s) checked "
        f"across {loc_df['FPO'].nunique()} FPO(s)")
    if not not_available:
        log("All locations are available in PMIS.")
    else:
        log(f"{len(not_available)} location(s) NOT available:")
        for m in not_available:
            log(f"  [FPO {m['FPO']}] [missing {m['Missing']}]  "
                f"{m['District']} > {m['Block']} > {m['GP']} > {m['Village']}")
        log("")
        log("Add the missing entries using 'Geographical Details', then re-run.")
    log("=" * 60)


# ---------------------------------------------------------------- runner
def run_geo_check(driver, excel_path, base_url, username, password,
                  sheet="Participant Master Data", status_col=None,
                  log=print, stop_event=None):
    """Entry point used by the app (config -> 'geo_check:run_geo_check')."""
    bot = GeoChecker(driver, log=log, stop_event=stop_event)
    bot.ensure_logged_in(username, password, base_url)

    loc_df = bot.build_location_df(excel_path, sheet)
    if loc_df.empty:
        log("No locations found in the participant sheet - nothing to check.")
        return 0
    log(f"{len(loc_df)} location(s) to check across "
        f"{loc_df['FPO'].nunique()} FPO(s).")

    try:
        not_available = bot.check_locations(loc_df, base_url)
    except StopRequested:
        log("Stopped by user - partial check only.")
        return 0

    _report(log, loc_df, not_available)
    return len(loc_df)
