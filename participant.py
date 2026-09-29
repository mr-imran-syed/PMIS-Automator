"""Participant Master Data - member registration form entry.

Case 1 only (FOAB member, not SHG): routed by fpo_code present + shg_code blank.

Ported from run_script() in PMIS_autobot SERP P3.ipynb. Differences:
  - dropdowns are matched case-insensitively with a wait, so 'HINDUISM' in the
    sheet still matches 'Hinduism' on the form
  - the geo cascade reuses GeoBot.find_and_select (AJAX settle + retry)
  - year of birth is derived from `age` into an auto-calculated sheet column
  - IsCoop is CHECKED and IsShg UNCHECKED (Case 1)

IMPORTANT: this fills the form only. It never calls Save/Continue and never
writes 'Updated' back to the sheet - nothing is submitted to PMIS.
"""

import time
from datetime import date

import pandas as pd
from openpyxl import load_workbook

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
)

import config
import settings
from pmis_geo import GeoBot, StopRequested
from geo_check import normalize_fpo, LEVEL_XPATHS, LEVEL_ORDER, _read_sheet

# EMPTY  = there is no data in the cell -> skip the field entirely.
# BLANKS = EMPTY plus "NA", meaning "not applicable" for routing purposes.
# They must stay separate: the SHG dropdown is SET to the option "NA", so
# "NA" is a real value to select, not an empty cell.
EMPTY = ("", "nan", "none", "null")
BLANKS = EMPTY + ("na",)

# Form controls (from the working notebook).
F = {
    "shg":           '//select[@id="ShgId"]',
    "name":          '//input[@name="FirstName"]',
    "gender":        '//select[@name="GenderCode"]',
    "relation":      '//select[@name="MemberContactPersonRelation"]',
    "relative_name": '//input[@name="MemberContactPersonName"]',
    "category":      '//select[@name="Category"]',
    "ethnicity":     '//select[@name="EthnicityId"]',
    "yob":           '//input[@id="YearOfBirthAttr"]',
    "education":     '//select[@id="HigherEducation"]',
    "status":        '//select[@name="StatusId"]',
    "contact_no":    '//input[@name="ContactNo"]',
    "household":     '//select[@name="StatusInHouseHold"]',
    "adhar":         '//input[@name="AdharCard"]',
    "member_type":   '//select[@name="MemberTypeId"]',
    "is_shg":        '//input[@id="IsShg"]',
    "is_coop":       '//input[@id="IsCoop"]',
}


def _blank(value):
    """No usable data, treating 'NA' as absent (used for case routing)."""
    return str(value).strip().lower() in BLANKS


def _empty(value):
    """Truly nothing to enter. 'NA' is NOT empty - it can be a real option."""
    return str(value).strip().lower() in EMPTY


def _truthy(value):
    """Sheet cell -> should the checkbox be ticked? 'NULL'/blank/False = no."""
    return str(value).strip().lower() in ("true", "yes", "y", "1")


def derive_year_of_birth(age, today=None):
    """age -> 'dd-mm-yyyy' for the YearOfBirthAttr field (which takes a full
    date, as in the notebook). Day/month default to 01-01."""
    if _blank(age):
        return ""
    try:
        years = int(float(str(age).strip()))
    except ValueError:
        return ""
    if years <= 0 or years > 120:
        return ""
    today = today or date.today()
    return f"01-01-{today.year - years}"


def classify_case(fpo_code, shg_code):
    """1 = FOAB only, 2 = SHG only, 3 = both, None = neither."""
    has_fpo = not _blank(fpo_code) and normalize_fpo(fpo_code)[1] \
        and str(fpo_code).strip().upper() != "NA"
    has_shg = not _blank(shg_code)
    if has_fpo and not has_shg:
        return 1
    if has_shg and not has_fpo:
        return 2
    if has_fpo and has_shg:
        return 3
    return None


def build_participants_df(excel_path, sheet):
    """Rows for the member form, with case + derived year of birth."""
    cols = config.PARTICIPANT_FORM_COLUMNS
    geo = config.PARTICIPANT_GEO_COLUMNS
    fpo_col = config.PARTICIPANT_FPO_COLUMN

    df = _read_sheet(excel_path, sheet)
    df["_row"] = df.index + 2                       # Excel row (incl. header)

    needed = [fpo_col] + list(geo.values()) + [cols["name"]]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(
            f"Sheet '{sheet}' is missing column(s): {', '.join(missing)}")

    # Skip anyone already saved to PMIS, so auto-save runs are resumable.
    status_col = "PMIS-Update"
    if status_col in df.columns:
        done = config.PARTICIPANT_DONE_MARKER.lower()
        df = df[df[status_col].astype(str).str.strip().str.lower() != done]
        df = df.reset_index(drop=True)

    shg_col = cols.get("shg_code")
    df["_case"] = [
        classify_case(f, df[shg_col].iloc[i] if shg_col in df.columns else "")
        for i, f in enumerate(df[fpo_col])
    ]
    age_col = cols["age"]
    df[config.PARTICIPANT_YOB_COLUMN] = [
        derive_year_of_birth(a) for a in df.get(age_col, [""] * len(df))
    ]
    return df


def write_yob_column(excel_path, sheet, df, log=print):
    """Persist the auto-calculated year-of-birth column into the sheet so it is
    visible/auditable. Best-effort: skipped if the workbook is open."""
    col_name = config.PARTICIPANT_YOB_COLUMN
    try:
        wb = load_workbook(excel_path)
        ws = wb[sheet]
        headers = {c.value: c.column for c in ws[1]}
        idx = headers.get(col_name)
        if idx is None:
            idx = ws.max_column + 1
            ws.cell(row=1, column=idx, value=col_name)
        for _, row in df.iterrows():
            ws.cell(row=int(row["_row"]), column=idx,
                    value=row[col_name] or None)
        wb.save(excel_path)
        log(f"Auto-calculated '{col_name}' written to the sheet.")
        return True
    except PermissionError:
        log(f"  ! '{col_name}' not written (workbook open in Excel) - "
            f"using the calculated values for this run anyway.")
    except Exception as e:
        log(f"  ! Could not write '{col_name}': {e}")
    return False


def mark_updated(excel_path, sheet, excel_row, log=print):
    """Stamp the status column so the row is skipped on the next run."""
    try:
        wb = load_workbook(excel_path)
        ws = wb[sheet]
        headers = {c.value: c.column for c in ws[1]}
        col = headers.get("PMIS-Update")
        if col is None:
            col = ws.max_column + 1
            ws.cell(row=1, column=col, value="PMIS-Update")
        ws.cell(row=int(excel_row), column=col,
                value=config.PARTICIPANT_DONE_MARKER)
        wb.save(excel_path)
        return True
    except PermissionError:
        log("  ! Could not mark the row (workbook open in Excel) - close it "
            "so progress can be saved.")
    except Exception as e:
        log(f"  ! Could not mark the row: {e}")
    return False


class MemberForm(GeoBot):
    """Fills the member registration form. Never submits."""

    # ---------------------------------------------------------- primitives
    def pick(self, xpath, value, timeout=10, poll=0.2):
        """Select an option by visible text, case/space-insensitive.

        'NA' is a normal value here - only a truly empty cell is skipped.
        """
        if _empty(value):
            return False
        target = str(value).strip().casefold()
        deadline = time.time() + timeout
        while True:
            try:
                sel = Select(self.driver.find_element(By.XPATH, xpath))
                for opt in sel.options:
                    if opt.text.strip().casefold() == target:
                        sel.select_by_visible_text(opt.text)
                        return True
            except (NoSuchElementException, StaleElementReferenceException):
                pass
            if time.time() >= deadline:
                self._log_options(xpath)
                return False
            time.sleep(poll)

    def _log_options(self, xpath, limit=14):
        """Show what the dropdown actually offers, so a text mismatch is
        obvious instead of just 'MISS'."""
        try:
            opts = [o.text.strip() for o in
                    Select(self.driver.find_element(By.XPATH, xpath)).options]
        except Exception:
            self.log("      (could not read the dropdown's options)")
            return
        if not opts:
            self.log("      dropdown is empty")
            return
        shown = ", ".join(repr(o) for o in opts[:limit])
        more = f" ...(+{len(opts) - limit} more)" if len(opts) > limit else ""
        self.log(f"      available options: {shown}{more}")

    def type_text(self, xpath, value, clear=True):
        if _empty(value):
            return False
        try:
            el = self.wait.until(EC.visibility_of_element_located(
                (By.XPATH, xpath)))
        except TimeoutException:
            return False
        try:
            if clear:
                el.clear()
            el.click()
            el.send_keys(str(value).strip())
            return True
        except Exception:
            return False

    def set_checkbox(self, xpath, checked):
        els = self.driver.find_elements(By.XPATH, xpath)
        if not els:
            return None                       # control not on this page
        el = els[0]
        try:
            if el.is_selected() != checked:
                try:
                    el.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", el)
            return el.is_selected() == checked
        except Exception:
            return False

    def checkbox_controls(self):
        """Every checkbox on the form as (element, value, label_text).

        Labels make the value-chain mapping verifiable instead of relying on
        magic ids.
        """
        found = []
        for el in self.driver.find_elements(
                By.XPATH, '//input[@type="checkbox"]'):
            try:
                value = el.get_attribute("value") or ""
                eid = el.get_attribute("id") or ""
                label = ""
                if eid:
                    labs = self.driver.find_elements(
                        By.XPATH, f'//label[@for="{eid}"]')
                    if labs:
                        label = labs[0].text.strip()
                if not label:
                    for rel in ("./ancestor::label[1]",
                                "./following-sibling::*[1]",
                                "./parent::*"):
                        try:
                            label = el.find_element(By.XPATH, rel).text.strip()
                            if label:
                                break
                        except Exception:
                            continue
                found.append((el, value, " ".join(label.split())[:60]))
            except Exception:
                continue
        return found

    def log_checkbox_map(self):
        """One-off dump so the BYP/Goat ids can be confirmed against reality."""
        controls = self.checkbox_controls()
        if not controls:
            self.log("    (no checkboxes found on this page)")
            return
        self.log("    checkboxes on this form (value | label):")
        for _, value, label in controls:
            self.log(f"      value={value or '(none)':<6} {label or '(no label)'}")

    def set_value_chain(self, name, spec, want, controls=None):
        """Tick/clear one value-chain box. Matches by label, falls back to id.
        Returns (ok, how) where `how` describes the control actually used."""
        controls = controls if controls is not None else self.checkbox_controls()

        chosen = None
        for el, value, label in controls:
            low = label.lower()
            if label and any(k in low for k in spec["labels"]):
                chosen = (el, value, label, "label")
                break
        if chosen is None:
            wanted_id = str(spec["value_id"])
            for el, value, label in controls:
                if value == wanted_id:
                    chosen = (el, value, label, "value_id")
                    break
        if chosen is None:
            return False, "control not found"

        el, value, label, how = chosen
        try:
            if el.is_selected() != want:
                try:
                    el.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", el)
            ok = el.is_selected() == want
        except Exception:
            return False, f"value={value} ({how}) - click failed"
        return ok, f"value={value} label='{label}' matched-by-{how}"

    # One round trip checks every candidate selector at once.
    _JS_ACTIVE_BAR = """
        // Pace.js (this site) flags a request in flight on <body>.
        var busyClass = arguments[1];
        if (busyClass && document.body &&
            document.body.classList.contains(busyClass)) {
            return 'body.' + busyClass;
        }
        // Fallback for other progress widgets. getClientRects() is used
        // instead of computed display because an element can report
        // display:block while a hidden ANCESTOR keeps it off screen --
        // .pace-progress does exactly that, and reading its own style
        // reports a bar that is permanently "visible".
        var sels = arguments[0];
        for (var i = 0; i < sels.length; i++) {
            var els = document.querySelectorAll(sels[i]);
            for (var j = 0; j < els.length; j++) {
                if (els[j].getClientRects().length > 0) {
                    return sels[i];
                }
            }
        }
        return null;
    """

    _JS_BAR_CANDIDATES = r"""
        var q = '[id*="progress"],[class*="progress"],[id*="loading"],' +
                '[class*="loading"],[class*="pace"],[id*="pace"],' +
                '[class*="spinner"],[class*="loader"]';
        var out = [];
        var els = document.querySelectorAll(q);
        for (var i = 0; i < els.length && out.length < 15; i++) {
            var e = els[i];
            out.push(e.tagName.toLowerCase() +
                     (e.id ? '#' + e.id : '') +
                     (e.className && typeof e.className === 'string'
                        ? '.' + e.className.trim().split(/\s+/).join('.') : ''));
        }
        return out;
    """

    def _active_progress_bar(self):
        """Selector of a genuinely-visible progress indicator, or None."""
        try:
            return self.driver.execute_script(
                self._JS_ACTIVE_BAR,
                list(config.PARTICIPANT_PROGRESS_SELECTORS),
                config.PARTICIPANT_BUSY_BODY_CLASS)
        except Exception:
            return None

    def wait_not_busy(self, timeout=15, poll=0.15):
        """Wait until Pace reports the page idle. Cheap and precise."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            if not self._active_progress_bar():
                return True
            time.sleep(poll)
        return False

    def log_progress_candidates(self):
        """Dump progress-bar-like elements so the right selector can be added
        to config.PARTICIPANT_PROGRESS_SELECTORS."""
        try:
            found = self.driver.execute_script(self._JS_BAR_CANDIDATES) or []
        except Exception:
            return
        if found:
            self.log("      progress-bar candidates on this page:")
            for f in found:
                self.log(f"        {f}")
        else:
            self.log("      no progress-bar-like elements found on this page")

    def wait_progress_bar(self, timeout=30, poll=0.15):
        """Wait for the page's thin loading bar to appear and then finish.

        Returns the selector that was tracked, or None if no bar ever showed
        (in which case the caller falls back to another readiness signal).
        """
        appear_deadline = time.time() + config.PARTICIPANT_PROGRESS_APPEAR
        active = None
        while time.time() < appear_deadline:
            active = self._active_progress_bar()
            if active:
                break
            time.sleep(poll)
        if not active:
            return None

        deadline = time.time() + timeout
        while time.time() < deadline:
            if not self._active_progress_bar():
                time.sleep(config.PARTICIPANT_PROGRESS_SETTLE)
                return active
            time.sleep(poll)
        return active          # finished waiting, but the bar never cleared

    def wait_reloaded(self, anchor, timeout=None):
        """Wait out the full page refresh that follows Save & Continue.

        The refresh IS the signal, so each stage is awaited directly rather
        than guessed at:
          1. the old document goes stale  -> the refresh actually started
          2. document.readyState complete -> the new document has loaded
          3. the State dropdown is populated -> the dropdowns are back
        Each stage returns the moment it is satisfied, so a fast page costs
        barely anything. Returns (ok, detail).
        """
        timeout = timeout or config.PARTICIPANT_RELOAD_TIMEOUT
        started = time.time()

        if anchor is not None:
            try:
                WebDriverWait(self.driver, timeout).until(
                    EC.staleness_of(anchor))
            except TimeoutException:
                return False, f"page never refreshed ({timeout}s)"

        remaining = max(2.0, timeout - (time.time() - started))
        try:
            WebDriverWait(self.driver, remaining).until(
                lambda d: d.execute_script(
                    "return document.readyState") == "complete")
        except Exception:
            pass        # readyState is a nicety; the dropdown check decides

        remaining = max(2.0, timeout - (time.time() - started))
        try:
            WebDriverWait(self.driver, remaining).until(
                lambda d: self.geography_loaded(("State",)))
        except TimeoutException:
            return False, "dropdowns did not repopulate after the refresh"

        time.sleep(config.PARTICIPANT_RELOAD_SETTLE)
        return True, f"refreshed in {time.time() - started:.1f}s"

    def geography_loaded(self, levels=("State", "District", "Block")):
        """True when each named cascade dropdown actually has options.

        This is the signal that the form has finished (re)loading: after
        'Save & Continue' the page keeps the previous geography, so State,
        District and Block all repopulate once the refresh lands.
        """
        try:
            for level in levels:
                sel = Select(self.driver.find_element(
                    By.XPATH, LEVEL_XPATHS[level]))
                if len(sel.options) <= 1:
                    return False
            return True
        except (NoSuchElementException, StaleElementReferenceException):
            return False
        except Exception:
            return False

    def wait_form_ready(self, levels=("State", "District", "Block"),
                        timeout=40, settle=0.4, poll=0.25, fallback=True):
        """Block until the form is genuinely usable.

        Checked twice `settle` apart, so we cannot catch a half-built cascade
        mid-rebuild and start typing into a page that is about to be replaced.
        """
        deadline = time.time() + timeout
        while time.time() < deadline:
            if self.geography_loaded(levels):
                time.sleep(settle)
                if self.geography_loaded(levels):
                    return True
            time.sleep(poll)

        # The form may legitimately come back with an empty cascade (a brand
        # new entry with no retained geography). Accept State alone rather
        # than failing the whole run.
        if fallback and len(levels) > 1 and self.geography_loaded(("State",)):
            self.log("      note: only State repopulated - continuing "
                     "(District/Block will load once State is selected)")
            return True
        return False

    def current_selection(self, xpath):
        """Visible text of the option currently selected, or ''."""
        try:
            sel = Select(self.driver.find_element(By.XPATH, xpath))
            return sel.first_selected_option.text.strip()
        except Exception:
            return ""

    def select_geo(self, level, value):
        """Set one cascade level. Returns (ok, skipped).

        'Save & Continue' keeps the previous participant's location, so when a
        level already shows the wanted value we leave it alone - re-selecting
        would cost a full AJAX settle and re-trigger the dependent lookups.
        """
        xpath = LEVEL_XPATHS[level]
        wanted = str(value).strip().casefold()
        if wanted and self.current_selection(xpath).casefold() == wanted:
            return True, True
        # Poll for the option and take it the moment it appears. State and
        # District are already fully populated on load, and the AJAX levels
        # resolve as soon as the request lands - so no fixed settle is needed
        # (that was costing ~0.6s per level for nothing).
        ok = self.pick(xpath, value, timeout=config.PARTICIPANT_GEO_TIMEOUT)
        return ok, False

    def _accept_alert(self, timeout=3):
        try:
            WebDriverWait(self.driver, timeout).until(EC.alert_is_present())
            self.alert.accept()
            return True
        except TimeoutException:
            return False

    def _report(self, label, ok, value):
        mark = "ok  " if ok else "MISS"
        shown = "(blank)" if _empty(value) else str(value).strip()
        self.log(f"    {mark} {label:<22} {shown}")

    # ------------------------------------------------------------ the form
    def fill_member(self, row):
        """Fill every mapped field for one participant. Returns a list of
        field names that could not be set."""
        cols = config.PARTICIPANT_FORM_COLUMNS
        geo = config.PARTICIPANT_GEO_COLUMNS
        failed = []

        def val(key):
            col = cols.get(key)
            return row[col] if col and col in row else ""

        def first_of(*keys):
            """First of these columns that actually holds something."""
            for k in keys:
                v = val(k)
                if not _empty(v):
                    return v
            return ""

        def do(label, ok, value):
            self._report(label, ok, value)
            if not ok and not _empty(value):
                failed.append(label)

        # ShgId must always end up selected: the SHG named in the sheet if
        # there is one, otherwise the literal option "NA".
        shg_value = val("shg_name") or val("shg_code")
        if _blank(shg_value):
            shg_value = config.PARTICIPANT_SHG_NA
        do("SHG", self.pick(F["shg"], shg_value), shg_value)

        do("Name", self.type_text(F["name"], val("name")), val("name"))
        do("Gender", self.pick(F["gender"], val("gender")), val("gender"))
        do("Relation", self.pick(F["relation"], val("relation")),
           val("relation"))
        do("Contact person", self.type_text(F["relative_name"],
                                            val("relative_name")),
           val("relative_name"))
        caste = first_of("caste", "category")
        if caste and str(caste).strip().upper() not in {
                o.upper() for o in config.PARTICIPANT_CASTE_OPTIONS}:
            self.log(f"      note: caste '{caste}' is not one of "
                     f"{', '.join(config.PARTICIPANT_CASTE_OPTIONS)}")
        do("Caste", self.pick(F["category"], caste), caste)
        do("Ethnicity", self.pick(F["ethnicity"], val("ethnicity")),
           val("ethnicity"))

        yob = row.get(config.PARTICIPANT_YOB_COLUMN, "")
        do("Year of birth", self.type_text(F["yob"], yob), yob)

        do("Status", self.pick(F["status"], config.PARTICIPANT_DEFAULT_STATUS),
           config.PARTICIPANT_DEFAULT_STATUS)

        # Blank / 0 mobile numbers get the placeholder used by the notebook.
        mob = val("contact_no")
        if _blank(mob) or str(mob).strip() in ("0", "0000000000"):
            mob = config.PARTICIPANT_PLACEHOLDER_MOBILE
        do("Mobile", self.type_text(F["contact_no"], mob), mob)

        do("Household status", self.pick(F["household"], val("household")),
           val("household"))
        do("Aadhaar", self.type_text(F["adhar"], val("adhar")), val("adhar"))
        do("Member type", self.pick(F["member_type"], val("member_type")),
           val("member_type"))

        # --- location cascade ------------------------------------------------
        for level in ("State", "District", "Block", "GP", "Village"):
            self._check_stop()
            value = row.get(geo[level], "")
            if _empty(value):
                do(f"Geo {level}", False, value)
                continue
            ok, skipped = self.select_geo(level, value)
            self._report(f"Geo {level}{' (kept)' if skipped else ''}",
                         ok, value)
            if not ok:
                failed.append(f"Geo {level}")
                # Everything below depends on this level, so it cannot
                # resolve either - stop instead of burning a timeout each.
                rest = LEVEL_ORDER[LEVEL_ORDER.index(level) + 1:]
                if rest:
                    self.log(f"      skipping {', '.join(rest)} - they depend "
                             f"on {level}")
                break
            # Only a real change can raise the confirmation dialog, so there
            # is nothing to wait for when the value was already set.
            if not skipped and level in ("GP", "Village"):
                self._accept_alert(
                    timeout=config.PARTICIPANT_ALERT_TIMEOUT)

        # --- value chains (driven by the sheet columns) --------------------
        self.log_checkbox_map()
        controls = self.checkbox_controls()
        for vc_name, vc_spec in config.PARTICIPANT_VALUE_CHAINS.items():
            cell = row.get(vc_spec["column"], "")
            want = _truthy(cell)
            ok, how = self.set_value_chain(vc_name, vc_spec, want, controls)
            state = "tick" if want else "clear"
            self._report(f"VC {vc_name} ({state})", ok, f"{cell}  [{how}]")
            if not ok:
                failed.append(f"VC {vc_name}")

        # --- membership flags ----------------------------------------------
        # IsCoop stays UNCHECKED: ticking it reveals FOAB fields (share
        # certificate no., registration date, role) that the sheet does not
        # carry. FOAB membership is already implied by the coopId in the URL.
        coop = self.set_checkbox(F["is_coop"], False)
        self._report("IsCoop (unchecked)", coop is True, "False")
        if coop is None:
            self.log("      note: IsCoop checkbox not present on this page")

        shg = self.set_checkbox(F["is_shg"], False)
        self._report("IsShg (unchecked)", shg is True, "False")
        if shg is None:
            self.log("      note: IsShg checkbox not present on this page")

        return failed

    def save_and_continue(self):
        """Submit the form. Returns (ok, message).

        The alert that follows is PMIS confirming the save, so it is accepted
        BEFORE the sheet is marked - otherwise a failed save could still be
        recorded as done.
        """
        btns = self.driver.find_elements(
            By.XPATH, config.PARTICIPANT_SAVE_BUTTON)
        if not btns:
            return False, "Save & Continue button not found"
        # Anchor on an element of the CURRENT page so the reload is detectable.
        try:
            anchor = self.driver.find_element(By.XPATH, F["name"])
        except Exception:
            anchor = None
        try:
            btns[0].click()
        except Exception as e:
            return False, f"could not click Save & Continue: {e}"

        if not self._accept_alert(timeout=10):
            return False, ("no confirmation dialog appeared - the form "
                           "probably has validation errors")

        # Save & Continue triggers a FULL page refresh, and the dropdowns
        # only come back with the new document. Wait for that refresh to
        # finish - otherwise the next participant is typed into a page that
        # is about to be replaced.
        ok, detail = self.wait_reloaded(anchor)
        if not ok:
            self.log_progress_candidates()
            return True, f"saved, but {detail}"
        return True, f"saved - {detail}"

    def goto_coop_form(self, base_url, coop_id):
        url = (f"{base_url.rstrip('/')}/{config.CREATE_DIRECT_PATH}"
               f"?coopId={coop_id}")
        self.driver.get(url)
        if self.is_signin_page():
            self.reauth_if_signed_out()
            self.driver.get(url)
        try:
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.XPATH, F["name"])))
        except TimeoutException:
            raise RuntimeError(
                f"Member registration form did not load for FPO {coop_id}.\n"
                f"  tried    : {url}\n"
                f"  landed on: {self.driver.current_url}")
        # On a fresh open only State is populated; District/Block fill in once
        # a State is chosen, so require State alone here.
        if not self.wait_form_ready(levels=("State",), timeout=25):
            self.log("      note: the State list did not populate in time")
        self.log(f"Member form open for FPO {coop_id}: {url}")


# ---------------------------------------------------------------- runner
def run_participant_entry(driver, excel_path, base_url, username, password,
                          sheet="Participant Master Data", status_col=None,
                          log=print, stop_event=None):
    """Enter Case 1 participants into the member registration form.

    Auto Save OFF (default): fills the FIRST pending participant and stops,
        leaving the form on screen. Nothing is submitted.
    Auto Save ON  (Settings tab): fills, saves, marks the sheet 'Updated' and
        moves to the next participant.
    """
    auto_save = bool(settings.get("participant_auto_save"))
    bot = MemberForm(driver, log=log, stop_event=stop_event)

    df = build_participants_df(excel_path, sheet)
    write_yob_column(excel_path, sheet, df, log)

    case1 = df[df["_case"] == 1]
    tally = {1: 0, 2: 0, 3: 0, "none": 0}
    for c in df["_case"]:
        tally[c if c in (1, 2, 3) else "none"] += 1
    log(f"Pending rows: {len(df)}  |  Case 1 (FOAB, not SHG): {tally[1]}  "
        f"|  Case 2 (SHG only): {tally[2]}  |  Case 3 (both): {tally[3]}  "
        f"|  neither: {tally['none']}")
    log(f"Auto Save is {'ON - entries WILL be saved to PMIS' if auto_save else 'OFF - nothing will be submitted'}.")

    if case1.empty:
        log("No Case 1 rows to enter (need a valid fpo_code and blank SHG).")
        return 0

    bot.ensure_logged_in(username, password, base_url)

    name_col = config.PARTICIPANT_FORM_COLUMNS["name"]
    saved = 0
    current_coop = None

    try:
        for _, row in case1.iterrows():
            bot._check_stop()
            coop_id = normalize_fpo(row[config.PARTICIPANT_FPO_COLUMN])[0]
            excel_row = int(row["_row"])

            # Only reload the form when the FPO changes.
            if coop_id != current_coop:
                bot.goto_coop_form(base_url, coop_id)
                current_coop = coop_id

            log("")
            log(f"Excel row {excel_row}  [FPO {coop_id}]  {row[name_col]}")
            failed = bot.fill_member(row)

            # The row is marked as soon as every field went in - before any
            # save - so a manually-saved record is not offered again.
            if not failed:
                mark_updated(excel_path, sheet, excel_row, log)
                log(f"  marked row {excel_row} "
                    f"'{config.PARTICIPANT_DONE_MARKER}' (fields complete)")

            if not auto_save:
                log("")
                log("=" * 60)
                if failed:
                    log(f"Filled, with {len(failed)} field(s) NOT set: "
                        f"{', '.join(failed)}")
                    log("Row NOT marked - fix the data and run again.")
                else:
                    log("All mapped fields were set successfully.")
                    log(f"Row {excel_row} is marked "
                        f"'{config.PARTICIPANT_DONE_MARKER}' and will be "
                        f"skipped next run - click Save on the form yourself.")
                log("NOT SAVED - Auto Save is off. Nothing was submitted.")
                log("=" * 60)
                return 1

            # --- Auto Save ------------------------------------------------
            if failed:
                log("")
                log("=" * 60)
                log(f"STOPPING before save: {len(failed)} field(s) could not "
                    f"be set ({', '.join(failed)}).")
                log("Fix the data or the mapping, then run again. Nothing was "
                    "submitted for this participant.")
                log("=" * 60)
                return saved

            ok, message = bot.save_and_continue()
            if not ok:
                log("")
                log("=" * 60)
                log(f"SAVE FAILED for Excel row {excel_row}: {message}")
                log("Stopping so the form can be checked. Earlier "
                    f"participants ({saved}) were saved.")
                log("=" * 60)
                return saved

            saved += 1
            log(f"  saved ({message})")

    except StopRequested:
        log("")
        log(f"Stopped by user. {saved} participant(s) saved.")
        return saved

    log("")
    log("=" * 60)
    log(f"Done. {saved} participant(s) saved to PMIS.")
    log("=" * 60)
    return saved
