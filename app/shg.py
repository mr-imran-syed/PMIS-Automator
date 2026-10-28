"""SHG creation from the `shg` sheet (OG SHGs only).

Flow per run:
  1. Open the SHG list and read the header - "Partner : X", "Project : CODE-X"
     and "Workorder No : N" - then confirm all three against the sheet.
  2. For each pending OG row: open Create SHG, re-check the workorder, fill
     the form, and (when Auto Save is on) save.
  3. Write the outcome into the PMIS-Update column: the done marker on
     success, or the actual error text ("SHG Name already exists", a missing
     hamlet, ...) so the sheet says why a row did not go through.

Verified against the live form: 9 required fields (SHG Name, Date Of
Formation, SHG Status, State, District, Block, Gram Panchayat, Village,
Hamlet); commodities are checkboxes named CommodityType; Save is
input#SubmitButton (a JS button, not a real submit).
"""

import re
import time
from datetime import datetime, date

from openpyxl import load_workbook

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import config
import data_file
import settings
from pmis_geo import StopRequested, LEVELS
from participant import MemberForm, _empty, _blank
from geo_check import LEVEL_XPATHS, near_miss

# Cascade on the SHG form - same control names as elsewhere, plus Hamlet.
# Same cascade controls as everywhere else - see pmis_geo.LEVELS.
SHG_LEVELS = [name for name, _, _, _ in LEVELS]
SHG_LEVEL_XPATHS = LEVEL_XPATHS
SHG_LEVEL_IDS = {name: sel for name, _, sel, _ in LEVELS}

# Written to PMIS-Update when PMIS already holds this SHG name.
NAME_EXISTS = "SHG Name already exists"


def _done_markers():
    """Statuses that mean the SHG is in PMIS - "Updated" and "Success" are
    what older runs wrote, and those rows must stay protected."""
    return (config.SHG_DONE_MARKER.strip().lower(), "updated", "success")

# Where inline validation errors show up after a failed save.
ERROR_SELECTORS = [
    ".validation-summary-errors", ".field-validation-error",
    ".alert-danger", ".alert-error", ".text-danger",
    "#errorMessage", ".error-message",
    "#toast-container .toast-error", "#toast-container .toast-warning",
    "#shgNameError",
]

# toastr markup for a win: <div class="toast toast-success">Success</div>.
# showNotification("success") fires this just before the page redirects.
SUCCESS_SELECTOR = "#toast-container .toast-success"


# ---------------------------------------------------------------- helpers
def format_date(value):
    """Any date cell -> 'dd-mm-yyyy' for DateOfFormationAttr."""
    if _empty(value):
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d-%m-%Y")
    text = str(value).strip()
    # Already dd-mm-yyyy (or d-m-yyyy) - normalise the padding.
    m = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$", text)
    if m:
        return "%02d-%02d-%s" % (int(m.group(1)), int(m.group(2)), m.group(3))
    for fmt in ("%Y-%m-%d", "%d-%b-%Y", "%d %b %Y", "%d/%m/%y",
                "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).strftime("%d-%m-%Y")
        except ValueError:
            continue
    return text


def parse_partner_header(page_text):
    """Pull Partner / Project / Workorder out of the page header.

    "Project : OSED-WOSCA" means project code OSED and partner WOSCA: the
    text before the hyphen is the project, the text after it is the partner.
    """
    text = " ".join((page_text or "").split())
    out = {"partner": None, "project": None, "partner_from_project": None,
           "workorder": None}
    m = re.search(r"Partner\s*:\s*([A-Za-z0-9_ ]+?)\s{2,}|Partner\s*:\s*(\S+)",
                  text)
    if m:
        out["partner"] = (m.group(1) or m.group(2) or "").strip()
    m = re.search(r"Project\s*:\s*([A-Za-z0-9_]+)-([A-Za-z0-9_ ]+?)(?:\s{2,}|$|\s[A-Z][a-z])",
                  text)
    if m:
        out["project"] = m.group(1).strip()
        out["partner_from_project"] = m.group(2).strip()
    m = re.search(r"Workorder\s*No\s*:?\s*([A-Za-z0-9/\-]+)", text, re.I)
    if m:
        out["workorder"] = m.group(1).strip()
    return out


def read_shg_sheet(excel_path, sheet=None):
    """Every row of the SHG sheet, with the columns we rely on checked.

    Read once per run and passed around: load_shgs() and known_names() both
    want the same rows, and parsing the workbook twice buys nothing.
    """
    sheet = sheet or config.SHG_SHEET
    headers, rows = data_file.read_rows(excel_path, sheet)
    missing = [c for c in config.SHG_COLUMNS.values() if c not in headers]
    if missing:
        raise ValueError(f"Sheet '{sheet}' is missing column(s): "
                         f"{', '.join(missing)}")
    return rows


def load_rows_only(rows):
    """Keep only the rows the user has marked with config.LOAD_MARKER.

    Both the creator and the geo checker run on exactly these rows, so
    queueing work is a matter of typing 'Load' in PMIS-Update.
    """
    want = config.LOAD_MARKER.strip().lower()
    return [r for r in rows
            if data_file.text(r, config.SHG_STATUS_COL).lower() == want]


def load_shgs(rows):
    """Marked rows of the one SHG type this app creates."""
    allowed = config.SHG_TYPE_ALLOWED.strip().lower()
    return [r for r in load_rows_only(rows)
            if data_file.text(r, config.SHG_COLUMNS["type"]).lower() == allowed]


def known_names(rows):
    """SHG names the sheet already shows as in PMIS - created earlier, or
    rejected as duplicates. Sheet names are not unique, so this is what stops
    the same name being attempted twice."""
    done = _done_markers() + (NAME_EXISTS.lower(),)
    return {data_file.text(r, config.SHG_COLUMNS["name"]).casefold()
            for r in rows
            if data_file.text(r, config.SHG_STATUS_COL).lower().startswith(done)
            and data_file.text(r, config.SHG_COLUMNS["name"])}


def write_status_bulk(excel_path, sheet, mapping, log=print,
                      protect_marker=True):
    """Write many statuses into PMIS-Update: {excel_row: text}.

    Rows already carrying a done marker are left alone by default, so a geo
    check can never wipe the record of an SHG that was actually created.
    """
    return data_file.write_column(
        excel_path, sheet, config.SHG_STATUS_COL, mapping, log,
        protect=_done_markers() if protect_marker else ())


class ShgForm(MemberForm):
    """Reuses the member form's primitives (pick / type_text / checkbox /
    alert handling / Pace-aware waits) for the SHG screens."""

    # ------------------------------------------------------- verification
    def verify_partner(self, base_url, sheet_partner, sheet_workorder):
        """Confirm the logged-in context matches the sheet. (ok, detail)."""
        url = base_url.rstrip("/") + "/" + config.SHG_LIST_PATH
        self.driver.get(url)
        if self.is_signin_page():
            self.reauth_if_signed_out()
            self.driver.get(url)
        try:
            self.wait.until(lambda d: d.find_element(By.TAG_NAME, "body").text)
        except TimeoutException:
            return False, f"could not open {url}"

        info = parse_partner_header(
            self.driver.find_element(By.TAG_NAME, "body").text)
        self.log(f"Page header: partner={info['partner']!r} "
                 f"project={info['project']!r} "
                 f"partner-from-project={info['partner_from_project']!r} "
                 f"workorder={info['workorder']!r}")

        problems = []
        page_partner = info["partner_from_project"] or info["partner"]
        if sheet_partner and page_partner and \
                page_partner.strip().lower() != str(sheet_partner).strip().lower():
            problems.append(f"partner {page_partner!r} != sheet "
                            f"{str(sheet_partner).strip()!r}")
        if sheet_workorder and info["workorder"] and \
                info["workorder"].strip().lower() != str(sheet_workorder).strip().lower():
            problems.append(f"workorder {info['workorder']!r} != sheet "
                            f"{str(sheet_workorder).strip()!r}")
        if not page_partner and not info["workorder"]:
            problems.append("no partner/workorder found on the page")

        if problems:
            return False, "; ".join(problems)
        return True, (f"partner {page_partner}, workorder {info['workorder']}")

    def open_create_form(self, base_url, expect_workorder=None):
        url = base_url.rstrip("/") + "/" + config.SHG_CREATE_PATH
        self.driver.get(url)
        if self.is_signin_page():
            self.reauth_if_signed_out()
            self.driver.get(url)
        try:
            WebDriverWait(self.driver, 25).until(
                EC.presence_of_element_located(
                    (By.XPATH, config.SHG_FORM["name"])))
            # The form is only usable once initializeCascadeLocation() has
            # filled the State list - the markup arrives well before that.
            WebDriverWait(self.driver, 25).until(
                lambda d: len(d.find_elements(
                    By.XPATH, SHG_LEVEL_XPATHS["State"] + "/option")) > 1)
        except TimeoutException:
            raise RuntimeError(f"Create SHG form did not load ({url}); "
                               f"landed on {self.driver.current_url}")
        self.wait_not_busy(timeout=20)
        if expect_workorder:
            info = parse_partner_header(
                self.driver.find_element(By.TAG_NAME, "body").text)
            if info["workorder"] and info["workorder"].strip().lower() != \
                    str(expect_workorder).strip().lower():
                raise RuntimeError(
                    f"Workorder on the create page ({info['workorder']}) does "
                    f"not match the sheet ({expect_workorder})")

    # ------------------------------------------------------------- filling
    def options(self, xpath):
        """Option texts of one dropdown.

        A dialog blocks every command, so the retry clears one - but asking
        for an alert that is not there costs a round trip and an exception,
        so only do it after something has actually failed.
        """
        from selenium.webdriver.support.select import Select
        for _ in range(2):
            try:
                sel = Select(self.driver.find_element(By.XPATH, xpath))
                return [o.text.strip() for o in sel.options]
            except Exception as e:
                self._note_alert(e)
        return []

    def hamlet_options(self):
        return self.options(SHG_LEVEL_XPATHS["Hamlet"])

    def hamlet_fallback(self, choices):
        """What to put in Hamlet when the sheet does not name a real one.

        PMIS makes Hamlet mandatory wherever the village has any, and offers
        no "none of these" entry - leaving it on "Select" fails validation
        silently, with no dialog and no message. So take an option that
        literally means no hamlet if the village has one, else the first.
        """
        for o in choices:
            if o.strip().upper() in config.SHG_NO_HAMLET_VALUES:
                return o
        return choices[0]

    def select_cascade(self, row, cols):
        """State -> Hamlet. Returns (failed_levels, notes)."""
        failed, notes = [], []
        self._cascade_set = {}        # level -> value, for verify_cascade
        level_to_col = {
            "State": cols["state"], "District": cols["district"],
            "Block": cols["block"], "GP": cols["gp"],
            "Village": cols["village"], "Hamlet": cols["hamlet"],
        }
        for level in SHG_LEVELS:
            self._check_stop()
            value = row.get(level_to_col[level], "")

            if level == "Hamlet":
                # Selecting the village may have raised "There are no Hamlets
                # for this village!" - clear it before touching the field.
                notice = self.consume_alert() or self.take_last_alert()
                if notice:
                    self.log(f"      PMIS says: {notice}")

                choices = [o for o in self.hamlet_options()
                           if o and not o.lower().startswith("select")]
                if (notice and "hamlet" in notice.lower()) or not choices:
                    # Nothing to choose; PMIS disables the field in this case.
                    notes.append("village has no hamlets")
                    self._report("Hamlet (none for village)", True, value)
                    continue

                wanted = "" if _empty(value) else str(value).strip()
                match = next((o for o in choices
                              if o.casefold() == wanted.casefold()), None)
                if match:
                    value = match
                elif not wanted or wanted.upper() in \
                        config.SHG_NO_HAMLET_VALUES:
                    value = self.hamlet_fallback(choices)
                    notes.append(f"no hamlet given - used '{value}'")
                else:
                    value = self.hamlet_fallback(choices)
                    notes.append(f"hamlet '{wanted}' not in PMIS - "
                                 f"used '{value}'")

            if _empty(value):
                failed.append(level)
                self._report(f"Geo {level}", False, value)
                break

            xpath = SHG_LEVEL_XPATHS[level]
            if self.current_selection(xpath).casefold() == \
                    str(value).strip().casefold():
                self._cascade_set[level] = value
                self._report(f"Geo {level} (kept)", True, value)
                continue
            ok = self.pick(xpath, value, timeout=config.PARTICIPANT_GEO_TIMEOUT)
            self.consume_alert()          # a pick can raise an info dialog
            self._report(f"Geo {level}", ok, value)
            if ok:
                self._cascade_set[level] = value
            if not ok:
                failed.append(level)
                # Say WHICH name PMIS does not have, and what is close to it -
                # "Not filled: Geo Village" alone sends you hunting.
                notes.append(near_miss(self, level, value))
                rest = SHG_LEVELS[SHG_LEVELS.index(level) + 1:]
                if rest:
                    self.log(f"      skipping {', '.join(rest)} - they depend "
                             f"on {level}")
                break
        return failed, notes

    # The page's own duplicate check, asked directly. #ShgName carries
    # onchange="checkSHGNameExist()", which calls this same endpoint and, on a
    # hit, shows #shgNameError and disables Save - but send_keys never blurs
    # the field, so driving the DOM would mean guessing when the AJAX landed.
    _NAME_CHECK_JS = """
        var name = arguments[0], done = arguments[arguments.length - 1];
        fetch('/SHG/IsSHGNameExists?name=' + encodeURIComponent(name),
              {credentials: 'same-origin'})
            .then(function (r) {
                if (!r.ok) { done(null); return; }      // endpoint is down
                return r.json().then(function (j) { done(!!j.exists); });
            })
            .catch(function () { done(null); });
    """

    # None = not asked yet, False = the endpoint is broken, stop asking.
    _name_check = None

    def shg_name_taken(self, name, retry=True):
        """True if PMIS already holds this SHG name, False if it does not,
        None if the check could not answer.

        A signed-out session answers the sign-in page instead of JSON, which
        reads as None - so reload the form once (that round trip re-auths)
        and ask again before giving up.
        """
        if _empty(name):
            return False
        try:
            self.driver.set_script_timeout(config.SHG_NAME_CHECK_TIMEOUT)
            answer = self.driver.execute_async_script(
                self._NAME_CHECK_JS, str(name).strip())
        except Exception:
            answer = None
        if answer is None and retry and self._base_url:
            self.log("      name check gave no answer - reloading the form...")
            try:
                self.open_create_form(self._base_url)
            except Exception:
                return None
            return self.shg_name_taken(name, retry=False)
        return answer

    def fill_shg(self, row):
        """Fill the Create SHG form. Returns (failed, notes)."""
        cols = config.SHG_COLUMNS
        F = config.SHG_FORM
        failed, notes = [], []

        name = row.get(cols["name"], "")
        ok = self.type_text(F["name"], name)
        self._report("SHG Name", ok, name)
        if not ok:
            failed.append("SHG Name")
            return failed, notes
        if self._name_check is not False:
            taken = self.shg_name_taken(name)
            if taken is None:
                self._name_check = False
                self.log("      note: PMIS's duplicate-name check is not "
                         "answering - Save will catch duplicates instead.")
            elif taken:
                # Nothing below this matters; PMIS will reject the form.
                self._name_check = True
                self._report("SHG Name", False, f"{name} (already in PMIS)")
                return [NAME_EXISTS], notes
            else:
                self._name_check = True

        formed = format_date(row.get(cols["date"], ""))
        ok = self.type_text(F["date"], formed)
        self._report("Date Of Formation", ok, formed)
        if not ok:
            failed.append("Date Of Formation")

        ok = self.pick(F["status"], config.SHG_DEFAULT_STATUS)
        self._report("SHG Status", ok, config.SHG_DEFAULT_STATUS)
        if not ok:
            failed.append("SHG Status")

        geo_failed, geo_notes = self.select_cascade(row, cols)
        failed.extend(f"Geo {g}" for g in geo_failed)
        notes.extend(geo_notes)

        # Commodities: match the checkbox by its visible label.
        commodity = row.get(cols["commodity"], "")
        if not _empty(commodity):
            hit = False
            for el, value, label in self.checkbox_controls():
                if label and label.strip().casefold() == \
                        str(commodity).strip().casefold():
                    try:
                        if not el.is_selected():
                            el.click()
                        hit = el.is_selected()
                    except Exception:
                        hit = False
                    self._report("Commodity", hit,
                                 f"{commodity} [value={value}]")
                    break
            if not hit:
                failed.append("Commodity")
                self._report("Commodity", False, commodity)

        if not failed:
            # Last thing before Save: make sure no level drifted back.
            refixed = self.verify_cascade()
            if refixed:
                notes.append("re-selected " + ", ".join(sorted(set(refixed))))
            still = self.drifted_levels()
            if still:
                failed.extend(f"Geo {lvl}" for lvl, _ in still)
        return failed, notes

    _SELECTED_JS = """
        return arguments[0].map(function (id) {
            var el = document.getElementById(id);
            if (!el || el.selectedIndex < 0) return "";
            return (el.options[el.selectedIndex].text || "").trim();
        });
    """

    def selected_levels(self, levels):
        """Visible text of several cascade dropdowns, in one round trip.

        The settle loop polls six levels several times a second; one by one
        that is a dozen WebDriver calls per poll, and they dominate the run.
        """
        try:
            texts = self.driver.execute_script(
                self._SELECTED_JS, [SHG_LEVEL_IDS[lvl] for lvl in levels])
            return dict(zip(levels, texts))
        except Exception:
            return {lvl: self.current_selection(SHG_LEVEL_XPATHS[lvl])
                    for lvl in levels}

    def drifted_levels(self):
        """Levels whose selection no longer matches what we chose."""
        now = self.selected_levels(list(self._cascade_set))
        return [(lvl, val) for lvl, val in self._cascade_set.items()
                if (now.get(lvl) or "").casefold()
                != str(val).strip().casefold()]

    def verify_cascade(self):
        """Hold the geography cascade until it stops drifting back to "Select".

        PMIS rebuilds a dropdown by replacing its whole option list when the
        AJAX for it lands (displayFieldBinder does select.html(...)), which
        wipes the selection. We choose a level as soon as its option appears,
        so a reply that arrives later resets it - District goes back to
        "Select", the form fails validation on Save, the "Are You Sure?"
        dialog never appears and PMIS reports nothing at all.

        Those requests do not go through Pace, so there is no idle signal to
        wait on. Only stability proves every reply has landed: re-pick what
        drifted (top-down, since a parent clears its children) and require the
        whole cascade to read correct for SHG_CASCADE_SETTLE seconds.

        Returns the levels that had to be re-picked.
        """
        deadline = time.time() + config.SHG_CASCADE_TIMEOUT
        fixed = []
        stable_since = None
        while time.time() < deadline:
            drifted = self.drifted_levels()
            if not drifted:
                if stable_since is None:
                    stable_since = time.time()
                elif time.time() - stable_since >= config.SHG_CASCADE_SETTLE:
                    return fixed
                time.sleep(0.25)
                continue

            stable_since = None
            for lvl, val in drifted:      # _cascade_set is ordered top-down
                self.log(f"      {lvl} was reset by PMIS - selecting "
                         f"'{val}' again")
                self.pick(SHG_LEVEL_XPATHS[lvl], val,
                          timeout=config.PARTICIPANT_GEO_TIMEOUT)
                self.consume_alert()
                fixed.append(lvl)
        return fixed

    # -------------------------------------------------------------- saving
    _SAVE_STATE_JS = """
        var out = {url: location.href, success: false, errors: []};
        out.success = document.querySelectorAll(arguments[1]).length > 0;
        arguments[0].forEach(function (sel) {
            document.querySelectorAll(sel).forEach(function (e) {
                if (e.getClientRects().length === 0) return;
                var t = (e.innerText || '').trim();
                if (t && out.errors.indexOf(t) === -1) out.errors.push(t);
            });
        });
        return out;
    """

    def save_state(self):
        """Where the page is, whether it toasted success, and any error text -
        all in one round trip, because the save wait asks on every poll."""
        try:
            st = self.driver.execute_script(
                self._SAVE_STATE_JS, list(ERROR_SELECTORS), SUCCESS_SELECTOR)
        except Exception:
            return {"url": "", "success": False, "errors": ""}
        errors = " | ".join(" ".join(e.split())
                            for e in (st.get("errors") or []))
        return {"url": st.get("url") or "",
                "success": bool(st.get("success")),
                "errors": errors[:400]}

    def collect_errors(self):
        """Visible validation text on the page, joined into one line."""
        return self.save_state()["errors"]

    def save(self):
        """Click Save, answer "Are You Sure?" with Yes, and report (ok, msg).

        The page's own contract, read off the live form:
          - Save is disabled while the typed SHG name already exists
          - an invalid form (or no commodity ticked) never reaches the confirm
          - Yes posts it: success redirects to the SHG list, failure stays on
            the form and toasts the server's message
        The message is what gets written into PMIS-Update, so it carries the
        real error text.
        """
        btns = self.driver.find_elements(By.XPATH, config.SHG_FORM["save"])
        if not btns:
            return False, "Save button not found"
        if not btns[0].is_enabled():
            # The page only disables Save for a duplicate name.
            return False, (self.collect_errors() or NAME_EXISTS)

        before = self.driver.current_url
        try:
            btns[0].click()
        except Exception as e:
            return False, f"could not click Save: {e}"

        # "Are You Sure?" -> Yes. No dialog means the form rejected itself
        # and nothing was posted.
        question = self.consume_alert(timeout=config.SHG_CONFIRM_TIMEOUT)
        if not question:
            return False, (self.collect_errors() or
                           "no confirmation dialog - the form was not submitted")
        self.log(f"      {question} -> Yes")

        # Success redirects to the SHG list; failure stays put and toasts.
        # Only error selectors are watched, so the success toast cannot
        # win this race.
        state = {}

        def decided(_):
            state.update(self.save_state())
            return (state["url"] != before or state["success"]
                    or state["errors"])

        try:
            WebDriverWait(self.driver, config.SHG_SAVE_TIMEOUT).until(decided)
        except TimeoutException:
            pass
        time.sleep(config.SHG_SAVE_SETTLE)

        # Either proof is enough: the success toast fires on ajax.done, and
        # the redirect to the SHG list follows it.
        after = self.save_state()
        if state.get("success") or after["success"] or \
                config.SHG_LIST_PATH.strip("/").lower() in after["url"].lower():
            return True, "saved"
        return False, (after["errors"] or
                       "save not confirmed (still on the form)")


# ---------------------------------------------------------------- runner
def run_shg_creation(driver, excel_path, base_url, username, password,
                     sheet=None, status_col=None, log=print, stop_event=None):
    sheet = sheet or config.SHG_SHEET
    auto_save = bool(settings.get("shg_auto_save"))
    bot = ShgForm(driver, log=log, stop_event=stop_event)
    cols = config.SHG_COLUMNS

    rows = read_shg_sheet(excel_path, sheet)      # one parse per run
    df = load_shgs(rows)
    taken = known_names(rows)
    log(f"{len(df)} {config.SHG_TYPE_ALLOWED} SHG(s) marked "
        f"'{config.LOAD_MARKER}' in '{sheet}'.")
    if taken:
        log(f"{len(taken)} SHG name(s) already in PMIS - those rows are skipped.")
    log(f"Auto Save is {'ON - SHGs WILL be created' if auto_save else 'OFF - nothing will be saved'}.")
    if not df:
        log(f"Nothing to do - mark the rows you want with "
            f"'{config.LOAD_MARKER}' in '{config.SHG_STATUS_COL}'.")
        return 0

    bot.ensure_logged_in(username, password, base_url)

    first = df[0]
    ok, detail = bot.verify_partner(base_url,
                                    first.get(cols["partner"], ""),
                                    first.get(cols["workorder"], ""))
    if not ok:
        log("")
        log("=" * 60)
        log(f"PARTNER CHECK FAILED: {detail}")
        log("You are signed into a different partner/project than the sheet "
            "describes. Nothing was entered.")
        log("=" * 60)
        return 0
    log(f"Partner check OK - {detail}")

    # ---- pass 1: the names ----------------------------------------------
    # Every name is settled before anything is created, so the sheet says
    # what PMIS already holds before a single form is filled.
    bot.open_create_form(base_url)
    log("")
    log(f"Checking {len(df)} SHG name(s) against PMIS...")

    names = {}                  # excel row -> verdict text
    todo = []                   # rows whose name is free
    checked = True
    try:
        for row in df:
            bot._check_stop()
            excel_row = int(row["_row"])
            name = data_file.text(row, cols["name"])
            key = name.casefold()

            if key in taken:    # already in PMIS, or earlier in this batch
                names[excel_row] = NAME_EXISTS
                log(f"  row {excel_row}: {name} - {NAME_EXISTS}")
                continue

            exists = bot.shg_name_taken(name)
            if exists is None:
                log("")
                log("  ! PMIS's duplicate-name check is not answering - "
                    "skipping the name pass.")
                log("    Duplicates will be caught as each SHG is saved.")
                checked = False
                todo = [r for r in df
                        if data_file.text(r, cols["name"]).casefold()
                        not in taken]
                break
            if exists:
                taken.add(key)
                names[excel_row] = NAME_EXISTS
                log(f"  row {excel_row}: {name} - {NAME_EXISTS}")
            else:
                todo.append(row)
    except StopRequested:
        log("")
        log("Stopped by user during the name check.")
        write_status_bulk(excel_path, sheet, names, log)
        return 0

    if names:
        write_status_bulk(excel_path, sheet, names, log)
    if checked:
        log(f"  {len(names)} already in PMIS, {len(todo)} to create.")
    if not todo:
        log("")
        log("Nothing left to create.")
        return 0

    # ---- pass 2: create what is left -------------------------------------
    # Statuses are buffered: rewriting the whole .xlsx after every row costs
    # far more than the run itself. At most SHG_FLUSH_EVERY rows are ever
    # unwritten, and every exit path flushes.
    pending = {}

    def record(excel_row, text):
        pending[int(excel_row)] = text
        if len(pending) >= config.SHG_FLUSH_EVERY:
            flush()

    def flush():
        if pending:
            write_status_bulk(excel_path, sheet, pending, log)
            pending.clear()

    done = 0
    try:
        for row in todo:
            bot._check_stop()
            excel_row = int(row["_row"])
            name = data_file.text(row, cols["name"])
            log("")
            log(f"Excel row {excel_row}: {name}")

            if name.casefold() in taken:
                log(f"  skipped - {NAME_EXISTS}")
                record(excel_row, NAME_EXISTS)
                continue

            bot.open_create_form(base_url,
                                 expect_workorder=row.get(cols["workorder"], ""))
            failed, notes = bot.fill_shg(row)

            if failed == [NAME_EXISTS]:
                log(f"  skipped - {NAME_EXISTS}")
                record(excel_row, NAME_EXISTS)
                taken.add(name.casefold())
                continue        # the next row reloads a fresh Create SHG form

            if failed:
                msg = "Not filled: " + ", ".join(failed)
                if notes:
                    msg += " (" + "; ".join(notes) + ")"
                log(f"  {msg}")
                record(excel_row, msg)
                continue

            if not auto_save:
                note = "; ".join(notes) if notes else ""
                log("  filled - NOT SAVED (Auto Save is off)"
                    + (f" [{note}]" if note else ""))
                log("=" * 60)
                log("Stopping after the first SHG so the form can be checked.")
                log("=" * 60)
                return 1

            saved, message = bot.save()
            if saved:
                status = config.SHG_DONE_MARKER
                if notes:
                    status += " (" + "; ".join(notes) + ")"
                record(excel_row, status)
                taken.add(name.casefold())
                done += 1
                log(f"  {config.SHG_DONE_MARKER}")
            else:
                record(excel_row, message)
                if NAME_EXISTS.lower() in message.lower():
                    taken.add(name.casefold())
                log(f"  NOT created - {message}")
    except StopRequested:
        log("")
        log(f"Stopped by user. {done} SHG(s) created.")
        return done
    finally:
        flush()

    log("")
    log("=" * 60)
    log(f"Done. {len(names)} already existed, {done} created.")
    log("=" * 60)
    return done


# ------------------------------------------------------- SHG geo checker
def build_shg_cascades(rows):
    """Unique geo cascades from already-loaded SHG rows, ordered so that
    consecutive rows share as long a prefix as possible (the checker then
    skips levels that are already selected)."""
    cols = config.SHG_COLUMNS
    want = [cols[k] for k in ("state", "district", "block", "gp", "village",
                              "hamlet")]
    seen, out = set(), []
    for r in rows:
        cascade = {lvl: data_file.text(r, col)
                   for lvl, col in zip(SHG_LEVELS, want)}
        # A cascade is only checkable down to Village; Hamlet is optional.
        if not all(cascade[lvl] for lvl in SHG_LEVELS[:5]):
            continue
        key = tuple(cascade[lvl] for lvl in SHG_LEVELS)
        if key in seen:
            continue
        seen.add(key)
        out.append(cascade)
    out.sort(key=lambda c: [c[lvl] for lvl in SHG_LEVELS])
    return out


def run_shg_geo_check(driver, excel_path, base_url, username, password,
                      sheet=None, status_col=None, log=print, stop_event=None):
    """Walk every unique cascade in the SHG sheet against the Create SHG page
    and write the verdict for each sheet row into PMIS-Update."""
    sheet = sheet or config.SHG_SHEET
    cols = config.SHG_COLUMNS
    bot = ShgForm(driver, log=log, stop_event=stop_event)

    # Every row, with its cascade, so results map back to all of them.
    want = [cols[k] for k in ("state", "district", "block", "gp", "village",
                              "hamlet")]
    full = load_rows_only(read_shg_sheet(excel_path, sheet))

    loc = build_shg_cascades(full)
    if not loc:
        log(f"No rows marked '{config.LOAD_MARKER}' in "
            f"'{config.SHG_STATUS_COL}' - nothing to check.")
        return 0
    log(f"{len(loc)} unique location(s) to check from '{sheet}' "
        f"({len(full)} row(s) marked '{config.LOAD_MARKER}').")

    bot.ensure_logged_in(username, password, base_url)
    bot.open_create_form(base_url)

    verdicts = {}            # cascade key -> result text
    bad, hamlet_bad = [], []
    total = len(loc)

    try:
        for i, row in enumerate(loc):
            bot._check_stop()
            key = tuple(str(row[l]) for l in SHG_LEVELS)
            trail = " > ".join(key[:5])
            log(f"#{i + 1}/{total}  {trail}")

            missing_at = None
            for level in SHG_LEVELS[:5]:          # State .. Village
                value = row[level]
                xpath = SHG_LEVEL_XPATHS[level]
                if bot.current_selection(xpath).casefold() == value.casefold():
                    continue                      # already selected
                if not bot.pick(xpath, value,
                                timeout=config.PARTICIPANT_GEO_TIMEOUT):
                    missing_at = level
                    break

            if missing_at:
                verdict = near_miss(bot, missing_at, row[missing_at])
                log(f"    NOT FOUND - {verdict}")
                verdicts[key] = verdict
                bad.append(dict(row, Missing=missing_at))
                continue

            # Selecting the village can raise "There are no Hamlets for
            # this village!" - clear it and use what it tells us.
            notice = bot.consume_alert() or bot.take_last_alert()
            no_hamlets = bool(notice) and "hamlet" in notice.lower()
            if notice:
                log(f"    PMIS says: {notice}")

            hamlet = str(row["Hamlet"]).strip()
            wants_hamlet = bool(hamlet) and \
                hamlet.upper() not in config.SHG_NO_HAMLET_VALUES

            if no_hamlets:
                # The village genuinely has none. Fine if the sheet agrees.
                if wants_hamlet:
                    # Nothing to suggest - the village has no hamlets at all.
                    log(f"    {config.NO_LOCATION} - hamlet {hamlet} "
                        f"(village has no hamlets)")
                    verdicts[key] = config.NO_LOCATION
                    hamlet_bad.append(dict(row, Missing="Hamlet"))
                    continue
                log("    ok (village has no hamlets)")
                verdicts[key] = config.GEO_OK_MARKER
                continue

            if wants_hamlet:
                available = {o.strip().casefold()
                             for o in bot.hamlet_options()}
                if hamlet.casefold() not in available:
                    verdict = near_miss(bot, "Hamlet", hamlet)
                    log(f"    {verdict}")
                    verdicts[key] = verdict
                    hamlet_bad.append(dict(row, Missing="Hamlet"))
                    continue
                log("    ok (incl. hamlet)")
            else:
                log("    ok")
            verdicts[key] = config.GEO_OK_MARKER
    except StopRequested:
        log("")
        log("Stopped by user - writing the results gathered so far.")
    except Exception as e:
        # Never throw away completed work: whatever was verified still gets
        # written to the sheet before the error is reported.
        import traceback
        log("")
        log(f"ERROR after {len(verdicts)} cascade(s): {e}")
        log(traceback.format_exc())
        log("Writing the results gathered so far before stopping.")

    # ---- map the verdicts onto every sheet row ----
    mapping = {}
    for r in full:
        key = tuple(data_file.text(r, c) for c in want)
        if key in verdicts:
            mapping[int(r["_row"])] = verdicts[key]
    written = write_status_bulk(excel_path, sheet, mapping, log)

    # ---- report ----
    log("")
    log("=" * 64)
    log(f"SHG Geo Checker: {len(verdicts)} of {total} cascade(s) checked")
    log(f"Results written to '{config.SHG_STATUS_COL}' for {written} row(s) "
        f"(rows already '{config.SHG_DONE_MARKER}' were left untouched).")
    if not bad and not hamlet_bad:
        log("Every location exists in PMIS.")
    if bad:
        log("")
        log(f"{len(bad)} cascade(s) failed. Distinct names to fix:")
        seen = set()
        for m in bad:
            k = (m["Missing"], str(m[m["Missing"]]))
            if k in seen:
                continue
            seen.add(k)
            log(f"  [{m['Missing']}] '{m[m['Missing']]}'"
                f"   in  " + " > ".join(str(m[l]) for l in SHG_LEVELS[:5]))
        log(f"  -> {len(seen)} distinct bad name(s)")
    if hamlet_bad:
        log("")
        log(f"{len(hamlet_bad)} cascade(s) have a hamlet PMIS does not know:")
        for h in sorted({str(x["Hamlet"]) for x in hamlet_bad}):
            log(f"  {h}")
        log("  (SHG creation still proceeds for these, leaving hamlet unset)")
    log("=" * 64)
    return total
