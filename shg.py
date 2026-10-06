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

import pandas as pd
from openpyxl import load_workbook

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import config
import settings
from pmis_geo import StopRequested
from participant import MemberForm, _empty, _blank

# Cascade on the SHG form - same control names as elsewhere, plus Hamlet.
SHG_LEVELS = ["State", "District", "Block", "GP", "Village", "Hamlet"]
SHG_LEVEL_XPATHS = {
    "State":    '//select[@name="ProvinceId"]',
    "District": '//select[@name="DistrictId"]',
    "Block":    '//select[@name="MunicipalityId"]',
    "GP":       '//select[@name="GrampanchayatId"]',
    "Village":  '//select[@name="VillageId"]',
    "Hamlet":   '//select[@name="HamletId"]',
}

# Where inline validation errors show up after a failed save.
ERROR_SELECTORS = [
    ".validation-summary-errors", ".field-validation-error",
    ".alert-danger", ".alert-error", ".text-danger",
    "#errorMessage", ".error-message", ".toast-error",
]


# ---------------------------------------------------------------- helpers
def format_date(value):
    """Any date cell -> 'dd-mm-yyyy' for DateOfFormationAttr."""
    if _empty(value):
        return ""
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return value.strftime("%d-%m-%Y")
    text = str(value).strip()
    # Already dd-mm-yyyy (or d-m-yyyy) - normalise the padding.
    m = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$", text)
    if m:
        return "%02d-%02d-%s" % (int(m.group(1)), int(m.group(2)), m.group(3))
    try:
        return pd.to_datetime(text, dayfirst=True).strftime("%d-%m-%Y")
    except Exception:
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


def load_shgs(excel_path, sheet=None):
    """Pending OG rows from the SHG sheet."""
    sheet = sheet or config.SHG_SHEET
    cols = config.SHG_COLUMNS
    df = pd.read_excel(excel_path, sheet_name=sheet,
                       keep_default_na=False, na_values=[])
    df["_row"] = df.index + 2

    missing = [c for c in cols.values() if c not in df.columns]
    if missing:
        raise ValueError(f"Sheet '{sheet}' is missing column(s): "
                         f"{', '.join(missing)}")

    status = config.SHG_STATUS_COL
    if status in df.columns:
        done = config.SHG_DONE_MARKER.strip().lower()
        df = df[df[status].astype(str).str.strip().str.lower() != done]

    # OG only.
    type_col = cols["type"]
    allowed = config.SHG_TYPE_ALLOWED.strip().lower()
    df = df[df[type_col].astype(str).str.strip().str.lower() == allowed]
    return df.reset_index(drop=True)


def write_status(excel_path, sheet, excel_row, text, log=print):
    """Record the outcome (done marker or error text) in PMIS-Update."""
    try:
        wb = load_workbook(excel_path)
        ws = wb[sheet]
        headers = {c.value: c.column for c in ws[1]}
        col = headers.get(config.SHG_STATUS_COL)
        if col is None:
            col = ws.max_column + 1
            ws.cell(row=1, column=col, value=config.SHG_STATUS_COL)
        ws.cell(row=int(excel_row), column=col, value=str(text)[:400])
        wb.save(excel_path)
        return True
    except PermissionError:
        log(f"  ! Could not write status for row {excel_row} - the workbook "
            f"is open in Excel.")
    except Exception as e:
        log(f"  ! Could not write status for row {excel_row}: {e}")
    return False


def write_status_bulk(excel_path, sheet, mapping, log=print,
                      protect_marker=True):
    """Write many statuses in one pass: {excel_row: text}.

    Rows already carrying the done marker are left alone by default, so a
    geo check can never wipe the record of an SHG that was actually created.
    """
    if not mapping:
        return 0
    try:
        wb = load_workbook(excel_path)
        ws = wb[sheet]
        headers = {c.value: c.column for c in ws[1]}
        col = headers.get(config.SHG_STATUS_COL)
        if col is None:
            col = ws.max_column + 1
            ws.cell(row=1, column=col, value=config.SHG_STATUS_COL)
        done = config.SHG_DONE_MARKER.strip().lower()
        written = 0
        for excel_row, text in mapping.items():
            cell = ws.cell(row=int(excel_row), column=col)
            if protect_marker and str(cell.value or "").strip().lower() \
                    .startswith(done):
                continue          # already created - keep that record
            cell.value = str(text)[:400]
            written += 1
        wb.save(excel_path)
        return written
    except PermissionError:
        log("  ! Could not write results - the workbook is open in Excel.")
    except Exception as e:
        log(f"  ! Could not write results: {e}")
    return 0


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
        except TimeoutException:
            raise RuntimeError(f"Create SHG form did not load ({url}); "
                               f"landed on {self.driver.current_url}")
        if expect_workorder:
            info = parse_partner_header(
                self.driver.find_element(By.TAG_NAME, "body").text)
            if info["workorder"] and info["workorder"].strip().lower() != \
                    str(expect_workorder).strip().lower():
                raise RuntimeError(
                    f"Workorder on the create page ({info['workorder']}) does "
                    f"not match the sheet ({expect_workorder})")

    # ------------------------------------------------------------- filling
    def hamlet_options(self):
        try:
            from selenium.webdriver.support.select import Select
            sel = Select(self.driver.find_element(
                By.XPATH, SHG_LEVEL_XPATHS["Hamlet"]))
            return [o.text.strip() for o in sel.options]
        except Exception:
            return []

    def select_cascade(self, row, cols):
        """State -> Hamlet. Returns (failed_levels, notes)."""
        failed, notes = [], []
        level_to_col = {
            "State": cols["state"], "District": cols["district"],
            "Block": cols["block"], "GP": cols["gp"],
            "Village": cols["village"], "Hamlet": cols["hamlet"],
        }
        for level in SHG_LEVELS:
            self._check_stop()
            value = row.get(level_to_col[level], "")

            if level == "Hamlet":
                # Only set a hamlet that actually exists; otherwise leave it.
                if _empty(value) or str(value).strip().upper() in \
                        config.SHG_NO_HAMLET_VALUES:
                    notes.append("hamlet not provided")
                    self._report("Hamlet (skipped)", True, value)
                    continue
                available = {o.strip().casefold()
                             for o in self.hamlet_options()}
                if str(value).strip().casefold() not in available:
                    notes.append(f"hamlet '{str(value).strip()}' not in PMIS")
                    self._report("Hamlet (not in list)", True, value)
                    continue

            if _empty(value):
                failed.append(level)
                self._report(f"Geo {level}", False, value)
                break

            xpath = SHG_LEVEL_XPATHS[level]
            if self.current_selection(xpath).casefold() == \
                    str(value).strip().casefold():
                self._report(f"Geo {level} (kept)", True, value)
                continue
            ok = self.pick(xpath, value, timeout=config.PARTICIPANT_GEO_TIMEOUT)
            self._report(f"Geo {level}", ok, value)
            if not ok:
                failed.append(level)
                rest = SHG_LEVELS[SHG_LEVELS.index(level) + 1:]
                if rest:
                    self.log(f"      skipping {', '.join(rest)} - they depend "
                             f"on {level}")
                break
        return failed, notes

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
        return failed, notes

    # -------------------------------------------------------------- saving
    def collect_errors(self):
        """Visible validation text on the page, joined into one line."""
        js = """
            var out = [];
            arguments[0].forEach(function(sel){
                document.querySelectorAll(sel).forEach(function(e){
                    if (e.getClientRects().length === 0) return;
                    var t = (e.innerText || '').trim();
                    if (t && out.indexOf(t) === -1) out.push(t);
                });
            });
            return out;
        """
        try:
            found = self.driver.execute_script(js, list(ERROR_SELECTORS)) or []
        except Exception:
            return ""
        text = " | ".join(" ".join(f.split()) for f in found)
        return text[:400]

    def save(self):
        """Click Save and report (ok, message). The message is what gets
        written into PMIS-Update, so it carries the real error text."""
        btns = self.driver.find_elements(By.XPATH, config.SHG_FORM["save"])
        if not btns:
            return False, "Save button not found"
        before = self.driver.current_url
        try:
            btns[0].click()
        except Exception as e:
            return False, f"could not click Save: {e}"

        # A dialog may carry the message ("SHG Name already exists").
        alert_text = ""
        try:
            WebDriverWait(self.driver, 6).until(EC.alert_is_present())
            alert = self.driver.switch_to.alert
            alert_text = (alert.text or "").strip()
            alert.accept()
        except TimeoutException:
            pass

        self.wait_not_busy(timeout=20)
        time.sleep(config.SHG_SAVE_SETTLE)

        inline = self.collect_errors()
        message = " | ".join(x for x in (alert_text, inline) if x).strip()

        looks_bad = bool(inline) or bool(
            re.search(r"exist|error|required|invalid|duplicate|fail",
                      alert_text, re.I))
        if looks_bad:
            return False, message or "save rejected (no message shown)"

        if self.driver.current_url != before:
            return True, alert_text or "saved"
        # Still on the form with nothing reported - treat as unconfirmed.
        return False, message or "save not confirmed (still on the form)"


# ---------------------------------------------------------------- runner
def run_shg_creation(driver, excel_path, base_url, username, password,
                     sheet=None, status_col=None, log=print, stop_event=None):
    sheet = sheet or config.SHG_SHEET
    auto_save = bool(settings.get("shg_auto_save"))
    bot = ShgForm(driver, log=log, stop_event=stop_event)
    cols = config.SHG_COLUMNS

    df = load_shgs(excel_path, sheet)
    log(f"{len(df)} pending {config.SHG_TYPE_ALLOWED} SHG(s) in '{sheet}'.")
    log(f"Auto Save is {'ON - SHGs WILL be created' if auto_save else 'OFF - nothing will be saved'}.")
    if df.empty:
        return 0

    bot.ensure_logged_in(username, password, base_url)

    first = df.iloc[0]
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

    done = 0
    try:
        for _, row in df.iterrows():
            bot._check_stop()
            excel_row = int(row["_row"])
            name = row.get(cols["name"], "")
            log("")
            log(f"Excel row {excel_row}: {name}")

            bot.open_create_form(base_url,
                                 expect_workorder=row.get(cols["workorder"], ""))
            failed, notes = bot.fill_shg(row)

            if failed:
                msg = "Not filled: " + ", ".join(failed)
                if notes:
                    msg += " (" + "; ".join(notes) + ")"
                log(f"  {msg}")
                write_status(excel_path, sheet, excel_row, msg, log)
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
                write_status(excel_path, sheet, excel_row, status, log)
                done += 1
                log(f"  created - {message}")
            else:
                write_status(excel_path, sheet, excel_row, message, log)
                log(f"  NOT created - {message}")
    except StopRequested:
        log("")
        log(f"Stopped by user. {done} SHG(s) created.")
        return done

    log("")
    log("=" * 60)
    log(f"Done. {done} SHG(s) created.")
    log("=" * 60)
    return done


# ------------------------------------------------------- SHG geo checker
def build_shg_cascades(excel_path, sheet=None):
    """Unique geo cascades from the SHG sheet, ordered so that consecutive
    rows share as long a prefix as possible (the checker then skips levels
    that are already selected)."""
    sheet = sheet or config.SHG_SHEET
    cols = config.SHG_COLUMNS
    want = [cols[k] for k in ("state", "district", "block", "gp", "village",
                              "hamlet")]
    df = pd.read_excel(excel_path, sheet_name=sheet,
                       keep_default_na=False, na_values=[])
    missing = [c for c in want if c not in df.columns]
    if missing:
        raise ValueError(f"Sheet '{sheet}' is missing column(s): "
                         f"{', '.join(missing)}")

    loc = df[want].copy()
    loc.columns = SHG_LEVELS
    for c in SHG_LEVELS:
        loc[c] = loc[c].astype(str).str.strip()
    # A cascade is only checkable down to Village; Hamlet is optional.
    for c in ["State", "District", "Block", "GP", "Village"]:
        loc = loc[loc[c] != ""]
    loc = loc.drop_duplicates(subset=SHG_LEVELS)
    loc = loc.sort_values(SHG_LEVELS).reset_index(drop=True)
    return loc


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
    full = pd.read_excel(excel_path, sheet_name=sheet,
                         keep_default_na=False, na_values=[])
    full["_row"] = full.index + 2
    missing_cols = [c for c in want if c not in full.columns]
    if missing_cols:
        raise ValueError(f"Sheet '{sheet}' is missing column(s): "
                         f"{', '.join(missing_cols)}")
    for c in want:
        full[c] = full[c].astype(str).str.strip()

    loc = build_shg_cascades(excel_path, sheet)
    if loc.empty:
        log("No cascades to check in the SHG sheet.")
        return 0
    log(f"{len(loc)} unique location(s) to check from '{sheet}' "
        f"({len(full)} rows).")

    bot.ensure_logged_in(username, password, base_url)
    bot.open_create_form(base_url)

    verdicts = {}            # cascade key -> result text
    bad, hamlet_bad = [], []
    total = len(loc)

    try:
        for i, row in loc.iterrows():
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
                log(f"    NOT FOUND - {missing_at}: {row[missing_at]}")
                verdicts[key] = f"{missing_at} not found: {row[missing_at]}"
                bad.append(dict(row, Missing=missing_at))
                continue

            hamlet = str(row["Hamlet"]).strip()
            if hamlet and hamlet.upper() not in config.SHG_NO_HAMLET_VALUES:
                available = {o.strip().casefold()
                             for o in bot.hamlet_options()}
                if hamlet.casefold() not in available:
                    log(f"    Hamlet not found: {hamlet}")
                    verdicts[key] = f"Hamlet not found: {hamlet}"
                    hamlet_bad.append(dict(row, Missing="Hamlet"))
                    continue
                log("    ok (incl. hamlet)")
            else:
                log("    ok")
            verdicts[key] = config.SHG_GEO_OK_MARKER
    except StopRequested:
        log("")
        log("Stopped by user - writing the results gathered so far.")

    # ---- map the verdicts onto every sheet row ----
    mapping = {}
    for _, r in full.iterrows():
        key = tuple(str(r[c]) for c in want)
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
