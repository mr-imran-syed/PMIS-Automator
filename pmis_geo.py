"""PMIS GEO automation logic (browser-agnostic).

Refactored from the PMIS_autobot GEO notebook into a reusable class so it can be
driven from a GUI. Holds all the hardening we developed:
  - wait_options_stable : wait for AJAX-loaded dropdowns to settle before judging
  - find_and_select     : settle + retry match (tolerates server lag; avoids dupes)
  - wait_modal_gone     : recover if the add-modal fails to auto-close
  - click_add           : click 'Add' only when enabled and unobscured
  - per-row State (from the Excel 'State' column)
  - PMIS='Checked' write-back to the source workbook for resumable runs
"""

import time
import random

import pandas as pd
from openpyxl import load_workbook

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
)

# NOTE: environment URLs live in config.ENVIRONMENTS (single source of truth).

REQUIRED_COLUMNS = ["State", "District", "Block", "GP", "Village"]


class StopRequested(Exception):
    """Raised internally to unwind the loop when the user presses Stop."""


class GeoBot:
    def __init__(self, driver, log=print, stop_event=None):
        self.driver = driver
        self.log = log                     # callable(str) for progress messages
        self.stop_event = stop_event       # threading.Event or None
        self.wait = WebDriverWait(driver, 10)
        self.alert = Alert(driver)
        # openpyxl write-back state
        self._wb = None
        self._ws = None
        self._pmis_col = None
        self._geo_file = None
        self._sheet = None
        self._status_col = "PMIS"
        self._creds = None          # kept so a signed-out page can re-auth
        self._base_url = None

    # ------------------------------------------------------------------ utils
    def _check_stop(self):
        if self.stop_event is not None and self.stop_event.is_set():
            raise StopRequested()

    def rest(self, a=0.3, b=0.7):
        time.sleep(random.uniform(a, b))

    # ------------------------------------------------------------------ login
    # --- sign-in helpers ------------------------------------------------
    def signin_form_present(self):
        return bool(self.driver.find_elements(By.XPATH, '//input[@name="Email"]'))

    def is_signin_page(self):
        """True when the browser is sitting on the sign-in page.

        PMIS bounces unauthenticated requests to
        /SignIn/Index?ReturnUrl=<the page you asked for>.
        """
        try:
            if "signin" in self.driver.current_url.lower():
                return True
        except Exception:
            return False
        return self.signin_form_present()

    def _submit_login_form(self, username, password):
        """Fill and submit whatever sign-in form is on the current page."""
        self.wait.until(EC.visibility_of_element_located(
            (By.XPATH, '//input[@name="Email"]')))
        email = self.driver.find_element(By.XPATH, '//input[@name="Email"]')
        email.clear()
        email.click()
        email.send_keys(username)
        self.rest()

        pw = self.driver.find_element(By.XPATH, '//input[@name="Password"]')
        pw.clear()
        pw.click()
        pw.send_keys(password)
        self.rest()

        self.driver.find_element(By.XPATH, '//input[@id="SubmitButton"]').click()

    def _open_signin_page(self, base_url):
        """Get to the sign-in form, however this environment exposes it."""
        d = self.driver
        base = base_url.rstrip("/")
        d.get(base_url)
        if self.signin_form_present():
            return
        # Home page usually has a Sign In link.
        links = d.find_elements(By.XPATH, '//a[@href="/SignIn"]')
        if links:
            links[0].click()
            try:
                self.wait.until(EC.visibility_of_element_located(
                    (By.XPATH, '//input[@name="Email"]')))
                return
            except TimeoutException:
                pass
        # Fall back to the known sign-in routes.
        for path in ("SignIn/Index", "SignIn"):
            d.get(f"{base}/{path}")
            if self.signin_form_present():
                return
        raise RuntimeError(
            f"Could not find the sign-in form at {base_url}. Check the "
            f"environment URL.")

    def login(self, username, password, base_url):
        self._creds = (username, password)
        self._base_url = base_url

        self._open_signin_page(base_url)
        self._submit_login_form(username, password)

        # Success => we leave the sign-in page. Bad credentials => we stay.
        try:
            WebDriverWait(self.driver, 20).until(
                lambda d: not self.is_signin_page())
        except TimeoutException:
            raise RuntimeError(
                "Login failed - check the username/password and environment.")
        self.log(f"Logged in as {username}")

    def ensure_logged_in(self, username, password, base_url):
        """Sign in only when the session is not already authenticated.

        Lets a reused browser skip the whole sign-in round trip. Anything
        opened afterwards still calls reauth_if_signed_out(), so a session
        that has actually expired still recovers.
        """
        self._creds = (username, password)
        self._base_url = base_url
        try:
            current = self.driver.current_url or ""
        except Exception:
            current = ""
        if base_url.rstrip("/") not in current:
            self.driver.get(base_url)
        if self.is_signin_page():
            self.login(username, password, base_url)
        else:
            self.log(f"Already signed in - reusing the session ({username}).")

    def reauth_if_signed_out(self):
        """If the current page bounced to sign-in, sign in again.

        The site's ReturnUrl normally lands us back on the requested page.
        Returns True if a re-login was performed.
        """
        if not self.is_signin_page():
            return False
        if not self._creds:
            raise RuntimeError("Signed out, and no credentials to retry with.")
        self.log("    not authenticated for this page - signing in again...")
        self._submit_login_form(*self._creds)
        try:
            WebDriverWait(self.driver, 20).until(
                lambda d: not self.is_signin_page())
        except TimeoutException:
            raise RuntimeError(
                "Could not re-authenticate - the account may not have access "
                "to this page.")
        return True

    def goto_geolocation(self, base_url):
        self.driver.get(base_url.rstrip("/") + "/geolocation")
        # Confirm the geolocation form loaded (also catches a bounce to sign-in).
        try:
            self.wait.until(EC.presence_of_element_located(
                (By.XPATH, '//select[@name="ProvinceId"]')))
        except TimeoutException:
            raise RuntimeError(
                "Could not open the geolocation page - are you logged in with an "
                "admin account that has access?")

    # -------------------------------------------------------- selenium helpers
    def save_button(self):
        self.driver.find_element(By.XPATH, '//input[@id="SubmitButton"]').click()
        self.rest()

    def wait_options_stable(self, select_xpath, timeout=10, settle=0.6,
                            lead=0.5, poll=0.15):
        """Wait until a dependent dropdown finishes loading via AJAX (its option
        count stops changing for `settle`s), so an empty list is trustworthy."""
        time.sleep(lead)
        deadline = time.time() + timeout
        prev = None
        stable_since = None
        while time.time() < deadline:
            try:
                el = self.driver.find_element(By.XPATH, select_xpath)
                count = len(el.find_elements(By.TAG_NAME, "option"))
            except (NoSuchElementException, StaleElementReferenceException):
                prev = None
                stable_since = None
                time.sleep(poll)
                continue
            if count == prev:
                if stable_since is None:
                    stable_since = time.time()
                elif time.time() - stable_since >= settle:
                    break
            else:
                prev = count
                stable_since = None
            time.sleep(poll)
        return Select(self.driver.find_element(By.XPATH, select_xpath))

    def find_and_select(self, select_xpath, name, settle_timeout=10,
                        retry=6, poll=0.3, lead=None, settle=None):
        """Select the option matching `name` (trimmed, case-insensitive); True on
        success, False only if it never appears. Settles the list, then retries
        for `retry`s to tolerate server lag on a just-created item."""
        target = str(name).strip().casefold()
        stable_kw = {}
        if lead is not None:
            stable_kw["lead"] = lead
        if settle is not None:
            stable_kw["settle"] = settle
        self.wait_options_stable(select_xpath, timeout=settle_timeout,
                                 **stable_kw)
        deadline = time.time() + retry
        while True:
            try:
                sel = Select(self.driver.find_element(By.XPATH, select_xpath))
                for opt in sel.options:
                    if opt.text.strip().casefold() == target:
                        sel.select_by_visible_text(opt.text)
                        return True
            except (NoSuchElementException, StaleElementReferenceException):
                pass
            if time.time() >= deadline:
                return False
            time.sleep(poll)

    def wait_modal_gone(self, timeout=10):
        """Ensure the shared add-modal (#pop-up-div) is dismissed. If it stalls
        open, close it ourselves (button, then JS force-hide) instead of raising."""
        modal_xpath = '//div[@id="pop-up-div"]'
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.invisibility_of_element_located((By.XPATH, modal_xpath)))
            return
        except TimeoutException:
            pass
        for xp in ('//div[@id="pop-up-div"]//button[@data-dismiss="modal"]',
                   '//div[@id="pop-up-div"]//button[contains(@class,"close")]'):
            try:
                b = self.driver.find_element(By.XPATH, xp)
                if b.is_displayed():
                    b.click()
                    break
            except (NoSuchElementException, StaleElementReferenceException):
                continue
        try:
            WebDriverWait(self.driver, 3).until(
                EC.invisibility_of_element_located((By.XPATH, modal_xpath)))
            return
        except TimeoutException:
            pass
        try:
            self.driver.execute_script(
                "var m=document.getElementById('pop-up-div');"
                "if(m){m.classList.remove('in');m.style.display='none';}"
                "document.querySelectorAll('.modal-backdrop')"
                ".forEach(function(e){e.remove();});"
                "document.body.classList.remove('modal-open');")
        except Exception:
            pass
        self.log("  ! add-modal did not auto-close; forced it shut and continued.")

    def click_add(self, add_xpath, timeout=10):
        WebDriverWait(self.driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, add_xpath))).click()

    # ---------------------------------------------------------- cascade levels
    def select_state(self, state_name):
        state = Select(self.driver.find_element(
            By.XPATH, '//select[@name="ProvinceId"]'))
        WebDriverWait(self.driver, 10).until(lambda d: len(state.options) > 1)
        state.select_by_visible_text(state_name)

    def select_district(self, district_name):
        district = Select(self.driver.find_element(
            By.XPATH, '//select[@name="DistrictId"]'))
        WebDriverWait(self.driver, 10).until(lambda d: len(district.options) > 1)
        district.select_by_visible_text(district_name)

    def select_block(self, block_name):
        if not pd.notna(block_name):
            return None
        self.wait.until(EC.visibility_of_element_located(
            (By.XPATH, '//select[@name="MunicipalityId"]')))
        if self.find_and_select('//select[@name="MunicipalityId"]', block_name):
            self.rest()
            return None
        try:
            self.click_add('//a[@id="addMunicipality"]')
            self.wait.until(EC.visibility_of_element_located(
                (By.XPATH, '//input[@id="SubmitButton"]')))
            box = self.driver.find_element(
                By.XPATH, '//div[@class="modal-header"]//input[@id="Name"]')
            box.click()
            box.send_keys(str(block_name))
            self.rest()
            self.save_button()
            WebDriverWait(self.driver, 3).until(EC.alert_is_present())
            self.rest()
            self.alert.accept()
            self.wait_modal_gone()
            self.log(f"Added Block: {block_name}")
            self.rest()
            return "New Block"
        except (NoSuchElementException, TimeoutException):
            self.log(f"  ! could not add Block: {block_name}")
            return None

    def select_gp(self, gp_name):
        if not pd.notna(gp_name):
            return None
        self.wait.until(EC.visibility_of_element_located(
            (By.XPATH, '//select[@name="GrampanchayatId"]')))
        if self.find_and_select('//select[@name="GrampanchayatId"]', gp_name):
            self.rest()
            return None
        try:
            self.click_add('//a[@id="addGrampanchayat"]')
            self.wait.until(EC.visibility_of_element_located(
                (By.XPATH, '//input[@id="SubmitButton"]')))
            box = self.driver.find_element(
                By.XPATH, '//div[@class="modal-header"]//input[@id="Name"]')
            box.click()
            box.send_keys(str(gp_name))
            self.rest()
            self.save_button()
            WebDriverWait(self.driver, 3).until(EC.alert_is_present())
            self.rest()
            self.alert.accept()
            self.wait_modal_gone()
            self.log(f"Added GP: {gp_name}")
            self.rest()
            return "New GP"
        except (NoSuchElementException, TimeoutException):
            self.log(f"  ! could not add GP: {gp_name}")
            return None

    def select_village(self, village_name):
        if not pd.notna(village_name):
            return None
        self.wait.until(EC.visibility_of_element_located(
            (By.XPATH, '//select[@name="VillageId"]')))
        if self.find_and_select('//select[@name="VillageId"]', village_name):
            return None
        try:
            self.click_add('//a[@id="addVillage"]')
            self.wait.until(EC.visibility_of_element_located(
                (By.XPATH, '//input[@id="SubmitButton"]')))
            box = self.driver.find_element(
                By.XPATH, '//div[@class="modal-header"]//input[@id="Name"]')
            box.click()
            box.send_keys(str(village_name))
            self.save_button()
            WebDriverWait(self.driver, 3).until(EC.alert_is_present())
            self.rest()
            self.alert.accept()
            self.wait_modal_gone()
            self.log(f"Added Village: {village_name}")
            return "New Village"
        except (NoSuchElementException, TimeoutException):
            self.log(f"  ! could not add Village: {village_name}")
            return None

    # ------------------------------------------------------------- excel + run
    def load_geo_data(self, geo_file, sheet, status_col="PMIS"):
        df = pd.read_excel(geo_file, sheet_name=sheet)
        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            raise ValueError(
                f"Sheet '{sheet}' is missing required column(s): "
                f"{', '.join(missing)}")
        if status_col not in df.columns:
            df[status_col] = pd.NA
        df = df[REQUIRED_COLUMNS + [status_col]]
        df = df.reset_index().rename(columns={"index": "src_row"})
        return df[df[status_col] != "Checked"].reset_index(drop=True)

    def _open_wb(self, geo_file, sheet):
        self._geo_file = geo_file
        self._sheet = sheet
        self._wb = load_workbook(geo_file)
        self._ws = self._wb[sheet]
        headers = {c.value: c.column for c in self._ws[1]}
        self._pmis_col = headers.get(self._status_col)
        if self._pmis_col is None:
            self._pmis_col = self._ws.max_column + 1
            self._ws.cell(row=1, column=self._pmis_col, value=self._status_col)

    def mark_checked(self, src_row):
        if self._ws is None:
            self._open_wb(self._geo_file, self._sheet)
        self._ws.cell(row=int(src_row) + 2, column=self._pmis_col, value="Checked")
        try:
            self._wb.save(self._geo_file)
        except PermissionError:
            self.log(f"  ! Could not save Excel (is it open?). Row {int(src_row)} "
                     f"not marked - close Excel to enable resume.")

    def run(self, excel_path, base_url, username, password,
            sheet="Geographical Details", status_col="PMIS"):
        """Full run: log in, open the geolocation page, and process every
        not-yet-Checked row. Returns the number of rows processed."""
        self._geo_file = excel_path
        self._sheet = sheet
        self._status_col = status_col

        self.ensure_logged_in(username, password, base_url)
        self.goto_geolocation(base_url)

        data = self.load_geo_data(excel_path, sheet, status_col)
        total = len(data)
        if total == 0:
            self.log("Nothing to do - every row is already marked Checked.")
            return 0
        self.log(f"{total} row(s) to process.")
        self._open_wb(excel_path, sheet)

        done = 0
        try:
            for i, row in data.iterrows():
                self._check_stop()
                state = row["State"]
                self.log(f"#{i} [{state}] B:{row['Block']}  "
                         f"GP:{row['GP']}  V:{row['Village']}")

                self.select_state(state)
                self.select_district(row["District"])
                x = self.select_block(row["Block"])
                if x == "New Block":
                    self.select_state(state)
                    self.select_district(row["District"])
                    self.select_block(row["Block"])

                y = self.select_gp(row["GP"])
                if y == "New GP":
                    self.select_state(state)
                    self.select_district(row["District"])
                    self.select_block(row["Block"])
                    self.select_gp(row["GP"])

                self.select_village(row["Village"])
                self.mark_checked(row["src_row"])
                done += 1
        except StopRequested:
            self.log(f"Stopped by user. {done} of {total} row(s) done "
                     f"(progress saved).")
            return done

        self.log(f"Done. Processed {done} row(s).")
        return done


# ---------------------------------------------------------------- runner
def run_geo(driver, excel_path, base_url, username, password,
            sheet="Geographical Details", status_col="PMIS",
            log=print, stop_event=None):
    """Entry point used by the app (config UPDATE_TYPES -> 'pmis_geo:run_geo')."""
    bot = GeoBot(driver, log=log, stop_event=stop_event)
    return bot.run(excel_path, base_url, username, password,
                   sheet=sheet, status_col=status_col)
