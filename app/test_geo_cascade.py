"""Self-check for the geolocation cascade (no browser, real workbook).

Only the browser is faked. The Excel side - reading the sheet, skipping rows
already Checked, and writing 'Checked' back through the batched flush - runs
for real against a temporary workbook.

Run:  python test_geo_cascade.py
"""
import os
import sys
import tempfile

sys.path.insert(0, "libs")

import openpyxl

import config
from pmis_geo import GeoBot, LEVELS

HEADERS = [col for _, col, _, _ in LEVELS] + ["PMIS"]
NAMES = [sel for _, _, sel, _ in LEVELS]
ROW = ["ODISHA", "KEONJHAR", "GHATGAON", "BADAPICHULLA", "JHARBEDA"]


def workbook(rows, sheet="Geographical Details"):
    """A real .xlsx: rows are (hamlet, pmis) pairs on the same cascade."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(HEADERS)
    for hamlet, pmis in rows:
        ws.append(ROW + [hamlet, pmis])
    path = os.path.join(tempfile.mkdtemp(), "geo.xlsx")
    wb.save(path)
    return path


def marks(path, sheet="Geographical Details"):
    ws = openpyxl.load_workbook(path)[sheet]
    return [r[-1].value for r in ws.iter_rows(min_row=2)]


class FakeBot(GeoBot):
    """GeoBot with only the browser replaced."""

    def __init__(self, missing=()):
        GeoBot.__init__(self, driver=None, log=lambda m: None)
        self._missing = set(missing)   # levels PMIS does not have yet
        self.picked = []               # (select name, value)

    def ensure_logged_in(self, *a):
        pass

    def goto_geolocation(self, *a):
        pass

    def select_level(self, select_name, value, add_id=None):
        self.picked.append((select_name, str(value).strip()))
        if select_name in self._missing:
            if add_id is None:
                raise RuntimeError(f"{value} missing and not creatable")
            self._missing.discard(select_name)   # created; now it exists
            return True
        return False


def demo():
    # 1. hamlet present -> all six levels selected, row marked in the file
    path = workbook([("TALA SAHI", "")])
    b = FakeBot()
    assert b.run(path, "u", "n", "p") == 1
    assert [s for s, _ in b.picked] == NAMES, b.picked
    assert b.picked[-1] == ("HamletId", "TALA SAHI")
    assert marks(path) == ["Checked"], marks(path)

    # 2. hamlet blank -> stops after Village, row still completed
    path = workbook([("", ""), ("   ", "")])
    b = FakeBot()
    assert b.run(path, "u", "n", "p") == 2
    assert [s for s, _ in b.picked] == NAMES[:-1] * 2, b.picked
    assert marks(path) == ["Checked", "Checked"]

    # 3. a missing hamlet is created, then the cascade replays from the top
    path = workbook([("NEW SAHI", "")])
    b = FakeBot(missing={"HamletId"})
    assert b.run(path, "u", "n", "p") == 1
    assert [s for s, _ in b.picked] == NAMES + NAMES, b.picked
    assert marks(path) == ["Checked"]

    # 4. a level that can never be selected -> row skipped and left unmarked,
    #    and the add is attempted only once (no infinite replay)
    path = workbook([("X", "")])
    b = FakeBot(missing={"ProvinceId"})
    assert b.run(path, "u", "n", "p") == 0
    assert marks(path) == [None], marks(path)

    # 5. rows already Checked are not processed again
    path = workbook([("A", "Checked"), ("B", "")])
    b = FakeBot()
    assert b.run(path, "u", "n", "p") == 1
    assert b.picked[-1] == ("HamletId", "B"), b.picked
    assert marks(path) == ["Checked", "Checked"]

    # 6. the batched flush really reaches disk, even mid-batch
    many = [("H%d" % i, "") for i in range(config.GEO_FLUSH_EVERY + 3)]
    path = workbook(many)
    b = FakeBot()
    assert b.run(path, "u", "n", "p") == len(many)
    assert marks(path) == ["Checked"] * len(many), marks(path)

    print("geo cascade: all checks passed")


if __name__ == "__main__":
    demo()
