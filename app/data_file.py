"""Bundled Excel data file helpers: locate it, open it in Excel for editing,
and detect whether it is currently open (locked) so we can ask the user to
close it before a run writes progress into it.
"""

import os
import shutil

from openpyxl import load_workbook

from config import data_file_path, data_template_path


def read_rows(path, sheet):
    """A worksheet as a list of dicts, keyed by header, plus '_row'.

    Values arrive exactly as the cell holds them - a literal "NA" stays the
    text "NA", and blanks are ''. Dates come back as datetime objects.
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet]
        rows = ws.iter_rows(values_only=True)
        headers = [str(h).strip() if h is not None else ""
                   for h in next(rows, ())]
        out = []
        for n, values in enumerate(rows, start=2):
            row = {h: ("" if v is None else v)
                   for h, v in zip(headers, values) if h}
            if not any(str(v).strip() for v in row.values()):
                continue                      # trailing blank row
            row["_row"] = n
            out.append(row)
        return headers, out
    finally:
        wb.close()


def write_column(path, sheet, column, mapping, log=print, protect=()):
    """Write {excel_row: text} into one column, creating it if absent.

    `protect` is a tuple of lowercase prefixes whose rows are left untouched,
    so a geo check can never wipe the record of something already created.
    Returns how many cells were written.
    """
    if not mapping:
        return 0
    try:
        wb = load_workbook(path)
        ws = wb[sheet]
        headers = {c.value: c.column for c in ws[1]}
        col = headers.get(column)
        if col is None:
            col = ws.max_column + 1
            ws.cell(row=1, column=col, value=column)
        written = 0
        for excel_row, value in mapping.items():
            cell = ws.cell(row=int(excel_row), column=col)
            if protect and str(cell.value or "").strip().lower() \
                    .startswith(protect):
                continue
            cell.value = str(value)[:400]
            written += 1
        wb.save(path)
        return written
    except PermissionError:
        log("  ! Could not write results - the workbook is open in Excel.")
    except Exception as e:
        log(f"  ! Could not write results: {e}")
    return 0


def text(row, column):
    """One cell as trimmed text ('' when absent)."""
    return str(row.get(column, "") or "").strip()


def exists():
    return os.path.exists(data_file_path())


def ensure_data_file():
    """A fresh clone ships only the template (the real workbook is gitignored
    because it holds participant PII). Create the working copy from it."""
    path = data_file_path()
    if os.path.exists(path):
        return path, False
    tpl = data_template_path()
    if os.path.exists(tpl):
        shutil.copyfile(tpl, path)
        return path, True
    return path, False


def open_in_excel():
    """Open the data file in the user's default spreadsheet app (Excel)."""
    path = data_file_path()
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    os.startfile(path)   # Windows: launches Excel (or default handler)


def _lock_file_path():
    """Excel creates a hidden owner file '~$<name>' next to an open workbook."""
    path = data_file_path()
    d, name = os.path.split(path)
    return os.path.join(d, "~$" + name)


def is_open(path=None):
    """Best-effort check for whether the workbook is currently open in Excel.

    Uses two signals:
      1. Excel's '~$' owner/lock file exists beside it.
      2. The file cannot be opened for writing (Excel holds a write lock).
    Either one being true means 'treat as open'.
    """
    path = path or data_file_path()
    if not os.path.exists(path):
        return False

    if os.path.exists(_lock_file_path()):
        return True

    # Try to open for read+write / append; Excel's lock makes this raise.
    try:
        with open(path, "a+b"):
            pass
        return False
    except (PermissionError, OSError):
        return True


if __name__ == "__main__":
    print("data file:", data_file_path())
    print("exists   :", exists())
    print("is_open  :", is_open())
