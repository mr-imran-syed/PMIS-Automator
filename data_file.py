"""Bundled Excel data file helpers: locate it, open it in Excel for editing,
and detect whether it is currently open (locked) so we can ask the user to
close it before a run writes progress into it.
"""

import os
import shutil

from config import data_file_path, data_template_path


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
