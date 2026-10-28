"""User settings, persisted next to the app in settings.json.

Deliberately NOT for credentials - those are typed every launch and never
stored. This is only for UI preferences such as the Auto Save toggle.
"""

import os
import json

from config import app_dir

SETTINGS_FILE = "settings.json"

# Auto Save defaults to OFF: it submits records to PMIS, so it must be an
# explicit, deliberate choice rather than something inherited silently.
DEFAULTS = {
    "participant_auto_save": False,
    "shg_auto_save": False,
    # When on, the username and a DPAPI-encrypted password are kept so the
    # fields are pre-filled next launch. Off by default.
    "keep_credentials": False,
    "saved_username": "",
    "saved_password": "",      # DPAPI ciphertext, never plaintext
}


def path():
    return os.path.join(app_dir(), SETTINGS_FILE)


def load():
    data = dict(DEFAULTS)
    try:
        with open(path(), encoding="utf-8") as f:
            stored = json.load(f)
        if isinstance(stored, dict):
            data.update({k: v for k, v in stored.items() if k in DEFAULTS})
    except Exception:
        pass
    return data


def save(data):
    try:
        with open(path(), "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        return True
    except Exception:
        return False


def get(key):
    return load().get(key, DEFAULTS.get(key))


def set_value(key, value):
    data = load()
    data[key] = value
    return save(data)
