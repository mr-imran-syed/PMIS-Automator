"""Optional credential storage, encrypted with Windows DPAPI.

Used only when the user ticks "Keep credentials" in Settings.

DPAPI (CryptProtectData) ties the ciphertext to the current Windows user
account on this machine: copying settings.json to another PC or another user
profile yields something that cannot be decrypted. That is meaningfully better
than writing the password in plain text, though it is NOT protection against
someone already logged in as that user.

If DPAPI is unavailable (non-Windows), the password is simply not stored -
we never silently fall back to plaintext.
"""

import base64
import ctypes
from ctypes import wintypes

DESCRIPTION = "PMIS Automation Tool credentials"


class _Blob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_char))]


def _available():
    return hasattr(ctypes, "windll")


def _to_blob(data):
    buf = ctypes.create_string_buffer(data, len(data))
    return _Blob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char))), buf


def _take(blob_out):
    """Copy the result out of the DPAPI buffer, then free it."""
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)


def protect(text):
    """Encrypt a string -> base64 token, or None if it cannot be done."""
    if not _available() or text is None:
        return None
    try:
        blob_in, _keep = _to_blob(str(text).encode("utf-8"))
        blob_out = _Blob()
        ok = ctypes.windll.crypt32.CryptProtectData(
            ctypes.byref(blob_in), DESCRIPTION,
            None, None, None, 0, ctypes.byref(blob_out))
        if not ok:
            return None
        return base64.b64encode(_take(blob_out)).decode("ascii")
    except Exception:
        return None


def unprotect(token):
    """Decrypt a base64 token -> string, or None if it cannot be read."""
    if not _available() or not token:
        return None
    try:
        blob_in, _keep = _to_blob(base64.b64decode(token))
        blob_out = _Blob()
        ok = ctypes.windll.crypt32.CryptUnprotectData(
            ctypes.byref(blob_in), None,
            None, None, None, 0, ctypes.byref(blob_out))
        if not ok:
            return None
        return _take(blob_out).decode("utf-8")
    except Exception:
        return None


if __name__ == "__main__":
    token = protect("hunter2")
    print("available :", _available())
    print("token     :", (token or "")[:40], "...")
    print("round-trip:", unprotect(token) == "hunter2")
