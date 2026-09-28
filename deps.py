"""Dependency checking + in-app installation/update.

- Packages are installed into an app-managed 'libs/' folder via pip
  (`pip install --target libs`), so the global Python stays untouched and the
  portable folder carries its own dependencies.
- pip runs as a hidden subprocess; its output is streamed to a callback so the
  UI can show progress. No visible command window (spec 3).
"""

import os
import sys
import json
import subprocess
import importlib
import urllib.request

import config

# Hide the console window pip would otherwise flash on Windows.
_NO_WINDOW = 0x08000000 if os.name == "nt" else 0

STATE_READY = "ready"
STATE_UPDATE = "update"
STATE_MISSING = "missing"
STATE_ERROR = "error"


# --------------------------------------------------------------- versions
def _installed_version(import_name):
    try:
        from importlib.metadata import version, PackageNotFoundError
    except ImportError:  # pragma: no cover
        return None
    try:
        return version(import_name)
    except Exception:
        return None


def _latest_version(pip_name, timeout=3):
    """Best-effort latest version from PyPI. Returns None if offline/unknown."""
    url = f"https://pypi.org/pypi/{pip_name}/json"
    try:
        with urllib.request.urlopen(url, timeout=timeout) as r:
            data = json.load(r)
        return data["info"]["version"]
    except Exception:
        return None


def _parse_version(text):
    """'v0.2.10' -> (0, 2, 10). Unparseable parts become 0."""
    if not text:
        return ()
    cleaned = str(text).strip().lstrip("vV").split("-")[0].split("+")[0]
    parts = []
    for chunk in cleaned.split("."):
        digits = "".join(ch for ch in chunk if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts)


def _newer(remote, local):
    """True when `remote` is a strictly higher version than `local`."""
    r, l = _parse_version(remote), _parse_version(local)
    if not r or not l:
        return False
    size = max(len(r), len(l))
    r += (0,) * (size - len(r))
    l += (0,) * (size - len(l))
    return r > l


def latest_app_version(repo=None, timeout=4):
    """Newest published version of this app on GitHub, or None if unknown.

    Tries the latest release tag first, then falls back to the VERSION file on
    the default branch (so it works whether or not releases are used).
    """
    repo = repo or getattr(config, "GITHUB_REPO", None)
    if not repo or "REPLACE_ME" in repo:
        return None

    try:
        url = f"https://api.github.com/repos/{repo}/releases/latest"
        req = urllib.request.Request(
            url, headers={"Accept": "application/vnd.github+json",
                          "User-Agent": "pmis-automation-tool"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            tag = json.load(r).get("tag_name")
        if tag:
            return str(tag).strip()
    except Exception:
        pass

    branch = getattr(config, "GITHUB_BRANCH", "main")
    for br in (branch, "master"):
        try:
            raw = f"https://raw.githubusercontent.com/{repo}/{br}/VERSION"
            req = urllib.request.Request(
                raw, headers={"User-Agent": "pmis-automation-tool"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                v = r.read().decode("utf-8").strip()
            if v:
                return v
        except Exception:
            continue
    return None


def _importable(import_name):
    config.ensure_libs_on_path()
    try:
        importlib.import_module(import_name)
        return True
    except Exception:
        return False


# --------------------------------------------------------------- checks
def check_dependency(dep, check_updates=True):
    """Return a status dict for one dependency entry from config.DEPENDENCIES."""
    kind = dep["kind"]
    out = {"key": dep["key"], "name": dep["name"], "kind": kind,
           "mandatory": dep.get("mandatory", True),
           "state": STATE_MISSING, "version": None, "latest": None, "detail": ""}

    if kind == "python":
        out["state"] = STATE_READY
        out["version"] = "%d.%d.%d" % sys.version_info[:3]
        return out

    if kind == "package":
        imp = dep["import_name"]
        if not _importable(imp):
            out["state"] = STATE_MISSING
            return out
        # metadata lookups use the distribution (pip) name, not the import name
        out["version"] = _installed_version(dep["pip_name"])
        out["state"] = STATE_READY
        if check_updates:
            latest = _latest_version(dep["pip_name"])
            out["latest"] = latest
            if latest and out["version"] and latest != out["version"]:
                out["state"] = STATE_UPDATE
        return out

    if kind == "app":
        out["version"] = config.APP_VERSION
        if not check_updates:
            out["state"] = STATE_READY
            out["detail"] = f"v{config.APP_VERSION}"
            return out
        latest = latest_app_version()
        out["latest"] = latest
        if latest is None:
            out["state"] = STATE_READY
            out["detail"] = f"v{config.APP_VERSION} (update check unavailable)"
        elif _newer(latest, config.APP_VERSION):
            out["state"] = STATE_UPDATE
            out["detail"] = f"v{config.APP_VERSION} -> {latest}"
        else:
            out["state"] = STATE_READY
            out["detail"] = f"v{config.APP_VERSION} (up to date)"
        return out

    if kind == "browser":
        try:
            config.ensure_libs_on_path()
            from browsers import detect_browser
            found = detect_browser()
        except Exception:
            found = None
        if found:
            out["state"] = STATE_READY
            out["detail"] = found[0].title()
        else:
            out["state"] = STATE_MISSING
            out["detail"] = "Install Chrome or Firefox"
        return out

    if kind == "driver":
        # Selenium Manager auto-fetches the matching driver on first run,
        # so 'ready' once Selenium and a browser are both present.
        sel_ok = _importable("selenium")
        try:
            from browsers import detect_browser
            has_browser = detect_browser() is not None
        except Exception:
            has_browser = False
        if sel_ok and has_browser:
            out["state"] = STATE_READY
            out["detail"] = "Auto-managed"
        else:
            out["state"] = STATE_MISSING
            out["detail"] = "Needs Selenium + a browser"
        return out

    return out


def check_all(check_updates=True):
    return [check_dependency(d, check_updates) for d in config.DEPENDENCIES]


# --------------------------------------------------------------- install
def install_package(pip_name, log_cb=print, upgrade=False):
    """pip-install (or upgrade) a package into the app-managed libs folder.
    Streams pip output to log_cb. Returns True on success."""
    target = config.libs_dir()
    os.makedirs(target, exist_ok=True)
    cmd = [sys.executable, "-m", "pip", "install", "--target", target,
           "--no-warn-script-location"]
    if upgrade:
        cmd.append("--upgrade")
    cmd.append(pip_name)   # package always last
    log_cb(f"$ pip install {'--upgrade ' if upgrade else ''}{pip_name}")
    try:
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, bufsize=1, creationflags=_NO_WINDOW)
    except Exception as e:
        log_cb(f"ERROR launching pip: {e}")
        return False

    for line in proc.stdout:
        log_cb(line.rstrip())
    proc.wait()
    ok = proc.returncode == 0
    log_cb("Done." if ok else f"pip exited with code {proc.returncode}")
    if ok:
        importlib.invalidate_caches()
    return ok


def dependency_by_key(key):
    for d in config.DEPENDENCIES:
        if d["key"] == key:
            return d
    return None


if __name__ == "__main__":
    for s in check_all(check_updates=False):
        print(f"{s['name']:16} {s['state']:8} {s.get('version') or s.get('detail') or ''}")
