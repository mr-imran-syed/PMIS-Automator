"""Browser detection + Selenium driver factory.

Detects an installed browser in this order of preference -- Firefox, Chrome,
then Edge (Edge ships with Windows 11, so it is the guaranteed fallback) -- and
builds the matching Selenium WebDriver. Selenium Manager (built into Selenium
4.6+) auto-downloads the correct driver version for whichever browser is found,
so there is no chromedriver/geckodriver version juggling.
"""

import os

try:
    import winreg  # Windows only
except ImportError:  # pragma: no cover - non-Windows dev machines
    winreg = None


class BrowserNotFound(Exception):
    """Raised when no supported browser is installed."""


# Detection order is set by _PREFERENCE below (edit that to re-rank browsers).
# Each entry: (friendly name, executable, list of common paths).
_ENV = os.environ
_BROWSERS = [
    (
        "chrome",
        "chrome.exe",
        [
            os.path.join(_ENV.get("PROGRAMFILES", r"C:\Program Files"),
                         r"Google\Chrome\Application\chrome.exe"),
            os.path.join(_ENV.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
                         r"Google\Chrome\Application\chrome.exe"),
            os.path.join(_ENV.get("LOCALAPPDATA", ""),
                         r"Google\Chrome\Application\chrome.exe"),
        ],
    ),
    (
        "firefox",
        "firefox.exe",
        [
            os.path.join(_ENV.get("PROGRAMFILES", r"C:\Program Files"),
                         r"Mozilla Firefox\firefox.exe"),
            os.path.join(_ENV.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
                         r"Mozilla Firefox\firefox.exe"),
        ],
    ),
    (
        "edge",
        "msedge.exe",
        [
            os.path.join(_ENV.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
                         r"Microsoft\Edge\Application\msedge.exe"),
            os.path.join(_ENV.get("PROGRAMFILES", r"C:\Program Files"),
                         r"Microsoft\Edge\Application\msedge.exe"),
        ],
    ),
]

# Firefox first, then Chrome, then Edge as the always-present fallback.
_PREFERENCE = ("firefox", "chrome", "edge")
_BROWSERS.sort(key=lambda b: _PREFERENCE.index(b[0]))


def _from_registry(exe_name):
    """Look up a browser path from the Windows 'App Paths' registry keys."""
    if winreg is None:
        return None
    sub = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\%s" % exe_name
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, sub) as key:
                path, _ = winreg.QueryValueEx(key, None)
                if path and os.path.exists(path):
                    return path
        except OSError:
            continue
    return None


def detect_browser():
    """Return (name, binary_path) for the most preferred installed browser,
    or None if none of Chrome/Firefox/Edge is found."""
    for name, exe, paths in _BROWSERS:
        for p in paths:
            if p and os.path.exists(p):
                return name, p
        reg = _from_registry(exe)
        if reg:
            return name, reg
    return None


def available_browsers():
    """Return a list of all detected browser names (for diagnostics)."""
    found = []
    for name, exe, paths in _BROWSERS:
        if any(p and os.path.exists(p) for p in paths) or _from_registry(exe):
            found.append(name)
    return found


def make_driver(preferred=None):
    """Detect a browser and return (driver, browser_name).

    `preferred` optionally forces one of 'chrome'/'firefox'/'edge' if installed;
    otherwise the first browser found in preference order is used.
    Raises BrowserNotFound if nothing suitable is installed.
    """
    detected = detect_browser()
    if detected is None:
        raise BrowserNotFound(
            "No supported browser found. Please install Google Chrome or "
            "Mozilla Firefox, then run this app again."
        )

    name, binary = detected
    if preferred and preferred in available_browsers():
        name = preferred

    if name == "chrome":
        from selenium.webdriver import Chrome, ChromeOptions
        opts = ChromeOptions()
        opts.add_argument("--start-maximized")
        opts.add_experimental_option("excludeSwitches", ["enable-logging"])
        driver = Chrome(options=opts)
    elif name == "edge":
        from selenium.webdriver import Edge, EdgeOptions
        opts = EdgeOptions()
        opts.add_argument("--start-maximized")
        opts.add_experimental_option("excludeSwitches", ["enable-logging"])
        driver = Edge(options=opts)
    else:  # firefox
        from selenium.webdriver import Firefox, FirefoxOptions
        opts = FirefoxOptions()
        driver = Firefox(options=opts)
        driver.maximize_window()

    return driver, name


if __name__ == "__main__":
    print("Detected (preferred):", detect_browser())
    print("All available:", available_browsers())
