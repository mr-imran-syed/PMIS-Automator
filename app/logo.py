"""The app's logo: a steaming idli, in pixels.

Drawn straight onto a tk.PhotoImage rather than shipped as a .png or .ico -
no binary asset to package, and it scales to whatever size is asked for. The
Windows taskbar insists on a real .ico file, so one is generated into the
temp folder on the fly.
"""

import os
import struct
import sys
import tempfile
import tkinter as tk

# Windows groups taskbar buttons - and picks their icon - by this id. Without
# it the app is grouped under the Python interpreter and wears its icon.
APP_ID = "HeiferIndia.PMISAutomationTool"

# ' ' is transparent. Each letter is a colour in PALETTE.
#   s steam   h highlight   i idli   d idli shade   p plate   q plate shade
IDLI = [
    "      s    s    ",
    "     s      s   ",
    "      s    s    ",
    "     s      s   ",
    "                ",
    "      hhhh      ",
    "    hhhhhhhh    ",
    "   hhiiiiiihh   ",
    "  hiiiiiiiiiih  ",
    "  iiiiiiiiiiii  ",
    "  iiiiiiiiiiii  ",
    "  diiiiiiiiiid  ",
    "   dddddddddd   ",
    "                ",
    "  pppppppppppp  ",
    " qqqqqqqqqqqqqq ",
]

PALETTE = {
    "s": "#BFD3DE",      # steam
    "h": "#FFFFFF",      # lit top of the idli
    "i": "#EDE7D5",      # idli body
    "d": "#CCC2A6",      # where it meets the plate
    "p": "#90A4B4",      # plate
    "q": "#6B8294",      # plate, underside
}


def image(scale=2):
    """A tk.PhotoImage of the logo, `scale` times 16x16.

    Needs a Tk root to exist already, so call it after the window is made.
    """
    img = tk.PhotoImage(width=len(IDLI[0]), height=len(IDLI))
    for y, line in enumerate(IDLI):
        for x, ch in enumerate(line):
            colour = PALETTE.get(ch)
            if colour:
                img.put(colour, (x, y))
            else:
                img.transparency_set(x, y, True)
    return img.zoom(scale) if scale > 1 else img


def set_app_id(app_id=APP_ID):
    """Tell Windows this is its own application, not 'Python'.

    Must run BEFORE the window is created, or the taskbar has already decided
    what this process is. A no-op anywhere but Windows.
    """
    if not sys.platform.startswith("win"):
        return False
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
        return True
    except Exception:
        return False


def ico_path(sizes=(16, 32, 48, 64)):      # exact multiples of 16
    """Write a multi-size .ico of the logo to the temp folder; return its path.

    Each entry is a PNG, which Windows has accepted inside .ico since Vista -
    so Tk writes the payloads and this only has to add the directory in front.
    Needs a Tk root to exist.
    """
    path = os.path.join(tempfile.gettempdir(), "pmis-idli.ico")
    pngs = []
    for size in sizes:
        scale = size // len(IDLI)          # 1x, 2x, 3x, 4x
        tmp = os.path.join(tempfile.gettempdir(), f"pmis-idli-{size}.png")
        image(scale).write(tmp, format="png")
        with open(tmp, "rb") as f:
            pngs.append((size, f.read()))
        os.remove(tmp)

    header = struct.pack("<HHH", 0, 1, len(pngs))     # reserved, type=icon, n
    offset = len(header) + 16 * len(pngs)
    entries, payload = b"", b""
    for size, data in pngs:
        entries += struct.pack(
            "<BBBBHHII",
            size if size < 256 else 0,   # width  (0 means 256)
            size if size < 256 else 0,   # height
            0, 0,                        # palette size, reserved
            1, 32,                       # colour planes, bits per pixel
            len(data), offset)
        payload += data
        offset += len(data)

    with open(path, "wb") as f:
        f.write(header + entries + payload)
    return path


def apply_to(window, scale=2):
    """Dress `window` in the logo: title bar, Alt-Tab and taskbar.

    iconphoto covers the title bar; Windows reads the taskbar button from the
    .ico handed to iconbitmap. Returns the PhotoImage, which must stay
    referenced or Tk garbage-collects it, or None if Tk would not take it.
    """
    img = None
    try:
        img = image(scale)
        window.iconphoto(True, img)
    except Exception:
        pass
    try:
        window.iconbitmap(default=ico_path())
    except Exception:
        pass
    return img


if __name__ == "__main__":          # eyeball it: python logo.py
    set_app_id()
    root = tk.Tk()
    root.title("steaming idli")
    root.configure(bg="#1E1E1E")
    apply_to(root)
    kept = [image(s) for s in (1, 2, 4, 8)]
    for img in kept:
        tk.Label(root, image=img, bg="#1E1E1E").pack(side="left", padx=8,
                                                     pady=8)
    root.mainloop()
