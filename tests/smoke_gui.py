"""Launch the real GUI, confirm it rendered, and optionally save a screenshot.

    python tests/smoke_gui.py --screenshot out.png

Run from the project root with the build venv. This drives the actual
``WaybillSorterApp`` rather than inspecting the source, so it catches failures
that only appear once the widget tree is built - a missing i18n key, a bad
geometry string, an exception during layout.

If the window is not found within the timeout the script exits non-zero and
prints the Tk window that *was* created, so a title mismatch is obvious.
"""

from __future__ import annotations

import argparse
import ctypes
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import tkinter as tk  # noqa: E402

from version import version_label  # noqa: E402


def find_window_by_title(title: str, timeout: float = 20.0) -> int:
    """Return the first top-level window whose text starts with ``title``."""
    user32 = ctypes.windll.user32
    buffer = ctypes.create_unicode_buffer(512)
    deadline = time.time() + timeout

    while time.time() < deadline:
        found: list[int] = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        def callback(hwnd, _lparam):
            if not user32.IsWindowVisible(hwnd):
                return True
            user32.GetWindowTextW(hwnd, buffer, 512)
            text = buffer.value
            if text and (text == title or text.startswith(title)):
                # Skip tool windows and hidden helpers.
                if user32.GetWindow(hwnd, 4) == 0:  # GW_OWNER
                    found.append(hwnd)
                    return False
            return True

        user32.EnumWindows(callback, 0)
        if found:
            return found[0]
        time.sleep(0.25)

    return 0


def hwnd_for_widget(widget: tk.Misc) -> int:
    """Return the Win32 handle backing a Tk widget.

    Much more reliable than searching windows by title: a Toplevel called
    "Settings" is ambiguous on a machine that also runs the Windows Settings
    app, and searching would happily capture the wrong window. Tk already knows
    its own handle, so ask it directly.
    """
    widget.update_idletasks()
    try:
        return int(widget.winfo_id())
    except (tk.TclError, ValueError):
        return 0


def capture_window(hwnd: int, destination: Path) -> bool:
    """Screenshot a single window by handle."""
    try:
        from PIL import Image
    except ImportError:
        print("Pillow not available; skipping screenshot.")
        return False

    import ctypes.wintypes as wintypes

    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32

    # Bring the window forward so the capture is not occlusion.
    user32.SetForegroundWindow(hwnd)
    time.sleep(1.0)

    rect = wintypes.RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        print("GetWindowRect failed.")
        return False
    width = rect.right - rect.left
    height = rect.bottom - rect.top

    hdc_window = user32.GetWindowDC(hwnd)
    hdc_mem = gdi32.CreateCompatibleDC(hdc_window)
    bitmap = gdi32.CreateCompatibleBitmap(hdc_window, width, height)
    gdi32.SelectObject(hdc_mem, bitmap)

    # PW_RENDERFULLCONTENT (2) captures composited content, so it works even
    # when the window is not the top-most one.
    if not user32.PrintWindow(hwnd, hdc_mem, 2):
        print("PrintWindow failed; falling back to BitBlt.")
        gdi32.BitBlt(hdc_mem, 0, 0, width, height, hdc_window, 0, 0, 0x00CC0020)

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", wintypes.DWORD),
            ("biWidth", wintypes.LONG),
            ("biHeight", wintypes.LONG),
            ("biPlanes", wintypes.WORD),
            ("biBitCount", wintypes.WORD),
            ("biCompression", wintypes.DWORD),
            ("biSizeImage", wintypes.DWORD),
            ("biXPelsPerMeter", wintypes.LONG),
            ("biYPelsPerMeter", wintypes.LONG),
            ("biClrUsed", wintypes.DWORD),
            ("biClrImportant", wintypes.DWORD),
        ]

    header = BITMAPINFOHEADER()
    header.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    header.biWidth = width
    header.biHeight = -height  # negative = top-down rows
    header.biPlanes = 1
    header.biBitCount = 32
    header.biCompression = 0

    buffer_size = width * height * 4
    buffer = ctypes.create_string_buffer(buffer_size)
    gdi32.GetDIBits(hdc_mem, bitmap, 0, height, buffer, ctypes.byref(header), 0)

    image = Image.frombuffer("RGBA", (width, height), buffer, "raw", "BGRA", 0, 1)
    destination.parent.mkdir(parents=True, exist_ok=True)
    image.convert("RGB").save(destination)

    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(hdc_mem)
    user32.ReleaseDC(hwnd, hdc_window)

    print(f"Screenshot saved: {destination}")
    return True


def _make_console_utf8_safe() -> None:
    """Stop the harness crashing when it prints non-ASCII window titles.

    A default Windows console is cp1252, so printing a Chinese window title
    raises UnicodeEncodeError. The application already reconfigures its own
    stdio; the test harness needs the same treatment.
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def main(argv: list[str] | None = None) -> int:
    _make_console_utf8_safe()
    parser = argparse.ArgumentParser(description="Verify the GUI launches and renders.")
    parser.add_argument("--screenshot", help="Save a PNG of the main window here.")
    parser.add_argument("--keep-open", type=float, default=6.0, help="Seconds to leave the window up.")
    parser.add_argument("--settings", action="store_true", help="Open the Settings window and check it renders.")
    parser.add_argument(
        "--update-dialog",
        action="store_true",
        help="Show the update-available dialog using a synthetic newer release.",
    )
    parser.add_argument("--language", choices=["en", "zh"], help="Switch language before capturing.")
    parser.add_argument(
        "--screen",
        choices=["selector", "sort", "driver", "unlock"],
        help="Open a specific screen before capturing.",
    )
    args = parser.parse_args(argv)

    from waybill_sorter import WaybillSorterApp

    app = WaybillSorterApp()
    expected_title = app.title()
    print(f"Tk reports window title: {expected_title!r}")

    if not expected_title or expected_title == "window_title":
        print("FAIL: the window title is empty or an unresolved i18n key.")
        app.destroy()
        return 1

    # Render one frame so layout errors surface here rather than on screen.
    app.update_idletasks()
    app.update()

    footer_expected = version_label()
    print(f"Expected footer text    : {footer_expected!r}")

    # Confirm the footer label widget actually carries the version string.
    footer_found = False
    def walk(widget):
        nonlocal footer_found
        for child in widget.winfo_children():
            if isinstance(child, tk.Label):
                try:
                    if footer_expected in str(child.cget("text")):
                        footer_found = True
                except tk.TclError:
                    pass
            walk(child)
    walk(app)
    if not footer_found:
        print("FAIL: no label in the widget tree shows the version footer.")
        app.destroy()
        return 1
    print("PASS: version footer present in the widget tree.")

    if args.language and app.language.get() != args.language:
        app.toggle_language()
        app.update_idletasks()
        app.update()
        print(f"PASS: switched language to {args.language!r} (title now {app.title()!r}).")

    if args.screen and args.screen != "selector":
        {
            "sort": app.select_sort_workflow,
            "driver": app.select_driver_workflow,
            "unlock": app.select_unlock_workflow,
        }[args.screen]()
        app.update_idletasks()
        app.update()
        print(f"PASS: opened the {args.screen!r} screen.")

        # A back button must be reachable from every section, otherwise there
        # is no way out except closing the app.
        back_text = app.tr("back_to_workflows")
        found_back = _find_button(app, lambda text: back_text in text)
        if not found_back:
            print(f"FAIL: no back button containing {back_text!r} on the {args.screen!r} screen.")
            app.destroy()
            return 1
        print(f"PASS: back button present and visible ({found_back!r}).")

    if args.settings:
        if not hasattr(app, "open_settings"):
            print("FAIL: the app has no Settings entry point.")
            app.destroy()
            return 1
        app.open_settings()
        app.update_idletasks()
        app.update()
        if app.settings_window is None or not app.settings_window.winfo_exists():
            print("FAIL: the Settings window did not open.")
            app.destroy()
            return 1
        print("PASS: Settings window opened.")
        _report_labels(app.settings_window, "Settings", indent="      ")

    if args.update_dialog:
        result = _synthetic_update_result()
        app.show_update_dialog(result)
        app.update_idletasks()
        app.update()
        print("PASS: update-available dialog rendered for a synthetic newer release.")

    _report_focus_rings(app)

    app.deiconify()
    app.lift()
    app.update()

    # Capture the dialog we just built, or the main window. Ask Tk for the
    # handle rather than searching by title, which is ambiguous ("Settings" is
    # also the name of a Windows app).
    target = None
    if args.update_dialog:
        target = _last_toplevel(app)
    elif args.settings and app.settings_window is not None:
        target = app.settings_window

    hwnd = hwnd_for_widget(target) if target is not None else 0
    if not hwnd:
        hwnd = find_window_by_title(expected_title)
        if not hwnd:
            print(f"FAIL: no visible top-level window titled {expected_title!r} appeared.")
            app.destroy()
            return 1
        print("NOTE: falling back to the main window for capture.")
    print(f"PASS: window handle 0x{hwnd:x} found and visible.")

    if args.screenshot:
        capture_window(hwnd, Path(args.screenshot))

    # Let it render, then close through Tk to prove clean shutdown.
    deadline = time.time() + args.keep_open
    while time.time() < deadline:
        app.update()
        time.sleep(0.05)

    app.destroy()
    print("PASS: GUI built, rendered, and shut down cleanly.")
    return 0


def _find_button(app: tk.Misc, predicate) -> str | None:
    """Text of the first visible Button satisfying ``predicate``, else None."""
    match: list[str] = []

    def walk(widget):
        for child in widget.winfo_children():
            if isinstance(child, tk.Button):
                try:
                    text = str(child.cget("text"))
                    # winfo_ismapped is false for widgets in a screen that was
                    # built but never shown.
                    if child.winfo_ismapped() and predicate(text):
                        match.append(text)
                except tk.TclError:
                    pass
            walk(child)

    walk(app)
    return match[0] if match else None


def _report_focus_rings(app: tk.Misc) -> None:
    """Flag buttons whose Tk focus ring would draw as a stray box.

    highlightthickness defaults to a non-zero value in Tk; left unset it shows
    as a blue box around whichever button takes keyboard focus.
    """
    offenders = []

    def walk(widget):
        for child in widget.winfo_children():
            if isinstance(child, tk.Button):
                try:
                    if int(child.cget("highlightthickness")) != 0:
                        offenders.append(str(child.cget("text")))
                except (tk.TclError, ValueError):
                    pass
            walk(child)

    walk(app)
    if offenders:
        print(f"WARNING: {len(offenders)} button(s) have a non-zero focus ring: {offenders}")
    else:
        print("PASS: no button has a focus ring (highlightthickness is 0 everywhere).")


def _last_toplevel(app: tk.Misc) -> tk.Misc | None:
    """The most recently created Toplevel child, i.e. the dialog just opened."""
    toplevels = [child for child in app.winfo_children() if isinstance(child, tk.Toplevel)]
    return toplevels[-1] if toplevels else None


def _report_labels(window: tk.Misc, heading: str, indent: str = "      ") -> None:
    """Print the text of every label in a window, to confirm nothing is a raw key."""
    texts = []

    def walk(widget):
        for child in widget.winfo_children():
            if isinstance(child, tk.Label):
                try:
                    texts.append(str(child.cget("text")))
                except tk.TclError:
                    pass
            walk(child)

    walk(window)
    print(f"{indent}{heading} labels:")
    for text in texts:
        flag = "  <-- UNRESOLVED KEY?" if text.islower() and "_" in text else ""
        print(f"{indent}  - {text!r}{flag}")


def _synthetic_update_result():
    """A fabricated 'newer release' so the update dialog can be seen today.

    0.2.0 does not exist yet, so the only way to verify this dialog renders
    before that release is to hand the UI a made-up result.
    """
    from updater import UpdateCheckResult

    result = UpdateCheckResult(status="update-available", current_version="0.1.0")
    result.remote_version = "0.2.0"
    result.download_url = "https://github.com/ElvaraStudio/unitrans-portal/releases/download/v0.2.0/Unitrans-Portal-0.2.0.exe"
    result.release_page = "https://github.com/ElvaraStudio/unitrans-portal/releases/tag/v0.2.0"
    result.size_bytes = 24_900_000
    result.sha256 = "b" * 64
    result.notes = {
        "en": {
            "headline": "Sample update",
            "notes": [
                "Synthetic note for layout verification - not a real release.",
                "Second line to confirm the notes pane scrolls and wraps properly.",
            ],
        },
        "zh": {
            "headline": "\u793a\u4f8b\u66f4\u65b0",
            "notes": [
                "\u7528\u4e8e\u68c0\u67e5\u5e03\u5c40\u7684\u793a\u4f8b\u6587\u5b57\u3002",
                "\u7b2c\u4e8c\u884c\u7528\u4e8e\u786e\u8ba4\u5907\u6ce8\u533a\u57df\u53ef\u4ee5\u6b63\u5e38\u6362\u884c\u3002",
            ],
        },
    }
    return result


if __name__ == "__main__":
    raise SystemExit(main())


if __name__ == "__main__":
    raise SystemExit(main())
