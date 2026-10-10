"""Inspect the widgets of a running Unitrans Portal window.

Verifies the *packaged* build's UI without relying on a screenshot: a Tk
button is a real Win32 child window, so enumerating children and reading their
text confirms which controls actually exist in the compiled executable.

Usage (with the app running):

    python tests/inspect_running_ui.py
    python tests/inspect_running_ui.py --click Settings
"""

from __future__ import annotations

import argparse
import ctypes
import ctypes.wintypes as wintypes
import sys
import time

user32 = ctypes.windll.user32

WANT_TITLE = "Unitrans Portal"
WINDOW_TEXT_LIMIT = 512


def find_window(title_prefix: str = WANT_TITLE, timeout: float = 60.0) -> int:
    """First visible top-level window whose title starts with ``title_prefix``."""
    buffer = ctypes.create_unicode_buffer(WINDOW_TEXT_LIMIT)
    deadline = time.time() + timeout

    while time.time() < deadline:
        found: list[int] = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        def callback(hwnd, _lparam):
            if not user32.IsWindowVisible(hwnd):
                return True
            user32.GetWindowTextW(hwnd, buffer, WINDOW_TEXT_LIMIT)
            if buffer.value.startswith(title_prefix):
                found.append(hwnd)
                return False
            return True

        user32.EnumWindows(callback, 0)
        if found:
            return found[0]
        time.sleep(0.5)

    return 0


def child_windows(parent: int) -> list[tuple[int, str, str, tuple[int, int, int, int]]]:
    """Every descendant window with its class, text and screen rectangle."""
    results: list[tuple[int, str, str, tuple[int, int, int, int]]] = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def callback(hwnd, _lparam):
        text = ctypes.create_unicode_buffer(WINDOW_TEXT_LIMIT)
        klass = ctypes.create_unicode_buffer(WINDOW_TEXT_LIMIT)
        user32.GetWindowTextW(hwnd, text, WINDOW_TEXT_LIMIT)
        user32.GetClassNameW(hwnd, klass, WINDOW_TEXT_LIMIT)

        rect = wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        # Translate to client coordinates relative to the parent, which is how
        # Tk places widgets - absolute screen coords are harder to reason about.
        point = wintypes.POINT(rect.left, rect.top)
        user32.ScreenToClient(parent, ctypes.byref(point))
        size = (point.x, point.y, rect.right - rect.left, rect.bottom - rect.top)

        results.append((hwnd, klass.value, text.value, size))
        return True

    user32.EnumChildWindows(parent, callback, 0)
    return results


def click_at(parent: int, x: int, y: int) -> None:
    """Send a click at client coordinates ``(x, y)`` of ``parent``."""
    lparam = (y << 16) | (x & 0xFFFF)
    user32.SendMessageW(parent, 0x0201, 1, lparam)  # WM_LBUTTONDOWN
    user32.SendMessageW(parent, 0x0202, 0, lparam)  # WM_LBUTTONUP


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Inspect the running app's widgets.")
    parser.add_argument("--click", help="Click the button whose text matches this exactly.")
    parser.add_argument("--timeout", type=float, default=60.0)
    args = parser.parse_args(argv)

    hwnd = find_window(timeout=args.timeout)
    if not hwnd:
        print(f"FAIL: no window titled {WANT_TITLE!r} found within {args.timeout:.0f}s")
        return 1

    title = ctypes.create_unicode_buffer(WINDOW_TEXT_LIMIT)
    user32.GetWindowTextW(hwnd, title, WINDOW_TEXT_LIMIT)
    rect = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    print(f"window : {title.value!r}  hwnd=0x{hwnd:x}")
    print(f"size   : {rect.right - rect.left} x {rect.bottom - rect.top}")

    children = child_windows(hwnd)
    labels = [text for _h, _c, text, _r in children if text.strip()]
    print(f"child windows: {len(children)}  (with text: {len(labels)})")
    print()
    print("Button-like controls found:")
    button_classes = {"Button", "TButton", "TkChild"}
    for _h, klass, text, _size in children:
        if text.strip() and klass in button_classes:
            print(f"  [{klass}] {text!r}")

    if args.click:
        matches = [c for c in children if c[2].strip() == args.click]
        if not matches:
            print()
            print(f"FAIL: no control with text {args.click!r}")
            return 1
        target_hwnd, _klass, text, (x, y, w, h) = matches[0]
        print()
        print(f"clicking {text!r} at client ({x + w // 2}, {y + h // 2}) in {w}x{h}")
        click_at(hwnd, x + w // 2, y + h // 2)
        return 0

    return 0


if __name__ == "__main__":
    sys.exit(main())
