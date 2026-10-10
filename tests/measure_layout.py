"""Measure how the interface changes size between languages.

Widget sizes are compared in English and Chinese so the source of any layout
shift is identified rather than guessed at. Run from the project root:

    python tests/measure_layout.py
"""

from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from waybill_sorter import WaybillSorterApp  # noqa: E402

#: Widgets whose size the user would notice changing when the language toggles.
WATCHED = (
    "sort_button",
    "make_pdf_button",
    "open_output_button",
    "driver_fill_button",
    "unlock_button",
)


def geometry_of(widget: tk.Misc) -> tuple[int, int, int, int]:
    widget.update_idletasks()
    return (widget.winfo_width(), widget.winfo_height(), widget.winfo_x(), widget.winfo_y())


def collect(app: WaybillSorterApp) -> dict[str, tuple[int, int, int, int]]:
    app.update_idletasks()
    snapshot: dict[str, tuple[int, int, int, int]] = {
        "<window>": (app.winfo_width(), app.winfo_height(), app.winfo_x(), app.winfo_y()),
        "<reqsize>": (app.winfo_reqwidth(), app.winfo_reqheight(), 0, 0),
    }
    for name in WATCHED:
        widget = getattr(app, name, None)
        if widget is not None:
            try:
                if widget.winfo_exists():
                    snapshot[name] = geometry_of(widget)
            except tk.TclError:
                pass
    # Header buttons are local variables, so find them by walking the tree.
    # Buttons are keyed by label so the same button can be compared across a
    # language switch even though its text changes; the index disambiguates
    # repeats such as the per-screen "Clear Log".
    seen: dict[str, int] = {}
    for child in app.winfo_children():
        for grandchild in _all_widgets(child):
            if isinstance(grandchild, tk.Button):
                try:
                    role = str(grandchild.cget("bg"))  # primary vs secondary
                    seen[role] = seen.get(role, 0) + 1
                    snapshot[f"btn#{seen[role]}:{role}"] = geometry_of(grandchild)
                except tk.TclError:
                    pass
    return snapshot


def _all_widgets(widget: tk.Misc):
    """Every descendant of ``widget``, depth first."""
    for child in widget.winfo_children():
        yield child
        yield from _all_widgets(child)


def wait_until_settled(app: WaybillSorterApp, seconds: float = 1.5) -> None:
    """Let any pending after() callbacks (e.g. show_maximized) run.

    Without this the measurement catches the window mid-maximise and the
    language delta is confounded by the startup behaviour.
    """
    import time

    deadline = time.time() + seconds
    while time.time() < deadline:
        app.update()
        time.sleep(0.05)


def main() -> int:
    app = WaybillSorterApp()
    wait_until_settled(app)

    print(f"window state after startup: {app.state()!r}")
    print()
    print("=== main screen (workflow selector) ===")
    english = collect(app)

    app.toggle_language()
    app.update()
    app.update_idletasks()
    wait_until_settled(app, 0.6)
    chinese = collect(app)
    print(f"window state after toggle : {app.state()!r}")
    print()

    width_deltas = []
    for key in sorted(set(english) | set(chinese)):
        before = english.get(key)
        after = chinese.get(key)
        marker = ""
        if before and after and before[:2] != after[:2]:
            marker = f"   <-- CHANGED ({before[0]}x{before[1]} -> {after[0]}x{after[1]})"
            width_deltas.append((key, before, after))
        print(f"{key:34} {str(before):>26}{marker if marker else ''}")

    print()
    if width_deltas:
        print(f"{len(width_deltas)} element(s) change size between languages:")
        for key, before, after in width_deltas:
            dw = after[0] - before[0]
            print(f"  {key:32} width {dw:+d}px")
    else:
        print("No measured element changes size.")

    # Buttons of the same kind must be the same width as each other, otherwise
    # they look ragged and the column is sized by whichever label is longest.
    print()
    by_role: dict[str, set[int]] = {}
    for key, geometry in english.items():
        if key.startswith("btn#"):
            by_role.setdefault(key.split(":", 1)[1], set()).add(geometry[0])
    for role, widths in sorted(by_role.items()):
        label = "primary (#0B5CAD)" if role.lower() == "#0b5cad" else f"secondary ({role})"
        if len(widths) == 1:
            print(f"  uniform: all {label} buttons are {widths.pop()}px wide")
        else:
            print(f"  RAGGED : {label} buttons differ: {sorted(widths)}")

    # The window itself is the thing the user notices most.
    print()
    print(f"window english : {english['<window>'][0]} x {english['<window>'][1]} at ({english['<window>'][2]},{english['<window>'][3]})")
    print(f"window chinese : {chinese['<window>'][0]} x {chinese['<window>'][1]} at ({chinese['<window>'][2]},{chinese['<window>'][3]})")
    print(f"requested english: {english['<reqsize>'][0]} x {english['<reqsize>'][1]}")
    print(f"requested chinese: {chinese['<reqsize>'][0]} x {chinese['<reqsize>'][1]}")

    app.destroy()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
