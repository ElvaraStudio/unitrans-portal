"""Check that buttons are equal-sized and that nothing resizes on a language switch.

Two separate properties, because they are easy to confuse:

1. **Uniformity** - buttons that form a group (the workflow cards, the action
   column) must all be the same width. Previously each sized to its own label,
   so "Unlock PDFs" was narrower than "Fill Driver Numbers".
2. **Stability** - switching language must not change any button's width, and
   must not change the window's size or position. Buttons are matched *by
   position in the widget tree*, not by label, because the label is exactly
   what changes.

    python tests/measure_buttons.py
"""

from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from waybill_sorter import WaybillSorterApp  # noqa: E402


def all_widgets(widget: tk.Misc):
    """Every descendant, depth first - a stable ordering across rebuilds."""
    for child in widget.winfo_children():
        yield child
        yield from all_widgets(child)


def button_widths(app: tk.Misc) -> list[tuple[str, int, int]]:
    """``(text, width, height)`` for every mapped button, in tree order."""
    result = []
    for widget in all_widgets(app):
        if isinstance(widget, tk.Button):
            try:
                if widget.winfo_ismapped():
                    result.append(
                        (str(widget.cget("text")), widget.winfo_width(), widget.winfo_height())
                    )
            except tk.TclError:
                pass
    return result


def settle(app: tk.Misc, seconds: float = 0.8) -> None:
    import time

    deadline = time.time() + seconds
    while time.time() < deadline:
        app.update()
        time.sleep(0.04)


def screen_buttons(app: WaybillSorterApp, screen: str) -> list[tuple[str, int, int]]:
    """Open ``screen`` and return its buttons."""
    if screen == "selector":
        app.return_to_workflows()
    else:
        {
            "sort": app.select_sort_workflow,
            "driver": app.select_driver_workflow,
            "unlock": app.select_unlock_workflow,
        }[screen]()
    settle(app, 0.5)
    return button_widths(app)


def group_widths(buttons: list[tuple[str, int, int]], exclude: set[str]) -> list[int]:
    """Widths of buttons other than the header's (whose labels differ anyway)."""
    return [width for text, width, _h in buttons if text not in exclude]


def main() -> int:
    app = WaybillSorterApp()
    settle(app, 1.2)

    failures: list[str] = []

    # Header buttons have genuinely different labels ("中文" vs "Settings"), so
    # they are excluded from the uniformity check but still included in the
    # stability check.
    header_labels = {app.tr("settings"), app.tr("language"), app.tr("back_to_workflows")}
    header_labels |= {f"\u2190  {app.tr('back_to_workflows')}"}

    print("=== uniformity: buttons in a group must share one width ===")
    # Two visual groups per screen: the tall action buttons, and the file
    # pickers above the path entry boxes (which are a shorter, narrower button
    # by design). Each group must be internally uniform, and the pickers must
    # also match the secondary button beneath them - a third distinct width
    # means some button is sizing to its own label again.
    for screen in ("selector", "sort", "driver", "unlock"):
        buttons = screen_buttons(app, screen)
        widths = group_widths(buttons, header_labels)
        distinct = sorted(set(widths))
        counts: dict[int, int] = {}
        for width in widths:
            counts[width] = counts.get(width, 0) + 1
        summary = ", ".join(f"{count}x{width}px" for width, count in sorted(counts.items()))
        print(f"  {screen:9} groups: {summary}")
        if len(distinct) > 2:
            failures.append(
                f"{screen}: {len(distinct)} distinct button widths {distinct} - "
                "expected at most two (action buttons, path pickers)"
            )

    print()
    print("=== stability: a language switch must change nothing ===")
    for screen in ("selector", "sort", "driver", "unlock"):
        app.return_to_workflows()
        settle(app, 0.4)
        if screen != "selector":
            {
                "sort": app.select_sort_workflow,
                "driver": app.select_driver_workflow,
                "unlock": app.select_unlock_workflow,
            }[screen]()
            settle(app, 0.5)

        before = button_widths(app)
        window_before = (app.winfo_width(), app.winfo_height(), app.winfo_x(), app.winfo_y())

        app.toggle_language()
        settle(app, 0.8)

        after = button_widths(app)
        window_after = (app.winfo_width(), app.winfo_height(), app.winfo_x(), app.winfo_y())

        if window_before != window_after:
            print(f"  {screen:9} WINDOW MOVED {window_before} -> {window_after}")
            failures.append(f"{screen}: window changed {window_before} -> {window_after}")
        elif len(before) != len(after):
            print(f"  {screen:9} BUTTON COUNT CHANGED {len(before)} -> {len(after)}")
            failures.append(f"{screen}: button count changed")
        else:
            # The header buttons legitimately change width - "中文" and
            # "Settings" are different words - so they are excluded here too.
            changed = [
                (b[0], b[1], a[1])
                for b, a in zip(before, after)
                if b[1] != a[1] and b[0] not in header_labels
            ]
            if changed:
                print(f"  {screen:9} RESIZED {changed}")
                failures.append(f"{screen}: buttons resized on language switch {changed}")
            else:
                print(f"  {screen:9} OK    {len(before)} buttons unchanged, "
                      f"window {window_after[0]}x{window_after[1]} at "
                      f"({window_after[2]},{window_after[3]})")

        # Leave the app in English for the next screen's comparison.
        if app.language.get() != "en":
            app.toggle_language()
            settle(app, 0.4)

    app.destroy()

    print()
    if failures:
        for failure in failures:
            print(f"FAIL: {failure}")
        return 1
    print("PASS: buttons are uniform within each group and nothing resizes "
          "when the language changes.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
