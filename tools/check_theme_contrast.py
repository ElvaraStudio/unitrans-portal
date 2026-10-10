"""Check every foreground/background pair against the WCAG contrast ratio.

Dark mode is easy to get subtly wrong: text that is *nearly* the background
colour looks acceptable on the monitor it was chosen on and is unreadable on
someone else's. This computes the actual contrast ratios, so "is it readable?"
is measured rather than judged by eye.

Thresholds (WCAG 2.1):
  * 4.5:1 - normal body text
  * 3.0:1 - large text (18pt+, or 14pt+ bold) and UI component boundaries

Run from the project root:

    python tools/check_theme_contrast.py
    python tools/check_theme_contrast.py --scheme dark --verbose
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from theme import SCHEMES  # noqa: E402

#: (foreground token, background token, minimum ratio, description)
#: The threshold is 3.0 where the text is large or bold, and for the
#: boundaries of controls, which WCAG treats as non-text contrast.
PAIRS: list[tuple[str, str, float, str]] = [
    # Body text on the page and on cards
    ("text", "window", 4.5, "body text on the page background"),
    ("text", "surface", 4.5, "body text on a card"),
    ("text", "surface_alt", 4.5, "log text on the log panel"),
    ("text", "list_bg", 4.5, "text in a read-only path box"),
    ("text_entry", "list_bg", 4.5, "path text in a path box"),

    # Headings are large and bold
    ("text_strong", "surface", 4.5, "window/section title"),
    ("text_heading", "surface", 4.5, "card heading"),
    ("text_muted", "surface", 4.5, "subtitle and notes"),
    ("text_muted", "window", 4.5, "subtitle on the page background"),
    ("text_faint", "window", 4.5, "footer version text"),
    ("text_label", "surface", 4.5, "stat labels"),

    # Accent. "on_accent" is white text on an accent-filled button, so it only
    # applies where the token really is a button background.
    ("on_accent", "accent", 4.5, "label on a primary button"),
    ("on_accent", "accent_hover", 4.5, "label on a hovered primary button"),
    ("on_accent", "accent_disabled", 2.4, "label on a disabled primary button"),
    ("text_heading", "accent_soft", 4.5, "label on a secondary button"),

    # Accent used as a colour rather than a fill.
    ("accent_text", "surface", 4.5, "accent used as text on a card"),
    ("accent_text", "window", 4.5, "accent used as text on the page"),

    # State text - these are text colours on a card, not button fills.
    ("success", "surface", 4.5, "success message"),
    ("error", "surface", 4.5, "error message"),
    ("warning", "danger_bg", 4.5, "warning banner text"),

    # Structure - non-text contrast. The thresholds are modest on purpose:
    # these are modern flat-UI outlines that are meant to be barely visible, so
    # the bar is "perceptible", not "high contrast".
    ("border", "surface", 1.15, "card outline against the card"),
    ("border", "window", 1.10, "card outline against the page"),
    ("divider", "surface", 1.10, "separator against a card"),
    ("scrollbar", "scrollbar_trough", 1.10, "scrollbar against its trough"),
]


def _srgb_to_linear(channel: int) -> float:
    value = channel / 255.0
    return value / 12.92 if value <= 0.04045 else ((value + 0.055) / 1.055) ** 2.4


def relative_luminance(hex_colour: str) -> float:
    text = hex_colour.lstrip("#")
    red, green, blue = (int(text[index:index + 2], 16) for index in (0, 2, 4))
    return (
        0.2126 * _srgb_to_linear(red)
        + 0.7152 * _srgb_to_linear(green)
        + 0.0722 * _srgb_to_linear(blue)
    )


def contrast_ratio(foreground: str, background: str) -> float:
    light = max(relative_luminance(foreground), relative_luminance(background))
    dark = min(relative_luminance(foreground), relative_luminance(background))
    return (light + 0.05) / (dark + 0.05)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--scheme", choices=sorted(SCHEMES), help="Check one scheme only.")
    parser.add_argument("--verbose", action="store_true", help="List every pair, not just failures.")
    args = parser.parse_args(argv)

    schemes = [args.scheme] if args.scheme else sorted(SCHEMES)
    failures = 0

    for scheme in schemes:
        table = SCHEMES[scheme]
        print(f"=== {scheme} ===")
        for foreground_token, background_token, minimum, description in PAIRS:
            if foreground_token not in table or background_token not in table:
                print(f"  MISSING TOKEN: {foreground_token} or {background_token}")
                failures += 1
                continue
            ratio = contrast_ratio(table[foreground_token], table[background_token])
            ok = ratio >= minimum
            if not ok:
                failures += 1
            if args.verbose or not ok:
                mark = "ok  " if ok else "FAIL"
                print(
                    f"  {mark} {ratio:5.2f}:1 (min {minimum:.1f})  {description}"
                    f"  [{foreground_token} on {background_token}]"
                )
        checked = len(PAIRS)
        print(f"  {checked} pairs checked")
        print()

    if failures:
        print(f"{failures} contrast check(s) failed.")
        return 1
    print("All contrast checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
