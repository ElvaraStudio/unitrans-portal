"""One-off: replace hardcoded hex colours with theme token lookups.

Every colour in ``waybill_sorter.py`` was written as a literal. Dark mode needs
one switch to change all of them, so each literal is replaced by a token
constant that is read from ``theme`` at startup.

The replacement is deliberately literal-for-literal, applied only as a *value*
(``bg="#FFFFFF"`` and friends). It never touches a hex string used as data - the
translation table contains colour-like words and the waybill parser handles
arbitrary text - because only these specific 21 values appear as styling, and
each occurrence is checked to sit in a styling attribute.

Run from the project root:

    python tools/apply_theme_tokens.py --dry-run
    python tools/apply_theme_tokens.py
"""

from __future__ import annotations

import argparse
import ast
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TARGET = ROOT / "waybill_sorter.py"

#: hex value -> (token constant name, theme token)
#:
#: #FFFFFF appears as both a background (cards) and a foreground (text on an
#: accent button). Each occurrence is inspected to decide which is meant.
REPLACEMENTS: dict[str, tuple[str, str]] = {
    "#E8EDF4": ("BG_WINDOW", "window"),
    "#F8FAFC": ("BG_PANEL", "surface_alt"),
    "#F6F8FB": ("BG_LIST", "list_bg"),
    "#EAF1F8": ("ACCENT_SOFT", "accent_soft"),
    "#DDE8F3": ("ACCENT_SOFT_HOVER", "accent_soft_hover"),
    "#D5DCE7": ("BORDER", "border"),
    "#0B5CAD": ("ACCENT", "accent"),
    "#084A8C": ("ACCENT_HOVER", "accent_hover"),
    "#6F9FCE": ("ACCENT_DISABLED", "accent_disabled"),
    "#082B52": ("TEXT_STRONG", "text_strong"),
    "#12355F": ("TEXT_HEADING", "text_heading"),
    "#4B5F7A": ("TEXT_MUTED", "text_muted"),
    "#6B7280": ("TEXT_FAINT", "text_faint"),
    "#627084": ("TEXT_LABEL", "text_label"),
    "#0F172A": ("TEXT", "text"),
    "#111827": ("TEXT_ENTRY", "text_entry"),
    "#1B7F3B": ("STATE_SUCCESS", "success"),
    "#B3261E": ("STATE_ERROR", "error"),
    "#9A3412": ("STATE_WARNING", "warning"),
    "#FFF7ED": ("BG_WARNING", "danger_bg"),
    "#FDBA74": ("BORDER_WARNING", "warning_border"),
}

#: Attributes whose value is a colour.
COLOUR_ATTRS = (
    "bg",
    "fg",
    "activebackground",
    "activeforeground",
    "disabledforeground",
    "readonlybackground",
    "selectcolor",
    "highlightbackground",
    "highlightcolor",
    "background",
    "foreground",
)

ATTR_PATTERN = re.compile(
    r"\b(" + "|".join(COLOUR_ATTRS) + r")\s*=\s*(?P<quote>[\"'])(?P<colour>#[0-9A-Fa-f]{6})(?P=quote)"
)

#: Attributes that take a *background* colour.
BACKGROUND_ATTRS = frozenset({
    "bg",
    "activebackground",
    "readonlybackground",
    "selectcolor",
    "highlightbackground",
    "highlightcolor",
    "background",
})

TOKEN_CONSTANTS = """# ---------------------------------------------------------------------------
# Colour tokens. Read from theme.py at startup and whenever the user switches
# between light and dark, so no widget needs to know which scheme is in use.
# Reassigning these and rebuilding the window is the whole of a theme change.
# ---------------------------------------------------------------------------
def _load_colour_tokens() -> dict[str, str]:
    \"\"\"Current value of every colour token, keyed by its constant name.\"\"\"
    return {
        "BG_WINDOW": theme.colour("window"),
        "BG_PANEL": theme.colour("surface_alt"),
        "BG_LIST": theme.colour("list_bg"),
        "BG_WARNING": theme.colour("danger_bg"),
        "SURFACE": theme.colour("surface"),
        "ON_ACCENT": theme.colour("on_accent"),
        "ACCENT": theme.colour("accent"),
        "ACCENT_HOVER": theme.colour("accent_hover"),
        "ACCENT_DISABLED": theme.colour("accent_disabled"),
        "ACCENT_SOFT": theme.colour("accent_soft"),
        "ACCENT_SOFT_HOVER": theme.colour("accent_soft_hover"),
        "ACCENT_TEXT": theme.colour("accent_text"),
        "TEXT": theme.colour("text"),
        "TEXT_STRONG": theme.colour("text_strong"),
        "TEXT_HEADING": theme.colour("text_heading"),
        "TEXT_MUTED": theme.colour("text_muted"),
        "TEXT_FAINT": theme.colour("text_faint"),
        "TEXT_LABEL": theme.colour("text_label"),
        "TEXT_ENTRY": theme.colour("text_entry"),
        "BORDER": theme.colour("border"),
        "DIVIDER": theme.colour("divider"),
        "BORDER_WARNING": theme.colour("warning_border"),
        "SCROLLBAR": theme.colour("scrollbar"),
        "SCROLLBAR_TROUGH": theme.colour("scrollbar_trough"),
        "STATE_SUCCESS": theme.colour("success"),
        "STATE_ERROR": theme.colour("error"),
        "STATE_WARNING": theme.colour("warning"),
    }


#: Populated by _load_colour_tokens(); see apply_theme().
(
    BG_WINDOW, BG_PANEL, BG_LIST, BG_WARNING, SURFACE, ON_ACCENT,
    ACCENT, ACCENT_HOVER, ACCENT_DISABLED, ACCENT_SOFT, ACCENT_SOFT_HOVER,
    ACCENT_TEXT, TEXT, TEXT_STRONG, TEXT_HEADING, TEXT_MUTED, TEXT_FAINT,
    TEXT_LABEL, TEXT_ENTRY, BORDER, DIVIDER, BORDER_WARNING, SCROLLBAR,
    SCROLLBAR_TROUGH, STATE_SUCCESS, STATE_ERROR, STATE_WARNING,
) = _load_colour_tokens().values()


def apply_theme(name: str) -> None:
    \"\"\"Switch the active scheme and refresh the token constants.\"\"\"
    global BG_WINDOW, BG_PANEL, BG_LIST, BG_WARNING, SURFACE, ON_ACCENT
    global ACCENT, ACCENT_HOVER, ACCENT_DISABLED, ACCENT_SOFT, ACCENT_SOFT_HOVER
    global ACCENT_TEXT, TEXT, TEXT_STRONG, TEXT_HEADING, TEXT_MUTED, TEXT_FAINT
    global TEXT_LABEL, TEXT_ENTRY, BORDER, DIVIDER, BORDER_WARNING
    global SCROLLBAR, SCROLLBAR_TROUGH, STATE_SUCCESS, STATE_ERROR, STATE_WARNING

    theme.set_active(name)
    tokens = _load_colour_tokens()
    BG_WINDOW = tokens["BG_WINDOW"]
    BG_PANEL = tokens["BG_PANEL"]
    BG_LIST = tokens["BG_LIST"]
    BG_WARNING = tokens["BG_WARNING"]
    SURFACE = tokens["SURFACE"]
    ON_ACCENT = tokens["ON_ACCENT"]
    ACCENT = tokens["ACCENT"]
    ACCENT_HOVER = tokens["ACCENT_HOVER"]
    ACCENT_DISABLED = tokens["ACCENT_DISABLED"]
    ACCENT_SOFT = tokens["ACCENT_SOFT"]
    ACCENT_SOFT_HOVER = tokens["ACCENT_SOFT_HOVER"]
    ACCENT_TEXT = tokens["ACCENT_TEXT"]
    TEXT = tokens["TEXT"]
    TEXT_STRONG = tokens["TEXT_STRONG"]
    TEXT_HEADING = tokens["TEXT_HEADING"]
    TEXT_MUTED = tokens["TEXT_MUTED"]
    TEXT_FAINT = tokens["TEXT_FAINT"]
    TEXT_LABEL = tokens["TEXT_LABEL"]
    TEXT_ENTRY = tokens["TEXT_ENTRY"]
    BORDER = tokens["BORDER"]
    DIVIDER = tokens["DIVIDER"]
    BORDER_WARNING = tokens["BORDER_WARNING"]
    SCROLLBAR = tokens["SCROLLBAR"]
    SCROLLBAR_TROUGH = tokens["SCROLLBAR_TROUGH"]
    STATE_SUCCESS = tokens["STATE_SUCCESS"]
    STATE_ERROR = tokens["STATE_ERROR"]
    STATE_WARNING = tokens["STATE_WARNING"]

"""


def token_for(attribute: str, colour: str) -> str | None:
    """Constant name for a colour used in a given attribute, or None."""
    if colour.upper() == "#FFFFFF":
        # White is a background on cards and the text colour on a filled
        # accent button; the attribute says which is meant.
        return "SURFACE" if attribute.lower() in BACKGROUND_ATTRS else "ON_ACCENT"

    if colour.upper() == "#0B5CAD":
        # The accent doubles as a text colour in the Settings window, where it
        # marks an available update; dark mode needs a lighter shade there.
        return "ACCENT" if attribute.lower() in BACKGROUND_ATTRS else "ACCENT_TEXT"

    entry = REPLACEMENTS.get(colour.upper())
    return entry[0] if entry else None


def rewrite(source: str) -> tuple[str, int, list[str]]:
    count = 0
    unmapped: list[str] = []

    def replacement(match: re.Match[str]) -> str:
        nonlocal count
        attribute = match.group(1)
        quote = match.group("quote")
        colour = match.group("colour")
        name = token_for(attribute, colour)
        if name is None:
            unmapped.append(f"{attribute}={colour}")
            return match.group(0)
        count += 1
        return f"{attribute}={name}"

    result = ATTR_PATTERN.sub(replacement, source)
    return result, count, unmapped


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args(argv)

    source = TARGET.read_text(encoding="utf-8")
    ast.parse(source)

    result, count, unmapped = rewrite(source)
    if unmapped:
        print("Colours with no token mapping (not replaced):")
        for item in sorted(set(unmapped)):
            print("  ", item)

    remaining = ATTR_PATTERN.findall(result)
    print(f"{count} colour values replaced; {len(remaining)} styling literals remain")

    if remaining:
        for attribute, _quote, colour in sorted(set(remaining)):
            print(f"   left behind: {attribute}={colour}")

    ast.parse(result)

    if args.dry_run:
        print("Dry run - nothing written.")
        return 0

    TARGET.write_text(result, encoding="utf-8")
    print(f"Wrote {TARGET.name}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
