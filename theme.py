"""Light and dark colour schemes, as named tokens.

Every colour in the interface comes from here. Widgets ask for a token by name
(``theme()["surface"]``) rather than hardcoding a hex value, which is what makes
a dark mode possible at all: before this, 21 different hex colours were written
directly into 140 places in ``waybill_sorter.py``.

Adding a token
--------------
A token must be defined in **both** schemes. ``assert_complete()`` fails at
import time otherwise, which turns "I forgot to add the dark value" from a
black-on-black widget discovered by a user into an error on the developer's
machine.

Contrast
--------
Foreground/background pairs are checked by ``tools/check_theme_contrast.py``
against the WCAG contrast ratio. Dark mode is easy to get subtly wrong - text
that is *nearly* background-coloured looks fine on the developer's monitor and
unreadable on someone else's - so the pairs are verified numerically rather
than by eye.
"""

from __future__ import annotations

LIGHT = "light"
DARK = "dark"
THEMES = (LIGHT, DARK)
DEFAULT_THEME = LIGHT


#: The light scheme. These are the original colours; the interface looked like
#: this before dark mode existed.
LIGHT_COLOURS: dict[str, str] = {
    # --- surfaces -------------------------------------------------------
    "window": "#E8EDF4",          # page background behind the cards
    "surface": "#FFFFFF",         # cards, header, dialogs
    "surface_alt": "#F8FAFC",     # log panel, notes pane
    "list_bg": "#F6F8FB",         # read-only path boxes
    "menu": "#E8EDF4",            # menubar background
    "danger_bg": "#FFF7ED",       # warning banner background

    # --- text -----------------------------------------------------------
    "text": "#0F172A",            # primary body text
    "text_strong": "#082B52",     # headings and titles
    "text_heading": "#12355F",    # section and card headings
    "text_muted": "#4B5F7A",      # subtitles, notes
    # Darker than the original #6B7280, which only reached 4.11:1 on the page
    # background - below the 4.5:1 needed for readable text.
    "text_faint": "#5B6470",      # footer, hints
    "text_label": "#627084",      # stat labels
    "text_entry": "#111827",      # text inside path boxes
    "on_accent": "#FFFFFF",       # text on an accent-coloured button

    # --- accent (primary buttons, focus, links) -------------------------
    "accent": "#0B5CAD",
    "accent_hover": "#084A8C",
    "accent_disabled": "#6F9FCE",
    "accent_soft": "#EAF1F8",     # secondary button background
    "accent_soft_hover": "#DDE8F3",
    "accent_text": "#0B5CAD",     # accent used as text on a card

    # --- structure ------------------------------------------------------
    "border": "#D5DCE7",          # card outlines
    "divider": "#D5DCE7",         # separators
    "scrollbar": "#C8D2E0",
    "scrollbar_trough": "#F1F5F9",

    # --- state ----------------------------------------------------------
    "success": "#1B7F3B",
    "error": "#B3261E",
    "warning": "#9A3412",
    "warning_border": "#FDBA74",
}


#: The dark scheme. Cool and blue-leaning, so it reads as the same product
#: rather than a generic grey inversion.
DARK_COLOURS: dict[str, str] = {
    # --- surfaces -------------------------------------------------------
    "window": "#0E1523",
    "surface": "#18243A",
    "surface_alt": "#131C2E",
    "list_bg": "#101828",
    "menu": "#131C2E",
    "danger_bg": "#2A1A1A",

    # --- text -----------------------------------------------------------
    "text": "#E6EDF6",
    "text_strong": "#F3F7FC",
    "text_heading": "#DCE7F5",
    "text_muted": "#9FB0C7",
    "text_faint": "#8497B0",
    "text_label": "#8FA1BA",
    "text_entry": "#DCE7F5",
    "on_accent": "#FFFFFF",

    # --- accent ---------------------------------------------------------
    # Lighter than the light-mode blue. On a dark surface the original #0B5CAD
    # is too low-contrast to read as a button, and white text on it only
    # reaches 3.9:1 - short of the 4.5:1 needed for body text. This shade
    # clears it while still reading as a button rather than a pastel.
    "accent": "#2F6FBF",
    "accent_hover": "#245C9F",
    "accent_disabled": "#2C4A6E",
    "accent_soft": "#22314B",
    "accent_soft_hover": "#2A3D5C",
    "accent_text": "#7FB2F0",     # accent used as text on a card

    # --- structure ------------------------------------------------------
    "border": "#2C3A52",
    "divider": "#2C3A52",
    "scrollbar": "#3A4A63",
    "scrollbar_trough": "#141D2E",

    # --- state ----------------------------------------------------------
    "success": "#4ADE80",
    "error": "#FF6B6B",
    "warning": "#FDBA74",
    "warning_border": "#7C4A21",
}


SCHEMES: dict[str, dict[str, str]] = {
    LIGHT: LIGHT_COLOURS,
    DARK: DARK_COLOURS,
}


class ThemeError(RuntimeError):
    """A scheme is missing a token, or a token name is unknown."""


def _assert_complete() -> None:
    """Every scheme must define exactly the same tokens."""
    reference = set(LIGHT_COLOURS)
    for name, scheme in SCHEMES.items():
        missing = reference - set(scheme)
        extra = set(scheme) - reference
        if missing or extra:
            raise ThemeError(
                f"scheme {name!r} does not match the light scheme: "
                f"missing={sorted(missing)} extra={sorted(extra)}"
            )


_assert_complete()


def normalise(name: str | None) -> str:
    """Map anything to a valid scheme name, defaulting rather than raising."""
    return name if name in SCHEMES else DEFAULT_THEME


def palette(name: str | None = None) -> dict[str, str]:
    """The colour tokens for a scheme."""
    return SCHEMES[normalise(name)]


def is_dark(name: str | None) -> bool:
    return normalise(name) == DARK


#: Set by the application at startup and whenever the user switches theme, so
#: module-level helpers and widget factories can read colours without every
#: call site having to thread the current scheme through.
_active: str = DEFAULT_THEME


def set_active(name: str | None) -> str:
    global _active
    _active = normalise(name)
    return _active


def active() -> str:
    return _active


def colour(token: str, scheme: str | None = None) -> str:
    """Look up one token. Raises on an unknown token - that is a code bug."""
    table = palette(scheme if scheme is not None else _active)
    try:
        return table[token]
    except KeyError:
        raise ThemeError(
            f"unknown colour token {token!r}; known tokens: {sorted(table)}"
        ) from None


def colours(scheme: str | None = None) -> dict[str, str]:
    """Shorthand for the active scheme, for call sites that need several."""
    return palette(scheme if scheme is not None else _active)
