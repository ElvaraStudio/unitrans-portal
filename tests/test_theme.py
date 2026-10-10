"""Tests for the colour schemes.

The theme layer is small but easy to get wrong in ways that only show up
visually: a token defined in one scheme and not the other, a foreground colour
that is too close to its background to read, or a lookup that silently returns
the wrong scheme's value.

Contrast is computed here rather than eyeballed, because "is this readable?" is
the one question a screenshot cannot answer reliably across monitors.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import theme  # noqa: E402


def _linear(channel: int) -> float:
    value = channel / 255.0
    return value / 12.92 if value <= 0.04045 else ((value + 0.055) / 1.055) ** 2.4


def luminance(hex_colour: str) -> float:
    text = hex_colour.lstrip("#")
    red, green, blue = (int(text[index:index + 2], 16) for index in (0, 2, 4))
    return 0.2126 * _linear(red) + 0.7152 * _linear(green) + 0.0722 * _linear(blue)


def contrast(foreground: str, background: str) -> float:
    high = max(luminance(foreground), luminance(background))
    low = min(luminance(foreground), luminance(background))
    return (high + 0.05) / (low + 0.05)


#: (foreground, background, minimum) - the pairs a user actually reads.
READABLE_PAIRS = [
    ("text", "window", 4.5),
    ("text", "surface", 4.5),
    ("text", "surface_alt", 4.5),
    ("text_entry", "list_bg", 4.5),
    ("text_strong", "surface", 4.5),
    ("text_heading", "surface", 4.5),
    ("text_muted", "surface", 4.5),
    ("text_muted", "window", 4.5),
    ("text_faint", "window", 4.5),
    ("text_label", "surface", 4.5),
    ("on_accent", "accent", 4.5),
    ("on_accent", "accent_hover", 4.5),
    ("text_heading", "accent_soft", 4.5),
    ("accent_text", "surface", 4.5),
    ("success", "surface", 4.5),
    ("error", "surface", 4.5),
    ("warning", "danger_bg", 4.5),
]


class TestSchemeIntegrity(unittest.TestCase):
    def test_both_schemes_exist(self):
        self.assertEqual(set(theme.SCHEMES), {theme.LIGHT, theme.DARK})

    def test_schemes_define_identical_tokens(self):
        # A token present in one scheme and missing from the other is the
        # classic dark-mode bug: the widget keeps its light colour and is
        # invisible.
        light = set(theme.LIGHT_COLOURS)
        dark = set(theme.DARK_COLOURS)
        self.assertEqual(light - dark, set(), "missing from the dark scheme")
        self.assertEqual(dark - light, set(), "missing from the light scheme")

    def test_every_value_is_a_hex_colour(self):
        for name, scheme in theme.SCHEMES.items():
            for token, value in scheme.items():
                with self.subTest(scheme=name, token=token):
                    self.assertRegex(value, r"^#[0-9A-Fa-f]{6}$")

    def test_schemes_are_actually_different(self):
        # A copy-paste that left dark identical to light would otherwise pass.
        self.assertNotEqual(theme.LIGHT_COLOURS, theme.DARK_COLOURS)

    def test_dark_surfaces_are_dark(self):
        for token in ("window", "surface", "surface_alt", "list_bg"):
            with self.subTest(token=token):
                self.assertLess(
                    luminance(theme.DARK_COLOURS[token]),
                    0.2,
                    f"{token} is not dark in the dark scheme",
                )

    def test_light_surfaces_are_light(self):
        for token in ("window", "surface", "surface_alt"):
            with self.subTest(token=token):
                self.assertGreater(
                    luminance(theme.LIGHT_COLOURS[token]),
                    0.7,
                    f"{token} is not light in the light scheme",
                )


class TestContrast(unittest.TestCase):
    def test_readable_pairs_meet_wcag(self):
        for scheme_name, scheme in theme.SCHEMES.items():
            for foreground, background, minimum in READABLE_PAIRS:
                with self.subTest(scheme=scheme_name, pair=f"{foreground}/{background}"):
                    ratio = contrast(scheme[foreground], scheme[background])
                    self.assertGreaterEqual(
                        ratio,
                        minimum,
                        f"{scheme_name}: {foreground} on {background} is only "
                        f"{ratio:.2f}:1, needs {minimum}:1",
                    )

    def test_dark_accent_is_lighter_than_light_accent(self):
        # White text on the light-mode blue only reaches about 3.9:1, so the
        # dark scheme needs a lighter accent to stay readable.
        self.assertGreater(
            luminance(theme.DARK_COLOURS["accent"]),
            luminance(theme.LIGHT_COLOURS["accent"]),
        )


class TestLookup(unittest.TestCase):
    def tearDown(self):
        theme.set_active(theme.DEFAULT_THEME)

    def test_default_is_light(self):
        self.assertEqual(theme.DEFAULT_THEME, theme.LIGHT)
        theme.set_active(None)
        self.assertEqual(theme.active(), theme.LIGHT)

    def test_unknown_scheme_falls_back_rather_than_raising(self):
        # A corrupt settings file must not stop the app from starting.
        for value in ("", "blue", None, "LIGHT", "  "):
            with self.subTest(value=value):
                self.assertEqual(theme.normalise(value), theme.DEFAULT_THEME)

    def test_colour_returns_the_active_scheme(self):
        theme.set_active(theme.DARK)
        self.assertEqual(theme.colour("window"), theme.DARK_COLOURS["window"])
        theme.set_active(theme.LIGHT)
        self.assertEqual(theme.colour("window"), theme.LIGHT_COLOURS["window"])

    def test_colour_of_an_unknown_token_raises(self):
        # A typo in a token name is a code bug, not user input, so it should be
        # loud rather than silently returning a default colour.
        with self.assertRaises(theme.ThemeError):
            theme.colour("no_such_token")

    def test_is_dark(self):
        self.assertTrue(theme.is_dark(theme.DARK))
        self.assertFalse(theme.is_dark(theme.LIGHT))
        self.assertFalse(theme.is_dark("nonsense"))


class TestSourceUsesTokens(unittest.TestCase):
    """The app must not reintroduce hardcoded colours."""

    def test_no_colour_literals_in_styling_attributes(self):
        import re

        source = (ROOT / "waybill_sorter.py").read_text(encoding="utf-8")
        attributes = (
            "bg|fg|activebackground|activeforeground|disabledforeground|"
            "readonlybackground|selectcolor|highlightbackground|highlightcolor|"
            "background|foreground"
        )
        pattern = re.compile(
            rf"\b(?:{attributes})\s*=\s*[\"']#[0-9A-Fa-f]{{6}}[\"']"
        )
        found = pattern.findall(source)
        self.assertEqual(
            found,
            [],
            "hardcoded colours found; use a theme token so dark mode works: "
            f"{found}",
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
