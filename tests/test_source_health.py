"""Guards against source-level defects that tests would not otherwise catch.

These check the *shape* of the source rather than its behaviour. They exist
because of a real problem found in this codebase: ``waybill_sorter.py`` had
accumulated two generations of UI code, and five methods were defined twice.
Python keeps only the last definition, so the earlier copies were unreachable -
they could not be called, could not be fixed, and existed only to mislead
whoever edited the file next. Nothing failed; the code just quietly did not do
what it appeared to do.
"""

from __future__ import annotations

import ast
import sys
import unittest
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

PYTHON_SOURCES = ("waybill_sorter.py", "version.py", "updater.py", "settings.py", "releasectl.py", "make_assets.py")


class TestNoShadowedMethods(unittest.TestCase):
    def test_no_class_defines_a_method_twice(self):
        for name in PYTHON_SOURCES:
            path = ROOT / name
            if not path.exists():
                continue
            tree = ast.parse(path.read_text(encoding="utf-8"))
            for node in ast.walk(tree):
                if not isinstance(node, ast.ClassDef):
                    continue
                seen: dict[str, list[int]] = defaultdict(list)
                for item in node.body:
                    if isinstance(item, ast.FunctionDef):
                        seen[item.name].append(item.lineno)
                shadowed = {key: lines for key, lines in seen.items() if len(lines) > 1}
                with self.subTest(file=name, cls=node.name):
                    self.assertEqual(
                        shadowed,
                        {},
                        f"{name}:{node.name} defines these methods more than once, so all "
                        f"but the last are dead code: {shadowed}",
                    )


class TestImportsResolve(unittest.TestCase):
    def test_first_party_modules_are_importable(self):
        # PyInstaller bundles these by following imports, so a typo here would
        # surface as a crash inside the frozen exe rather than at build time.
        import settings  # noqa: F401
        import updater  # noqa: F401
        import version  # noqa: F401

    def test_app_imports_without_a_display(self):
        # Importing must not construct widgets or open a window; otherwise
        # --check-updates and the CLI paths would fail on a headless machine.
        import waybill_sorter

        for symbol in ("main", "parse_args", "WaybillSorterApp"):
            self.assertTrue(hasattr(waybill_sorter, symbol), f"missing {symbol}")


class TestCliSurface(unittest.TestCase):
    def test_check_updates_flag_exists(self):
        import waybill_sorter

        args = waybill_sorter.parse_args(["--check-updates"])
        self.assertTrue(args.check_updates)

    def test_check_updates_is_off_by_default(self):
        import waybill_sorter

        args = waybill_sorter.parse_args([])
        self.assertFalse(args.check_updates)


if __name__ == "__main__":
    unittest.main(verbosity=2)
