"""One-off: report how each hardcoded colour is used, to plan the theming work.

Run from the project root.
"""

from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path

SOURCE = Path(__file__).resolve().parent.parent / "waybill_sorter.py"

ATTRIBUTES = (
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


def main() -> int:
    src = SOURCE.read_text(encoding="utf-8")

    usage: dict[str, set[str]] = defaultdict(set)
    for attribute in ATTRIBUTES:
        for match in re.finditer(
            attribute + r"\s*=\s*[\"']?(#[0-9A-Fa-f]{6})", src
        ):
            usage[match.group(1).upper()].add(attribute)

    print("colour usage by attribute:")
    for colour in sorted(usage, key=lambda key: (-len(usage[key]), key)):
        print(f"  {colour}: {', '.join(sorted(usage[colour]))}")

    print()
    print("non-hex colour keywords:")
    found = re.findall(
        r"(?:bg|fg|background|foreground)\s*=\s*[\"']([A-Za-z]+)[\"']", src
    )
    keywords = sorted(set(found))
    print(f"  {keywords if keywords else 'none'}")

    print()
    print("ttk style configuration:")
    for line_number, line in enumerate(src.splitlines(), start=1):
        if "ttk.Style" in line or "style.configure" in line or "theme_use" in line:
            print(f"  L{line_number}: {line.strip()}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
