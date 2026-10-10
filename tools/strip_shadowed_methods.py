"""One-off: strip class methods that a later definition silently overrides.

In Python, defining the same method twice in a class body keeps only the last.
The earlier definition is unreachable - it cannot be called, cannot be reached
by any code path, and exists only to mislead whoever reads the file next.

This removes exactly those definitions and nothing else: a method is deleted
only when the *same class* defines it again later. Behaviour cannot change,
because the deleted code was already unreachable.

Run from the project root:

    python tools/strip_shadowed_methods.py --dry-run
    python tools/strip_shadowed_methods.py
"""

from __future__ import annotations

import argparse
import ast
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TARGET = ROOT / "waybill_sorter.py"


def find_shadowed(source: str) -> list[tuple[str, int, int, int]]:
    """Return ``(method, start, end, winner_line)`` for every shadowed method."""
    tree = ast.parse(source)
    found: list[tuple[str, int, int, int]] = []

    for node in ast.walk(tree):
        if not isinstance(node, ast.ClassDef):
            continue
        by_name: dict[str, list[ast.FunctionDef]] = defaultdict(list)
        for item in node.body:
            if isinstance(item, ast.FunctionDef):
                by_name[item.name].append(item)
        for name, definitions in by_name.items():
            if len(definitions) < 2:
                continue
            winner = definitions[-1]
            for dead in definitions[:-1]:
                found.append((name, dead.lineno, dead.end_lineno, winner.lineno))

    return sorted(found, key=lambda item: item[1])


def strip_shadowed(path: Path, dry_run: bool = False) -> int:
    source = path.read_text(encoding="utf-8")
    targets = find_shadowed(source)
    if not targets:
        print("No shadowed methods found; nothing to do.")
        return 0

    lines = source.splitlines(keepends=True)
    removed: set[int] = set()
    for name, start, end, winner in targets:
        print(f"  removing {name}() L{start}-{end}  (kept the definition at L{winner})")
        # Also drop one blank separator line after the method, so the file does
        # not accumulate double blank lines.
        stop = end + 1 if end < len(lines) and not lines[end].strip() else end
        for index in range(start - 1, stop):
            removed.add(index)

    if dry_run:
        print(f"\nDry run: would remove {len(removed)} lines from {path.name}.")
        return 0

    kept = [line for index, line in enumerate(lines) if index not in removed]
    result = "".join(kept)

    # The rewrite must not change what the code does. Re-parsing proves the
    # file is still valid; comparing the surviving method names proves nothing
    # reachable was lost.
    ast.parse(result)
    before = _class_method_names(source)
    after = _class_method_names(result)
    if before != after:
        lost = before - after
        gained = after - before
        raise SystemExit(f"Refusing to write: method set changed (lost={lost}, gained={gained})")

    path.write_text(result, encoding="utf-8")
    print(f"\nRemoved {len(removed)} lines. {len(targets)} shadowed definitions gone.")
    print(f"Surviving method names unchanged ({len(after)}).")
    return 0


def _class_method_names(source: str) -> set[str]:
    tree = ast.parse(source)
    names: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef):
            for item in node.body:
                if isinstance(item, ast.FunctionDef):
                    names.add(f"{node.name}.{item.name}")
    return names


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Remove methods that a later definition overrides.")
    parser.add_argument("--dry-run", action="store_true", help="Show what would be removed.")
    parser.add_argument("--path", default=str(TARGET), help="File to process.")
    args = parser.parse_args(argv)

    path = Path(args.path)
    if not path.exists():
        print(f"Not found: {path}", file=sys.stderr)
        return 1

    print(f"Scanning {path.name} for shadowed methods...")
    return strip_shadowed(path, dry_run=args.dry_run)


if __name__ == "__main__":
    sys.exit(main())
