"""One-off: convert action-column buttons to the width-sharing stack helper.

Buttons were created and then packed individually with ``fill=tk.X``. Tk's
``width`` option is only a *minimum*, so a button whose label was shorter than
that minimum came out narrower than the rest of the column - the ragged blue
boxes. ``_stack_button`` gives each button its own weighted grid row with
``sticky="ew"``, so every button adopts the widest button's width.

A regex was tried first and proved brittle: the calls wrap across lines, some
``.pack`` calls sit on the same line and some on the next, and the arguments
contain nested calls with commas inside string literals. This walks the source
instead, matching brackets to find where each call really ends.

Run from the project root:

    python tools/apply_stack_buttons.py --dry-run
    python tools/apply_stack_buttons.py
"""

from __future__ import annotations

import argparse
import ast
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TARGET = ROOT / "waybill_sorter.py"

#: Matches the start of a factory call, with an optional attribute assignment
#: immediately before it, anchored at the beginning of a line.
CALL_START_RE = re.compile(
    r"^(?P<indent>[ ]+)"
    r"(?:(?P<assign>[A-Za-z_][A-Za-z0-9_.]*)\s*=\s*)?"
    r"self\._(?P<kind>primary|secondary)_button\(",
    re.MULTILINE,
)


def find_call_end(source: str, open_paren: int) -> int:
    """Index just past the ``)`` that closes the call starting at ``open_paren``."""
    depth = 0
    quote: str | None = None
    index = open_paren
    while index < len(source):
        char = source[index]
        if quote:
            if char == quote:
                quote = None
        elif char in "\"'":
            quote = char
        elif char in "([{":
            depth += 1
        elif char in ")]}":
            depth -= 1
            if depth == 0:
                return index + 1
        index += 1
    raise ValueError("unbalanced brackets")


def split_args(text: str) -> list[str]:
    """Split on top-level commas, respecting quotes and nested brackets."""
    parts: list[str] = []
    depth = 0
    quote: str | None = None
    current = ""

    for char in text:
        if quote:
            current += char
            if char == quote:
                quote = None
            continue
        if char in "\"'":
            quote = char
            current += char
            continue
        if char in "([{":
            depth += 1
        elif char in ")]}":
            depth -= 1
        if char == "," and depth == 0:
            parts.append(current.strip())
            current = ""
        else:
            current += char

    if current.strip():
        parts.append(current.strip())
    return [part for part in parts if part]


def rewrite(source: str) -> tuple[str, int, list[str]]:
    """Return ``(new_source, buttons_rewritten, problems)``."""
    count = 0
    problems: list[str] = []
    out: list[str] = []
    cursor = 0

    while True:
        match = CALL_START_RE.search(source, cursor)
        if not match:
            out.append(source[cursor:])
            break

        indent = match.group("indent")
        assign = match.group("assign")
        kind = match.group("kind")
        open_paren = match.end() - 1

        try:
            call_end = find_call_end(source, open_paren)
        except ValueError:
            problems.append(f"unbalanced call near offset {match.start()}")
            out.append(source[cursor:match.end()])
            cursor = match.end()
            continue

        # Is this call immediately followed by .pack(fill=tk.X...)?
        #
        # The pack is either on the same line as the closing bracket or on the
        # next line, indented the same amount, naming the attribute again:
        #
        #     self.sort_button = self._primary_button(
        #         action_stack, self.tr("sort"), self.start_sort, label_key="sort"
        #     )
        #     self.sort_button.pack(fill=tk.X, pady=(0, 8))
        #
        # Both shapes are accepted here.
        pack_pattern = r"\s*\.pack\(fill=tk\.X(?P<spaced>,\s*pady=\(0,\s*8\))?\)"
        if assign:
            pack_pattern = (
                r"(?:\s*\n" + indent + r"(?:" + re.escape(assign) + r"\s*)?)?"
                + pack_pattern
            )
        pack_match = re.match(pack_pattern, source[call_end:call_end + 400])
        if not pack_match:
            # Not an action-column button (the factory result is used
            # differently); leave it alone.
            out.append(source[cursor:call_end])
            cursor = call_end
            continue

        body = source[open_paren + 1:call_end - 1]
        pieces = split_args(body)
        if len(pieces) < 3:
            problems.append(f"too few arguments near offset {match.start()}: {body[:60]}")
            out.append(source[cursor:call_end])
            cursor = call_end
            continue

        target, text, command = pieces[0], pieces[1], pieces[2]
        extra = pieces[3:]

        count += 1
        kwargs: list[str] = []
        for item in extra:
            if item.startswith("label_key=") or item == "compact=True":
                kwargs.append(item)
        if kind == "secondary":
            kwargs.append("primary=False")
        if pack_match.group("spaced") is None:
            kwargs.append("pady=(0, 0)")

        args = [target, str(count), text, command] + kwargs
        prefix = f"{assign} = " if assign else ""
        single = f"{indent}{prefix}self._stack_button({', '.join(args)})"

        out.append(source[cursor:match.start()])
        if len(single) <= 100:
            out.append(single)
        else:
            body_indent = indent + "    "
            rendered = ",\n".join(f"{body_indent}{part}" for part in args)
            out.append(f"{indent}{prefix}self._stack_button(\n{rendered},\n{indent})")

        cursor = call_end + pack_match.end()

    return "".join(out), count, problems


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args(argv)

    source = TARGET.read_text(encoding="utf-8")
    ast.parse(source)

    result, count, problems = rewrite(source)
    if problems:
        print("Problems found; not writing:")
        for problem in problems:
            print("  ", problem)
        return 1
    if count == 0:
        print("No matching button patterns found - nothing to do.")
        return 1

    ast.parse(result)

    if args.dry_run:
        print(f"Dry run: would rewrite {count} buttons in {TARGET.name}.")
        for line in result.splitlines():
            if "_stack_button(" in line:
                print("   ", line.strip()[:110])
        return 0

    TARGET.write_text(result, encoding="utf-8")
    print(f"Rewrote {count} buttons to use _stack_button in {TARGET.name}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
