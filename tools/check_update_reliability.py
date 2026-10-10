"""Run the update check several times to see whether failures are intermittent.

A single failed check is usually a network blip and is reported as such, but
that is worth distinguishing from a systematic problem before trusting the
updater. Run from the project root.

    python tools/check_update_reliability.py --attempts 5
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from updater import check_for_update  # noqa: E402


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--attempts", type=int, default=5)
    parser.add_argument("--delay", type=float, default=0.5)
    args = parser.parse_args(argv)

    failures = 0
    for index in range(args.attempts):
        result = check_for_update()
        if result.update_available:
            state = "update offered"
        elif result.ok:
            state = "up to date"
        else:
            state = "CHECK FAILED"
            failures += 1
        print(
            f"  attempt {index + 1}: status={result.status:16} "
            f"remote={result.remote_version}  {state}"
        )
        if result.error:
            print(f"              error: {result.error}")
        if index < args.attempts - 1:
            time.sleep(args.delay)

    print()
    if failures == 0:
        print(f"All {args.attempts} checks succeeded.")
        return 0
    if failures == args.attempts:
        print(f"All {args.attempts} checks failed - this is systematic, not a blip.")
        return 1
    print(
        f"{failures} of {args.attempts} checks failed - intermittent network, "
        "and each failure was reported rather than crashing."
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
