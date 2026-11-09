"""Fail when the README's "Tests: N" line disagrees with the suite.

Rule from the handover: a claim not backed by a test is unverified, and a count that
drifts from the suite is the cheapest way for a README to become dishonest. This
collects the suite, reads the real number, and refuses to pass on a mismatch.
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
COUNT = re.compile(r"^\*\*Tests:\*\* (\d+)", re.MULTILINE)


def collected() -> int:
    result = subprocess.run(
        [sys.executable, "-m", "pytest", "--collect-only", "-q"],
        cwd=ROOT, capture_output=True, text=True,
    )
    match = re.search(r"(\d+) tests? collected", result.stdout)
    if match is None:
        print("could not read a test count from pytest:\n" + result.stdout + result.stderr)
        raise SystemExit(1)
    return int(match.group(1))


def main() -> int:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    match = COUNT.search(readme)
    if match is None:
        print('no "**Tests:** N" line in README.md')
        return 1
    claimed, actual = int(match.group(1)), collected()
    if claimed != actual:
        print(f"README claims {claimed} tests, suite has {actual}. Update the README (or the tests).")
        return 1
    print(f"test count matches README: {actual}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
