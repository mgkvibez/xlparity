"""Compute the compatibility scoreboard, and fail when the committed page is stale.

The metric is pre-registered in `docs/scoreboard.md`: per-formula match, cohort scores
reported separately, and no headline number allowed to hide a cohort. This tool is the
only thing that writes the generated section, so the page cannot drift from the corpus.

Expectations currently come from Microsoft's documented semantics, tagged `docs` per
entry. When a real Excel capture is frozen for an entry, its tag becomes `excel`, and the
scoreboard prints the split, so a number derived from documentation is never presented as
an Excel-captured result.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Tuple

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from xlparity import ExcelError, evaluate, parse  # noqa: E402

CORPUS = ROOT / "corpus" / "scoreboard.json"
PAGE = ROOT / "docs" / "scoreboard.md"
BEGIN = "<!-- BEGIN GENERATED: scoreboard -->"
END = "<!-- END GENERATED: scoreboard -->"
TOLERANCE = 1e-12

COHORT_TITLES = {
    "core": "1. Core arithmetic and precedence",
    "text": "2. Text, comparison and coercion",
    "common": "3. Common functions",
    "financial": "4. Obscure financial",
    "arrays": "5. Dynamic arrays",
}


def as_display(value: Any) -> str:
    """Render a result the way the corpus stores it, so a match prints identically on both sides."""
    if isinstance(value, ExcelError):
        return value.code
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if value is None:
        return "(blank)"
    if isinstance(value, float):
        return repr(round(value, 12))
    return str(value)


def matches(actual: Any, expected: Any) -> bool:
    """The comparison rules from the pre-registered metric, and nothing beyond them."""
    expected_is_error = isinstance(expected, str) and expected.startswith("#")
    if isinstance(actual, ExcelError) or expected_is_error:
        left = actual.code if isinstance(actual, ExcelError) else as_display(actual)
        right = expected.code if isinstance(expected, ExcelError) else expected
        return left == right
    if isinstance(actual, bool) != isinstance(expected, bool):
        return False
    if isinstance(actual, bool):
        return actual == expected
    if isinstance(actual, (int, float)) and isinstance(expected, (int, float)):
        return abs(actual - expected) <= TOLERANCE * max(1.0, abs(float(expected)))
    if actual is None and expected == "":
        return True
    if expected is None and actual == "":
        return True
    if isinstance(actual, str) and isinstance(expected, str):
        return actual == expected
    return False


def run_entry(entry: Dict[str, Any]) -> Tuple[bool, Any]:
    try:
        actual = evaluate(parse(entry["formula"]), entry.get("inputs") or None)
    except Exception as error:  # noqa: BLE001 - a crash is a finding to report, not a reason to stop
        return False, error.__class__.__name__
    return matches(actual, entry["expected"]), actual


def build() -> Tuple[str, int, int]:
    corpus = json.loads(CORPUS.read_text(encoding="utf-8"))["entries"]
    rows: List[Dict[str, Any]] = []
    for entry in corpus:
        ok, actual = run_entry(entry)
        rows.append({**entry, "ok": ok, "actual": actual})

    totals: Dict[str, List[Dict[str, Any]]] = {name: [] for name in COHORT_TITLES}
    for row in rows:
        totals.setdefault(row["cohort"], []).append(row)

    lines = [BEGIN, ""]
    source_counts: Dict[str, int] = {}
    for row in rows:
        source_counts[row.get("source", "docs")] = source_counts.get(row.get("source", "docs"), 0) + 1
    provenance = ", ".join(f"{count} {name}" for name, count in sorted(source_counts.items()))
    lines.append(f"**Expectation provenance:** {provenance}. `docs` means Microsoft's documented")
    lines.append("semantics; `excel` means captured from Excel and frozen; `derived` means computed here")
    lines.append("from the documented closed form with exact rational arithmetic, which is a second")
    lines.append("implementation of the definition rather than a capture of Excel's own result. No `docs`")
    lines.append("or `derived` entry claims to be an `excel` capture.")
    lines.append("")
    lines.append("| cohort | passing | total | score |")
    lines.append("| --- | --- | --- | --- |")
    for name, title in COHORT_TITLES.items():
        bucket = totals.get(name, [])
        passing = sum(1 for row in bucket if row["ok"])
        score = "n/a" if not bucket else f"{100.0 * passing / len(bucket):.1f}%"
        lines.append(f"| {title} | {passing} | {len(bucket)} | {score} |")
    passing_all = sum(1 for row in rows if row["ok"])
    lines.append("")
    share = 100.0 * passing_all / len(rows)
    lines.append(f"**All cohorts:** {passing_all} of {len(rows)} formulas match ({share:.1f}%).")
    lines.append("")

    failures = [row for row in rows if not row["ok"]]
    if failures:
        lines.append("### Failures, listed by name")
        lines.append("")
        lines.append("| function | formula | expected | actual |")
        lines.append("| --- | --- | --- | --- |")
        for row in failures:
            lines.append(
                f"| `{row['function']}` | `{row['formula']}` | `{as_display(row['expected'])}` | "
                f"`{as_display(row['actual'])}` |"
            )
        lines.append("")
    else:
        lines.append("Every corpus formula matches. The corpus is grown, never trimmed, so this is only")
        lines.append("true until the next difficult formula is added.")
        lines.append("")
    lines.append(END)
    return "\n".join(lines) + "\n", passing_all, len(rows)


def rewrite(page_text: str, generated: str) -> str:
    """Splice the generated block in, normalising blank lines so repeated writes are idempotent."""
    if BEGIN not in page_text or END not in page_text:
        raise SystemExit(f"{PAGE} needs the {BEGIN} and {END} markers")
    head = page_text.split(BEGIN)[0].rstrip("\n")
    tail = page_text.split(END)[1].lstrip("\n")
    return f"{head}\n\n{generated.rstrip(chr(10))}\n\n{tail}"


def main(argv: List[str]) -> int:
    generated, passing, total = build()
    page = PAGE.read_text(encoding="utf-8")
    expected_page = rewrite(page, generated)
    if "--write" in argv:
        PAGE.write_text(expected_page, encoding="utf-8")
        print(f"scoreboard written: {passing}/{total} matching")
        return 0
    if expected_page != page:
        print("docs/scoreboard.md is stale: run `python tools/scoreboard.py --write`")
        return 1
    print(f"scoreboard current: {passing}/{total} matching")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
