"""Turning a workbook Excel wrote into corpus entries.
Turning a workbook Excel wrote into corpus entries.
This is the bridge the scoreboard's `excel` provenance needs. A cell in a real workbook carries a formula and
the value Excel cached for it, which is exactly the pair a corpus entry is: `formula`, `expected`, and the
inputs the formula references. So the importer reads the sheet, extracts each formula cell, reconstructs its
inputs from the neighbouring cached values, and emits entries tagged `excel`.
Deliberately, it does not compute anything. If the engine disagrees with Excel, that is a finding for the
scoreboard to print, not something the importer smooths over.
Two rules it will not bend:
- A cell whose own value Excel did not cache is skipped, with the reason recorded. An entry without an
  expected value is a claim with nothing behind it. - A cell that depends on another uncached formula, or on
  another sheet, is skipped with that reason. Replaying it with an empty input would test the engine against a
  different workbook from the one Excel calculated, and any verdict would be about the difference rather than
  the formula.
Usage:
python tools/import_workbook.py book.xlsx --cohort core --out corpus/imported.json
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Tuple

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from xlparity.xlsx import Cell, Sheet, Workbook, read_workbook  # noqa: E402

# A1, $A$1, and the sheet-qualified forms; the range form is matched first so its two ends are not read
# separately
_SHEET = r"(?:(?:'(?P<sheet>[^']+)'|(?P<bare_sheet>[A-Za-z_][A-Za-z0-9_.]*))!)?"
_CELL = r"\$?[A-Z]{1,3}\$?\d+"
RANGE = re.compile(_SHEET + f"(?P<start>{_CELL}):(?P<end>{_CELL})")
SINGLE = re.compile(_SHEET + f"(?P<reference>{_CELL})")
# function names look like references only when followed by "(", so "LOG10(" must not be read as cell LOG10
NOT_A_CALL = re.compile(r"[A-Za-z0-9_.]+\s*\(")


class Unsupported(Exception):
    """A cell this importer will not turn into an entry, with the reason it gives."""


def _split(absolute: str) -> Tuple[int, int]:
    letters = "".join(character for character in absolute if character.isalpha())
    digits = "".join(character for character in absolute if character.isdigit())
    column = 0
    for character in letters:
        column = column * 26 + (ord(character) - ord("A") + 1)
    return int(digits), column


def _render(row: int, column: int) -> str:
    letters = ""
    while column > 0:
        column, remainder = divmod(column - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return f"{letters}{row}"


def references_in(formula: str) -> List[str]:
    """Every cell reference a formula names, with ranges expanded and absolute markers dropped."""
    stripped = NOT_A_CALL.sub("", formula)
    references: List[str] = []
    for match in RANGE.finditer(stripped):
        sheet = match.group("sheet") or match.group("bare_sheet")
        if sheet is not None:
            raise Unsupported(f"references sheet {sheet!r}")
        start_row, start_column = _split(match.group("start").replace("$", ""))
        end_row, end_column = _split(match.group("end").replace("$", ""))
        for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
            for column in range(min(start_column, end_column), max(start_column, end_column) + 1):
                references.append(_render(row, column))
    stripped = RANGE.sub("", stripped)
    for match in SINGLE.finditer(stripped):
        sheet = match.group("sheet") or match.group("bare_sheet")
        if sheet is not None:
            raise Unsupported(f"references sheet {sheet!r}")
        reference = match.group("reference").replace("$", "")
        if reference not in references:
            references.append(reference)
    return references


def _input_of(cell: Cell) -> Any:
    if cell.kind == "uncached":
        raise Unsupported(f"depends on {cell.reference}, which has no cached value")
    if cell.kind == "error":
        raise Unsupported(f"depends on {cell.reference}, which is {cell.value!r} in Excel")
    if cell.kind == "date":
        # Excel stores a date as a serial; the engine's inputs are values, so the serial is what it must see
        raise Unsupported(f"depends on {cell.reference}, whose value is a formatted date")
    return cell.value


def entries_from_sheet(sheet: Sheet, cohort: str, workbook_name: str) -> Tuple[List[Dict[str, Any]],
List[Dict[str, str]]]:
    """Entries and skips for one sheet, in cell order so an import is reproducible."""
    entries: List[Dict[str, Any]] = []
    skips: List[Dict[str, str]] = []
    for cell in sorted(sheet, key=lambda item: (item.row, item.column)):
        if cell.formula is None:
            continue
        where = {"workbook": workbook_name, "sheet": sheet.name, "cell": cell.reference}
        try:
            if cell.formula.startswith("{"):
                raise Unsupported("is an array formula, whose spill range is not modelled here")
            if cell.value is None or cell.kind == "uncached":
                raise Unsupported("has no cached value, so there is nothing to compare against")
            if cell.kind == "error":
                expected = str(cell.value)
            elif cell.kind == "date":
                raise Unsupported("is a formatted date; the corpus compares numbers")
            else:
                expected = cell.value
            inputs: Dict[str, Any] = {}
            for reference in references_in(cell.formula):
                if reference == cell.reference:
                    # a formula that names itself is a circular reference Excel would have refused
                    raise Unsupported("refers to itself")
                neighbour = sheet.cells.get(reference)
                if neighbour is None:
                    # an unwritten cell is empty in Excel too, so leaving it out of the inputs is faithful
                    continue
                inputs[reference] = _input_of(neighbour)
            entries.append(
                {
                    "function": _function_of(cell.formula),
                    "cohort": cohort,
                    "formula": cell.formula,
                    "inputs": inputs,
                    "expected": expected,
                    "source": "excel",
                    "provenance": where,
                }
            )
        except Unsupported as reason:
            skips.append({**where, "formula": cell.formula, "reason": str(reason)})
    return entries, skips


def _function_of(formula: str) -> str:
    match = re.match(r"\s*([A-Za-z_][A-Za-z0-9_.]*)\s*\(", formula)
    return match.group(1).upper() if match else "EXPRESSION"


def entries_from_workbook(workbook: Workbook, cohort: str, workbook_name: str = "") -> Tuple[List[Dict[str,
Any]], List[Dict[str, str]]]:
    entries: List[Dict[str, Any]] = []
    skips: List[Dict[str, str]] = []
    for sheet in workbook.sheets:
        sheet_entries, sheet_skips = entries_from_sheet(sheet, cohort, workbook_name)
        entries.extend(sheet_entries)
        skips.extend(sheet_skips)
    return entries, skips


def main(argv: List[str]) -> int:
    parser = argparse.ArgumentParser(description="Import formula cells from a workbook as corpus entries.")
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--cohort", default="core", choices=["core", "text", "common", "financial", "arrays"])
    parser.add_argument("--out", type=Path, default=None)
    arguments = parser.parse_args(argv)

    workbook = read_workbook(str(arguments.workbook))
    entries, skips = entries_from_workbook(workbook, arguments.cohort, arguments.workbook.name)
    payload = {"entries": entries, "skipped": skips}
    if arguments.out is None:
        json.dump(payload, sys.stdout, indent=2)
        sys.stdout.write("\n")
    else:
        arguments.out.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    print(f"{len(entries)} entries, {len(skips)} skipped", file=sys.stderr)
    for skip in skips:
        print(f"  skipped {skip['sheet']}!{skip['cell']}: {skip['reason']}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
