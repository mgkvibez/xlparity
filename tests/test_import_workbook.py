"""The importer, from a workbook to entries the scoreboard can run.
The importer, from a workbook to entries the scoreboard can run.
The point of these tests is the join: an entry made here must be replayable, meaning its inputs are the values
Excel had, and its expected value is the one Excel cached. A test that only checked the JSON shape would miss
the failure that matters, which is an entry that replays a slightly different workbook and therefore scores
the engine against the wrong question.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

from xlparity import evaluate, parse
from xlparity.xlsx import read_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))

from import_workbook import entries_from_workbook, references_in  # noqa: E402
from test_xlsx import build_xlsx  # noqa: E402


def import_sheet(sheet_xml: str, cohort: str = "core", **kwargs: Any):
    workbook = read_workbook(build_xlsx(sheet_xml, **kwargs))
    return entries_from_workbook(workbook, cohort, "book.xlsx")


def test_expands_ranges_and_drops_absolute_markers() -> None:
    assert references_in("SUM($A$1:B2)+C3") == ["A1", "B1", "A2", "B2", "C3"]
    assert references_in("A1*2") == ["A1"]


def test_does_not_mistake_a_function_name_for_a_cell() -> None:
    assert references_in("LOG10(A1)+ROUNDUP(B2,0)") == ["A1", "B2"]


def test_an_imported_entry_replays_to_the_value_excel_cached() -> None:
    sheet = (
        '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c><c r="C1"><f>A1*B1</f><v>6</v></c></row>'
    )
    entries, skips = import_sheet(sheet)
    assert skips == []
    entry = entries[0]
    assert entry["formula"] == "A1*B1"
    assert entry["inputs"] == {"A1": 2, "B1": 3}
    assert entry["expected"] == 6
    assert entry["source"] == "excel"
    assert entry["provenance"] == {"workbook": "book.xlsx", "sheet": "Sheet1", "cell": "C1"}
    assert evaluate(parse(entry["formula"]), entry["inputs"]) == entry["expected"]


def test_a_range_formula_gets_every_cell_as_an_input() -> None:
    sheet = (
        '<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>'
        '<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f>SUM(A1:A3)</f><v>6</v></c></row>'
    )
    entries, skips = import_sheet(sheet, cohort="common")
    assert skips == []
    entry = entries[0]
    assert entry["inputs"] == {"A1": 1, "A2": 2, "A3": 3}
    assert evaluate(parse(entry["formula"]), entry["inputs"]) == 6


def test_names_the_function_the_formula_calls() -> None:
    entries, _ = import_sheet('<row r="1"><c r="A1"><v>4</v></c><c r="B1"><f>SQRT(A1)</f><v>2</v></c></row>')
    assert entries[0]["function"] == "SQRT"


def test_skips_a_formula_whose_own_value_excel_never_cached() -> None:
    entries, skips = import_sheet('<row r="1"><c r="A1"><f>1+1</f></c></row>')
    assert entries == []
    assert skips[0]["reason"].startswith("has no cached value")


def test_skips_a_formula_that_depends_on_an_uncached_cell() -> None:
    sheet = '<row r="1"><c r="A1"><f>1/0</f></c><c r="B1"><f>A1+1</f><v>2</v></c></row>'
    entries, skips = import_sheet(sheet)
    assert entries == []
    assert any("A1" in skip["reason"] for skip in skips)


def test_skips_a_cross_sheet_reference_rather_than_replaying_the_wrong_workbook() -> None:
    entries, skips = import_sheet('<row r="1"><c r="A1"><f>Sheet2!A1*2</f><v>4</v></c></row>')
    assert entries == []
    assert skips[0]["reason"] == "references sheet 'Sheet2'"


def test_keeps_an_error_result_as_its_code_so_the_scoreboard_compares_codes() -> None:
    sheet = '<row r="1"><c r="A1"><v>0</v></c><c r="B1" t="e"><f>1/A1</f><v>#DIV/0!</v></c></row>'
    entries, skips = import_sheet(sheet)
    assert skips == []
    assert entries[0]["expected"] == "#DIV/0!"


def test_skips_a_cell_that_refers_to_itself() -> None:
    # Excel would refuse this, but a hand-edited file can contain it; no circular entry may come out
    entries, skips = import_sheet('<row r="1"><c r="A1"><f>A1+1</f><v>1</v></c></row>')
    assert entries == []
    assert skips[0]["reason"] == "refers to itself"


def test_entries_come_out_in_cell_order_so_an_import_is_reproducible() -> None:
    sheet = (
        '<row r="1"><c r="A1"><v>1</v></c><c r="C1"><f>A1+1</f><v>2</v></c></row>'
        '<row r="2"><c r="B2"><v>5</v></c><c r="D2"><f>B2*2</f><v>10</v></c></row>'
    )
    entries, _ = import_sheet(sheet)
    assert [entry["provenance"]["cell"] for entry in entries] == ["C1", "D2"]


def test_inputs_carry_the_cached_values_as_excel_stored_them() -> None:
    sheet = (
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><f>LEN(A1)</f><v>5</v></c></row>'
    )
    entries, skips = import_sheet(sheet, shared_strings="<si><t>hello</t></si>")
    assert skips == []
    assert entries[0]["inputs"] == {"A1": "hello"}
    assert evaluate(parse(entries[0]["formula"]), entries[0]["inputs"]) == 5
