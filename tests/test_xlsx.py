"""The reader, tested against workbooks assembled part by part.

There is no Excel on this machine, so the fixtures are written as the parts Excel writes: a zip with
workbook.xml,
its relationships, sharedStrings.xml, styles.xml and sheet XML. That is spec-shaped rather than Excel-written,
which
is a real limit and is stated in the README. What it does check is what is easy to get wrong and hard to
notice: which
cells are dates, which strings are rich text, and what a formula cell with no cached value means.
"""

from __future__ import annotations

import io
import zipfile

import pytest

from xlparity.xlsx import XlsxError, read_workbook

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


def build_xlsx(
    sheet_xml: str,
    *,
    shared_strings: str | None = None,
    styles: str | None = None,
    date1904: bool = False,
    defined_names: str = "",
    sheet_name: str = "Sheet1",
) -> bytes:
    """A minimal but valid .xlsx, with only the parts the reader needs."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "xl/workbook.xml",
            f'<workbook xmlns="{MAIN}" xmlns:r="{REL}">'
            f'<workbookPr date1904="{"1" if date1904 else "0"}"/>'
            f'<sheets><sheet name="{sheet_name}" sheetId="1" r:id="rId1"/></sheets>'
            f"{defined_names}</workbook>",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            f'<Relationships xmlns="{PKG}">'
            '<Relationship Id="rId1"'
            ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
            ' Target="worksheets/sheet1.xml"/></Relationships>',
        )
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            f'<worksheet xmlns="{MAIN}"><sheetData>{sheet_xml}</sheetData></worksheet>',
        )
        if shared_strings is not None:
            archive.writestr(
                "xl/sharedStrings.xml",
                f'<sst xmlns="{MAIN}">{shared_strings}</sst>',
            )
        if styles is not None:
            archive.writestr("xl/styles.xml", f'<styleSheet xmlns="{MAIN}">{styles}</styleSheet>')
    return buffer.getvalue()


def style_sheet(number_formats: dict[int, str], cells: str) -> str:
    """A styles part whose cellXfs map style index n to number_formats[n], plus the cells that use it."""
    custom = "".join(
        f'<numFmt numFmtId="{identifier}" formatCode="{code}"/>'
        for identifier, code in number_formats.items()
    )
    formats = "".join(
        f'<xf numFmtId="{identifier}"/>' for identifier in sorted(number_formats) or [0]
    )
    return f'<numFmts count="{len(number_formats)}">{custom}</numFmts><cellXfs count="1">{formats}</cellXfs>'


def test_reads_a_number_a_shared_string_and_a_formula_with_its_cached_value() -> None:
    sheet = (
        '<row r="1"><c r="A1"><v>42</v></c>'
        '<c r="B1" t="s"><v>0</v></c>'
        '<c r="C1"><f>A1*2</f><v>84</v></c></row>'
    )
    workbook = read_workbook(build_xlsx(sheet, shared_strings="<si><t>hello</t></si>"))
    cells = workbook.sheet("Sheet1")
    assert cells.value("A1") == 42
    assert cells.value("B1") == "hello"
    assert cells.value("C1") == 84
    assert cells.cells["C1"].formula == "A1*2"


def test_joins_rich_text_runs_into_the_string_the_cell_displays() -> None:
    sheet = '<row r="1"><c r="A1" t="s"><v>0</v></c></row>'
    shared = '<si><r><t>bold</t></r><r><t xml:space="preserve"> and </t></r><r><t>plain</t></r></si>'
    workbook = read_workbook(build_xlsx(sheet, shared_strings=shared))
    assert workbook.sheet("Sheet1").value("A1") == "bold and plain"


def test_reports_a_formula_with_no_cached_value_as_uncached_rather_than_empty() -> None:
    sheet = '<row r="1"><c r="A1"><f>1+1</f></c></row>'
    cells = read_workbook(build_xlsx(sheet)).sheet("Sheet1")
    assert cells.cells["A1"].kind == "uncached"
    with pytest.raises(XlsxError, match="no cached value"):
        cells.value("A1")


def test_a_cell_that_was_never_written_reads_as_empty_not_as_an_error() -> None:
    cells = read_workbook(build_xlsx('<row r="1"><c r="A1"><v>1</v></c></row>')).sheet("Sheet1")
    assert cells.value("Z99") is None


def test_reads_booleans_errors_and_inline_strings_by_type() -> None:
    sheet = (
        '<row r="1"><c r="A1" t="b"><v>1</v></c>'
        '<c r="B1" t="b"><v>0</v></c>'
        '<c r="C1" t="e"><v>#DIV/0!</v></c>'
        '<c r="D1" t="inlineStr"><is><t>inline</t></is></c></row>'
    )
    cells = read_workbook(build_xlsx(sheet)).sheet("Sheet1")
    assert cells.value("A1") is True
    assert cells.value("B1") is False
    assert cells.value("C1") == "#DIV/0!"
    assert cells.value("D1") == "inline"


def test_converts_a_date_styled_serial_using_the_1900_epoch() -> None:
    styles = style_sheet({164: "yyyy-mm-dd"}, "")
    sheet = '<row r="1"><c r="A1" s="0"><v>45000</v></c><c r="B1"><v>45000</v></c></row>'
    cells = read_workbook(build_xlsx(sheet, styles=styles)).sheet("Sheet1")
    assert cells.cells["A1"].kind == "date"
    assert str(cells.value("A1").date()) == "2023-03-15"
    # the same serial without the style is a number: the difference between the two cells is formatting alone
    assert cells.value("B1") == 45000


def test_converts_a_date_serial_using_the_1904_epoch_when_the_writer_said_so() -> None:
    styles = style_sheet({14: "mm-dd-yy"}, "")
    sheet = '<row r="1"><c r="A1" s="0"><v>45000</v></c></row>'
    cells = read_workbook(build_xlsx(sheet, styles=styles, date1904=True)).sheet("Sheet1")
    assert str(cells.value("A1").date()) == "2027-03-16"


def test_treats_a_custom_format_with_date_letters_as_a_date_and_time_zones_as_numbers() -> None:
    styles = style_sheet({165: "[$-409]d-mmm-yy;@", 166: "0.00"}, "")
    sheet = '<row r="1"><c r="A1" s="0"><v>45000</v></c><c r="B1" s="1"><v>45000</v></c></row>'
    cells = read_workbook(build_xlsx(sheet, styles=styles)).sheet("Sheet1")
    assert cells.cells["A1"].kind == "date"
    assert cells.cells["B1"].kind == "number"


def test_reads_defined_names_and_sheet_names_in_workbook_order() -> None:
    names = '<definedNames><definedName name="Rate">Sheet1!$B$1</definedName></definedNames>'
    workbook = read_workbook(build_xlsx('<row r="1"/>', defined_names=names, sheet_name="Inputs"))
    assert workbook.defined_names == {"Rate": "Sheet1!$B$1"}
    assert workbook.sheets[0].name == "Inputs"
    with pytest.raises(XlsxError, match="no sheet named"):
        workbook.sheet("Missing")


def test_refuses_a_shared_string_index_that_is_not_in_the_table() -> None:
    sheet = '<row r="1"><c r="A1" t="s"><v>7</v></c></row>'
    with pytest.raises(XlsxError, match="shared string 7"):
        read_workbook(build_xlsx(sheet, shared_strings="<si><t>only one</t></si>"))


def test_refuses_a_cell_type_it_does_not_understand() -> None:
    sheet = '<row r="1"><c r="A1" t="d"><v>2023-03-15</v></c></row>'
    with pytest.raises(XlsxError, match="cell type 'd'"):
        read_workbook(build_xlsx(sheet))


def test_reports_a_missing_part_by_the_sheet_that_named_it() -> None:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr(
            "xl/workbook.xml",
            f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets>'
            '<sheet name="Gone" sheetId="1" r:id="rId1"/></sheets></workbook>',
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            f'<Relationships xmlns="{PKG}">'
            '<Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
        )
    with pytest.raises(XlsxError, match="not in the archive"):
        read_workbook(buffer.getvalue())
