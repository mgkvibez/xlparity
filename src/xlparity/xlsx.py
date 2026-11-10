"""Reading .xlsx without Excel.
xlparity's whole claim is that it agrees with Excel. That claim needs Excel's own answers to compare against,
and the only honest place to get them is a workbook Excel wrote: a file carries both the formula text and the
value Excel cached for it. So the reader's job is narrow and deliberate. It reports, per cell, the formula and
the cached value, and it does not evaluate anything itself.

The reader is stdlib only (`zipfile` and `xml.etree.ElementTree`), which is why it can exist here at all: an
.xlsx is a zip of XML parts, and Python already reads zips.

What it handles, and why each part is needed:

- shared strings, including rich-text runs, which Excel splits into several `<r>` elements when part of a cell
  is bolded. Concatenating the runs is what the cell actually displays. - the 1904 date system, because a
  workbook written on a Mac by a user who ticked the box counts days from 1904-01-01 and every date reads four
  years out if that flag is ignored. - cell types: shared, inline, boolean, error and number, plus a formula
  whose cached value is missing. - styles, only as far as "is this number a date", which is the one formatting
  question that changes a value rather than its appearance.

Deliberate refusals, named rather than guessed: a formula cell with no cached value is reported as such
instead of being read as an empty string, because an empty string is a different value and would corrupt a
parity comparison in the direction that flatters the engine. A cell range Excel stored as an array formula, an
unsupported cell type, and a part the reader cannot find are errors with the part and the cell named.
"""

from __future__ import annotations

import datetime as _datetime
import io
import re
import zipfile
from dataclasses import dataclass, field
from typing import Iterator
from xml.etree import ElementTree

_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkg": "http://schemas.openxmlformats.org/package/2006/relationships",
}

_CELL_REFERENCE = re.compile(r"^([A-Z]+)(\d+)$")

# Built-in number formats that are dates or times. Ids 14-22 and 45-47 are the documented ones; anything else
# is checked by looking for date letters in a custom format code.
_BUILTIN_DATE_FORMATS = frozenset({14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57})


class XlsxError(Exception):
    """A workbook the reader will not guess about."""


@dataclass(frozen=True)
class Cell:
    """One cell: where it is, what formula produced it, and the value Excel cached for it."""

    row: int
    column: int
    reference: str
    formula: str | None
    value: object
    """`str`, `float`, `bool`, `datetime.datetime`, or `None` when Excel cached nothing."""
    kind: str
    """`number`, `string`, `boolean`, `error`, `date`, `blank`, or `uncached`."""

    def __post_init__(self) -> None:
        if self.kind == "uncached" and self.value is not None:
            raise XlsxError(f"{self.reference}: a cell with no cached value cannot carry one")


@dataclass
class Sheet:
    name: str
    cells: dict[str, Cell] = field(default_factory=dict)
    """Keyed by A1-style reference, which is how a formula names them."""

    def value(self, reference: str) -> object:
        cell = self.cells.get(reference.upper())
        if cell is None:
            # Excel treats a cell that was never written as empty, not as an error, so this is not one either
            return None
        if cell.kind == "uncached":
            raise XlsxError(
                f"{self.name}!{reference} has a formula and no cached value: recalculate it in Excel first"
            )
        return cell.value

    def __iter__(self) -> Iterator[Cell]:
        return iter(self.cells.values())


@dataclass
class Workbook:
    sheets: list[Sheet]
    date1904: bool
    defined_names: dict[str, str] = field(default_factory=dict)

    def sheet(self, name: str) -> Sheet:
        for sheet in self.sheets:
            if sheet.name == name:
                return sheet
        raise XlsxError(f"no sheet named {name!r}; this workbook has {[sheet.name for sheet in self.sheets]}")


def _column_index(letters: str) -> int:
    index = 0
    for character in letters:
        index = index * 26 + (ord(character) - ord("A") + 1)
    return index


def _parse_reference(reference: str) -> tuple[int, int]:
    match = _CELL_REFERENCE.match(reference)
    if match is None:
        raise XlsxError(f"{reference!r} is not an A1-style cell reference")
    letters, digits = match.groups()
    return int(digits), _column_index(letters)


def _text_of(node: ElementTree.Element | None) -> str:
    """The concatenated text of an element's `<t>` children, which is what a rich-text cell displays."""
    if node is None:
        return ""
    pieces: list[str] = []
    for text in node.iter(f"{{{_NS['main']}}}t"):
        pieces.append(text.text or "")
    return "".join(pieces)


def _read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        raw = archive.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ElementTree.fromstring(raw)
    return [_text_of(item) for item in root.findall("main:si", _NS)]


def _read_sheet_targets(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    """Sheet name and part path, in workbook order, resolved through the workbook's relationships part."""
    workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    relationships = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        relationship.get("Id"): relationship.get("Target", "")
        for relationship in relationships.findall("pkg:Relationship", _NS)
    }
    sheets: list[tuple[str, str]] = []
    for sheet in workbook.findall("main:sheets/main:sheet", _NS):
        identifier = sheet.get(f"{{{_NS['rel']}}}id")
        if identifier is None:
            raise XlsxError(
                f"sheet {sheet.get('name')!r} has no relationship id, so its part cannot be found"
            )
        target = targets.get(identifier)
        if target is None:
            raise XlsxError(
                f"sheet {sheet.get('name')!r} points at relationship {identifier!r}, which is not declared"
            )
        path = target.lstrip("/")
        if not path.startswith("xl/"):
            # a relative target is relative to xl/, which is where workbook.xml lives
            path = f"xl/{path}"
        sheets.append((sheet.get("name", ""), path))
    return sheets


def _date_style_indexes(archive: zipfile.ZipFile) -> set[int]:
    """Style indexes (`s` attribute values) whose number format is a date."""
    try:
        raw = archive.read("xl/styles.xml")
    except KeyError:
        return set()
    root = ElementTree.fromstring(raw)
    custom_formats: dict[int, str] = {}
    for number_format in root.findall("main:numFmts/main:numFmt", _NS):
        identifier = number_format.get("numFmtId")
        code = number_format.get("formatCode", "")
        if identifier is not None:
            custom_formats[int(identifier)] = code
    date_styles: set[int] = set()
    for index, cell_format in enumerate(root.findall("main:cellXfs/main:xf", _NS)):
        identifier = int(cell_format.get("numFmtId", "0"))
        if identifier in _BUILTIN_DATE_FORMATS or _looks_like_a_date(custom_formats.get(identifier, "")):
            date_styles.add(index)
    return date_styles


def _looks_like_a_date(format_code: str) -> bool:
    """Whether a custom format code is a date, judged by its date letters outside literals and brackets."""
    stripped = re.sub(r'"[^"]*"', "", format_code)
    stripped = re.sub(r"\[[^\]]*\]", "", stripped)
    stripped = stripped.lower()
    return any(letter in stripped for letter in ("yy", "dd", "mmm", "d/", "/d", "mm:", ":mm"))


def _serial_to_datetime(serial: float, date1904: bool) -> _datetime.datetime:
    # 1899-12-30 is Excel's day zero because of the 1900 leap-year bug it inherited from Lotus 1-2-3
    epoch = _datetime.datetime(1904, 1, 1) if date1904 else _datetime.datetime(1899, 12, 30)
    return epoch + _datetime.timedelta(days=serial)


def _cell_from_node(
    node: ElementTree.Element,
    reference: str,
    shared_strings: list[str],
    date_styles: set[int],
    date1904: bool,
) -> Cell:
    row, column = _parse_reference(reference)
    style = node.get("s")
    kind_attribute = node.get("t", "n")
    formula_node = node.find("main:f", _NS)
    formula = formula_node.text if formula_node is not None else None
    value_node = node.find("main:v", _NS)
    text = value_node.text if value_node is not None else None

    if kind_attribute == "inlineStr":
        return Cell(row, column, reference, formula, _text_of(node.find("main:is", _NS)), "string")
    if text is None:
        if formula is not None:
            # Excel stores a formula it has not calculated without a <v>; reading it as empty would be wrong
            return Cell(row, column, reference, formula, None, "uncached")
        return Cell(row, column, reference, None, None, "blank")
    if kind_attribute == "s":
        index = int(text)
        if index >= len(shared_strings):
            raise XlsxError(
                f"{reference}: shared string {index} is not in sharedStrings.xml, "
                f"which has {len(shared_strings)}"
            )
        return Cell(row, column, reference, formula, shared_strings[index], "string")
    if kind_attribute == "str":
        return Cell(row, column, reference, formula, text, "string")
    if kind_attribute == "b":
        return Cell(row, column, reference, formula, text.strip() == "1", "boolean")
    if kind_attribute == "e":
        return Cell(row, column, reference, formula, text, "error")
    if kind_attribute == "n":
        number = float(text)
        is_date = style is not None and int(style) in date_styles
        value = _serial_to_datetime(number, date1904) if is_date else number
        return Cell(row, column, reference, formula, value, "date" if is_date else "number")
    raise XlsxError(f"{reference}: cell type {kind_attribute!r} is not one this reader understands")


def _read_sheet(
    path: str,
    name: str,
    raw: bytes,
    shared_strings: list[str],
    date_styles: set[int],
    date1904: bool,
) -> Sheet:
    sheet = Sheet(name=name)
    root = ElementTree.fromstring(raw)
    for node in root.iter(f"{{{_NS['main']}}}c"):
        reference = node.get("r")
        if reference is None:
            # a cell without a reference is only legal inside an array; guessing its position would corrupt
            # every following cell's position
            raise XlsxError(f"{path}: a cell has no reference, which this reader will not guess at")
        cell = _cell_from_node(node, reference.upper(), shared_strings, date_styles, date1904)
        sheet.cells[cell.reference] = cell
    return sheet


def read_workbook(source: str | bytes | io.BufferedIOBase) -> Workbook:
    """Read an .xlsx into sheets. Accepts a path, raw bytes, or an open binary file."""
    if isinstance(source, str):
        source = io.BytesIO(_read(source))
    elif isinstance(source, bytes):
        # zipfile needs a file-like object; a caller holding a downloaded file's bytes should not need to know
        source = io.BytesIO(source)
    archive = zipfile.ZipFile(source)
    with archive:
        shared_strings = _read_shared_strings(archive)
        date_styles = _date_style_indexes(archive)
        workbook_part = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        properties = workbook_part.find("main:workbookPr", _NS)
        date1904 = properties is not None and properties.get("date1904") in {"1", "true"}
        defined_names: dict[str, str] = {}
        for defined_name in workbook_part.findall("main:definedNames/main:definedName", _NS):
            identifier = defined_name.get("name")
            if identifier is not None:
                defined_names[identifier] = (defined_name.text or "").strip()
        sheets: list[Sheet] = []
        for name, path in _read_sheet_targets(archive):
            try:
                raw = archive.read(path)
            except KeyError as error:
                raise XlsxError(f"sheet {name!r} names part {path!r}, which is not in the archive") from error
            sheets.append(_read_sheet(path, name, raw, shared_strings, date_styles, date1904))
    return Workbook(sheets=sheets, date1904=date1904, defined_names=defined_names)


def _read(path: str) -> bytes:
    with open(path, "rb") as handle:
        return handle.read()
