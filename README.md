# xlparity

**Excel parity, not Excel-ish.**

Formula engines are claimed: IronCalc alone ships 300+ functions and XLSX support. Parity is unclaimed. Parity means the
full function set including the obscure financial and dynamic-array ones, plus a writer that round-trips charts, pivots,
conditional formats and defined names without losing them.

**The scoreboard is the product.** A number you can read, that goes up, measured against Excel's own output on a corpus of
real spreadsheets.

## Why this fits a fixture-gate workflow

Parity is not a matter of opinion if you pre-register the metric. [`docs/scoreboard.md`](docs/scoreboard.md) defines what
counts as a match, cohort by cohort, before any result exists. The corpus grows and is never trimmed, and a failing formula
stays on the board after it starts passing, so the score cannot be quietly improved by redefining the test.

## What works today

- Parser and evaluator: Excel operator precedence (`-2^2` is `4`, `^` is right-associative), comparisons across
  types in Excel's order (numbers < text < logicals, text compared case-insensitively), error *values* that
  propagate instead of raising.
- Ranges: `A1:C3`, `$A$1:$B$2`, reversed corners. Aggregates ignore text, blanks and logicals **inside** a block
  while a scalar argument is coerced, so `SUM(1,"x",3)` is `#VALUE!` and `SUM(A1:A4)` skips the text, which is what
  Excel actually does.
- 32 functions: SUM, MIN, MAX, AVERAGE, COUNT, COUNTA, IF, IFERROR, AND, OR, NOT, ISBLANK, ISNUMBER, ISTEXT,
  ISERROR, ROUND, ROUNDUP, ROUNDDOWN, INT, MOD, SQRT, ABS, LEN, LEFT, RIGHT, MID, TRIM, UPPER, LOWER, CONCAT,
  VLOOKUP, INDEX, MATCH, NPV, IRR, XIRR, SLN, SYD, DB, DDB.
- The financial cohort, filled rather than promised: `NPV` discounts its first value by one whole period, `IRR` and
  `XIRR` solve for the root and refuse `#NUM!` when a cash flow has none, `DB` rounds its rate to three decimals before
  depreciating (that rounding is the definition, not a tidy-up), and `DDB` stops at salvage rather than following its
  formula past it.
- `SEQUENCE`, `UNIQUE`, `SORT` and `FILTER` return real array values, so `SUM(SEQUENCE(5))` is 15, `INDEX(SORT(A1:A4),1)` is
  the smallest value, and `SUM(FILTER(A1:A4,B1:B4))` adds only the flagged rows. A computed array that reaches the top of a
  formula is `#SPILL!`, not `#VALUE!`.
- 135 tests.

## Dynamic arrays, and the half of them that needs a sheet

A dynamic array is two things: a value, and a place to put it. The value is here. `SEQUENCE(4)` is four numbers and
`SUM` of them is 10 whether or not anything spills, which makes the inside-a-formula half genuinely testable without a
grid.

The place is not. This engine evaluates one formula for one cell, so a result that wants several cells has nowhere to land,
and the honest error for that is `#SPILL!`: the formula is right and the sheet is missing. `#VALUE!` would have said the
opposite, so the two cases are kept apart, and a bare range reference still answers `#VALUE!` because a range used where a
value was asked for is a different mistake from an array that could not be placed.

Three limits, stated rather than discovered later:

1. There is no broadcasting, so `FILTER(A1:A10, B1:B10>5)` cannot be written. The include argument has to be a range of
   its own, one flag per value: `FILTER(A1:A10,B1:B10)` where the flags are already in column B. That is the arrangement
   the function needs; it is not the arrangement people write.
2. `UNIQUE` and `SORT` handle an array of one row or one column. De-duplicating rows of a matrix, or sorting by its second
   column, needs a walk this does not do, and is `#VALUE!` rather than the first column's answer.
3. `#` is not a reference operator here. `A1#` is a syntax error, and the cells a spill occupies do not exist to name.

## Reading Excel rather than imitating it

The scoreboard's `excel` column needs Excel's own answers, and the only honest source of those is a workbook Excel wrote:
an `.xlsx` carries the formula text and the value Excel cached for it, side by side. `xlparity.xlsx.read_workbook` reads
that, and deliberately does not evaluate anything.

It is stdlib only (`zipfile` and `xml.etree`), which is why it can exist here: an `.xlsx` is a zip of XML parts. What it
handles, and why each part matters:

1. Shared strings, including rich text. Excel splits a cell into several `<r>` runs when part of it is bolded, and joining
   those runs is what the cell displays; taking only the first run silently truncates.
2. The 1904 date system. A workbook written with that flag counts days from 1904-01-01, and ignoring the flag reads every
   date four years out, which is the kind of error that looks like an evaluation bug.
3. Cell types: shared, inline, boolean, error, number, plus the date question. A date here is a number with a date format,
   so the reader reports `date` and a `datetime` only for cells whose style says so: the same serial with an ordinary
   format stays a number, and that distinction is formatting, not value.
4. Defined names, read as written rather than resolved.
5. A formula cell with no cached value is reported as `uncached` and raises when asked for its value. It is never read as
   an empty string: an empty string is a different value, and treating one as the other would flatter the engine in a
   parity comparison.

The limit is worth stating plainly: the tests assemble their workbooks part by part with Python's `zipfile`, because there
is no Excel on this machine. They check the semantics I can state from the specification, not that Excel's own writer emits
what I expect. A real workbook added to the corpus is what would settle that.

`tools/import_workbook.py` is the bridge from a workbook to corpus entries: it walks each sheet, takes every formula cell
whose value Excel cached, reconstructs the inputs from the neighbouring cached values, and emits an entry tagged `excel`
with the workbook, sheet and cell recorded. It computes nothing itself. A cell whose own value is uncached, one that
depends on an uncached cell, one that reaches another sheet, a self-reference, an array formula and a formatted date are
each skipped with their reason printed, because replaying them would compare the engine against a different workbook from
the one Excel calculated. The `excel` column stays at zero until someone runs a real workbook through it, which is the one
claim this repository cannot make on this machine.

## What does not work yet

- The scoreboard is real but not yet Excel-captured: 60 corpus formulas, none tagged `excel`. 35 are tagged `docs`
  (Microsoft's documented semantics) and 25 `derived` (computed here from the documented closed form with exact
  rational arithmetic, which is an independent implementation of the definition, not a capture of Excel's output).
  Cohort 5 is still empty: no dynamic arrays, so no `FILTER`/`SORT`/`UNIQUE`/`SEQUENCE` and no `XLOOKUP`. The XLSX
  reader can now supply those captured values, but no real workbook has been run through the scoreboard yet, so every
  figure it prints still rests on documentation and closed forms rather than on Excel.
- No round-trip fidelity measurement yet: the second half of the metric, re-saving an untouched XLSX and diffing it, is
  defined but unbuilt.
- No locale handling: Excel's separators and formats vary by locale, and this engine assumes the English one.

## Status

**Stage 0: scaffolding.** The parser and evaluator exist with 135 passing tests covering the details that "Excel-ish"
engines get wrong:

**Tests:** 135

Lookups (`VLOOKUP`, `INDEX`, `MATCH`) exist with Excel's semantics rather than the convenient ones: approximate match is
the default, text comparison ignores case, `INDEX` truncates a fractional index, "not found" is `#N/A` while "out of
range" is `#REF!`, and the array forms (`INDEX(range,0,1)`, whole-column results) now return the whole axis, because the
engine has array values and returning the first cell of an axis would be a confident lie.


- Excel's precedence exactly, including right-associative `^` and `-2^2 = 4` rather than `-4`
- error **values** (`#DIV/0!`, `#NAME?`, `#VALUE!`) that propagate through arithmetic instead of raising
- `ROUND` half away from zero, not Python's banker's rounding
- blank cells that are `0` in arithmetic and `""` in concatenation
- `SUM` ignoring text, `ISERROR` testing a propagated error

Financial functions exist with the definitions rather than the convenient approximations, including the two that are
easy to get subtly wrong: `DB` rounds its rate to three decimals, and `DDB` caps at salvage, so its last year is
`22.12` on a 2400/300/10 asset rather than the `64.42` its own formula suggests. `IRR` and `XIRR` use Newton from the
guess and fall back to a bracketed scan, so a series with several sign changes returns the root nearest the guess; that
is a deviation from Excel's iterate-from-the-guess-only behaviour, which returns `#NUM!` where this returns a rate.

No XLSX reading or writing yet, and no dynamic arrays.

```bash
pip install -e ".[dev]" && bash scripts/verify.sh
```

`scripts/verify.sh` is the gate CI runs: ruff, the suite, the README test count, and the scoreboard page against the
corpus. The scoreboard check was previously a tool nobody called, which meant the README's claim that CI keeps the page
current was false; it is enforced now.

## Roadmap

1. **Parser and evaluator** (done, growing function by function; 32 of them, including the financial cohort).
2. **Corpus**: real spreadsheets plus per-function cases with Excel-captured expected values.
3. **Scoreboard**: cohort scores, published, and gated in `scripts/verify.sh`, so a page that disagrees with the corpus
   fails CI instead of ageing quietly.
4. **XLSX reader and corpus import** (done for values and formulas): shared strings with rich text, cell types, styles as far as dates,
   formulas with their cached values, defined names, and both date systems. Not yet: array and shared formulas, and
   anything a chart or pivot carries.
5. **XLSX writer with round-trip fidelity**: charts, pivots, conditional formats preserved or listed as lost by name.
6. **Dynamic arrays**: `FILTER`/`SORT`/`UNIQUE`/`SEQUENCE` and `INDEX`'s whole-axis forms are done, as values inside a
   formula. Not done: the cells a result spills into and the `#` operator that names them, which need a sheet.

## Non-goals

- Not a UI. A library plus a scoreboard; an editor can sit on top.
- No macro/VBA execution.
- No attempt to match Excel's floating-point quirks beyond the documented 15-significant-digit behaviour.

## Licence

MIT.
