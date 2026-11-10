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
- 91 tests.

## What does not work yet

- The scoreboard is real but not yet Excel-captured: 60 corpus formulas, none tagged `excel`. 35 are tagged `docs`
  (Microsoft's documented semantics) and 25 `derived` (computed here from the documented closed form with exact
  rational arithmetic, which is an independent implementation of the definition, not a capture of Excel's output).
  Cohort 5 is still empty: no dynamic arrays, so no `FILTER`/`SORT`/`UNIQUE`/`SEQUENCE`, no `XLOOKUP`, and no XLSX
  reading.
- No round-trip fidelity measurement yet: the second half of the metric, re-saving an untouched XLSX and diffing it, is
  defined but unbuilt.
- No locale handling: Excel's separators and formats vary by locale, and this engine assumes the English one.

## Status

**Stage 0: scaffolding.** The parser and evaluator exist with 91 passing tests covering the details that "Excel-ish"
engines get wrong:

**Tests:** 91

Lookups (`VLOOKUP`, `INDEX`, `MATCH`) exist with Excel's semantics rather than the convenient ones: approximate match is
the default, text comparison ignores case, `INDEX` truncates a fractional index, "not found" is `#N/A` while "out of
range" is `#REF!`, and the array forms (`INDEX(range,0,1)`, whole-column results) are refused with `#VALUE!` because this
engine has no array values.


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
4. **XLSX reader**: shared strings, styles, formulas, defined names, date systems.
5. **XLSX writer with round-trip fidelity**: charts, pivots, conditional formats preserved or listed as lost by name.
6. **Dynamic arrays**: spill semantics, `#` operator, `FILTER`/`SORT`/`UNIQUE`/`SEQUENCE`.

## Non-goals

- Not a UI. A library plus a scoreboard; an editor can sit on top.
- No macro/VBA execution.
- No attempt to match Excel's floating-point quirks beyond the documented 15-significant-digit behaviour.

## Licence

MIT.
