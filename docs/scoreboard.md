# The compatibility scoreboard

The product is the number, not the library. This page defines the number so it cannot be quietly redefined later, which is
the same rule this project's fixture gate follows elsewhere: **pre-register the metric, never redefine it after seeing
results.**

## What is measured

For every (function, formula, input) triple in the corpus:

| field | meaning |
| --- | --- |
| `formula` | the formula text, exactly as Excel would store it |
| `inputs` | the cell values the formula reads |
| `expected` | Excel's own result, captured once and frozen |
| `actual` | what this engine returns |
| `match` | `actual == expected` under the comparison rules below |

## How matches are judged

- Numbers compare with a relative tolerance of 1e-12, because a floating-point result that differs in the fifteenth digit
  is not a parity failure, it is a printing difference.
- Error values compare by code: `#DIV/0!` matches only `#DIV/0!`.
- Booleans and text compare exactly.
- Rounding mode differences are **failures**, not tolerance cases. `ROUND(2.5,0)=3` is parity; `2` is not.

## Cohort scores, reported separately

1. **Core arithmetic and precedence** — the baseline. Should be 100%.
2. **Text, comparison and coercion** — where "Excel-ish" engines usually diverge.
3. **Common functions** — SUM/IF/ROUND/INDEX/MATCH/VLOOKUP class.
4. **Obscure financial** — IRR, XIRR, NPV, and the depreciation family.
5. **Dynamic arrays** — FILTER, SORT, UNIQUE, SEQUENCE, spill behaviour.

A single headline number hides exactly the failures users hit, so the scoreboard reports every cohort and never only a
total.

## Round-trip fidelity, measured separately

Separate from calculation: open a real XLSX, write it back untouched, and diff the re-saved file against the original.
Charts, pivot caches, conditional formats and defined names must survive. Any loss is listed by element name.

## The current score

<!-- BEGIN GENERATED: scoreboard -->

**Expectation provenance:** 25 derived, 35 docs. `docs` means Microsoft's documented
semantics; `excel` means captured from Excel and frozen; `derived` means computed here
from the documented closed form with exact rational arithmetic, which is a second
implementation of the definition rather than a capture of Excel's own result. No `docs`
or `derived` entry claims to be an `excel` capture.

| cohort | passing | total | score |
| --- | --- | --- | --- |
| 1. Core arithmetic and precedence | 10 | 10 | 100.0% |
| 2. Text, comparison and coercion | 7 | 7 | 100.0% |
| 3. Common functions | 18 | 18 | 100.0% |
| 4. Obscure financial | 25 | 25 | 100.0% |
| 5. Dynamic arrays | 0 | 0 | n/a |

**All cohorts:** 60 of 60 formulas match (100.0%).

Every corpus formula matches. The corpus is grown, never trimmed, so this is only
true until the next difficult formula is added.

<!-- END GENERATED: scoreboard -->

## Rule

The corpus is grown, never trimmed. A failing formula stays in the scoreboard after it starts passing, so the score stays
honest over time.
