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

## Rule

The corpus is grown, never trimmed. A failing formula stays in the scoreboard after it starts passing, so the score stays
honest over time.
