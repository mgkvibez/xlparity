"""VLOOKUP, INDEX and MATCH, where spreadsheet reimplementations diverge most.

These are the cases that silently differ between engines: approximate match as the default, text comparison
ignoring case, blanks taking the type of the thing they are compared against, an index that truncates rather
than rounds, and the difference between "not found" (#N/A) and "out of range" (#REF!).
"""

from __future__ import annotations

from xlparity import ExcelError, evaluate, parse

SHEET = {
    "A1": 1, "A2": 3, "A3": 5, "A4": 7,
    "B1": "one", "B2": "three", "B3": "five", "B4": "seven",
    "C1": "Apple", "C2": "Banana", "C3": "Cherry",
}


def run(formula: str, cells: dict | None = None):
    return evaluate(parse(formula), SHEET if cells is None else cells)


def code(value):
    assert isinstance(value, ExcelError), f"expected an error value, got {value!r}"
    return value.code


class TestVlookup:
    def test_exact_match_returns_the_column_to_the_right(self):
        assert run("VLOOKUP(3,A1:B4,2,FALSE)") == "three"

    def test_approximate_is_the_default_and_takes_the_largest_value_at_or_below(self):
        # 4 is not in the first column: approximate returns the row for 3, which is Excel's answer
        assert run("VLOOKUP(4,A1:B4,2)") == "three"

    def test_approximate_and_exact_disagree_when_the_lookup_value_is_absent(self):
        assert run("VLOOKUP(4,A1:B4,2)") == "three"
        assert code(run("VLOOKUP(4,A1:B4,2,FALSE)")) == "#N/A"

    def test_a_lookup_below_every_key_is_not_found(self):
        assert code(run("VLOOKUP(0,A1:B4,2)")) == "#N/A"

    def test_a_column_past_the_table_is_a_reference_error_not_a_value_error(self):
        assert code(run("VLOOKUP(3,A1:B4,3,FALSE)")) == "#REF!"

    def test_column_zero_is_a_value_error(self):
        assert code(run("VLOOKUP(3,A1:B4,0,FALSE)")) == "#VALUE!"

    def test_text_lookup_ignores_case(self):
        cells = {"A1": "Alpha", "B1": 10, "A2": "beta", "B2": 20}
        assert run("VLOOKUP(\"BETA\",A1:B2,2,FALSE)", cells) == 20

    def test_a_scalar_table_is_refused_rather_than_guessed(self):
        assert code(run("VLOOKUP(1,5,1)")) == "#VALUE!"

    def test_an_error_in_the_lookup_value_propagates_unchanged(self):
        assert code(run("VLOOKUP(1/0,A1:B4,2,FALSE)")) == "#DIV/0!"


class TestMatch:
    def test_exact_match_returns_a_one_based_position(self):
        assert run("MATCH(5,A1:A4,0)") == 3

    def test_approximate_ascending_takes_the_largest_value_at_or_below(self):
        assert run("MATCH(4,A1:A4,1)") == 2

    def test_descending_match_walks_the_other_way(self):
        cells = {"A1": 9, "A2": 6, "A3": 3, "A4": 1}
        assert run("MATCH(5,A1:A4,-1)", cells) == 2

    def test_not_found_is_na(self):
        assert code(run("MATCH(4,A1:A4,0)")) == "#N/A"

    def test_text_match_is_case_insensitive(self):
        assert run("MATCH(\"cherry\",C1:C3,0)") == 3

    def test_an_invalid_match_type_is_na(self):
        assert code(run("MATCH(5,A1:A4,5)")) == "#N/A"


class TestIndex:
    def test_two_dimensional_index_addresses_row_and_column(self):
        assert run("INDEX(A1:B4,2,2)") == "three"

    def test_single_argument_form_indexes_a_vector(self):
        assert run("INDEX(B1:B4,4)") == "seven"

    def test_out_of_range_row_is_a_reference_error(self):
        assert code(run("INDEX(A1:B4,9,1)")) == "#REF!"

    def test_out_of_range_column_is_a_reference_error(self):
        assert code(run("INDEX(A1:B4,1,9)")) == "#REF!"

    def test_a_zero_index_returns_the_whole_axis_now_that_array_values_exist(self):
        # INDEX(range,0,1) is "the whole first column" in Excel: an array, which is a value the engine now has
        cells = {"A1": 1, "B1": 9, "A2": 2, "B2": 8, "A3": 3, "B3": 7, "A4": 4, "B4": 6}
        assert run("SUM(INDEX(A1:B4,0,1))", cells) == 10
        assert run("SUM(INDEX(A1:B4,2,0))", cells) == 10
        assert run("SUM(INDEX(A1:B4,0,0))", cells) == 40
        # a negative index is still not a thing
        assert code(run("INDEX(A1:B4,-1,1)")) == "#VALUE!"

    def test_a_blank_cell_comes_back_as_zero_in_a_numeric_context(self):
        cells = {"A1": 1, "B1": None, "A2": 2, "B2": 4}
        assert run("INDEX(A1:B2,1,2)+0", cells) == 0
