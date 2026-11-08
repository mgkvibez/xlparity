"""Coverage for ranges, Excel type ordering, and the wider function set."""

import pytest

from xlparity import evaluate, parse

SHEET = {
    "A1": 1.0,
    "A2": 2.0,
    "A3": 3.0,
    "A4": "text",
    "B1": 10.0,
    "B2": 20.0,
    "B3": None,
    "C1": True,
    "C2": False,
}


def value(source, cells=None):
    return evaluate(parse(source), cells if cells is not None else SHEET)


# --- ranges -----------------------------------------------------------------


def test_range_reads_a_rectangle_row_major():
    assert value("SUM(A1:B2)") == 33.0


def test_range_ignores_text_and_blanks_inside_the_block():
    assert value("SUM(A1:A4)") == 6.0
    assert value("COUNT(A1:B3)") == 5.0
    assert value("COUNTA(A1:A4)") == 4.0


def test_range_accepts_absolute_markers_and_reversed_corners():
    assert value("SUM($A$1:$A$3)") == 6.0
    assert value("SUM(B2:A1)") == 33.0


def test_a_bare_range_outside_a_function_is_a_value_error():
    assert value("A1:B2").code == "#VALUE!"


def test_min_max_average_over_a_range():
    assert value("MIN(A1:A3)") == 1.0
    assert value("MAX(A1:A3)") == 3.0
    assert value("AVERAGE(A1:A3)") == 2.0
    assert value("AVERAGE(B3:B3)").code == "#DIV/0!"


def test_range_needs_references_on_both_sides():
    with pytest.raises(Exception):
        parse("SUM(A1:3)")


# --- ordering ---------------------------------------------------------------


def test_numbers_sort_before_text_and_text_before_logicals():
    assert value('1<"a"') is True
    assert value('"a"<TRUE') is True
    assert value("1<TRUE") is True


def test_text_comparison_is_case_insensitive():
    assert value('"Apple"="apple"') is True
    assert value('"apple"<"banana"') is True


def test_a_blank_takes_the_type_it_is_compared_with():
    blank = {"A1": None}
    assert value("A1=0", blank) is True
    assert value('A1=""', blank) is True
    assert value("A1=FALSE", blank) is True


def test_false_sorts_before_true():
    assert value("FALSE<TRUE") is True
    assert value("TRUE>1") is True


# --- aggregates and logic ---------------------------------------------------


def test_aggregates_propagate_errors():
    assert value("SUM(A1,1/0)").code == "#DIV/0!"


def test_and_or_not_coerce_like_excel():
    assert value("AND(TRUE,1)") is True
    assert value("AND(TRUE,0)") is False
    assert value("OR(FALSE,0)") is False
    assert value("OR(FALSE,1)") is True
    assert value("NOT(0)") is True
    assert value('AND("yes")').code == "#VALUE!"


def test_iferror_swallows_only_errors():
    assert value("IFERROR(1/0,\"fallback\")") == "fallback"
    assert value("IFERROR(7,\"fallback\")") == 7.0


def test_type_predicates():
    assert value("ISBLANK(B3)") is True
    assert value("ISNUMBER(A1)") is True
    assert value("ISNUMBER(A4)") is False
    assert value('ISTEXT("x")') is True
    assert value("ISTEXT(A1)") is False


# --- text functions ---------------------------------------------------------


def test_text_functions_follow_excel_edge_cases():
    assert value('LEN("hello")') == 5.0
    assert value('LEFT("hello",2)') == "he"
    assert value('RIGHT("hello",2)') == "lo"
    assert value('RIGHT("hello",0)') == ""
    assert value('MID("hello",2,3)') == "ell"
    assert value('LEFT("hello",-1)').code == "#VALUE!"
    assert value('TRIM("  a   b  ")') == "a b"
    assert value('UPPER("aB")') == "AB"
    assert value('LOWER("aB")') == "ab"


# --- numeric functions ------------------------------------------------------


def test_rounding_variants():
    assert value("ROUND(2.5,0)") == 3.0
    assert value("ROUNDUP(2.1,0)") == 3.0
    assert value("ROUNDDOWN(2.9,0)") == 2.0
    assert value("INT(-2.1)") == -3.0
    assert value("ROUNDUP(-2.1,0)") == -3.0


def test_mod_takes_the_sign_of_the_divisor():
    assert value("MOD(-3,2)") == 1.0
    assert value("MOD(3,-2)") == -1.0
    assert value("MOD(3,0)").code == "#DIV/0!"


def test_sqrt_refuses_negatives_with_num_not_value():
    assert value("SQRT(9)") == 3.0
    assert value("SQRT(-1)").code == "#NUM!"


def test_unknown_function_stays_a_name_error():
    assert value("FLOOR(1)").code == "#NAME?"
