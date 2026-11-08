"""Precedence, error propagation and the Excel rules that a naive calculator gets wrong."""

import pytest

from xlparity import ExcelError, FormulaError, evaluate, parse


def value(source: str, cells=None):
    return evaluate(parse(source), cells or {})


def test_multiplication_binds_tighter_than_addition():
    assert value("1+2*3") == 7


def test_parentheses_override_precedence():
    assert value("(1+2)*3") == 9


def test_exponent_is_right_associative():
    assert value("2^3^2") == 512


def test_unary_minus_binds_tighter_than_exponent_in_excel():
    assert value("-2^2") == 4  # Excel: (-2)^2, unlike most languages


def test_percent_is_a_postfix_divide_by_100():
    assert value("50%") == 0.5


def test_concatenation_joins_text_and_numbers():
    assert value('"a"&1&"b"') == "a1b"


def test_comparison_returns_boolean():
    assert value("2>1") is True
    assert value("2<>2") is False


def test_division_by_zero_is_an_error_value_not_an_exception():
    result = value("1/0")
    assert isinstance(result, ExcelError)
    assert result.code == "#DIV/0!"


def test_errors_propagate_through_arithmetic():
    assert value("1/0+1").code == "#DIV/0!"


def test_iserror_detects_a_propagated_error():
    assert value("ISERROR(1/0)") is True


def test_unknown_function_is_a_name_error_value():
    assert value("NOPE(1)").code == "#NAME?"


def test_sum_ignores_text_and_blanks():
    assert value("SUM(1,2,3)") == 6
    assert value('SUM(1,"x",3)') == 4


def test_sum_reads_cells():
    assert value("SUM(A1,B1)", {"A1": 2, "B1": 3.5}) == 5.5


def test_blank_cell_is_zero_in_arithmetic_but_empty_in_text():
    assert value("A1+1", {}) == 1
    assert value('A1&"x"', {}) == "x"


def test_round_uses_half_away_from_zero_not_bankers_rounding():
    assert value("ROUND(2.5,0)") == 3
    assert value("ROUND(-2.5,0)") == -3
    assert value("ROUND(1.2345,2)") == 1.23


def test_if_returns_branches_and_defaults_to_false():
    assert value('IF(1,"yes","no")') == "yes"
    assert value("IF(0,1)") is False


def test_string_literal_with_doubled_quote():
    assert value('"he said ""hi"""') == 'he said "hi"'


def test_syntax_errors_raise_because_they_are_not_values():
    with pytest.raises(FormulaError):
        parse("1+")


def test_unterminated_string_raises():
    with pytest.raises(FormulaError):
        parse('"abc')


def test_error_compares_as_text_like_excel():
    assert value('1/0="x"').code == "#DIV/0!"
